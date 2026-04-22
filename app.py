"""
Agentic Win Story Generator - Flask Backend v2
Serves the frontend, handles AI extraction, and generates a single-slide
agentic use case PPTX.

Required env var: ANTHROPIC_API_KEY
"""
import json
import os
import traceback
from pathlib import Path

import requests as http_requests
from flask import Flask, request, send_from_directory, Response, jsonify
from flask_cors import CORS

from generate_pptx import build_pptx

app = Flask(__name__, static_folder='static')
CORS(app)


EXTRACTION_PROMPT = """You are a UiPath customer success writer. A rep has shared raw notes about an agentic automation win. Transform these notes into a polished, structured story for a single-slide agent use case template.

Rules (follow strictly):
- TITLE: SHORT and punchy. 3-6 words. Just the subject of the use case (e.g. "Medical claim adjudication", "Automating denials intake", "KYC refresh triage"). Do NOT include dollar amounts, percentages, or outcomes in the title — those go in the outcomes tiles. Default to the cleanest possible short title.
- SUBTITLE: one line, 8-14 words, summarizing the actor orchestration (e.g. "Agents reason over denials; bots execute; humans approve the edge cases.")
- PROBLEM DESCRIPTION: 2 concise sentences describing the pain. Professional tone. Max ~240 characters.
- PROBLEM STATS: extract up to 4 quantitative problem metrics if present (dollars, percentages, volumes, time). Format value cleanly (e.g. "$181", "12%+", "21 days"). Empty array if fewer than 2 are clearly stated.
- SOLUTION DESCRIPTION: 1-2 concise sentences describing how agents, bots, and humans resolve the problem. Max ~220 characters.
- CAPABILITIES: 3-6 UiPath product names/capabilities used (e.g. "Agents", "Unattended Robots", "Doc Understanding", "Action Center", "Maestro", "Data Service", "API Integration").
- STEPS: 3-9 discrete steps. Each has a role:
  - "AGENT" = AI reasoning, classification, routing, decisions
  - "BOT"   = deterministic RPA (data entry, API calls, portal polling, system updates)
  - "HUMAN" = human in the loop (review, approve, sign-off)
  Step description: 3-6 words, imperative.
- OUTCOMES: 1-5 measured outcome tiles. Each: value (e.g. "$558K", "90%", "9 min") + short label (e.g. "revenue released", "of workflow automated").
- ATTRIBUTABLE IMPACT (optional): list of directional metrics directly moved. Each item: {"direction": "up" | "down", "text": "metric name"}. 3-5 items. Empty list if none clearly inferable. Use "down" for reductions (cycle time, touches, backlogs) and "up" for improvements (yield, satisfaction, throughput).
- DOWNSTREAM IMPACT (optional): list of second-order effects (staff retention, NPS, compliance posture). Same format as attributable. Empty list if none inferable.
- BREADCRUMB: three items: [industry, function/department, use case name]
- COMPANY: customer company name
- THEME: "light" by default

Return ONLY a valid JSON object. No markdown fences, no preamble:
{
  "breadcrumb": ["industry", "function", "use case name"],
  "title": "string (short, 3-6 words)",
  "subtitle": "string",
  "company": "string",
  "problem_desc": "string",
  "problem_stats": [{"value": "string", "label": "string"}],
  "solution_desc": "string",
  "capabilities": ["string", "..."],
  "steps": [{"role": "AGENT|BOT|HUMAN", "description": "short step name"}],
  "outcomes": [{"value": "string", "label": "string"}],
  "attributable": [{"direction": "up|down", "text": "metric name"}],
  "downstream": [{"direction": "up|down", "text": "metric name"}],
  "theme": "light"
}"""




SUGGESTIONS_PROMPT = """You are a UiPath CSM helping build a stronger agentic win story. Be concise and specific.

Review the story data and return ONLY a JSON object with exactly these fields:
{
  "topSuggestion": "The single most impactful thing to do right now (1 sentence)",
  "missingData": [max 3 items: {"what": "data to find", "where": "Salesforce/Gainsight/CSM notes/AE/QBR deck", "why": "1 sentence"}],
  "agenticValue": [max 2 items: {"question": "probing question", "insight": "what it reveals"}],
  "storyAngles": [max 2 items: {"angle": "angle name", "suggestion": "action to take"}]
}

Focus on: missing metrics (ROI, cycle time, throughput, quality), agentic value beyond time saved (revenue impact, exception resolution, capacity), and where to find this data internally (Salesforce, Gainsight, QBR decks, CSM notes, AE). Keep every field under 20 words. Return ONLY the JSON, no explanation."""


STEPS_PROMPT = """You are a UiPath process analyst. Given a description of an agentic automation process, break it into 3-9 discrete steps.

For each step, classify the actor role:
- "AGENT" = AI reasoning, classification, routing, decisions
- "BOT"   = deterministic RPA (data entry, API calls, portal polling, system updates)
- "HUMAN" = human in the loop (review, approve, sign-off)

Step description: 3-6 words, imperative (e.g. "Pull & classify inbound denials", "Extract fields via Doc Understanding").

Return ONLY valid JSON:
{"steps": [{"role": "AGENT|BOT|HUMAN", "description": "short step name"}]}"""


def _call_claude(system, user_text, max_tokens=2048, timeout=45):
    api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
    if not api_key:
        raise ValueError('ANTHROPIC_API_KEY not configured on server.')
    resp = http_requests.post(
        'https://api.anthropic.com/v1/messages',
        headers={
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
        },
        json={
            'model': 'claude-haiku-4-5-20251001',
            'max_tokens': max_tokens,
            'system': system,
            'messages': [{'role': 'user', 'content': user_text}],
        },
        timeout=timeout,
    )
    if resp.status_code != 200:
        try:
            msg = resp.json().get('error', {}).get('message', resp.text)
        except Exception:
            msg = resp.text
        raise RuntimeError(f'Anthropic API error ({resp.status_code}): {msg}')
    raw = resp.json()['content'][0]['text'].strip()
    cleaned = raw.replace('```json', '').replace('```', '').strip()
    depth = 0
    start = cleaned.find('{')
    if start == -1:
        raise json.JSONDecodeError('No JSON object found', cleaned, 0)
    for i, ch in enumerate(cleaned[start:], start):
        if ch == '{': depth += 1
        elif ch == '}': depth -= 1
        if depth == 0:
            return json.loads(cleaned[start:i+1])
    return json.loads(cleaned)


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)


@app.route('/extract', methods=['POST'])
def extract():
    """Takes raw text notes, returns structured JSON for the agentic template."""
    try:
        body = request.get_json(force=True)
        text = (body.get('text') or '').strip()
        if not text:
            return jsonify(error='No input text provided.'), 400
        parsed = _call_claude(EXTRACTION_PROMPT, f'Rep notes:\n\n{text}')
        return jsonify(parsed)
    except json.JSONDecodeError as e:
        return jsonify(error=f'Could not parse AI response as JSON: {e}'), 500
    except http_requests.Timeout:
        return jsonify(error='AI request timed out. Try again.'), 504
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Extraction failed: {str(e)}'), 500


@app.route('/extract-steps', methods=['POST'])
def extract_steps():
    """Takes a process description, returns just the steps list."""
    try:
        body = request.get_json(force=True)
        text = (body.get('text') or '').strip()
        if not text:
            return jsonify(error='No process description provided.'), 400
        parsed = _call_claude(STEPS_PROMPT, f'Process description:\n\n{text}',
                              max_tokens=1024, timeout=30)
        return jsonify(parsed)
    except json.JSONDecodeError as e:
        return jsonify(error=f'Could not parse AI response as JSON: {e}'), 500
    except http_requests.