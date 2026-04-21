"""
Agentic Win Story Generator — Flask Backend
Serves the frontend, handles AI extraction, and generates a single-slide
agentic use case PPTX matching the UiPath Agent Use Case Template v10.

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

Rules:
- TITLE: sharp 6-12 word title describing the outcome (e.g. "Unblocking appeals by automating the indexing bottleneck")
- SUBTITLE: one line, 8-14 words, summarizing the actor orchestration (e.g. "Agents reason over denials; bots execute; humans approve the edge cases.")
- PROBLEM DESCRIPTION: 2 concise sentences describing the pain. Professional tone.
- PROBLEM STATS: extract up to 4 quantitative problem metrics if present (dollar amounts, percentages, volumes, time values). Format value cleanly (e.g. "$181", "12%+", "$250B+", "21 days"). If fewer than 2 are clearly present, return fewer or empty array.
- SOLUTION DESCRIPTION: 1-2 concise sentences describing how agents, bots, and humans together resolve the problem. Reference UiPath capabilities/products naturally.
- CAPABILITIES: list of UiPath product names/capabilities used, 3-6 items (e.g. "Agents", "Unattended Robots", "Doc Understanding", "Action Center", "Maestro", "Data Service").
- STEPS: break the automation into 3-9 discrete steps. For each, classify the actor role:
  - "AGENT" = AI reasoning, classification, routing, decision
  - "BOT"   = deterministic RPA (data entry, API calls, portal polling, system updates, notifications)
  - "HUMAN" = human in the loop (review, approve, sign off on exceptions)
  Each step description should be 3-6 words, imperative.
- OUTCOMES: 1-4 measured outcome tiles. Each has a value (e.g. "$558K", "90%", "9 min") and a short label (e.g. "revenue released", "of workflow automated"). Only include what the notes actually measure.
- ATTRIBUTABLE IMPACT (optional): qualitative directional metrics directly attributable to the solution but not yet quantified. Format as "↓ Days in AR · ↓ Touches per claim · ↑ First-pass yield". Use ↓ for reductions, ↑ for improvements, separate with " · ". 3-5 items. Null if none inferable.
- DOWNSTREAM IMPACT (optional): second-order effects beyond direct solution outputs (e.g. staff retention, customer NPS, compliance posture). Same format as attributable. Null if none inferable.
- BREADCRUMB: three items for the top-of-slide breadcrumb: [industry], [function/department], [use case name]
- COMPANY: the customer company name
- THEME: "dark" by default

Return ONLY a valid JSON object. No markdown fences, no explanation:
{
  "breadcrumb": ["industry", "function", "use case name"],
  "title": "string",
  "subtitle": "string",
  "company": "string",
  "problem_desc": "string",
  "problem_stats": [{"value": "string", "label": "string"}],
  "solution_desc": "string",
  "capabilities": ["string", ...],
  "steps": [{"role": "AGENT|BOT|HUMAN", "description": "short step name"}],
  "outcomes": [{"value": "string", "label": "string"}],
  "attributable": "string or null",
  "downstream": "string or null",
  "theme": "dark"
}"""


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


@app.route('/generate', methods=['POST'])
def generate():
    """Takes structured JSON, returns PPTX bytes."""
    try:
        data = request.get_json(force=True)
        pptx_bytes, slide_num = build_pptx(data)
        customer = (data.get('company') or 'agentic_win_story').strip()
        safe_name = ''.join(c if c.isalnum() or c in ' _-' else '_' for c in customer)
        filename = f'{safe_name}_agentic_win_story.pptx'
        return Response(
            pptx_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            headers={
                'Content-Disposition': f'attachment; filename="{filename}"',
                'X-Slide-Number': str(slide_num),
            }
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Generation failed: {str(e)}'), 500


@app.route('/health')
def health():
    has_key = bool(os.environ.get('ANTHROPIC_API_KEY', '').strip())
    return jsonify(status='ok', ai_enabled=has_key)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    print(f'\n  Agentic Win Story Generator running at http://localhost:{port}\n')
    app.run(host='0.0.0.0', port=port, debug=False)
