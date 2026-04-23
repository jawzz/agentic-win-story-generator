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
- TITLE: SHORT and punchy. 3-6 words. Just the subject of the use case (e.g. "Medical claim adjudication", "Automating denials intake", "KYC refresh triage"). Do NOT include dollar amounts, percentages, or outcomes in the title - those go in the outcomes tiles. Default to the cleanest possible short title.
- SUBTITLE: one line, 8-14 words, summarizing the actor orchestration (e.g. "Agents reason over denials; bots execute; humans approve the edge cases.")
- PROBLEM DESCRIPTION: 2 concise sentences describing the pain. Professional tone. Max ~240 characters.
- PROBLEM STATS: extract up to 4 quantitative problem metrics if present (dollars, percentages, volumes, time). Format value cleanly (e.g. "$181", "12%+", "21 days"). Empty array if fewer than 2 are clearly stated.
- SOLUTION DESCRIPTION: 1-2 concise sentences describing how agents, bots, and humans resolve the problem. Max ~220 characters.
- CAPABILITIES: 3-6 UiPath product names/capabilities used. Canonical names (use these exactly):
    - "Agents" - AI reasoning, autonomous decisioning
    - "Maestro" - the orchestration layer. ALWAYS include Maestro when the solution involves coordinating agents, bots, or humans across a process, a workflow engine, process/workload orchestration, routing, handoffs, SLA management, or anything the notes call "workload registration", "workflow registration", "orchestration", or "process controller". Default to including Maestro on any multi-step agentic solution.
    - "IXP" - intelligent document processing / data extraction from docs, emails, forms, invoices, claims, PDFs, faxes, images. NEVER use "Doc Understanding" or "Document Understanding" - always say "IXP".
    - "Unattended Robots" - deterministic RPA, system-to-system automation
    - "Attended Robots" - desktop assistant bots with humans
    - "Action Center" - human-in-the-loop approvals and reviews
    - "Data Service" - structured data storage
    - "API Integration" - integrations to external systems
    - "Test Suite", "Insights", "Apps" - only if clearly referenced
  If notes mention orchestration/routing/handoff/workflow engine: include Maestro. If notes mention any document/form/invoice/email extraction: include IXP (not Doc Understanding).
- STEPS: 3-9 discrete steps. Each has a role:
  - "AGENT" = AI reasoning, classification, routing, decisions
  - "BOT"   = deterministic RPA (data entry, API calls, portal polling, system updates)
  - "HUMAN" = human in the loop (review, approve, sign-off)
  Step description: 3-6 words, imperative. When a step extracts data from documents/emails/forms, say IXP (never "Doc Understanding"). When a step orchestrates or routes work across agents/bots/humans, say Maestro.
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


SUGGESTIONS_PROMPT = """You are a senior UiPath CSM helping a rep strengthen an agentic win story. Be brutally concise. You are talking to an internal UiPath seller (CSM or AE), not the customer.

Review the story data and return 1-4 suggestions — ONLY the most impactful gaps or angles. Fewer is better. Skip anything obvious or already covered.

Return ONLY this JSON:
{
  "suggestions": [
    {
      "priority": "high" | "medium",
      "title": "short action (3-8 words, imperative)",
      "detail": "1 sentence on why this matters (max 20 words)",
      "tip": "where someone might look for this data (max 12 words, see list below)"
    }
  ]
}

For "tip", suggest plausible internal sources like: Salesforce opportunity record, Gainsight CSM notes, the AE, QBR/EBR deck, UiPath Insights, Automation Hub, the customer champion, internal Slack channels, product marketing. Pick what actually fits — don't force it. If a suggestion isn't about missing data (e.g. a reframe or angle), the tip can be an action instead (e.g. "Reframe the intro to lead with revenue unlocked").

Rules:
- MAX 4 items. Prefer 2-3.
- High priority only for the 1-2 most impactful.
- Focus on: missing quant metrics (ROI, cycle time, FTE capacity, error rate, revenue), agentic value beyond time saved, story angles that reframe cost savings as revenue/capacity/decision-quality.
- No filler. Every suggestion must be concretely actionable.
- Return ONLY the JSON. No markdown, no preamble."""


STEPS_PROMPT = """You are a UiPath process analyst. Given a description of an agentic automation process, break it into 3-9 discrete steps.

For each step, classify the actor role:
- "AGENT" = AI reasoning, classification, routing, decisions
- "BOT"   = deterministic RPA (data entry, API calls, portal polling, system updates)
- "HUMAN" = human in the loop (review, approve, sign-off)

Step description: 3-6 words, imperative (e.g. "Pull & classify inbound denials", "Extract fields via IXP", "Route via Maestro"). When a step describes extracting data from documents/emails/forms/invoices, refer to it as IXP (never "Doc Understanding"). When a step describes orchestrating, routing, or handing off work across agents/bots/humans, refer to it as Maestro.

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
    except http_requests.Timeout:
        return jsonify(error='AI request timed out. Try again.'), 504
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Step extraction failed: {str(e)}'), 500


@app.route('/suggest', methods=['POST'])
def suggest():
    """Takes current story JSON, returns AI suggestions for strengthening it."""
    try:
        body = request.get_json(force=True) or {}
        story_json = json.dumps(body, ensure_ascii=False)
        parsed = _call_claude(SUGGESTIONS_PROMPT,
                              f'Current story data:\n\n{story_json}',
                              max_tokens=1024, timeout=30)
        return jsonify(parsed)
    except json.JSONDecodeError as e:
        return jsonify(error=f'Could not parse AI response as JSON: {e}'), 500
    except http_requests.Timeout:
        return jsonify(error='AI request timed out. Try again.'), 504
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Suggestions failed: {str(e)}'), 500


@app.route('/generate', methods=['POST'])
def generate():
    """Takes structured story JSON, returns the built .pptx as a download."""
    try:
        body = request.get_json(force=True) or {}
        result = build_pptx(body)
        pptx_bytes = result[0] if isinstance(result, tuple) else result
        company = (body.get('company') or 'agentic_win_story').strip()
        safe = ''.join(c if c.isalnum() or c in ('_', '-') else '_' for c in company)
        filename = f'{safe}_agentic_win_story.pptx'
        return Response(
            pptx_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Generate failed: {str(e)}'), 500


@app.route('/health')
def health():
    return jsonify(status='ok', has_api_key=bool(os.environ.get('ANTHROPIC_API_KEY', '').strip()))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
