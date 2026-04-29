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

TRANSPARENCY RULES FOR NUMBERS (follow carefully):
Every numerical metric (in problem_stats or outcomes) MUST be tagged with a "source" indicating where it came from:
- "stated"     = the exact number (or clear equivalent) appears in the notes. No derivation needed.
- "calculated" = you did simple arithmetic from numbers in the notes. Show the math in "note".
- "estimated"  = you inferred a plausible number from industry benchmarks or scope context. Label as estimate in "note".
- "qualitative"= no number given and none can be responsibly estimated. Use a word like "Reduced", "Faster", "Yes" as the value and leave source=qualitative.

RULES:
1. Every outcome and problem_stat item MUST have a "source" field.
2. Items with source = "calculated" or "estimated" MUST have a "note" field (max ~80 chars) explaining the derivation. Examples:
   - "From $500K/yr budget ÷ 12 months"
   - "Industry avg for claim triage (est.)"
   - "3,000 tickets × $120 handle cost"
3. Items with source = "stated" should still include a short "note" only if helpful (e.g. "per rep notes"). Usually the note can be empty.
4. Prefer "stated" whenever possible. Use "calculated" when basic arithmetic is obvious. Use "estimated" sparingly — and when you do, be honest about it in the note.
5. Never fabricate a specific number with source = "stated". That's a lie. If the rep didn't give you the number, mark it "estimated" or "calculated" and show your work, or use qualitative.
6. Do not invent capabilities, products, steps, agents, bots, or humans that aren't described in the notes. Anti-fabrication still applies to NON-numeric fields.
7. Encourage the rep to go find real numbers — every estimated/calculated outcome will show the derivation note in the UI so they can verify or replace it.

Rules (follow strictly):
- TITLE: SHORT and punchy. 3-6 words. Just the subject of the use case (e.g. "Medical claim adjudication", "Automating denials intake", "KYC refresh triage"). Do NOT include dollar amounts, percentages, or outcomes in the title - those go in the outcomes tiles. Default to the cleanest possible short title.
- SUBTITLE: one line, 8-14 words, summarizing the actor orchestration (e.g. "Agents reason over denials; bots execute; humans approve the edge cases.")
- PROBLEM DESCRIPTION: 2 concise sentences describing the pain. Professional tone. Max ~240 characters.
- PROBLEM STATS: up to 4 quantitative problem metrics (dollars, percentages, volumes, time). Each item: {value, label, source, note}. See transparency rules above. Prefer stated numbers. Empty array if you can't produce even one honestly.
- SOLUTION DESCRIPTION: 1-2 concise sentences describing how agents, bots, and humans resolve the problem. Max ~220 characters.
- CAPABILITIES: 3-6 UiPath product names/capabilities used. Canonical names (use these exactly):
    - "Agents" - AI reasoning, autonomous decisioning
    - "Maestro" - the orchestration layer. ALWAYS include Maestro when the solution involves coordinating agents, bots, or humans across a process, a workflow engine, process/workload orchestration, routing, handoffs, SLA management, or anything the notes call "workload registration", "workflow registration", "orchestration", or "process controller". Default to including Maestro on any multi-step agentic solution.
    - "IXP" - the umbrella term for intelligent document + communications processing. IXP includes Document Understanding AND Communications Mining. ALWAYS use "IXP" — never list "Doc Understanding", "Document Understanding", or "Communications Mining" as separate capabilities. Whenever the notes mention parsing PDFs, emails, forms, invoices, claims, faxes, images, chat transcripts, or call transcripts, that's IXP.
    - "Unattended Robots" - deterministic RPA, system-to-system automation
    - "Attended Robots" - desktop assistant bots with humans
    - "Action Center" - human-in-the-loop approvals and reviews
    - "Data Service" - structured data storage
    - "API Integration" - integrations to external systems
    - "Test Suite", "Insights", "Apps" - only if clearly referenced
  If notes mention orchestration/routing/handoff/workflow engine: include Maestro. If notes mention any document/form/invoice/email extraction: include IXP (not Doc Understanding).
- STEPS: 3-9 discrete steps that are actually described in the notes. Each has a role:
  - "AGENT" = AI reasoning, classification, routing, decisions
  - "BOT"   = deterministic RPA (data entry, API calls, portal polling, system updates)
  - "HUMAN" = human in the loop (review, approve, sign-off)
  - "IXP"   = documents and communications processing — the umbrella for Document Understanding AND Communications Mining. Use IXP for any step that parses PDFs, emails, forms, invoices, claims, faxes, images, chat transcripts, or call transcripts. Do NOT use BOT for these.
  Step description: 3-6 words, imperative. When a step orchestrates or routes work across agents/bots/humans, say Maestro. DO NOT invent steps, agents, bots, humans, or IXP steps that the notes don't mention. If the notes only describe 3 steps, return 3 steps.
- OUTCOMES: 1-5 outcome tiles. Each item: {value, label, source, note}. See transparency rules above. Value examples: "$558K", "90%", "9 min", or qualitative "Reduced" / "Faster" / "Fewer" / "Yes". Label examples: "revenue released", "cycle time", "of workflow automated". Empty array if the notes truly have no outcomes.
- ATTRIBUTABLE IMPACT (optional): list of directional metrics directly moved. Each item: {"direction": "up" | "down", "text": "metric name"}. 3-5 items. Empty list if none clearly inferable from the notes. Use "down" for reductions (cycle time, touches, backlogs) and "up" for improvements (yield, satisfaction, throughput). Only list metrics that the notes actually discuss or clearly imply — do NOT invent metrics the rep didn't mention.
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
  "problem_stats": [{"value": "string", "label": "string", "source": "stated|calculated|estimated|qualitative", "note": "string (required if calculated/estimated, else optional)"}],
  "solution_desc": "string",
  "capabilities": ["string", "..."],
  "steps": [{"role": "AGENT|BOT|HUMAN", "description": "short step name"}],
  "outcomes": [{"value": "string", "label": "string", "source": "stated|calculated|estimated|qualitative", "note": "string (required if calculated/estimated, else optional)"}],
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
- "IXP"   = documents and communications processing (extracting data from PDFs, emails, forms, invoices, claims, faxes, images). Use IXP — not BOT — for any step that parses unstructured documents.

Step description: 3-6 words, imperative (e.g. "Pull & classify inbound denials", "Extract fields via IXP", "Route via Maestro"). Use Maestro for orchestration, routing, or handoff steps.

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


def _extract_text_from_upload(filename, blob):
    """Best-effort text extraction from .pdf, .docx, .txt, .md."""
    name = (filename or '').lower()
    if name.endswith('.pdf'):
        try:
            import io as _io
            from pypdf import PdfReader
            reader = PdfReader(_io.BytesIO(blob))
            return '\n\n'.join((p.extract_text() or '') for p in reader.pages).strip()
        except Exception as e:
            return f"[PDF parse error: {e}]"
    if name.endswith('.docx'):
        try:
            import io as _io
            import docx
            d = docx.Document(_io.BytesIO(blob))
            parts = [p.text for p in d.paragraphs if p.text.strip()]
            for tbl in d.tables:
                for row in tbl.rows:
                    parts.append(' | '.join(c.text for c in row.cells))
            return '\n'.join(parts).strip()
        except Exception as e:
            return f"[DOCX parse error: {e}]"
    if name.endswith('.txt') or name.endswith('.md'):
        try:
            return blob.decode('utf-8', errors='replace').strip()
        except Exception as e:
            return f"[Text decode error: {e}]"
    return f"[Unsupported file type: {filename}]"


@app.route('/parse-docs', methods=['POST'])
def parse_docs():
    """Accept multiple uploaded files, return concatenated extracted text per file."""
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify(error='No files uploaded.'), 400
        results = []
        MAX_BYTES = 10 * 1024 * 1024  # 10 MB per file
        for f in files:
            blob = f.read(MAX_BYTES + 1)
            if len(blob) > MAX_BYTES:
                results.append({'name': f.filename, 'text': f'[File too large: {f.filename}, max 10MB]'})
                continue
            text = _extract_text_from_upload(f.filename, blob)
            # Cap per-file text to keep prompt size reasonable
            if len(text) > 20000:
                text = text[:20000] + '\n[...truncated...]'
            results.append({'name': f.filename, 'text': text})
        return jsonify(files=results)
    except Exception as e:
        traceback.print_exc()
        return jsonify(error=f'Parse failed: {str(e)}'), 500


@app.route('/health')
def health():
    return jsonify(status='ok', has_api_key=bool(os.environ.get('ANTHROPIC_API_KEY', '').strip()))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
