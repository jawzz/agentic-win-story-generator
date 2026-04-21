"""
Agentic Win Story Generator — PPTX Engine
Builds a single-slide agentic use case deck using python-pptx.
Matches the UiPath Agent Use Case Template v10.
"""
from pathlib import Path
import io
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def C(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


DARK = {
    'BG':C('14222A'),'CARD_BG':C('1E3038'),'CARD_BG_2':C('253A45'),'CARD_BG_3':C('2E4552'),
    'DIVIDER':C('2A3F4A'),'ORANGE':C('FA4616'),'ORANGE_CARD':C('7A2E18'),
    'TEAL':C('0BA2B3'),'TEAL_CARD':C('0E5C68'),'GOLD':C('DA9100'),
    'WHITE':C('FFFFFF'),'CREAM':C('FFE8DC'),'TEXT':C('FFFFFF'),
    'TEXT_MUTED':C('8AABB5'),'TEXT_DIM':C('5C7480'),
}
LIGHT = {
    'BG':C('F7F8FA'),'CARD_BG':C('FFFFFF'),'CARD_BG_2':C('EDEFF2'),'CARD_BG_3':C('FFFFFF'),
    'DIVIDER':C('D8DDE3'),'ORANGE':C('FA4616'),'ORANGE_CARD':C('FA4616'),
    'TEAL':C('0BA2B3'),'TEAL_CARD':C('0BA2B3'),'GOLD':C('DA9100'),
    'WHITE':C('FFFFFF'),'CREAM':C('FFE8DC'),'TEXT':C('1A2330'),
    'TEXT_MUTED':C('5A6B78'),'TEXT_DIM':C('8A98A5'),
}

F = "Poppins"
BASE_DIR = Path(__file__).parent
LOGO_WHITE = BASE_DIR / 'static' / 'uipath_logo_2400.png'
LOGO_DARK = BASE_DIR / 'static' / 'uipath_logo_2400_dark.png'


# ---------- shape helpers ----------
def _rect(s, x, y, w, h, fill):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background(); sh.shadow.inherit = False
    return sh


def _rect_bordered(s, x, y, w, h, fill, border, bw=0.5):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = border; sh.line.width = Pt(bw)
    sh.shadow.inherit = False
    return sh


def _pill(s, x, y, w, h, fill):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.adjustments[0] = 0.5
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background(); sh.shadow.inherit = False
    return sh


def _chevron(s, x, y, w, h, fill):
    sh = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background(); sh.shadow.inherit = False
    return sh


def _text(s, x, y, w, h, text, *, size=12, bold=False, italic=False, color=None,
          align='left', anchor='top', tracking=None, font=F):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    if anchor == 'middle': tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    if anchor == 'bottom': tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = tf.paragraphs[0]
    if align == 'center': p.alignment = PP_ALIGN.CENTER
    if align == 'right': p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = text
    r.font.name = font; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    if color is not None: r.font.color.rgb = color
    if tracking is not None:
        rPr = r._r.get_or_add_rPr()
        rPr.set('spc', str(int(tracking * 100)))
    return tb


def _mixed(s, x, y, w, h, parts, *, align='left', anchor='top'):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    if anchor == 'middle': tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    if align == 'center': p.alignment = PP_ALIGN.CENTER
    if align == 'right': p.alignment = PP_ALIGN.RIGHT
    for part in parts:
        r = p.add_run(); r.text = part['text']
        r.font.name = part.get('font', F)
        r.font.size = Pt(part.get('size', 12))
        r.font.bold = part.get('bold', False)
        r.font.italic = part.get('italic', False)
        if 'color' in part: r.font.color.rgb = part['color']
        if part.get('tracking') is not None:
            rPr = r._r.get_or_add_rPr()
            rPr.set('spc', str(int(part['tracking'] * 100)))
    return tb


# ---------- main build ----------
def _build_slide(slide, *, theme, data):
    T = theme
    is_light = theme is LIGHT
    logo_path = str(LOGO_DARK if is_light else LOGO_WHITE)

    # Short title heuristic — tighter top zone + more body space when short
    title_str = data.get('title', '')
    is_short_title = len(title_str) < 50
    if is_short_title:
        divider_y = 1.50
        row_y = 1.65
        row_h = 1.95
        orch_ch = 1.45
        step_h = 0.95
    else:
        divider_y = 1.90
        row_y = 2.05
        row_h = 1.85
        orch_ch = 1.30
        step_h = 0.80
    gap = 0.30
    col_w = (12.33 - gap) / 2

    # Background + top orange accent
    _rect(slide, 0, 0, 13.333, 7.5, T['BG'])
    _rect(slide, 0.5, 0.18, 0.5, 0.05, T['ORANGE'])

    # Breadcrumb + logo
    bc = data.get('breadcrumb', ('Industry', 'Function', 'Use case name'))
    _mixed(slide, 0.5, 0.32, 9.0, 0.28, [
        {'text': bc[0], 'size': 10, 'color': T['TEXT_MUTED'], 'bold': True},
        {'text': '   /   ', 'size': 10, 'color': T['TEXT_DIM']},
        {'text': bc[1], 'size': 10, 'color': T['TEXT_MUTED'], 'bold': True},
        {'text': '   /   ', 'size': 10, 'color': T['TEXT_DIM']},
        {'text': bc[2], 'size': 10, 'color': T['TEXT'], 'bold': True},
    ])
    slide.shapes.add_picture(logo_path, Inches(11.85), Inches(0.25), height=Inches(0.40))

    # Title + subtitle (adaptive)
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.60), Inches(9.0), Inches(1.20))
    tf = tb.text_frame
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    p1 = tf.paragraphs[0]; p1.space_after = Pt(4)
    r1 = p1.add_run(); r1.text = title_str
    r1.font.name = F; r1.font.size = Pt(26); r1.font.bold = True
    r1.font.color.rgb = T['TEXT']
    p2 = tf.add_paragraph()
    r2 = p2.add_run(); r2.text = data.get('subtitle', '')
    r2.font.name = F; r2.font.size = Pt(13)
    r2.font.color.rgb = T['TEXT_MUTED']

    # Company name (right aligned)
    _text(slide, 9.7, 1.05, 3.2, 0.6, data.get('company', ''),
          size=22, bold=True, color=T['TEXT'], align='right', anchor='middle')

    _rect(slide, 0.5, divider_y, 12.33, 0.015, T['DIVIDER'])

    # PROBLEM card
    px = 0.5
    _rect(slide, px, row_y, col_w, row_h, T['ORANGE_CARD'])
    _text(slide, px+0.35, row_y+0.22, col_w-0.7, 0.42,
          'The problem', size=18, bold=True, color=T['WHITE'])

    problem_desc = data.get('problem_desc', '')
    problem_stats = data.get('problem_stats') or []  # list of (num, label) or dicts with 'value', 'label'
    # Normalize stats format
    norm_stats = []
    for st in problem_stats:
        if isinstance(st, dict):
            norm_stats.append((st.get('value') or st.get('number') or '', st.get('label') or ''))
        else:
            norm_stats.append((st[0], st[1]))

    if norm_stats:
        _text(slide, px+0.35, row_y+0.72, col_w-0.7, 0.40,
              problem_desc, size=10, color=T['WHITE'])
        stats_area_h = 0.58
        stats_y = row_y + row_h - 0.12 - stats_area_h
        sa_w = col_w - 0.7
        sc = min(len(norm_stats), 4)
        sw = (sa_w - 0.15*(sc-1)) / sc
        for i, (n, l) in enumerate(norm_stats[:sc]):
            sx = px + 0.35 + i*(sw+0.15)
            _text(slide, sx, stats_y, sw, 0.34, n, size=22, bold=True, color=T['WHITE'])
            _text(slide, sx, stats_y+0.36, sw, 0.26, l, size=9, color=T['CREAM'])
    else:
        _text(slide, px+0.35, row_y+0.72, col_w-0.7, row_h-0.85,
              problem_desc, size=13, color=T['WHITE'])

    # SOLUTION card (with auto-wrap pills)
    sx0 = px + col_w + gap
    _rect(slide, sx0, row_y, col_w, row_h, T['TEAL_CARD'])
    _text(slide, sx0+0.35, row_y+0.22, col_w-0.7, 0.42,
          'The solution', size=18, bold=True, color=T['WHITE'])

    pills = data.get('capabilities') or []
    max_pill_area_w = col_w - 0.7
    rows = [[]]; row_w = 0.0
    pill_h = 0.28; pill_row_gap = 0.06
    for L in pills:
        w = 0.22 + len(L)*0.075
        if rows[-1] and row_w + w > max_pill_area_w:
            rows.append([]); row_w = 0.0
        rows[-1].append((L, w)); row_w += w + 0.08
    n_rows = len(rows) if pills else 0
    pill_block_h = n_rows*pill_h + max(0, n_rows-1)*pill_row_gap if pills else 0
    bottom_margin = 0.10 if n_rows >= 2 else 0.20
    pill_block_bottom = row_y + row_h - bottom_margin
    pill_block_top = pill_block_bottom - pill_block_h
    desc_top = row_y + 0.72
    if pills:
        desc_h = max(0.30, pill_block_top - 0.15 - desc_top)
    else:
        desc_h = row_h - 0.85
    desc_size = 10 if n_rows >= 2 else 11
    _text(slide, sx0+0.35, desc_top, col_w-0.7, desc_h,
          data.get('solution_desc', ''), size=desc_size, color=T['WHITE'])

    if pills:
        pill_fill = T['WHITE']
        pill_text = C('0E5C68')
        for ri, row in enumerate(rows):
            py_row = pill_block_top + ri*(pill_h + pill_row_gap)
            pxp = sx0 + 0.35
            for L, w in row:
                _pill(slide, pxp, py_row, w, pill_h, pill_fill)
                _text(slide, pxp, py_row, w, pill_h, L,
                      size=9, bold=True, color=pill_text, align='center', anchor='middle')
                pxp += w + 0.08

    # What the automation does (orchestration)
    cy = row_y + row_h + 0.10
    ch = orch_ch
    _rect(slide, 0.5, cy, 12.33, ch, T['CARD_BG_2'])
    _text(slide, 0.7, cy+0.12, 8.0, 0.30,
          'What the automation does', size=15, bold=True, color=T['TEXT'])

    steps = data.get('steps') or []
    # Normalize steps: each item dict with 'role' in {agent, bot, human} and 'description'
    norm_steps = []
    for st in steps:
        if isinstance(st, dict):
            role = (st.get('role') or st.get('type') or 'AGENT').upper()
            if role in ('ROBOT',): role = 'BOT'
            if role in ('PERSON',): role = 'HUMAN'
            if role not in ('AGENT', 'BOT', 'HUMAN'):
                role = 'AGENT'
            desc = st.get('description') or st.get('name') or ''
            norm_steps.append((role, desc))
        else:
            norm_steps.append((str(st[0]).upper(), str(st[1])))

    role_c = {'AGENT': T['ORANGE'], 'BOT': T['TEAL'], 'HUMAN': T['GOLD']}
    n_steps = len(norm_steps)

    if n_steps > 0:
        # Adaptive sizing
        if n_steps <= 5:
            arrow_w, gap_s, show_num = 0.22, 0.10, True
        elif n_steps <= 6:
            arrow_w, gap_s, show_num = 0.18, 0.08, True
        elif n_steps <= 7:
            arrow_w, gap_s, show_num = 0.16, 0.07, True
        else:
            arrow_w, gap_s, show_num = 0.14, 0.06, False

        step_y = cy + 0.45
        inner_pad = 0.30
        avail_w = 12.33 - 2*inner_pad
        tile_w = (avail_w - arrow_w*(n_steps-1) - gap_s*(n_steps-1)) / n_steps

        def desc_size_for_step(desc, tw):
            avail_text_w = tw - 0.28
            chars_at_10 = int(avail_text_w / 0.08)
            chars_at_9 = int(avail_text_w / 0.072)
            if len(desc) <= chars_at_10 * 3: return 10
            if len(desc) <= chars_at_9 * 3: return 9
            return 8

        x = 0.5 + inner_pad
        for i, (r, desc) in enumerate(norm_steps):
            if is_light:
                _rect_bordered(slide, x, step_y, tile_w, step_h, T['CARD_BG_3'], T['DIVIDER'], 0.75)
            else:
                _rect(slide, x, step_y, tile_w, step_h, T['CARD_BG_3'])
            _rect(slide, x, step_y, tile_w, 0.05, role_c[r])
            label = f"{i+1:02d}  {r}" if show_num else r
            _text(slide, x+0.14, step_y+0.12, tile_w-0.24, 0.22,
                  label, size=9, bold=True, color=role_c[r], tracking=1.5)
            ds = desc_size_for_step(desc, tile_w)
            _text(slide, x+0.14, step_y+0.34, tile_w-0.24, step_h-0.38,
                  desc, size=ds, bold=True, color=T['TEXT'])
            x += tile_w
            if i < n_steps - 1:
                ax = x + gap_s/3
                ay = step_y + step_h/2 - arrow_w/2
                _chevron(slide, ax, ay, arrow_w, arrow_w, T['TEXT_MUTED'])
                x += arrow_w + gap_s*2/3

    # Measured outcomes
    oy = cy + ch + 0.10
    _text(slide, 0.5, oy, 8.0, 0.30,
          'Measured outcomes', size=15, bold=True, color=T['TEXT'])

    outcomes = data.get('outcomes') or []
    norm_outcomes = []
    for o in outcomes:
        if isinstance(o, dict):
            norm_outcomes.append((o.get('value') or o.get('number') or '',
                                   o.get('label') or ''))
        else:
            norm_outcomes.append((o[0], o[1]))

    if norm_outcomes:
        nt = len(norm_outcomes)
        tile_y = oy + 0.38
        tile_h_out = 0.85
        tg = 0.20
        tw2 = (12.33 - tg*(nt-1)) / nt
        for i, (num, label) in enumerate(norm_outcomes):
            tx = 0.5 + i*(tw2 + tg)
            _rect(slide, tx, tile_y, tw2, tile_h_out, T['ORANGE'])
            _text(slide, tx+0.32, tile_y+0.12, tw2-0.5, 0.50,
                  num, size=32, bold=True, color=T['WHITE'])
            _text(slide, tx+0.32, tile_y+0.60, tw2-0.5, 0.22,
                  label, size=11, color=T['CREAM'])

    # Impact cards: Attributable + Downstream (both optional)
    ay = oy + 1.25
    ah = 0.65
    acol_w = (12.33 - gap) / 2

    attributable = (data.get('attributable') or '').strip()
    downstream = (data.get('downstream') or '').strip()

    # Only render impact cards if at least one present
    if attributable or downstream:
        # Attributable (left)
        if attributable:
            if is_light:
                _rect_bordered(slide, 0.5, ay, acol_w, ah, T['CARD_BG'], T['DIVIDER'], 0.75)
            else:
                _rect(slide, 0.5, ay, acol_w, ah, T['CARD_BG'])
            _mixed(slide, 0.5+0.2, ay+0.12, acol_w-0.4, 0.26, [
                {'text': 'Attributable impact', 'size': 12, 'bold': True, 'color': T['TEAL']},
                {'text': '   directly moved, not yet quantified',
                 'size': 9, 'italic': True, 'color': T['TEXT_DIM']},
            ])
            metric_size = 9 if len(attributable) > 80 else 10
            _text(slide, 0.5+0.2, ay+0.40, acol_w-0.4, 0.22,
                  attributable, size=metric_size, color=T['TEXT'])

        # Downstream (right)
        if downstream:
            dx = 0.5 + acol_w + gap
            if is_light:
                _rect_bordered(slide, dx, ay, acol_w, ah, T['CARD_BG'], T['DIVIDER'], 0.75)
            else:
                _rect(slide, dx, ay, acol_w, ah, T['CARD_BG'])
            _mixed(slide, dx+0.2, ay+0.12, acol_w-0.4, 0.26, [
                {'text': 'Downstream impact', 'size': 12, 'bold': True, 'color': T['GOLD']},
                {'text': '   second-order effects',
                 'size': 9, 'italic': True, 'color': T['TEXT_DIM']},
            ])
            metric_size = 9 if len(downstream) > 80 else 10
            _text(slide, dx+0.2, ay+0.40, acol_w-0.4, 0.22,
                  downstream, size=metric_size, color=T['TEXT'])


def build_pptx(data, template_path=None):
    """Builds a single-slide agentic use case PPTX.

    Args:
        data: dict with keys: breadcrumb (tuple), title, subtitle, company,
              problem_desc, problem_stats, solution_desc, capabilities, steps,
              outcomes, attributable, downstream, theme ("dark" or "light")
        template_path: unused (kept for API compatibility with old signature)

    Returns:
        (pptx_bytes, slide_count)
    """
    theme_name = (data.get('theme') or 'dark').lower()
    theme = LIGHT if theme_name == 'light' else DARK

    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    slide = p.slides.add_slide(p.slide_layouts[6])
    _build_slide(slide, theme=theme, data=data)

    buf = io.BytesIO()
    p.save(buf)
    buf.seek(0)
    return buf.read(), 1
