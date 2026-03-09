#!/usr/bin/env python3
"""股权架构图生成工具 v3 - 支持多层架构/递归子公司/分界线/交叉持股"""
import io, json
from http.server import HTTPServer, BaseHTTPRequestHandler
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from lxml import etree

BLACK = RGBColor(0,0,0)
FONT_SZ = Pt(7)
BOX_BORDER = int(Pt(1).pt * 12700)
LINE_W = int(Pt(0.25).pt * 12700)
NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

BW_MIN, BH_MIN = 1.6, 1.0
BW_PAD, BH_PAD = 0.65, 0.45
LINE_H = 0.40
PW, PH = 1.5, 0.5
SW, SH = 40.9, 19.1
GAP_V = 0.75
GAP_H = 0.5
GAP_H_MIN = 0.25
PCT_OFF = 0.12
NL = "\n"

_CJK = [('\u4e00','\u9fff'),('\u3000','\u303f'),('\uff00','\uffef'),('\u2e80','\u2eff')]
def _is_cjk(c): return any(lo<=c<=hi for lo,hi in _CJK)

def box_width(name):
    lines = [l for l in name.split(NL) if l.strip()]
    if not lines: return BW_MIN
    def lw(l): return sum(0.247 if _is_cjk(c) else 0.135 for c in l)
    return max(BW_MIN, max(lw(l) for l in lines) + BW_PAD)

def box_height(name):
    n = max(1, len(name.split(NL)))
    return max(BH_MIN, n * LINE_H + BH_PAD)

def calc_positions(widths, center=None):
    if center is None: center = SW / 2
    n = len(widths)
    gap = GAP_H
    total = sum(widths) + (n-1)*gap
    usable = SW - 2.4
    if total > usable:
        gap = max(GAP_H_MIN, (usable - sum(widths)) / max(n-1,1))
        total = sum(widths) + (n-1)*gap
    x = center - total/2
    cxs = []
    for w in widths:
        cxs.append(x + w/2); x += w + gap
    return cxs

# ── 递归树布局 ────────────────────────────────────────────────────────────────
def count_leaves(node):
    ch = node.get('children', [])
    return sum(count_leaves(c) for c in ch) if ch else 1

def assign_tree_compact(nodes, center):
    """Position nodes with fixed GAP_H gaps, centered. Children inherit parent center."""
    ws = [box_width(n['name']) for n in nodes]
    cxs = calc_positions(ws, center)
    for i, node in enumerate(nodes):
        node['_cx'] = cxs[i]
        node['_w']  = ws[i]
        node['_h']  = box_height(node['name'])
        if node.get('children'):
            assign_tree_compact(node['children'], cxs[i])

def assign_tree_x(nodes, left, col_w):
    x = left
    for node in nodes:
        leaves = count_leaves(node)
        node['_cx'] = x + leaves * col_w / 2
        node['_w']  = box_width(node['name'])
        node['_h']  = box_height(node['name'])
        if node.get('children'):
            assign_tree_x(node['children'], x, col_w)
        x += leaves * col_w

def get_level_max_h(nodes, level, cur=0):
    mh = BH_MIN
    for node in nodes:
        if cur == level:
            mh = max(mh, box_height(node['name']))
        if node.get('children') and cur < level:
            mh = max(mh, get_level_max_h(node['children'], level, cur+1))
    return mh

def tree_depth(nodes, cur=0):
    if not nodes: return cur
    return max(tree_depth(n.get('children',[]), cur+1) for n in nodes)

# ── PPTX helpers ─────────────────────────────────────────────────────────────
def _rm_shadow(shape):
    spPr = shape._element.find(qn('p:spPr'))
    if spPr is not None and spPr.find(qn('a:effectLst')) is None:
        etree.SubElement(spPr, qn('a:effectLst'))

def _set_border(tb, w, dashed=False, color='000000'):
    sp = tb._element; spPr = sp.find(qn('p:spPr'))
    for e in spPr.findall(qn('a:ln')): spPr.remove(e)
    dash = '<a:prstDash xmlns:a="{}" val="dash"/>'.format(NS) if dashed else ''
    spPr.append(etree.fromstring(
        '<a:ln xmlns:a="{}" w="{}"><a:solidFill><a:srgbClr val="{}"/>'
        '</a:solidFill>{}</a:ln>'.format(NS, w, color, dash)))

def _set_text(tf, lines, align=PP_ALIGN.CENTER):
    bp = tf._txBody.find(qn('a:bodyPr'))
    bp.set('anchor','ctr'); bp.set('lIns','36001'); bp.set('rIns','36001')
    tf.word_wrap = False
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment = align
        if line.strip():
            run = p.add_run()
            run.text = line; run.font.size=FONT_SZ; run.font.color.rgb=BLACK
            run.font.name = 'Times New Roman'
            rPr = run._r.find(qn('a:rPr'))
            if rPr is None: rPr = etree.SubElement(run._r, qn('a:rPr'))
            for ea in rPr.findall(qn('a:ea')): rPr.remove(ea)
            ea = etree.SubElement(rPr, qn('a:ea')); ea.set('typeface','微软雅黑')

def add_box(slide, x, y, w, h, lines, dashed=False, filled=False):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    _set_border(tb, LINE_W if dashed else BOX_BORDER, dashed)
    if filled:
        spPr = tb._element.find(qn('p:spPr'))
        spPr.append(etree.fromstring(
            '<a:solidFill xmlns:a="{}"><a:srgbClr val="D9D9D9"/></a:solidFill>'.format(NS)))
    _set_text(tb.text_frame, lines if lines else [''])
    return tb

def add_label(slide, x, y, w, h, text, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    sp = tb._element; spPr = sp.find(qn('p:spPr'))
    for e in spPr.findall(qn('a:ln')): spPr.remove(e)
    _set_text(tb.text_frame, [text], align=align)
    return tb

def add_line(slide, x1, y1, x2, y2):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1),Cm(y1),Cm(x2),Cm(y2))
    c.line.color.rgb = BLACK; c.line.width = Pt(0.25); _rm_shadow(c); return c

def add_hline_dashed(slide, x1, y, x2):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1),Cm(y),Cm(x2),Cm(y))
    c.line.color.rgb = BLACK; c.line.width = Pt(0.5); _rm_shadow(c)
    ln = c._element.find(qn('p:spPr')).find(qn('a:ln'))
    if ln is not None:
        ln.append(etree.fromstring('<a:prstDash xmlns:a="{}" val="dash"/>'.format(NS)))
    return c

def add_arrow(slide, x, y1, y2):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x),Cm(y1),Cm(x),Cm(y2))
    c.line.color.rgb = BLACK; c.line.width = Pt(0.25); _rm_shadow(c)
    ln = c._element.find(qn('p:spPr')).find(qn('a:ln'))
    if ln is not None:
        ln.append(etree.fromstring('<a:tailEnd xmlns:a="{}" type="triangle" w="sm" len="sm"/>'.format(NS)))
    return c

def add_harrow(slide, x1, y, x2):
    """Horizontal arrow for cross-links"""
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1),Cm(y),Cm(x2),Cm(y))
    c.line.color.rgb = BLACK; c.line.width = Pt(0.25); _rm_shadow(c)
    ln = c._element.find(qn('p:spPr')).find(qn('a:ln'))
    if ln is not None:
        ln.append(etree.fromstring('<a:tailEnd xmlns:a="{}" type="triangle" w="sm" len="sm"/>'.format(NS)))
    return c

# ── Main PPTX generator ──────────────────────────────────────────────────────
def generate_pptx(data):
    prs = Presentation()
    prs.slide_width = Cm(SW); prs.slide_height = Cm(SH)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    shs   = data['shareholders']
    subs  = data['subsidiaries']
    comp  = data.get('companyName','本公司')
    mid   = data.get('middleTier', [])   # 中间控股层
    div   = data.get('divider', None)    # 分界线
    xlink = data.get('crossLinks', [])   # 交叉持股

    R1Y = 1.4

    # ── 股东层 ──────────────────────────────────────────────────────────────
    sh_w = [box_width(s['name']) for s in shs]
    sh_h = [box_height(s['name']) for s in shs]
    sh_cx = calc_positions(sh_w)
    msh = max(sh_h)

    # Snap lines that are very close to center or to each other
    SNAP = 0.8
    snapped_cx = list(sh_cx)
    for i in range(len(snapped_cx)):
        # snap to center
        if abs(snapped_cx[i] - SW/2) < SNAP:
            snapped_cx[i] = SW/2
        # snap to neighbor
        for j in range(i):
            if abs(snapped_cx[i] - snapped_cx[j]) < SNAP:
                snapped_cx[i] = snapped_cx[j]
    sh_cx = snapped_cx

    for i, sh in enumerate(shs):
        add_box(slide, sh_cx[i]-sh_w[i]/2, R1Y, sh_w[i], msh, sh['name'].split(NL))

    ctrl = [i for i,s in enumerate(shs) if s.get('isControl')]
    if ctrl:
        lx = sh_cx[ctrl[0]]  - sh_w[ctrl[0]]/2  - 0.3
        rx = sh_cx[ctrl[-1]] + sh_w[ctrl[-1]]/2 + 0.3
        add_box(slide, lx, R1Y-0.3, rx-lx, msh+0.6, [], dashed=True)

    CCX = SW / 2
    cur_y = R1Y + msh  # bottom of shareholder row

    # ── 中间控股层 ────────────────────────────────────────────────────────────
    if mid:
        mid_w = [box_width(m['name']) for m in mid]
        mid_h = [box_height(m['name']) for m in mid]
        mid_cx = calc_positions(mid_w)
        mmid = max(mid_h)
        MID_Y = cur_y + GAP_V + GAP_V

        for i, m in enumerate(mid):
            add_box(slide, mid_cx[i]-mid_w[i]/2, MID_Y, mid_w[i], mmid, m['name'].split(NL))
            # 连线到本公司 merge point
            add_line(slide, mid_cx[i], MID_Y, mid_cx[i], MID_Y - GAP_V)
            # label
            add_label(slide, mid_cx[i]+PCT_OFF, MID_Y - GAP_V/2 - PH/2, PW, PH, m.get('mainPct',''))
            # 从股东连线到中间层
            parents = m.get('parentShareholders', [])
            for pi in parents:
                if 0 <= pi < len(sh_cx):
                    px = sh_cx[pi]
                    add_line(slide, px, cur_y, px, MID_Y)
                    add_line(slide, px, MID_Y, mid_cx[i], MID_Y)

        # 对没有进中间层的股东，直接画竖线到汇聚线
        mid_parents_flat = set(pi for m in mid for pi in m.get('parentShareholders',[]))
        MERGE_Y = MID_Y - GAP_V
        for i, sh in enumerate(shs):
            if i not in mid_parents_flat:
                add_line(slide, sh_cx[i], cur_y, sh_cx[i], MERGE_Y)
                add_label(slide, sh_cx[i]+PCT_OFF, cur_y+(MERGE_Y-cur_y)/2-PH/2, PW, PH, sh['pct'])

        # 中间层汇聚线
        all_x = [sh_cx[i] for i in range(len(shs)) if i not in mid_parents_flat] + list(mid_cx)
        if len(all_x) > 1:
            add_line(slide, min(all_x), MERGE_Y, max(all_x), MERGE_Y)
        add_arrow(slide, CCX, MERGE_Y, MERGE_Y + GAP_V)
        CY = MERGE_Y + GAP_V
    else:
        # 标准：股东直接连本公司
        MERGE_Y = cur_y + GAP_V
        CY      = MERGE_Y + GAP_V
        for i, sh in enumerate(shs):
            add_line(slide, sh_cx[i], cur_y, sh_cx[i], MERGE_Y)
            add_label(slide, sh_cx[i]+PCT_OFF, cur_y+(MERGE_Y-cur_y)/2-PH/2, PW, PH, sh['pct'])
        add_line(slide, sh_cx[0], MERGE_Y, sh_cx[-1], MERGE_Y)
        add_arrow(slide, CCX, MERGE_Y, CY)

    # ── 本公司 ────────────────────────────────────────────────────────────────
    comp_w = box_width(comp); comp_h = box_height(comp)
    filled = data.get('companyFilled', False)
    add_box(slide, CCX-comp_w/2, CY, comp_w, comp_h, [comp], filled=filled)

    # ── 子公司（递归树） ──────────────────────────────────────────────────────
    if not subs:
        buf = io.BytesIO(); prs.save(buf); return buf.getvalue()

    # Compact layout: use fixed GAP_H between boxes, centered on slide
    total_lv = count_leaves({'children': subs}) if subs else 1
    avail_w  = SW - 2.4
    # Use min of: fixed-gap width OR equal-column width (when tree is wide)
    sub_leaf_w = [box_width(s['name']) for s in subs]
    natural_w  = sum(sub_leaf_w) + (len(subs)-1) * GAP_H
    col_w = avail_w / total_lv if natural_w > avail_w else None
    if col_w:
        assign_tree_x(subs, (SW - avail_w)/2, col_w)
    else:
        assign_tree_compact(subs, SW/2)
    # Snap top-level sub lines close to center or each other
    SNAP_S = 0.8
    for i in range(len(subs)):
        if abs(subs[i]['_cx'] - SW/2) < SNAP_S: subs[i]['_cx'] = SW/2
        for j in range(i):
            if abs(subs[i]['_cx'] - subs[j]['_cx']) < SNAP_S: subs[i]['_cx'] = subs[j]['_cx']

    depth = tree_depth(subs)

    # Pre-compute max height at each depth level
    level_h = [get_level_max_h(subs, lv) for lv in range(depth)]

    # Y positions for each level
    level_y = []
    y = CY + comp_h + GAP_V
    SM = y  # merge line for level 0
    y += GAP_V
    level_y.append(y)
    for lv in range(1, depth):
        y += level_h[lv-1] + GAP_V + GAP_V
        level_y.append(y)

    # Divider line
    div_y = None
    if div:
        after_lv = div.get('afterLevel', 0)
        if after_lv < len(level_y):
            div_y = level_y[after_lv] + level_h[after_lv] + GAP_V * 0.6

    # Draw merge line from company to subs
    add_line(slide, CCX, CY+comp_h, CCX, SM)
    if len(subs) > 1:
        leftmost  = min(n['_cx'] for n in subs)
        rightmost = max(n['_cx'] for n in subs)
        add_line(slide, leftmost, SM, rightmost, SM)

    # Track box positions for cross-links
    box_pos = {}

    def draw_tree(nodes, parent_cx, parent_sm, level):
        mh = level_h[level]
        SY = level_y[level]
        for node in nodes:
            cx = node['_cx']; w = node['_w']
            add_arrow(slide, cx, parent_sm, SY)
            mid_pct = parent_sm + (SY - parent_sm)/2 - PH/2
            add_label(slide, cx+PCT_OFF, mid_pct, PW, PH, node['pct'])
            add_box(slide, cx-w/2, SY, w, mh, node['name'].split(NL))
            box_pos[node['name'].replace(NL,'')] = {'cx': cx, 'w': w, 'y': SY, 'h': mh}
            if node.get('children'):
                ch = node['children']
                ch_lx = min(c['_cx'] for c in ch)
                ch_rx = max(c['_cx'] for c in ch)
                ch_sm = SY + mh + GAP_V
                add_line(slide, cx, SY+mh, cx, ch_sm)
                if len(ch) > 1:
                    add_line(slide, ch_lx, ch_sm, ch_rx, ch_sm)
                # Divider between this level and next?
                if div_y and level == div.get('afterLevel', 0):
                    pass  # drawn globally below
                draw_tree(ch, cx, ch_sm, level+1)

    draw_tree(subs, CCX, SM, 0)

    # Draw divider line
    if div_y:
        add_hline_dashed(slide, 1.0, div_y, SW-1.0)
        ll = div.get('leftLabel','境外')
        rl = div.get('rightLabel','境内')
        add_label(slide, 1.2, div_y-0.35, 2.0, 0.4, ll)
        add_label(slide, SW-3.2, div_y+0.05, 2.0, 0.4, rl)

    # Draw cross-links
    for xl in xlink:
        fn = xl.get('from','').replace(NL,''); tn = xl.get('to','').replace(NL,'')
        if fn in box_pos and tn in box_pos:
            fp = box_pos[fn]; tp = box_pos[tn]
            # Route below boxes: down → across → up into target box bottom
            y_bot = max(fp['y']+fp['h'], tp['y']+tp['h']) + 0.4
            x1 = fp['cx']
            x2 = tp['cx']
            add_line(slide, x1, fp['y']+fp['h'], x1, y_bot)
            add_line(slide, x1, y_bot, x2, y_bot)
            add_arrow(slide, x2, y_bot, tp['y']+tp['h'])
            mid_x = (x1+x2)/2
            add_label(slide, mid_x-PW/2, y_bot+0.05, PW, PH, xl.get('pct',''))

    buf = io.BytesIO(); prs.save(buf); return buf.getvalue()


# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"/><title>股权架构图生成工具 v3</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;700&family=Noto+Sans+SC:wght@400;500;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Noto Sans SC',system-ui,-apple-system,sans-serif;background:#f0f2f8;}
header{background:#1a2640;color:#fff;padding:14px 28px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:10;}
header h1{font-size:20px;letter-spacing:1px;font-weight:700;}
header p{font-size:12px;color:#8fa0b8;margin-top:2px;}
.page{max-width:960px;margin:0 auto;padding:20px 16px 48px;}
.card{background:#fff;border-radius:8px;box-shadow:0 1px 6px rgba(0,0,0,.08);padding:16px 18px;margin-bottom:14px;}
.card-title{font-size:15px;font-weight:700;color:#1a2640;border-left:3px solid #1a2640;padding-left:9px;margin-bottom:12px;}
.hint{font-size:13px;color:#999;margin-bottom:10px;line-height:1.7;}
.row{display:flex;gap:6px;align-items:flex-start;margin-bottom:8px;}
.row-num{width:18px;font-size:11px;color:#ccc;padding-top:7px;flex-shrink:0;}
textarea.ni{flex:1;padding:5px 8px;border:1px solid #ddd;border-radius:4px;font-size:13px;outline:none;font-family:inherit;resize:none;overflow:hidden;line-height:1.5;min-height:30px;}
textarea.ni:focus{border-color:#1a2640;}
input.pi{flex:0 0 78px;padding:6px 8px;border:1px solid #ddd;border-radius:4px;font-size:13px;outline:none;font-family:inherit;}
input.pi:focus{border-color:#1a2640;}
input.wi{width:260px;padding:6px 9px;border:1px solid #ddd;border-radius:4px;font-size:13px;outline:none;font-family:inherit;}
input.si{width:120px;padding:5px 8px;border:1px solid #ddd;border-radius:4px;font-size:12px;outline:none;font-family:inherit;}
input.si:focus{border-color:#1a2640;}
label.ctrl{font-size:11px;color:#555;display:flex;align-items:center;gap:3px;cursor:pointer;white-space:nowrap;flex-shrink:0;padding-top:7px;}
button.rm{padding:4px 7px;border:1px solid #e0e0e0;border-radius:4px;cursor:pointer;background:#fff;font-size:11px;color:#bbb;flex-shrink:0;margin-top:3px;}
button.rm:hover{background:#fff0f0;color:#c33;}
button.add-ch{padding:3px 8px;border:1px solid #9ab0e0;border-radius:4px;cursor:pointer;background:#e8f4ff;font-size:11px;color:#2477cc;flex-shrink:0;margin-top:3px;}
.sub-card{background:#f9f9fb;border:1px solid #eaeaee;border-radius:6px;padding:9px 11px;margin-bottom:8px;}
.ch-indent{margin-left:20px;margin-top:6px;border-left:2px solid #e0e4f0;padding-left:10px;}
button.add-btn{width:100%;margin-top:4px;padding:9px;border:1px dashed #9ab0e0;border-radius:5px;background:#f0f5ff;color:#2455aa;font-size:13px;cursor:pointer;font-family:inherit;}
button.add-btn:hover{background:#e0ecff;}
.section-toggle{font-size:14px;color:#2477cc;cursor:pointer;margin-left:8px;font-weight:400;}
.extra-section{display:none;margin-top:10px;padding-top:10px;border-top:1px solid #eee;}
.gen-wrap{text-align:center;margin:6px 0 18px;}
#genBtn{padding:13px 52px;background:#1a2640;color:#fff;border:none;border-radius:7px;font-size:16px;font-weight:700;cursor:pointer;letter-spacing:1px;box-shadow:0 3px 12px rgba(26,38,64,.3);font-family:inherit;}
#genBtn:hover{background:#243560;}
#genBtn:disabled{background:#8899bb;cursor:not-allowed;}
#status{margin-top:9px;font-size:12px;min-height:18px;}
.ok{color:#2a7a2a;}.err{color:#c22;}
.preview-card{background:#fff;border-radius:8px;box-shadow:0 1px 6px rgba(0,0,0,.08);padding:16px 18px;margin-bottom:14px;}
.preview-title{font-size:14px;color:#888;font-weight:400;margin-bottom:12px;display:flex;align-items:center;gap:8px;}
.preview-title strong{color:#1a2640;font-weight:700;}
#svgWrap{width:100%;overflow-x:auto;overflow-y:scroll;height:340px;border:1px solid #e8ecf4;border-radius:5px;background:#fafbff;}
#svgWrap svg{display:block;}
.xl-row{display:flex;gap:6px;align-items:center;margin-bottom:6px;flex-wrap:wrap;}
.xl-row select{padding:5px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px;font-family:inherit;}
</style>
</head>
<body>
<header>
  <div style="font-size:22px">&#9878;</div>
  <div><h1>股权架构图生成工具 v3</h1><p>港股IPO专用 &middot; 支持多层架构 / 递归子公司 / 分界线 / 交叉持股</p></div>
</header>
<div class="page">

<!-- 本公司 -->
<div class="card">
  <div class="card-title">本公司</div>
  <div class="row">
    <input type="text" class="wi" id="companyName" value="本公司" oninput="renderPreview()"/>
    <label class="ctrl" style="margin-left:12px">
      <input type="checkbox" id="companyFilled" onchange="renderPreview()"/> 灰底色（填充）
    </label>
  </div>
</div>

<!-- 股东层 -->
<div class="card">
  <div class="card-title">股东层（第一层）</div>
  <p class="hint">名称直接回车换行。勾选"控股方"加虚线框。</p>
  <div id="shareholderList"></div>
  <button class="add-btn" onclick="addShareholder()">+ 添加股东</button>
</div>

<!-- 中间控股层 -->
<div class="card">
  <div class="card-title">
    中间控股层（可选）
    <span class="section-toggle" onclick="toggleSection('midSection')">▶ 展开</span>
  </div>
  <p class="hint">适用于：股东通过控股公司间接持股本公司的情形（如图4样式）。</p>
  <div id="midSection" class="extra-section">
    <div id="midList"></div>
    <button class="add-btn" onclick="addMid()">+ 添加中间控股公司</button>
  </div>
</div>

<!-- 子公司层 -->
<div class="card">
  <div class="card-title">子公司层（支持多级递归）</div>
  <p class="hint">名称直接回车换行。每个子公司可无限添加下级子公司。</p>
  <div id="subsidiaryList"></div>
  <button class="add-btn" onclick="addSub()">+ 添加子公司</button>
</div>



<!-- 交叉持股 -->
<div class="card">
  <div class="card-title">
    交叉持股箭头（可选）
    <span class="section-toggle" onclick="toggleSection('xlSection')">▶ 展开</span>
  </div>
  <p class="hint">在两个子公司之间画一条横向箭头，标注持股比例（如图5样式）。</p>
  <div id="xlSection" class="extra-section">
    <div id="xlList"></div>
    <button class="add-btn" onclick="addXL()">+ 添加交叉持股</button>
  </div>
</div>

<!-- 生成 -->
<div class="gen-wrap">
  <button id="genBtn" onclick="generate()">&#8595; 生成并下载 PPTX</button>
  <p id="status"></p>
</div>

<!-- 预览 -->
<div class="preview-card">
  <div class="preview-title"><strong>实时预览</strong><span>布局与导出PPT完全一致</span></div>
  <div id="svgWrap"></div>
</div>

</div>
<script>
// ── 数据 ─────────────────────────────────────────────────────────────────────
let shareholders=[
  {name:"王平",          pct:"34.51%", isControl:true},
  {name:"兆格投资",       pct:"8.85%",  isControl:true},
  {name:"我们的董事及\\n高级管理层成员", pct:"0.19%", isControl:false},
  {name:"其他A股股东",    pct:"44.66%", isControl:false},
  {name:"H股股东",       pct:"11.79%", isControl:false},
];
let middleTier=[];  // {name, mainPct, parentShareholders:[idx,...], ownerPcts:[...]}
let subsidiaries=[
  {name:"美格智联\\n（中国）", pct:"100%", children:[]},
  {name:"眾格上海\\n（中国）", pct:"100%", children:[]},
  {name:"西安兆格\\n（中国）", pct:"100%", children:[]},
  {name:"上海美胧\\n（中国）", pct:"100%", children:[]},
  {name:"眾格南通\\n（中国）", pct:"100%", children:[]},
  {name:"美格智投\\n（中国）", pct:"100%", children:[]},
  {name:"方格國際\\n（香港）", pct:"100%", children:[{name:"MeiG Smart\\nTechnology\\nFrance\\n（法国）", pct:"100%", children:[]}]},
  {name:"MeiG Smart\\nTechnology\\n(Europe)\\nGmbH\\n（德国）", pct:"100%", children:[]},
];
let crossLinks=[];  // {from, to, pct}

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function ex(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function autoH(ta){ta.style.height='auto';ta.style.height=(ta.scrollHeight)+'px';}

function toggleSection(id){
  const el=document.getElementById(id);
  const span=el.previousElementSibling;
  if(el.style.display==='block'){el.style.display='none';span.textContent='▶ 展开';}
  else{el.style.display='block';span.textContent='▼ 收起';}
}

// ── 股东渲染 ─────────────────────────────────────────────────────────────────
function renderShareholders(){
  document.getElementById('shareholderList').innerHTML=shareholders.map((sh,i)=>`
    <div class="row">
      <div class="row-num">${i+1}</div>
      <textarea class="ni" id="sh_${i}" oninput="shareholders[${i}].name=this.value;autoH(this);renderPreview()">${esc(sh.name)}</textarea>
      <input class="pi" type="text" value="${esc(sh.pct)}" placeholder="持股%" oninput="shareholders[${i}].pct=this.value;renderPreview()"/>
      <label class="ctrl"><input type="checkbox" ${sh.isControl?'checked':''} onchange="shareholders[${i}].isControl=this.checked;renderPreview()"/> 控股方</label>
      <button class="rm" onclick="shareholders.splice(${i},1);renderShareholders();renderPreview()">&#x2715;</button>
    </div>`).join('');
  shareholders.forEach((_,i)=>{const t=document.getElementById('sh_'+i);if(t)autoH(t);});
}
function addShareholder(){shareholders.push({name:'',pct:'',isControl:false});renderShareholders();renderPreview();}

// ── 中间控股层渲染 ────────────────────────────────────────────────────────────
function renderMid(){
  const shNames=shareholders.map((s,i)=>`<option value="${i}">${esc(s.name.replace(/\\n/g,' '))}</option>`).join('');
  document.getElementById('midList').innerHTML=middleTier.map((m,i)=>`
    <div class="sub-card">
      <div class="row">
        <div class="row-num">${i+1}</div>
        <textarea class="ni" id="mid_${i}" oninput="middleTier[${i}].name=this.value;autoH(this);renderPreview()">${esc(m.name)}</textarea>
        <input class="pi" type="text" value="${esc(m.mainPct||'')}" placeholder="持本公司%" oninput="middleTier[${i}].mainPct=this.value;renderPreview()"/>
        <button class="rm" onclick="middleTier.splice(${i},1);renderMid();renderPreview()">&#x2715;</button>
      </div>
      <div style="font-size:11px;color:#888;margin:4px 0 4px 24px">上级股东（可多选）：
        <select multiple size="3" style="width:200px;font-size:11px;border:1px solid #ddd;border-radius:4px"
          onchange="middleTier[${i}].parentShareholders=[...this.selectedOptions].map(o=>+o.value);renderPreview()">
          ${shareholders.map((s,si)=>`<option value="${si}" ${(m.parentShareholders||[]).includes(si)?'selected':''}>${esc(s.name.replace(/\\n/g,' '))}</option>`).join('')}
        </select>
      </div>
    </div>`).join('');
  middleTier.forEach((_,i)=>{const t=document.getElementById('mid_'+i);if(t)autoH(t);});
}
function addMid(){middleTier.push({name:'',mainPct:'',parentShareholders:[]});renderMid();renderPreview();}

// ── 子公司递归渲染 ────────────────────────────────────────────────────────────
function renderNode(node, path, depth){
  const indent=depth*16;
  const id='n_'+path.join('_');
  return `<div style="margin-bottom:6px">
    <div class="row" style="margin-left:${indent}px">
      <textarea class="ni" id="${id}" oninput="setNodeField('${path.join(',')}','name',this.value);autoH(this);renderPreview()">${esc(node.name)}</textarea>
      <input class="pi" type="text" value="${esc(node.pct)}" placeholder="持股%" oninput="setNodeField('${path.join(',')}','pct',this.value);renderPreview()"/>
      <button class="add-ch" onclick="addChild('${path.join(',')}')">+子公司</button>
      <button class="rm" onclick="removeNode('${path.join(',')}')">&#x2715;</button>
    </div>
    <div class="ch-indent" ${(node.children&&node.children.length)?'':'style="display:none"'} id="ch_${path.join('_')}">
      ${(node.children||[]).map((c,ci)=>renderNode(c,[...path,ci],depth+1)).join('')}
    </div>
  </div>`;
}

function renderSubsidiaries(){
  document.getElementById('subsidiaryList').innerHTML=subsidiaries.map((s,i)=>renderNode(s,[i],0)).join('');
  // autoH all textareas
  document.querySelectorAll('#subsidiaryList textarea').forEach(t=>autoH(t));
}

function getNode(path){
  let arr=subsidiaries, node=null;
  for(let i=0;i<path.length;i++){
    node=arr[path[i]];
    if(i<path.length-1) arr=node.children;
  }
  return node;
}
function setNodeField(pathStr,field,val){
  const path=pathStr.split(',').map(Number);
  const node=getNode(path);
  if(node) node[field]=val;
}
function addChild(pathStr){
  const path=pathStr.split(',').map(Number);
  const node=getNode(path);
  if(node){
    if(!node.children) node.children=[];
    node.children.push({name:'',pct:'100%',children:[]});
    renderSubsidiaries();renderPreview();
  }
}
function removeNode(pathStr){
  const path=pathStr.split(',').map(Number);
  if(path.length===1){subsidiaries.splice(path[0],1);}
  else{
    const parent=getNode(path.slice(0,-1));
    parent.children.splice(path[path.length-1],1);
  }
  renderSubsidiaries();renderXL();renderPreview();
}
function addSub(){subsidiaries.push({name:'',pct:'100%',children:[]});renderSubsidiaries();renderXL();renderPreview();}

// ── 交叉持股渲染 ──────────────────────────────────────────────────────────────
function getAllNodeNames(nodes,acc=[]){
  const _nl=new RegExp(String.fromCharCode(92,110),'g'),_lf=new RegExp(String.fromCharCode(10),'g');
  nodes.forEach(n=>{
    acc.push(n.name.replace(_nl,'').replace(_lf,''));
    if(n.children) getAllNodeNames(n.children,acc);
  });
  return acc;
}
function renderXL(){
  const names=getAllNodeNames(subsidiaries);
  const opts=names.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join('');
  document.getElementById('xlList').innerHTML=crossLinks.map((xl,i)=>`
    <div class="xl-row">
      <span style="font-size:12px;color:#555">从</span>
      <select style="padding:5px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px;font-family:inherit;max-width:180px" onchange="crossLinks[${i}].from=this.value;renderPreview()">
        <option value="">-- 选择公司 --</option>${opts.replace(`value="${esc(xl.from)}"`,`value="${esc(xl.from)}" selected`)}
      </select>
      <span style="font-size:12px;color:#555">→</span>
      <select style="padding:5px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px;font-family:inherit;max-width:180px" onchange="crossLinks[${i}].to=this.value;renderPreview()">
        <option value="">-- 选择公司 --</option>${opts.replace(`value="${esc(xl.to)}"`,`value="${esc(xl.to)}" selected`)}
      </select>
      <input class="pi" type="text" value="${esc(xl.pct)}" placeholder="持股%" oninput="crossLinks[${i}].pct=this.value;renderPreview()"/>
      <button class="rm" onclick="crossLinks.splice(${i},1);renderXL();renderPreview()">&#x2715;</button>
    </div>`).join('');
}
function addXL(){crossLinks.push({from:'',to:'',pct:''});renderXL();}

// ── SVG 预览（与Python完全一致的布局算法） ────────────────────────────────────
const BW_MIN=1.6,BH_MIN=1.0,BW_PAD=0.65,BH_PAD=0.45,LINE_H=0.40;
const SW=40.9,GAP_V=0.75,GAP_H=0.5,GAP_H_MIN=0.25,POFF=0.12,PH=0.5,PW=1.5;
const _CJK=/[\\u4e00-\\u9fff\\u3000-\\u303f\\uff00-\\uffef\\u2e80-\\u2eff]/;
function boxW(name){
  const lines=name.split('\\n').filter(l=>l.trim());
  if(!lines.length)return BW_MIN;
  const lw=l=>[...l].reduce((a,c)=>a+(_CJK.test(c)?0.247:0.135),0);
  return Math.max(BW_MIN,Math.max(...lines.map(lw))+BW_PAD);
}
function boxH(name){return Math.max(BH_MIN,Math.max(1,name.split('\\n').length)*LINE_H+BH_PAD);}

function calcPos(widths,center){
  if(center===undefined)center=SW/2;
  const n=widths.length; let gap=GAP_H;
  const total=widths.reduce((a,w)=>a+w,0)+(n-1)*gap;
  const usable=SW-2.4;
  if(total>usable&&center===SW/2) gap=Math.max(GAP_H_MIN,(usable-widths.reduce((a,w)=>a+w,0))/Math.max(n-1,1));
  let x=center-(widths.reduce((a,w)=>a+w,0)+(n-1)*gap)/2;
  return widths.map(w=>{const cx=x+w/2;x+=w+gap;return cx;});
}

function countLeaves(n){const ch=n.children||[];return ch.length?ch.reduce((a,c)=>a+countLeaves(c),0):1;}
function assignCompact(nodes,center){
  const ws=nodes.map(n=>boxW(n.name));
  const cxs=calcPos(ws,center);
  nodes.forEach((n,i)=>{
    n._cx=cxs[i]; n._w=ws[i]; n._h=boxH(n.name);
    if(n.children&&n.children.length) assignCompact(n.children,cxs[i]);
  });
}
function assignX(nodes,left,colW){
  let x=left;
  nodes.forEach(n=>{
    const lv=countLeaves(n);
    n._cx=x+lv*colW/2; n._w=boxW(n.name); n._h=boxH(n.name);
    if(n.children&&n.children.length) assignX(n.children,x,colW);
    x+=lv*colW;
  });
}
function getLevelH(nodes,level,cur=0){
  let mh=BH_MIN;
  nodes.forEach(n=>{
    if(cur===level)mh=Math.max(mh,boxH(n.name));
    if(n.children&&cur<level)mh=Math.max(mh,getLevelH(n.children,level,cur+1));
  });
  return mh;
}
function treeDepth(nodes,cur=0){
  if(!nodes||!nodes.length)return cur;
  return Math.max(...nodes.map(n=>treeDepth(n.children||[],cur+1)));
}
function getAllNames(nodes,acc=[]){
  nodes.forEach(n=>{acc.push(n.name);if(n.children)getAllNames(n.children,acc);});
  return acc;
}

function renderPreview(){
  const comp=document.getElementById('companyName').value||'本公司';
  const shs=shareholders, subs=subsidiaries, mid=middleTier;

  const compFilled=document.getElementById('companyFilled').checked;
  const divEnabled=false, divAfterLevel=0, divLeft='境外', divRight='境内';
  const n=shs.length;

  const sh_w=shs.map(s=>boxW(s.name)), sh_h=shs.map(s=>boxH(s.name));
  const sh_cx=calcPos(sh_w);
  const msh=Math.max(...sh_h,BH_MIN);
  // Snap lines close to center or each other
  const SNAP=0.8;
  for(let i=0;i<sh_cx.length;i++){
    if(Math.abs(sh_cx[i]-SW/2)<SNAP) sh_cx[i]=SW/2;
    for(let j=0;j<i;j++){
      if(Math.abs(sh_cx[i]-sh_cx[j])<SNAP) sh_cx[i]=sh_cx[j];
    }
  }
  const R1Y=1.4;
  const CCX=SW/2;
  const comp_w=boxW(comp), comp_h=boxH(comp);

  const wrap=document.getElementById('svgWrap');
  const sc=40;
  const p=v=>+(v*sc).toFixed(1);

  const el=[];
  el.push('<defs><marker id="tri" markerWidth="7" markerHeight="7" refX="7" refY="3.5" orient="auto"><polygon points="0,0 7,3.5 0,7" fill="#000"/></marker></defs>');

  let _ci=0;
  const rect=(x,y,w,h,sw=1,dash='',fill='none')=>
    `<rect x="${p(x)}" y="${p(y)}" width="${p(w)}" height="${p(h)}" fill="${fill}" stroke="#000" stroke-width="${sw}" ${dash?'stroke-dasharray="'+dash+'"':''} shape-rendering="crispEdges"/>`;
  const ln=(x1,y1,x2,y2,sw=0.7,dash='')=>
    `<line x1="${p(x1)}" y1="${p(y1)}" x2="${p(x2)}" y2="${p(y2)}" stroke="#000" stroke-width="${sw}" ${dash?'stroke-dasharray="'+dash+'"':''} shape-rendering="crispEdges"/>`;
  const arr=(x,y1,y2)=>
    `<line x1="${p(x)}" y1="${p(y1)}" x2="${p(x)}" y2="${p(y2)}" stroke="#000" stroke-width="0.7" marker-end="url(#tri)" shape-rendering="crispEdges"/>`;
  const harr=(x1,y,x2)=>
    `<line x1="${p(x1)}" y1="${p(y)}" x2="${p(x2)}" y2="${p(y)}" stroke="#000" stroke-width="0.7" marker-end="url(#tri)" shape-rendering="crispEdges"/>`;

  const txt=(x,y,w,h,lines,fs=8)=>{
    const id='cl'+(++_ci);
    const lh=fs*1.5,tot=lines.length*lh,sy=p(y)+p(h)/2-tot/2+fs;
    return `<clipPath id="${id}"><rect x="${p(x+0.05)}" y="${p(y+0.05)}" width="${p(w-0.1)}" height="${p(h-0.1)}"/></clipPath>`
      +lines.map((l,i)=>`<text x="${p(x)+p(w)/2}" y="${+(sy+i*lh).toFixed(1)}" text-anchor="middle" font-size="${fs}" font-family="serif" fill="#000" clip-path="url(#${id})">${ex(l)}</text>`).join('');
  };
  const lbl=(x,y,text,anchor='start')=>
    `<text x="${p(x)}" y="${p(y)+p(PH)/2+3}" text-anchor="${anchor}" font-size="7" font-family="serif" fill="#000">${ex(text)}</text>`;

  // 股东
  const ctrl=shs.map((s,i)=>s.isControl?i:-1).filter(i=>i>=0);
  if(ctrl.length){
    const lx=sh_cx[ctrl[0]]-sh_w[ctrl[0]]/2-0.3,rx=sh_cx[ctrl[ctrl.length-1]]+sh_w[ctrl[ctrl.length-1]]/2+0.3;
    el.push(rect(lx,R1Y-0.3,rx-lx,msh+0.6,0.5,'4,3'));
  }
  shs.forEach((sh,i)=>{
    el.push(rect(sh_cx[i]-sh_w[i]/2,R1Y,sh_w[i],msh,1.2));
    el.push(txt(sh_cx[i]-sh_w[i]/2,R1Y,sh_w[i],msh,sh.name.split('\\n')));
  });

  let cur_y=R1Y+msh, CY;

  // 中间层
  if(mid.length){
    const mid_w=mid.map(m=>boxW(m.name)), mid_h=mid.map(m=>boxH(m.name));
    const mid_cx=calcPos(mid_w);
    const mmid=Math.max(...mid_h,BH_MIN);
    const MID_Y=cur_y+GAP_V+GAP_V;
    const MERGE_Y=MID_Y-GAP_V;
    mid.forEach((m,i)=>{
      el.push(rect(mid_cx[i]-mid_w[i]/2,MID_Y,mid_w[i],mmid,1.2));
      el.push(txt(mid_cx[i]-mid_w[i]/2,MID_Y,mid_w[i],mmid,m.name.split('\\n')));
      el.push(ln(mid_cx[i],MID_Y,mid_cx[i],MERGE_Y));
      el.push(lbl(mid_cx[i]+POFF,MERGE_Y+(MID_Y-MERGE_Y)/2-PH/2,m.mainPct||''));
      (m.parentShareholders||[]).forEach(pi=>{
        if(pi<sh_cx.length){
          el.push(ln(sh_cx[pi],cur_y,sh_cx[pi],MID_Y));
          el.push(ln(sh_cx[pi],MID_Y,mid_cx[i],MID_Y));
        }
      });
    });
    const midParents=new Set(mid.flatMap(m=>m.parentShareholders||[]));
    shs.forEach((_,i)=>{
      if(!midParents.has(i)){
        el.push(ln(sh_cx[i],cur_y,sh_cx[i],MERGE_Y));
        el.push(lbl(sh_cx[i]+POFF,cur_y+(MERGE_Y-cur_y)/2-PH/2,shs[i].pct));
      }
    });
    const allX=[...shs.map((_,i)=>midParents.has(i)?null:sh_cx[i]),...mid_cx].filter(x=>x!==null);
    if(allX.length>1) el.push(ln(Math.min(...allX),MERGE_Y,Math.max(...allX),MERGE_Y));
    el.push(arr(CCX,MERGE_Y,MERGE_Y+GAP_V));
    CY=MERGE_Y+GAP_V;
  } else {
    const MERGE_Y=cur_y+GAP_V, _CY=MERGE_Y+GAP_V;
    shs.forEach((sh,i)=>{
      el.push(ln(sh_cx[i],cur_y,sh_cx[i],MERGE_Y));
      el.push(lbl(sh_cx[i]+POFF,cur_y+(MERGE_Y-cur_y)/2-PH/2,sh.pct));
    });
    if(n>1) el.push(ln(sh_cx[0],MERGE_Y,sh_cx[n-1],MERGE_Y));
    el.push(arr(CCX,MERGE_Y,_CY));
    CY=_CY;
  }

  // 本公司
  const cfill=compFilled?'#D9D9D9':'none';
  el.push(rect(CCX-comp_w/2,CY,comp_w,comp_h,1.2,'',cfill));
  el.push(txt(CCX-comp_w/2,CY,comp_w,comp_h,[comp]));

  if(!subs.length){
    const maxH=CY+comp_h+1;
    const svgW=Math.max(SW*sc,wrap.clientWidth-4);
    wrap.innerHTML=`<svg width="${svgW}" height="${p(maxH)}" xmlns="http://www.w3.org/2000/svg"><g transform="translate(${((svgW-SW*sc)/2).toFixed(1)},0)">${el.join('')}</g></svg>`;
    return;
  }

  // 子公司树
  const totalLeaves=subs.reduce((a,s)=>a+countLeaves(s),0);
  const availW=SW-2.4;
  // Use compact fixed-gap layout when content fits, else equal-column
  const topW=subs.reduce((a,s)=>a+boxW(s.name),0)+(subs.length-1)*GAP_H;
  if(topW>availW){
    const colW=availW/totalLeaves;
    assignX(subs,(SW-availW)/2,colW);
  } else {
    assignCompact(subs,SW/2);
  }

  const depth=treeDepth(subs);
  const levelH=Array.from({length:Math.max(depth,1)},(_,lv)=>getLevelH(subs,lv));
  const levelY=[]; let y=CY+comp_h+GAP_V;
  const SM=y; y+=GAP_V; levelY.push(y);
  // Snap sub lines close to center or each other
  subs.forEach(s=>{
    if(Math.abs(s._cx-SW/2)<SNAP) s._cx=SW/2;
  });
  for(let i=1;i<subs.length;i++){
    for(let j=0;j<i;j++){
      if(Math.abs(subs[i]._cx-subs[j]._cx)<SNAP) subs[i]._cx=subs[j]._cx;
    }
  }
  for(let lv=1;lv<depth;lv++){y+=levelH[lv-1]+GAP_V+GAP_V;levelY.push(y);}

  let divY=null;
  if(divEnabled&&divAfterLevel<levelY.length){
    divY=levelY[divAfterLevel]+levelH[divAfterLevel]+GAP_V*0.6;
  }

  // sub merge line
  el.push(ln(CCX,CY+comp_h,CCX,SM));
  if(subs.length>1){
    const lx=Math.min(...subs.map(s=>s._cx)), rx=Math.max(...subs.map(s=>s._cx));
    el.push(ln(lx,SM,rx,SM));
  }

  const boxPos={};

  function drawTree(nodes,parentSM,level){
    const mh=levelH[level]||BH_MIN;
    const SY=levelY[level];
    nodes.forEach(node=>{
      const cx=node._cx,w=node._w;
      el.push(arr(cx,parentSM,SY));
      el.push(lbl(cx+POFF,parentSM+(SY-parentSM)/2-PH/2,node.pct));
      el.push(rect(cx-w/2,SY,w,mh,1.2));
      el.push(txt(cx-w/2,SY,w,mh,node.name.split('\\n')));
      boxPos[node.name.replace(new RegExp(String.fromCharCode(92,110),'g'),'').replace(new RegExp(String.fromCharCode(10),'g'),'')]={cx,w,y:SY,h:mh};
      if(node.children&&node.children.length){
        const ch=node.children;
        const ch_lx=Math.min(...ch.map(c=>c._cx)), ch_rx=Math.max(...ch.map(c=>c._cx));
        const ch_sm=SY+mh+GAP_V;
        el.push(ln(cx,SY+mh,cx,ch_sm));
        if(ch.length>1) el.push(ln(ch_lx,ch_sm,ch_rx,ch_sm));
        drawTree(ch,ch_sm,level+1);
      }
    });
  }
  drawTree(subs,SM,0);

  // 分界线
  if(divY!==null){
    el.push(ln(1.0,divY,SW-1.0,divY,1,'6,4'));
    el.push(lbl(1.2,divY-0.35,divLeft));
    el.push(lbl(SW-3.2,divY+0.05,divRight));
  }

  // 交叉持股
  crossLinks.forEach(xl=>{
    const nl=new RegExp(String.fromCharCode(92,110),'g'), lf=new RegExp(String.fromCharCode(10),'g');
    const fp=boxPos[xl.from.replace(nl,'').replace(lf,'')], tp=boxPos[xl.to.replace(nl,'').replace(lf,'')];
    if(fp&&tp){
      const yBot=Math.max(fp.y+fp.h, tp.y+tp.h)+0.4;
      const x1=fp.cx, x2=tp.cx;
      el.push(ln(x1,fp.y+fp.h,x1,yBot));
      el.push(ln(x1,yBot,x2,yBot));
      el.push(arr(x2,yBot,tp.y+tp.h));
      el.push(lbl((x1+x2)/2,yBot+0.05,xl.pct,'middle'));
    }
  });

  // 计算SVG尺寸
  let maxH=SM+GAP_V;
  levelY.forEach((ly,lv)=>{maxH=Math.max(maxH,ly+(levelH[lv]||BH_MIN)+0.8);});
  // Account for cross-link routes below boxes
  crossLinks.forEach(xl=>{
    const nl2=new RegExp(String.fromCharCode(92,110),'g'),lf2=new RegExp(String.fromCharCode(10),'g');
    const fp=boxPos[xl.from.replace(nl2,'').replace(lf2,'')];
    const tp=boxPos[xl.to.replace(nl2,'').replace(lf2,'')];
    if(fp&&tp) maxH=Math.max(maxH,Math.max(fp.y+fp.h,tp.y+tp.h)+1.5);
  });

  const allRX=[...subs.map(s=>s._cx+s._w/2), CCX+comp_w/2];
  const allLX=[...subs.map(s=>s._cx-s._w/2), CCX-comp_w/2];
  const contentL=Math.min(...allLX)-1.0, contentR=Math.max(...allRX)+1.0;
  const contentW=contentR-contentL;
  const containerW=wrap.clientWidth-4;
  const svgW=Math.max(contentW*sc,containerW);
  const offsetX=(svgW-contentW*sc)/2-contentL*sc;

  wrap.innerHTML=`<svg width="${svgW}" height="${p(maxH)}" xmlns="http://www.w3.org/2000/svg"><g transform="translate(${offsetX.toFixed(1)},0)">${el.join('')}</g></svg>`;
}

// ── 生成下载 ──────────────────────────────────────────────────────────────────
async function generate(){
  const btn=document.getElementById('genBtn'),st=document.getElementById('status');
  btn.disabled=true;st.textContent='正在生成...';st.className='';
  try{
    const divEnabled=document.getElementById('divEnabled').checked;
    const body={
      companyName:document.getElementById('companyName').value||'本公司',
      companyFilled:document.getElementById('companyFilled').checked,
      shareholders,
      middleTier,
      subsidiaries,
      crossLinks,
      divider:divEnabled?{
        afterLevel:parseInt(document.getElementById('divAfterLevel').value)||0,
        leftLabel:document.getElementById('divLeft').value||'境外',
        rightLabel:document.getElementById('divRight').value||'境内',
      }:null,
    };
    const resp=await fetch('/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    if(!resp.ok)throw new Error(await resp.text());
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download='股权架构图.pptx';
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
    st.textContent='文件已生成并下载！';st.className='ok';
  }catch(e){st.textContent='生成失败：'+e.message;st.className='err';}
  btn.disabled=false;
}

// 初始化
renderShareholders();renderMid();renderSubsidiaries();renderXL();
window.addEventListener('load',renderPreview);
window.addEventListener('resize',renderPreview);
</script>
</body></html>"""

class Handler(BaseHTTPRequestHandler):
    def log_message(self,fmt,*args): pass
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-Type','text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(HTML.encode('utf-8'))
    def do_POST(self):
        if self.path!='/generate':
            self.send_response(404);self.end_headers();return
        body=self.rfile.read(int(self.headers.get('Content-Length',0)))
        try:
            pptx=generate_pptx(json.loads(body.decode('utf-8')))
            self.send_response(200)
            self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition',"attachment; filename*=UTF-8''%E8%82%A1%E6%9D%83%E6%9E%B6%E6%9E%84%E5%9B%BE.pptx")
            self.send_header('Content-Length',str(len(pptx)))
            self.end_headers();self.wfile.write(pptx)
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type','text/plain; charset=utf-8')
            self.end_headers();self.wfile.write(str(e).encode('utf-8'))

if __name__=='__main__':
    import os
    port=int(os.environ.get('PORT', 5001))
    host='0.0.0.0'  # 允许外部访问
    print("股权架构图生成工具 v3 已启动")
    print("请在浏览器打开: http://localhost:{}".format(port))
    print("按 Ctrl+C 停止")
    server=HTTPServer((host, port), Handler)
    try: server.serve_forever()
    except KeyboardInterrupt: print("\n已停止")
