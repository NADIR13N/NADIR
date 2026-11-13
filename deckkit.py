# -*- coding: utf-8 -*-
"""Helpers: design system + native PowerPoint animation/transition XML injection."""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

In = Inches

# ---------------------------------------------------------------- palette
GREEN_D = "174C3C"
GREEN_E = "2D8061"
SAGE    = "A9BDA5"
IVORY   = "F6F5EE"
WHITE   = "FFFFFF"
TEAL    = "267C83"
GOLD    = "C5A45A"
CHARCOAL= "303B38"
TERRA   = "A65A3A"
BROWN   = "7A4A21"
GREY_TX = "6E7B74"

F_TITLE = "Aptos Display"
F_BODY  = "Aptos"


def C(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------- shapes
def rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75, round_=None, alpha=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if round_ is not None else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, In(x), In(y), In(w), In(h))
    if round_ is not None:
        try:
            shp.adjustments[0] = round_
        except Exception:
            pass
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
        if alpha is not None:
            _fill_alpha(shp, alpha)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(line_w)
    return shp


def _fill_alpha(shp, pct):
    """pct = opacity percent (0-100)."""
    sp = shp.fill._xPr.find(qn('a:solidFill'))
    srgb = sp.find(qn('a:srgbClr'))
    a = etree.SubElement(srgb, qn('a:alpha'))
    a.set('val', str(int(pct * 1000)))


def oval(slide, x, y, w, h, fill=None, line=None, line_w=1.0, alpha=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, In(x), In(y), In(w), In(h))
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
        if alpha is not None:
            _fill_alpha(shp, alpha)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(line_w)
    return shp


def shape(slide, mso, x, y, w, h, fill=None, line=None, line_w=1.0, rot=None):
    shp = slide.shapes.add_shape(mso, In(x), In(y), In(w), In(h))
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = C(fill)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line)
        shp.line.width = Pt(line_w)
    if rot is not None:
        shp.rotation = rot
    return shp


def shape_text(shp, text, size, color, bold=True, font=F_BODY, spc=None, italic=False):
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = In(0.05); tf.margin_right = In(0.05)
    tf.margin_top = In(0.02); tf.margin_bottom = In(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.name = font; r.font.color.rgb = C(color)
    if spc:
        r.font._rPr.set('spc', str(spc))
    return shp


# ---------------------------------------------------------------- text
def tb(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: list of dicts:
       {'runs':[{'t','sz','c','b','i','f','spc'}...], 'align', 'space_after', 'space_before', 'line'}"""
    box = slide.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = para.get('align', PP_ALIGN.LEFT)
        if 'space_after' in para:
            p.space_after = Pt(para['space_after'])
        if 'space_before' in para:
            p.space_before = Pt(para['space_before'])
        if 'line' in para:
            p.line_spacing = para['line']
        for rs in para['runs']:
            run = p.add_run()
            run.text = rs['t']
            f = run.font
            f.size = Pt(rs.get('sz', 12))
            f.name = rs.get('f', F_BODY)
            f.bold = rs.get('b', False)
            f.italic = rs.get('i', False)
            f.color.rgb = C(rs.get('c', CHARCOAL))
            if rs.get('spc'):
                f._rPr.set('spc', str(rs['spc']))
    return box


def bullets(slide, x, y, w, h, items, sz=11.5, color=CHARCOAL, dot=GOLD, gap=5, line=1.08, dot_sz=None):
    paras = []
    for it in items:
        runs = [{'t': '\u25AA  ', 'sz': dot_sz or sz, 'c': dot, 'b': True}]
        if isinstance(it, list):
            for seg in it:
                seg = dict(seg)
                seg.setdefault('sz', sz); seg.setdefault('c', color)
                runs.append(seg)
        else:
            runs.append({'t': it, 'sz': sz, 'c': color})
        paras.append({'runs': runs, 'space_after': gap, 'line': line})
    return tb(slide, x, y, w, h, paras)


# ---------------------------------------------------------------- pictures
def pic(slide, path, x, y, w, h, line=None, line_w=1.5):
    im = Image.open(path)
    iw, ih = im.size
    box_ar = w / h
    im_ar = iw / ih
    p = slide.shapes.add_picture(path, In(x), In(y), In(w), In(h))
    if im_ar > box_ar:
        frac = (1 - box_ar / im_ar) / 2
        p.crop_left = frac; p.crop_right = frac
    else:
        frac = (1 - im_ar / box_ar) / 2
        p.crop_top = frac; p.crop_bottom = frac
    p.shadow.inherit = False
    if line is not None:
        p.line.color.rgb = C(line)
        p.line.width = Pt(line_w)
    return p


# ---------------------------------------------------------------- page furniture
def header(slide, title, kicker=None):
    rect(slide, 0.55, 0.52, 0.085, 0.60, fill=GOLD)
    paras = [{'runs': [{'t': title, 'sz': 26, 'c': GREEN_D, 'b': True, 'f': F_TITLE}]}]
    if kicker:
        paras.append({'runs': [{'t': kicker, 'sz': 11.5, 'c': TEAL, 'b': False}], 'space_before': 2})
    return tb(slide, 0.80, 0.44, 10.5, 1.0, paras)


def footer(slide, idx):
    tb(slide, 0.55, 7.08, 7.0, 0.3,
       [{'runs': [{'t': "Embryogenèse somatique chez les plantes — Biotechnologie Appliquée",
                   'sz': 9, 'c': GREY_TX}]}])
    tb(slide, 11.6, 7.08, 1.2, 0.3,
       [{'runs': [{'t': f"{idx:02d} / 07", 'sz': 9, 'c': GREY_TX, 'b': True}], 'align': PP_ALIGN.RIGHT}])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ---------------------------------------------------------------- animations (native PPTX)
P  = "http://schemas.openxmlformats.org/presentationml/2006/main"
A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"

# effect -> (presetID, presetSubtype, animEffect filter, needs scale)
EFFECTS = {
    'fade':   (10, 0, 'fade', False),
    'appear': (1, 0, None, False),
    'wipe_l': (22, 8, 'wipe(right)', False),
    'wipe_b': (22, 4, 'wipe(up)', False),
    'wipe_t': (22, 1, 'wipe(down)', False),
    'zoom':   (23, 16, 'fade', True),
}


def add_animations(slide, specs):
    """specs: list of dicts {spid, effect, trigger: click|after|with|auto, delay: ms, dur: ms}.
    First spec must be trigger 'click' or 'auto'."""
    sld = slide._element
    for t in sld.findall(qn('p:timing')):
        sld.remove(t)
    timing = etree.SubElement(sld, qn('p:timing'))
    tnLst = etree.SubElement(timing, qn('p:tnLst'))
    par0 = etree.SubElement(tnLst, qn('p:par'))
    ctn0 = etree.SubElement(par0, qn('p:cTn'))
    ctn0.set('id', '1'); ctn0.set('dur', 'indefinite')
    ctn0.set('restart', 'never'); ctn0.set('nodeType', 'tmRoot')
    child0 = etree.SubElement(ctn0, qn('p:childTnLst'))
    seq = etree.SubElement(child0, qn('p:seq'))
    seq.set('concurrent', '1'); seq.set('nextAc', 'seek')
    ctn1 = etree.SubElement(seq, qn('p:cTn'))
    ctn1.set('id', '2'); ctn1.set('dur', 'indefinite'); ctn1.set('nodeType', 'mainSeq')
    main = etree.SubElement(ctn1, qn('p:childTnLst'))
    prev = etree.SubElement(seq, qn('p:prevCondLst'))
    c = etree.SubElement(prev, qn('p:cond')); c.set('evt', 'onPrev'); c.set('delay', '0')
    tgt = etree.SubElement(c, qn('p:tgtEl')); etree.SubElement(tgt, qn('p:sldTgt'))
    nxt = etree.SubElement(seq, qn('p:nextCondLst'))
    c = etree.SubElement(nxt, qn('p:cond')); c.set('evt', 'onNext'); c.set('delay', '0')
    tgt = etree.SubElement(c, qn('p:tgtEl')); etree.SubElement(tgt, qn('p:sldTgt'))

    idc = [2]
    def nid():
        idc[0] += 1
        return str(idc[0])

    group = None          # current click-group childTnLst
    for s in specs:
        trig = s['trigger']
        if trig in ('click', 'auto') or group is None:
            gpar = etree.SubElement(main, qn('p:par'))
            gctn = etree.SubElement(gpar, qn('p:cTn'))
            gctn.set('id', nid()); gctn.set('fill', 'hold')
            gst = etree.SubElement(gctn, qn('p:stCondLst'))
            g1 = etree.SubElement(gst, qn('p:cond')); g1.set('delay', 'indefinite')
            if trig == 'auto':
                g2 = etree.SubElement(gst, qn('p:cond'))
                g2.set('evt', 'onBegin'); g2.set('delay', '0')
                tn = etree.SubElement(g2, qn('p:tn')); tn.set('val', '2')
            group = etree.SubElement(gctn, qn('p:childTnLst'))
        node_type = {'click': 'clickEffect', 'after': 'afterEffect',
                     'with': 'withEffect', 'auto': 'afterEffect'}[trig]
        spid = str(s['spid'])
        dur = int(s.get('dur', 500))
        pid, sub, filt, scale = EFFECTS[s['effect']]

        epar = etree.SubElement(group, qn('p:par'))
        ectn = etree.SubElement(epar, qn('p:cTn'))
        ectn.set('id', nid()); ectn.set('fill', 'hold')
        est = etree.SubElement(ectn, qn('p:stCondLst'))
        ec = etree.SubElement(est, qn('p:cond')); ec.set('delay', str(int(s.get('delay', 0))))
        ecchild = etree.SubElement(ectn, qn('p:childTnLst'))
        p2 = etree.SubElement(ecchild, qn('p:par'))
        c2 = etree.SubElement(p2, qn('p:cTn'))
        c2.set('id', nid())
        c2.set('presetID', str(pid)); c2.set('presetClass', 'entr')
        c2.set('presetSubtype', str(sub)); c2.set('fill', 'hold')
        c2.set('grpId', '0'); c2.set('nodeType', node_type)
        c2st = etree.SubElement(c2, qn('p:stCondLst'))
        c2c = etree.SubElement(c2st, qn('p:cond')); c2c.set('delay', '0')
        c2child = etree.SubElement(c2, qn('p:childTnLst'))

        st = etree.SubElement(c2child, qn('p:set'))
        bh = etree.SubElement(st, qn('p:cBhvr'))
        bctn = etree.SubElement(bh, qn('p:cTn'))
        bctn.set('id', nid()); bctn.set('dur', '1'); bctn.set('fill', 'hold')
        bst = etree.SubElement(bctn, qn('p:stCondLst'))
        bc = etree.SubElement(bst, qn('p:cond')); bc.set('delay', '0')
        tg = etree.SubElement(bh, qn('p:tgtEl'))
        sp = etree.SubElement(tg, qn('p:spTgt')); sp.set('spid', spid)
        an = etree.SubElement(bh, qn('p:attrNameLst'))
        anm = etree.SubElement(an, qn('p:attrName')); anm.text = 'style.visibility'
        to = etree.SubElement(st, qn('p:to'))
        sv = etree.SubElement(to, qn('p:strVal')); sv.set('val', 'visible')

        if filt:
            ae = etree.SubElement(c2child, qn('p:animEffect'))
            ae.set('transition', 'in'); ae.set('filter', filt)
            bh = etree.SubElement(ae, qn('p:cBhvr'))
            bctn = etree.SubElement(bh, qn('p:cTn')); bctn.set('id', nid()); bctn.set('dur', str(dur))
            tg = etree.SubElement(bh, qn('p:tgtEl'))
            sp = etree.SubElement(tg, qn('p:spTgt')); sp.set('spid', spid)
        if scale:
            asc = etree.SubElement(c2child, qn('p:animScale'))
            bh = etree.SubElement(asc, qn('p:cBhvr'))
            bctn = etree.SubElement(bh, qn('p:cTn'))
            bctn.set('id', nid()); bctn.set('dur', str(dur)); bctn.set('fill', 'hold')
            tg = etree.SubElement(bh, qn('p:tgtEl'))
            sp = etree.SubElement(tg, qn('p:spTgt')); sp.set('spid', spid)
            fr = etree.SubElement(asc, qn('p:from'))
            fr.set('x', '60000'); fr.set('y', '60000')
            to2 = etree.SubElement(asc, qn('p:to'))
            to2.set('x', '100000'); to2.set('y', '100000')

    bld = etree.SubElement(timing, qn('p:bldLst'))
    seen = set()
    for s in specs:
        if s['spid'] in seen:
            continue
        seen.add(s['spid'])
        bp = etree.SubElement(bld, qn('p:bldP'))
        bp.set('spid', str(s['spid'])); bp.set('grpId', '0')


def add_transition(slide, kind):
    """kind: fade | fade_black | push_up | wipe_r | zoom_in | morph"""
    sld = slide._element
    for t in sld.findall(qn('p:transition')):
        sld.remove(t)
    for t in sld.findall('{%s}AlternateContent' % MC):
        sld.remove(t)
    if kind == 'morph':
        ac = etree.SubElement(sld, '{%s}AlternateContent' % MC)
        ch = etree.SubElement(ac, '{%s}Choice' % MC)
        ch.set('Requires', 'p159')
        tr = etree.SubElement(ch, qn('p:transition'))
        tr.set('spd', 'slow')
        tr.set('{%s}dur' % P14, '1400')
        etree.SubElement(tr, '{%s}morph' % P159).set('option', 'byObject')
        fb = etree.SubElement(ac, '{%s}Fallback' % MC)
        tr2 = etree.SubElement(fb, qn('p:transition'))
        tr2.set('spd', 'slow')
        etree.SubElement(tr2, qn('p:fade'))
        return
    tr = etree.SubElement(sld, qn('p:transition'))
    tr.set('spd', 'med')
    if kind == 'fade':
        etree.SubElement(tr, qn('p:fade'))
    elif kind == 'fade_black':
        etree.SubElement(tr, qn('p:fade')).set('thruBlk', '1')
    elif kind == 'push_up':
        etree.SubElement(tr, qn('p:push')).set('dir', 'u')
    elif kind == 'wipe_r':
        etree.SubElement(tr, qn('p:wipe')).set('dir', 'r')
    elif kind == 'zoom_in':
        etree.SubElement(tr, qn('p:zoom')).set('dir', 'in')
