"""
Xelix brand normalizer (v2)
===========================
Aligns an assembled Xelix deck to the brand guidelines, robustly and without
damaging well-designed slides. Per slide it:

  1. Embeds the Barlow fonts in the deck so it renders correctly on ANY machine
     (the #1 cause of "broken" decks is the viewer not having the font, which
     makes a wider fallback overflow headers, stat numbers and cards).
  2. Sets every run in Barlow (titles in Barlow Black), preserving weight.
  3. Ensures exactly ONE Xelix wordmark bottom-right: if the slide already has a
     corner logo it is left alone; otherwise one is added, coloured for the
     background directly behind that corner (white on dark, blue on light).
  4. Repairs text contrast LOCALLY: for each text shape / table cell it looks at
     the background immediately behind it (its own fill, the panel beneath it,
     or the slide background) and only recolours text that would be unreadable
     (dark-on-dark -> white, light-on-light -> navy). Correctly-contrasted and
     mid-tone brand-coloured text is left untouched.

Backgrounds and layouts are never replaced — the deck's design is preserved.

    from normalizer import normalize_deck
    normalize_deck("in.pptx", "out.pptx",
                   logo_blue="xelix_blue.png", logo_white="xelix_white.png")
"""
import os, re, zipfile, posixpath, io
from lxml import etree

NS_P   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'

PRIMARY_BLUE = '002749'
NEAR_WHITE   = 'FFFFFF'
BODY_FONT    = 'Barlow'
HEADING_FONT = 'Barlow Black'

SLIDE_W_169 = 12192000
SLIDE_H_169 = 6858000

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, 'fonts')
# typeface -> {style: ttf filename}; embedded so the deck is self-contained
EMBED_FONTS = {
    'Barlow':       {'regular': 'Barlow-Regular.ttf', 'bold': 'Barlow-Bold.ttf'},
    'Barlow Black': {'regular': 'Barlow-Black.ttf'},
}


def _luminance(hex6):
    try:
        r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
        return (r * 299 + g * 587 + b * 114) / 1000
    except Exception:
        return 128


# ── background detection ───────────────────────────────────────────────────────
def _bg_luma_from_xml(xml):
    s = xml.decode('utf-8', 'ignore') if isinstance(xml, bytes) else xml
    m = re.search(r'<p:bg>.*?</p:bg>', s, re.DOTALL)
    if not m:
        return None
    bg = m.group(0)
    if '<a:blipFill' in bg or '<a:blip ' in bg:
        return 'img'
    cols = re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', bg)
    if cols:
        return min(_luminance(c) for c in cols)
    if re.search(r'val="(?:dk1|dk2|tx1|tx2)"', bg):
        return 20.0
    if re.search(r'val="(?:bg1|bg2|lt1|lt2)"', bg):
        return 240.0
    return None


def _resolve_rels(zin, part_name):
    d = posixpath.dirname(part_name)
    rp = posixpath.join(d, '_rels', posixpath.basename(part_name) + '.rels')
    out = {}
    if rp not in zin.namelist():
        return out
    for rel in etree.fromstring(zin.read(rp)):
        tgt, mode = rel.get('Target'), rel.get('TargetMode')
        if mode == 'External' or tgt is None:
            continue
        out[rel.get('Id')] = posixpath.normpath(posixpath.join(d, tgt))
    return out


def _image_region_luma(zin, media_part, region=None):
    """Mean luminance of a media image, optionally a sub-region given as
    (x0,y0,x1,y1) fractions."""
    if media_part not in zin.namelist():
        return None
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(zin.read(media_part))).convert('L')
        if region:
            w, h = im.size
            im = im.crop((int(region[0] * w), int(region[1] * h),
                          max(int(region[2] * w), 1), max(int(region[3] * h), 1)))
        im.thumbnail((48, 48))
        hgram = im.histogram()
        tot = sum(hgram)
        return (sum(i * c for i, c in enumerate(hgram)) / tot) if tot else None
    except Exception:
        return None


def _emu(v):
    try:
        return int(v)
    except Exception:
        return 0


def _shape_bbox(el):
    """Absolute bbox (x,y,cx,cy) from a top-level shape's xfrm, or None."""
    a = '{%s}' % NS_A
    xfrm = el.find('.//' + a + 'xfrm')
    if xfrm is None:
        return None
    off = xfrm.find(a + 'off'); ext = xfrm.find(a + 'ext')
    if off is None or ext is None:
        return None
    return (_emu(off.get('x')), _emu(off.get('y')),
            _emu(ext.get('cx')), _emu(ext.get('cy')))


def _fill_luma(el, zin=None, part=None):
    """Luminance of a shape/pic's own fill, or None if it has no solid/gradient
    fill. Pictures are sampled. blipFill on a shape -> sampled too."""
    a = '{%s}' % NS_A
    p = '{%s}' % NS_P
    spPr = el.find(p + 'spPr')
    tag = etree.QName(el).localname
    search_root = spPr if spPr is not None else el
    if search_root is not None:
        sf = search_root.find(a + 'solidFill')
        if sf is not None:
            c = sf.find(a + 'srgbClr')
            if c is not None:
                return _luminance(c.get('val'))
            sc = sf.find(a + 'schemeClr')
            if sc is not None and sc.get('val') in ('dk1', 'dk2', 'tx1', 'tx2'):
                return 20.0
            if sc is not None and sc.get('val') in ('bg1', 'bg2', 'lt1', 'lt2'):
                return 240.0
        gf = search_root.find(a + 'gradFill')
        if gf is not None:
            cols = [g.find(a + 'srgbClr').get('val')
                    for g in gf.iter(a + 'gs') if g.find(a + 'srgbClr') is not None]
            if cols:
                return min(_luminance(c) for c in cols)
    # picture or blipFill image
    if tag == 'pic' and zin is not None and part is not None:
        blip = el.find('.//' + a + 'blip')
        if blip is not None:
            rid = blip.get('{%s}embed' % NS_R)
            tgt = _resolve_rels(zin, part).get(rid)
            if tgt:
                return _image_region_luma(zin, tgt)
    return None


def _fullbleed_bg_luma(zin, part_name, xml_str, w, h, region=None):
    tree_m = re.search(r'<p:spTree>(.*)</p:spTree>', xml_str, re.DOTALL)
    if not tree_m:
        return None
    tree = tree_m.group(1)
    rels = None
    cnt = 0
    for m in re.finditer(r'<p:(sp|pic)>.*?</p:\1>', tree, re.DOTALL):
        if cnt >= 5:
            break
        cnt += 1
        blk, kind = m.group(0), m.group(1)
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', blk)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', blk)
        if not off or not ext:
            continue
        ox, oy, cx, cy = (int(off.group(1)), int(off.group(2)),
                          int(ext.group(1)), int(ext.group(2)))
        if not (abs(ox) < w * 0.04 and abs(oy) < h * 0.04
                and cx > w * 0.92 and cy > h * 0.92):
            continue
        if kind == 'pic':
            emb = re.search(r'r:embed="([^"]+)"', blk)
            if emb:
                if rels is None:
                    rels = _resolve_rels(zin, part_name)
                tgt = rels.get(emb.group(1))
                if tgt:
                    luma = _image_region_luma(zin, tgt, region)
                    return luma if luma is not None else 'img'
        else:
            sppr = re.search(r'<p:spPr>.*?</p:spPr>', blk, re.DOTALL)
            cols = re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"',
                              sppr.group(0) if sppr else '')
            if cols:
                return min(_luminance(c) for c in cols)
    return None


def _effective_bg(zin, slide_name, slide_xml, w, h, region=None):
    """(regime, confident, luma). region = bottom-right corner etc. fractions."""
    fb = _fullbleed_bg_luma(zin, slide_name, slide_xml, w, h, region)
    if fb == 'img':
        return 'dark', False, 20.0
    if isinstance(fb, (int, float)):
        return ('dark' if fb < 128 else 'light'), True, fb
    luma = _bg_luma_from_xml(slide_xml)
    if luma == 'img':
        return 'dark', False, 20.0
    if isinstance(luma, (int, float)):
        return ('dark' if luma < 128 else 'light'), True, luma
    rels = _resolve_rels(zin, slide_name)
    layout = next((t for t in rels.values() if '/slideLayouts/' in t), None)
    if layout and layout in zin.namelist():
        lxml = zin.read(layout).decode('utf-8', 'ignore')
        fb = _fullbleed_bg_luma(zin, layout, lxml, w, h, region)
        if fb == 'img':
            return 'dark', False, 20.0
        if isinstance(fb, (int, float)):
            return ('dark' if fb < 128 else 'light'), True, fb
        luma = _bg_luma_from_xml(lxml)
        if luma == 'img':
            return 'dark', False, 20.0
        if isinstance(luma, (int, float)):
            return ('dark' if luma < 128 else 'light'), True, luma
        lrels = _resolve_rels(zin, layout)
        master = next((t for t in lrels.values() if '/slideMasters/' in t), None)
        if master and master in zin.namelist():
            luma = _bg_luma_from_xml(zin.read(master))
            if isinstance(luma, (int, float)):
                return ('dark' if luma < 128 else 'light'), True, luma
    return 'light', False, 240.0   # default: assume white page


# ── fonts ──────────────────────────────────────────────────────────────────────
def _swap_fonts_in_runs(slide_xml):
    p, a = '{%s}' % NS_P, '{%s}' % NS_A
    try:
        root = etree.fromstring(slide_xml.encode('utf-8'))
    except Exception:
        return re.sub(r'(<a:(?:latin|ea|cs)\b[^>]*\btypeface=")[^"]*"',
                      r'\g<1>%s"' % BODY_FONT, slide_xml)
    for sp in root.iter(p + 'sp'):
        ph = sp.find('.//' + p + 'ph')
        is_title = ph is not None and ph.get('type') in ('title', 'ctrTitle')
        font = HEADING_FONT if is_title else BODY_FONT
        for rpr in sp.iter():
            if etree.QName(rpr).localname not in ('rPr', 'defRPr', 'endParaRPr'):
                continue
            for kind in ('latin', 'ea', 'cs'):
                el = rpr.find(a + kind)
                if el is None:
                    el = etree.SubElement(rpr, a + kind)
                    el.set('typeface', font)
                elif is_title:
                    el.set('typeface', font)       # titles always get heading font
                elif not el.get('typeface', '').startswith('Barlow'):
                    el.set('typeface', font)        # non-Barlow -> Barlow
                # else: already a Barlow variant (ExtraBold, SemiBold, etc.) — leave
    for rpr in root.iter():
        if etree.QName(rpr).localname not in ('rPr', 'defRPr', 'endParaRPr'):
            continue
        for kind in ('latin', 'ea', 'cs'):
            el = rpr.find(a + kind)
            if el is not None and not el.get('typeface', '').startswith('Barlow'):
                el.set('typeface', BODY_FONT)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                          standalone=True).decode('utf-8')


def _patch_theme_fonts(theme_xml):
    s = theme_xml.decode('utf-8', 'ignore')
    s = re.sub(r'<a:majorFont>.*?</a:majorFont>',
               lambda m: re.sub(r'(<a:latin\b[^>]*\btypeface=")[^"]*"',
                                r'\g<1>%s"' % HEADING_FONT, m.group(0)),
               s, flags=re.DOTALL)
    s = re.sub(r'<a:minorFont>.*?</a:minorFont>',
               lambda m: re.sub(r'(<a:latin\b[^>]*\btypeface=")[^"]*"',
                                r'\g<1>%s"' % BODY_FONT, m.group(0)),
               s, flags=re.DOTALL)
    return s.encode('utf-8')


# ── local contrast repair ──────────────────────────────────────────────────────
def _run_color_luma(rpr, a):
    """Return ('explicit', luma) / ('scheme-dark', None) / ('scheme-light', None)
    / ('none', None) for a run-property element."""
    if rpr is None:
        return ('none', None)
    sf = rpr.find(a + 'solidFill')
    if sf is None:
        return ('none', None)
    c = sf.find(a + 'srgbClr')
    if c is not None:
        return ('explicit', _luminance(c.get('val')))
    sc = sf.find(a + 'schemeClr')
    if sc is not None:
        v = sc.get('val')
        if v in ('dk1', 'dk2', 'tx1', 'tx2'):
            return ('scheme-dark', None)
        if v in ('bg1', 'bg2', 'lt1', 'lt2'):
            return ('scheme-light', None)
    return ('other', None)


def _set_run_color(rpr, a, hexval=None, gradient=False):
    for sf in rpr.findall(a + 'solidFill'):
        rpr.remove(sf)
    for gf in rpr.findall(a + 'gradFill'):
        rpr.remove(gf)
    if gradient:
        fill = etree.Element(a + 'gradFill')
        gsLst = etree.SubElement(fill, a + 'gsLst')
        gs0 = etree.SubElement(gsLst, a + 'gs'); gs0.set('pos', '0')
        etree.SubElement(gs0, a + 'srgbClr').set('val', 'EB3FC7')
        gs1 = etree.SubElement(gsLst, a + 'gs'); gs1.set('pos', '100000')
        etree.SubElement(gs1, a + 'srgbClr').set('val', 'E450FB')
        lin = etree.SubElement(fill, a + 'lin')
        lin.set('ang', '5400000'); lin.set('scaled', '1')
    else:
        fill = etree.Element(a + 'solidFill')
        etree.SubElement(fill, a + 'srgbClr').set('val', hexval)
    ln = rpr.find(a + 'ln')
    if ln is not None:
        ln.addnext(fill)
    else:
        rpr.insert(0, fill)


def _is_stat_run(r, a):
    """True if this run is a hero stat number: large-ish, short, has digits."""
    t = r.find(a + 't')
    txt = (t.text or '').strip() if t is not None else ''
    if not txt or len(txt) > 15:
        return False
    rpr = r.find(a + 'rPr')
    sz = int(rpr.get('sz', '0')) if rpr is not None else 0
    if sz < 1800:
        return False
    return any(c.isdigit() for c in txt)


def _fix_para_runs(para, a, dark):
    """Recolour runs in a paragraph for a dark/light local background.
    On dark backgrounds, hero stat numbers get the brand pink gradient instead
    of plain white."""
    target = NEAR_WHITE if dark else PRIMARY_BLUE
    changed = False
    runs = para.findall(a + 'r')
    for r in runs:
        rpr = r.find(a + 'rPr')
        kind, lum = _run_color_luma(rpr, a)
        need = False
        if dark:
            # unreadable if text is dark / dark-scheme / has no explicit colour
            if kind == 'explicit' and lum is not None and lum < 120:
                need = True
            elif kind in ('scheme-dark', 'none'):
                need = True
        else:
            if kind == 'explicit' and lum is not None and lum > 165:
                need = True
            elif kind == 'scheme-light':
                need = True
        if need:
            if rpr is None:
                rpr = etree.Element(a + 'rPr')
                r.insert(0, rpr)
            if dark and _is_stat_run(r, a):
                _set_run_color(rpr, a, gradient=True)
            else:
                _set_run_color(rpr, a, hexval=target)
            changed = True
    return changed


def _fix_contrast_local(slide_xml, zin, slide_name, w, h):
    a, p = '{%s}' % NS_A, '{%s}' % NS_P
    try:
        root = etree.fromstring(slide_xml.encode('utf-8'))
    except Exception:
        return slide_xml
    spTree = root.find('.//' + p + 'spTree')
    if spTree is None:
        return slide_xml

    slide_regime, _, slide_lum = _effective_bg(zin, slide_name, slide_xml, w, h)

    # Collect top-level filled panels (bbox + luma + element ref), in z-order.
    panels = []
    for el in list(spTree):
        tag = etree.QName(el).localname
        if tag not in ('sp', 'pic'):
            continue
        bbox = _shape_bbox(el)
        lum = _fill_luma(el, zin, slide_name)
        if bbox and lum is not None:
            panels.append((bbox, lum, el))

    changed = False

    # text shapes (including inside groups — root.iter catches nested sp)
    for sp in root.iter(p + 'sp'):
        txbody = sp.find(p + 'txBody')
        if txbody is None:
            continue
        if not any((t.text or '').strip() for t in sp.iter(a + 't')):
            continue
        bbox = _shape_bbox(sp)
        # for shapes inside groups, use the group's absolute bbox for panel
        # detection (the shape's own bbox is in group-local coordinates)
        detect_bbox = bbox
        parent = sp.getparent()
        if parent is not None and etree.QName(parent).localname == 'grpSp':
            grp_bbox = _shape_bbox(parent)
            if grp_bbox is not None:
                detect_bbox = grp_bbox
        own = _fill_luma(sp)
        # local bg luminance behind this text shape
        if own is not None:
            lum = own
        else:
            lum = slide_lum
            if detect_bbox is not None:
                cx, cy = detect_bbox[0] + detect_bbox[2] / 2, detect_bbox[1] + detect_bbox[3] / 2
                for pb, pl, pel in reversed(panels):
                    if pel is sp:    # identity check — different shapes with the
                        continue    # same bbox (e.g. text stacked on a coloured
                                    # tile) must still see each other
                    # require centre to be inside the panel with a margin,
                    # but cap the margin for large panels (a full-height panel's
                    # 10% margin would exclude stat numbers near the top)
                    mx = min(pb[2] * 0.10, w * 0.03)
                    my = min(pb[3] * 0.10, h * 0.03)
                    if pb[0] + mx <= cx <= pb[0] + pb[2] - mx and \
                       pb[1] + my <= cy <= pb[1] + pb[3] - my:
                        lum = pl
                        break
        dark = lum < 128
        for para in txbody.findall(a + 'p'):
            if _fix_para_runs(para, a, dark):
                changed = True

    # table cells (use each cell's own fill, else the table/slide bg)
    for tbl in root.iter(a + 'tbl'):
        # table position for slide-bg fallback
        for tr in tbl.findall(a + 'tr'):
            for tc in tr.findall(a + 'tc'):
                tcpr = tc.find(a + 'tcPr')
                lum = slide_lum
                if tcpr is not None:
                    sf = tcpr.find(a + 'solidFill')
                    if sf is not None:
                        c = sf.find(a + 'srgbClr')
                        if c is not None:
                            lum = _luminance(c.get('val'))
                        else:
                            sc = sf.find(a + 'schemeClr')
                            if sc is not None and sc.get('val') in ('dk1', 'dk2', 'tx1', 'tx2'):
                                lum = 20.0
                            elif sc is not None and sc.get('val') in ('bg1', 'bg2', 'lt1', 'lt2'):
                                lum = 240.0
                dark = lum < 128
                txbody = tc.find(a + 'txBody')
                if txbody is None:
                    continue
                for para in txbody.findall(a + 'p'):
                    if _fix_para_runs(para, a, dark):
                        changed = True

    if not changed:
        return slide_xml
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                          standalone=True).decode('utf-8')


# ── logo ────────────────────────────────────────────────────────────────────────
LOGO_RID = 'rIdXelixLogo'
LOGO_AR = 2.80
LOGO_W = 1188720
LOGO_H = int(LOGO_W / LOGO_AR)
LOGO_MARGIN = 274320


def _has_corner_logo(zin, slide_name, slide_xml, w, h):
    """True if the slide (or its layout/master) already has a small picture in
    the bottom-right corner — i.e. an existing wordmark/badge."""
    def corner_pic(xml):
        for m in re.finditer(r'<p:pic>.*?</p:pic>', xml, re.DOTALL):
            blk = m.group(0)
            off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', blk)
            ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', blk)
            if not off or not ext:
                continue
            x, y, cx, cy = (int(off.group(1)), int(off.group(2)),
                            int(ext.group(1)), int(ext.group(2)))
            ccx, ccy = (x + cx / 2) / w, (y + cy / 2) / h
            if ccx > 0.80 and ccy > 0.82 and cx < w * 0.28:
                return True
        return False
    if corner_pic(slide_xml):
        return True
    rels = _resolve_rels(zin, slide_name)
    layout = next((t for t in rels.values() if '/slideLayouts/' in t), None)
    if layout and layout in zin.namelist():
        if corner_pic(zin.read(layout).decode('utf-8', 'ignore')):
            return True
    return False


def _logo_pic_xml(slide_w, slide_h):
    x = slide_w - LOGO_W - LOGO_MARGIN
    y = slide_h - LOGO_H - LOGO_MARGIN
    return (
        '<p:pic><p:nvPicPr>'
        '<p:cNvPr id="9001" name="Xelix logo"/>'
        '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>'
        '</p:nvPicPr><p:blipFill>'
        f'<a:blip r:embed="{LOGO_RID}"/><a:stretch><a:fillRect/></a:stretch>'
        '</p:blipFill><p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{LOGO_W}" cy="{LOGO_H}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</p:spPr></p:pic>'
    )


def _strip_our_logo(slide_xml):
    return re.sub(r'<p:pic>(?:(?!</p:pic>).)*?Xelix logo.*?</p:pic>', '',
                  slide_xml, flags=re.DOTALL)


def _add_logo(slide_xml, slide_w, slide_h):
    slide_xml = _strip_our_logo(slide_xml)
    return slide_xml.replace('</p:spTree>', _logo_pic_xml(slide_w, slide_h) + '</p:spTree>', 1)


def _slide_rels_with_logo(rels_bytes, media_target):
    root = (etree.fromstring(rels_bytes) if rels_bytes
            else etree.fromstring(f'<Relationships xmlns="{NS_REL}"/>'.encode()))
    for rel in list(root):
        if rel.get('Id') == LOGO_RID:
            root.remove(rel)
    rel = etree.SubElement(root, '{%s}Relationship' % NS_REL)
    rel.set('Id', LOGO_RID)
    rel.set('Type', NS_R + '/image')
    rel.set('Target', media_target)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ── structural deck fixes ──────────────────────────────────────────────────────
def _solid_fill_hex(sp, a):
    spPr = sp.find('{%s}spPr' % NS_P)
    if spPr is None:
        return None
    sf = spPr.find(a + 'solidFill')
    if sf is None:
        return None
    c = sf.find(a + 'srgbClr')
    return c.get('val').upper() if c is not None else None


def _has_visible_fill(sp, a):
    spPr = sp.find('{%s}spPr' % NS_P)
    if spPr is None:
        return False
    if spPr.find(a + 'noFill') is not None:
        return False
    return (spPr.find(a + 'solidFill') is not None or
            spPr.find(a + 'gradFill') is not None or
            spPr.find(a + 'blipFill') is not None)


def _is_pink(hexval):
    if not hexval or len(hexval) != 6:
        return False
    try:
        r, g, b = int(hexval[0:2], 16), int(hexval[2:4], 16), int(hexval[4:6], 16)
    except Exception:
        return False
    return r > 180 and g < 100 and b > 80 and r > g + 80


def _set_shape_hero_gradient(sp, a):
    """Replace a shape's fill with the Xelix hero gradient (6115A6 → 030312)."""
    spPr = sp.find('{%s}spPr' % NS_P)
    if spPr is None:
        return
    for tag in ('solidFill', 'gradFill', 'noFill', 'blipFill', 'pattFill'):
        for el in spPr.findall(a + tag):
            spPr.remove(el)
    gf = etree.Element(a + 'gradFill')
    gf.set('flip', 'none'); gf.set('rotWithShape', '1')
    gsLst = etree.SubElement(gf, a + 'gsLst')
    gs0 = etree.SubElement(gsLst, a + 'gs'); gs0.set('pos', '0')
    etree.SubElement(gs0, a + 'srgbClr').set('val', '6115A6')
    gs1 = etree.SubElement(gsLst, a + 'gs'); gs1.set('pos', '100000')
    etree.SubElement(gs1, a + 'srgbClr').set('val', '030312')
    lin = etree.SubElement(gf, a + 'lin')
    lin.set('ang', '13500000')   # -45° (135° clockwise)
    lin.set('scaled', '0')
    # spPr child order: xfrm, prstGeom/custGeom, fill, ... — insert after geom
    geom = spPr.find(a + 'prstGeom')
    if geom is not None:
        geom.addnext(gf)
    else:
        spPr.insert(0, gf)


def _round_corners(sp, a, val):
    """Change prstGeom rect → roundRect with avLst adj=val."""
    prst = sp.find('.//' + a + 'prstGeom')
    if prst is None or prst.get('prst') != 'rect':
        return False
    prst.set('prst', 'roundRect')
    for av in prst.findall(a + 'avLst'):
        prst.remove(av)
    av = etree.SubElement(prst, a + 'avLst')
    gd = etree.SubElement(av, a + 'gd')
    gd.set('name', 'adj'); gd.set('fmla', 'val %d' % val)
    return True


def _set_shape_fill_solid(sp, a, hexval):
    """Replace a shape's fill with a solid colour."""
    spPr = sp.find('{%s}spPr' % NS_P)
    if spPr is None:
        return
    for tag in ('solidFill', 'gradFill', 'noFill', 'blipFill', 'pattFill'):
        for el in spPr.findall(a + tag):
            spPr.remove(el)
    sf = etree.Element(a + 'solidFill')
    etree.SubElement(sf, a + 'srgbClr').set('val', hexval)
    geom = spPr.find(a + 'prstGeom')
    if geom is not None:
        geom.addnext(sf)
    else:
        spPr.insert(0, sf)


def _apply_structural_fixes(slide_xml, w, h):
    """Run the deck-cleanup pass: remove off-brand decorations, replace navy
    panels with the hero gradient, round harsh-edged cards and eyebrows.
    Returns (new_xml, info) where info['footer_removed'] indicates if a footer
    was stripped (so caller can add a wordmark in its place)."""
    a = '{%s}' % NS_A; p = '{%s}' % NS_P
    info = {'footer_removed': False, 'pink_bar_removed': False,
            'eyebrows_rounded': 0, 'cards_rounded': 0,
            'navy_panels_replaced': 0, 'corner_decorations_removed': 0,
            'card_strips_removed': 0, 'navy_blocks_recoloured': 0}
    try:
        root = etree.fromstring(slide_xml.encode('utf-8'))
    except Exception:
        return slide_xml, info
    spTree = root.find('.//' + p + 'spTree')
    if spTree is None:
        return slide_xml, info

    # Pre-scan: does this slide have a navy footer bar? If so we'll sweep
    # the whole bottom band (footer text, redundant wordmark text) when
    # we remove the bar.
    has_footer = False
    for sp in spTree.findall(p + 'sp'):
        bbox = _shape_bbox(sp)
        if bbox is None:
            continue
        x, y, cx, cy = bbox
        if y / h > 0.92 and cx / w > 0.85 and cy / h < 0.10:
            fill = _solid_fill_hex(sp, a)
            if fill and _luminance(fill) < 50:
                has_footer = True
                break

    for sp in list(spTree.findall(p + 'sp')):
        bbox = _shape_bbox(sp)
        if bbox is None:
            continue
        x, y, cx, cy = bbox
        xr, yr, wr, hr = x / w, y / h, cx / w, cy / h
        cxr, cyr = xr + wr / 2, yr + hr / 2
        fill_hex = _solid_fill_hex(sp, a)
        prst = sp.find('.//' + a + 'prstGeom')
        prst_kind = prst.get('prst') if prst is not None else None
        has_text = any((t.text or '').strip() for t in sp.iter(a + 't'))

        # 1) sweep bottom band when footer is present — bar + footer text +
        #    redundant text-logo are all replaced by a proper wordmark
        if has_footer and cyr > 0.93:
            spTree.remove(sp)
            info['footer_removed'] = True
            continue

        # 2) thin pink vertical bar on left edge
        if xr < 0.05 and wr < 0.06 and hr > 0.40 and _is_pink(fill_hex):
            spTree.remove(sp); info['pink_bar_removed'] = True; continue

        # 2b) thin pink horizontal strip on a card (above/below a card body) —
        #     very wide vs tall, height tiny — pure decoration to remove
        if _is_pink(fill_hex) and hr < 0.02 and 0.08 < wr < 0.35 and yr > 0.10:
            spTree.remove(sp); info['card_strips_removed'] += 1; continue

        # 3) decorative corner squares (off-brand): small solid rect in a corner
        #    with no text content (purely visual blocks)
        in_top_right = cxr > 0.80 and cyr < 0.25
        in_top_left  = cxr < 0.20 and cyr < 0.10 and yr < 0.03
        if (in_top_right or in_top_left) and prst_kind == 'rect' \
                and fill_hex and not has_text \
                and wr < 0.20 and hr < 0.30:
            spTree.remove(sp); info['corner_decorations_removed'] += 1; continue

        # 4a) large dark navy panel on left → hero gradient
        if xr < 0.05 and wr > 0.20 and hr > 0.50 and fill_hex \
                and _luminance(fill_hex) < 60:
            _set_shape_hero_gradient(sp, a)
            info['navy_panels_replaced'] += 1
            continue

        # 4b) prominent solid-navy blocks (panels, headers, callout cards) → brand
        #     purple. The master deck uses navy mostly inside tables (graphicFrame),
        #     so standalone navy <p:sp> rectangles tend to be the heavy off-brand
        #     blocks. Skip tiny accent dots and full-slide backgrounds.
        if fill_hex and _luminance(fill_hex) < 50 and prst_kind in ('rect', 'roundRect') \
                and wr >= 0.04 and hr >= 0.03 \
                and not (wr > 0.95 and hr > 0.95):
            try:
                r, g, b = (int(fill_hex[0:2], 16), int(fill_hex[2:4], 16),
                           int(fill_hex[4:6], 16))
                # blue-leaning (not pure black, not red-leaning)
                is_navy = b > r and b > g and b > 30
            except Exception:
                is_navy = False
            if is_navy:
                _set_shape_fill_solid(sp, a, '6115A6')
                info['navy_blocks_recoloured'] += 1
                # fall through to corner-rounding so the new purple block also
                # gets rounded for consistency

        # 5) pink eyebrow banner near top → fully rounded pill (matches the
        #    master deck's button style for product labels)
        is_eyebrow = (yr < 0.10 and 0.05 < wr < 0.30 and hr < 0.06
                      and _is_pink(fill_hex))
        if is_eyebrow and prst is not None and prst_kind in ('rect', 'roundRect'):
            prst.set('prst', 'roundRect')
            for av in prst.findall(a + 'avLst'):
                prst.remove(av)
            av = etree.SubElement(prst, a + 'avLst')
            gd = etree.SubElement(av, a + 'gd')
            gd.set('name', 'adj'); gd.set('fmla', 'val 50000')
            info['eyebrows_rounded'] += 1
            continue

        # 6) other visible filled rectangles → subtle rounding (cards/panels)
        if prst_kind == 'rect' and _has_visible_fill(sp, a) \
                and wr > 0.05 and hr > 0.03 and not (wr > 0.95 and hr > 0.95):
            if _round_corners(sp, a, 10000):
                info['cards_rounded'] += 1

    return (etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                           standalone=True).decode('utf-8'), info)


# ── presentation size / content-types ──────────────────────────────────────────
def _normalize_size(prs_xml):
    s = prs_xml.decode('utf-8', 'ignore')
    m = re.search(r'<p:sldSz\b[^>]*/>', s)
    orig = None
    if m:
        cx = re.search(r'cx="(\d+)"', m.group(0)); cy = re.search(r'cy="(\d+)"', m.group(0))
        if cx and cy:
            orig = (int(cx.group(1)), int(cy.group(1)))
        s = s.replace(m.group(0),
                      f'<p:sldSz cx="{SLIDE_W_169}" cy="{SLIDE_H_169}" type="screen16x9"/>')
    return s.encode('utf-8'), orig


def _ensure_ct_default(ct_xml, ext, ctype):
    s = ct_xml.decode('utf-8', 'ignore')
    if re.search(r'<Default\s+Extension="%s"' % ext, s, re.IGNORECASE):
        return ct_xml.encode('utf-8') if isinstance(ct_xml, str) else ct_xml
    s = s.replace('<Types xmlns="%s">' % NS_CT,
                  '<Types xmlns="%s"><Default Extension="%s" ContentType="%s"/>'
                  % (NS_CT, ext, ctype), 1)
    return s.encode('utf-8')


# ── font embedding ──────────────────────────────────────────────────────────────
def _embed_fonts(merged, names):
    """Add ppt/fonts/*.fntdata, wire embeddedFontLst + rels + content-types +
    presentation flags so the deck renders with Barlow everywhere."""
    if not os.path.isdir(FONT_DIR):
        return   # fonts not bundled — skip silently
    prs_path, rels_path, ct_path = ('ppt/presentation.xml',
                                    'ppt/_rels/presentation.xml.rels',
                                    '[Content_Types].xml')
    prs = merged[prs_path].decode('utf-8', 'ignore')
    rels = etree.fromstring(merged[rels_path])
    rid_n = max([int(re.search(r'\d+', r.get('Id')).group())
                 for r in rels if re.search(r'\d+', r.get('Id'))] + [0])
    font_idx = 0
    embedded = []   # (typeface, {style: rId})
    for typeface, styles in EMBED_FONTS.items():
        rid_map = {}
        for style, fname in styles.items():
            fp = os.path.join(FONT_DIR, fname)
            if not os.path.exists(fp):
                continue
            font_idx += 1
            part = 'ppt/fonts/font%d.fntdata' % font_idx
            with open(fp, 'rb') as f:
                merged[part] = f.read()
            rid_n += 1
            rid = 'rId%d' % rid_n
            rel = etree.SubElement(rels, '{%s}Relationship' % NS_REL)
            rel.set('Id', rid)
            rel.set('Type', NS_R + '/font')
            rel.set('Target', 'fonts/font%d.fntdata' % font_idx)
            rid_map[style] = rid
        if rid_map:
            embedded.append((typeface, rid_map))
    if not embedded:
        return

    # build <p:embeddedFontLst>
    lst = '<p:embeddedFontLst>'
    for typeface, rid_map in embedded:
        lst += '<p:embeddedFont><p:font typeface="%s"/>' % typeface
        if 'regular' in rid_map:
            lst += '<p:regular r:id="%s"/>' % rid_map['regular']
        if 'bold' in rid_map:
            lst += '<p:bold r:id="%s"/>' % rid_map['bold']
        lst += '</p:embeddedFont>'
    lst += '</p:embeddedFontLst>'

    # insert after <p:notesSz .../> (correct schema position)
    m = re.search(r'<p:notesSz\b[^>]*/>', prs)
    if m:
        prs = prs[:m.end()] + lst + prs[m.end():]
    else:
        prs = prs.replace('</p:presentation>', lst + '</p:presentation>', 1)
    # set embed flags on <p:presentation ...>
    def _flags(mm):
        tag = mm.group(0)
        if 'embedTrueTypeFonts' not in tag:
            tag = tag[:-1] + ' embedTrueTypeFonts="1">'
        if 'saveSubsetFonts' not in tag:
            tag = tag[:-1] + ' saveSubsetFonts="0">'
        return tag
    prs = re.sub(r'<p:presentation\b[^>]*>', _flags, prs, count=1)

    merged[prs_path] = prs.encode('utf-8')
    merged[rels_path] = etree.tostring(rels, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)
    merged[ct_path] = _ensure_ct_default(merged[ct_path], 'fntdata',
                                         'application/x-fontdata')


# ── orchestrator ────────────────────────────────────────────────────────────────
def normalize_deck(in_path, out_path, *, logo_blue, logo_white,
                   swap_fonts=True, embed_fonts=True, add_logo=False,
                   fix_contrast=True, force_169=True, **_ignored):
    report = {'slides': [], 'warnings': [], 'size': None, 'logos_added': 0,
              'logos_kept': 0}
    zin = zipfile.ZipFile(in_path, 'r')
    names = zin.namelist()
    slide_names = sorted(
        [n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
        key=lambda n: int(re.search(r'(\d+)', n).group(1)))

    used_blue = used_white = False
    merged = {n: zin.read(n) for n in names}

    for sn in slide_names:
        sxml = merged[sn].decode('utf-8', 'ignore')

        # structural deck fixes (remove off-brand decorations, replace navy
        # panels with hero gradient, round harsh-edged shapes)
        sxml, struct_info = _apply_structural_fixes(sxml, SLIDE_W_169, SLIDE_H_169)
        for k, v in struct_info.items():
            if isinstance(v, bool) and v:
                report.setdefault(k + '_count', 0)
                report[k + '_count'] += 1
            elif isinstance(v, int) and v:
                report.setdefault(k, 0)
                report[k] += v

        if swap_fonts:
            sxml = _swap_fonts_in_runs(sxml)
        if fix_contrast:
            sxml = _fix_contrast_local(sxml, zin, sn, SLIDE_W_169, SLIDE_H_169)

        logo_choice = None
        # add a wordmark bottom-right on slides where the footer was removed
        # (the footer was acting as a presence marker — replace with the logo)
        if struct_info['footer_removed'] and not _has_corner_logo(
                zin, sn, merged[sn].decode('utf-8', 'ignore'),
                SLIDE_W_169, SLIDE_H_169):
            region = (0.74, 0.80, 1.0, 1.0)
            regime, _, _ = _effective_bg(zin, sn,
                                         merged[sn].decode('utf-8', 'ignore'),
                                         SLIDE_W_169, SLIDE_H_169, region)
            dark = (regime == 'dark')
            logo_choice = 'white' if dark else 'blue'
            sxml = _add_logo(sxml, SLIDE_W_169, SLIDE_H_169)
            tgt = ('../media/xelix_logo_white.png' if dark
                   else '../media/xelix_logo_blue.png')
            rp = 'ppt/slides/_rels/%s.rels' % posixpath.basename(sn)
            merged[rp] = _slide_rels_with_logo(merged.get(rp), tgt)
            used_white = used_white or dark
            used_blue = used_blue or (not dark)
            report['logos_added'] += 1

        merged[sn] = sxml.encode('utf-8')
        report['slides'].append({'slide': posixpath.basename(sn),
                                  'logo': logo_choice or 'existing'})

    # theme fonts
    if swap_fonts:
        for n in names:
            if re.match(r'ppt/theme/theme\d+\.xml$', n):
                merged[n] = _patch_theme_fonts(merged[n])

    # size
    if force_169 and 'ppt/presentation.xml' in merged:
        merged['ppt/presentation.xml'], orig = _normalize_size(merged['ppt/presentation.xml'])
        report['size'] = orig
        if orig and abs(orig[0] / orig[1] - SLIDE_W_169 / SLIDE_H_169) > 0.02:
            report['warnings'].append(
                'Source aspect ratio %.3f differs from 16:9 — only the canvas was '
                'resized, shapes were not rescaled.' % (orig[0] / orig[1]))

    # logo media + content types
    if used_blue or used_white:
        if '[Content_Types].xml' in merged:
            merged['[Content_Types].xml'] = _ensure_ct_default(
                merged['[Content_Types].xml'], 'png', 'image/png')
        if used_blue:
            with open(logo_blue, 'rb') as f:
                merged['ppt/media/xelix_logo_blue.png'] = f.read()
        if used_white:
            with open(logo_white, 'rb') as f:
                merged['ppt/media/xelix_logo_white.png'] = f.read()

    # embed fonts (do last; touches presentation.xml + rels + content-types)
    if embed_fonts:
        try:
            _embed_fonts(merged, names)
            report['fonts_embedded'] = True
        except Exception as e:
            report['warnings'].append('Font embedding skipped: %s' % e)
            report['fonts_embedded'] = False

    zin.close()
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in merged.items():
            z.writestr(n, data)
    return report


if __name__ == '__main__':
    import sys, json
    a = sys.argv
    print(json.dumps(normalize_deck(a[1], a[2], logo_blue=a[3], logo_white=a[4]),
                     indent=2))
