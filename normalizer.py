"""
Xelix brand normalizer
=======================

Deterministic, in-place brand layer for PPTX decks. Takes one department deck
and returns the SAME deck with a consistent Xelix brand skin applied to every
slide, so several departments' decks can be stitched into one combined meeting
deck that reads as a single on-brand deck.

It does NOT restructure content, reflow text, or rebuild slides. The input is
assumed to be a well-structured slide; we only unify the brand layer:

  1. Normalize slide size to 16:9 (warn if source aspect differs).
  2. Swap all fonts to Barlow (theme major/minor + explicit run overrides;
     titles get Barlow Black).
  3. Standardize the background WITHIN each slide's existing regime:
       - a slide already on a dark background -> the one canonical hero gradient
         (6115A6 -> 030312)
       - a slide on a light/no background     -> pure white
     (We do not flip dark<->light, so deliberately-designed slides survive.)
  4. Contrast-repair text colour so it stays readable on the new background
       - light slide: near-white runs -> Primary Blue 002749
       - dark slide : near-dark runs  -> near-white
     (This is repair, not a wholesale recolour: mid-tone / intentional brand
      colours are left alone.)
  5. Drop the Xelix wordmark bottom-right on every slide (blue on light,
     white on dark).

Everything is regex/byte level on the slide XML where possible (re-serializing
DrawingML with a generic XML lib reorders namespaces and triggers PowerPoint's
"repaired" dialog). lxml is used only for read-only inspection.

Usage:
    from normalizer import normalize_deck
    report = normalize_deck("in.pptx", "out.pptx",
                            logo_blue="xelix_blue.png",
                            logo_white="xelix_white.png")
"""

import os, re, zipfile, posixpath
from lxml import etree

# ── Namespaces ────────────────────────────────────────────────────────────────
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

# ── Brand spec (from Xelix Brand Guidelines, April 2023) ───────────────────────
PRIMARY_BLUE = '002749'   # default text colour, light backgrounds
NEAR_WHITE   = 'FFFFFF'   # text colour on dark backgrounds
HERO_TOP     = '6115A6'   # hero gradient start (purple)
HERO_BOTTOM  = '030312'   # hero gradient end (near-black)
BODY_FONT    = 'Barlow'
HEADING_FONT = 'Barlow Black'

# 16:9 canonical size (EMU) — 13.333in x 7.5in
SLIDE_W_169 = 12192000
SLIDE_H_169 = 6858000

# ── Background fragments ───────────────────────────────────────────────────────
WHITE_BG = (
    '<p:bg><p:bgPr>'
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    '<a:effectLst/>'
    '</p:bgPr></p:bg>'
)

# Hero gradient: purple (top-left) -> near-black (bottom-right).
# ang is in 60000ths of a degree, clockwise from 3 o'clock; 45deg points
# down-and-right so the 0% stop lands top-left. (Verified by render.)
HERO_BG = (
    '<p:bg><p:bgPr>'
    '<a:gradFill rotWithShape="1">'
    '<a:gsLst>'
    f'<a:gs pos="0"><a:srgbClr val="{HERO_TOP}"/></a:gs>'
    f'<a:gs pos="100000"><a:srgbClr val="{HERO_BOTTOM}"/></a:gs>'
    '</a:gsLst>'
    '<a:lin ang="2700000" scaled="1"/>'
    '</a:gradFill>'
    '<a:effectLst/>'
    '</p:bgPr></p:bg>'
)


def _luminance(hex6):
    try:
        r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
        return (r * 299 + g * 587 + b * 114) / 1000
    except Exception:
        return 128


# ── Background-regime detection ────────────────────────────────────────────────
def _bg_luma_from_xml(xml_bytes):
    """Return luminance of the first explicit <p:bg> *solid/gradient* fill in
    this part. Returns ('img', None) for a blipFill image background, or
    (None) if the part declares no background of its own."""
    if isinstance(xml_bytes, bytes):
        s = xml_bytes.decode('utf-8', 'ignore')
    else:
        s = xml_bytes
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
    """Return {rId: target_part_name} for a given part."""
    d = posixpath.dirname(part_name)
    rels_path = posixpath.join(d, '_rels', posixpath.basename(part_name) + '.rels')
    out = {}
    if rels_path not in zin.namelist():
        return out
    root = etree.fromstring(zin.read(rels_path))
    for rel in root:
        tgt, mode = rel.get('Target'), rel.get('TargetMode')
        if mode == 'External' or tgt is None:
            continue
        out[rel.get('Id')] = posixpath.normpath(posixpath.join(d, tgt))
    return out


def _image_mean_luma(zin, media_part):
    """Sample a media image's mean luminance with PIL. None if unreadable
    (e.g. EMF/WMF vector formats PIL can't decode)."""
    if media_part not in zin.namelist():
        return None
    try:
        from PIL import Image
        import io
        im = Image.open(io.BytesIO(zin.read(media_part))).convert('L')
        im.thumbnail((48, 48))
        h = im.histogram()
        total = sum(h)
        if not total:
            return None
        return sum(i * c for i, c in enumerate(h)) / total
    except Exception:
        return None


def _fullbleed_bg_luma(zin, part_name, xml_str, w, h):
    """Look for a backmost full-bleed picture or rectangle covering the slide
    and return its luminance (sampling the image if it's a picture)."""
    tree_m = re.search(r'<p:spTree>(.*)</p:spTree>', xml_str, re.DOTALL)
    if not tree_m:
        return None
    tree = tree_m.group(1)
    rels = None
    # Examine the first few shapes/pics (a bg element is normally backmost)
    count = 0
    for m in re.finditer(r'<p:(sp|pic)>.*?</p:\1>', tree, re.DOTALL):
        if count >= 4:
            break
        count += 1
        blk, kind = m.group(0), m.group(1)
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', blk)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', blk)
        if not off or not ext:
            continue
        ox, oy = int(off.group(1)), int(off.group(2))
        cx, cy = int(ext.group(1)), int(ext.group(2))
        covers = (abs(ox) < w * 0.04 and abs(oy) < h * 0.04
                  and cx > w * 0.92 and cy > h * 0.92)
        if not covers:
            continue
        if kind == 'pic':
            emb = re.search(r'r:embed="([^"]+)"', blk)
            if emb:
                if rels is None:
                    rels = _resolve_rels(zin, part_name)
                tgt = rels.get(emb.group(1))
                if tgt:
                    luma = _image_mean_luma(zin, tgt)
                    if luma is not None:
                        return luma
                    return 'img'   # unreadable full-bleed image — treat as image bg
        else:  # sp rectangle
            cols = re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"',
                              re.search(r'<p:spPr>.*?</p:spPr>', blk, re.DOTALL).group(0)
                              if re.search(r'<p:spPr>.*?</p:spPr>', blk, re.DOTALL) else '')
            if cols:
                return min(_luminance(c) for c in cols)
    return None


def _effective_bg(zin, slide_name, slide_xml, w, h):
    """Resolve the slide's visible background regime.
    Returns (regime, confident) where regime in {'light','dark'}.
    Walks: full-bleed shape/pic on the slide -> slide <p:bg> ->
    layout (full-bleed + bg) -> master <p:bg> -> text-colour fallback.
    Image backgrounds are sampled for luminance; if unreadable, regime is the
    best guess with confident=False."""
    # 1. Full-bleed visual on the slide itself (most common for these decks)
    fb = _fullbleed_bg_luma(zin, slide_name, slide_xml, w, h)
    if fb == 'img':
        return 'dark', False          # unknown image — guess dark, low confidence
    if isinstance(fb, (int, float)):
        return ('dark' if fb < 128 else 'light'), True

    # 2. Slide's own <p:bg>
    luma = _bg_luma_from_xml(slide_xml)
    if luma == 'img':
        return 'dark', False
    if isinstance(luma, (int, float)):
        return ('dark' if luma < 128 else 'light'), True

    # 3. Layout (full-bleed + its own bg)
    rels = _resolve_rels(zin, slide_name)
    layout = next((t for t in rels.values() if '/slideLayouts/' in t), None)
    if layout and layout in zin.namelist():
        lxml = zin.read(layout).decode('utf-8', 'ignore')
        fb = _fullbleed_bg_luma(zin, layout, lxml, w, h)
        if fb == 'img':
            return 'dark', False
        if isinstance(fb, (int, float)):
            return ('dark' if fb < 128 else 'light'), True
        luma = _bg_luma_from_xml(lxml)
        if luma == 'img':
            return 'dark', False
        if isinstance(luma, (int, float)):
            return ('dark' if luma < 128 else 'light'), True
        # 4. Master
        lrels = _resolve_rels(zin, layout)
        master = next((t for t in lrels.values() if '/slideMasters/' in t), None)
        if master and master in zin.namelist():
            luma = _bg_luma_from_xml(zin.read(master))
            if isinstance(luma, (int, float)):
                return ('dark' if luma < 128 else 'light'), True

    # 5. Fallback: dominant text colour (low confidence)
    cols = re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', slide_xml)
    if cols:
        light = sum(1 for c in cols if _luminance(c) > 160)
        return ('dark' if light > len(cols) / 2 else 'light'), False
    return 'light', False


# ── Pass: background ───────────────────────────────────────────────────────────
def _set_background(slide_xml, dark):
    frag = HERO_BG if dark else WHITE_BG
    slide_xml = re.sub(r'<p:bg>.*?</p:bg>', '', slide_xml, flags=re.DOTALL)
    return slide_xml.replace('<p:cSld>', '<p:cSld>' + frag, 1)


# ── Pass: fonts ────────────────────────────────────────────────────────────────
def _swap_fonts_in_runs(slide_xml):
    """Force every explicit run font to Barlow. Title-placeholder runs get
    Barlow Black; all other latin/ea/cs typefaces become Barlow."""
    p, a = '{%s}' % NS_P, '{%s}' % NS_A
    try:
        root = etree.fromstring(slide_xml.encode('utf-8'))
    except Exception:
        # Fall back to a blunt global swap if the slide won't parse
        slide_xml = re.sub(r'(<a:(?:latin|ea|cs)\b[^>]*\btypeface=")[^"]*"',
                           r'\g<1>%s"' % BODY_FONT, slide_xml)
        return slide_xml

    for sp in root.iter(p + 'sp'):
        ph = sp.find('.//' + p + 'ph')
        is_title = ph is not None and ph.get('type') in ('title', 'ctrTitle')
        font = HEADING_FONT if is_title else BODY_FONT
        for rpr in sp.iter():
            tag = etree.QName(rpr).localname
            if tag not in ('rPr', 'defRPr', 'endParaRPr'):
                continue
            for kind in ('latin', 'ea', 'cs'):
                el = rpr.find(a + kind)
                if el is None:
                    el = etree.SubElement(rpr, a + kind)
                el.set('typeface', font if kind == 'latin' else font)
    # Non-shape text (tables, smartart runs) -> Barlow
    for rpr in root.iter():
        if etree.QName(rpr).localname not in ('rPr', 'defRPr', 'endParaRPr'):
            continue
        for kind in ('latin', 'ea', 'cs'):
            el = rpr.find(a + kind)
            if el is not None and el.get('typeface') not in (BODY_FONT, HEADING_FONT):
                el.set('typeface', BODY_FONT)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                          standalone=True).decode('utf-8')


def _patch_theme_fonts(theme_xml):
    """Set theme major font -> Barlow Black, minor font -> Barlow, so every
    placeholder that inherits its font from the theme becomes on-brand."""
    s = theme_xml.decode('utf-8', 'ignore')

    def _fix(block, font):
        block = re.sub(r'(<a:latin\b[^>]*\btypeface=")[^"]*"',
                       r'\g<1>%s"' % font, block)
        return block

    def _major(m):  return _fix(m.group(0), HEADING_FONT)
    def _minor(m):  return _fix(m.group(0), BODY_FONT)
    s = re.sub(r'<a:majorFont>.*?</a:majorFont>', _major, s, flags=re.DOTALL)
    s = re.sub(r'<a:minorFont>.*?</a:minorFont>', _minor, s, flags=re.DOTALL)
    return s.encode('utf-8')


# ── Pass: contrast-repair text colour ──────────────────────────────────────────
def _fix_text_contrast(slide_xml, dark):
    """Repair only text that would be unreadable on the new background.
    light slide: near-white run colours -> Primary Blue
    dark slide : near-dark run colours  -> near-white
    Operates only inside run-property blocks so shape fills are untouched."""
    target_dark_text = PRIMARY_BLUE   # used on light slides
    target_light_text = NEAR_WHITE    # used on dark slides

    def _fix_block(m):
        block = m.group(0)

        def _swap(m2):
            val = m2.group(1)
            lum = _luminance(val)
            if not dark:
                # light slide: near-white (unreadable) OR near-black (off-brand
                # generic black) -> brand navy. Mid-tone brand colours untouched.
                if lum > 170 or lum < 90:
                    return f'<a:srgbClr val="{target_dark_text}"'
            else:
                if lum < 90:                    # dark text on gradient -> white
                    return f'<a:srgbClr val="{target_light_text}"'
            return m2.group(0)
        block = re.sub(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', _swap, block)

        # scheme colours that resolve wrong on the new bg
        if not dark:
            # white-ish scheme text on white -> navy
            block = re.sub(r'<a:schemeClr val="(?:bg1|bg2|lt1|lt2)"\s*/>',
                           f'<a:srgbClr val="{target_dark_text}"/>', block)
        else:
            # dark scheme text on gradient -> white
            block = re.sub(r'<a:schemeClr val="(?:tx1|tx2|dk1|dk2)"\s*/>',
                           f'<a:srgbClr val="{target_light_text}"/>', block)
        return block

    for tag in ('rPr', 'defRPr', 'endParaRPr'):
        slide_xml = re.sub(r'<a:%s\b[^>]*>.*?</a:%s>' % (tag, tag),
                           _fix_block, slide_xml, flags=re.DOTALL)
        # self-closing runs carry no colour; on dark bg they inherit dark -> fix
        if dark:
            slide_xml = re.sub(
                r'<a:%s(\b[^>]*?)/>' % tag,
                lambda m: f'<a:{tag}{m.group(1)}>'
                          f'<a:solidFill><a:srgbClr val="{target_light_text}"/></a:solidFill>'
                          f'</a:{tag}>',
                slide_xml)
    return slide_xml


# ── Pass: logo ─────────────────────────────────────────────────────────────────
LOGO_RID = 'rIdXelixLogo'
LOGO_AR = 2.80           # wordmark aspect ratio (w/h)
LOGO_W = 1188720         # ~1.3in
LOGO_H = int(LOGO_W / LOGO_AR)
LOGO_MARGIN = 274320     # ~0.3in


def _logo_pic_xml(slide_w, slide_h):
    x = slide_w - LOGO_W - LOGO_MARGIN
    y = slide_h - LOGO_H - LOGO_MARGIN
    return (
        '<p:pic>'
        '<p:nvPicPr>'
        '<p:cNvPr id="9001" name="Xelix logo"/>'
        '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        '<p:nvPr/>'
        '</p:nvPicPr>'
        '<p:blipFill>'
        f'<a:blip r:embed="{LOGO_RID}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</p:blipFill>'
        '<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{LOGO_W}" cy="{LOGO_H}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</p:spPr>'
        '</p:pic>'
    )


def _strip_existing_logo(slide_xml):
    """Remove any pic we previously injected (idempotency) so re-running is safe."""
    return re.sub(r'<p:pic>(?:(?!</p:pic>).)*?Xelix logo.*?</p:pic>', '',
                  slide_xml, flags=re.DOTALL)


def _add_logo(slide_xml, slide_w, slide_h):
    slide_xml = _strip_existing_logo(slide_xml)
    pic = _logo_pic_xml(slide_w, slide_h)
    # insert as last child of the slide's main <p:spTree>, before its close
    return slide_xml.replace('</p:spTree>', pic + '</p:spTree>', 1)


def _slide_rels_with_logo(rels_bytes, media_target):
    """Add (or replace) the logo relationship in a slide's .rels part."""
    rel_ns = NS_REL
    if rels_bytes is None:
        root = etree.fromstring(
            f'<Relationships xmlns="{rel_ns}"/>'.encode('utf-8'))
    else:
        root = etree.fromstring(rels_bytes)
    # remove any prior logo rel
    for rel in list(root):
        if rel.get('Id') == LOGO_RID:
            root.remove(rel)
    rel = etree.SubElement(root, '{%s}Relationship' % rel_ns)
    rel.set('Id', LOGO_RID)
    rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/'
                    'relationships/image')
    rel.set('Target', media_target)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8',
                          standalone=True)


# ── Pass: presentation size ────────────────────────────────────────────────────
def _normalize_size(prs_xml):
    s = prs_xml.decode('utf-8', 'ignore')
    m = re.search(r'<p:sldSz\b[^>]*/>', s)
    orig = None
    if m:
        cx = re.search(r'cx="(\d+)"', m.group(0))
        cy = re.search(r'cy="(\d+)"', m.group(0))
        if cx and cy:
            orig = (int(cx.group(1)), int(cy.group(1)))
        s = s.replace(m.group(0),
                      f'<p:sldSz cx="{SLIDE_W_169}" cy="{SLIDE_H_169}" type="screen16x9"/>')
    return s.encode('utf-8'), orig


def _ensure_png_default(ct_xml):
    s = ct_xml.decode('utf-8', 'ignore')
    if re.search(r'<Default\s+Extension="png"', s, re.IGNORECASE):
        return ct_xml
    s = s.replace('<Types xmlns="%s">' % NS_CT,
                  '<Types xmlns="%s"><Default Extension="png" '
                  'ContentType="image/png"/>' % NS_CT, 1)
    return s.encode('utf-8')


# ── Orchestrator ────────────────────────────────────────────────────────────────
def normalize_deck(in_path, out_path, *, logo_blue, logo_white,
                   unify_background=False, swap_fonts=True,
                   fix_contrast=None, add_logo=True, force_169=True):
    """Apply the Xelix brand layer to a deck.

    Default (safe) mode applies only the universally non-destructive layers:
    font -> Barlow and the wordmark bottom-right. Existing backgrounds, colours
    and layouts are untouched, so well-structured department slides survive
    intact while gaining a consistent typeface and logo.

    unify_background=True (full mode): also imposes the canonical white /
    hero-gradient background per slide regime, plus contrast-repair of text
    against it. Only use for decks known to be plain — it will recolour text
    and replace backgrounds.

    fix_contrast: None -> follow unify_background (off in safe mode, on in full
    mode). Set True/False to override. Contrast-repair on its own (without
    controlling the background) can darken white text that sits in a local dark
    panel, so it is not part of the safe default.
    """
    if fix_contrast is None:
        fix_contrast = unify_background
    report = {'slides': [], 'warnings': [], 'size': None,
              'mode': 'full' if unify_background else 'safe'}
    zin = zipfile.ZipFile(in_path, 'r')
    names = zin.namelist()

    slide_names = sorted(
        [n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
        key=lambda n: int(re.search(r'(\d+)', n).group(1)))

    blue_media = 'ppt/media/xelix_logo_blue.png'
    white_media = 'ppt/media/xelix_logo_white.png'
    used_blue = used_white = False

    transformed = {}
    for sn in slide_names:
        sxml = zin.read(sn).decode('utf-8', 'ignore')
        regime, confident = _effective_bg(zin, sn, sxml, SLIDE_W_169, SLIDE_H_169)
        dark = (regime == 'dark')

        if swap_fonts:
            sxml = _swap_fonts_in_runs(sxml)
        if unify_background:
            sxml = _set_background(sxml, dark)
            confident = True            # we now control the background, so we're sure
        # contrast-repair only when we know the effective background
        if fix_contrast and confident:
            sxml = _fix_text_contrast(sxml, dark)
        if add_logo:
            sxml = _add_logo(sxml, SLIDE_W_169, SLIDE_H_169)
            tgt = ('../media/xelix_logo_white.png' if dark
                   else '../media/xelix_logo_blue.png')
            rels_path = 'ppt/slides/_rels/%s.rels' % posixpath.basename(sn)
            rels_bytes = zin.read(rels_path) if rels_path in names else None
            transformed[rels_path] = _slide_rels_with_logo(rels_bytes, tgt)
            if dark:
                used_white = True
            else:
                used_blue = True

        transformed[sn] = sxml.encode('utf-8')
        report['slides'].append({
            'slide': posixpath.basename(sn),
            'regime': regime,
            'confident': confident,
            'logo': 'white' if dark else 'blue',
            'contrast_fixed': bool(fix_contrast and confident),
        })

    # theme parts
    if swap_fonts:
        for n in names:
            if re.match(r'ppt/theme/theme\d+\.xml$', n):
                transformed[n] = _patch_theme_fonts(zin.read(n))

    # presentation size
    prs = 'ppt/presentation.xml'
    if force_169 and prs in names:
        new_prs, orig = _normalize_size(zin.read(prs))
        transformed[prs] = new_prs
        report['size'] = orig
        if orig and abs(orig[0] / orig[1] - SLIDE_W_169 / SLIDE_H_169) > 0.02:
            report['warnings'].append(
                'Source aspect ratio %.3f differs from 16:9 — shapes were not '
                'rescaled, only the canvas. Review layout.' % (orig[0] / orig[1]))

    # content types (png default)
    ct = '[Content_Types].xml'
    if add_logo and ct in names:
        transformed[ct] = _ensure_png_default(zin.read(ct))

    # write output
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if n in transformed:
                zout.writestr(n, transformed[n])
            else:
                zout.writestr(n, zin.read(n))
        # remaining transformed parts not in original (e.g. new rels) 
        for n, data in transformed.items():
            if n not in names:
                zout.writestr(n, data)
        # add logo media
        if add_logo and used_blue:
            with open(logo_blue, 'rb') as f:
                zout.writestr(blue_media, f.read())
        if add_logo and used_white:
            with open(logo_white, 'rb') as f:
                zout.writestr(white_media, f.read())
    zin.close()
    return report


if __name__ == '__main__':
    import sys, json
    a = sys.argv
    rep = normalize_deck(a[1], a[2], logo_blue=a[3], logo_white=a[4])
    print(json.dumps(rep, indent=2))
