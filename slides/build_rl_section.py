#!/usr/bin/env python3
"""Build the RL section from remaining HKUST template prototypes.

Layouts used (auto-detected; deleted demo slides are skipped):
- Cover (first slide)
- Section divider Style A if present, else any section divider
- Title and Content: short title (box is ~4.3 in) + spaced body paragraphs
- Four-card Two Content: one idea per card, one sentence
- Wide two-column Two Content (title spans the slide): comparisons
"""

from __future__ import annotations

import shutil
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "templates" / "HKUST_Teaching.pptx"
OUTPUT = ROOT / "CourseProject.pptx"
IMAGES = ROOT / "images"

NAVY = RGBColor(0x00, 0x33, 0x66)
INK = RGBColor(0x26, 0x26, 0x26)
MUTED = RGBColor(0x6B, 0x6B, 0x6B)

A_T = qn("a:t")
A_P = qn("a:p")
A_R = qn("a:r")
A_RPR = qn("a:rPr")


def delete_slide(prs: Presentation, index: int) -> None:
    slide_id_list = prs.slides._sldIdLst
    r_id = slide_id_list[index].rId
    prs.part.drop_rel(r_id)
    del slide_id_list[index]


def duplicate_slide(prs: Presentation, index: int):
    """Copy a prototype. Drop shapes that add_slide inserts, or titles and placeholders double up."""
    source = prs.slides[index]
    new_slide = prs.slides.add_slide(source.slide_layout)
    tree = new_slide.shapes._spTree
    for child in list(tree):
        if child.tag in {qn("p:sp"), qn("p:pic"), qn("p:cxnSp"), qn("p:grpSp"), qn("p:graphicFrame")}:
            tree.remove(child)
    for shape in source.shapes:
        tree.insert_element_before(deepcopy(shape.element), "p:extLst")
    return new_slide


def slide_blob(slide) -> str:
    return "\n".join(
        shape.text_frame.text for shape in slide.shapes if shape.has_text_frame
    )


def set_run_text(paragraph, text: str) -> None:
    runs = paragraph._p.findall(A_R)
    if not runs:
        r = etree.SubElement(paragraph._p, A_R)
        rpr = etree.SubElement(r, A_RPR)
        rpr.set("lang", "en-US")
        t = etree.SubElement(r, A_T)
        t.text = text
        return
    texts = runs[0].findall(A_T)
    if texts:
        texts[0].text = text
        for extra in texts[1:]:
            extra.text = ""
    else:
        t = etree.SubElement(runs[0], A_T)
        t.text = text
    for extra_run in runs[1:]:
        for t in extra_run.findall(A_T):
            t.text = ""


def rewrite_paragraphs(shape, lines: list[str], *, blank_between: bool = False) -> None:
    """Replace paragraph text, cloning the first paragraph's formatting."""
    tf = shape.text_frame
    body = tf._txBody
    prototypes = body.findall(A_P)
    if not prototypes:
        return
    proto = deepcopy(prototypes[0])
    for child in list(body):
        if child.tag == A_P:
            body.remove(child)
    sequence: list[str] = []
    for i, line in enumerate(lines):
        if blank_between and i:
            sequence.append("")
        sequence.append(line)
    for line in sequence:
        p = deepcopy(proto)
        body.append(p)
        # paragraph wrapper
        from pptx.text.text import _Paragraph

        set_run_text(_Paragraph(p, tf), line)


def shape_by_name(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def _full_bleed_dark(shape_el) -> bool:
    """True for a full-slide picture or a dark solid panel (the dark HKUST variants)."""
    sp_pr = shape_el.find(qn("p:spPr"))
    if sp_pr is None:
        return False
    xfrm = sp_pr.find(qn("a:xfrm"))
    if xfrm is None:
        return False
    ext = xfrm.find(qn("a:ext"))
    if ext is None:
        return False
    if int(ext.get("cx")) < int(8.5 * 914400) or int(ext.get("cy")) < int(4.2 * 914400):
        return False
    if shape_el.find(".//" + qn("a:blip")) is not None:
        return True
    for srgb in shape_el.findall(".//" + qn("a:srgbClr")):
        val = (srgb.get("val") or "").upper()
        if len(val) == 6:
            r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
            if (r + g + b) / 3 < 90:
                return True
    solid = sp_pr.find(qn("a:solidFill"))
    if solid is not None:
        for sch in solid.findall(".//" + qn("a:schemeClr")):
            if sch.get("val") in {"tx1", "tx2", "dk1", "dk2"}:
                return True
    return False


def uses_dark_background(slide) -> bool:
    for part in (slide.slide_layout._element, slide.slide_layout.slide_master._element):
        c_sld = part.find(qn("p:cSld"))
        tree = c_sld.find(qn("p:spTree")) if c_sld is not None else None
        if tree is None:
            continue
        for child in tree:
            if child.tag in {qn("p:sp"), qn("p:pic")} and _full_bleed_dark(child):
                return True
    return False


def discover(prs: Presentation) -> dict:
    found = {
        "cover": 0,
        "bullets": None,
        "cards": None,
        "columns": None,
        "picture": None,
        "thank_you": None,
        "qa": None,
    }
    card_best = (-1, None)
    skipped_dark = []
    for i, slide in enumerate(prs.slides):
        blob = slide_blob(slide)
        if "Thank you" in blob:
            found["thank_you"] = i
        elif blob.strip().startswith("Q & A"):
            found["qa"] = i
        if i != 0 and uses_dark_background(slide):
            skipped_dark.append(i + 1)
            continue
        layout = slide.slide_layout.name if slide.slide_layout else ""
        if layout == "Title and Content" and "Lorem ipsum" in blob and found["bullets"] is None:
            names = {s.name for s in slide.shapes}
            if "TextBox 5" in names or "TextBox 3" in names:
                found["bullets"] = i
        if layout == "Two Content":
            subheads = [
                s
                for s in slide.shapes
                if s.has_text_frame and s.text_frame.text.strip().lower().startswith("subhead")
            ]
            if len(subheads) >= 4 and len(subheads) > card_best[0]:
                card_best = (len(subheads), i)
            names = {s.name for s in slide.shapes}
            if "Content Placeholder 2" in names and "Content Placeholder 3" in names:
                if found["columns"] is None:
                    found["columns"] = i
        kinds = {s.placeholder_format.type for s in slide.placeholders}
        if found["picture"] is None and PP_PLACEHOLDER.PICTURE in kinds and PP_PLACEHOLDER.BODY in kinds:
            found["picture"] = i
    found["cards"] = card_best[1]
    found["skipped_dark"] = skipped_dark
    missing = [k for k in ("bullets", "columns", "picture") if found[k] is None]
    if missing:
        raise SystemExit(
            "Template has no light-background "
            + " or ".join(missing)
            + " slide left. Dark section dividers and the dark content set are not used."
        )
    return found


def fill_cover(slide) -> None:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape.name.startswith("Title"):
            set_run_text(shape.text_frame.paragraphs[0], "CourseProject")
        elif "Placeholder" in shape.name or shape.name.startswith("Subtitle"):
            paras = shape.text_frame.paragraphs
            if paras:
                set_run_text(paras[0], "Human–Robot Handover")
            if len(paras) > 1:
                set_run_text(paras[1], "Created by: Tong Ren     Date: 2026.11.17")


def add_light_section(prs: Presentation, exemplar_index: int):
    """Light Section Header from the same master as the content slides. Dark dividers are not used."""
    master = prs.slides[exemplar_index].slide_layout.slide_master
    layout = None
    for candidate in master.slide_layouts:
        c_sld = candidate._element.find(qn("p:cSld"))
        if c_sld is not None and c_sld.get("name") == "Section Header":
            layout = candidate
            break
    if layout is None:
        raise SystemExit("No light Section Header layout on the content master.")
    slide = prs.slides.add_slide(layout)
    for shape in slide.placeholders:
        kind = shape.placeholder_format.type
        if kind == PP_PLACEHOLDER.TITLE:
            shape.text_frame.clear()
            add_muli_run(shape.text_frame.paragraphs[0], "Reinforcement Learning", "Muli ExtraBold", 32, True)
        elif kind == PP_PLACEHOLDER.BODY:
            shape.text_frame.clear()
            add_muli_run(shape.text_frame.paragraphs[0], "Part 1 of 3", "Muli", 18, False)
    return slide


def fill_bullets(slide, title: str, lines: list[str]) -> None:
    title_shape = shape_by_name(slide, "TextBox 3") or shape_by_name(slide, "TextBox 4")
    body = shape_by_name(slide, "TextBox 5")
    if title_shape is None or body is None:
        raise SystemExit("Bullet prototype is missing its title or body text box.")
    if len(title) > 22:
        raise SystemExit(f"Title too long for this layout ({len(title)} chars): {title}")
    set_run_text(title_shape.text_frame.paragraphs[0], title)
    rewrite_paragraphs(body, lines, blank_between=True)


def fill_cards(slide, title: str, cards: list[tuple[str, str]]) -> None:
    if len(title) > 22:
        raise SystemExit(f"Card-slide title too long: {title}")
    title_shape = next(
        (s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip() in ("Add title", title)),
        None,
    )
    # Title may be Title 1 or TextBox; pick the topmost short text box that says Add title.
    candidates = [
        s
        for s in slide.shapes
        if s.has_text_frame and "add title" in s.text_frame.text.lower()
    ]
    if candidates:
        title_shape = sorted(candidates, key=lambda s: s.top)[0]
        set_run_text(title_shape.text_frame.paragraphs[0], title)

    bodies = [
        s
        for s in slide.shapes
        if s.has_text_frame and s.text_frame.text.strip().lower().startswith("subhead")
    ]
    bodies.sort(key=lambda s: s.left)
    if len(bodies) < len(cards):
        raise SystemExit("Card prototype does not have enough cards.")
    for shape, (head, body) in zip(bodies, cards):
        paras = shape.text_frame.paragraphs
        set_run_text(paras[0], head)
        # Template cards are: heading, blank line, body.
        if len(paras) >= 3:
            set_run_text(paras[2], body)
        elif len(paras) >= 2:
            set_run_text(paras[1], body)


CS285_CREDIT = "Sergey Levine, CS 285: Deep Reinforcement Learning, UC Berkeley, Lecture 4."


def add_cs285_credit(slide) -> None:
    """Bottom-right, above the slide number and clear of the logo."""
    box = slide.shapes.add_textbox(
        Emu(int(2.40 * 914400)),
        Emu(int(4.66 * 914400)),
        Emu(int(7.28 * 914400)),
        Emu(int(0.34 * 914400)),
    )
    box.text_frame.word_wrap = False
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    add_muli_run(p, CS285_CREDIT, "Muli", 10, False, MUTED)


def place_lecture_page(slide, image: Path) -> None:
    """Drop the prototype's text and picture frame. The master keeps the school mark."""
    tree = slide.shapes._spTree
    for child in list(tree):
        if child.tag not in {qn("p:sp"), qn("p:pic"), qn("p:cxnSp"), qn("p:grpSp"), qn("p:graphicFrame")}:
            continue
        name = child.find(".//" + qn("p:cNvPr"))
        if name is not None and "Slide Number" in (name.get("name") or ""):
            continue
        tree.remove(child)
    slide_w, slide_h = 9144000, 5143500
    top = int(0.06 * 914400)
    bottom_keep = int(1.18 * 914400)  # blue mark sits in the last inch
    side = int(0.18 * 914400)
    avail_w = slide_w - 2 * side
    avail_h = slide_h - top - bottom_keep
    from PIL import Image

    with Image.open(image) as im:
        ratio = im.width / im.height
    if avail_w / avail_h > ratio:
        height = avail_h
        width = int(height * ratio)
    else:
        width = avail_w
        height = int(width / ratio)
    left = (slide_w - width) // 2
    slide.shapes.add_picture(str(image), Emu(left), Emu(top), width=Emu(width), height=Emu(height))
    add_cs285_credit(slide)


def place_definition_with_photo(slide, text_image: Path, photo: Path, credit: str) -> None:
    """Lecture definitions on the left, a related photo on the right, logo left clear."""
    tree = slide.shapes._spTree
    for child in list(tree):
        if child.tag not in {qn("p:sp"), qn("p:pic"), qn("p:cxnSp"), qn("p:grpSp"), qn("p:graphicFrame")}:
            continue
        name = child.find(".//" + qn("p:cNvPr"))
        if name is not None and "Slide Number" in (name.get("name") or ""):
            continue
        tree.remove(child)

    from PIL import Image

    slide_w, slide_h = 9144000, 5143500
    top = int(0.12 * 914400)
    bottom_keep = int(1.05 * 914400)
    left = int(0.12 * 914400)
    right = int(0.18 * 914400)
    gap = int(0.10 * 914400)
    avail_h = slide_h - top - bottom_keep
    caption_h = int(0.58 * 914400)

    with Image.open(photo) as im:
        photo_ratio = im.width / im.height
    photo_w = int(4.35 * 914400)
    photo_h = int(photo_w / photo_ratio)
    max_h = avail_h - caption_h - int(0.06 * 914400)
    if photo_h > max_h:
        photo_h = max_h
        photo_w = int(photo_h * photo_ratio)
    photo_left = slide_w - right - photo_w
    slide.shapes.add_picture(str(photo), Emu(photo_left), Emu(top), width=Emu(photo_w), height=Emu(photo_h))

    note = slide.shapes.add_textbox(
        Emu(photo_left),
        Emu(top + photo_h + int(0.04 * 914400)),
        Emu(photo_w),
        Emu(caption_h),
    )
    note.text_frame.word_wrap = True
    p = note.text_frame.paragraphs[0]
    add_muli_run(p, credit, "Muli", 9, False, MUTED)

    text_w = photo_left - gap - left
    with Image.open(text_image) as im:
        ratio = im.width / im.height
    if text_w / avail_h > ratio:
        height = avail_h
        width = int(height * ratio)
    else:
        width = text_w
        height = int(width / ratio)
    slide.shapes.add_picture(
        str(text_image),
        Emu(left),
        Emu(top + (avail_h - height) // 2),
        width=Emu(width),
        height=Emu(height),
    )
    add_cs285_credit(slide)


def add_banner(slide, text: str) -> None:
    """One line under the four cards, above the footer."""
    box = slide.shapes.add_textbox(
        Emu(int(0.4 * 914400)),
        Emu(int(4.15 * 914400)),
        Emu(int(9.2 * 914400)),
        Emu(int(0.45 * 914400)),
    )
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_muli_run(p, text, "Muli", 16, False, NAVY)


def add_muli_run(
    paragraph,
    text: str,
    face: str,
    size_pt: int,
    bold: bool,
    color: RGBColor | None = None,
    italic: bool = False,
) -> None:
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = face
    if color is not None:
        run.font.color.rgb = color
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = r_pr.find(qn(tag))
        if el is None:
            el = etree.SubElement(r_pr, qn(tag))
        el.set("typeface", face)


def fill_picture_slide(
    slide,
    title: str,
    lead: str,
    rows: list[tuple[str, str]],
    closing: str,
    tag: str,
    image: Path,
    credit: str,
    *,
    fit: str = "fill",
    crop_bottom: float = 0.0,
) -> None:
    """Text on the left, image on the right. Labels are the RL terms the example maps to."""
    if len(title) > 16:
        raise SystemExit(f"Picture-slide title runs into the image: {title}")
    title_shape = shape_by_name(slide, "TextBox 4")
    set_run_text(title_shape.text_frame.paragraphs[0], title)

    body = next(s for s in slide.placeholders if s.placeholder_format.type == PP_PLACEHOLDER.BODY)
    tf = body.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    add_muli_run(p, lead, "Muli ExtraBold", 18, True, NAVY)
    p.space_after = Pt(12)
    for label, text in rows:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        add_muli_run(p, f"{label}  ", "Muli ExtraBold", 14, True, NAVY)
        add_muli_run(p, text, "Muli", 14, False, INK)
    p = tf.add_paragraph()
    p.space_before = Pt(16)
    add_muli_run(p, closing, "Muli", 12, False, MUTED, italic=True)

    pill = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        body.left,
        Emu(int(3.95 * 914400)),
        body.width,
        Emu(int(0.42 * 914400)),
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = NAVY
    pill.line.fill.background()
    pill.shadow.inherit = False
    ptf = pill.text_frame
    ptf.word_wrap = True
    pp = ptf.paragraphs[0]
    pp.alignment = PP_ALIGN.CENTER
    add_muli_run(pp, tag, "Muli ExtraBold", 12, True, RGBColor(0xFF, 0xFF, 0xFF))

    frame = next(s for s in slide.placeholders if s.placeholder_format.type == PP_PLACEHOLDER.PICTURE)
    left, top, width, height = frame.left, frame.top, frame.width, frame.height
    if fit == "fill":
        pic = frame.insert_picture(str(image))
        pic.crop_bottom = max(pic.crop_bottom, crop_bottom)
    else:
        frame._element.getparent().remove(frame._element)
        from PIL import Image

        with Image.open(image) as im:
            ratio = im.width / (im.height * (1 - crop_bottom))
        if ratio > width / height:
            w, h = width, int(width / ratio)
        else:
            w, h = int(height * ratio), height
        pic = slide.shapes.add_picture(
            str(image), left + (width - w) // 2, top + (height - h) // 2, width=w, height=h
        )
        pic.crop_bottom = crop_bottom

    note = slide.shapes.add_textbox(left, top + height + Emu(int(0.04 * 914400)), width, Emu(int(0.28 * 914400)))
    note.text_frame.word_wrap = True
    np_ = note.text_frame.paragraphs[0]
    np_.alignment = PP_ALIGN.RIGHT
    add_muli_run(np_, credit, "Muli", 8, False, MUTED)


def _cell_props(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    return tcPr


def _set_cell_fill(cell, hex_color: str) -> None:
    tcPr = _cell_props(cell)
    for child in list(tcPr):
        if child.tag == qn("a:solidFill"):
            tcPr.remove(child)
    solid = etree.Element(qn("a:solidFill"))
    color = etree.SubElement(solid, qn("a:srgbClr"))
    color.set("val", hex_color)
    tcPr.insert(0, solid)


def _set_cell_borders(cell, hex_color: str = "D5DCE3") -> None:
    tcPr = _cell_props(cell)
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        for old in tcPr.findall(qn(f"a:{edge}")):
            tcPr.remove(old)
        line = etree.SubElement(tcPr, qn(f"a:{edge}"))
        line.set("w", "6350")
        line.set("cap", "flat")
        line.set("cmpd", "sng")
        line.set("algn", "ctr")
        fill = etree.SubElement(line, qn("a:solidFill"))
        color = etree.SubElement(fill, qn("a:srgbClr"))
        color.set("val", hex_color)
        etree.SubElement(line, qn("a:prstDash")).set("val", "solid")


def _write_cell(cell, parts: list[tuple[str, bool, RGBColor]], size: int, align) -> None:
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    for text, bold, color in parts:
        add_muli_run(p, text, "Muli", size, bold, color)
    _cell_props(cell).set("anchor", "ctr")
    cell.margin_left = Emu(int(0.08 * 914400))
    cell.margin_right = Emu(int(0.06 * 914400))
    cell.margin_top = Emu(int(0.02 * 914400))
    cell.margin_bottom = Emu(int(0.02 * 914400))


def fill_family_table(slide) -> None:
    """One taxonomy table. The first column merges rows that share an axis."""
    title = "Algorithm families"
    title_shape = shape_by_name(slide, "TextBox 3") or shape_by_name(slide, "TextBox 4")
    if title_shape is None:
        raise SystemExit("Family-table slide is missing its title box.")
    set_run_text(title_shape.text_frame.paragraphs[0], title)

    body = shape_by_name(slide, "TextBox 5")
    if body is not None:
        body._element.getparent().remove(body._element)

    note = slide.shapes.add_textbox(
        Emu(int(5.25 * 914400)),
        Emu(int(0.68 * 914400)),
        Emu(int(4.35 * 914400)),
        Emu(int(0.36 * 914400)),
    )
    note.text_frame.word_wrap = False
    p = note.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    add_muli_run(p, "Bold names are used in this course.", "Muli", 12, False, MUTED)

    rows_n, cols_n = 10, 3
    left = int(0.40 * 914400)
    top = int(1.28 * 914400)
    width = int(9.20 * 914400)
    height = int(3.18 * 914400)
    graphic = slide.shapes.add_table(rows_n, cols_n, Emu(left), Emu(top), Emu(width), Emu(height))
    table = graphic.table
    tbl_pr = table._tbl.tblPr
    style = tbl_pr.find(qn("a:tableStyleId"))
    if style is not None:
        tbl_pr.remove(style)

    table.columns[0].width = Emu(int(1.70 * 914400))
    table.columns[1].width = Emu(int(4.35 * 914400))
    table.columns[2].width = Emu(int(3.15 * 914400))
    row_h = Emu(height // rows_n)
    for row in table.rows:
        row.height = row_h

    white = RGBColor(0xFF, 0xFF, 0xFF)
    header = [
        "Axis",
        "Type",
        "Representative algorithms",
    ]
    for col, text in enumerate(header):
        cell = table.cell(0, col)
        _set_cell_fill(cell, "003366")
        _set_cell_borders(cell, "003366")
        _write_cell(cell, [(text, True, white)], 12, PP_ALIGN.LEFT if col else PP_ALIGN.CENTER)

    course = NAVY
    plain = INK
    body_rows = [
        (
            "Model",
            [("Model-based. Plan with a known or learned model.", False, plain)],
            [("MPC, PETS, MuZero", False, plain)],
        ),
        (
            None,
            [("Model-free. Learn from samples only.", False, plain)],
            [("DQN, ", False, plain), ("PPO", True, course), (", ", False, plain), ("SAC", True, course), (", ", False, plain), ("TD3", True, course)],
        ),
        (
            "What is learned",
            [("Value-based. Learn Q or V only.", False, plain)],
            [("Q-learning, DQN", False, plain)],
        ),
        (
            None,
            [("Policy-based. Learn the policy only.", False, plain)],
            [("REINFORCE", False, plain)],
        ),
        (
            None,
            [("Actor-critic. Learn a policy and a value together.", False, plain)],
            [("A2C, ", False, plain), ("PPO", True, course), (", DDPG, ", False, plain), ("TD3", True, course), (", ", False, plain), ("SAC", True, course)],
        ),
        (
            "Data reuse",
            [("On-policy. Use only the current policy's data.", False, plain)],
            [("REINFORCE, A2C, TRPO, ", False, plain), ("PPO", True, course)],
        ),
        (
            None,
            [("Off-policy. Reuse past data from a replay buffer.", False, plain)],
            [("DQN, DDPG, ", False, plain), ("TD3", True, course), (", ", False, plain), ("SAC", True, course)],
        ),
        (
            "Policy",
            [("Stochastic. One state can yield different actions.", False, plain)],
            [("A2C, ", False, plain), ("PPO", True, course), (", ", False, plain), ("SAC", True, course)],
        ),
        (
            None,
            [("Deterministic. One state maps to one action.", False, plain)],
            [("DDPG, ", False, plain), ("TD3", True, course)],
        ),
    ]
    spans = [(1, 2), (3, 5), (6, 7), (8, 9)]
    for offset, (axis, type_parts, algo_parts) in enumerate(body_rows, start=1):
        for col in range(3):
            cell = table.cell(offset, col)
            _set_cell_fill(cell, "F4F7FA" if col == 0 else "FFFFFF")
            _set_cell_borders(cell, "D5DCE3")
        if axis:
            _write_cell(table.cell(offset, 0), [(axis, True, NAVY)], 12, PP_ALIGN.CENTER)
        _write_cell(table.cell(offset, 1), type_parts, 12, PP_ALIGN.LEFT)
        _write_cell(table.cell(offset, 2), algo_parts, 12, PP_ALIGN.LEFT)

    for start, end in spans:
        table.cell(start, 0).merge(table.cell(end, 0))
        _set_cell_fill(table.cell(start, 0), "F4F7FA")
        _cell_props(table.cell(start, 0)).set("anchor", "ctr")


def fill_columns(slide, title: str, left_head: str, left: list[str], right_head: str, right: list[str]) -> None:
    title_shape = next((s for s in slide.shapes if s.name.startswith("Title")), None)
    if title_shape is None:
        title_shape = shape_by_name(slide, "TextBox 4")
    if title_shape is not None:
        if title_shape.width < Emu(6 * 914400) and len(title) > 18:
            raise SystemExit(f"Column title too long for the light layout: {title}")
        set_run_text(title_shape.text_frame.paragraphs[0], title)
    left_shape = shape_by_name(slide, "Content Placeholder 2")
    right_shape = shape_by_name(slide, "Content Placeholder 3")
    if left_shape is None or right_shape is None:
        raise SystemExit("Two-column prototype is missing content placeholders.")

    def write(shape, head: str, lines: list[str]) -> None:
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        add_muli_run(p, head, "Muli ExtraBold", 20, True)
        for line in lines:
            p = tf.add_paragraph()
            add_muli_run(p, line, "Muli", 16, False)

    write(left_shape, left_head, left)
    write(right_shape, right_head, right)


def main() -> None:
    shutil.copy(TEMPLATE, OUTPUT)
    prs = Presentation(str(OUTPUT))
    proto = discover(prs)
    print(
        "Light prototypes (1-based):",
        {k: (None if v is None else v + 1) for k, v in proto.items() if k != "skipped_dark"},
    )
    print("Skipped dark template slides:", proto["skipped_dark"])

    fill_cover(prs.slides[proto["cover"]])
    before = len(prs.slides)
    add_light_section(prs, proto["bullets"])

    s = duplicate_slide(prs, proto["picture"])
    fill_picture_slide(
        s,
        "Tesla Autopilot",
        "A chain of decisions.",
        [
            ("See", "8 cameras around the car"),
            ("Act", "steer, accelerate, brake"),
            ("Goal", "arrive safely and smoothly"),
        ],
        "Learned mostly from human driving video. Can it get better than its teachers?",
        "Partially observed: cameras miss things",
        IMAGES / "tesla_autopilot_model_x.jpg",
        "Autopilot engaged in a Model X. Photo: Isnoop, CC BY-SA 4.0, Wikimedia Commons",
    )

    s = duplicate_slide(prs, proto["picture"])
    fill_picture_slide(
        s,
        "AlphaGo, 2016",
        "Beat Lee Sedol 4–1.",
        [
            ("See", "the whole board, nothing hidden"),
            ("Act", "where to place the next stone"),
            ("Reward", "+1 win, −1 loss, at the end"),
        ],
        "Started from human games, then improved by playing against itself.",
        "Fully observed: the board is the state",
        IMAGES / "alphago_lee_sedol_game1.png",
        "Lee Sedol (black) vs AlphaGo (white), Game 1. Diagram: Wesalius, CC BY-SA 4.0, Wikimedia Commons",
        fit="contain",
        crop_bottom=0.06,
    )

    # Original CS285 pages, scaled to leave the HKUST mark in the bottom-left.
    s = duplicate_slide(prs, proto["picture"])
    place_lecture_page(s, IMAGES / "lec4_terminology.png")
    s = duplicate_slide(prs, proto["picture"])
    place_lecture_page(s, IMAGES / "lec4_reward.png")

    s = duplicate_slide(prs, proto["picture"])
    place_definition_with_photo(
        s,
        IMAGES / "lec4_mdp_def_text.png",
        IMAGES / "markov_decision_process.png",
        "Markov decision process. Diagram: waldoalvarez, CC BY-SA 4.0, Wikimedia Commons",
    )

    s = duplicate_slide(prs, proto["bullets"])
    fill_family_table(s)

    if proto["cards"] is not None:
        s = duplicate_slide(prs, proto["cards"])
        fill_cards(
            s,
            "Four ideas",
            [
                ("On-policy", "Current policy only. PPO."),
                ("Off-policy", "Replay buffer. SAC and TD3."),
                ("Explore", "Entropy bonus in SAC."),
                ("Exploit", "Use the best policy."),
            ],
        )

    s = duplicate_slide(prs, proto["columns"])
    fill_columns(
        s,
        "Algorithms",
        "PPO",
        [
            "On-policy",
            "Fresh data each update",
            "Stable, easy default",
            "Notebook exercise",
        ],
        "SAC  (and TD3)",
        [
            "Off-policy, replay buffer",
            "Strong on continuous control",
            "We train handover with SAC",
            "TD3: deterministic actor",
        ],
    )

    s = duplicate_slide(prs, proto["bullets"])
    fill_bullets(
        s,
        "How we train",
        [
            "SB3: SAC(\"MlpPolicy\", env).learn(…).",
            "Test with predict(obs, deterministic=True).",
            "Handover setup: learning_starts 1000.",
            "Entropy coefficient ent_coef = 0.005.",
        ],
    )

    s = duplicate_slide(prs, proto["bullets"])
    fill_bullets(
        s,
        "Check the run",
        [
            "The same reward can fail on another seed.",
            "Save checkpoints and read the learning curve.",
            "Evaluate in simulation before the live demo.",
            "Notebook Part 6 is this loop on our env.",
        ],
    )

    from sim_slides import add_simulation_slides
    from proj_slides import add_project_slides

    add_simulation_slides(prs, proto["bullets"])
    add_project_slides(prs, proto["bullets"])

    for key in ("thank_you", "qa"):
        if proto[key] is not None:
            duplicate_slide(prs, proto[key])

    keep = {proto["cover"]} | set(range(before, len(prs.slides)))
    for i in sorted(set(range(len(prs.slides))) - keep, reverse=True):
        delete_slide(prs, i)

    # PowerPoint recomputes slide-number fields, but Quick Look and Keynote show the cached text.
    for number, slide in enumerate(prs.slides, 1):
        for fld in slide._element.iter(qn("a:fld")):
            if fld.get("type") == "slidenum":
                for t in fld.iter(A_T):
                    t.text = str(number)

    prs.save(str(OUTPUT))
    print(f"Wrote {len(prs.slides)} slides → {OUTPUT}")


if __name__ == "__main__":
    main()
