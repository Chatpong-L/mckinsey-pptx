"""Editorial slide templates for the human-crafted design pass.

Implements the design-audit specs: flat fills (no shadows), drawn diagrams
instead of box grids, one cyan accent per slide, hairline separators, and the
cover's skyline art returning only at the dividers and the close.
"""
from __future__ import annotations
from typing import Optional, Sequence

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from lxml import etree

from mckinsey_pptx.base import (
    blank_slide, add_chrome, add_rect, add_oval, add_line, add_textbox,
    write_paragraph, enable_text_shrink,
)
from mckinsey_pptx.theme import Theme, rgb
from mckinsey_pptx.builder import _REGISTRY

from max_theme import MAX_THEME, PH_BORDER, PH_FILL, PH_TEXT
from max_slides import placeholder_box, ASSETS, LOGO_WHITE

STRIP_PANEL = f"{ASSETS}/gen/skyline-strip-panel.png"
STRIP_WIDE = f"{ASSETS}/gen/skyline-strip-wide.png"

STEEL = rgb("55708C")      # muted steel for de-emphasized data
SLATE = rgb("6B7B8C")      # decline bars
PALE_ROW = rgb("F4F7FA")

import os


def cover_crop(path, w_in, h_in):
    """Center-crop an image file to the w:h aspect of the target box and
    cache the crop next to the original. Returns the cropped path."""
    from PIL import Image
    out = f"{os.path.splitext(path)[0]}.crop{w_in:g}x{h_in:g}.png"
    if os.path.exists(out) and os.path.getmtime(out) >= os.path.getmtime(path):
        return out
    im = Image.open(path)
    target = w_in / h_in
    w, h = im.size
    if w / h > target:
        nw = int(h * target)
        im = im.crop(((w - nw) // 2, 0, (w + nw) // 2, h))
    else:
        nh = int(w / target)
        im = im.crop((0, (h - nh) // 2, w, (h + nh) // 2))
    im.save(out)
    return out


def slide_icon(slide, path, x, y, size):
    """Place a square icon PNG (transparent background) at x,y."""
    return slide.shapes.add_picture(path, Inches(x), Inches(y),
                                    width=Inches(size), height=Inches(size))


# ---------- global flattening ----------

def flatten_all_shadows(prs):
    """Force truly flat fills: an explicit empty effect list on every shape,
    and strip the preset <p:style> effect reference that LibreOffice still
    honors even when effectLst is empty."""
    for slide in prs.slides:
        for shape in slide.shapes:
            el = shape._element
            try:
                spPr = el.spPr
            except AttributeError:
                continue
            if spPr is None:
                continue
            for e in spPr.findall(qn("a:effectLst")):
                spPr.remove(e)
            etree.SubElement(spPr, qn("a:effectLst"))
            for style in el.findall(qn("p:style")):
                el.remove(style)


# ---------- shared devices ----------

def takeaway_line(slide, text, theme: Theme = MAX_THEME, top=1.32):
    """One bold editorial takeaway under the title, replacing the right rail."""
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    tb = add_textbox(slide, layout.margin_left_in, top,
                     layout.slide_width_in - layout.margin_left_in
                     - layout.margin_right_in, 0.34)
    write_paragraph(tb.text_frame, text, size=typo.body_size + 2, bold=True,
                    color=pal.dark_navy, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)


def divider_strip(slide, theme: Theme = MAX_THEME):
    """Low-contrast skyline strip along the bottom of a divider's navy panel."""
    slide.shapes.add_picture(STRIP_PANEL, Inches(0), Inches(6.9),
                             width=Inches(4.5), height=Inches(0.6))


def add_art_divider(prs, *,
                    section_number: str,
                    section_title: str,
                    subtitle: Optional[str] = None,
                    art_path: Optional[str] = None,
                    page_number=None, section_marker=None,
                    source=None, footnote=None,
                    theme: Theme = MAX_THEME):
    """Section divider whose left panel is a themed art image (cover-art
    sibling) with the big section number over it. Falls back to the flat
    navy panel when art is missing."""
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    panel_w = 4.5
    if art_path and os.path.exists(art_path):
        crop = cover_crop(art_path, panel_w, layout.slide_height_in)
        slide.shapes.add_picture(crop, Inches(0), Inches(0),
                                 width=Inches(panel_w),
                                 height=Inches(layout.slide_height_in))
        # gentle navy scrim behind the number so it always reads
        tb = add_textbox(slide, 0.5, layout.slide_height_in / 2 - 1.5,
                         panel_w - 1.0, 2.0, anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_rect(slide, 0, 0, panel_w, layout.slide_height_in,
                 fill=pal.deep_navy)
        divider_strip(slide, theme)
        tb = add_textbox(slide, 0.5, layout.slide_height_in / 2 - 1.5,
                         panel_w - 1.0, 2.0, anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, str(section_number),
                    size=typo.title_size + 56, bold=True,
                    color=pal.bright_blue, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    right_left = panel_w + 0.6
    right_w = layout.slide_width_in - right_left - layout.margin_right_in
    tb = add_textbox(slide, right_left, layout.slide_height_in / 2 - 1.0,
                     right_w, 1.4, anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, section_title,
                    size=typo.title_size + 8, bold=True,
                    color=pal.text_dark, family=typo.family, first=True)
    add_line(slide, right_left, layout.slide_height_in / 2 + 0.40,
             right_left + 1.6, layout.slide_height_in / 2 + 0.40,
             color=pal.bright_blue, width_pt=2.5)
    if subtitle:
        tb = add_textbox(slide, right_left,
                         layout.slide_height_in / 2 + 0.55, right_w, 1.0)
        write_paragraph(tb.text_frame, subtitle,
                        size=typo.body_size + 2, color=pal.footer_gray,
                        family=typo.family, first=True)
    return slide


def closing_strip(slide, theme: Theme = MAX_THEME, height=0.8):
    """Full-width skyline strip for the closing navy slides."""
    layout = theme.layout
    slide.shapes.add_picture(STRIP_WIDE, Inches(0),
                             Inches(layout.slide_height_in - height),
                             width=Inches(layout.slide_width_in),
                             height=Inches(height))


def browser_frame(slide, left, top, width, height, theme: Theme = MAX_THEME):
    """Navy browser chrome: top bar with three dots + hairline outline.
    Returns the inner content box (l, t, w, h)."""
    pal = theme.palette
    bar_h = 0.30
    add_rect(slide, left, top, width, bar_h, fill=pal.dark_navy)
    for i in range(3):
        add_oval(slide, left + 0.14 + i * 0.16, top + 0.105, 0.09, 0.09,
                 fill=pal.light_blue if i == 2 else pal.mid_blue)
    add_rect(slide, left, top + bar_h, width, height - bar_h, fill=None,
             line=pal.grid_gray, line_width=1.0)
    return (left, top + bar_h, width, height - bar_h)


def _kicker(slide, text, theme, top=1.32, left=None):
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    tb = add_textbox(slide, left if left is not None else layout.margin_left_in,
                     top, 6.0, 0.28)
    write_paragraph(tb.text_frame, text.upper(), size=typo.small_size,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    first=True)


# ---------- 03 · route map agenda ----------

def add_route_map(prs, *,
                  title: str,
                  subtitle: Optional[str] = None,
                  stops: Sequence[dict],
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    """Horizontal 6-node route echoing the bottom tracker.
    stops: [{keyword, descriptor}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.32, 8.0, 0.30)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)

    n = len(stops)
    mid_y = 4.15
    left0 = layout.margin_left_in + 0.55
    right0 = layout.slide_width_in - layout.margin_right_in - 0.55
    add_line(slide, left0, mid_y, right0, mid_y, color=pal.dark_navy,
             width_pt=2.0)
    step = (right0 - left0) / (n - 1)
    d = 0.55
    for i, st in enumerate(stops):
        cx = left0 + i * step
        fill = pal.bright_blue if i == n - 1 else pal.dark_navy
        add_oval(slide, cx - d / 2, mid_y - d / 2, d, d, fill=fill)
        tb = add_textbox(slide, cx - d / 2, mid_y - d / 2, d, d,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, str(i + 1), size=typo.body_size + 2,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        above = (i % 2 == 0)
        block_top = mid_y - 1.55 if above else mid_y + 0.45
        add_line(slide, cx, mid_y - d / 2 - 0.08 if above else mid_y + d / 2,
                 cx, block_top + (1.0 if above else 0.0),
                 color=pal.grid_gray, width_pt=0.75)
        lx = max(0.25, min(cx - 1.05, layout.slide_width_in - 2.35))
        tb = add_textbox(slide, lx, block_top, 2.1, 0.34)
        write_paragraph(tb.text_frame, st["keyword"], size=typo.body_size + 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, lx, block_top + 0.36, 2.1, 0.55)
        write_paragraph(tb.text_frame, st["descriptor"], size=typo.small_size,
                        color=pal.footer_gray, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- 04 · asymmetric speakers ----------

def add_speaker_panels(prs, *,
                       title: str = "Your guides tonight",
                       speakers: Sequence[dict],
                       page_number=None, section_marker=None,
                       source=None, footnote=None,
                       theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    top = 1.65
    panel_h = layout.footer_top_in - top - 0.25
    panel_w = 2.9
    gap = 0.35
    for i, sp in enumerate(speakers):
        left = layout.margin_left_in + i * (panel_w + gap)
        photo_h = panel_h * 0.62
        placeholder_box(slide, left, top, panel_w, photo_h,
                        sp.get("photo_label", "Speaker photo"), theme=theme)
        band_h = 0.95
        add_rect(slide, left, top + photo_h, panel_w, band_h,
                 fill=pal.dark_navy)
        tb = add_textbox(slide, left + 0.2, top + photo_h + 0.12,
                         panel_w - 0.4, 0.40)
        write_paragraph(tb.text_frame, sp["name"], size=typo.body_size + 3,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.2, top + photo_h + 0.52,
                         panel_w - 0.4, 0.36)
        write_paragraph(tb.text_frame, sp["role"], size=typo.body_size - 1,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)

    # Shared credential column right of the panels
    cleft = layout.margin_left_in + 2 * panel_w + gap + 0.6
    cwidth = layout.slide_width_in - layout.margin_right_in - cleft
    cy = top + 0.25
    for sp in speakers:
        tb = add_textbox(slide, cleft, cy, cwidth, 0.28)
        write_paragraph(tb.text_frame, sp["name"].upper(),
                        size=typo.small_size, bold=True,
                        color=pal.footer_gray, family=typo.family, first=True)
        add_line(slide, cleft, cy + 0.32, cleft + 1.2, cy + 0.32,
                 color=pal.bright_blue, width_pt=1.5)
        tb = add_textbox(slide, cleft, cy + 0.45, cwidth, 1.15)
        for j, bl in enumerate(sp.get("bullets", [])[:2]):
            write_paragraph(tb.text_frame, bl, size=typo.body_size + 1,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=6)
        enable_text_shrink(tb.text_frame)
        cy += 1.85
    return slide


# ---------- 07/08 · horizontal ranked bars ----------

def add_hbar_ranked(prs, *,
                    title: str,
                    takeaway: str,
                    items: Sequence[dict],
                    direction: str = "right",
                    annotation: Optional[dict] = None,
                    unit_note: str = "",
                    page_number=None, section_marker=None,
                    source=None, footnote=None,
                    theme: Theme = MAX_THEME):
    """Horizontal ranked bars. items: [{label, value, display}].
    direction 'right' = growth (navy ramp); 'left' = decline (slate, bars
    extend left from a right baseline)."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    takeaway_line(slide, takeaway, theme)
    if unit_note:
        tb = add_textbox(slide, layout.margin_left_in, 1.72, 8.0, 0.26)
        write_paragraph(tb.text_frame, unit_note, size=typo.small_size,
                        color=pal.footer_gray, family=typo.family, first=True)

    top = 2.15
    row_h = 0.62
    gap = 0.28
    max_val = max(abs(it["value"]) for it in items)
    label_w = 2.9
    bar_area = (layout.slide_width_in - layout.margin_left_in
                - layout.margin_right_in - label_w - 1.0)
    ramp = [rgb("0B2E5C"), rgb("14406F"), rgb("1D5282"), rgb("266495"),
            rgb("2F76A8")]
    for i, it in enumerate(items):
        y = top + i * (row_h + gap)
        bl = abs(it["value"]) / max_val * bar_area
        if direction == "right":
            lx = layout.margin_left_in
            tb = add_textbox(slide, lx, y, label_w - 0.15, row_h,
                             anchor=MSO_ANCHOR.MIDDLE)
            write_paragraph(tb.text_frame, it["label"], size=typo.body_size + 1,
                            bold=True, color=pal.dark_navy,
                            family=typo.family, first=True)
            enable_text_shrink(tb.text_frame)
            bx = lx + label_w
            add_rect(slide, bx, y, bl, row_h, fill=ramp[min(i, len(ramp) - 1)])
            tb = add_textbox(slide, bx + bl + 0.12, y, 0.95, row_h,
                             anchor=MSO_ANCHOR.MIDDLE)
            write_paragraph(tb.text_frame, it["display"],
                            size=typo.body_size + 2, bold=True,
                            color=pal.dark_navy, family=typo.family,
                            first=True)
        else:
            rx = layout.slide_width_in - layout.margin_right_in
            tb = add_textbox(slide, rx - label_w + 0.15, y, label_w - 0.15,
                             row_h, anchor=MSO_ANCHOR.MIDDLE)
            write_paragraph(tb.text_frame, it["label"], size=typo.body_size + 1,
                            bold=True, color=pal.dark_navy,
                            family=typo.family, align=PP_ALIGN.RIGHT,
                            first=True)
            enable_text_shrink(tb.text_frame)
            bx = rx - label_w - bl
            add_rect(slide, bx, y, bl, row_h, fill=SLATE)
            tb = add_textbox(slide, bx - 1.07, y, 0.95, row_h,
                             anchor=MSO_ANCHOR.MIDDLE)
            write_paragraph(tb.text_frame, it["display"],
                            size=typo.body_size + 2, bold=True,
                            color=SLATE, family=typo.family,
                            align=PP_ALIGN.RIGHT, first=True)
    if annotation:
        ax = layout.margin_left_in + label_w + annotation.get("x", 3.2)
        ay = top + annotation.get("row", 2) * (row_h + gap) + 0.05
        aw = annotation.get("w", 3.3)
        add_rect(slide, ax, ay, aw, 0.52, fill=pal.white,
                 line=pal.bright_blue, line_width=1.0)
        tb = add_textbox(slide, ax + 0.12, ay, aw - 0.24, 0.52,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, annotation["text"],
                        size=typo.small_size, color=pal.text_dark,
                        family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- 09 · sector stat matrix ----------

def add_sector_matrix(prs, *,
                      title: str,
                      subtitle: Optional[str] = None,
                      rows: Sequence[dict],
                      callout: Optional[str] = None,
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """Hairline stat rows: sector | count | growth mini-bar | margin swing.
    rows: [{name, count, growth (int %), margin_from, margin_to, note}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.30, 10.0, 0.30)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)

    left = layout.margin_left_in
    total_w = layout.slide_width_in - left - layout.margin_right_in
    c_name, c_count, c_growth = left, left + 3.1, left + 5.8
    c_margin = left + 9.5
    top = 1.85
    # header row
    for x, label in ((c_name, "SECTOR"), (c_count, "OPERATING COMPANIES"),
                     (c_growth, "REVENUE GROWTH SINCE FY2020"),
                     (c_margin, "NET MARGIN SWING")):
        tb = add_textbox(slide, x, top, 3.2, 0.26)
        write_paragraph(tb.text_frame, label, size=typo.small_size,
                        color=pal.footer_gray, family=typo.family, first=True)
    add_line(slide, left, top + 0.32, left + total_w, top + 0.32,
             color=pal.dark_navy, width_pt=1.0)

    row_h = 1.02
    max_growth = max(r["growth"] for r in rows)
    y = top + 0.45
    for i, r in enumerate(rows):
        cy = y + row_h / 2
        name_x = c_name
        if r.get("icon") and os.path.exists(r["icon"]):
            slide_icon(slide, r["icon"], c_name, cy - 0.27, 0.54)
            name_x = c_name + 0.72
        tb = add_textbox(slide, name_x, cy - 0.30, 2.9 - (name_x - c_name),
                         0.60, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["name"], size=typo.body_size + 3,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, c_count, cy - 0.32, 2.4, 0.44,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["count"], size=typo.title_size - 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        tb = add_textbox(slide, c_count, cy + 0.12, 2.4, 0.26)
        write_paragraph(tb.text_frame, r.get("note", ""), size=typo.small_size,
                        color=pal.footer_gray, family=typo.family, first=True)
        # growth mini-bar
        bar_max = 2.5
        bw = max(0.15, r["growth"] / max_growth * bar_max)
        add_rect(slide, c_growth, cy - 0.14, bw, 0.28, fill=pal.bright_blue)
        tb = add_textbox(slide, c_growth + bw + 0.10, cy - 0.20, 0.95, 0.40,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, f"+{r['growth']}%",
                        size=typo.body_size + 2, bold=True,
                        color=pal.dark_navy, family=typo.family, first=True)
        # margin swing
        tb = add_textbox(slide, c_margin, cy - 0.20, 2.9, 0.40,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame,
                        f"{r['margin_from']} → {r['margin_to']}",
                        size=typo.body_size + 1, bold=True,
                        color=pal.text_dark, family=typo.family, first=True)
        if i < len(rows) - 1:
            add_line(slide, left, y + row_h, left + total_w, y + row_h,
                     color=pal.grid_gray, width_pt=0.5)
        y += row_h + 0.10
    if callout:
        tb = add_textbox(slide, left, y + 0.05, total_w, 0.32)
        write_paragraph(tb.text_frame, callout, size=typo.body_size, bold=True,
                        color=pal.bright_blue, family=typo.family, first=True)
    return slide


# ---------- 10 · fork in the road ----------

def add_fork_road(prs, *,
                  title: str,
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    # Start node
    sx, sy = 1.15, 4.15
    add_oval(slide, sx - 0.5, sy - 0.5, 1.0, 1.0, fill=pal.dark_navy)
    tb = add_textbox(slide, sx - 0.5, sy - 0.5, 1.0, 1.0,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "You today", size=typo.small_size,
                    bold=True, color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    end_x = 11.6
    # ORGANIC path: long shallow gray with 4 nodes
    oy = 2.65
    add_line(slide, sx + 0.42, sy - 0.25, 2.4, oy, color=STEEL, width_pt=2.5)
    add_line(slide, 2.4, oy, end_x, oy, color=STEEL, width_pt=2.5)
    org_nodes = ["Build sales", "Hire a team", "Reach scale", "Enter market"]
    for i, lab in enumerate(org_nodes):
        nx = 3.3 + i * 2.15
        add_oval(slide, nx - 0.10, oy - 0.10, 0.20, 0.20, fill=STEEL)
        tb = add_textbox(slide, nx - 0.85, oy - 0.55, 1.7, 0.3)
        write_paragraph(tb.text_frame, lab, size=typo.small_size,
                        color=pal.footer_gray, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
    tb = add_textbox(slide, 1.7, oy - 1.05, 2.4, 0.32)
    write_paragraph(tb.text_frame, "ORGANIC: build it",
                    size=typo.body_size + 1, bold=True, color=STEEL,
                    family=typo.family, first=True)

    # INORGANIC path: short steep cyan chevrons
    iy = 5.55
    add_line(slide, sx + 0.42, sy + 0.25, 2.3, iy, color=pal.bright_blue,
             width_pt=2.5)
    chev_w, chev_h = 1.15, 0.72
    for i in range(2):
        cx = 2.45 + i * (chev_w - 0.18)
        s = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(cx),
                                   Inches(iy - chev_h / 2), Inches(chev_w),
                                   Inches(chev_h))
        s.shadow.inherit = False
        s.fill.solid(); s.fill.fore_color.rgb = pal.bright_blue
        s.line.fill.background()
    tb = add_textbox(slide, 2.45, iy + 0.55, 3.6, 0.32)
    write_paragraph(tb.text_frame, "INORGANIC: buy it",
                    size=typo.body_size + 1, bold=True, color=pal.bright_blue,
                    family=typo.family, first=True)
    # Arrival square (shared destination size) reached early
    arr_x = 4.75
    add_rect(slide, arr_x, iy - 0.30, 0.6, 0.6, fill=pal.dark_navy)
    add_line(slide, arr_x + 0.6, iy, end_x, iy, color=pal.grid_gray,
             width_pt=1.0)
    # Cyan flag at arrival
    add_line(slide, arr_x + 0.30, iy - 1.15, arr_x + 0.30, iy - 0.30,
             color=pal.bright_blue, width_pt=1.5)
    tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                 Inches(arr_x + 0.30), Inches(iy - 1.18),
                                 Inches(0.5), Inches(0.3))
    tri.shadow.inherit = False
    tri.rotation = 90
    tri.fill.solid(); tri.fill.fore_color.rgb = pal.bright_blue
    tri.line.fill.background()
    tb = add_textbox(slide, arr_x + 0.9, iy - 1.30, 4.6, 0.55)
    write_paragraph(tb.text_frame,
                    "Day one: revenue, staff, licences, customers",
                    size=typo.body_size, bold=True, color=pal.dark_navy,
                    family=typo.family, first=True)
    # Organic arrival square (far right)
    add_rect(slide, end_x, oy - 0.30, 0.6, 0.6, fill=STEEL)

    # Time axis
    ty = 6.45
    add_line(slide, sx, ty, end_x + 0.6, ty, color=pal.grid_gray,
             width_pt=1.0)
    for x, lab in ((sx, "Year 0"), (arr_x + 0.30, "Months"),
                   (end_x + 0.3, "Year 5+")):
        add_line(slide, x, ty - 0.06, x, ty + 0.06, color=pal.footer_gray,
                 width_pt=1.0)
        tb = add_textbox(slide, x - 0.6, ty + 0.10, 1.2, 0.26)
        write_paragraph(tb.text_frame, lab, size=typo.small_size,
                        color=pal.footer_gray, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
    tb = add_textbox(slide, layout.margin_left_in, 1.32, 11.0, 0.34)
    write_paragraph(tb.text_frame,
                    "Same destination. The skill on the fast road is picking the right target.",
                    size=typo.body_size + 2, bold=True, color=pal.dark_navy,
                    family=typo.family, first=True)
    return slide


# ---------- 11/42 · dimension transformation table ----------

def add_dimension_table(prs, *,
                        title: str,
                        left_header: str, right_header: str,
                        rows: Sequence[dict],
                        chip: Optional[dict] = None,
                        check_right: bool = False,
                        page_number=None, section_marker=None,
                        source=None, footnote=None,
                        theme: Theme = MAX_THEME):
    """Spine ledger: rows: [{dim, left, right}]. Optional bottom-right stat
    chip {text}. check_right draws a small cyan tick before right cells."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    left = layout.margin_left_in + 0.2
    total_w = layout.slide_width_in - left - layout.margin_right_in - 0.2
    spine_w = 1.9
    cell_w = (total_w - spine_w) / 2
    spine_x = left + cell_w
    top = 1.95
    row_h = 0.95

    # headers
    tb = add_textbox(slide, left, top - 0.42, cell_w, 0.30)
    write_paragraph(tb.text_frame, left_header.upper(), size=typo.small_size,
                    bold=True, color=pal.footer_gray, family=typo.family,
                    first=True)
    tb = add_textbox(slide, spine_x + spine_w, top - 0.42, cell_w, 0.30)
    write_paragraph(tb.text_frame, right_header.upper(), size=typo.small_size,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    first=True)

    for i, r in enumerate(rows):
        y = top + i * row_h
        cy = y + row_h / 2
        tb = add_textbox(slide, left, y, cell_w - 0.3, row_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["left"], size=typo.body_size + 1,
                        italic=True, color=pal.footer_gray,
                        family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        # spine
        tb = add_textbox(slide, spine_x, y, spine_w, row_h - 0.28,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["dim"].upper(), size=typo.small_size,
                        bold=True, color=pal.footer_gray, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        ay = y + row_h - 0.30
        add_line(slide, spine_x + spine_w / 2 - 0.22, ay,
                 spine_x + spine_w / 2 + 0.14, ay, color=pal.bright_blue,
                 width_pt=1.5)
        tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                     Inches(spine_x + spine_w / 2 + 0.10),
                                     Inches(ay - 0.055), Inches(0.11),
                                     Inches(0.11))
        tri.shadow.inherit = False
        tri.rotation = 90
        tri.fill.solid(); tri.fill.fore_color.rgb = pal.bright_blue
        tri.line.fill.background()
        # right cell
        rx = spine_x + spine_w
        if check_right:
            add_line(slide, rx + 0.02, cy + 0.02, rx + 0.09, cy + 0.09,
                     color=pal.bright_blue, width_pt=2.0)
            add_line(slide, rx + 0.09, cy + 0.09, rx + 0.22, cy - 0.08,
                     color=pal.bright_blue, width_pt=2.0)
        tb = add_textbox(slide, rx + (0.34 if check_right else 0.0), y,
                         cell_w - 0.4, row_h, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["right"], size=typo.body_size + 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        if i < len(rows) - 1:
            add_line(slide, left, y + row_h, left + total_w, y + row_h,
                     color=pal.grid_gray, width_pt=0.5)

    if chip:
        cw = 4.6
        cx = layout.slide_width_in - layout.margin_right_in - cw
        cyy = top + len(rows) * row_h + 0.25
        add_rect(slide, cx, cyy, cw, 0.62, fill=pal.dark_navy)
        tb = add_textbox(slide, cx + 0.2, cyy, cw - 0.4, 0.62,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, chip["text"], size=typo.body_size,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- 13 · navy hero with waffle ----------

def add_stat_hero_navy(prs, *,
                       title_eyebrow: str,
                       stat: str,
                       stat_label: str,
                       waffle_filled: int = 81,
                       waffle_caption: str,
                       closing: str,
                       bg_path: Optional[str] = None,
                       page_number=None, section_marker=None,
                       source=None, footnote=None,
                       theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    add_rect(slide, 0, 0, layout.slide_width_in, layout.slide_height_in,
             fill=rgb("022859"))
    if bg_path and os.path.exists(bg_path):
        crop = cover_crop(bg_path, layout.slide_width_in,
                          layout.slide_height_in)
        slide.shapes.add_picture(crop, Inches(0), Inches(0),
                                 width=Inches(layout.slide_width_in),
                                 height=Inches(layout.slide_height_in))
    tb = add_textbox(slide, layout.margin_left_in, 0.45, 10.0, 0.35)
    write_paragraph(tb.text_frame, title_eyebrow.upper(),
                    size=typo.body_size, bold=True, color=pal.light_blue,
                    family=typo.family, first=True)

    # Left: giant stat
    tb = add_textbox(slide, layout.margin_left_in - 0.05, 1.15, 6.4, 3.4,
                     anchor=MSO_ANCHOR.MIDDLE)
    p = tb.text_frame.paragraphs[0]
    r = p.add_run(); r.text = stat.rstrip("%")
    r.font.size = Pt(230); r.font.bold = True
    r.font.color.rgb = pal.white; r.font.name = typo.family
    r2 = p.add_run(); r2.text = "%"
    r2.font.size = Pt(110); r2.font.bold = True
    r2.font.color.rgb = pal.bright_blue; r2.font.name = typo.family
    tb = add_textbox(slide, layout.margin_left_in, 4.55, 5.9, 0.95)
    write_paragraph(tb.text_frame, stat_label, size=typo.body_size + 5,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    first=True)
    enable_text_shrink(tb.text_frame)

    # Right: 10x10 waffle
    cell = 0.235
    gap = 0.075
    gx = 7.35
    gy = 1.35
    for idx in range(100):
        row_i, col_i = divmod(idx, 10)
        x = gx + col_i * (cell + gap)
        y = gy + row_i * (cell + gap)
        if idx < waffle_filled:
            add_rect(slide, x, y, cell, cell, fill=pal.bright_blue)
        else:
            add_rect(slide, x, y, cell, cell, fill=None, line=pal.white,
                     line_width=0.75)
    tb = add_textbox(slide, gx - 0.15, gy + 10 * (cell + gap) + 0.12,
                     10 * (cell + gap) + 0.3, 0.65)
    write_paragraph(tb.text_frame, waffle_caption, size=typo.small_size + 1,
                    color=pal.light_blue, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    enable_text_shrink(tb.text_frame)

    tb = add_textbox(slide, layout.margin_left_in, 6.35, 9.5, 0.40)
    write_paragraph(tb.text_frame, closing, size=typo.body_size + 3,
                    bold=True, color=pal.white, family=typo.family,
                    first=True)
    if source:
        tb = add_textbox(slide, layout.margin_left_in, 6.90, 10.5, 0.25)
        write_paragraph(tb.text_frame, f"Source: {source}",
                        size=typo.footer_size, color=rgb("7C93B5"),
                        family=typo.family, first=True)
    return slide


# ---------- 14 · hairline stat band with glyphs ----------

def add_stat_band(prs, *,
                  title: str,
                  stats: Sequence[dict],
                  kicker: Optional[str] = None,
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    """stats: [{glyph: 'donut20'|'bar80'|'onethree'|'steps', value, label, context}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    n = len(stats)
    total_w = layout.slide_width_in - layout.margin_left_in - layout.margin_right_in
    mod_w = total_w / n
    top = 2.15
    for i, st in enumerate(stats):
        x = layout.margin_left_in + i * mod_w
        cxm = x + mod_w / 2
        # glyph
        g = st.get("glyph")
        gy = top + 0.1
        if g == "donut20":
            add_oval(slide, cxm - 0.35, gy, 0.7, 0.7, fill=None,
                     line=pal.grid_gray, line_width=6.0)
            pie = slide.shapes.add_shape(MSO_SHAPE.PIE, Inches(cxm - 0.35),
                                         Inches(gy), Inches(0.7), Inches(0.7))
            pie.adjustments[0] = -90.0
            pie.adjustments[1] = -18.0
            pie.shadow.inherit = False
            pie.fill.solid(); pie.fill.fore_color.rgb = pal.dark_navy
            pie.line.fill.background()
        elif g == "bar80":
            add_rect(slide, cxm - 0.75, gy + 0.22, 1.5, 0.26, fill=None,
                     line=pal.grid_gray, line_width=1.0)
            add_rect(slide, cxm - 0.75, gy + 0.22, 1.2, 0.26,
                     fill=pal.dark_navy)
        elif g == "onethree":
            for k in range(3):
                ox = cxm - 0.55 + k * 0.4
                if k == 0:
                    add_oval(slide, ox, gy + 0.18, 0.32, 0.32,
                             fill=pal.dark_navy)
                else:
                    add_oval(slide, ox, gy + 0.18, 0.32, 0.32, fill=None,
                             line=pal.footer_gray, line_width=1.25)
        elif g == "steps":
            for k in range(4):
                h = 0.14 + k * 0.14
                add_rect(slide, cxm - 0.60 + k * 0.32, gy + 0.62 - h,
                         0.24, h, fill=pal.dark_navy)
        # value / label / context
        tb = add_textbox(slide, x + 0.15, top + 1.05, mod_w - 0.3, 0.75)
        write_paragraph(tb.text_frame, st["value"], size=44, bold=True,
                        color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, x + 0.2, top + 1.9, mod_w - 0.4, 0.62)
        write_paragraph(tb.text_frame, st["label"], size=typo.body_size,
                        color=pal.text_dark, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, x + 0.2, top + 2.55, mod_w - 0.4, 0.55)
        write_paragraph(tb.text_frame, st["context"], size=typo.small_size,
                        italic=True, color=pal.footer_gray,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        if i < n - 1:
            add_line(slide, x + mod_w, top + 0.15, x + mod_w, top + 3.0,
                     color=pal.grid_gray, width_pt=0.75)
    if kicker:
        tb = add_textbox(slide, layout.margin_left_in, top + 3.55, total_w,
                         0.4)
        write_paragraph(tb.text_frame, kicker, size=typo.body_size + 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
    return slide


# ---------- 15 · quote breather ----------

def add_quote_breather(prs, *,
                       title: str,
                       quote: str,
                       author: str,
                       author_title: str,
                       photo_label: Optional[str] = None,
                       photo_path: Optional[str] = None,
                       page_number=None, section_marker=None,
                       source=None, footnote=None,
                       theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    tb = add_textbox(slide, 1.1, 2.05, 1.3, 1.3)
    write_paragraph(tb.text_frame, "“", size=90, bold=True,
                    color=pal.bright_blue, family=typo.family, first=True)
    tb = add_textbox(slide, 2.4, 2.30, 9.4, 1.9)
    write_paragraph(tb.text_frame, quote, size=26, color=pal.text_dark,
                    family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)

    ay = 4.75
    if photo_path and os.path.exists(photo_path):
        crop = cover_crop(photo_path, 1.1, 1.1)
        slide.shapes.add_picture(crop, Inches(2.4), Inches(ay - 0.15),
                                 width=Inches(1.1), height=Inches(1.1))
        ax = 3.75
    elif photo_label:
        placeholder_box(slide, 2.4, ay - 0.15, 1.1, 1.1, photo_label,
                        theme=theme)
        ax = 3.75
    else:
        ax = 2.4
    add_line(slide, ax, ay, ax + 0.6, ay, color=pal.bright_blue, width_pt=2.0)
    tb = add_textbox(slide, ax, ay + 0.10, 6.0, 0.35)
    write_paragraph(tb.text_frame, author, size=typo.body_size + 2, bold=True,
                    color=pal.text_dark, family=typo.family, first=True)
    tb = add_textbox(slide, ax, ay + 0.48, 6.0, 0.32)
    write_paragraph(tb.text_frame, author_title, size=typo.body_size,
                    color=pal.footer_gray, family=typo.family, first=True)
    return slide


# ---------- 20 · Green 5 chevron signal bar ----------

def add_chevron_flags(prs, *,
                      title: str,
                      subtitle: Optional[str] = None,
                      flags: Sequence[dict],
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """Five chevrons in a navy→cyan ramp, one-line descriptor under each.
    flags: [{keyword, descriptor}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.32, 11.0, 0.30)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)

    n = len(flags)
    ramp = [rgb("0B2E5C"), rgb("124173"), rgb("19548A"), rgb("2067A1"),
            rgb("18B9FF")]
    overlap = 0.28
    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)
    chev_w = (total_w + (n - 1) * overlap) / n
    chev_h = 1.15
    cy = 3.0
    for i, f in enumerate(flags):
        x = layout.margin_left_in + i * (chev_w - overlap)
        s = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(cy),
                                   Inches(chev_w), Inches(chev_h))
        s.shadow.inherit = False
        s.fill.solid(); s.fill.fore_color.rgb = ramp[i % len(ramp)]
        s.line.fill.background()
        tb = add_textbox(slide, x + 0.42, cy + 0.14, chev_w - 0.75, 0.36)
        write_paragraph(tb.text_frame, str(i + 1), size=typo.body_size + 4,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        tb = add_textbox(slide, x + 0.42, cy + 0.50, chev_w - 0.75, 0.55,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, f["keyword"].upper(),
                        size=typo.body_size, bold=True, color=pal.white,
                        family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        # tick + descriptor
        tick_x = x + chev_w / 2 - overlap / 2
        add_line(slide, tick_x, cy + chev_h + 0.05, tick_x,
                 cy + chev_h + 0.35, color=pal.grid_gray, width_pt=0.75)
        tb = add_textbox(slide, x + 0.05, cy + chev_h + 0.42,
                         chev_w - overlap - 0.10, 0.85)
        write_paragraph(tb.text_frame, f["descriptor"],
                        size=typo.body_size - 1, color=pal.text_dark,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- 21 · Red 3 keyline panels ----------

def add_keyline_panels(prs, *,
                       title: str,
                       subtitle: Optional[str] = None,
                       panels: Sequence[dict],
                       closing: Optional[str] = None,
                       page_number=None, section_marker=None,
                       source=None, footnote=None,
                       theme: Theme = MAX_THEME):
    """Flat panels with a red top keyline and ghost numerals; the deck's only
    red accent family. panels: [{label, bullets(max 2)}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    RED = rgb("B03A2E")
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.32, 11.0, 0.30)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)

    n = len(panels)
    gap = 0.45
    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)
    p_w = (total_w - gap * (n - 1)) / n
    top = 1.85
    p_h = 3.6
    for i, p in enumerate(panels):
        x = layout.margin_left_in + i * (p_w + gap)
        add_rect(slide, x, top, p_w, p_h, fill=PALE_ROW)
        add_rect(slide, x, top, p_w, 0.055, fill=RED)
        tb = add_textbox(slide, x + p_w - 1.55, top + 0.12, 1.35, 1.0)
        write_paragraph(tb.text_frame, f"0{i + 1}", size=60, bold=True,
                        color=rgb("D8DEE6"), family=typo.family,
                        align=PP_ALIGN.RIGHT, first=True)
        tb = add_textbox(slide, x + 0.3, top + 0.55, p_w - 0.6, 0.45)
        write_paragraph(tb.text_frame, p["label"], size=typo.body_size + 4,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, x + 0.3, top + 1.25, p_w - 0.6, p_h - 1.5)
        for j, bl in enumerate(p["bullets"][:2]):
            write_paragraph(tb.text_frame, bl, size=typo.body_size,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=8)
        enable_text_shrink(tb.text_frame)
    if closing:
        cy2 = top + p_h + 0.35
        add_rect(slide, layout.margin_left_in, cy2, total_w, 0.62,
                 fill=pal.deep_navy)
        tb = add_textbox(slide, layout.margin_left_in, cy2, total_w, 0.62,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, closing, size=typo.body_size + 2,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
    return slide


# ---------- 22 · the two ledgers ----------

def add_ledgers(prs, *,
                title: str,
                page_number=None, section_marker=None,
                source=None, footnote=None,
                theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    def book(x, y, w, h, spine_color, page_color, label, label_color,
             lines, text_color):
        add_rect(slide, x, y, 0.28, h, fill=spine_color)
        add_rect(slide, x + 0.28, y, w - 0.28, h, fill=page_color)
        tb = add_textbox(slide, x + 0.55, y + 0.30, w - 0.85, 0.75)
        write_paragraph(tb.text_frame, label, size=typo.body_size + 2,
                        bold=True, color=label_color, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        add_line(slide, x + 0.55, y + 1.10, x + w - 0.35, y + 1.10,
                 color=label_color, width_pt=1.0)
        tb = add_textbox(slide, x + 0.55, y + 1.30, w - 0.85, h - 1.55)
        for j, l in enumerate(lines):
            write_paragraph(tb.text_frame, l, size=typo.body_size,
                            color=text_color, family=typo.family,
                            bullet=True, first=(j == 0), space_after=10)
        enable_text_shrink(tb.text_frame)

    by, bh, bw = 2.35, 3.4, 4.4
    book(1.0, by, bw, bh, rgb("8A97A5"), rgb("EEF1F5"),
         "THE BOOKS THEY SHOW", rgb("6B7B8C"),
         ["Prepared for the tax office",
          "Understated revenue and profit",
          "Looks cheap, and is unbankable"],
         pal.footer_gray)
    book(7.65, by, bw, bh, pal.bright_blue, pal.deep_navy,
         "THE BOOKS THAT MATTER", pal.bright_blue,
         ["Actual cash flows and contracts",
          "Real margins, verified in diligence",
          "Supports financing and a fair price"],
         pal.white)

    # Diligence lens bridging both
    lens_cx, lens_cy, lens_d = 6.52, 2.6, 1.5
    add_oval(slide, lens_cx - lens_d / 2, lens_cy - lens_d / 2, lens_d,
             lens_d, fill=None, line=pal.bright_blue, line_width=3.0)
    add_line(slide, lens_cx + lens_d * 0.33, lens_cy + lens_d * 0.33,
             lens_cx + lens_d * 0.62, lens_cy + lens_d * 0.62,
             color=pal.bright_blue, width_pt=4.0)
    tb = add_textbox(slide, lens_cx - 1.02, lens_cy + lens_d * 0.72, 2.04, 0.62)
    write_paragraph(tb.text_frame, "Diligence reconciles the two",
                    size=typo.body_size, bold=True, color=pal.bright_blue,
                    family=typo.family, align=PP_ALIGN.CENTER, first=True)

    tb = add_textbox(slide, layout.margin_left_in, 6.15, 12.0, 0.40)
    write_paragraph(tb.text_frame,
                    "You are buying the second book. Diligence is how you make sure it exists.",
                    size=typo.body_size + 2, bold=True, color=pal.dark_navy,
                    family=typo.family, first=True)
    return slide


# ---------- 25 · tapering funnel ----------

def add_taper_funnel(prs, *,
                     title: str,
                     stages: Sequence[dict],
                     chip: Optional[str] = None,
                     page_number=None, section_marker=None,
                     source=None, footnote=None,
                     theme: Theme = MAX_THEME):
    """Centered tapering funnel with a dramatic final stage.
    stages: [{name, value, description}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    widths = [7.6, 6.1, 4.9, 3.3, 2.1, 1.0]
    ramp = [rgb("011C40"), rgb("022859"), rgb("0A3A73"), rgb("14508C"),
            rgb("1D74B8"), rgb("18B9FF")]
    n = len(stages)
    top = 1.70
    row_h = 0.72
    gap = 0.14
    cx = 4.55
    for i, st in enumerate(stages):
        w = widths[i] if i < len(widths) else 1.0
        x = cx - w / 2
        y = top + i * (row_h + gap)
        add_rect(slide, x, y, w, row_h, fill=ramp[i % len(ramp)])
        if i == n - 1:
            add_rect(slide, x - 0.07, y - 0.07, w + 0.14, row_h + 0.14,
                     fill=None, line=pal.bright_blue, line_width=2.0)
        inner_w = max(w - 0.3, 1.6)
        tb = add_textbox(slide, cx - inner_w / 2, y + 0.05, inner_w, 0.30,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, st["name"], size=typo.small_size + 1,
                        bold=True,
                        color=rgb("022859") if i == n - 1 else pal.white,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, cx - inner_w / 2, y + 0.34, inner_w, 0.34,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, st["value"], size=typo.body_size + 3,
                        bold=True,
                        color=rgb("022859") if i == n - 1 else pal.light_blue,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        # leader line + description on the right
        dx = 9.05
        add_line(slide, cx + w / 2 + 0.08, y + row_h / 2, dx - 0.12,
                 y + row_h / 2, color=pal.grid_gray, width_pt=0.6)
        tb = add_textbox(slide, dx, y + 0.06, 4.0, row_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(
            tb.text_frame,
            st["description"] + ("  ·  This is the needle." if i == n - 1 else ""),
            size=typo.small_size + 1,
            bold=(i == n - 1),
            color=pal.dark_navy if i == n - 1 else pal.text_dark,
            family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
    if chip:
        cy2 = top + (n - 1) * (row_h + gap) - 0.72
        add_rect(slide, 0.55, cy2, 2.6, 0.85, fill=pal.dark_navy)
        tb = add_textbox(slide, 0.68, cy2 + 0.08, 2.34, 0.69,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, chip, size=typo.small_size + 1,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- 26 · channel ceilings ----------

def add_ceilings(prs, *,
                 title: str,
                 takeaway: str,
                 channels: Sequence[dict],
                 breakout: str,
                 page_number=None, section_marker=None,
                 source=None, footnote=None,
                 theme: Theme = MAX_THEME):
    """Rising columns each capped by a gray 'ceiling' line.
    channels: [{name, ceiling, note}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    takeaway_line(slide, takeaway, theme)

    base_y = 5.55
    heights = [1.2, 1.9, 2.6]
    col_w = 2.1
    xs = [1.6, 5.2, 8.8]
    for i, ch in enumerate(channels):
        x, h = xs[i], heights[i]
        add_rect(slide, x, base_y - h, col_w, h, fill=pal.dark_navy)
        # ceiling cap
        add_line(slide, x - 0.35, base_y - h - 0.12, x + col_w + 0.35,
                 base_y - h - 0.12, color=rgb("8A97A5"), width_pt=3.0)
        tb = add_textbox(slide, x - 0.35, base_y - h - 0.48,
                         col_w + 0.7, 0.30)
        write_paragraph(tb.text_frame, ch["ceiling"], size=typo.small_size,
                        italic=True, color=pal.footer_gray,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, x - 0.35, base_y + 0.12, col_w + 0.7, 0.34)
        write_paragraph(tb.text_frame, ch["name"], size=typo.body_size + 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, x - 0.45, base_y + 0.50, col_w + 0.9, 0.68)
        write_paragraph(tb.text_frame, ch["note"], size=typo.small_size + 1,
                        color=pal.text_dark, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
    # breakout arrow above the tallest ceiling, label sitting to its left
    bx = xs[2] + col_w / 2
    top_h = heights[2]
    add_line(slide, bx, base_y - top_h - 0.50, bx,
             base_y - top_h - 1.00, color=pal.bright_blue, width_pt=2.5)
    tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                 Inches(bx - 0.14),
                                 Inches(base_y - top_h - 1.20),
                                 Inches(0.28), Inches(0.22))
    tri.shadow.inherit = False
    tri.fill.solid(); tri.fill.fore_color.rgb = pal.bright_blue
    tri.line.fill.background()
    tb = add_textbox(slide, bx - 7.0, base_y - top_h - 0.96, 6.60, 0.32)
    write_paragraph(tb.text_frame, breakout, size=typo.body_size, bold=True,
                    color=pal.bright_blue, family=typo.family,
                    align=PP_ALIGN.RIGHT, first=True)
    enable_text_shrink(tb.text_frame)
    return slide


# ---------- 30 · metro line ----------

def add_metro_line(prs, *,
                   title: str,
                   stops: Sequence[dict],
                   input_note: str,
                   output_note: str,
                   page_number=None, section_marker=None,
                   source=None, footnote=None,
                   theme: Theme = MAX_THEME):
    """Horizontal metro pipeline; final stop cyan = the payoff.
    stops: [{name, description}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    n = len(stops)
    y = 3.9
    left0, right0 = 1.7, 11.6
    add_line(slide, left0 - 0.5, y, right0 + 0.5, y, color=pal.dark_navy,
             width_pt=3.0)
    step = (right0 - left0) / (n - 1)
    d = 0.55
    for i, st in enumerate(stops):
        cx = left0 + i * step
        last = (i == n - 1)
        if st.get("icon") and os.path.exists(st["icon"]):
            slide_icon(slide, st["icon"], cx - 0.30, y - 1.78, 0.60)
        add_oval(slide, cx - d / 2, y - d / 2, d, d,
                 fill=pal.bright_blue if last else pal.dark_navy)
        tb = add_textbox(slide, cx - d / 2, y - d / 2, d, d,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, f"{i + 1:02d}", size=typo.body_size,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, cx - 1.05, y + 0.45, 2.1, 0.34)
        write_paragraph(tb.text_frame, st["name"], size=typo.body_size + 1,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, cx - 1.05, y + 0.80, 2.1, 0.85)
        write_paragraph(tb.text_frame, st["description"],
                        size=typo.small_size, color=pal.footer_gray,
                        family=typo.family, align=PP_ALIGN.CENTER,
                        first=True)
        enable_text_shrink(tb.text_frame)
    # input / output annotations
    tb = add_textbox(slide, left0 - 1.35, y - 1.05, 2.6, 0.55)
    write_paragraph(tb.text_frame, input_note.upper(), size=typo.small_size,
                    bold=True, color=pal.footer_gray, family=typo.family,
                    first=True)
    tb = add_textbox(slide, right0 - 1.15, y - 1.05, 2.6, 0.55)
    write_paragraph(tb.text_frame, output_note.upper(), size=typo.small_size,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    align=PP_ALIGN.RIGHT, first=True)
    return slide


# ---------- 33 · rising road ----------

def add_rising_road(prs, *,
                    title: str,
                    subtitle: Optional[str] = None,
                    steps: Sequence[dict],
                    page_number=None, section_marker=None,
                    source=None, footnote=None,
                    theme: Theme = MAX_THEME):
    """Stepped path climbing left→right; final node cyan with a key glyph.
    steps: [{name, description}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.32, 11.5, 0.30)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.text_dark, family=typo.family, first=True)

    n = len(steps)
    x0, y0 = 0.9, 5.85
    x1, y1 = 11.9, 2.45
    pts = []
    for i in range(n):
        t = i / (n - 1)
        pts.append((x0 + t * (x1 - x0), y0 + t * (y1 - y0)))
    for i in range(n - 1):
        add_line(slide, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1],
                 color=pal.dark_navy, width_pt=6.0)
    d = 0.52
    for i, (st, (cx, cy)) in enumerate(zip(steps, pts)):
        last = (i == n - 1)
        add_oval(slide, cx - d / 2, cy - d / 2, d, d,
                 fill=pal.bright_blue if last else pal.dark_navy)
        tb = add_textbox(slide, cx - d / 2, cy - d / 2, d, d,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, str(i + 1), size=typo.body_size + 1,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        above = (i % 2 == 1)
        bx = min(max(cx - 1.15, 0.3), layout.slide_width_in - 2.6)
        if above:
            ny_ = cy - 1.42
        else:
            ny_ = cy + 0.42
        tb = add_textbox(slide, bx, ny_, 2.3, 0.32)
        write_paragraph(tb.text_frame, st["name"], size=typo.body_size + 1,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, bx, ny_ + 0.32, 2.3, 0.62)
        write_paragraph(tb.text_frame, st["description"],
                        size=typo.small_size, color=pal.footer_gray,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
    # Key glyph beside the final node
    kx, ky = pts[-1][0] + 0.45, pts[-1][1] - 0.05
    add_oval(slide, kx, ky - 0.12, 0.26, 0.26, fill=None,
             line=pal.bright_blue, line_width=2.5)
    add_line(slide, kx + 0.26, ky, kx + 0.62, ky, color=pal.bright_blue,
             width_pt=2.5)
    add_line(slide, kx + 0.48, ky, kx + 0.48, ky + 0.10,
             color=pal.bright_blue, width_pt=2.5)
    add_line(slide, kx + 0.58, ky, kx + 0.58, ky + 0.13,
             color=pal.bright_blue, width_pt=2.5)
    return slide


# ---------- 34 · buyer/seller mirror spine ----------

def add_mirror_spine(prs, *,
                     title: str,
                     stages: Sequence[dict],
                     left_header: str = "BUYER'S PATH",
                     right_header: str = "SELLER'S MIRROR",
                     page_number=None, section_marker=None,
                     source=None, footnote=None,
                     theme: Theme = MAX_THEME):
    """Central spine with stage chips; inward-pointing plates each side.
    stages: [{stage, left, right}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    spine_x = layout.slide_width_in / 2
    top, bottom = 2.15, 6.35
    add_line(slide, spine_x, top - 0.15, spine_x, bottom,
             color=pal.dark_navy, width_pt=1.5)
    tb = add_textbox(slide, spine_x - 5.6, 1.55, 3.6, 0.3)
    write_paragraph(tb.text_frame, left_header, size=typo.small_size + 1,
                    bold=True, color=pal.footer_gray, family=typo.family,
                    first=True)
    tb = add_textbox(slide, spine_x + 2.0, 1.55, 3.6, 0.3)
    write_paragraph(tb.text_frame, right_header, size=typo.small_size + 1,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    align=PP_ALIGN.RIGHT, first=True)

    n = len(stages)
    row_h = (bottom - top) / n
    plate_w, plate_h = 3.7, 0.72
    for i, st in enumerate(stages):
        cy = top + i * row_h + row_h / 2
        # stage chip on the spine
        chip_w, chip_h = 1.75, 0.44
        add_rect(slide, spine_x - chip_w / 2, cy - chip_h / 2, chip_w,
                 chip_h, fill=pal.dark_navy)
        tb = add_textbox(slide, spine_x - chip_w / 2, cy - chip_h / 2,
                         chip_w, chip_h, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, st["stage"].upper(),
                        size=typo.small_size, bold=True, color=pal.white,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        # left plate (points right)
        lp = slide.shapes.add_shape(MSO_SHAPE.PENTAGON,
                                    Inches(spine_x - chip_w / 2 - 0.25
                                           - plate_w),
                                    Inches(cy - plate_h / 2),
                                    Inches(plate_w), Inches(plate_h))
        lp.shadow.inherit = False
        lp.fill.solid(); lp.fill.fore_color.rgb = rgb("E8EDF3")
        lp.line.fill.background()
        tb = add_textbox(slide, spine_x - chip_w / 2 - 0.25 - plate_w + 0.15,
                         cy - plate_h / 2, plate_w - 0.5, plate_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, st["left"], size=typo.body_size,
                        color=pal.dark_navy, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        # right plate (points left)
        rp = slide.shapes.add_shape(MSO_SHAPE.PENTAGON,
                                    Inches(spine_x + chip_w / 2 + 0.25),
                                    Inches(cy - plate_h / 2),
                                    Inches(plate_w), Inches(plate_h))
        rp.shadow.inherit = False
        rp.rotation = 180
        rp.fill.solid(); rp.fill.fore_color.rgb = pal.deep_navy
        rp.line.fill.background()
        tb = add_textbox(slide, spine_x + chip_w / 2 + 0.60,
                         cy - plate_h / 2, plate_w - 0.5, plate_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, st["right"], size=typo.body_size,
                        color=pal.white, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- 41 · causality band ----------

def add_causality_band(prs, *,
                       title: str,
                       subtitle: Optional[str] = None,
                       blocks: Sequence[dict],
                       result: str,
                       page_number=None, section_marker=None,
                       source=None, footnote=None,
                       theme: Theme = MAX_THEME):
    """blocks: [{label, support}] joined by chevrons, then '=' then result chip."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.32, 11.0, 0.3)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)

    bw, bh = 2.75, 1.75
    gap = 0.55
    cy = 3.9
    x = layout.margin_left_in + 0.1
    for i, blk in enumerate(blocks):
        add_rect(slide, x, cy - bh / 2, bw, bh, fill=pal.deep_navy)
        tb = add_textbox(slide, x + 0.18, cy - bh / 2 + 0.12, 0.5, 0.3)
        write_paragraph(tb.text_frame, str(i + 1), size=typo.body_size,
                        bold=True, color=pal.light_blue, family=typo.family,
                        first=True)
        tb = add_textbox(slide, x + 0.18, cy - bh / 2 + 0.45, bw - 0.36,
                         0.62)
        write_paragraph(tb.text_frame, blk["label"], size=15, bold=True,
                        color=pal.white, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, x + 0.18, cy - bh / 2 + 1.05, bw - 0.36,
                         0.60)
        write_paragraph(tb.text_frame, blk["support"],
                        size=typo.small_size + 0.5, color=pal.light_blue,
                        family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        x += bw
        if i < len(blocks) - 1:
            ch = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x + 0.06),
                                        Inches(cy - 0.26), Inches(0.42),
                                        Inches(0.52))
            ch.shadow.inherit = False
            ch.fill.solid(); ch.fill.fore_color.rgb = pal.bright_blue
            ch.line.fill.background()
            x += gap
    tb = add_textbox(slide, x + 0.02, cy - 0.35, 0.5, 0.7,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "=", size=28, bold=True,
                    color=pal.footer_gray, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    rx = x + 0.55
    rw = layout.slide_width_in - layout.margin_right_in - rx
    add_rect(slide, rx, cy - bh / 2 + 0.2, rw, bh - 0.4,
             fill=pal.bright_blue)
    tb = add_textbox(slide, rx + 0.2, cy - bh / 2 + 0.2, rw - 0.4, bh - 0.4,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, result, size=15, bold=True,
                    color=rgb("022859"), family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)
    return slide


# ---------- 38 · split stat hero ----------

def add_stat_hero_split(prs, *,
                        title: str,
                        stat: str,
                        stat_label: str,
                        rows: Sequence[dict],
                        kicker: Optional[str] = None,
                        page_number=None, section_marker=None,
                        source=None, footnote=None,
                        theme: Theme = MAX_THEME):
    """Giant number left; hairline-separated evidence rows right.
    rows: [{value, label}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    tb = add_textbox(slide, layout.margin_left_in, 2.0, 6.2, 2.4,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, stat, size=110, bold=True,
                    color=pal.dark_navy, family=typo.family, first=True)
    tb = add_textbox(slide, layout.margin_left_in, 4.45, 5.9, 1.0)
    write_paragraph(tb.text_frame, stat_label, size=18, bold=True,
                    color=pal.bright_blue, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)

    rx = 7.3
    rw = layout.slide_width_in - layout.margin_right_in - rx
    row_h = 1.02
    y = 2.05
    for i, r in enumerate(rows):
        tb = add_textbox(slide, rx, y, 1.7, row_h, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["value"], size=24, bold=True,
                        color=pal.dark_navy, family=typo.family, first=True)
        tb = add_textbox(slide, rx + 1.8, y, rw - 1.8, row_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["label"], size=typo.body_size,
                        color=pal.text_dark, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        if i < len(rows) - 1:
            add_line(slide, rx, y + row_h, rx + rw, y + row_h,
                     color=pal.grid_gray, width_pt=0.75)
        y += row_h + 0.06
    if kicker:
        tb = add_textbox(slide, layout.margin_left_in, 6.15, 12.0, 0.4)
        write_paragraph(tb.text_frame, kicker, size=typo.body_size + 1,
                        italic=True, color=pal.footer_gray,
                        family=typo.family, first=True)
    return slide


# ---------- 46 · Q&A ----------

def add_qa_slide(prs, *,
                 line: str,
                 qr_caption: str,
                 bg_path: Optional[str] = None,
                 page_number=None, section_marker=None,
                 source=None, footnote=None,
                 theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    add_rect(slide, 0, 0, layout.slide_width_in, layout.slide_height_in,
             fill=rgb("022859"))
    if bg_path and os.path.exists(bg_path):
        crop = cover_crop(bg_path, layout.slide_width_in,
                          layout.slide_height_in)
        slide.shapes.add_picture(crop, Inches(0), Inches(0),
                                 width=Inches(layout.slide_width_in),
                                 height=Inches(layout.slide_height_in))
    tb = add_textbox(slide, 0.7, 0.85, 8.0, 2.3)
    p = tb.text_frame.paragraphs[0]
    r = p.add_run(); r.text = "Q&"
    r.font.size = Pt(120); r.font.bold = True
    r.font.color.rgb = pal.white; r.font.name = typo.family
    r2 = p.add_run(); r2.text = "A"
    r2.font.size = Pt(120); r2.font.bold = True
    r2.font.color.rgb = pal.bright_blue; r2.font.name = typo.family
    tb = add_textbox(slide, 0.75, 3.35, 7.4, 0.6)
    write_paragraph(tb.text_frame, line, size=18, color=pal.white,
                    family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)
    placeholder_box(slide, 9.5, 2.1, 2.6, 2.6, "QR: booking page",
                    theme=theme)
    tb = add_textbox(slide, 9.0, 4.85, 3.6, 0.55)
    write_paragraph(tb.text_frame, qr_caption, size=typo.small_size + 1,
                    color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    closing_strip(slide, theme)
    return slide


# ---------- 47 · credits ----------

def add_credits_slide(prs, *,
                      title: str = "With gratitude",
                      groups: Sequence[dict],
                      finale: str,
                      finale_name: str,
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """Editorial credits page. groups: [{to, line}] in rising emotional order;
    finale gets its own set-apart moment."""
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    tb = add_textbox(slide, 0, 0.75, layout.slide_width_in, 0.35)
    write_paragraph(tb.text_frame, "WITH GRATITUDE", size=typo.body_size,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    add_line(slide, layout.slide_width_in / 2 - 0.6, 1.25,
             layout.slide_width_in / 2 + 0.6, 1.25,
             color=pal.bright_blue, width_pt=2.0)

    # vertically center the composition: short lists breathe instead of
    # hugging the header
    spacing = 1.02 if len(groups) >= 3 else 1.25
    block_h = len(groups) * spacing + 1.95
    y = max(1.75, 1.6 + (5.55 - block_h) / 2)
    for g in groups:
        tb = add_textbox(slide, 2.2, y, 8.93, 0.34)
        write_paragraph(tb.text_frame, g["to"].upper(),
                        size=typo.small_size + 1, bold=True,
                        color=pal.footer_gray, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, 1.7, y + 0.34, 9.93, 0.52)
        write_paragraph(tb.text_frame, g["line"], size=typo.body_size + 2,
                        color=pal.text_dark, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        y += spacing

    # Finale: partner moment
    add_line(slide, 4.4, y + 0.12, 8.93, y + 0.12, color=pal.grid_gray,
             width_pt=0.75)
    tb = add_textbox(slide, 1.35, y + 0.35, 10.63, 0.85)
    write_paragraph(tb.text_frame, finale, size=20, italic=True,
                    color=pal.dark_navy, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    enable_text_shrink(tb.text_frame)
    tb = add_textbox(slide, 0, y + 1.28, layout.slide_width_in, 0.5)
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = f"{finale_name}  "
    r.font.size = Pt(22); r.font.bold = True
    r.font.color.rgb = pal.dark_navy; r.font.name = typo.family
    r2 = p.add_run(); r2.text = "♥"
    r2.font.size = Pt(20); r2.font.bold = True
    r2.font.color.rgb = pal.bright_blue; r2.font.name = typo.family
    return slide


# ---------- 50 · vertical process ladder ----------

def add_vertical_ladder(prs, *,
                        title: str,
                        steps: Sequence[dict],
                        page_number=None, section_marker=None,
                        source=None, footnote=None,
                        theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    x = 1.4
    top, bottom = 1.8, 6.4
    add_line(slide, x, top, x, bottom, color=pal.dark_navy, width_pt=2.0)
    n = len(steps)
    d = 0.45
    for i, st in enumerate(steps):
        cy = top + i * (bottom - top) / (n - 1) if n > 1 else top
        last = (i == n - 1)
        add_oval(slide, x - d / 2, cy - d / 2, d, d,
                 fill=pal.bright_blue if last else pal.dark_navy)
        tb = add_textbox(slide, x - d / 2, cy - d / 2, d, d,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, f"{i + 1:02d}", size=typo.small_size,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, x + 0.55, cy - 0.30, 10.6, 0.6,
                         anchor=MSO_ANCHOR.MIDDLE)
        p = tb.text_frame.paragraphs[0]
        r = p.add_run(); r.text = st["name"] + "  ·  "
        r.font.size = Pt(typo.body_size + 1); r.font.bold = True
        r.font.color.rgb = pal.dark_navy; r.font.name = typo.family
        r2 = p.add_run(); r2.text = st["description"]
        r2.font.size = Pt(typo.body_size)
        r2.font.color.rgb = pal.footer_gray; r2.font.name = typo.family
    return slide


# ---------- 51 · definition list ----------

def add_definition_list(prs, *,
                        title: str,
                        subtitle: Optional[str] = None,
                        terms: Sequence[dict],
                        page_number=None, section_marker=None,
                        source=None, footnote=None,
                        theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, 1.32, 10.0, 0.28)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)
    left = layout.margin_left_in
    total_w = layout.slide_width_in - left - layout.margin_right_in
    top = 1.95
    row_h = 0.92
    add_line(slide, left, top, left + total_w, top, color=pal.dark_navy,
             width_pt=1.0)
    for i, t in enumerate(terms):
        y = top + 0.08 + i * row_h
        tb = add_textbox(slide, left, y, 0.7, row_h - 0.16,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, f"{i + 1:02d}", size=typo.body_size,
                        color=pal.footer_gray, family=typo.family, first=True)
        tb = add_textbox(slide, left + 0.8, y, 2.3, row_h - 0.16,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, t["term"], size=typo.body_size + 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 3.2, y, total_w - 3.2, row_h - 0.16,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, t["definition"], size=typo.body_size,
                        color=pal.text_dark, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        add_line(slide, left, y + row_h - 0.04, left + total_w,
                 y + row_h - 0.04, color=pal.grid_gray, width_pt=0.75)
    return slide


# ---------- 52 · methodology grid ----------

def add_method_grid(prs, *,
                    title: str,
                    blocks: Sequence[dict],
                    page_number=None, section_marker=None,
                    source=None, footnote=None,
                    theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    col_w = 5.9
    gut = 0.55
    row_h = 2.45
    for i, blk in enumerate(blocks):
        ci, ri = i % 2, i // 2
        x = layout.margin_left_in + ci * (col_w + gut)
        y = 1.65 + ri * (row_h + 0.25)
        tb = add_textbox(slide, x, y, col_w, 0.28)
        write_paragraph(tb.text_frame, blk["header"].upper(),
                        size=typo.small_size + 1, bold=True,
                        color=pal.bright_blue, family=typo.family, first=True)
        add_line(slide, x, y + 0.32, x + 1.2, y + 0.32, color=pal.dark_navy,
                 width_pt=1.5)
        tb = add_textbox(slide, x, y + 0.45, col_w, row_h - 0.5)
        write_paragraph(tb.text_frame, blk["body"], size=typo.body_size - 1,
                        color=pal.text_dark, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


_REGISTRY.update({
    "art_divider": add_art_divider,
    "route_map": add_route_map,
    "speaker_panels": add_speaker_panels,
    "hbar_ranked": add_hbar_ranked,
    "sector_matrix": add_sector_matrix,
    "fork_road": add_fork_road,
    "dimension_table": add_dimension_table,
    "stat_hero_navy": add_stat_hero_navy,
    "stat_band": add_stat_band,
    "quote_breather": add_quote_breather,
    "chevron_flags": add_chevron_flags,
    "keyline_panels": add_keyline_panels,
    "ledgers": add_ledgers,
    "taper_funnel": add_taper_funnel,
    "ceilings": add_ceilings,
    "metro_line": add_metro_line,
    "rising_road": add_rising_road,
    "mirror_spine": add_mirror_spine,
    "causality_band": add_causality_band,
    "stat_hero_split": add_stat_hero_split,
    "qa_slide": add_qa_slide,
    "credits_slide": add_credits_slide,
    "vertical_ladder": add_vertical_ladder,
    "definition_list": add_definition_list,
    "method_grid": add_method_grid,
})
