"""Custom Max Solutions slide templates for the M&A webinar deck.

Extends the mckinsey_pptx engine with webinar-specific layouts:
 - max_cover: branded cover with the Max Solutions logo
 - speaker_slide: presenter cards with highlighted photo placeholders
 - poll_slide: live-poll moment with lettered options
 - screenshot_slide: product walkthrough with a highlighted screenshot placeholder
 - case_slide: deal case study (situation -> outcome + KPI band)
 - access_ladder: 3-step ascending "where deals come from" diagram
 - cta_slide: segmented call-to-action with QR placeholder
 - thank_you: navy closing slide with white logo

Every unfinished asset renders as a loud amber PLACEHOLDER box so nothing
ships looking accidentally empty.
"""
from __future__ import annotations
from typing import Optional, Sequence

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.util import Inches, Pt

from mckinsey_pptx.base import (
    blank_slide, add_chrome, add_footer, add_rect, add_oval, add_line,
    add_textbox, write_paragraph, enable_text_shrink,
)
from mckinsey_pptx.theme import Theme, rgb
from mckinsey_pptx.builder import _REGISTRY

from max_theme import MAX_THEME, PH_BORDER, PH_FILL, PH_TEXT

ASSETS = __file__.rsplit("/", 1)[0] + "/assets"
LOGO = f"{ASSETS}/maxsolutions-logo.png"          # navy on transparent
LOGO_WHITE = f"{ASSETS}/maxsolutions-logo-white.png"
MAXDATA_ICON = f"{ASSETS}/maxdata-icon.png"


# ---------- helpers ----------

def placeholder_box(slide, left_in, top_in, width_in, height_in, label,
                    sublabel=None, theme: Theme = MAX_THEME):
    """Loud amber dashed box marking an asset the owner must drop in."""
    typo = theme.typography
    box = add_rect(slide, left_in, top_in, width_in, height_in,
                   fill=rgb(PH_FILL), line=rgb(PH_BORDER), line_width=1.5)
    box.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    tb = add_textbox(slide, left_in + 0.15, top_in, width_in - 0.3,
                     height_in, anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "PLACEHOLDER", size=typo.small_size,
                    bold=True, color=rgb(PH_BORDER), family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    write_paragraph(tb.text_frame, label, size=typo.body_size, bold=True,
                    color=rgb(PH_TEXT), family=typo.family,
                    align=PP_ALIGN.CENTER, space_before=2)
    if sublabel:
        write_paragraph(tb.text_frame, sublabel, size=typo.small_size,
                        color=rgb(PH_TEXT), family=typo.family,
                        align=PP_ALIGN.CENTER, space_before=2)
    enable_text_shrink(tb.text_frame)
    return box


def add_logo(slide, path, left_in, top_in, width_in):
    return slide.shapes.add_picture(path, Inches(left_in), Inches(top_in),
                                    width=Inches(width_in))


# ---------- cover ----------

def add_max_cover(prs, *,
                  title: str,
                  subtitle: Optional[str] = None,
                  event_line: Optional[str] = None,
                  date: Optional[str] = None,
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    # Right panel: generated Bangkok-skyline hero art if present, else navy
    # with the Max Data icon watermark.
    panel_w = 4.2
    hero = f"{ASSETS}/gen/cover-hero-crop.png"
    add_rect(slide, layout.slide_width_in - panel_w, 0, panel_w,
             layout.slide_height_in, fill=pal.deep_navy)
    import os as _os
    if _os.path.exists(hero):
        slide.shapes.add_picture(hero,
                                 Inches(layout.slide_width_in - panel_w),
                                 Inches(0), width=Inches(panel_w),
                                 height=Inches(layout.slide_height_in))
        add_logo(slide, MAXDATA_ICON,
                 layout.slide_width_in - panel_w / 2 - 0.4,
                 layout.slide_height_in - 1.25, 0.8)
    else:
        add_logo(slide, MAXDATA_ICON, layout.slide_width_in - panel_w + 1.05,
                 layout.slide_height_in / 2 - 1.05, 2.1)

    # Max Solutions logo top-left
    add_logo(slide, LOGO, layout.margin_left_in + 0.15, 0.45, 1.9)

    left = layout.margin_left_in + 0.15
    text_w = layout.slide_width_in - panel_w - left - 0.5

    if event_line:
        tb = add_textbox(slide, left, 2.05, text_w, 0.35)
        write_paragraph(tb.text_frame, event_line.upper(),
                        size=typo.body_size, bold=True,
                        color=pal.bright_blue, family=typo.family, first=True)

    tb = add_textbox(slide, left, 2.45, text_w, 1.9)
    write_paragraph(tb.text_frame, title, size=typo.title_size + 14,
                    bold=True, color=pal.dark_navy, family=typo.family,
                    first=True)
    enable_text_shrink(tb.text_frame)

    if subtitle:
        tb = add_textbox(slide, left, 4.45, text_w, 0.95)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size + 4,
                        color=pal.footer_gray, family=typo.family, first=True)

    add_line(slide, left, layout.slide_height_in - 1.15,
             left + 5.8, layout.slide_height_in - 1.15,
             color=pal.bright_blue, width_pt=2.0)
    if date:
        tb = add_textbox(slide, left, layout.slide_height_in - 1.0, 7.0, 0.35)
        write_paragraph(tb.text_frame, date, size=typo.body_size + 1,
                        bold=True, color=pal.text_dark, family=typo.family,
                        first=True)
    return slide


# ---------- speakers ----------

def add_speaker_slide(prs, *,
                      title: str = "Your guides tonight",
                      speakers: Sequence[dict],
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """2-3 presenter cards: photo placeholder, name, role, credential bullets."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    n = len(speakers)
    gap = 0.5
    total_w = layout.slide_width_in - layout.margin_left_in - layout.margin_right_in
    card_w = min(4.6, (total_w - gap * (n - 1)) / n)
    start_left = (layout.slide_width_in - (card_w * n + gap * (n - 1))) / 2
    top = 1.85
    card_h = 4.7

    for i, sp in enumerate(speakers):
        left = start_left + i * (card_w + gap)
        add_rect(slide, left, top, card_w, card_h, fill=pal.soft_gray)
        add_rect(slide, left, top, card_w, 0.14, fill=pal.dark_navy)

        # Photo placeholder
        ph_w = 1.7
        placeholder_box(slide, left + (card_w - ph_w) / 2, top + 0.45,
                        ph_w, 1.7, sp.get("photo_label", "Speaker photo"),
                        theme=theme)

        tb = add_textbox(slide, left + 0.3, top + 2.35, card_w - 0.6, 0.4)
        write_paragraph(tb.text_frame, sp["name"], size=typo.body_size + 4,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, left + 0.3, top + 2.78, card_w - 0.6, 0.4)
        write_paragraph(tb.text_frame, sp["role"], size=typo.body_size,
                        color=pal.bright_blue, bold=True, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)

        tb = add_textbox(slide, left + 0.45, top + 3.25, card_w - 0.9, 1.3)
        for j, b in enumerate(sp.get("bullets", [])):
            write_paragraph(tb.text_frame, b, size=typo.body_size - 1,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=4)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- live poll ----------

def add_poll_slide(prs, *,
                   title: str = "Live poll",
                   question: str,
                   options: Sequence[str],
                   instruction: str = "Answer in the poll panel now",
                   page_number=None, section_marker=None,
                   source=None, footnote=None,
                   theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    # LIVE POLL badge
    badge_w = 1.55
    add_rect(slide, layout.margin_left_in, 1.55, badge_w, 0.38,
             fill=pal.bright_blue)
    tb = add_textbox(slide, layout.margin_left_in, 1.55, badge_w, 0.38,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "LIVE POLL", size=typo.body_size, bold=True,
                    color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    tb = add_textbox(slide, layout.margin_left_in, 2.15,
                     layout.slide_width_in - layout.margin_left_in
                     - layout.margin_right_in, 0.9)
    write_paragraph(tb.text_frame, question, size=typo.title_size + 2,
                    bold=True, color=pal.dark_navy, family=typo.family,
                    first=True)
    enable_text_shrink(tb.text_frame)

    top = 3.25
    row_h = 0.62
    letters = "ABCDEFG"
    for i, opt in enumerate(options):
        y = top + i * (row_h + 0.16)
        add_oval(slide, layout.margin_left_in + 0.15, y + 0.06, 0.5, 0.5,
                 fill=pal.dark_navy)
        tb = add_textbox(slide, layout.margin_left_in + 0.15, y + 0.06,
                         0.5, 0.5, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, letters[i], size=typo.body_size + 2,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        add_rect(slide, layout.margin_left_in + 0.85, y, 8.6, row_h,
                 fill=pal.soft_gray)
        tb = add_textbox(slide, layout.margin_left_in + 1.1, y, 8.2, row_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, opt, size=typo.body_size + 2,
                        color=pal.text_dark, family=typo.family, first=True)

    # Instruction ribbon, right side
    rib_left = layout.slide_width_in - layout.margin_right_in - 2.7
    add_rect(slide, rib_left, 3.25, 2.7, 1.15, fill=pal.deep_navy)
    tb = add_textbox(slide, rib_left + 0.2, 3.25, 2.3, 1.15,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, instruction, size=typo.body_size,
                    bold=True, color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    return slide


# ---------- product screenshot walkthrough ----------

def add_screenshot_slide(prs, *,
                         title: str,
                         placeholder_label: str,
                         placeholder_note: Optional[str] = None,
                         image_path: Optional[str] = None,
                         image_caption: Optional[str] = None,
                         bullets: Sequence[str] = (),
                         stats: Sequence[dict] = (),
                         page_number=None, section_marker=None,
                         source=None, footnote=None,
                         theme: Theme = MAX_THEME):
    """Left: talking points + optional stat tiles. Right: large highlighted
    screenshot placeholder."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    left_w = 4.3
    top = 1.7
    tb = add_textbox(slide, layout.margin_left_in, top, left_w, 3.0)
    for j, b in enumerate(bullets):
        write_paragraph(tb.text_frame, b, size=typo.body_size + 1,
                        color=pal.text_dark, family=typo.family, bullet=True,
                        first=(j == 0), space_after=8)
    enable_text_shrink(tb.text_frame)

    # Stat tiles under the bullets
    if stats:
        tile_top = 4.55
        tile_w = (left_w - 0.2 * (len(stats) - 1)) / len(stats)
        for i, st in enumerate(stats):
            tl = layout.margin_left_in + i * (tile_w + 0.2)
            add_rect(slide, tl, tile_top, tile_w, 1.5, fill=pal.dark_navy)
            tb = add_textbox(slide, tl + 0.1, tile_top + 0.18, tile_w - 0.2,
                             0.55)
            write_paragraph(tb.text_frame, st["value"],
                            size=typo.title_size - 2, bold=True,
                            color=pal.bright_blue, family=typo.family,
                            align=PP_ALIGN.CENTER, first=True)
            enable_text_shrink(tb.text_frame)
            tb = add_textbox(slide, tl + 0.1, tile_top + 0.78, tile_w - 0.2,
                             0.62)
            write_paragraph(tb.text_frame, st["label"], size=typo.small_size,
                            color=pal.white, family=typo.family,
                            align=PP_ALIGN.CENTER, first=True)
            enable_text_shrink(tb.text_frame)

    # Right: real screenshot if available, else the loud placeholder
    shot_left = layout.margin_left_in + left_w + 0.4
    shot_w = layout.slide_width_in - layout.margin_right_in - shot_left
    shot_h = 5.0
    if image_path:
        from PIL import Image as _Image
        iw, ih = _Image.open(image_path).size
        scale = min(shot_w / iw, shot_h / ih)
        w, h = iw * scale, ih * scale
        px = shot_left + (shot_w - w) / 2
        py = top + (shot_h - h) / 2
        add_rect(slide, px - 0.03, py - 0.03, w + 0.06, h + 0.06,
                 fill=pal.dark_navy)
        slide.shapes.add_picture(image_path, Inches(px), Inches(py),
                                 width=Inches(w))
        if image_caption:
            tb = add_textbox(slide, px, py + h + 0.05, w, 0.22)
            write_paragraph(tb.text_frame, image_caption,
                            size=typo.footer_size, italic=True,
                            color=pal.footer_gray, family=typo.family,
                            align=PP_ALIGN.CENTER, first=True)
    else:
        placeholder_box(slide, shot_left, top, shot_w, shot_h,
                        placeholder_label, placeholder_note, theme=theme)
    return slide


# ---------- case study ----------

def add_case_slide(prs, *,
                   title: str,
                   case_name: str,
                   sector_chip: str,
                   situation: Sequence[str],
                   outcome: Sequence[str],
                   kpis: Sequence[dict],
                   photo_label: Optional[str] = None,
                   page_number=None, section_marker=None,
                   source=None, footnote=None,
                   theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    # Case header band
    top = 1.55
    band_h = 0.52
    add_rect(slide, layout.margin_left_in, top,
             layout.slide_width_in - layout.margin_left_in
             - layout.margin_right_in, band_h, fill=pal.dark_navy)
    tb = add_textbox(slide, layout.margin_left_in + 0.25, top, 7.5, band_h,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, case_name, size=typo.body_size + 4,
                    bold=True, color=pal.white, family=typo.family, first=True)
    chip_w = 2.6
    chip_left = layout.slide_width_in - layout.margin_right_in - chip_w - 0.15
    add_rect(slide, chip_left, top + 0.10, chip_w, band_h - 0.20,
             fill=pal.bright_blue)
    tb = add_textbox(slide, chip_left, top + 0.10, chip_w, band_h - 0.20,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, sector_chip, size=typo.small_size,
                    bold=True, color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    # Two columns: situation -> outcome
    col_top = top + band_h + 0.25
    col_h = 2.95
    col_w = 5.35
    right_left = layout.slide_width_in - layout.margin_right_in - col_w

    for col_left, head, items, head_fill in (
        (layout.margin_left_in, "THE SITUATION", situation, pal.mid_blue),
        (right_left, "AFTER THE DEAL", outcome, pal.bright_blue),
    ):
        add_rect(slide, col_left, col_top, col_w, 0.38, fill=head_fill)
        tb = add_textbox(slide, col_left + 0.2, col_top, col_w - 0.4, 0.38,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, head, size=typo.body_size, bold=True,
                        color=pal.white, family=typo.family, first=True)
        add_rect(slide, col_left, col_top + 0.38, col_w, col_h - 0.38,
                 fill=pal.soft_gray)
        tb = add_textbox(slide, col_left + 0.25, col_top + 0.55,
                         col_w - 0.5, col_h - 0.65)
        for j, b in enumerate(items):
            write_paragraph(tb.text_frame, b, size=typo.body_size,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=6)
        enable_text_shrink(tb.text_frame)

    # Connecting arrow between the columns
    mid_y = col_top + col_h / 2
    ar_left = layout.margin_left_in + col_w + 0.12
    ar_w = right_left - ar_left - 0.12
    tb = add_textbox(slide, ar_left, mid_y - 0.4, ar_w, 0.8,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "→", size=typo.title_size + 14, bold=True,
                    color=pal.bright_blue, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    # KPI band at the bottom
    kpi_top = col_top + col_h + 0.22
    n = max(len(kpis), 1)
    kpi_w = (layout.slide_width_in - layout.margin_left_in
             - layout.margin_right_in - 0.2 * (n - 1)) / n
    for i, k in enumerate(kpis):
        kl = layout.margin_left_in + i * (kpi_w + 0.2)
        add_rect(slide, kl, kpi_top, kpi_w, 1.05, fill=pal.deep_navy)
        tb = add_textbox(slide, kl + 0.1, kpi_top + 0.12, kpi_w - 0.2, 0.45)
        write_paragraph(tb.text_frame, k["value"], size=typo.title_size - 4,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, kl + 0.1, kpi_top + 0.58, kpi_w - 0.2, 0.42)
        write_paragraph(tb.text_frame, k["label"], size=typo.small_size,
                        color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- access ladder ----------

def add_access_ladder(prs, *,
                      title: str,
                      steps: Sequence[dict],
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """3 ascending blocks left->right (community -> marketplace -> advisory).
    Each step: name, stat, description bullets."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    n = len(steps)
    gap = 0.45
    total_w = layout.slide_width_in - layout.margin_left_in - layout.margin_right_in
    col_w = (total_w - gap * (n - 1)) / n
    base_y = 6.55
    heights = [2.6, 3.4, 4.2][:n]
    fills = [pal.mid_blue, pal.dark_navy, pal.deep_navy][:n]

    for i, st in enumerate(steps):
        left = layout.margin_left_in + i * (col_w + gap)
        h = heights[i]
        top = base_y - h
        add_rect(slide, left, top, col_w, h, fill=fills[i])
        # Step number pill
        add_oval(slide, left + 0.25, top + 0.22, 0.46, 0.46,
                 fill=pal.bright_blue)
        tb = add_textbox(slide, left + 0.25, top + 0.22, 0.46, 0.46,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, str(i + 1), size=typo.body_size + 2,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, left + 0.85, top + 0.22, col_w - 1.0, 0.5,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, st["name"], size=typo.body_size + 2,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.3, top + 0.85, col_w - 0.6, 0.55)
        write_paragraph(tb.text_frame, st["stat"], size=typo.title_size - 2,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.3, top + 1.45, col_w - 0.6,
                         h - 1.6)
        for j, b in enumerate(st.get("bullets", [])):
            write_paragraph(tb.text_frame, b, size=typo.body_size - 1,
                            color=pal.white, family=typo.family, bullet=True,
                            first=(j == 0), space_after=5)
        enable_text_shrink(tb.text_frame)

    # Ascending arrow above the steps
    tb = add_textbox(slide, layout.margin_left_in, 1.62, total_w, 0.4)
    write_paragraph(tb.text_frame,
                    "Wider access, deeper support →",
                    size=typo.body_size + 1, bold=True, italic=False,
                    color=pal.footer_gray, family=typo.family,
                    align=PP_ALIGN.RIGHT, first=True)
    return slide


# ---------- CTA ----------

def add_cta_slide(prs, *,
                  title: str,
                  paths: Sequence[dict],
                  bottom_actions: Sequence[str],
                  qr_label: str = "QR code to booking page",
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    n = len(paths)
    gap = 0.4
    cards_w = 9.0
    card_w = (cards_w - gap * (n - 1)) / n
    top = 1.75
    card_h = 3.15

    for i, p in enumerate(paths):
        left = layout.margin_left_in + i * (card_w + gap)
        add_rect(slide, left, top, card_w, card_h, fill=pal.soft_gray)
        add_rect(slide, left, top, card_w, 0.5, fill=pal.dark_navy)
        tb = add_textbox(slide, left + 0.15, top, card_w - 0.3, 0.5,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, p["who"], size=typo.body_size + 1,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.25, top + 0.7, card_w - 0.5, 1.1)
        write_paragraph(tb.text_frame, p["action"], size=typo.body_size + 1,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.25, top + 1.8, card_w - 0.5,
                         card_h - 1.95)
        write_paragraph(tb.text_frame, p["detail"], size=typo.body_size - 1,
                        color=pal.text_dark, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)

    # QR placeholder on the right
    qr_left = layout.margin_left_in + cards_w + 0.45
    qr_w = layout.slide_width_in - layout.margin_right_in - qr_left
    placeholder_box(slide, qr_left, top, qr_w, card_h, qr_label,
                    "Scan to book", theme=theme)

    # Bottom action ribbon
    rib_top = top + card_h + 0.35
    rib_w = layout.slide_width_in - layout.margin_left_in - layout.margin_right_in
    add_rect(slide, layout.margin_left_in, rib_top, rib_w, 1.15,
             fill=pal.deep_navy)
    tb = add_textbox(slide, layout.margin_left_in + 0.4, rib_top + 0.12,
                     rib_w - 0.8, 0.95)
    for j, a in enumerate(bottom_actions):
        write_paragraph(tb.text_frame, a, size=typo.body_size + 2,
                        bold=(j == 0), color=pal.white if j == 0
                        else pal.light_blue, family=typo.family,
                        first=(j == 0), space_after=4)
    enable_text_shrink(tb.text_frame)
    return slide


# ---------- thank you ----------

def add_thank_you(prs, *,
                  headline: str = "Thank you",
                  lines: Sequence[str] = (),
                  contact_placeholder: Optional[str] = None,
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    add_rect(slide, 0, 0, layout.slide_width_in, layout.slide_height_in,
             fill=pal.deep_navy)
    add_logo(slide, LOGO_WHITE, layout.slide_width_in / 2 - 1.25, 1.0, 2.5)

    tb = add_textbox(slide, 1.5, 3.0, layout.slide_width_in - 3.0, 1.0)
    write_paragraph(tb.text_frame, headline, size=typo.title_size + 22,
                    bold=True, color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    tb = add_textbox(slide, 1.5, 4.15, layout.slide_width_in - 3.0, 1.3)
    for j, l in enumerate(lines):
        write_paragraph(tb.text_frame, l, size=typo.body_size + 3,
                        color=pal.light_blue, family=typo.family,
                        align=PP_ALIGN.CENTER, first=(j == 0), space_after=6)

    if contact_placeholder:
        placeholder_box(slide, layout.slide_width_in / 2 - 2.4,
                        5.65, 4.8, 1.15, contact_placeholder, theme=theme)
    return slide


# ---------- bottom progress tracker ----------

TRACKER_STOPS = ["WHY NOW", "THE LENS", "ACCESS", "THE PATH", "PROOF",
                 "NEXT STEP"]


def add_progress_tracker(slide, active_index, theme: Theme = MAX_THEME):
    """Slim 6-stop journey bar along the very bottom edge: completed stops in
    navy, the current stop in cyan, upcoming in gray. Lets a viewer who dozed
    off re-orient instantly."""
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    h = 0.17
    top = layout.slide_height_in - h
    n = len(TRACKER_STOPS)
    seg_w = layout.slide_width_in / n
    for i, name in enumerate(TRACKER_STOPS):
        left = i * seg_w
        if i == active_index:
            fill, txt = pal.bright_blue, pal.white
        elif i < active_index:
            fill, txt = pal.dark_navy, pal.white
        else:
            fill, txt = pal.light_gray, pal.footer_gray
        add_rect(slide, left, top, seg_w, h, fill=fill)
        tb = add_textbox(slide, left + 0.06, top, seg_w - 0.12, h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, f"{i + 1} · {name}", size=7.5,
                        bold=(i == active_index), color=txt,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
    return slide


# ---------- three cards (full-width column version of three trends) ----------

def add_three_cards(prs, *,
                    title: str,
                    subtitle: Optional[str] = None,
                    cards: Sequence[dict],
                    page_number=None, section_marker=None,
                    source=None, footnote=None,
                    theme: Theme = MAX_THEME):
    """Three equal columns spanning the full width: icon chip, label, bullets.
    Replaces the engine's left-stacked three_trends layout for webinar
    full-screen sharing."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    top = 1.55
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, top,
                         layout.slide_width_in - layout.margin_left_in
                         - layout.margin_right_in, 0.30)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.text_dark, family=typo.family, first=True)
        top += 0.45

    n = len(cards)
    gap = 0.45
    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)
    card_w = (total_w - gap * (n - 1)) / n
    card_h = layout.footer_top_in - top - 0.30

    for i, c in enumerate(cards):
        left = layout.margin_left_in + i * (card_w + gap)
        add_rect(slide, left, top, card_w, card_h, fill=pal.soft_gray)
        add_rect(slide, left, top, card_w, 0.12, fill=pal.bright_blue)
        # Icon chip
        chip_d = 0.85
        add_oval(slide, left + (card_w - chip_d) / 2, top + 0.35,
                 chip_d, chip_d, fill=pal.dark_navy)
        tb = add_textbox(slide, left + (card_w - chip_d) / 2, top + 0.35,
                         chip_d, chip_d, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, c.get("icon", str(i + 1)),
                        size=typo.title_size - 2, color=pal.white,
                        family=typo.family, align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, left + 0.25, top + 1.35, card_w - 0.5, 0.45)
        write_paragraph(tb.text_frame, c["label"], size=typo.body_size + 3,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.35, top + 1.95, card_w - 0.7,
                         card_h - 2.15)
        for j, bl in enumerate(c.get("bullets", [])):
            write_paragraph(tb.text_frame, bl, size=typo.body_size,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=8)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- recap cards (climax summary) ----------

def add_recap_cards(prs, *,
                    title: str,
                    cards: Sequence[dict],
                    conclusion: str,
                    page_number=None, section_marker=None,
                    source=None, footnote=None,
                    theme: Theme = MAX_THEME):
    """Three navy-headed takeaway cards + full-width navy conclusion banner."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    top = 1.90
    n = len(cards)
    gap = 0.40
    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)
    card_w = (total_w - gap * (n - 1)) / n
    card_h = 2.80

    for i, c in enumerate(cards):
        left = layout.margin_left_in + i * (card_w + gap)
        add_rect(slide, left, top, card_w, 0.55, fill=pal.dark_navy)
        tb = add_textbox(slide, left + 0.2, top, card_w - 0.4, 0.55,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, c["takeaway"], size=typo.body_size + 1,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        add_rect(slide, left, top + 0.55, card_w, card_h - 0.55,
                 fill=pal.soft_gray)
        tb = add_textbox(slide, left + 0.25, top + 0.75, card_w - 0.5,
                         card_h - 0.95)
        for j, bl in enumerate(c.get("bullets", [])):
            write_paragraph(tb.text_frame, bl, size=typo.body_size,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=8)
        enable_text_shrink(tb.text_frame)

    band_top = top + card_h + 0.35
    add_rect(slide, layout.margin_left_in, band_top, total_w, 1.0,
             fill=pal.deep_navy)
    tb = add_textbox(slide, layout.margin_left_in + 0.4, band_top,
                     total_w - 0.8, 1.0, anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, conclusion, size=typo.body_size + 4,
                    bold=True, color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    enable_text_shrink(tb.text_frame)
    return slide


# ---------- threshold scorecard ----------

def add_scorecard_slide(prs, *,
                        title: str,
                        subtitle: Optional[str] = None,
                        groups: Sequence[dict],
                        page_number=None, section_marker=None,
                        source=None, footnote=None,
                        theme: Theme = MAX_THEME):
    """Grouped criteria table with 'weak looks like' / 'strong looks like'
    threshold phrases, so the take-home photo teaches thresholds, not just
    labels. groups: [{name, color: 'green'|'red', rows: [{name, weak, strong}]}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    top = 1.50
    if subtitle:
        tb = add_textbox(slide, layout.margin_left_in, top,
                         layout.slide_width_in - layout.margin_left_in
                         - layout.margin_right_in, 0.28)
        write_paragraph(tb.text_frame, subtitle, size=typo.body_size,
                        color=pal.text_dark, family=typo.family, first=True)
        top += 0.40

    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)
    name_w = 3.1
    col_w = (total_w - name_w) / 2
    group_colors = {"green": rgb("2E7D32"), "red": rgb("C62828")}

    # Header row
    hdr_h = 0.34
    add_rect(slide, layout.margin_left_in + name_w, top, col_w, hdr_h,
             fill=pal.light_gray)
    add_rect(slide, layout.margin_left_in + name_w + col_w, top, col_w, hdr_h,
             fill=pal.dark_navy)
    for x, label, color in (
        (layout.margin_left_in + name_w, "WEAK LOOKS LIKE", pal.footer_gray),
        (layout.margin_left_in + name_w + col_w, "STRONG LOOKS LIKE",
         pal.white),
    ):
        tb = add_textbox(slide, x + 0.15, top, col_w - 0.3, hdr_h,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, label, size=typo.small_size, bold=True,
                        color=color, family=typo.family, first=True)
    y = top + hdr_h + 0.06

    n_rows = sum(len(g["rows"]) for g in groups)
    n_groups = len(groups)
    avail = layout.footer_top_in - 0.15 - y - n_groups * 0.30
    row_h = min(0.52, avail / n_rows)

    for g in groups:
        gc = group_colors.get(g.get("color", "green"), pal.dark_navy)
        add_rect(slide, layout.margin_left_in, y, total_w, 0.26, fill=gc)
        tb = add_textbox(slide, layout.margin_left_in + 0.15, y,
                         total_w - 0.3, 0.26, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, g["name"], size=typo.small_size,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        y += 0.30
        for ri, row in enumerate(g["rows"]):
            if ri % 2 == 0:
                add_rect(slide, layout.margin_left_in, y, total_w, row_h,
                         fill=pal.soft_gray)
            tb = add_textbox(slide, layout.margin_left_in + 0.12, y,
                             name_w - 0.2, row_h, anchor=MSO_ANCHOR.MIDDLE)
            write_paragraph(tb.text_frame, row["name"], size=typo.body_size - 1,
                            bold=True, color=pal.dark_navy,
                            family=typo.family, first=True)
            enable_text_shrink(tb.text_frame)
            for x, key, color in (
                (layout.margin_left_in + name_w, "weak", pal.footer_gray),
                (layout.margin_left_in + name_w + col_w, "strong",
                 pal.text_dark),
            ):
                tb = add_textbox(slide, x + 0.15, y, col_w - 0.3, row_h,
                                 anchor=MSO_ANCHOR.MIDDLE)
                write_paragraph(tb.text_frame, row[key],
                                size=typo.body_size - 2, color=color,
                                family=typo.family, first=True)
                enable_text_shrink(tb.text_frame)
            y += row_h
    return slide


# ---------- feature pick (hero panel + runners) ----------

def add_feature_pick(prs, *,
                     title: str,
                     hero: dict,
                     runners: Sequence[dict],
                     caveat: Optional[str] = None,
                     page_number=None, section_marker=None,
                     source=None, footnote=None,
                     theme: Theme = MAX_THEME):
    """Left: one hero pick on a navy panel with big numbers. Right: ranked
    runner rows. hero: {name, headline, stats:[{value,label}], line}
    runners: [{name, stat, note}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    top = 1.65
    panel_w = 5.4
    panel_h = layout.footer_top_in - top - 0.25
    add_rect(slide, layout.margin_left_in, top, panel_w, panel_h,
             fill=pal.deep_navy)
    add_rect(slide, layout.margin_left_in, top, panel_w, 0.14,
             fill=pal.bright_blue)
    tb = add_textbox(slide, layout.margin_left_in + 0.35, top + 0.35,
                     panel_w - 0.7, 0.35)
    write_paragraph(tb.text_frame, "OUR PICK", size=typo.small_size, bold=True,
                    color=pal.bright_blue, family=typo.family, first=True)
    tb = add_textbox(slide, layout.margin_left_in + 0.35, top + 0.70,
                     panel_w - 0.7, 0.55)
    write_paragraph(tb.text_frame, hero["name"], size=typo.title_size,
                    bold=True, color=pal.white, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)
    tb = add_textbox(slide, layout.margin_left_in + 0.35, top + 1.30,
                     panel_w - 0.7, 0.42)
    write_paragraph(tb.text_frame, hero["headline"], size=typo.body_size + 1,
                    color=pal.light_blue, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)
    # Stat tiles inside the panel
    stats = hero.get("stats", [])
    if stats:
        tile_w = (panel_w - 0.7 - 0.2 * (len(stats) - 1)) / len(stats)
        for i, st in enumerate(stats):
            tl = layout.margin_left_in + 0.35 + i * (tile_w + 0.2)
            tt = top + 1.95
            tb = add_textbox(slide, tl, tt, tile_w, 0.55)
            write_paragraph(tb.text_frame, st["value"],
                            size=typo.title_size - 2, bold=True,
                            color=pal.bright_blue, family=typo.family,
                            first=True)
            enable_text_shrink(tb.text_frame)
            tb = add_textbox(slide, tl, tt + 0.55, tile_w, 0.55)
            write_paragraph(tb.text_frame, st["label"], size=typo.small_size,
                            color=pal.white, family=typo.family, first=True)
            enable_text_shrink(tb.text_frame)
    if hero.get("line"):
        add_line(slide, layout.margin_left_in + 0.35, top + panel_h - 1.05,
                 layout.margin_left_in + panel_w - 0.35, top + panel_h - 1.05,
                 color=pal.mid_blue, width_pt=0.75)
        tb = add_textbox(slide, layout.margin_left_in + 0.35,
                         top + panel_h - 0.90, panel_w - 0.7, 0.75)
        write_paragraph(tb.text_frame, hero["line"], size=typo.body_size,
                        color=pal.white, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)

    # Runner rows on the right
    rleft = layout.margin_left_in + panel_w + 0.45
    rwidth = layout.slide_width_in - layout.margin_right_in - rleft
    row_h = 1.05
    ry = top + 0.10
    for i, r in enumerate(runners):
        add_line(slide, rleft, ry + row_h, rleft + rwidth, ry + row_h,
                 color=pal.grid_gray, width_pt=0.5)
        tb = add_textbox(slide, rleft, ry + 0.06, rwidth - 1.9, 0.38)
        write_paragraph(tb.text_frame, r["name"], size=typo.body_size + 1,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, rleft, ry + 0.46, rwidth - 1.9, 0.52)
        write_paragraph(tb.text_frame, r["note"], size=typo.body_size - 1,
                        color=pal.text_dark, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, rleft + rwidth - 1.8, ry + 0.10, 1.8, 0.8,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, r["stat"], size=typo.title_size - 4,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        align=PP_ALIGN.RIGHT, first=True)
        enable_text_shrink(tb.text_frame)
        ry += row_h + 0.12
    if caveat:
        tb = add_textbox(slide, rleft, ry + 0.05, rwidth, 0.6)
        write_paragraph(tb.text_frame, caveat, size=typo.small_size,
                        italic=True, color=pal.footer_gray,
                        family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
    return slide


# ---------- succession profile cards ----------

def add_profile_cards(prs, *,
                      title: str,
                      headline_stat: str,
                      headline_label: str,
                      section_chips: Sequence[dict],
                      profiles: Sequence[dict],
                      disclaimer: str,
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """Headline count band + per-section chips + 3 anonymized company profile
    cards styled like Max Data records. profiles: {sector, province, facts:[...]}"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)

    # Headline band: big count + label + section chips
    top = 1.60
    band_h = 1.05
    add_rect(slide, layout.margin_left_in, top, total_w, band_h,
             fill=pal.deep_navy)
    tb = add_textbox(slide, layout.margin_left_in + 0.35, top, 2.9, band_h,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, headline_stat, size=typo.title_size + 10,
                    bold=True, color=pal.bright_blue, family=typo.family,
                    first=True)
    tb = add_textbox(slide, layout.margin_left_in + 3.35, top + 0.12, 3.55,
                     band_h - 0.24, anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, headline_label, size=typo.body_size + 1,
                    color=pal.white, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)
    chip_w = 1.18
    cx = layout.slide_width_in - layout.margin_right_in - 0.2 \
        - chip_w * len(section_chips) - 0.12 * (len(section_chips) - 1)
    for i, ch in enumerate(section_chips):
        left = cx + i * (chip_w + 0.12)
        add_rect(slide, left, top + 0.18, chip_w, band_h - 0.36,
                 fill=pal.mid_blue)
        tb = add_textbox(slide, left, top + 0.22, chip_w, 0.34)
        write_paragraph(tb.text_frame, ch["value"], size=typo.body_size + 1,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, left, top + 0.56, chip_w, 0.30)
        write_paragraph(tb.text_frame, ch["label"], size=7.5,
                        color=pal.light_blue, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)

    # Profile cards
    n = len(profiles)
    gap = 0.4
    card_w = (total_w - gap * (n - 1)) / n
    ctop = top + band_h + 0.30
    card_h = layout.footer_top_in - ctop - 0.55
    for i, p in enumerate(profiles):
        left = layout.margin_left_in + i * (card_w + gap)
        add_rect(slide, left, ctop, card_w, card_h, fill=pal.soft_gray)
        add_rect(slide, left, ctop, card_w, 0.50, fill=pal.dark_navy)
        tb = add_textbox(slide, left + 0.2, ctop, card_w - 0.4, 0.50,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, p["sector"], size=typo.body_size,
                        bold=True, color=pal.white, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.25, ctop + 0.62, card_w - 0.5, 0.30)
        write_paragraph(tb.text_frame, p["province"], size=typo.small_size,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        first=True)
        tb = add_textbox(slide, left + 0.25, ctop + 0.95, card_w - 0.5,
                         card_h - 1.10)
        for j, f in enumerate(p["facts"]):
            write_paragraph(tb.text_frame, f, size=typo.body_size - 1,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=5)
        enable_text_shrink(tb.text_frame)

    tb = add_textbox(slide, layout.margin_left_in,
                     layout.footer_top_in - 0.42, total_w, 0.30)
    write_paragraph(tb.text_frame, disclaimer, size=typo.small_size,
                    italic=True, color=pal.footer_gray, family=typo.family,
                    first=True)
    return slide


# ---------- register everything ----------

_REGISTRY.update({
    "max_cover": add_max_cover,
    "speaker_slide": add_speaker_slide,
    "poll_slide": add_poll_slide,
    "screenshot_slide": add_screenshot_slide,
    "case_slide": add_case_slide,
    "access_ladder": add_access_ladder,
    "cta_slide": add_cta_slide,
    "thank_you": add_thank_you,
    "three_cards": add_three_cards,
    "recap_cards": add_recap_cards,
    "scorecard_slide": add_scorecard_slide,
    "feature_pick": add_feature_pick,
    "profile_cards": add_profile_cards,
})
