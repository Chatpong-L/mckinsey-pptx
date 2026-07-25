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

    # Right navy panel with the Max Data icon watermark
    panel_w = 4.2
    add_rect(slide, layout.slide_width_in - panel_w, 0, panel_w,
             layout.slide_height_in, fill=pal.deep_navy)
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

    # Right: screenshot placeholder
    shot_left = layout.margin_left_in + left_w + 0.4
    shot_w = layout.slide_width_in - layout.margin_right_in - shot_left
    placeholder_box(slide, shot_left, top, shot_w, 5.05, placeholder_label,
                    placeholder_note, theme=theme)
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
})
