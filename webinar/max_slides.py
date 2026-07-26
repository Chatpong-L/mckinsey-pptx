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
import os

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from mckinsey_pptx.base import (
    blank_slide, add_chrome, add_footer, add_rect, add_oval, add_line,
    add_textbox, write_paragraph, enable_text_shrink,
)
from mckinsey_pptx.theme import Theme, rgb
from mckinsey_pptx.builder import _REGISTRY

from max_theme import MAX_THEME, PH_BORDER, PH_FILL, PH_TEXT

ASSETS = __file__.rsplit("/", 1)[0] + "/assets"
QR_BOOKING = f"{ASSETS}/qr-booking.png"


def qr_block(slide, left_in, top_in, size_in, theme, on_dark=False,
             path=QR_BOOKING):
    """Real QR on a white quiet-zone plate. Falls back to the amber
    placeholder if the code has not been generated yet."""
    if not os.path.exists(path):
        return False
    if on_dark:
        add_rect(slide, left_in - 0.09, top_in - 0.09, size_in + 0.18,
                 size_in + 0.18, fill=theme.palette.white)
    slide.shapes.add_picture(path, Inches(left_in), Inches(top_in),
                             width=Inches(size_in), height=Inches(size_in))
    return True
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
    first_line = True
    if width_in >= 1.6:
        # the header word only fits comfortably in wider boxes; narrow ones
        # rely on the amber dashed border to signal "placeholder"
        write_paragraph(tb.text_frame, "PLACEHOLDER", size=typo.small_size,
                        bold=True, color=rgb(PH_BORDER), family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        first_line = False
    write_paragraph(tb.text_frame, label, size=typo.body_size, bold=True,
                    color=rgb(PH_TEXT), family=typo.family,
                    align=PP_ALIGN.CENTER, first=first_line,
                    space_before=None if first_line else 2)
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

    # Cyan seam where the white field meets the hero panel
    add_rect(slide, layout.slide_width_in - panel_w - 0.055, 0, 0.055,
             layout.slide_height_in, fill=pal.bright_blue)

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

    add_line(slide, left, 5.45, left + 5.8, 5.45,
             color=pal.bright_blue, width_pt=2.0)
    if date:
        tb = add_textbox(slide, left, 5.60, 7.0, 0.35)
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
                   ornament_path: Optional[str] = None,
                   page_number=None, section_marker=None,
                   source=None, footnote=None,
                   theme: Theme = MAX_THEME):
    """Dark interrupt family: full-bleed navy, cyan LIVE POLL chip, 2x2 answer
    tiles, cyan bottom strip. Visually signals a mode change in the webinar."""
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    add_rect(slide, 0, 0, layout.slide_width_in, layout.slide_height_in,
             fill=rgb("022859"))

    if ornament_path and os.path.exists(ornament_path):
        slide.shapes.add_picture(ornament_path, Inches(9.55), Inches(2.95),
                                 width=Inches(3.78), height=Inches(3.78))
    chip_w = 1.55
    add_rect(slide, layout.margin_left_in, 0.55, chip_w, 0.40,
             fill=pal.bright_blue)
    tb = add_textbox(slide, layout.margin_left_in, 0.55, chip_w, 0.40,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "LIVE POLL", size=typo.body_size, bold=True,
                    color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    tb = add_textbox(slide, layout.margin_left_in, 1.25,
                     layout.slide_width_in - layout.margin_left_in
                     - layout.margin_right_in, 1.0)
    write_paragraph(tb.text_frame, question, size=32, bold=True,
                    color=pal.white, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)

    # 2x2 answer tiles
    letters = "ABCD"
    tile_w = 5.9
    tile_h = 1.35
    gx, gy = layout.margin_left_in, 2.65
    gapx, gapy = 0.55, 0.45
    for i, opt in enumerate(options[:4]):
        col_i, row_i = i % 2, i // 2
        x = gx + col_i * (tile_w + gapx)
        y = gy + row_i * (tile_h + gapy)
        add_rect(slide, x, y, tile_w, tile_h, fill=rgb("0A3A73"))
        d = 0.62
        add_oval(slide, x + 0.28, y + (tile_h - d) / 2, d, d,
                 fill=pal.bright_blue)
        tb = add_textbox(slide, x + 0.28, y + (tile_h - d) / 2, d, d,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, letters[i], size=typo.body_size + 4,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tb = add_textbox(slide, x + 1.15, y + 0.15, tile_w - 1.4,
                         tile_h - 0.3, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, opt, size=typo.body_size + 3,
                        color=pal.white, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)

    # Bottom cyan strip
    strip_h = 0.62
    add_rect(slide, 0, layout.slide_height_in - 0.17 - strip_h,
             layout.slide_width_in, strip_h, fill=pal.bright_blue)
    tb = add_textbox(slide, 0, layout.slide_height_in - 0.17 - strip_h,
                     layout.slide_width_in, strip_h,
                     anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, instruction, size=typo.body_size + 3,
                    bold=True, color=rgb("022859"), family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)
    return slide


# ---------- product screenshot walkthrough ----------

def add_screenshot_slide(prs, *,
                         title: str,
                         placeholder_label: str,
                         placeholder_note: Optional[str] = None,
                         image_path: Optional[str] = None,
                         image_caption: Optional[str] = None,
                         kicker: Optional[str] = None,
                         claim: Optional[str] = None,
                         bullets: Sequence[str] = (),
                         stats: Sequence[dict] = (),
                         layout_mode: str = "right",
                         shot_width: Optional[float] = None,
                         overlap_stat: Optional[dict] = None,
                         page_number=None, section_marker=None,
                         source=None, footnote=None,
                         theme: Theme = MAX_THEME):
    """Product-shot slide with a browser frame. layout_mode: 'right' (text
    left, shot right), 'left' (mirrored), 'hero' (centered shot, stats
    overlapping its bottom edge)."""
    from editorial_slides import browser_frame

    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    top = 1.70

    def draw_shot(sx, sy, sw, sh):
        inner = browser_frame(slide, sx, sy, sw, sh, theme=theme)
        il, it, iw, ih = inner
        if image_path:
            from PIL import Image as _Image
            pw, ph = _Image.open(image_path).size
            scale = min((iw - 0.06) / pw, (ih - 0.06) / ph)
            w, h = pw * scale, ph * scale
            px = il + (iw - w) / 2
            py = it + (ih - h) / 2
            slide.shapes.add_picture(image_path, Inches(px), Inches(py),
                                     width=Inches(w))
            if image_caption and not (layout_mode == "hero" and stats):
                # keep the caption clear of the overlap chip, which hangs
                # below the shot on one side; in hero mode the stat chips
                # straddle the shot's bottom edge and would cover it
                cap_x, cap_w = sx, sw
                if overlap_stat and layout_mode == "right":
                    cap_x, cap_w = sx + 1.75, sw - 1.75
                elif overlap_stat and layout_mode == "left":
                    cap_w = sw - 1.75
                tb = add_textbox(slide, cap_x, sy + sh + 0.06, cap_w, 0.22)
                write_paragraph(tb.text_frame, image_caption,
                                size=typo.footer_size, italic=True,
                                color=pal.footer_gray, family=typo.family,
                                align=PP_ALIGN.CENTER, first=True)
        else:
            placeholder_box(slide, il + 0.04, it + 0.04, iw - 0.08,
                            ih - 0.08, placeholder_label, placeholder_note,
                            theme=theme)

    def draw_text_col(tx, tw):
        ty = top + 0.15
        if kicker:
            tb = add_textbox(slide, tx, ty, tw, 0.28)
            write_paragraph(tb.text_frame, kicker.upper(),
                            size=typo.small_size, bold=True,
                            color=pal.bright_blue, family=typo.family,
                            first=True)
            ty += 0.42
        if claim:
            tb = add_textbox(slide, tx, ty, tw, 0.85)
            write_paragraph(tb.text_frame, claim, size=typo.body_size + 4,
                            bold=True, color=pal.dark_navy,
                            family=typo.family, first=True)
            enable_text_shrink(tb.text_frame)
            ty += 1.0
        tb = add_textbox(slide, tx, ty, tw, 2.1)
        for j, bl in enumerate(bullets):
            write_paragraph(tb.text_frame, bl, size=typo.body_size + 1,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=8)
        enable_text_shrink(tb.text_frame)
        # open stat strip, hairline separated
        if stats:
            sy = 5.15
            seg_w = tw / len(stats)
            for i, st in enumerate(stats):
                sx2 = tx + i * seg_w
                tb = add_textbox(slide, sx2, sy, seg_w - 0.1, 0.5)
                write_paragraph(tb.text_frame, st["value"], size=26,
                                bold=True, color=pal.dark_navy,
                                family=typo.family, first=True)
                tb = add_textbox(slide, sx2, sy + 0.52, seg_w - 0.1, 0.55)
                write_paragraph(tb.text_frame, st["label"].upper(),
                                size=8.5, color=pal.footer_gray,
                                family=typo.family, first=True)
                enable_text_shrink(tb.text_frame)
                if i < len(stats) - 1:
                    add_line(slide, sx2 + seg_w - 0.18, sy + 0.05,
                             sx2 + seg_w - 0.18, sy + 0.95,
                             color=pal.grid_gray, width_pt=0.75)

    if layout_mode == "hero":
        if shot_width:
            sw = shot_width
            # size the frame so the image exactly fills the inner box
            if image_path and os.path.exists(image_path):
                from PIL import Image as _Im
                _w, _h = _Im.open(image_path).size
                sh = (sw - 0.06) * _h / _w + 0.30 + 0.06
            else:
                sh = sw * 0.52
        else:
            sw, sh = 8.8, 4.35
        sx = (layout.slide_width_in - sw) / 2
        if claim:
            tb = add_textbox(slide, layout.margin_left_in, 1.42,
                             layout.slide_width_in - 2 * layout.margin_left_in,
                             0.42)
            write_paragraph(tb.text_frame, claim, size=18, bold=True,
                            color=pal.dark_navy, family=typo.family,
                            align=PP_ALIGN.CENTER, first=True)
        draw_shot(sx, 1.95, sw, sh)
        if stats:
            chip_w, chip_h = 1.9, 0.95
            total = chip_w * len(stats) + 0.35 * (len(stats) - 1)
            cx0 = (layout.slide_width_in - total) / 2
            cy0 = 1.95 + sh - chip_h / 2
            for i, st in enumerate(stats):
                cx = cx0 + i * (chip_w + 0.35)
                add_rect(slide, cx, cy0, chip_w, chip_h, fill=pal.deep_navy)
                tb = add_textbox(slide, cx, cy0 + 0.08, chip_w, 0.45)
                write_paragraph(tb.text_frame, st["value"], size=22,
                                bold=True, color=pal.bright_blue,
                                family=typo.family, align=PP_ALIGN.CENTER,
                                first=True)
                tb = add_textbox(slide, cx, cy0 + 0.55, chip_w, 0.34)
                write_paragraph(tb.text_frame, st["label"].upper(), size=8,
                                color=pal.white, family=typo.family,
                                align=PP_ALIGN.CENTER, first=True)
                enable_text_shrink(tb.text_frame)
    elif layout_mode == "left":
        sw = 7.0
        sh = 4.7
        draw_shot(layout.margin_left_in, top, sw, sh)
        tx = layout.margin_left_in + sw + 0.5
        draw_text_col(tx, layout.slide_width_in - layout.margin_right_in - tx)
        if overlap_stat:
            _overlap_chip(slide, layout.margin_left_in + sw - 1.6,
                          top + sh - 0.55, overlap_stat, theme)
    else:
        text_w = 4.5
        sx = layout.margin_left_in + text_w + 0.5
        sw = layout.slide_width_in - layout.margin_right_in - sx
        draw_shot(sx, top, sw, 4.7)
        draw_text_col(layout.margin_left_in, text_w)
        if overlap_stat:
            _overlap_chip(slide, sx - 0.5, top + 4.7 - 0.55, overlap_stat,
                          theme)
    return slide


def _overlap_chip(slide, x, y, stat, theme):
    """White chip with cyan outline overlapping a screenshot corner."""
    pal, typo = theme.palette, theme.typography
    w, h = 2.0, 1.1
    add_rect(slide, x, y, w, h, fill=pal.white, line=pal.bright_blue,
             line_width=1.5)
    tb = add_textbox(slide, x + 0.12, y + 0.10, w - 0.24, 0.55)
    write_paragraph(tb.text_frame, stat["value"], size=30, bold=True,
                    color=theme.palette.dark_navy, family=typo.family,
                    first=True)
    tb = add_textbox(slide, x + 0.12, y + 0.66, w - 0.24, 0.36)
    write_paragraph(tb.text_frame, stat["label"], size=typo.small_size,
                    color=theme.palette.footer_gray, family=typo.family,
                    first=True)


# ---------- case study ----------

def add_case_slide(prs, *,
                   title: str,
                   case_name: str,
                   sector_chip: str,
                   situation: Sequence[str],
                   outcome: Sequence[str],
                   kpis: Sequence[dict],
                   bridge_stat: Optional[str] = None,
                   photo_label: Optional[str] = None,
                   photo_path: Optional[str] = None,
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

    # Two open columns with left rules, bridged by a cyan chevron carrying
    # the transformation stat.
    col_top = top + band_h + 0.35
    col_h = 2.75
    col_w = 4.35
    right_left = layout.slide_width_in - layout.margin_right_in - col_w

    for col_left, head, items, accent in (
        (layout.margin_left_in, "THE SITUATION", situation, pal.mid_blue),
        (right_left, "AFTER THE DEAL", outcome, pal.bright_blue),
    ):
        add_rect(slide, col_left, col_top, 0.03, col_h, fill=accent)
        tb = add_textbox(slide, col_left + 0.25, col_top, col_w - 0.4, 0.30)
        write_paragraph(tb.text_frame, head, size=typo.small_size + 1,
                        bold=True, color=accent, family=typo.family,
                        first=True)
        tb = add_textbox(slide, col_left + 0.25, col_top + 0.45,
                         col_w - 0.5, col_h - 0.5)
        for j, b in enumerate(items[:3]):
            write_paragraph(tb.text_frame, b, size=typo.body_size + 1,
                            color=pal.text_dark, family=typo.family,
                            bullet=True, first=(j == 0), space_after=8)
        enable_text_shrink(tb.text_frame)

    # Chevron bridge with the transformation stat
    if bridge_stat:
        mid_y = col_top + col_h / 2
        ch_w = right_left - (layout.margin_left_in + col_w) + 0.5
        ch_x = layout.margin_left_in + col_w - 0.25
        ch = slide.shapes.add_shape(
            MSO_SHAPE.CHEVRON, Inches(ch_x), Inches(mid_y - 0.34),
            Inches(ch_w), Inches(0.68))
        ch.shadow.inherit = False
        ch.fill.solid(); ch.fill.fore_color.rgb = pal.bright_blue
        ch.line.fill.background()
        tb = add_textbox(slide, ch_x + 0.15, mid_y - 0.34, ch_w - 0.45,
                         0.68, anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, bridge_stat, size=typo.body_size,
                        bold=True, color=rgb("022859"), family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        enable_text_shrink(tb.text_frame)

    # Scene panel filling the centre gap; the chevron bridges across it
    if photo_path and os.path.exists(photo_path):
        from editorial_slides import cover_crop
        gap_left = layout.margin_left_in + col_w + 0.18
        gap_w = right_left - gap_left - 0.18
        crop = cover_crop(photo_path, gap_w, col_h)
        slide.shapes.add_picture(crop, Inches(gap_left), Inches(col_top),
                                 width=Inches(gap_w), height=Inches(col_h))

    # Open stat band at the bottom, hairline separated
    kpi_top = col_top + col_h + 0.30
    n = max(len(kpis), 1)
    seg_w = (layout.slide_width_in - layout.margin_left_in
             - layout.margin_right_in) / n
    add_line(slide, layout.margin_left_in, kpi_top,
             layout.slide_width_in - layout.margin_right_in, kpi_top,
             color=pal.grid_gray, width_pt=0.75)
    for i, k in enumerate(kpis):
        kl = layout.margin_left_in + i * seg_w
        tb = add_textbox(slide, kl + 0.1, kpi_top + 0.12, seg_w - 0.3, 0.5)
        write_paragraph(tb.text_frame, k["value"], size=26, bold=True,
                        color=pal.dark_navy, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, kl + 0.1, kpi_top + 0.64, seg_w - 0.3, 0.34)
        write_paragraph(tb.text_frame, k["label"].upper(), size=8.5,
                        color=pal.footer_gray, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        if i < n - 1:
            add_line(slide, kl + seg_w - 0.15, kpi_top + 0.15,
                     kl + seg_w - 0.15, kpi_top + 0.9,
                     color=pal.grid_gray, width_pt=0.75)
    return slide


# ---------- access ladder ----------

def add_access_ladder(prs, *,
                      title: str,
                      steps: Sequence[dict],
                      page_number=None, section_marker=None,
                      source=None, footnote=None,
                      theme: Theme = MAX_THEME):
    """Ascending steps: wide-and-shallow to narrow-and-deep (the geometry IS
    the tradeoff). steps: [{kicker, name, stat, bullets(max 2)}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    n = len(steps)
    gap = 0.45
    base_y = 6.55
    widths = [4.6, 3.9, 3.2][:n]
    heights = [2.6, 3.4, 4.2][:n]
    fills = [pal.mid_blue, pal.dark_navy, pal.deep_navy][:n]
    left = layout.margin_left_in

    corners = []
    for i, st in enumerate(steps):
        col_w = widths[i]
        h = heights[i]
        top = base_y - h
        add_rect(slide, left, top, col_w, h, fill=fills[i])
        corners.append((left, top))
        if st.get("icon") and os.path.exists(st["icon"]):
            slide.shapes.add_picture(st["icon"],
                                     Inches(left + col_w - 0.75),
                                     Inches(top + 0.16), width=Inches(0.5),
                                     height=Inches(0.5))
        tb = add_textbox(slide, left + 0.28, top + 0.20, col_w - 0.56, 0.28)
        write_paragraph(tb.text_frame, st["kicker"].upper(),
                        size=typo.small_size, bold=True,
                        color=pal.light_blue, family=typo.family, first=True)
        tb = add_textbox(slide, left + 0.28, top + 0.52, col_w - 0.56, 0.55)
        write_paragraph(tb.text_frame, st["stat"], size=typo.title_size - 2,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.28, top + 1.15, col_w - 0.56,
                         h - 1.35)
        for j, bl in enumerate(st.get("bullets", [])[:2]):
            write_paragraph(tb.text_frame, bl, size=typo.body_size - 1,
                            color=pal.white, family=typo.family, bullet=True,
                            first=(j == 0), space_after=5)
        enable_text_shrink(tb.text_frame)
        left += col_w + gap

    # Cyan ascent line across the step corners
    for i in range(len(corners) - 1):
        x1, y1 = corners[i][0] + widths[i], corners[i][1]
        x2, y2 = corners[i + 1]
        add_line(slide, x1 - 0.1, y1 - 0.15, x2 + 0.1, y2 - 0.15,
                 color=pal.bright_blue, width_pt=2.5)
    tip_x, tip_y = corners[-1][0] + widths[-1] * 0.55, corners[-1][1] - 0.15
    tb = add_textbox(slide, corners[0][0] + 0.2, corners[-1][1] - 0.75,
                     9.5, 0.32)
    write_paragraph(tb.text_frame, "WIDER ACCESS  →  DEEPER SUPPORT",
                    size=typo.small_size + 1, bold=True,
                    color=pal.footer_gray, family=typo.family, first=True)
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
    """Conversion hero: three scannable audience rows + QR right + navy
    action ribbon. paths: [{num, who, action, detail}]"""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout

    rows_w = 8.6
    top = 1.85
    row_h = 1.18
    for i, p in enumerate(paths):
        y = top + i * (row_h + 0.14)
        d = 0.5
        add_oval(slide, layout.margin_left_in, y + (row_h - d) / 2, d, d,
                 fill=pal.bright_blue)
        tb = add_textbox(slide, layout.margin_left_in,
                         y + (row_h - d) / 2, d, d,
                         anchor=MSO_ANCHOR.MIDDLE)
        write_paragraph(tb.text_frame, str(p["num"]), size=typo.body_size + 4,
                        bold=True, color=pal.white, family=typo.family,
                        align=PP_ALIGN.CENTER, first=True)
        tx = layout.margin_left_in + 0.75
        tb = add_textbox(slide, tx, y + 0.02, rows_w - 0.75, 0.26)
        write_paragraph(tb.text_frame, p["who"].upper(),
                        size=typo.small_size, bold=True,
                        color=pal.footer_gray, family=typo.family, first=True)
        tb = add_textbox(slide, tx, y + 0.28, rows_w - 0.75, 0.42)
        write_paragraph(tb.text_frame, p["action"], size=16, bold=True,
                        color=pal.dark_navy, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, tx, y + 0.72, rows_w - 0.75, 0.40)
        write_paragraph(tb.text_frame, p["detail"], size=typo.body_size - 1,
                        color=pal.footer_gray, family=typo.family, first=True)
        enable_text_shrink(tb.text_frame)
        if p.get("icon") and os.path.exists(p["icon"]):
            d2 = 0.55
            slide.shapes.add_picture(
                p["icon"], Inches(layout.margin_left_in + rows_w - d2),
                Inches(y + (row_h - d2) / 2), width=Inches(d2),
                height=Inches(d2))
        if i < len(paths) - 1:
            add_line(slide, layout.margin_left_in, y + row_h + 0.07,
                     layout.margin_left_in + rows_w, y + row_h + 0.07,
                     color=pal.grid_gray, width_pt=0.75)

    # QR placeholder on the right
    qr_left = layout.margin_left_in + rows_w + 0.55
    qr_w = layout.slide_width_in - layout.margin_right_in - qr_left
    if not qr_block(slide, qr_left, top + 0.1, qr_w, theme):
        placeholder_box(slide, qr_left, top + 0.1, qr_w, qr_w, qr_label,
                        "Scan to book", theme=theme)
    tb = add_textbox(slide, qr_left, top + 0.1 + qr_w + 0.08, qr_w, 0.3)
    write_paragraph(tb.text_frame, "Scan to book your session",
                    size=typo.small_size, color=pal.footer_gray,
                    family=typo.family, align=PP_ALIGN.CENTER, first=True)

    # Bottom action ribbon
    rib_top = top + 3 * (row_h + 0.14) + 0.25
    rib_w = layout.slide_width_in - layout.margin_left_in - layout.margin_right_in
    add_rect(slide, layout.margin_left_in, rib_top, rib_w, 0.95,
             fill=pal.deep_navy)
    tb = add_textbox(slide, layout.margin_left_in + 0.4, rib_top + 0.08,
                     rib_w - 0.8, 0.82)
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
                  bg_path: Optional[str] = None,
                  page_number=None, section_marker=None,
                  source=None, footnote=None,
                  theme: Theme = MAX_THEME):
    slide = blank_slide(prs)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    add_rect(slide, 0, 0, layout.slide_width_in, layout.slide_height_in,
             fill=pal.deep_navy)

    if bg_path and os.path.exists(bg_path):
        from editorial_slides import cover_crop
        crop = cover_crop(bg_path, layout.slide_width_in,
                          layout.slide_height_in)
        slide.shapes.add_picture(crop, Inches(0), Inches(0),
                                 width=Inches(layout.slide_width_in),
                                 height=Inches(layout.slide_height_in))
    add_logo(slide, LOGO_WHITE, layout.slide_width_in / 2 - 1.25, 1.0, 2.5)

    tb = add_textbox(slide, 1.5, 3.0, layout.slide_width_in - 3.0, 1.0)
    write_paragraph(tb.text_frame, headline, size=typo.title_size + 22,
                    bold=True, color=pal.white, family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

    tb = add_textbox(slide, 1.5, 4.02, layout.slide_width_in - 3.0, 1.0)
    for j, l in enumerate(lines):
        write_paragraph(tb.text_frame, l, size=typo.body_size + 3,
                        color=pal.light_blue, family=typo.family,
                        align=PP_ALIGN.CENTER, first=(j == 0), space_after=6)

    if contact_placeholder:
        qs = 1.42
        gap = 1.30
        line_qr = f"{ASSETS}/qr-line-group.png"
        pair = os.path.exists(line_qr)
        total = qs * 2 + gap if pair else qs
        qx = layout.slide_width_in / 2 - total / 2
        qy = 4.80
        placed = qr_block(slide, qx, qy, qs, theme, on_dark=True)
        if placed:
            tb = add_textbox(slide, qx - 0.55, qy + qs + 0.16, qs + 1.1, 0.30)
            write_paragraph(tb.text_frame, "Book a session",
                            size=typo.small_size, bold=True,
                            color=pal.light_blue, family=typo.family,
                            align=PP_ALIGN.CENTER, first=True)
            if pair:
                lx = qx + qs + gap
                qr_block(slide, lx, qy, qs, theme, on_dark=True, path=line_qr)
                tb = add_textbox(slide, lx - 0.55, qy + qs + 0.16, qs + 1.1,
                                 0.30)
                write_paragraph(tb.text_frame, "Join the LINE group",
                                size=typo.small_size, bold=True,
                                color=pal.light_blue, family=typo.family,
                                align=PP_ALIGN.CENTER, first=True)
        else:
            placeholder_box(slide, layout.slide_width_in / 2 - 1.1,
                            4.85, 2.2, 1.6, contact_placeholder, theme=theme)
    from editorial_slides import closing_strip
    closing_strip(slide, theme, height=0.7)
    return slide


# ---------- bottom progress tracker ----------

TRACKER_STOPS = ["WHY NOW", "ACCESS", "THE LENS", "THE PATH", "PROOF",
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
    """Inverted hierarchy: the conclusion IS the headline; three flat recap
    chips sit beneath it."""
    slide = blank_slide(prs)
    add_chrome(slide, title=title, theme=theme, page_number=page_number,
               section_marker=section_marker, source=source, footnote=footnote)
    pal, typo, layout = theme.palette, theme.typography, theme.layout
    total_w = (layout.slide_width_in - layout.margin_left_in
               - layout.margin_right_in)

    tb = add_textbox(slide, layout.margin_left_in, 2.0, total_w, 1.15)
    write_paragraph(tb.text_frame, conclusion, size=26, bold=True,
                    color=pal.dark_navy, family=typo.family, first=True)
    enable_text_shrink(tb.text_frame)
    add_line(slide, layout.margin_left_in, 3.25,
             layout.margin_left_in + 2.5, 3.25, color=pal.bright_blue,
             width_pt=3.0)

    n = len(cards)
    gap = 0.4
    card_w = (total_w - gap * (n - 1)) / n
    top = 3.75
    card_h = 1.75
    for i, c in enumerate(cards):
        left = layout.margin_left_in + i * (card_w + gap)
        add_rect(slide, left, top, card_w, card_h, fill=None,
                 line=pal.grid_gray, line_width=1.0)
        if c.get("icon") and os.path.exists(c["icon"]):
            slide.shapes.add_picture(c["icon"],
                                     Inches(left + card_w - 0.78),
                                     Inches(top + 0.18), width=Inches(0.55),
                                     height=Inches(0.55))
        tb = add_textbox(slide, left + 0.22, top + 0.15, 0.9, 0.4)
        write_paragraph(tb.text_frame, f"0{i + 1}", size=typo.body_size + 4,
                        bold=True, color=pal.bright_blue, family=typo.family,
                        first=True)
        tb = add_textbox(slide, left + 0.22, top + 0.55, card_w - 0.44,
                         0.40)
        write_paragraph(tb.text_frame, c["takeaway"], size=typo.body_size + 2,
                        bold=True, color=pal.dark_navy, family=typo.family,
                        first=True)
        enable_text_shrink(tb.text_frame)
        tb = add_textbox(slide, left + 0.22, top + 1.0, card_w - 0.44,
                         0.65)
        write_paragraph(tb.text_frame, c["line"], size=typo.body_size - 1,
                        color=pal.footer_gray, family=typo.family,
                        first=True)
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
    group_colors = {"green": rgb("1E6E42"), "red": rgb("B03A2E")}

    # Screenshot-this tag, top-right
    tag_w = 1.85
    add_rect(slide, layout.slide_width_in - layout.margin_right_in - tag_w,
             top - 0.40, tag_w, 0.34, fill=pal.bright_blue)
    tb = add_textbox(slide,
                     layout.slide_width_in - layout.margin_right_in - tag_w,
                     top - 0.40, tag_w, 0.34, anchor=MSO_ANCHOR.MIDDLE)
    write_paragraph(tb.text_frame, "SCREENSHOT THIS", size=typo.small_size,
                    bold=True, color=rgb("022859"), family=typo.family,
                    align=PP_ALIGN.CENTER, first=True)

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
            for x, key, color, dot in (
                (layout.margin_left_in + name_w, "weak", pal.footer_gray,
                 "open"),
                (layout.margin_left_in + name_w + col_w, "strong",
                 pal.text_dark, "filled"),
            ):
                dcy = y + row_h / 2 - 0.06
                if dot == "open":
                    add_oval(slide, x + 0.12, dcy, 0.12, 0.12, fill=None,
                             line=pal.footer_gray, line_width=1.0)
                else:
                    add_oval(slide, x + 0.12, dcy, 0.12, 0.12,
                             fill=pal.bright_blue)
                tb = add_textbox(slide, x + 0.36, y, col_w - 0.5, row_h,
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
    if hero.get("icon") and os.path.exists(hero["icon"]):
        d3 = 0.85
        slide.shapes.add_picture(
            hero["icon"],
            Inches(layout.margin_left_in + panel_w / 2 - d3 / 2),
            Inches(top + panel_h - 2.05), width=Inches(d3), height=Inches(d3))
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
