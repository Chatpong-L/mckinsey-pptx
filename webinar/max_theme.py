"""Max Solutions brand theme for the mckinsey_pptx engine.

Palette per the Max Solutions house style (report-style module 02):
navy #022859 carries authority, cyan #18B9FF points the eye.
"""
from dataclasses import replace

from mckinsey_pptx.theme import DEFAULT_THEME, Palette, Typography, rgb

MAX_NAVY = "022859"
MAX_NAVY_DEEP = "011C40"
MAX_CYAN = "18B9FF"
MAX_LIGHT_NAVY = "2C4B73"
MAX_SLATE = "4A5A70"
MAX_TEAL = "2E8B8B"
MAX_SOFT_BG = "EEF3F8"
MAX_GRID = "D9DEE6"
MAX_MUTED = "6B7686"
MAX_BODY = "232B36"

# Placeholder treatment: deliberately off-brand amber so it screams "replace me".
PH_BORDER = "F59E0B"
PH_FILL = "FFF7E6"
PH_TEXT = "B45309"

MAX_PALETTE = Palette(
    dark_navy=rgb(MAX_NAVY),
    deep_navy=rgb(MAX_NAVY_DEEP),
    bright_blue=rgb(MAX_CYAN),
    mid_blue=rgb(MAX_LIGHT_NAVY),
    light_blue=rgb("6FCFFF"),
    royal_blue=rgb(MAX_TEAL),
    text_dark=rgb(MAX_BODY),
    rule_gray=rgb(MAX_GRID),
    light_gray=rgb("E4EAF1"),
    soft_gray=rgb(MAX_SOFT_BG),
    grid_gray=rgb(MAX_GRID),
    footer_gray=rgb(MAX_MUTED),
    placeholder_gray=rgb("9AA6B5"),
)

MAX_THEME = replace(
    DEFAULT_THEME,
    palette=MAX_PALETTE,
    typography=replace(DEFAULT_THEME.typography, family="Arial"),
    # Footer sits slightly higher than the engine default to clear the
    # bottom progress tracker strip.
    layout=replace(DEFAULT_THEME.layout, footer_top_in=6.98),
    copyright_text="© 2026 Max Solutions",
)
