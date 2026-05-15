#!/usr/bin/env python3
"""
Blood Bowl Aarhus Wargaming Grid 2026 – PowerPoint + PDF Generator
Produces AarhusWargmingGrid.pptx and AarhusWargmingGrid.pdf with:
  Page 1 – Team grid + star players
  Page 2 – All special rules

To update teams:        edit TEAMS dict  – values are (team_name, tier) tuples
To update star players: edit STAR_PLAYERS dict
To update rules:        edit SPECIAL_RULES dict
"""

import io
import math
import os
import subprocess

import cairosvg
from PIL import Image, ImageFilter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

# ─── PATHS ──────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SVG_PATH   = os.path.join(SCRIPT_DIR, "Warhammer Aarhus Logo_vector.svg")
OUT_PATH   = os.path.join(SCRIPT_DIR, "AarhusWargmingGrid.pptx")
BG_PATH    = os.path.join(SCRIPT_DIR, "_bg_temp.png")
LOGO_PATH  = os.path.join(SCRIPT_DIR, "_logo_temp.png")

# ─── GRID DATA ──────────────────────────────────────────────────────────────

# Gold-piece rows (team roster cost)
GP_ROWS = [650, 675, 700, 725, 750]

# Star-player-point columns (skill pick budget)
SPP_COLS = [3, 5, 6, 7, 8]

# Teams per (gp, spp) cell.  Each entry is (team_name, tier) where tier is 1/2/3.
# Slann is intentionally excluded.
# Tier colour key: T1 = green, T2 = yellow, T3 = orange, T4 = red  (see TIER_COLORS below).
TEAMS: dict[tuple[int, int], list[tuple[str, int]]] = {
    (650, 3):  [("Amazons", 1), ("OWA", 2)],
    (650, 5):  [("Orc", 2), ("Undead", 2), ("Wood Elf", 1)],
    (650, 6):  [("Human", 2), ("Dark Elf", 2), ("High Elf", 1)],
    (650, 8):  [("Bretonnian", 3), ("Snotlings", 3)],
    (650, 8): [],
    (675, 3):  [],
    (675, 5):  [("Norse", 1), ("Necromantic", 2), ("Lizardmen", 1)],
    (675, 6):  [("Tomb Kings", 3), ("Underworld", 1), ("Imperial Nobility", 2)],
    (675, 7):  [("Chaos Chosen", 2)],
    (675, 8): [],
    (700, 3):  [],
    (700, 5):  [("Nurgle", 3), ("Vampire", 1)],
    (700, 6):  [("Skaven", 1), ("Elven Union", 2)],
    (700, 7):  [("Khorne Renegades", 3), ("Chaos Dwarf", 2)],
    (700, 8): [],
    (725, 3):  [],
    (725, 5):  [],
    (725, 6):  [],
    (725, 7):  [("Dwarf", 1), ("Goblins", 4), ("Black Orc", 3)],
    (725, 8): [],
    (750, 3):  [],
    (750, 5):  [],
    (750, 6):  [],
    (750, 7):  [("Gnomes", 4)],
    (750, 8): [("Halfling", 4), ("Ogres", 3)],
}

# ─── STAR PLAYERS ───────────────────────────────────────────────────────────

# Keyed by SPP cost.  Add/remove stars here.
STAR_PLAYERS: dict[int, list[str]] = {
    1: [
        "Akhorne", "Barik Farblast", "Rashnak Backstabber", "Scrull Halfheight",
        "Josef Bugman", "The White Dwarf", "Willow Rosebark", "Rowana Forestfoot",
        "Frank n' Stein", "Glotl Stop", "Puggy Baconbreath", "Gretchen Wachter",
        "The Black Gobbo",
    ],
    2: [
        "Fungus the Loon", "Bilerot Vomitflesh", "Grim Ironjaw", "Grashnak Blackhoof",
        "Guffle Pushmaw", "Helmut Wulf", "Kiroth Krakeneye", "Max Spleenripper",
        "Scrappa Sorehead", "Glimmershard", "Thorson Stoutmead", "Zharg Madeye",
        "Rodney Roachbait", "Rumbelow Sheepskin", "Boa Kon'ssstriktr", "Bryce the Slice",
    ],
    3: [
        "Karla von Kill", "Ivar Eriksson", "Karina von Riesz", "Glart Smashrip",
        "Gloriel Summerbloom", "Ivan the Animal", "Morg n'Thighgrove", "Mighty Zug",
        "Scyla Anfingrimm", "The Swift Twins", "Wilhelm Chaney", "Zolcath the Zoat",
        "Angi Pangi", "Withergrasp", "Nobbla Blackwart", "Cindy Piewhistle",
        "Bomber Dribblesnot",
    ],
    5: [
        "H'Thark", "Deeproot", "Kreek Rustgouger", "Estelle la Veneaux",
        "Ripper Bolgrot", "Varag Ghoul-Chewer", "Count Luthor", "Lord Borak",
        "Dribb & Drill", "Skorg Snowpelt", "Eldril Sidewinder",
    ],
    8: [
        "Griff Oberwald", "Hakflem", "Skitter Stab Stab", "Grak & Crumbleberry",
        "Jordell Freshbreeze", "Roxanna Darknail", "Jeremiah Kool",
    ],
}

# ─── SPECIAL RULES ──────────────────────────────────────────────────────────

# Edit items/table in each section.
# Sections support three optional keys:
#   intro : str            – plain text shown above bullets/table
#   items : list[str]      – bullet-point lines
#   table : dict           – rendered as a data table; keys:
#       headers    : list[str]   – column header labels (optional)
#       rows       : list[list]  – one inner list per row
#       col_widths : list[float] – relative column widths (optional)
SPECIAL_RULES: dict[str, dict] = {
    "General": {
        "intro": None,
        "items": [
            "Official rules apply unless stated below",
            "Reroll double cost",
            "Allowed positionals: 3 + TIER",
            "Allowed rerolls: 1 + TIER",
            "Apo cost 75.000 and works as normal",
            "Positionals: All non 0-16 linemen!",
        ],
    },
    "Skill Picks & Stacking": {
        "intro": None,
        "items": [
            "Players can be added 1 Skill",
            "Primary = 1 SPP,  Secondary = 2 SPP",
            "Add max 3 Elite Skills to Players in total",
            "Access to 2 Primary Skill Stacks = 1 SPP",
            "Remove Elite Skill cap = 1 SPP",
        ],
    },
    "Allowed Inducements": {
        "intro": None,
        "items": ["Lone Fouler on team = No Bribes"],
        "table": {
            "headers": ["Qty", "Name", "Cost"],
            "col_widths": [0.10, 0.58, 0.32],
            "rows": [
                ["0-6",   "Extra team Training",      "100.000 GP"],
                ["0-3/6", "Bribes",                   "100.000 / 50.000 GP"],
                ["0-5",   "Temp Agency Cheerleaders", "15.000 GP"],
                ["0-5",   "Part-time Asst. Coaches",  "25.000 GP"],
                ["0-1",   "Riotous Rookies",           "120.000 GP"],
                ["0-1",   "Mortuary Assistant",        "70.000 GP"],
                ["0-1",   "Plague Doctor",             "70.000 GP"],
                ["0-2",   "Blitzers Best Kegs",        "30.000 GP"],
                ["0-1",   "Halfling Master Chef",      "70.000 / 200.000 GP"],
                ["0-2",   "Experimental Drugs",       "50.000 GP"],
                ["0-1",   "Star Players",             "Prices vary"],
            ],
        },
    },
    "Experimental Drugs": {
        "intro": "Roll 2d6 before each match. Select player(s) and roll – same player can be chosen multiple times.",
        "items": [],
        "table": {
            "headers": ["2d6", "Effect"],
            "col_widths": [0.18, 0.82],
            "rows": [
                ["2",     "Random Secondary"],
                ["3-4",   "Random Primary"],
                ["5-9",   "Selected Primary"],
                ["10-11", "Selected Secondary"],
                ["12",    "Random Stat / Selected Secondary"],
            ],
        },
    },
    "Stars": {
        "intro": None,
        "items": [
            "7 Players before purchase",
            "Use Team GP to purchase Stars",
            "Stars also cost Team SPP",
            "If teams share the same Star, both play",
        ],
    },
}

# ─── COLOURS ────────────────────────────────────────────────────────────────

BLUE_DARK    = RGBColor(0x0A, 0x18, 0x3C)   # slide background / darkest cells
BLUE_MID     = RGBColor(0x1A, 0x3A, 0x7C)   # normal grid cells
BLUE_LIGHT   = RGBColor(0x23, 0x50, 0xA0)   # alternate rows / star section
HEADER_BG    = RGBColor(0x05, 0x0F, 0x26)   # column/row header cells
SECTION_BG   = RGBColor(0x10, 0x28, 0x60)   # rule section header bg
GOLD         = RGBColor(0xFF, 0xC0, 0x00)   # header text / highlights
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)   # body text
STAR_GOLD    = RGBColor(0xFF, 0xD7, 0x00)   # star icon

# Tier badge colours (displayed inline after each team name)
TIER_COLORS = {
    1: RGBColor(0x00, 0xCC, 0x44),   # T1 – green
    2: RGBColor(0xFF, 0xDD, 0x00),   # T2 – yellow
    3: RGBColor(0xFF, 0x77, 0x00),   # T3 – orange
    4: RGBColor(0xDD, 0x11, 0x11),   # T4 – red
}

# ─── HELPERS ────────────────────────────────────────────────────────────────

def _hex(color: RGBColor) -> str:
    return str(color)  # RGBColor.__str__ returns the 6-char hex string


def fill_cell(cell, color: RGBColor):
    """Set table cell solid fill colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("a:solidFill")):
        tcPr.remove(old)
    sf = etree.SubElement(tcPr, qn("a:solidFill"))
    clr = etree.SubElement(sf, qn("a:srgbClr"))
    clr.set("val", _hex(color))


def cell_margins(cell, top=Pt(2), bottom=Pt(2), left=Pt(3), right=Pt(3)):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.set("marT", str(int(top)))
    tcPr.set("marB", str(int(bottom)))
    tcPr.set("marL", str(int(left)))
    tcPr.set("marR", str(int(right)))


def write_cell(
    cell,
    bg: RGBColor,
    lines: list[str],
    font_size: float = 9,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = WHITE,
    align=PP_ALIGN.CENTER,
    v_anchor=None,
):
    """Fill a cell, clear it, then write one paragraph per line."""
    fill_cell(cell, bg)
    cell_margins(cell)
    tf = cell.text_frame
    tf.word_wrap = True
    if v_anchor is not None:
        tf.vertical_anchor = v_anchor

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color


def write_teams_cell(cell, teams: list[tuple[str, int]], bg: RGBColor):
    """
    Write team entries into a grid cell.
    Each entry is (name, tier).  The tier is shown as a small coloured badge
    (T1/T2/T3) on the same line as the team name.
    """
    fill_cell(cell, bg)
    cell_margins(cell)
    tf = cell.text_frame
    tf.word_wrap = True

    for i, (name, tier) in enumerate(teams):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(0)
        p.space_after = Pt(1)

        r_name = p.add_run()
        r_name.text = name + " "
        r_name.font.size = Pt(9.5)
        r_name.font.color.rgb = WHITE

        r_tier = p.add_run()
        r_tier.text = f"T{tier}"
        r_tier.font.size = Pt(7)
        r_tier.font.bold = True
        r_tier.font.color.rgb = TIER_COLORS.get(tier, GOLD)


def _remove_border(cell):
    """Remove all borders on a cell (no-border look)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(side)):
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(side))
        ln.set("w", "0")
        nofill = etree.SubElement(ln, qn("a:noFill"))  # noqa: F841


def set_thin_border(cell, color: RGBColor = GOLD, width_pt: float = 0.5):
    """Set all four borders to a thin line."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    w = str(int(Pt(width_pt)))
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for old in tcPr.findall(qn(side)):
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(side))
        ln.set("w", w)
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        c = etree.SubElement(sf, qn("a:srgbClr"))
        c.set("val", _hex(color))


def add_textbox(
    slide,
    text: str,
    left, top, width, height,
    font_size: float = 12,
    bold: bool = False,
    color: RGBColor = WHITE,
    bg: RGBColor | None = None,
    align=PP_ALIGN.LEFT,
):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    if bg:
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = bg
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def send_to_back(slide, shape):
    sp_tree = slide.shapes._spTree
    elem = shape._element
    sp_tree.remove(elem)
    sp_tree.insert(2, elem)  # index 2 = behind all content shapes


# ─── BACKGROUND GENERATION ──────────────────────────────────────────────────

def build_logo_png(svg_path: str, out_path: str, height_px: int = 80):
    """Render the SVG logo as a small white-tinted PNG for the corner badge."""
    logo_bytes = cairosvg.svg2png(url=svg_path, output_height=height_px)
    logo = Image.open(io.BytesIO(logo_bytes)).convert("RGBA")
    r, g, b, a = logo.split()
    white = Image.new("L", logo.size, 255)
    bright_a = a.point(lambda v: int(v * 0.85))
    white_logo = Image.merge("RGBA", (white, white, white, bright_a))
    canvas = Image.new("RGBA", logo.size, (0, 0, 0, 0))
    canvas.paste(white_logo, (0, 0), white_logo)
    canvas.save(out_path, "PNG")
    return logo.width, logo.height


def build_background_png(svg_path: str, out_path: str, width: int = 1920, height: int = 1080):
    """
    Render the SVG logo as a white watermark centred on a dark-navy background.
    The resulting PNG is used as the slide background on every page.
    """
    # Convert SVG → RGBA PNG
    logo_bytes = cairosvg.svg2png(url=svg_path, output_height=int(height * 0.85))
    logo = Image.open(io.BytesIO(logo_bytes)).convert("RGBA")

    # Recolour: original paths are black; tint them white at ~20 % opacity
    r_logo, g_logo, b_logo, a_logo = logo.split()
    white_r = Image.new("L", logo.size, 255)
    white_g = Image.new("L", logo.size, 255)
    white_b = Image.new("L", logo.size, 255)
    # 20 % opacity watermark
    faded_a = a_logo.point(lambda v: int(v * 0.20))
    white_logo = Image.merge("RGBA", (white_r, white_g, white_b, faded_a))

    # Navy-blue background
    bg = Image.new("RGBA", (width, height), (10, 24, 60, 255))

    # Centre the logo
    x = (width - logo.width) // 2
    y = (height - logo.height) // 2
    bg.paste(white_logo, (x, y), white_logo)

    bg.convert("RGB").save(out_path, "PNG")
    return out_path


def add_bg_picture(slide, img_path: str, prs):
    """Insert image as the very first element (background)."""
    pic = slide.shapes.add_picture(img_path, 0, 0, prs.slide_width, prs.slide_height)
    send_to_back(slide, pic)


def add_logo_corner(slide, prs, margin):
    """Place the logo PNG in the upper-right corner, centred within the title bar."""
    with Image.open(LOGO_PATH) as img:
        w_px, h_px = img.size
    h = Inches(0.38)
    w = int(h * w_px / h_px)
    left = prs.slide_width - margin - w
    top  = margin + int((Inches(0.45) - h) / 2)
    slide.shapes.add_picture(LOGO_PATH, left, top, w, int(h))


# ─── PAGE 1 – GRID ──────────────────────────────────────────────────────────

def build_grid_table(slide, left, top, width, height):
    """
    Build the main team-grid table.
    Rows = GP costs, Columns = SPP budgets.
    """
    n_rows = 1 + len(GP_ROWS)       # header + data rows
    n_cols = 1 + len(SPP_COLS)      # GP label col + SPP cols

    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table

    # Column widths  (GP label narrower, SPP cols equal)
    gp_col_w = int(width * 0.095)
    spp_col_w = int((width - gp_col_w) / len(SPP_COLS))
    tbl.columns[0].width = gp_col_w
    for c in range(1, n_cols):
        tbl.columns[c].width = spp_col_w

    # Row heights (header slightly shorter)
    hdr_h = int(height * 0.13)
    data_h = int((height - hdr_h) / len(GP_ROWS))
    tbl.rows[0].height = hdr_h
    for r in range(1, n_rows):
        tbl.rows[r].height = data_h

    # ── header row ──
    write_cell(tbl.cell(0, 0), HEADER_BG, ["GP / SPP"], 9, bold=True, color=GOLD)
    for ci, spp in enumerate(SPP_COLS):
        write_cell(tbl.cell(0, ci + 1), HEADER_BG, [str(spp)], 11, bold=True, color=GOLD)

    # ── data rows ──
    for ri, gp in enumerate(GP_ROWS):
        row_idx = ri + 1
        row_bg = BLUE_MID if ri % 2 == 0 else BLUE_LIGHT
        # GP label
        write_cell(tbl.cell(row_idx, 0), HEADER_BG, [str(gp)], 9, bold=True, color=GOLD)
        # Team cells
        for ci, spp in enumerate(SPP_COLS):
            teams = TEAMS.get((gp, spp), [])
            write_teams_cell(tbl.cell(row_idx, ci + 1), teams, row_bg)

    # Borders
    for r in range(n_rows):
        for c in range(n_cols):
            set_thin_border(tbl.cell(r, c), GOLD, 0.4)

    return tbl


def build_star_table(slide, left, top, width, height):
    """
    Build the star-players reference table.
    One column per SPP cost bucket, rows = individual star names.
    """
    spp_costs = sorted(STAR_PLAYERS.keys())
    n_cols = 1 + len(spp_costs)   # label col + cost cols
    n_rows = 2                     # header + single data row (multi-line text)

    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table

    label_w = int(width * 0.075)
    cost_w  = int((width - label_w) / len(spp_costs))
    tbl.columns[0].width = label_w
    for c in range(1, n_cols):
        tbl.columns[c].width = cost_w

    hdr_h = int(height * 0.18)
    tbl.rows[0].height = int(hdr_h)
    tbl.rows[1].height = int(height - hdr_h)

    # Header
    write_cell(tbl.cell(0, 0), HEADER_BG, ["★", "Cost"], 8, bold=True, color=STAR_GOLD)
    for ci, cost in enumerate(spp_costs):
        write_cell(tbl.cell(0, ci + 1), HEADER_BG, [f"{cost} SPP"], 9, bold=True, color=GOLD)

    # Star names – one block per cost column
    write_cell(tbl.cell(1, 0), BLUE_DARK, [], 7)
    for ci, cost in enumerate(spp_costs):
        names = STAR_PLAYERS[cost]
        write_cell(tbl.cell(1, ci + 1), BLUE_MID, names, 7, color=WHITE, align=PP_ALIGN.LEFT)

    for r in range(n_rows):
        for c in range(n_cols):
            set_thin_border(tbl.cell(r, c), GOLD, 0.4)

    return tbl


def build_page1(prs, bg_path: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    add_bg_picture(slide, bg_path, prs)

    sw = prs.slide_width
    sh = prs.slide_height
    margin = Inches(0.12)

    # Title bar
    title_h = Inches(0.45)
    add_textbox(
        slide, "Aarhus Wargaming Grid 2026",
        margin, margin, sw - 2 * margin, title_h,
        font_size=22, bold=True, color=GOLD,
        bg=BLUE_DARK, align=PP_ALIGN.CENTER,
    )

    # Grid table
    grid_top  = margin + title_h + Inches(0.08)
    grid_h    = Inches(3.75)
    build_grid_table(slide, margin, grid_top, sw - 2 * margin, grid_h)

    # Star players table
    star_top = grid_top + grid_h + Inches(0.08)
    star_h   = sh - star_top - margin
    build_star_table(slide, margin, star_top, sw - 2 * margin, star_h)

    add_logo_corner(slide, prs, margin)


# ─── PAGE 2 – RULES ─────────────────────────────────────────────────────────

def _build_inline_table(slide, tbl_def: dict, left, top, width, height):
    """Render a compact data table inside a rules section."""
    headers    = tbl_def.get("headers", [])
    rows       = tbl_def["rows"]
    col_widths = tbl_def.get("col_widths")
    n_cols     = len(headers) if headers else len(rows[0])
    has_hdr    = bool(headers)
    n_rows     = len(rows) + (1 if has_hdr else 0)

    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, int(width), int(height)).table

    # Column widths
    if col_widths:
        total = sum(col_widths)
        for c, w in enumerate(col_widths):
            tbl.columns[c].width = int(width * w / total)
    else:
        cw = int(width / n_cols)
        for c in range(n_cols):
            tbl.columns[c].width = cw

    # Row heights
    HDR_ROW_H = int(Inches(0.23))
    DATA_ROW_H = int(Inches(0.195))
    row_offset = 0
    if has_hdr:
        tbl.rows[0].height = HDR_ROW_H
        row_offset = 1
    for r in range(len(rows)):
        tbl.rows[r + row_offset].height = DATA_ROW_H

    # Header row
    if has_hdr:
        for ci, hdr in enumerate(headers):
            write_cell(tbl.cell(0, ci), SECTION_BG, [hdr], 7.5, bold=True, color=GOLD)
            set_thin_border(tbl.cell(0, ci), GOLD, 0.3)

    # Data rows – alternate BLUE_DARK / BLUE_MID; first column centred
    for ri, row_data in enumerate(rows):
        row_bg = BLUE_DARK if ri % 2 == 0 else BLUE_MID
        for ci, val in enumerate(row_data):
            align = PP_ALIGN.CENTER if ci == 0 else PP_ALIGN.LEFT
            write_cell(tbl.cell(ri + row_offset, ci), row_bg, [str(val)], 7.5,
                       color=WHITE, align=align)
            set_thin_border(tbl.cell(ri + row_offset, ci), GOLD, 0.3)


def build_rules_block(slide, section: str, rules: dict, left, top, width, max_h):
    """
    Draw a rule section: coloured header bar, optional bullet list, optional table.
    Sections can have any combination of 'intro', 'items', and 'table' keys.
    Returns the bottom edge Y position.
    """
    HDR_H     = Inches(0.32)
    LINE_H    = Inches(0.205)
    ROW_H     = Inches(0.195)
    HDR_ROW_H = Inches(0.23)
    PAD       = Inches(0.05)

    items   = rules.get("items", [])
    intro   = rules.get("intro")
    tbl_def = rules.get("table")

    # Estimate text lines (intro may wrap)
    n_intro = math.ceil(len(intro) / 80) if intro else 0
    n_text  = n_intro + len(items)

    text_h = (PAD + n_text * LINE_H + PAD) if n_text > 0 else Inches(0)

    if tbl_def:
        n_data = len(tbl_def["rows"])
        n_hdr  = 1 if tbl_def.get("headers") else 0
        tbl_h  = n_hdr * HDR_ROW_H + n_data * ROW_H + PAD
        gap    = PAD if n_text > 0 else Inches(0)
    else:
        tbl_h = Inches(0)
        gap   = Inches(0)

    # ── section header bar ──
    add_textbox(
        slide, section,
        left, top, width, HDR_H,
        font_size=11, bold=True, color=GOLD,
        bg=SECTION_BG, align=PP_ALIGN.CENTER,
    )

    y = top + HDR_H

    # ── bullet / intro text ──
    if n_text > 0:
        body_box = slide.shapes.add_textbox(left, y, width, text_h)
        fill = body_box.fill
        fill.solid()
        fill.fore_color.rgb = BLUE_LIGHT

        tf = body_box.text_frame
        tf.word_wrap = True

        def _para(text: str, bullet: bool, first: bool):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.space_before = Pt(1)
            p.space_after  = Pt(0)
            run = p.add_run()
            run.text = ("• " if bullet else "  ") + text
            run.font.size = Pt(8.5)
            run.font.color.rgb = WHITE

        first = True
        if intro:
            _para(intro, bullet=False, first=first)
            first = False
        for item in items:
            _para(item, bullet=True, first=first)
            first = False

        y += text_h + gap

    # ── inline table ──
    if tbl_def:
        _build_inline_table(slide, tbl_def, left, y, width, tbl_h)
        y += tbl_h

    return y + PAD


def build_star_rules_table(slide, left, top, width, height):
    """
    Re-use the same star players layout on the rules page for reference.
    """
    build_star_table(slide, left, top, width, height)


def build_page2(prs, bg_path: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg_picture(slide, bg_path, prs)

    sw = prs.slide_width
    sh = prs.slide_height
    margin = Inches(0.12)
    gap    = Inches(0.10)

    # Title
    title_h = Inches(0.45)
    add_textbox(
        slide, "Aarhus Wargaming Grid 2026 – Special Rules",
        margin, margin, sw - 2 * margin, title_h,
        font_size=20, bold=True, color=GOLD,
        bg=BLUE_DARK, align=PP_ALIGN.CENTER,
    )

    # Two-column layout for rule sections
    rules_top  = margin + title_h + gap
    col_w      = (sw - 2 * margin - gap) / 2
    col_left   = [margin, margin + col_w + gap]

    sections   = list(SPECIAL_RULES.items())
    mid        = (len(sections) + 1) // 2   # split into two halves
    left_secs  = sections[:mid]
    right_secs = sections[mid:]

    rules_bottom = rules_top

    for col_idx, col_secs in enumerate((left_secs, right_secs)):
        y = rules_top
        for section, rules in col_secs:
            bottom = build_rules_block(
                slide, section, rules,
                col_left[col_idx], y, col_w,
                sh - y - margin,
            )
            y = bottom + gap
        rules_bottom = max(rules_bottom, y)

    # Star players at the bottom
    stars_top = rules_bottom + gap * 0.5
    stars_h   = sh - stars_top - margin
    if stars_h > Inches(1.5):
        build_star_rules_table(slide, margin, stars_top, sw - 2 * margin, stars_h)

    add_logo_corner(slide, prs, margin)


# ─── PDF EXPORT ─────────────────────────────────────────────────────────────

def export_pdf(pptx_path: str) -> str:
    """Convert the saved PPTX to PDF using LibreOffice (headless)."""
    out_dir = os.path.dirname(pptx_path) or "."
    result = subprocess.run(
        [
            "libreoffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            pptx_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    pdf_path = os.path.splitext(pptx_path)[0] + ".pdf"
    if result.returncode != 0 or not os.path.exists(pdf_path):
        raise RuntimeError(result.stderr.strip() or "LibreOffice returned non-zero exit code")
    return pdf_path


# ─── MAIN ───────────────────────────────────────────────────────────────────

def main():
    print("Building background image…")
    build_background_png(SVG_PATH, BG_PATH)
    build_logo_png(SVG_PATH, LOGO_PATH)

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    print("Building page 1 (grid)…")
    build_page1(prs, BG_PATH)

    print("Building page 2 (rules)…")
    build_page2(prs, BG_PATH)

    prs.save(OUT_PATH)
    print(f"Saved → {OUT_PATH}")

    print("Exporting PDF…")
    try:
        pdf_path = export_pdf(OUT_PATH)
        print(f"Saved → {pdf_path}")
    except (RuntimeError, FileNotFoundError) as e:
        print(f"Warning: PDF export failed – {e}")

    # Clean up temp files
    for tmp in (BG_PATH, LOGO_PATH):
        try:
            os.remove(tmp)
        except OSError:
            pass


if __name__ == "__main__":
    main()
