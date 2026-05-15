# Aarhus Wargaming Grid – Blood Bowl 7s

Generates a two-page **PowerPoint (.pptx)** and **PDF** reference sheet for the
Aarhus Wargaming Blood Bowl 7s tournament grid.

- **Page 1** – Team grid (GP rows × SPP columns) with tier badges, plus a star
  players table organised by SPP cost.
- **Page 2** – Special rules, inducements table, experimental-drugs table, and a
  repeat of the star players for quick reference.

The Warhammer Aarhus club logo appears both as a watermark background and as a
small badge in the upper-right corner of each page.

The grid was based on Steely Grid from 2026.

---

## Requirements

| Requirement | Notes |
|---|---|
| Python 3.10+ | Type hints use `dict[…]` and `list[…]` syntax |
| [python-pptx](https://python-pptx.readthedocs.io/) | PowerPoint generation |
| [cairosvg](https://cairosvg.org/) | SVG → PNG conversion for the background |
| [Pillow](https://pillow.readthedocs.io/) | Image compositing |
| [lxml](https://lxml.de/) | XML manipulation inside python-pptx |
| [reportlab](https://www.reportlab.com/) | Pure-Python PDF generation |

Install all Python packages in one step:

```bash
pip install python-pptx cairosvg pillow lxml reportlab
```

---

## Files

```
generate_AarhusWargmingGrid.py   ← main script (edit data here)
Warhammer Aarhus Logo_vector.svg ← club logo used as slide background
```

The script writes two output files next to itself:

```
AarhusWargmingGrid.pptx
AarhusWargmingGrid.pdf
```

---

## Usage

```bash
python3 generate_AarhusWargmingGrid.py
```

That's it. Both files are (re)created every time the script runs.

---

## Updating the content

All data lives at the top of `generate_AarhusWargmingGrid.py` in clearly
labelled sections. No other parts of the file need to be touched.

### Teams (`TEAMS`)

Each key is a `(gp_cost, spp_budget)` tuple. Each value is a list of
`(team_name, tier)` tuples where `tier` is `1`, `2`, or `3`.

```python
(700, 6): [("Skaven", 1), ("Elven Union", 2)],
```

- To **move** a team, change its key.
- To **add** a team, append `("Name", tier)` to the relevant cell.
- To **remove** a team, delete its tuple.
- Tier colours: **T1** = green, **T2** = yellow, **T3** = orange, **T4** = red.

### Star players (`STAR_PLAYERS`)

Keyed by SPP cost. Add or remove names from any list.

```python
5: ["H'Thark", "Deeproot", ...],
```

### Special rules (`SPECIAL_RULES`)

Each section is a dict with up to three optional keys:

| Key | Type | Purpose |
|---|---|---|
| `intro` | `str \| None` | Plain text shown above bullets or table |
| `items` | `list[str]` | Bullet-point lines |
| `table` | `dict` | Rendered as a data table (see below) |

**Adding a bullet section:**

```python
"My Section": {
    "intro": "Optional header sentence.",
    "items": [
        "Rule one",
        "Rule two",
    ],
},
```

**Adding a table section:**

```python
"My Table Section": {
    "intro": "Optional sentence above the table.",
    "table": {
        "headers": ["Col A", "Col B", "Col C"],   # omit for no header row
        "col_widths": [0.15, 0.55, 0.30],          # relative widths (optional)
        "rows": [
            ["val", "val", "val"],
            ["val", "val", "val"],
        ],
    },
},
```

A section can have both `items` **and** `table` – bullets appear first, table
below.

### Colours (`BLUE_DARK`, `BLUE_MID`, …)

Defined just below the rules data. Change the `RGBColor(r, g, b)` values to
retheme the whole sheet at once.

---

## License

Licensed under the [EUPL v1.2](LICENSE).
