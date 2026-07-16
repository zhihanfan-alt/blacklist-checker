# Blacklist Checker / 分拣员工黑名单检查器

**[Use Online / 在线使用](https://zhihanfan-alt.github.io/blacklist-checker/)** — no installation needed, runs in your browser

A Python CLI tool that checks scheduling spreadsheets against an employee blacklist using **6-layer name matching**. Supports configurable column positions for different spreadsheet formats.

## Features

- **6-Layer Name Matching**: exact, reversed word order, no-space, token subset, fuzzy (rapidfuzz), and n-gram similarity
- **Configurable Columns**: customize which columns to scan via JSON config file
- **Selective Export**: only outputs files that have matches
- **Visual Highlighting**: matched cells are highlighted in red with a summary row
- **Batch Processing**: check multiple schedule files at once
- **Multi-Sheet Support**: scans all worksheets in each Excel file
- **JSON Report**: detailed machine-readable output for integration

## Installation

```bash
# Clone the repo
git clone https://github.com/zhihanfan-alt/blacklist-checker.git
cd blacklist-checker

# Install dependencies
pip install -r requirements.txt
```

## Quick Start

### Basic Usage

```bash
# Provide files as arguments (first file = blacklist, rest = schedules)
python checker.py blacklist.xlsx schedule_mon.xlsx schedule_tue.xlsx

# Or place files in the current directory (auto-detect blacklist by filename)
python checker.py
```

### Blacklist File Format

Create an Excel file (`.xlsx`) with:
- A sheet named `黑名单` (or `Blacklist`, or any name with `--blacklist-sheet`)
- Column A containing employee names, starting from row 2 (row 1 = header)

| A |
|---|
| Name |
| John Smith |
| Maria Garcia |

### Schedule File Format

The tool scans specific columns in your schedule files for employee names. By default, it checks columns 1, 4, 8, 11, 15, 19, 23, 28 — but this is fully configurable (see below).

## Configuration

Create a JSON config file to customize the tool's behavior:

```bash
python checker.py --config my_config.json blacklist.xlsx schedule.xlsx
```

### Config File Format

See `config.example.json` for a full example:

```json
{
  "name_columns": [1, 4, 8],
  "sections": {
    "1": "Scanner",
    "4": "Tables",
    "8": "Name List"
  },
  "fuzzy_threshold": 85,
  "ngram_threshold": 0.70,
  "blacklist_sheet": "Blacklist"
}
```

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `name_columns` | `int[]` | `[1,4,8,11,15,19,23,28]` | Column numbers (1-based) to scan for names |
| `sections` | `object` | (see code) | Label for each column number |
| `fuzzy_threshold` | `int` | `85` | Minimum fuzzy match score (0-100) |
| `ngram_threshold` | `float` | `0.70` | Minimum n-gram similarity (0-1) |
| `blacklist_sheet` | `string` | `"黑名单"` | Sheet name in the blacklist file |

## CLI Options

```
python checker.py [files...] [options]

Positional:
  files                 Excel files (first = blacklist, rest = schedules)

Options:
  --version             Show version
  --config PATH         Path to JSON config file
  --blacklist-sheet N   Sheet name in the blacklist file
  --output-dir DIR      Output directory (default: output/)
  --min-score N         Minimum match score to include (0-100)
```

### Examples

```bash
# Use a custom config
python checker.py --config config.json blacklist.xlsx schedule.xlsx

# Set minimum match score to 90 (exclude low-confidence matches)
python checker.py --min-score 90 blacklist.xlsx schedule.xlsx

# Specify blacklist sheet name and output directory
python checker.py --blacklist-sheet "Blocklist" --output-dir results/ blacklist.xlsx schedule.xlsx
```

## Matching Algorithm

The tool uses a 6-layer matching strategy, from most to least precise:

| Layer | Method | Score | Description |
|-------|--------|-------|-------------|
| 1 | Exact | 100 | Names are identical after normalization |
| 2 | Reversed | 99 | Same words in different order (e.g., "Smith John" vs "John Smith") |
| 3 | No-space | 98 | Same characters ignoring spaces |
| 4 | Token subset | 95 | One name contains all words of the other |
| 5 | Fuzzy | 85-100 | Uses `rapidfuzz` ratio with configurable threshold |
| 6 | N-gram | 80-89 | Bigram/trigram similarity with configurable threshold |

Normalization (applied before matching):
- Lowercase conversion
- Diacritics removal (via `unidecode`, e.g., "Jose" -> "jose")
- Special character removal
- Whitespace collapsing

## Output

- **Console**: step-by-step progress with match details
- **Excel files**: only files with matches are exported to `output/` with red-highlighted cells
- **JSON report**: printed to stdout after `===JSON_REPORT===` marker for programmatic consumption

## Integration

The `===JSON_REPORT===` marker in stdout makes it easy to integrate with other tools:

```bash
python checker.py blacklist.xlsx schedule.xlsx | sed -n '/===JSON_REPORT===/,$ p' | tail -n +2 | jq .
```

## License

MIT
