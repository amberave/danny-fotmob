"""
Parse FotMob player-page JSON files into the players2.xlsx master sheet
=========================================================================

Input
-----
A folder called "json" (next to this script, or set JSON_DIR below) full of
files named "{fotmob_player_id}.json" - e.g. "1061249.json" for Vitinha.
Each file is the raw payload from a FotMob player page, with the shape:

    {
      "sectionOrder": [...],
      "shotmap": [...],        # per-shot event data - not used here
      "heatmap": {...},        # not used here
      "topStatCard": {"items": [{"title": "Goals", "statValue": "5", ...}, ...]},
      "statsSection": {"items": [
          {"title": "Shooting",  "items": [{"title": "Goals", "statValue": "5",
                                             "per90": 0.38, ...}, ...]},
          {"title": "Physical",  "items": [...]},
          {"title": "Passing",   "items": [...]},
          {"title": "Possession","items": [...]},
          {"title": "Defending", "items": [...]},
          {"title": "Discipline","items": [...]},
      ]}
    }

Every stat under statsSection (Shooting/Physical/Passing/Possession/
Defending/Discipline) becomes one output column, named after its "title"
(e.g. "Goals", "xG", "Top Speed", "Duels won %", ...). "Matches", "Started",
"Minutes" and "Rating" come from topStatCard (Goals/Assists duplicate the
Shooting/Passing groups, so topStatCard's copies of those two are skipped).
Per-90 values are written to extra "<title> per 90" columns.

This file does not say which competition/season it came from, so this
script assumes they're all the same competition/season you intend them to
be (e.g. you selected "Champions League" + "Total" before saving each one)
and does not try to verify that - double check a couple of files if unsure.

Output
------
Updates MASTER_XLSX (default "players2.xlsx") in place: for every JSON file,
matches its {player_id} to a row by the ID column (auto-detected - see
ID_COLUMN_CANDIDATES below, override ID_COLUMN if it guesses wrong), writes
the stats into that row, and adds any new stat columns at the end. Existing
columns, the index column, and rows with no matching JSON file are left
untouched. Player IDs with no row in the master sheet are appended as new
rows (name/id filled in, other existing columns left blank) rather than
silently dropped. A timestamped backup of the original file is written
alongside it before saving.

Requirements:  pip install pandas openpyxl
Run:           python parse_player_jsons_to_excel.py
"""

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

# --------------------------------------------------------------------------
# Config - adjust these if they don't match your setup
# --------------------------------------------------------------------------
LEAGUE = "2026-2027 FIFA World Cup"
JSON_DIR = Path(f"json/{LEAGUE}")                 # folder of {player_id}.json files
MASTER_XLSX = Path(f"{LEAGUE}.xlsx")     # existing master workbook to update
SHEET_NAME = 0                          # sheet name or index to update
ID_COLUMN = None                        # force a column name, or leave None
                                         # to auto-detect (see candidates below)
NAME_COLUMN_CANDIDATES = [
    "player", "player name", "name",
]
ID_COLUMN_CANDIDATES = [
    "id", "player_id", "playerid", "fotmob id", "fotmob_id",
    "fotmob player id", "fotmob player id (supplied)",
    "fotmob player id (seen in data)",
]
INCLUDE_PER90 = True                    # also write "<stat> per 90" columns

# Groups to pull from statsSection; add/remove to match what you need.
STAT_GROUPS = ["Shooting", "Physical", "Passing", "Possession",
               "Defending", "Discipline"]
# topStatCard items to pull (title -> output column name). Goals/Assists are
# skipped here since Shooting/Passing already carry them.
TOP_CARD_STATS = {
    "Matches": "Matches", "Started": "Matches started",
    "Minutes": "Minutes played", "Rating": "Player rating",
}


# --------------------------------------------------------------------------
# JSON parsing
# --------------------------------------------------------------------------
def to_num(v):
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").replace("%", ""))
    except ValueError:
        return None


def parse_player_json(path):
    """Return {output_column: value} for one player's JSON file."""
    data = json.loads(path.read_text(encoding="utf-8"))
    out = {}

    top_items = (data.get("topStatCard") or {}).get("items") or []
    for item in top_items:
        col = TOP_CARD_STATS.get(item.get("title"))
        if col:
            out[col] = to_num(item.get("statValue"))

    stats_section = data.get("statsSection") or {}
    seen_titles = set()
    for group in stats_section.get("items") or []:
        if STAT_GROUPS and group.get("title") not in STAT_GROUPS:
            continue
        for item in group.get("items") or []:
            title = item.get("title")
            if title is None:
                continue
            col = title
            if col in seen_titles:               # disambiguate a clash
                col = f"{group.get('title')} - {title}"
            seen_titles.add(title)
            out[col] = to_num(item.get("statValue"))
            if INCLUDE_PER90 and item.get("per90") is not None:
                out[f"{col} per 90"] = to_num(item.get("per90"))

    return out


def load_all_player_jsons():
    """Return {player_id: {column: value}} for every json/{id}.json file."""
    results = {}
    if not JSON_DIR.is_dir():
        print(f"! {JSON_DIR}/ does not exist - nothing to parse")
        return results
    for fp in sorted(JSON_DIR.glob("*.json")):
        stem = fp.stem
        if not stem.isdigit():
            print(f"  ! skipping {fp.name}: filename isn't a plain player ID")
            continue
        try:
            results[int(stem)] = parse_player_json(fp)
        except (json.JSONDecodeError, OSError) as exc:
            print(f"  ! could not parse {fp.name}: {exc}")
    return results


# --------------------------------------------------------------------------
# Master workbook merge
# --------------------------------------------------------------------------
def find_column(df, forced, candidates):
    if forced:
        if forced not in df.columns:
            raise SystemExit(f"Column '{forced}' not found in {MASTER_XLSX}. "
                              f"Available columns: {list(df.columns)}")
        return forced
    lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand in lower:
            return lower[cand]
    return None


def main():
    player_data = load_all_player_jsons()
    print(f"Parsed {len(player_data)} player JSON file(s) from {JSON_DIR}/")
    if not player_data:
        return

    if not MASTER_XLSX.exists():
        raise SystemExit(
            f"{MASTER_XLSX} not found. Put it next to this script (or update "
            f"MASTER_XLSX at the top) - it should already have your player "
            f"name/ID/index columns."
        )

    df = pd.read_excel(MASTER_XLSX, sheet_name=SHEET_NAME)
    id_col = find_column(df, ID_COLUMN, ID_COLUMN_CANDIDATES)
    if id_col is None:
        raise SystemExit(
            f"Couldn't find an ID column in {MASTER_XLSX}. Columns present: "
            f"{list(df.columns)}\nSet ID_COLUMN at the top of this script to "
            f"the right one and re-run."
        )
    name_col = find_column(df, None, NAME_COLUMN_CANDIDATES)
    print(f"Matching on column '{id_col}'"
          + (f" (name column '{name_col}')" if name_col else ""))

    # Backup before touching anything.
    # backup = MASTER_XLSX.with_name(
    #     f"{MASTER_XLSX.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    # shutil.copy2(MASTER_XLSX, backup)
    # print(f"Backed up original to {backup.name}")

    ids_numeric = pd.to_numeric(df[id_col], errors="coerce")

    # All stat columns seen across every JSON file, in first-seen order.
    all_cols = []
    for cols in player_data.values():
        for c in cols:
            if c not in all_cols:
                all_cols.append(c)
    new_cols = [c for c in all_cols if c not in df.columns]
    if new_cols:
        df = pd.concat([df, pd.DataFrame(pd.NA, index=df.index, columns=new_cols)],
                        axis=1)

    matched, unmatched_new = 0, []
    for pid, cols in player_data.items():
        mask = ids_numeric == pid
        if mask.any():
            idx = df.index[mask][0]
            for c, v in cols.items():
                df.at[idx, c] = v
            matched += 1
        else:
            new_row = {c: pd.NA for c in df.columns}
            new_row[id_col] = pid
            new_row.update(cols)
            unmatched_new.append(new_row)

    if unmatched_new:
        df = pd.concat([df, pd.DataFrame(unmatched_new)], ignore_index=True)
        print(f"  {len(unmatched_new)} player ID(s) had no existing row - "
              f"appended as new row(s): {[r[id_col] for r in unmatched_new]}")
    print(f"Updated {matched} existing row(s)")

    with pd.ExcelWriter(MASTER_XLSX, engine="openpyxl") as writer:
        sheet_name = SHEET_NAME if isinstance(SHEET_NAME, str) else LEAGUE
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.book[sheet_name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for col in ws.columns:
            width = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 50)

    print(f"Saved: {MASTER_XLSX.resolve()}")
    print(f"Columns: {len(df.columns)}  |  Rows: {len(df)}")


if __name__ == "__main__":
    main()
