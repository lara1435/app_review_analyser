"""Validate any spreadsheet (.xlsx / .csv) and analyze its rows with Jev.

Usage:
  uv run --with openpyxl jev_sheet.py FILE [options]

Examples:
  jev_sheet.py data.xlsx --validate-only
  jev_sheet.py data.xlsx --columns "Title,Review Text"
  jev_sheet.py data.xlsx --questions review_questions.json --output out.xlsx

Options:
  --sheet NAME          Sheet to use (default: first sheet)
  --columns A,B         Columns sent to Jev as the row's state (default: all text columns)
  --id-column NAME      Column used to label rows in logs (default: first column)
  --questions FILE      JSON file of Jev questions (default: generic sentiment/urgency/action)
  --output FILE         Output .xlsx (default: <input>_jev.xlsx; the input is never modified)
  --threshold 0.6       Confidence below this sets "Needs Human Review"
  --limit N             Analyze only the first N rows (for testing)
  --workers 4           Parallel Jev requests
  --validate-only       Run checks and stop; no API calls

Questions JSON format (same as the TypeSafe API):
  {
    "sentiment": {"type": "choice", "instructions": "...", "criteria": {"Positive": null, "Negative": null}},
    "urgency":   {"type": "score",  "instructions": "...", "criteria": ["Low", "Medium", "High"]},
    "is_bug":    {"type": "noul",   "instructions": "This row reports a bug."}
  }
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

load_dotenv()

DEFAULT_QUESTIONS = {
    "sentiment": {
        "type": "choice",
        "instructions": "What is the overall sentiment of this record?",
        "criteria": {"Positive": None, "Neutral": None, "Negative": None},
    },
    "urgency": {
        "type": "score",
        "instructions": "How urgently does someone need to act on this record?",
        "criteria": ["No action needed", "Low", "Medium", "High / critical"],
    },
    "needs_action": {
        "type": "noul",
        "instructions": "This record describes a problem or request that someone should act on.",
    },
}
MAX_STATE_CHARS = 20_000
HDR_FILL = PatternFill("solid", fgColor="DDEBF7")


# ---------------------------------------------------------------- loading
def load_table(path: Path, sheet: str | None):
    """Return (workbook, worksheet, headers, rows) — rows are lists of cell values (data rows only)."""
    if path.suffix.lower() == ".csv":
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        with path.open(newline="", encoding="utf-8-sig") as f:
            for r in csv.reader(f):
                ws.append([_num(v) for v in r])
    elif path.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(path)
        if sheet and sheet not in wb.sheetnames:
            raise SystemExit(f"ERROR: sheet '{sheet}' not found. Available: {wb.sheetnames}")
        ws = wb[sheet] if sheet else wb.worksheets[0]
    else:
        raise SystemExit(f"ERROR: unsupported file type '{path.suffix}' (use .xlsx, .xlsm or .csv)")
    all_rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not all_rows:
        raise SystemExit("ERROR: sheet is empty")
    headers = [str(h).strip() if h is not None else "" for h in all_rows[0]]
    return wb, ws, headers, all_rows[1:]


def _num(v: str):
    """CSV cells are strings; turn numeric ones into numbers, blanks into None."""
    if v.strip() == "":
        return None
    for cast in (int, float):
        try:
            return cast(v)
        except ValueError:
            pass
    return v


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


# ---------------------------------------------------------------- validation
def validate(headers, rows, columns, questions):
    errors, warnings, info = [], [], []
    width = len(headers)

    # header checks
    if all(h == "" for h in headers):
        errors.append("Header row is empty.")
    blanks = [get_column_letter(i + 1) for i, h in enumerate(headers) if h == ""]
    if blanks:
        warnings.append(f"Columns without a header: {', '.join(blanks)}")
    dups = [h for h, n in Counter(h for h in headers if h).items() if n > 1]
    if dups:
        errors.append(f"Duplicate column headers: {dups}")

    # row checks
    nonempty = [r for r in rows if not all(is_blank(v) for v in r)]
    empty = len(rows) - len(nonempty)
    # notes/footers: rows with a value only in the first cell
    notes = [i + 2 for i, r in enumerate(rows)
             if width > 1 and not is_blank(r[0]) and all(is_blank(v) for v in r[1:])]
    if notes:
        warnings.append(f"Rows with only the first cell filled (notes/footers?) — skipped: {notes}")
    data = [r for r in nonempty if not (width > 1 and all(is_blank(v) for v in r[1:]))]
    info.append(f"Rows: {len(data)} data, {empty} empty, {len(notes)} note")
    if not data:
        errors.append("No data rows.")

    dup_rows = sum(n - 1 for n in Counter(tuple(r) for r in data).values() if n > 1)
    if dup_rows:
        warnings.append(f"{dup_rows} fully duplicated row(s)")

    # per-column profile
    profile = []
    for i, h in enumerate(headers):
        vals = [r[i] if i < len(r) else None for r in data]
        filled = [v for v in vals if not is_blank(v)]
        types = Counter(
            "text" if isinstance(v, str) else
            "number" if isinstance(v, (int, float)) else
            "date" if isinstance(v, (datetime, date)) else type(v).__name__
            for v in filled
        )
        kind = types.most_common(1)[0][0] if types else "empty"
        mixed = len(types) > 1
        missing = len(vals) - len(filled)
        avg_len = sum(len(str(v)) for v in filled) / len(filled) if filled else 0
        profile.append((h or get_column_letter(i + 1), kind, len(filled), missing, len(set(map(str, filled))), mixed, avg_len))
        if mixed:
            warnings.append(f"Column '{h}' has mixed types: {dict(types)}")

    # columns to send
    if columns:
        unknown = [c for c in columns if c not in headers]
        if unknown:
            errors.append(f"--columns not found: {unknown}. Available: {headers}")
    else:
        text = [p for p in profile if p[1] == "text" and p[0] in headers]
        # prefer free-text columns (avg ≥ 10 chars); skips IDs, usernames, codes
        columns = [p[0] for p in text if p[6] >= 10] or [p[0] for p in text]
        if not columns:
            errors.append("No text columns found to analyze; pass --columns.")
        else:
            info.append(f"Auto-selected text columns: {columns}")

    # questions
    if not isinstance(questions, dict) or not questions:
        errors.append("Questions must be a non-empty JSON object.")
    else:
        for name, q in questions.items():
            t = q.get("type") if isinstance(q, dict) else None
            if t not in {"choice", "score", "noul"}:
                errors.append(f"Question '{name}': type must be choice, score or noul")
            elif t == "choice" and not (isinstance(q.get("criteria"), dict) and len(q["criteria"]) >= 2):
                errors.append(f"Question '{name}': choice needs a criteria object with ≥2 labels")
            elif t == "score" and not (isinstance(q.get("criteria"), list) and len(q["criteria"]) >= 2):
                errors.append(f"Question '{name}': score needs a criteria list with ≥2 levels")

    return errors, warnings, info, profile, columns


def print_report(path, errors, warnings, info, profile):
    print(f"\n=== Validation: {path.name} ===")
    for m in info:
        print(f"  ℹ {m}")
    print(f"\n  {'Column':<24}{'Type':<8}{'Filled':>7}{'Missing':>9}{'Unique':>8}")
    for name, kind, filled, missing, uniq, mixed, _ in profile:
        print(f"  {name[:23]:<24}{kind + ('*' if mixed else ''):<8}{filled:>7}{missing:>9}{uniq:>8}")
    for m in warnings:
        print(f"  ⚠ {m}")
    for m in errors:
        print(f"  ✖ {m}")
    print(f"\n  Result: {'FAILED' if errors else 'PASSED'} ({len(errors)} errors, {len(warnings)} warnings)\n")


# ---------------------------------------------------------------- analysis
def output_columns(questions):
    cols = []
    for name, q in questions.items():
        if q["type"] == "noul":
            cols.append((f"Jev {name} (prob)", name, "noul"))
        else:
            cols.append((f"Jev {name}", name, "value"))
            cols.append((f"Jev {name} conf.", name, "conf"))
    cols.append(("Jev Needs Human Review", None, "review"))
    return cols


def ask_jev(client, state, questions):
    r = client.system_one(state, questions)
    out = {}
    for name, a in r.answers.items():
        if a.type == "choice":
            out[name] = (a.choice, a.confidence)
        elif a.type == "score":
            out[name] = (round(a.score, 2), a.confidence)
        else:
            out[name] = (round(a.noul, 3), None)
    return out


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("file", type=Path)
    p.add_argument("--sheet")
    p.add_argument("--columns")
    p.add_argument("--id-column")
    p.add_argument("--questions", type=Path)
    p.add_argument("--output", type=Path)
    p.add_argument("--threshold", type=float, default=0.6)
    p.add_argument("--limit", type=int)
    p.add_argument("--workers", type=int, default=4)
    p.add_argument("--validate-only", action="store_true")
    a = p.parse_args()

    if not a.file.exists():
        raise SystemExit(f"ERROR: file not found: {a.file}")
    questions = json.loads(a.questions.read_text()) if a.questions else DEFAULT_QUESTIONS
    wb, ws, headers, rows = load_table(a.file, a.sheet)
    cols = [c.strip() for c in a.columns.split(",")] if a.columns else None
    errors, warnings, info, profile, cols = validate(headers, rows, cols, questions)
    print_report(a.file, errors, warnings, info, profile)
    if errors:
        sys.exit(1)
    if a.validate_only:
        return

    from typesafe_sdk import TypeSafeClient  # imported late so --validate-only works without the SDK/key
    client = TypeSafeClient()

    idx = {h: i for i, h in enumerate(headers)}
    id_i = idx.get(a.id_column, 0)
    width = len(headers)

    def json_safe(v):
        return v.isoformat() if isinstance(v, (datetime, date)) else v

    # rows to analyze: has data in at least one selected column, and isn't a note row
    targets = []
    for n, r in enumerate(rows, start=2):
        r = r + [None] * (width - len(r))
        if all(is_blank(r[idx[c]]) for c in cols):
            continue
        if width > 1 and all(is_blank(v) for v in r[1:]):
            continue
        state = {c: json_safe(r[idx[c]]) for c in cols if not is_blank(r[idx[c]])}
        if len(json.dumps(state, default=str)) > MAX_STATE_CHARS:
            print(f"  ⚠ row {n}: state too long, skipped")
            continue
        targets.append((n, r[id_i], state))
    if a.limit:
        targets = targets[: a.limit]
    print(f"Sending {len(targets)} rows to Jev ({len(questions)} questions each)…")

    def work(t):
        n, rid, state = t
        try:
            return n, rid, ask_jev(client, state, questions), None
        except Exception as e:  # keep going; record the error on the row
            return n, rid, None, f"{type(e).__name__}: {e}"

    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        results = list(ex.map(work, targets))

    # write output columns
    ocols = output_columns(questions) + [("Jev Error", None, "error")]
    start = ws.max_column + 1
    for j, (h, _, _) in enumerate(ocols):
        c = ws.cell(1, start + j, h)
        c.font, c.fill = Font(bold=True), HDR_FILL
        ws.column_dimensions[get_column_letter(start + j)].width = max(14, len(h) + 2)

    ok = [r for r in results if r[2]]
    failed = [r for r in results if not r[2]]
    for n, rid, ans, err in results:
        if err:
            ws.cell(n, start + len(ocols) - 1, err)
            print(f"  ✖ row {n} ({rid}): {err}")
            continue
        confs = [c for _, c in ans.values() if c is not None]
        for j, (_, qn, kind) in enumerate(ocols):
            v = (ans[qn][0] if kind in ("value", "noul") else
                 round(ans[qn][1], 2) if kind == "conf" else
                 ("YES" if confs and min(confs) < a.threshold else "") if kind == "review" else None)
            if v is not None:
                ws.cell(n, start + j, v)

    # summary sheet
    s = wb.create_sheet("Jev Summary")
    s.append(["Source", a.file.name]); s.append(["Columns sent", ", ".join(cols)])
    s.append(["Rows analyzed", len(ok)]); s.append(["Rows failed", len(failed)])
    s.append(["Needs human review", sum(1 for _, _, ans, _ in ok
              if [c for _, c in ans.values() if c is not None] and
              min(c for _, c in ans.values() if c is not None) < a.threshold)])
    s.append([])
    for name, q in questions.items():
        vals = [ans[name][0] for _, _, ans, _ in ok]
        s.append([f"{name} ({q['type']})"]); s.cell(s.max_row, 1).font = Font(bold=True)
        if q["type"] == "choice":
            for label, cnt in Counter(vals).most_common():
                s.append([label, cnt, f"{cnt / len(vals):.0%}" if vals else ""])
        elif q["type"] == "score":
            s.append(["average", round(sum(vals) / len(vals), 2) if vals else ""])
            s.append(["max", max(vals) if vals else ""])
        else:
            s.append(["likely yes (p>0.5)", sum(v > 0.5 for v in vals)])
        s.append([])
    s.column_dimensions["A"].width = 28; s.column_dimensions["B"].width = 18

    out = a.output or a.file.with_name(f"{a.file.stem}_jev.xlsx")
    if out.resolve() == a.file.resolve():
        raise SystemExit("ERROR: --output must differ from the input file")
    wb.save(out)
    print(f"\nDone: {len(ok)} analyzed, {len(failed)} failed → {out}")


if __name__ == "__main__":
    main()
