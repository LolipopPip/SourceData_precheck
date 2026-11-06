#!/usr/bin/env python3
"""Read-only, conservative numeric audit of scientific Source Data.

Python >=3.9. CSV uses the standard library; XLSX requires openpyxl; XLS needs xlrd.
Optional configured t-tests require scipy. No input files are modified.
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import math
import re
import statistics
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from itertools import combinations
from pathlib import Path


FIELDS = ["severity", "priority", "rule", "sheet", "cells", "related_sheet",
          "related_cells", "evidence", "next_step"]


def number(value):
    """Return finite Decimal or None; do not interpret dates, booleans or labels."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        raw = str(value)
    elif isinstance(value, str):
        raw = value.strip()
        if not raw or raw.endswith("%") or raw.startswith(("=", "<", ">")):
            return None
    else:
        return None
    try:
        result = Decimal(raw)
        return result if result.is_finite() else None
    except (InvalidOperation, ValueError):
        return None


def column_label(index):
    """1-based spreadsheet column index to Excel letters."""
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def cell_name(row, col):
    return f"{column_label(col)}{row}"


def finding(rule, sheet, cells, evidence, next_step, severity="review", priority=None):
    if priority is None:
        priority = "P3" if severity == "info" else "P2"
    return dict(severity=severity, priority=priority, rule=rule, sheet=sheet,
                cells=", ".join(map(str, cells)) if isinstance(cells, (list, tuple)) else str(cells),
                evidence=str(evidence), next_step=next_step)


def load_grids(path):
    suffix = path.suffix.lower()
    if suffix in (".csv", ".tsv"):
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            reader = csv.reader(stream, delimiter="\t" if suffix == ".tsv" else ",")
            rows = [list(row) for row in reader]
        return {path.stem: rows}, []
    if suffix == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise RuntimeError("Install xlrd to audit legacy .xls files: pip install xlrd>=2") from exc
        book = xlrd.open_workbook(str(path))
        grids = {}
        for ws in book.sheets():
            rows = []
            for ridx in range(ws.nrows):
                row = []
                for cidx in range(ws.ncols):
                    cell = ws.cell(ridx, cidx)
                    # Legacy Excel dates (ctype 3) are stored as floats, but are not measurements.
                    if cell.ctype in (0, 3, 5, 6):
                        value = None
                    elif cell.ctype == 4:
                        value = bool(cell.value)
                    else:
                        value = cell.value
                    row.append(value)
                rows.append(row)
            grids[ws.name] = rows
        return grids, []
    if suffix != ".xlsx":
        raise ValueError("Supported formats: .xls, .xlsx, .csv, .tsv")
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("Install openpyxl to audit .xlsx files: pip install openpyxl") from exc
    wb = load_workbook(path, read_only=True, data_only=True)
    formula_wb = load_workbook(path, read_only=True, data_only=False)
    grids = {}
    uncached = []
    try:
        for ws in wb.worksheets:
            grids[ws.title] = [list(row) for row in ws.iter_rows(values_only=True)]
            formula_sheet = formula_wb[ws.title]
            for ridx, row in enumerate(formula_sheet.iter_rows(), start=1):
                cached_row = grids[ws.title][ridx - 1] if ridx <= len(grids[ws.title]) else []
                for cidx, c in enumerate(row, start=1):
                    cached_value = cached_row[cidx - 1] if cidx <= len(cached_row) else None
                    if c.data_type == "f" and cached_value is None:
                        uncached.append((ws.title, c.coordinate))
    finally:
        wb.close()
        formula_wb.close()
    return grids, uncached


def row_segments(rows):
    """Consecutive rows containing numeric observations; blank/header rows split blocks."""
    segments, current = [], []
    for ridx, row in enumerate(rows, start=1):
        if any(number(v) is not None for v in row):
            current.append(ridx)
        elif current:
            segments.append(current)
            current = []
    if current:
        segments.append(current)
    return segments


def _informative_decimal_run(items):
    """Screen long, varied, non-integer observations; not short count/dose blocks."""
    values = [value for _, value in items]
    return (len(values) >= 7 and len(set(values)) >= 5
            and sum(value != value.to_integral_value() for value in values) >= 6
            and sum(-value.as_tuple().exponent >= 2 for value in values) >= 3)


def audit_repeated_columns(sheet, segment, records, identifier_cols):
    """Use seven-value fingerprints to avoid quadratic scans on wide workbooks."""
    issues = []
    candidates = sorted(set().union(*(set(records[r]) for r in segment)) - identifier_cols)
    fingerprints = defaultdict(set)
    for col in candidates:
        observations = [(row, records[row].get(col)) for row in segment]
        for start in range(len(observations) - 6):
            window = observations[start:start + 7]
            if (all(v is not None for _, v in window)
                    and all(window[i][0] + 1 == window[i + 1][0] for i in range(6))
                    and _informative_decimal_run(window)):
                fingerprints[(window[0][0], tuple(v for _, v in window))].add(col)

    candidate_pairs = set()
    capped = False
    max_pairs = 50
    for cols in fingerprints.values():
        if len(cols) < 2:
            continue
        for pair in combinations(sorted(cols), 2):
            candidate_pairs.add(pair)
            if len(candidate_pairs) >= max_pairs:
                capped = True
                break
        if capped:
            break
    if capped:
        issues.append(finding("duplicate_pair_search_limited", sheet,
            cell_name(segment[0], candidates[0]),
            "More than 50 long matching column pairs; showing the first 50 to bound processing.",
            "Narrow the panel range or inspect other matching columns manually; coverage is limited.",
            "info"))

    for a, b in sorted(candidate_pairs):
        matched, runs = [], []
        for row in segment:
            left, right = records[row].get(a), records[row].get(b)
            if left is not None and left == right:
                if matched and row != matched[-1][0] + 1:
                    runs.append(matched)
                    matched = []
                matched.append((row, left))
            elif matched:
                runs.append(matched)
                matched = []
        if matched:
            runs.append(matched)
        informative = [run for run in runs if _informative_decimal_run(run)]
        if not informative:
            continue
        longest = max(informative, key=len)
        a_rows = [r for r in segment if a in records[r]]
        b_rows = [r for r in segment if b in records[r]]
        entire_profiles = a_rows == b_rows == [r for r, _ in longest]
        rule = "identical_numeric_columns" if entire_profiles else "duplicate_numeric_block"
        left_cells = [cell_name(r, a) for r, _ in longest]
        right_cells = [cell_name(r, b) for r, _ in longest]
        evidence = (f"Columns {column_label(a)} and {column_label(b)} share "
                    f"{len(longest)} exact aligned decimal observations. "
                    f"Distinct numeric values: {len(set(v for _, v in longest))}.")
        issues.append(finding(rule, sheet, left_cells + right_cells, evidence,
            "Confirm labels, sample identity and independently counted source observations; "
            "this is not a determination of fabrication.",
            priority="P1" if entire_profiles and len(longest) >= 9 else "P2"))
    return issues


def audit_grid(sheet, rows):
    issues = []
    for segment in row_segments(rows):
        if len(segment) < 2:
            continue
        header = rows[segment[0] - 2] if segment[0] >= 2 else []
        identifier_cols = set()
        for j, title in enumerate(header, start=1):
            if isinstance(title, str) and re.search(
                r"\b(id|index|mouse\s*(number|no)|sample\s*(number|no)|animal\s*(number|no))\b",
                title, flags=re.IGNORECASE
            ):
                identifier_cols.add(j)
        records = {}
        for ridx in segment:
            records[ridx] = {j: number(v) for j, v in enumerate(rows[ridx - 1], start=1)
                             if number(v) is not None}

        # Duplicate multi-column numeric records, not individual rounded values.
        groups = defaultdict(list)
        for ridx, values in records.items():
            if len(values) >= 3 and any(val != 0 for val in values.values()):
                groups[tuple(sorted(values.items()))].append(ridx)
        for duplicate_rows in groups.values():
            if len(duplicate_rows) > 1:
                first = records[duplicate_rows[0]]
                cells = [cell_name(r, c) for r in duplicate_rows for c in first]
                issues.append(finding("duplicate_numeric_row", sheet, cells,
                    f"Identical multi-column numeric records on rows {duplicate_rows}.",
                    "Compare sample IDs and original instrument exports; intentional repeats may be valid."))

        columns = defaultdict(list)
        for ridx, values in records.items():
            for col, val in values.items():
                columns[col].append((ridx, val))

        issues.extend(audit_repeated_columns(sheet, segment, records, identifier_cols))

        # Internal gaps only: differing group n at the end is not an error.
        for col, items in columns.items():
            positions = {r for r, _ in items}
            if len(items) < 3:
                continue
            for ridx in range(min(positions) + 1, max(positions)):
                if ridx not in positions and records.get(ridx):
                    issues.append(finding("internal_missing_numeric_cell", sheet,
                        cell_name(ridx, col), "Blank/non-numeric value between numeric observations.",
                        "Check whether this is an intended missing value, a formula error, or a pasted blank."))

        for col, items in columns.items():
            if col in identifier_cols:
                continue
            # Do not use arbitrary digit tests on small datasets or integer counts.
            precise = [(r, v) for r, v in items if -v.as_tuple().exponent >= 2]
            if len(precise) >= 30:
                counts = Counter(abs(v.as_tuple().digits[-1]) for _, v in precise)
                digit, count = counts.most_common(1)[0]
                if count / len(precise) >= 0.8:
                    issues.append(finding("last_digit_concentration", sheet,
                        [cell_name(r, col) for r, _ in precise[:5]],
                        f"Last stored decimal digit {digit} occurs {count}/{len(precise)} times.",
                        "Check export precision, rounding settings and instrument quantization; this is not evidence of fabrication."))

            # High-precision exact duplicates, where ordinary rounded repeats are less informative.
            high_precision = defaultdict(list)
            for ridx, value in items:
                if -value.as_tuple().exponent >= 4:
                    high_precision[value].append(ridx)
            if len(items) >= 5:
                for value, repeated in high_precision.items():
                    if len(repeated) >= 2:
                        issues.append(finding("duplicate_high_precision_value", sheet,
                            [cell_name(r, col) for r in repeated],
                            f"Exact high-precision numeric value {value} occurs {len(repeated)} times.",
                            "Determine whether repeated values share an input, formula, or legitimate measurement limit."))

            # Report a progression only for >=5 consecutive rows; no Benford test by default.
            if len(items) >= 5:
                for start in range(len(items) - 4):
                    run = items[start:start + 5]
                    if any(run[i+1][0] != run[i][0] + 1 for i in range(4)):
                        continue
                    diffs = [run[i+1][1] - run[i][1] for i in range(4)]
                    if diffs[0] != 0 and all(x == diffs[0] for x in diffs):
                        issues.append(finding("arithmetic_progression", sheet,
                            [cell_name(r, col) for r, _ in run],
                            f"Five consecutive values have common difference {diffs[0]}.",
                            "Check whether these are designed doses/time points, transformed values, or actual observations."))
                        break

        # Compare columns in the same contiguous numeric section. Purely a review signal.
        candidate_columns = sorted(set(columns) - identifier_cols)
        if len(candidate_columns) > 40:
            issues.append(finding("wide_table_offset_check_skipped", sheet,
                cell_name(segment[0], candidate_columns[0]),
                f"{len(candidate_columns)} numeric columns: O(n^2) column-offset comparison skipped.",
                "Use a focused panel range/configuration for pairwise comparison if required.", "info"))
            continue
        for pos, a_col in enumerate(candidate_columns):
            for b_col in candidate_columns[pos + 1:]:
                aligned = [(r, records[r][b_col] - records[r][a_col]) for r in segment
                           if a_col in records[r] and b_col in records[r]]
                if len(aligned) >= 5 and aligned[0][1] != 0 and all(
                    diff == aligned[0][1] for _, diff in aligned
                ):
                    issues.append(finding("constant_column_offset", sheet,
                        [cell_name(r, a_col) + "/" + cell_name(r, b_col)
                         for r, _ in aligned[:5]],
                        f"Across {len(aligned)} aligned rows, column {column_label(b_col)} - "
                        f"{column_label(a_col)} is exactly {aligned[0][1]}.",
                        "Check whether this relationship follows from legitimate normalization or a calculation."))
    return issues


def _looks_like_identifier_or_design(values):
    """Exclude index rows, deliberately spaced doses, and near-constant scores."""
    if len(set(values)) < 3:
        return True
    if all(v == v.to_integral_value() and 0 <= v <= 5 for v in values):
        return True
    differences = [b - a for a, b in zip(values, values[1:])]
    if differences and differences[0] != 0 and all(d == differences[0] for d in differences):
        return True
    if (all(v == v.to_integral_value() and 1 <= v <= 100 for v in values)
            and values[0] == 1
            and all(b == a + 1 or b == 1 for a, b in zip(values, values[1:]))):
        return True
    return False


def _header_label(rows, row, col, orientation):
    """Use labels as context, never as a fourth numeric observation."""
    candidates = []
    if orientation == "horizontal":
        for preceding_col in range(col - 1, max(0, col - 4), -1):
            if preceding_col >= 1 and preceding_col <= len(rows[row - 1]):
                candidates.append(rows[row - 1][preceding_col - 1])
    else:
        for preceding_row in range(row - 1, max(0, row - 4), -1):
            if preceding_row >= 1 and col <= len(rows[preceding_row - 1]):
                candidates.append(rows[preceding_row - 1][col - 1])
    for value in candidates:
        if not isinstance(value, str) or not value.strip() or number(value) is not None:
            continue
        label = re.sub(r"\s+", " ", value.strip())
        if re.match(r"^(source\b|row title$|x$|id$|index$)", label, re.I):
            continue
        return label
    return ""


def _candidate_numeric_series(grids):
    """Extract true contiguous vertical and horizontal observations with A1 cells."""
    series = []
    for sheet, rows in grids.items():
        if not rows:
            continue
        designed_rows = set()
        for ridx, source_row in enumerate(rows, 1):
            numeric = [number(value) for value in source_row]
            run = [value for value in numeric if value is not None]
            if len(run) < 4:
                continue
            serial = (run[0] == 1 and all(v == v.to_integral_value() for v in run)
                      and all(b == a + 1 or b == 1 for a, b in zip(run, run[1:])))
            doses = (max(run) <= 1 and min(run) >= 0
                     and all(b - a == run[1] - run[0] for a, b in zip(run, run[1:])))
            if serial or doses:
                designed_rows.add(ridx)
        for orientation in ("vertical", "horizontal"):
            axis_count = (max((len(row) for row in rows), default=0)
                          if orientation == "vertical" else len(rows))
            for fixed in range(1, axis_count + 1):
                entries = []
                axis_length = len(rows) if orientation == "vertical" else len(rows[fixed - 1])
                for moving in range(1, axis_length + 2):
                    row, col = ((moving, fixed) if orientation == "vertical"
                                else (fixed, moving))
                    raw = (rows[row - 1][col - 1]
                           if row <= len(rows) and col <= len(rows[row - 1]) else None)
                    val = None if orientation == "vertical" and row in designed_rows else number(raw)
                    if val is not None:
                        entries.append((row, col, val))
                    elif entries:
                        if len(entries) >= 3:
                            vals = tuple(item[2] for item in entries)
                            if not _looks_like_identifier_or_design(vals):
                                start_row, start_col, _ = entries[0]
                                series.append(dict(sheet=sheet, orientation=orientation,
                                    label=_header_label(rows, start_row, start_col, orientation),
                                    values=vals,
                                    cells=tuple(cell_name(r, c) for r, c, _ in entries)))
                        entries = []
    return series


def _informative_repeat(values, left_label, right_label):
    """Classify numeric equality by specificity, never as a misconduct score."""
    if _looks_like_identifier_or_design(values):
        return None
    distinct = len(set(values))
    decimals = sum(v != v.to_integral_value() for v in values)
    precision = sum(-v.as_tuple().exponent >= 2 for v in values)
    if len(values) >= 7 and distinct >= 5 and decimals >= 6 and precision >= 3:
        return "P1" if len(values) >= 9 else "P2"
    if len(values) >= 5 and distinct >= 4 and decimals >= 4 and precision >= 3:
        return "P2"
    if (len(values) >= 7 and distinct >= 6 and decimals == 0
            and max(values) >= 10 and max(values) - min(values) >= 10
            and not any(re.search(r"\b(time|dose|index|serial|sample\s*(?:id|no))\b",
                                  label, re.I) for label in (left_label, right_label))):
        return "P2"
    if (len(values) >= 3 and distinct >= 3 and decimals == 0
            and max(values) >= 10 and min(values) >= 0
            and left_label and left_label.casefold() == right_label.casefold()
            and bool(re.search(r"(?:μm|µm|mm\b|nm\b|ng[/ ]|mg[/ ]|\d\s*(?:h|d|day))", left_label, re.I))):
        return "P2"
    return None


def audit_cross_panel_repeats(grids):
    """Compare disjoint numeric runs across panels/sheets by indexed triples."""
    series = _candidate_numeric_series(grids)
    fingerprints = defaultdict(list)
    for si, run in enumerate(series):
        values = run["values"]
        for start in range(len(values) - 2):
            triple = values[start:start + 3]
            if len(set(triple)) == 3 and not _looks_like_identifier_or_design(triple):
                fingerprints[triple].append((si, start))

    matches = {}
    skipped_fingerprints = 0
    first_skipped = None
    for occurrences in fingerprints.values():
        if len(occurrences) > 80:
            skipped_fingerprints += 1
            if first_skipped is None:
                first_skipped = series[occurrences[0][0]]
            continue
        for (ai, start_a), (bi, start_b) in combinations(occurrences, 2):
            if ai == bi:
                continue
            a, b = series[ai], series[bi]
            if a["sheet"] == b["sheet"] and set(a["cells"]) & set(b["cells"]):
                continue
            if (a["sheet"] == b["sheet"] and a["orientation"] == b["orientation"] == "vertical"
                    and [re.search(r"\d+$", cell).group() for cell in a["cells"]]
                    == [re.search(r"\d+$", cell).group() for cell in b["cells"]]):
                # v1.1 already checks aligned columns, including full and partial runs.
                continue
            av, bv = a["values"], b["values"]
            left, right = start_a, start_b
            while left > 0 and right > 0 and av[left - 1] == bv[right - 1]:
                left -= 1
                right -= 1
            end_a, end_b = start_a + 3, start_b + 3
            while end_a < len(av) and end_b < len(bv) and av[end_a] == bv[end_b]:
                end_a += 1
                end_b += 1
            values = av[left:end_a]
            priority = _informative_repeat(values, a["label"], b["label"])
            if not priority:
                continue
            if priority == "P1" and (len(values) != len(av) or len(values) != len(bv)):
                priority = "P2"  # A partial match is not two complete profiles.
            key = (min(ai, bi), max(ai, bi))
            if key not in matches or len(values) > len(matches[key][2]):
                matches[key] = (a, b, values, left, right, priority)

    issues = []
    for a, b, values, left, right, priority in matches.values():
        # The matching values are equal; use each run's own unaltered coordinates.
        a_cells = a["cells"][left:left + len(values)]
        b_cells = b["cells"][right:right + len(values)]
        note = f"; shared group label: {a['label']}" if a["label"] and a["label"].casefold() == b["label"].casefold() else ""
        item = finding("cross_panel_repeated_series", a["sheet"], list(a_cells),
            f"{len(values)} consecutive identical {a['orientation']}/{b['orientation']} observations"
            f" across panels{note}; values: {', '.join(map(str, values[:10]))}.",
            "Compare original sample identities, markers, formulas and legitimate shared controls; short integer coincidences may be expected.",
            priority=priority)
        item["related_sheet"] = b["sheet"]
        item["related_cells"] = ", ".join(b_cells)
        issues.append(item)
    if skipped_fingerprints:
        issues.append(finding("cross_panel_search_limited", first_skipped["sheet"],
            first_skipped["cells"][0],
            f"Skipped {skipped_fingerprints} extremely common three-value fingerprints "
            "(more than 80 occurrences each) to bound cross-sheet comparisons.",
            "Review heavily repeated or duplicated columns manually; this is a coverage notice, not an anomaly.",
            severity="info"))
    return sorted(issues, key=lambda item: (item["sheet"], item.get("related_sheet", ""), item["cells"]))


def values_for_range(ws, addr):
    try:
        from openpyxl.utils.cell import range_boundaries
    except ImportError as exc:
        raise RuntimeError("Configured cell-range audits require openpyxl") from exc
    min_col, min_row, max_col, max_row = range_boundaries(addr)
    if min_col is None or min_row is None:
        raise ValueError(f"Use explicit A1-style cell ranges: {addr!r}")
    return [number(ws[r - 1][c - 1] if r - 1 < len(ws) and c - 1 < len(ws[r - 1]) else None)
            for r in range(min_row, max_row + 1) for c in range(min_col, max_col + 1)]


def reported(ws, addr):
    values = values_for_range(ws, addr)
    if len(values) != 1 or values[0] is None:
        raise ValueError(f"Reported cell must be a single numeric cell: {addr}")
    return values[0]


def matches(expected, observed, decimals=None):
    if decimals is None:
        decimals = max(0, -observed.as_tuple().exponent)
    tolerance = Decimal("0.5") * (Decimal(10) ** -int(decimals)) + Decimal("0.000000001")
    return abs(Decimal(str(expected)) - observed) <= tolerance


def audit_config(grids, config, coverage):
    issues = []
    for item in config.get("aggregates", []):
        sheet = item["sheet"]
        ws = grids[sheet]
        values = [float(v) for v in values_for_range(ws, item["range"]) if v is not None]
        if not values:
            raise ValueError(f"No numeric values in {sheet}!{item['range']}")
        expected = {"n": len(values), "mean": statistics.mean(values)}
        if len(values) > 1:
            expected["sd"] = statistics.stdev(values)
            expected["sem"] = expected["sd"] / math.sqrt(len(values))
        for field, calc in expected.items():
            if field not in item:
                continue
            observed = reported(ws, item[field])
            if not matches(calc, observed, item.get("decimals")):
                issues.append(finding(f"reported_{field}_mismatch", sheet, item[field],
                    f"Reported {observed}; recalculated {calc:.12g} from {item['range']}.",
                    "Verify the exact plotted observations, statistical unit and rounding precision.", "confirmed"))
        coverage["configured_aggregates_checked"] += 1

    for item in config.get("identities", []):
        sheet = item["sheet"]
        ws = grids[sheet]
        total = reported(ws, item["total"])
        part_sum = sum((reported(ws, p) for p in item["parts"]), Decimal(0))
        if total != part_sum:
            issues.append(finding("count_identity_mismatch", sheet,
                [item["total"], *item["parts"]],
                f"Total {total} does not equal sum of parts {part_sum}.",
                "Check mutually exclusive categories, missing cells, and counting definitions.", "confirmed"))
        coverage["configured_identities_checked"] += 1

    for item in config.get("ratios", []):
        sheet = item["sheet"]
        ws = grids[sheet]
        numerator, denominator = reported(ws, item["numerator"]), reported(ws, item["denominator"])
        if denominator == 0:
            issues.append(finding("zero_denominator", sheet, item["denominator"],
                "Configured ratio has a zero denominator.", "Check the formula and data source.", "confirmed"))
        else:
            scale = Decimal(str(item.get("scale", 100)))
            calc = numerator / denominator * scale
            observed = reported(ws, item["reported"])
            if not matches(calc, observed, item.get("decimals")):
                issues.append(finding("reported_ratio_mismatch", sheet,
                    [item["numerator"], item["denominator"], item["reported"]],
                    f"Reported {observed}; recalculated {calc:.12g} (scale={scale}).",
                    "Recalculate from original counts and verify the percentage denominator.", "confirmed"))
        coverage["configured_ratios_checked"] += 1

    for item in config.get("ttests", []):
        try:
            from scipy import stats
        except ImportError:
            issues.append(finding("ttest_not_checked", item["sheet"], item["reported_p"],
                "scipy is unavailable; no p-value was recalculated.",
                "Install scipy and rerun; check statistical design and biological replication."))
            continue
        sheet = item["sheet"]
        ws = grids[sheet]
        a = [float(v) for v in values_for_range(ws, item["a"]) if v is not None]
        b = [float(v) for v in values_for_range(ws, item["b"]) if v is not None]
        design = item["design"].lower()
        if len(a) < 2 or len(b) < 2:
            raise ValueError(f"At least two observations per group required: {sheet}")
        if design == "welch":
            res = stats.ttest_ind(a, b, equal_var=False, alternative="two-sided")
        elif design == "student":
            res = stats.ttest_ind(a, b, equal_var=True, alternative="two-sided")
        elif design == "paired":
            if len(a) != len(b):
                raise ValueError(f"Paired groups need matching lengths: {sheet}")
            res = stats.ttest_rel(a, b, alternative="two-sided")
        else:
            raise ValueError(f"Unsupported t-test design: {design}; use welch, student or paired")
        p = float(res.pvalue)
        observed = reported(ws, item["reported_p"])
        if not math.isfinite(p):
            issues.append(finding("ttest_not_computable", sheet, item["reported_p"],
                "SciPy returned a non-finite p-value.",
                "Inspect sample size, zero variance and the statistical design."))
        elif not matches(p, observed, item.get("decimals")):
            issues.append(finding("reported_p_mismatch", sheet, item["reported_p"],
                f"Reported p={observed}; {design} two-sided t-test recalculates p={p:.12g}.",
                "Confirm paired/independent samples, test assumptions, tails, correction and graph settings.",
                "confirmed"))
        coverage["configured_ttests_checked"] += 1
    return issues


def audit_file(path, config=None):
    path = Path(path)
    grids, uncached = load_grids(path)
    coverage = dict(sheets_scanned=len(grids), configured_aggregates_checked=0,
                    configured_identities_checked=0, configured_ratios_checked=0,
                    configured_ttests_checked=0, uncached_formula_cells=len(uncached),
                    legacy_xls_formula_provenance_unavailable=(path.suffix.lower() == ".xls"))
    findings = []
    for sheet, rows in grids.items():
        findings.extend(audit_grid(sheet, rows))
    findings.extend(audit_cross_panel_repeats(grids))
    for sheet, coord in uncached[:100]:
        findings.append(finding("uncached_formula", sheet, coord,
            "Excel formula is present but its cached result is empty.",
            "Recalculate the workbook in Excel and save a copy before auditing."))
    if config:
        findings.extend(audit_config(grids, config, coverage))
    return findings, coverage


CHECKLIST_ZH = """# Source Data 人工复核清单

- [ ] **逐项追溯**：HTML 的每条提示都对应正确的工作表、Excel 单元格和 Figure/panel；记录人工解释及证据。
- [ ] **独立样本**：每个散点可追溯到动物/细胞/样本 ID 及原始采集记录；确认生物学重复单位 `n`，避免伪重复。
- [ ] **重复来源**：对长重复列，区分不同 marker、共享输入、嵌套细胞群、同组横竖重复排版和真正独立重复。
- [ ] **设计数值**：核查 0、剂量、时间轴、整数计数和百分比计算；不要自动把相同短序列定为异常。
- [ ] **缺失值**：核对内部空白、未检测、排除与检出限，绝不直接用 0 填补。
- [ ] **统计汇总**：对照 GraphPad/R/Python 原脚本复算 Mean、SD、SEM、n、比例、归一化和坐标值。
- [ ] **统计检验**：核对组别、配对、检验单双侧、多重比较校正及原始 P 值。
- [ ] **计数守恒**：按真实互斥关系核验分类总数，例如总数 = A 类 + B 类（仅当 A、B 互斥且穷尽）。
- [ ] **图表覆盖**：确认所有主图和补图 panel 都有对应的数值 Source Data，图注中的单位和样本数正确。
- [ ] **留痕**：保留原始数据、计数文件、脚本、审查日期、复核者、解释与后续动作，不为消除提示而修改原始观测。

**边界**：本工具只筛查数值模式；它不能认证原始实验、图像真实性、科研诚信、公开数据库提交或所有统计检验。
"""


CHECKLIST_EN = """# Source Data Manual Review Checklist

- [ ] **Trace every finding**: Link each HTML finding to its worksheet, exact Excel cells and figure panel; record the review outcome and evidence.
- [ ] **Independent units**: Trace plotted points to animal/cell/sample IDs and acquisition records. Verify the biological unit `n` and avoid pseudoreplication.
- [ ] **Repeated profiles**: Distinguish independent markers, shared inputs, nested groups, repeated layout and legitimate shared controls.
- [ ] **Designed values**: Check zeros, doses, time points, counts and derived percentages before flagging short repetitions.
- [ ] **Missing data**: Check blanks, exclusions, limits of detection and non-measurements; never substitute zeros automatically.
- [ ] **Summary statistics**: Recalculate Mean, SD, SEM, n, ratios and transformed values from the original plotted observations.
- [ ] **Statistical tests**: Check pairing, test direction, assumptions, multiplicity correction and original P values.
- [ ] **Count identities**: Verify totals against mutually exclusive and exhaustive components, e.g. total = group A + group B only when justified.
- [ ] **Figure coverage**: Ensure every main and supplementary quantitative panel has source values, units and appropriate sample-size information.
- [ ] **Provenance**: Preserve originals, instrument exports, analysis scripts, review date, reviewer and documented explanations; never edit raw observations merely to suppress alerts.

**Scope:** This scanner finds numerical patterns only. It does not certify experimental provenance, image authenticity, research integrity or compliance with any journal policy.
"""


def write_reports(input_path, findings, coverage, output_dir, lang="en"):
    if lang not in {"en", "zh"}:
        raise ValueError("Supported report languages: en, zh")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    renderer_path = Path(__file__).with_name("report_html.py")
    spec = importlib.util.spec_from_file_location("source_data_report_html", renderer_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load HTML renderer: {renderer_path}")
    renderer = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(renderer)

    # In Chinese mode, translate explanations and suggested actions while
    # retaining exact detector evidence (including numeric values and A1 cells).
    localized = []
    for item in findings:
        entry = dict(item)
        if lang == "zh":
            _, explanation, action = renderer._meaning(item, lang="zh")
            entry["evidence"] = f"{explanation} 检测细节（原始技术记录）：{item.get('evidence', '')}"
            entry["next_step"] = action
        localized.append(entry)
    with (output_dir / "findings.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(localized)
    by_severity = Counter(f["severity"] for f in findings)
    if lang == "zh":
        summary = [
            "# Source Data 完整性核查 — 数值审查", "", f"输入文件：`{Path(input_path).name}`",
            f"扫描工作表：{coverage['sheets_scanned']}",
            f"已映射的算术不一致：{by_severity.get('confirmed', 0)}",
            f"需人工复核的提示：{by_severity.get('review', 0)}",
            f"P1 优先提示：{sum(f.get('priority') == 'P1' for f in findings)}",
            "", "## 覆盖范围",
            f"- 已配置的 Mean/SD/SEM/n 区域：{coverage['configured_aggregates_checked']}",
            f"- 已配置的计数恒等式：{coverage['configured_identities_checked']}",
            f"- 已配置的比例：{coverage['configured_ratios_checked']}",
            f"- 已配置的检验：{coverage['configured_ttests_checked']}",
            f"- 未缓存的公式：{coverage['uncached_formula_cells']}",
            f"- 旧版 .xls 无法核实公式来源：{coverage.get('legacy_xls_formula_provenance_unavailable', False)}",
            "", "## 重要边界",
            "- 未配置明确映射且没有原始证据时，原始值、排除数据、样本独立性、实验意义及 P 值均未独立验证。",
            "- 自动提示不意味着数据造假或科研不端。",
            "- 没有提示不代表数据有效或符合投稿要求。",
            "- 不修改输入文件，也不会自动把数据上传到网络。",
            "", "## 核查详情", "打开 `report.html` 查看中文解释、两个工作表的高亮位置和 A1 单元格，或查阅 `findings.csv`。",
        ]
    else:
        summary = [
            "# Source Data Integrity Check — numerical audit",
            "", f"Input: `{Path(input_path).name}`",
            f"Sheets scanned: {coverage['sheets_scanned']}",
            f"Confirmed arithmetic inconsistencies: {by_severity.get('confirmed', 0)}",
            f"Review signals: {by_severity.get('review', 0)}",
            f"P1 priority (long identical numeric profiles): {sum(f.get('priority') == 'P1' for f in findings)}",
            "", "## Coverage",
            f"- Configured mean/SD/SEM/n blocks: {coverage['configured_aggregates_checked']}",
            f"- Configured count identities: {coverage['configured_identities_checked']}",
            f"- Configured ratios: {coverage['configured_ratios_checked']}",
            f"- Configured tests: {coverage['configured_ttests_checked']}",
            f"- Formula cells without cached values: {coverage['uncached_formula_cells']}",
            f"- Legacy .xls formula provenance unavailable: {coverage.get('legacy_xls_formula_provenance_unavailable', False)}",
            "", "## Important limits",
            "- Source values, excluded observations, sample independence, biological meaning, and exact p-values are not independently verified unless explicitly checked with source evidence and configuration.",
            "- A review signal is not evidence of fabrication or research misconduct.",
            "- No flags does not mean the dataset is valid or publication-ready.",
            "- The original input file is never modified or uploaded automatically.",
            "", "## Finding details", "Open `report.html` to inspect findings with highlighted A1 cells, or use `findings.csv`.",
        ]
    (output_dir / "summary.md").write_text("\n".join(summary) + "\n", encoding="utf-8")
    (output_dir / "manual-review-checklist.md").write_text(
        CHECKLIST_ZH if lang == "zh" else CHECKLIST_EN, encoding="utf-8")
    grids, _ = load_grids(Path(input_path))
    (output_dir / "report.html").write_text(
        renderer.render_html(Path(input_path).name, localized, grids, coverage, lang=lang), encoding="utf-8")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Input .xls, .xlsx, .csv or .tsv")
    parser.add_argument("--config", type=Path, help="Optional explicit cell mapping in JSON")
    parser.add_argument("--out-dir", type=Path, default=Path("source-data-audit"))
    parser.add_argument("--lang", choices=("en", "zh"), default="en",
                        help="Report language: en (default) or zh (Simplified Chinese)")
    args = parser.parse_args(argv)
    if not args.input.is_file():
        parser.error(f"File does not exist: {args.input}")
    config = json.loads(args.config.read_text(encoding="utf-8")) if args.config else None
    findings, coverage = audit_file(args.input, config)
    write_reports(args.input, findings, coverage, args.out_dir, lang=args.lang)
    print(f"{len(findings)} findings. Reports: {args.out_dir.resolve()}")


if __name__ == "__main__":
    main()
