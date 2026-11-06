---
name: source-data-integrity-check
description: Use when a researcher asks to audit scientific Source Data spreadsheets before submission, find repeated numerical values across worksheets, columns, rows or panels, inspect blanks and numerical inconsistencies, recalculate explicitly mapped mean/SD/SEM/n/percentages/p-values, or generate an offline HTML report with highlighted Excel cells. Supports English by default and Chinese on request; accepts .xls/.xlsx/.csv/.tsv, not image forensics.
---

# Source Data Integrity Check v1.3

Review numerical Source Data locally and conservatively. Produce a **standalone English HTML audit by default** with an expandable explanation per finding, original Figure/sheet name, exact A1 cell addresses, and a highlighted source-value excerpt. If the researcher explicitly requests Chinese, produce the report in Simplified Chinese. Never modify original measurements or infer research misconduct from a software signal.

## 1. Run the automatic read-only audit

Locate this skill directory and execute:

```bash
python3 scripts/audit_source_data.py "Source Data.xlsx" --out-dir audit-report
```

Language controls **all user-facing report files** (HTML, CSV explanatory text, summary and manual-review checklist). Use English unless the user explicitly requests Chinese or supplies `--lang zh`:

```bash
python3 scripts/audit_source_data.py "Source Data.xlsx" --lang zh --out-dir audit-report-zh
```

The default is `--lang en`. Language choice changes explanations only, never detection thresholds, numerical findings, source values, or A1 coordinates. Exact technical evidence may retain short English tokens in the Chinese report for traceability.

Accepted inputs: `.xlsx` (requires `openpyxl`), legacy `.xls` (requires `xlrd>=2.0.1`), `.csv`, `.tsv` (standard library). Check `requirements.txt` if optional packages are missing. Do not require internet or upload unpublished source data to third-party services.

On success, hand the researcher these files:

- `report.html` **primary output** — open in any browser, search/filter by Figure and P1/P2/P3, expand every rule hit, and view yellow-highlighted A1 cells within the corresponding source table snippet. It contains copied raw values: treat it as confidential if the original data are unpublished.
- `findings.csv` — machine-readable finding with priority, sheet, A1 cells, optional `related_sheet` / `related_cells` for cross-sheet matching, evidence and next step.
- `summary.md` — scan coverage and explicit unverified analyses.
- `manual-review-checklist.md` — manual provenance/replicate/statistics checklist.

If the user requests an explanation, **read the generated report and explain prioritized findings individually**. Never present a bare count of alerts as a fraud score or a publisher compliance certification.

## 2. Interpret Geng-inspired repeated-data checks correctly

The detector includes the original conservative within-sheet rules and a new **panel-aware cross-sheet / cross-orientation comparison**:

- `identical_numeric_columns`: at least seven **row-aligned** values match over complete columns, with enough varying, non-integer observations; nine matching values receive P1 priority. For different markers this merits investigation of original counts and cell/sample identity.
- `duplicate_numeric_block`: a long partial row-aligned repeated decimal run in two columns; seven or more varied observations, rather than four coincidental integer counts, are required.
- `cross_panel_repeated_series`: find contiguous numeric runs across worksheets, separate regions and horizontal/vertical orientations. Compare only disjoint original A1 cells; keep the longest relevant match per pair. Both worksheet names, both A1 ranges and the highlighted excerpts appear in the same HTML finding.

The cross-panel detector intentionally uses tiered conditions to reduce false positives:

| Repeated data | Default treatment |
| --- | --- |
| ≥9 varied decimal values, both **complete** profiles | P1 review effort; confirm that the series are genuinely independent |
| ≥7 varied decimal values, including partial profiles | P2; shorter ≥5 varied decimal runs require several precise non-integer values |
| ≥7 varied, non-linear integer measurements (not sample numbers) | P2; review original count provenance |
| ≥3 varied integers with a matching nearby informative dose/time label, e.g. `2.5 mg/mL` | P2 *candidate*; a short integer coincidence does not establish an error |
| `1,2,3...` sample indexes, repeated index resets, equal-spacing dose/time sequences, all-zero/near-constant low ordinal scores | Suppress; these are weak or designed-value patterns |

Text labels such as `2.5 mg/mL` are used as **context**, not counted as a fourth measurement in a numeric repeated block. The label search uses nearby cells; complex merged/multi-level headers may require manual inspection. Existing aligned within-sheet column checks remain authoritative, so the cross-panel detector avoids issuing duplicate findings for them.

When a three-value fingerprint occurs more than 80 times, the cross-panel matcher skips that overly common fingerprint and emits `cross_panel_search_limited` as P3 coverage information. Do not interpret a skipped region as checked and unique.

It retains checks for repeated multi-column rows, repeated high-precision points, five-point progressions, fixed column offsets, internal blank cells, and conservative last-digit concentration. It suppresses all-zero duplicate rows and does not treat small-count 4–6-value blocks as high-confidence evidence. **Do not import Geng's generic four-integer-block high-risk label or apply Benford's law globally to heterogeneous biomedical percentages, time bins, counts and instrument data.**

Priority means manual review effort only:

| Level | Meaning | Recommended response |
| --- | --- | --- |
| P1 | Long full numeric profiles exactly equal | Prioritize source records, marker labels and sample IDs. This is **not** a misconduct verdict. |
| P2 | Potential numerical discrepancy or qualified repetition | Compare with raw records, formulas, figure legends and units. |
| P3 | Low-specificity or coverage information | Distinguish designed values, zeros, counts, display precision and unsupported checks. |

The HTML preview is an **embedded excerpt**, not a live Excel editor: for a flagged location, show the referenced workbook sheet and address, then consult the actual `.xls/.xlsx` file before drawing conclusions. Formulas in legacy `.xls` cannot be independently reconstructed from `xlrd` and must be marked unverified. A blank or zero can be valid.

## 3. Recalculate reported statistics only with explicit mapping

When figure legends and plotting groups identify biological replicates, read [mapping-and-reporting.md](references/mapping-and-reporting.md), confirm the mapping with the researcher, then run:

```bash
python3 scripts/audit_source_data.py "Source Data.xlsx" --config audit-config.json --out-dir audit-report
```

Supported explicit mappings:

- `aggregates`: mean, sample SD, SEM, n.
- `identities`: total count and disjoint component counts.
- `ratios`: numerator, denominator and report scale.
- `ttests`: paired, Welch or Student two-sided t-tests with `scipy`.

Confirm experimental unit, pairing, exclusions, rounding, numerator/denominator, transformed axes, and multiplicity correction. Unspecified statistics remain **unchecked**, not passed. A `confirmed` label means an arithmetic mismatch *under the specified mapping*, not fabrication.

## 4. Read the report, then manually verify

1. Open `report.html`; examine every P1 finding first and use its highlighted Excel excerpt and sheet/cell addresses to locate the original values. For cross-sheet findings, compare the **two separately highlighted worksheet excerpts** in the same card.
2. Compare every reported match against sample IDs, image-count sheets, instrument exports, timing, dose, processing scripts, formulas and overlapping biological markers. For rows/columns carrying the same source by design, document that explanation.
3. Inspect P2 and P3, including high-precision repetitions and missing values. Do not replace missing values or discard outliers to eliminate alerts.
4. Read `manual-review-checklist.md`. Confirm every main and supplementary panel, biological n, Mean/SD/SEM, exact test, independent units, plotted observations and caption definitions.
5. Record the finding ID, reviewer, original evidence, explanation and any proposed correction in an audit trail. Preserve the untouched source file and rerun after separately documented edits if authorized.

Never claim this numerical scanner verifies microscopy image reuse, gel integrity, instrument authenticity, raw sequencing provenance, all p-values, or journal policy compliance. If no findings are returned, explain the checks' blind spots.

## 5. Verify Skill behavior

```bash
python3 -B -m unittest discover -s tests -v
```

Before using the Skill on an unfamiliar XLS layout, check whether unusual merged cells, horizontal sample series or multiple groups within one sheet require manually defined panel boundaries. The renderer escapes workbook text and uses no external CSS, JS or online services.

## 6. Privacy and public-release boundary

- Never incorporate a researcher's workbook, file name, cell values, figure labels, individual identifiers, data-derived examples, audit output or unpublished results into `SKILL.md`, source code, tests, sample data, documentation, archives, or public repositories.
- The repository's automated tests must use invented synthetic data unrelated to any user session. For a user-authorized end-to-end check, use a public article's example dataset **outside** the skill directory; keep all generated reports outside the skill directory and do not include them in a release.
- This Skill runs locally, requires no remote API, and must not upload user data, telemetry, reports or temporary files. Only give user data and report files back to the person who supplied them, not to another Skill user.
- Before publication, inspect every included file and ZIP member for identifying strings and data-derived snippets. Remove such content even if it came from an earlier debugging session. A previously published repository may retain information in Git history; deleting it from a current file alone does not erase that history.

**Coverage boundary:** Cross-sheet and horizontal/vertical repeated series are now scanned conservatively, but arbitrary shuffled order, non-contiguous observations, complex header hierarchies, different rounded precisions, and short unlabeled integer repeats can remain undetected. The scanner does not infer whether the repeated source is a legitimate shared control, nor validate source provenance or all statistical panels without explicit mapping. A missed match is **unchecked**, never proof of uniqueness.
