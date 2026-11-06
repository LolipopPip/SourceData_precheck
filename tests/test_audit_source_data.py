"""Regression tests for a read-only Source Data audit."""

import csv
import tempfile
import unittest
from pathlib import Path
from importlib.util import module_from_spec, spec_from_file_location


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "audit_source_data.py"
SPEC = spec_from_file_location("audit_source_data", SCRIPT)
audit_module = module_from_spec(SPEC)
SPEC.loader.exec_module(audit_module)


class SourceDataAuditTests(unittest.TestCase):
    def test_skill_tree_never_bundles_research_workbooks_or_audit_outputs(self):
        root = SCRIPT.parents[1]
        private_artifacts = {".xlsx", ".xls", ".csv", ".tsv", ".pdf", ".html"}
        bundled = [str(path.relative_to(root)) for path in root.rglob("*")
                   if path.is_file() and path.suffix.lower() in private_artifacts]
        self.assertEqual(bundled, [], "Real Source Data and generated reports must stay outside the published Skill")

    def make_csv(self, folder, rows):
        path = Path(folder) / "Source Data.csv"
        with path.open("w", newline="", encoding="utf-8") as handle:
            csv.writer(handle).writerows(rows)
        return path

    def test_contextual_cross_sheet_integer_run_is_located_without_dose_as_data(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            wb = Workbook()
            first = wb.active
            first.title = "Panel Alpha"
            second = wb.create_sheet("Panel Beta")
            for ws, start in ((first, 14), (second, 6)):
                ws[f"D{start-1}"] = "3.2 mg/mL"
                for offset, value in enumerate((28, 71, 94)):
                    ws[f"D{start+offset}"] = value
            path = Path(folder) / "synthetic.xlsx"
            wb.save(path)
            findings, coverage = audit_module.audit_file(path)
            hits = [f for f in findings if f["rule"] == "cross_panel_repeated_series"]
            self.assertEqual(len(hits), 1)
            self.assertEqual(hits[0]["priority"], "P2")
            self.assertEqual(hits[0]["sheet"], "Panel Alpha")
            self.assertEqual(hits[0]["cells"], "D14, D15, D16")
            self.assertEqual(hits[0]["related_sheet"], "Panel Beta")
            self.assertEqual(hits[0]["related_cells"], "D6, D7, D8")
            self.assertIn("3.2 mg/mL", hits[0]["evidence"])
            output = Path(folder) / "report"
            audit_module.write_reports(path, findings, coverage, output)
            with (output / "findings.csv").open(encoding="utf-8-sig", newline="") as handle:
                exported = list(csv.DictReader(handle))
            self.assertTrue(any(f["related_sheet"] == "Panel Beta"
                                and "D8" in f["related_cells"] for f in exported))

    def test_cross_orientation_long_decimal_match_keeps_both_coordinates(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            wb = Workbook()
            a = wb.active
            a.title = "Figure A"
            b = wb.create_sheet("Figure B")
            observations = [11.44, 15.31, 16.47, 22.56, 26.32, 10.35, 29.91, 16.38]
            a["B2"] = "Intensity"
            b["A3"] = "Intensity"
            for index, value in enumerate(observations):
                a[f"B{index+3}"] = value
                b.cell(row=3, column=index + 2, value=value)
            path = Path(folder) / "cross.xlsx"
            wb.save(path)
            findings, _ = audit_module.audit_file(path)
            hits = [f for f in findings if f["rule"] == "cross_panel_repeated_series"]
            self.assertEqual(len(hits), 1)
            self.assertIn("B3", hits[0]["cells"])
            self.assertIn("B10", hits[0]["cells"])
            self.assertIn("I3", hits[0]["related_cells"])

    def test_cross_sheet_index_dose_and_small_ordinal_repetitions_are_filtered(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            wb = Workbook()
            a = wb.active
            a.title = "Figure 1"
            b = wb.create_sheet("Figure 2")
            for ws in (a, b):
                for col in range(1, 12):
                    ws.cell(row=2, column=col, value=col)
                for col, value in enumerate((0, 0.2, 0.4, 0.6, 0.8, 1.0), 1):
                    ws.cell(row=4, column=col, value=value)
                ws["A6"] = "Behavior score"
                for col, value in enumerate((2, 5, 4, 2, 5, 4), 2):
                    ws.cell(row=6, column=col, value=value)
            path = Path(folder) / "labels.xlsx"
            wb.save(path)
            findings, _ = audit_module.audit_file(path)
            self.assertFalse(any(f["rule"] == "cross_panel_repeated_series" for f in findings))

    def test_short_integer_run_requires_matching_nontrivial_group_label(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            wb = Workbook()
            a = wb.active
            a.title = "Figure A"
            b = wb.create_sheet("Figure B")
            a["B2"], b["B2"] = "3.2 mg/mL", "1.6 mg/mL"
            for row, value in enumerate((28, 71, 94), 3):
                a[f"B{row}"] = value
                b[f"B{row}"] = value
            path = Path(folder) / "different_groups.xlsx"
            wb.save(path)
            findings, _ = audit_module.audit_file(path)
            self.assertFalse(any(f["rule"] == "cross_panel_repeated_series" for f in findings))

    def test_long_varied_integer_profile_matches_across_sheets_without_same_label(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as folder:
            wb = Workbook()
            a = wb.active
            a.title = "Figure A"
            b = wb.create_sheet("Figure B")
            a["B2"], b["B2"] = "Marker A", "Marker B"
            for row, val in enumerate((12, 39, 24, 81, 63, 45, 117, 54), 3):
                a[f"B{row}"] = val
                b[f"B{row}"] = val
            path = Path(folder) / "integer_counts.xlsx"
            wb.save(path)
            findings, _ = audit_module.audit_file(path)
            hits = [f for f in findings if f["rule"] == "cross_panel_repeated_series"]
            self.assertEqual(len(hits), 1)
            self.assertEqual(hits[0]["priority"], "P2")

    def test_repeated_horizontal_profile_within_one_sheet(self):
        first = [None, 10.23, 13.47, 17.59, 22.61, 24.71, 28.83, 32.97, 36.12]
        grids = {"Figure A": [["Control"], first,
                              ["Other"], ["Spacer"], ["Treated"], first]}
        hits = audit_module.audit_cross_panel_repeats(grids)
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0]["sheet"], "Figure A")
        self.assertEqual(hits[0]["related_sheet"], "Figure A")
        self.assertIn("B2", hits[0]["cells"])
        self.assertIn("I6", hits[0]["related_cells"])

    def test_aligned_same_sheet_decimal_columns_not_reported_twice(self):
        values = [11.44, 15.31, 16.47, 22.56, 26.32, 10.35, 29.91, 16.38, 19.79]
        grids = {"Figure A": [["Marker A", "Marker B"]] + [[v, v] for v in values]}
        old_hits = audit_module.audit_grid("Figure A", grids["Figure A"])
        new_hits = audit_module.audit_cross_panel_repeats(grids)
        self.assertEqual(sum(f["rule"] == "identical_numeric_columns" for f in old_hits), 1)
        self.assertEqual(new_hits, [])

    def test_partial_nine_decimal_match_is_not_p1_full_profile(self):
        values = [11.44, 15.31, 16.47, 22.56, 26.32, 10.35, 29.91, 16.38, 19.79]
        grids = {"Figure A": [["A"], [1.23]] + [[v] for v in values],
                 "Figure B": [["B"]] + [[v] for v in values]}
        hits = audit_module.audit_cross_panel_repeats(grids)
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0]["priority"], "P2")

    def test_cross_sheet_search_is_bounded_and_discloses_coverage_limit(self):
        vals = (11.44, 15.31, 16.47, 22.56, 26.32, 10.35, 29.91, 16.38, 19.79)
        matrix = [[f"Marker {c}" for c in range(90)]] + [[v] * 90 for v in vals]
        grids = {"Figure A": matrix, "Figure B": matrix}
        hits = audit_module.audit_cross_panel_repeats(grids)
        self.assertTrue(any(f["rule"] == "cross_panel_search_limited" for f in hits))
        self.assertLess(len(hits), 10)

    def test_numeric_sample_header_row_is_not_prepended_to_vertical_measurements(self):
        values = (11.44, 15.31, 16.47, 22.56, 26.32, 10.35, 29.91, 16.38)
        matrix = [["Sample"], list(range(1, 9))] + [[None, v] for v in values]
        hits = audit_module.audit_cross_panel_repeats({"Figure A": matrix, "Figure B": matrix})
        repeated = [f for f in hits if f["rule"] == "cross_panel_repeated_series"]
        self.assertEqual(len(repeated), 1)
        self.assertEqual(repeated[0]["priority"], "P2")
        self.assertEqual(repeated[0]["cells"].split(", ")[0], "B3")
        self.assertNotIn("B2", repeated[0]["cells"])

    def test_detects_numeric_duplicates_and_constant_offsets(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [
                ["Mouse ID", "Control", "Treated", "Other"],
                ["M1", 1.12, 1.42, 4.1],
                ["M2", 2.23, 2.53, 5.2],
                ["M3", 3.37, 3.67, 6.3],
                ["M4", 4.51, 4.81, 7.4],
                ["M5", 5.68, 5.98, 8.5],
                ["M6", 5.68, 5.98, 8.5],
            ])
            findings, coverage = audit_module.audit_file(path)
            kinds = {item["rule"] for item in findings}
            self.assertIn("duplicate_numeric_row", kinds)
            self.assertIn("constant_column_offset", kinds)
            self.assertTrue(all(item["severity"] == "review" for item in findings))
            self.assertEqual(coverage["sheets_scanned"], 1)

    def test_short_integer_count_column_does_not_get_last_digit_alarm(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [["Count"], [5], [15], [25], [35], [45]])
            findings, _ = audit_module.audit_file(path)
            self.assertNotIn("last_digit_concentration", {f["rule"] for f in findings})

    def test_configured_mean_sem_n_and_count_relation_are_recalculated(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            self.skipTest("openpyxl is not installed")
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Fig 5"
            for row, value in enumerate([1, 2, 3, 4], start=2):
                ws[f"B{row}"] = value
            ws["B7"] = 99  # incorrect mean; 2.5 expected
            ws["B8"] = 0.7  # incorrect SEM; sqrt(5/3)/2 expected
            ws["B9"] = 4
            ws["C2"] = 10
            ws["D2"] = 4
            ws["E2"] = 5  # C2 != D2+E2
            wb.save(path)
            config = {
                "aggregates": [{"sheet": "Fig 5", "range": "B2:B5", "mean": "B7", "sem": "B8", "n": "B9"}],
                "identities": [{"sheet": "Fig 5", "total": "C2", "parts": ["D2", "E2"]}],
            }
            before = path.read_bytes()
            findings, coverage = audit_module.audit_file(path, config=config)
            after = path.read_bytes()
            rules = [f["rule"] for f in findings]
            self.assertIn("reported_mean_mismatch", rules)
            self.assertIn("reported_sem_mismatch", rules)
            self.assertIn("count_identity_mismatch", rules)
            self.assertNotIn("reported_n_mismatch", rules)
            self.assertEqual(before, after)
            self.assertEqual(coverage["configured_aggregates_checked"], 1)

    def test_generates_reports_and_explains_unchecked_statistics(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [["Group A", "Group B"], [1.2, 2.3], [3.4, 4.5]])
            findings, coverage = audit_module.audit_file(path)
            output = Path(folder) / "reports"
            audit_module.write_reports(path, findings, coverage, output)
            self.assertTrue((output / "findings.csv").is_file())
            self.assertTrue((output / "report.html").is_file())
            self.assertIn("Source Data Integrity Audit", (output / "report.html").read_text(encoding="utf-8"))
            self.assertTrue((output / "summary.md").is_file())
            self.assertTrue((output / "manual-review-checklist.md").is_file())
            self.assertIn("not independently verified", (output / "summary.md").read_text(encoding="utf-8"))
            self.assertIn("Manual Review Checklist", (output / "manual-review-checklist.md").read_text(encoding="utf-8"))

    def test_explicit_chinese_language_applies_to_all_report_files(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [["Group A", "Group B"], [1.2, 2.3], [3.4, 4.5]])
            findings, coverage = audit_module.audit_file(path)
            output = Path(folder) / "reports-zh"
            audit_module.write_reports(path, findings, coverage, output, lang="zh")
            self.assertIn('lang="zh-CN"', (output / "report.html").read_text(encoding="utf-8"))
            self.assertIn("人工复核清单", (output / "manual-review-checklist.md").read_text(encoding="utf-8"))
            self.assertIn("未独立验证", (output / "summary.md").read_text(encoding="utf-8"))

    def test_language_flag_is_exposed_by_cli(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [["A", "B"], [1.2, 1.3], [2.1, 2.4]])
            target = Path(folder) / "out"
            audit_module.main([str(path), "--out-dir", str(target), "--lang", "zh"])
            self.assertIn('lang="zh-CN"', (target / "report.html").read_text(encoding="utf-8"))

    def test_explicit_welch_ttest_detects_incorrect_reported_p(self):
        try:
            from openpyxl import Workbook
            from scipy import stats  # noqa: F401
        except ImportError:
            self.skipTest("openpyxl and scipy are required for t-test verification")
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "p_values.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Figure 2"
            for idx, (a, b) in enumerate(zip([1, 2, 3, 4, 5], [5, 6, 7, 8, 9]), 2):
                ws[f"B{idx}"] = a
                ws[f"C{idx}"] = b
            ws["E2"] = 0.8
            wb.save(path)
            config = {"ttests": [{"sheet": "Figure 2", "a": "B2:B6", "b": "C2:C6",
                                    "design": "welch", "reported_p": "E2"}]}
            findings, coverage = audit_module.audit_file(path, config=config)
            self.assertIn("reported_p_mismatch", {f["rule"] for f in findings})
            self.assertEqual(coverage["configured_ttests_checked"], 1)

    def test_numeric_sample_ids_are_not_flagged_as_artificial_progressions(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [
                ["Mouse ID", "GFAP intensity"],
                [1, 1.23], [2, 3.44], [3, 2.15], [4, 5.66], [5, 2.88], [6, 1.79],
            ])
            findings, _ = audit_module.audit_file(path)
            self.assertFalse(any(f["rule"] == "arithmetic_progression" and "A" in f["cells"]
                                 for f in findings))

    def test_formula_without_cached_value_is_explained(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            self.skipTest("openpyxl is not installed")
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "formula.xlsx"
            wb = Workbook()
            ws = wb.active
            ws["A1"] = 1.5
            ws["A2"] = 2.5
            ws["A3"] = "=AVERAGE(A1:A2)"
            wb.save(path)
            findings, coverage = audit_module.audit_file(path)
            self.assertEqual(coverage["uncached_formula_cells"], 1)
            self.assertTrue(any(f["rule"] == "uncached_formula" and f["cells"] == "A3"
                                for f in findings))

    def test_long_matching_distinct_marker_columns_get_priority_with_actual_cells(self):
        with tempfile.TemporaryDirectory() as folder:
            observations = [11.44, 15.30, 16.45, 22.56, 26.30, 10.30, 29.90, 16.34, 22.10]
            path = self.make_csv(folder, [["Marker A", "Marker B", "Marker C", "Marker D"]] +
                                 [[98 + k / 10, val, 5 + k, val]
                                  for k, val in enumerate(observations)])
            findings, _ = audit_module.audit_file(path)
            matches = [x for x in findings if x["rule"] == "identical_numeric_columns"]
            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0]["priority"], "P1")
            self.assertIn("B2", matches[0]["cells"])
            self.assertIn("D10", matches[0]["cells"])

    def test_integer_count_profiles_do_not_trigger_high_priority_column_duplication(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.make_csv(folder, [["DRD1", "DRD2"]] +
                                 [[i % 9, i % 9] for i in range(100)])
            findings, _ = audit_module.audit_file(path)
            self.assertFalse(any(x["rule"] in ("identical_numeric_columns", "duplicate_numeric_block")
                                 for x in findings))

    def test_many_identical_decimal_columns_are_bounded_with_coverage_warning(self):
        with tempfile.TemporaryDirectory() as folder:
            values = [11.11, 12.22, 13.33, 14.44, 15.55, 16.66, 17.77, 18.88, 19.99]
            path = self.make_csv(folder, [[f"Marker{i}" for i in range(70)]
                                 + ["Label"]] + [[value] * 70 + [f"S{j}"]
                                              for j, value in enumerate(values)])
            findings, _ = audit_module.audit_file(path)
            self.assertTrue(any(x["rule"] == "duplicate_pair_search_limited" for x in findings))
            self.assertLessEqual(sum(x["rule"] == "identical_numeric_columns" for x in findings), 50)

    def test_partial_long_decimal_block_is_located(self):
        with tempfile.TemporaryDirectory() as folder:
            source = [10.11, 11.22, 12.33, 13.44, 14.55, 15.66, 16.77]
            a = [2.41] + source + [0.41]
            b = [9.23] + source + [0.93]
            path = self.make_csv(folder, [["A", "B"]] + [[x, y] for x, y in zip(a, b)])
            findings, _ = audit_module.audit_file(path)
            matches = [x for x in findings if x["rule"] == "duplicate_numeric_block"]
            self.assertEqual(len(matches), 1)
            self.assertIn("A3", matches[0]["cells"])
            self.assertIn("B9", matches[0]["cells"])

    def test_legacy_xls_reader_uses_xlrd_numeric_cells_only(self):
        from unittest.mock import patch

        class Cell:
            def __init__(self, value, ctype):
                self.value, self.ctype = value, ctype

        class Sheet:
            name, nrows, ncols = "Figure X", 2, 3

            def cell(self, row, col):
                return [[Cell("Marker", 1), Cell("Date", 1), Cell("Value", 1)],
                        [Cell("IT", 1), Cell(45000, 3), Cell(11.44, 2)]][row][col]

        class Book:
            def sheets(self):
                return [Sheet()]

        try:
            import xlrd  # noqa: F401
        except ImportError:
            self.skipTest("xlrd is optional")
        with patch("xlrd.open_workbook", return_value=Book()):
            grids, uncached = audit_module.load_grids(Path("dummy.xls"))
        self.assertEqual(grids["Figure X"][1][1], None)
        self.assertEqual(grids["Figure X"][1][2], 11.44)
        self.assertEqual(uncached, [])


if __name__ == "__main__":
    unittest.main()
