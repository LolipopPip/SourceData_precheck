"""A reviewer must be able to open every issue offline at its source cells."""

import tempfile
import unittest
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SPEC = spec_from_file_location("report_html", ROOT / "scripts" / "report_html.py")


class HtmlReportTests(unittest.TestCase):
    def setUp(self):
        self.renderer = module_from_spec(SPEC)
        SPEC.loader.exec_module(self.renderer)

    def test_each_chinese_issue_links_to_highlighted_excel_cells_and_escapes_text(self):
        findings = [{
            "rule": "identical_numeric_columns", "sheet": "Figure S1I", "priority": "P1",
            "severity": "review", "cells": "B2, D2, B3, D3",
            "evidence": "Two columns match.", "next_step": "Check counts.",
        }]
        grids = {"Figure S1I": [
            ["Marker A", "Marker B", "Marker C", "Marker D", "<img src=x onerror=alert(1)>"],
            [98.6, 11.44, 5.2, 11.44, ""],
            [97.67, 15.3, 12.34, 15.3, ""],
        ]}
        html = self.renderer.render_html("Example.xls", findings, grids, {}, lang="zh")
        self.assertIn("完整重复列", html)
        self.assertIn("Figure S1I", html)
        self.assertIn('href="#issue-1"', html)
        self.assertIn('id="issue-1"', html)
        self.assertIn("B2", html)
        self.assertIn("D3", html)
        self.assertIn('class="flagged"', html)
        self.assertIn("&lt;img src=x onerror=alert(1)&gt;", html)
        self.assertNotIn("<img src=x onerror=alert(1)>", html)

    def test_empty_findings_still_report_non_verification(self):
        html = self.renderer.render_html("Source.csv", [], {"Fig 1": [["A"], [1.2]]}, {}, lang="zh")
        self.assertIn("未发现自动规则命中", html)
        self.assertIn("不代表数据已经通过真实性验证", html)

    def test_preview_outside_sheet_does_not_crash(self):
        findings = [{"rule": "internal_missing_numeric_cell", "sheet": "Fig 1",
                     "cells": "B100", "priority": "P2", "severity": "review"}]
        html = self.renderer.render_html("Source.csv", findings, {"Fig 1": [[1.2]]}, {}, lang="zh")
        self.assertIn("工作表", html)
        self.assertIn("B100", html)

    def test_later_p1_is_open_and_navigation_resets_filters(self):
        findings = [
            {"rule": "internal_missing_numeric_cell", "sheet": "Fig 1", "cells": "A2",
             "priority": "P2", "severity": "review"},
            {"rule": "identical_numeric_columns", "sheet": "Fig 2", "cells": "A2, B2",
             "priority": "P1", "severity": "review"},
        ]
        html = self.renderer.render_html("Source.csv", findings, {"Fig 1": [[1.2], [2.4]],
                                                              "Fig 2": [[1.1, 1.1], [2.2, 2.2]]}, {}, lang="zh")
        card = html.split('id="issue-2"', 1)[1].split("</details>", 1)[0]
        self.assertIn('data-priority="P1"', card)
        self.assertIn(' open>', card)
        self.assertIn("search.value=''; sheet.value=''; priority.value=''; apply();", html)

    def test_repeated_columns_title_never_assumes_independent_marker_identity(self):
        title = self.renderer.RULES_ZH["identical_numeric_columns"][0]
        self.assertNotIn("不同指标", title)

    def test_cross_sheet_repeat_previews_both_original_sheets_and_highlights(self):
        issue = {"rule": "cross_panel_repeated_series", "sheet": "Panel Alpha",
                 "cells": "D3, D4, D5", "related_sheet": "Panel Beta",
                 "related_cells": "D8, D9, D10", "priority": "P2",
                 "evidence": "3 exact values", "severity": "review"}
        first = [[None] * 4 for _ in range(5)]
        second = [[None] * 4 for _ in range(10)]
        for idx, val in enumerate((28, 71, 94)):
            first[idx + 2][3] = val
            second[idx + 7][3] = val
        first[1][3] = "3.2 mg/mL"
        second[6][3] = "3.2 mg/mL"
        html = self.renderer.render_html("Synthetic.xlsx", [issue],
            {"Panel Alpha": first, "Panel Beta": second}, {}, lang="zh")
        card = html.split('id="issue-1"', 1)[1].split("</details>", 1)[0]
        self.assertIn("Panel Alpha!D3", card)
        self.assertIn("Panel Beta!D8", card)
        self.assertIn('title="D3"', card)
        self.assertIn('title="D8"', card)
        self.assertEqual(card.count('class="flagged"'), 6)

    def test_default_report_is_english_even_for_cross_sheet_findings(self):
        issue = {"rule": "cross_panel_repeated_series", "sheet": "Panel A",
                 "cells": "A2, A3, A4", "related_sheet": "Panel B",
                 "related_cells": "A2, A3, A4", "priority": "P2",
                 "evidence": "Three values match", "severity": "review"}
        html = self.renderer.render_html("Example.xlsx", [issue],
            {"Panel A": [["Count"], [28], [71], [94]],
             "Panel B": [["Count"], [28], [71], [94]]}, {})
        self.assertIn('lang="en"', html)
        self.assertIn("Source Data Integrity Audit", html)
        self.assertIn("Matched source location", html)
        self.assertNotIn("为什么会出现", html)
