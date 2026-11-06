"""Standalone, escaped, bilingual HTML review of Source Data numeric audit findings.

The output embeds small excerpts around flagged A1 coordinates. It never fetches
remote resources and does not alter the input workbook.
"""

from __future__ import annotations

import html
import re
from collections import Counter
from hashlib import sha1


RULES_ZH = {
    "cross_panel_repeated_series": (
        "跨工作表 / 跨数据区域的数值重复",
        "两处数据区域存在按顺序完全相同的观测片段。程序过滤了常见编号与设计梯度；短整数匹配只有在实验分组标签相同时才提示，仍可能是巧合或合法数据复用。",
        "对照两处原始样本 ID、原始采集记录、分组标签及分析公式，确认是否属于同一来源的合法重复呈现。",
    ),
    "cross_panel_search_limited": (
        "跨数据区域重复检索的覆盖提醒",
        "部分三值指纹在太多列中重复，为控制组合数量，程序跳过了这些高度常见的指纹。这是检测覆盖限制，不代表数值异常。",
        "对大规模重复列、共享对照或宽表，必要时人工分面板检查。",
    ),
    "identical_numeric_columns": (
        "完整重复列：两列数值序列完全相同",
        "两列在同一批行中出现整段相同的非整数测量值。若列标题代表不同 marker、实验组或独立观测，应优先核对原始记录；重复本身不是科研不端证据。",
        "核对两列列名、样本 ID、原始计数/仪器导出、计算公式和 Excel 录入历史，排除同一来源复用或整列误粘贴。",
    ),
    "duplicate_numeric_block": (
        "长连续数值片段重复",
        "不同列有一段较长的、变化充分的小数完全相同。算法已过滤短整数与零值为主的片段，但仍可能是共享分母、嵌套细胞群或衍生指标。",
        "确认相同值是否来自独立样本和独立测量；比对原始计数、表头及公式，不以程序分级代替判断。",
    ),
    "duplicate_numeric_row": (
        "多列数值整行相同",
        "有两行或多行出现相同的数值记录。时间点、技术重复和受检测限约束的记录可能自然相同。",
        "查阅行标签、样本 ID、时间轴和原始导出，判断是否重复呈现同一来源。",
    ),
    "duplicate_high_precision_value": (
        "高精度数值重复",
        "同一列有至少两次相同的多位小数。百分比与整数计数换算时，这种情况也可能正常。",
        "将每个重复值追溯到原始分子/分母或仪器读数，核对是否有共同计算输入。",
    ),
    "internal_missing_numeric_cell": (
        "数值区内部存在空格",
        "同列上下有测量值，中间的单元格为空或非数值。可能表示未测量、不适用或无缓存公式。",
        "查看该列对应指标的 n、缺失原因和图注；不要自动填零。",
    ),
    "constant_column_offset": (
        "两列保持固定差值",
        "同一区域两列在多个观测之间的差值完全不变，可能来自规范化、衍生计算或数据录入。",
        "核对标准化公式、处理前数值和处理组设计。",
    ),
    "arithmetic_progression": (
        "连续测量呈严格等差",
        "连续数值间隔完全相等；若为预设剂量、时间或刻度则属于正常设计。",
        "首先确认这一列是否为人为设置的横坐标、时间点或处理浓度。",
    ),
    "last_digit_concentration": (
        "存储小数末位集中",
        "某一存储小数末位过度集中；四舍五入、导出精度、仪器分辨率都可能解释这一现象。",
        "核对 Excel 显示格式、仪器精度、计算方式；该提示不能推断真实性。",
    ),
    "wide_table_offset_check_skipped": (
        "宽表检测覆盖提醒",
        "该区域数值列较多，部分两两固定差值规则为避免组合爆炸而跳过。这不是数据异常。",
        "需要时按 Figure/panel 划定小范围再运行固定差值分析。",
    ),
    "duplicate_pair_search_limited": (
        "重复列组合数量过多，部分扫描受限",
        "找到大量长序列相同的列组合，为保证程序在宽表上能够完成，报告只展示前 50 组并明确提示覆盖范围。此项本身不是数据问题。",
        "按 Figure/panel 缩小列范围，必要时逐组核对其余重复列。",
    ),
    "uncached_formula": (
        "公式缺少可读取的缓存结果",
        "Excel 中有公式，但只读程序没有取到计算结果。不能把它当成已经复核的数值。",
        "在 Excel 中重算并另存副本，再对副本重新审计，保留原文件。",
    ),
    "count_identity_mismatch": (
        "计数恒等式不一致",
        "在提供的统计映射假设下，总数与指定子类别之和不相等。",
        "核对子类别是否互斥、是否有遗漏项及统计口径；保持原始计数不变。",
    ),
    "reported_ratio_mismatch": (
        "报告比例与复算结果不符",
        "根据确认的分子、分母和比例尺度复算后，结果与报告单元格不一致。",
        "核实分母定义、比例尺度、数值精度及公式。",
    ),
    "reported_p_mismatch": (
        "报告 P 值与指定检验复算不符",
        "依据用户映射的统计检验与分组计算，报告 P 值不同。仍须确认配对关系、多重检验和样本单位。",
        "核对 GraphPad/R 设置、检验方向、独立重复单位及校正方法。",
    ),
}

RULES_EN = {
    "cross_panel_repeated_series": (
        "Repeated values across panels or data regions",
        "Two disjoint regions contain the same consecutive observations in the same order. Common sample indexes and designed gradients were screened out; short integer matches require the same informative group label and may still occur by chance or through valid reuse.",
        "Compare sample identifiers, original acquisition records, group labels and formulas for both locations; document any legitimate shared source."),
    "cross_panel_search_limited": (
        "Cross-panel search coverage limitation",
        "Some common three-value fingerprints occurred too often for an exhaustive comparison; these matches were skipped to bound processing. This is not evidence of an anomaly.",
        "Review heavily repeated tables or shared controls manually if necessary."),
    "identical_numeric_columns": (
        "Identical complete numeric columns",
        "Two columns share a long, row-aligned run of matching, varied decimal observations. Confirm whether their underlying measurements are expected to be independent.",
        "Check labels, sample identifiers, raw exports, formulas and data-entry history."),
    "duplicate_numeric_block": (
        "Repeated continuous numeric block",
        "Two regions contain an identical, sufficiently long and varied decimal sequence. Shared denominators, nested populations or derived metrics may explain it.",
        "Compare original samples, raw values, headers and formulas before interpreting the signal."),
    "duplicate_numeric_row": (
        "Identical numeric rows",
        "Multiple rows contain the same measurement record; technical replicates or time points can legitimately coincide.",
        "Check row labels, sample IDs, timing and acquisition records."),
    "duplicate_high_precision_value": (
        "Repeated high-precision value",
        "The same precise decimal occurs more than once in a column; calculations from shared inputs can produce this naturally.",
        "Trace each value to its raw measurement or numerator and denominator."),
    "internal_missing_numeric_cell": (
        "Blank inside a numeric region",
        "A cell between numeric observations is blank or nonnumeric. This may mean not measured, not applicable or missing a cached formula result.",
        "Check sample size, exclusions, missingness and the figure legend; never fill blanks with zeros automatically."),
    "constant_column_offset": (
        "Constant offset between columns",
        "Two columns differ by the same amount across multiple observations; normalization or derived calculations may explain this.",
        "Review transformation formulas and the original measurements."),
    "arithmetic_progression": (
        "Exact arithmetic progression",
        "Consecutive numeric values have an identical increment; planned doses, time points and axis ticks are normal examples.",
        "Confirm whether this is a designed independent variable rather than a measured outcome."),
    "last_digit_concentration": (
        "Concentrated trailing decimal digits",
        "Stored decimal endings are unusually concentrated. Export precision, rounding and instrument resolution may explain this.",
        "Inspect Excel display format and measurement or calculation precision."),
    "wide_table_offset_check_skipped": (
        "Wide-table comparison coverage limitation",
        "Some pairwise constant-offset checks were skipped in a very wide table to bound processing; this is not an anomaly.",
        "If needed, inspect narrower panel-specific ranges manually."),
    "duplicate_pair_search_limited": (
        "Too many repeated column pairs to display",
        "Many long matching column pairs were found; only the first 50 are shown to bound processing. Remaining pairs have not been individually reviewed.",
        "Split the table by panel and verify additional matching columns."),
    "uncached_formula": (
        "Formula has no readable cached value",
        "A spreadsheet formula exists, but the read-only loader could not obtain its calculated value; it has not been verified.",
        "Recalculate and save a separate workbook copy in Excel, then audit that copy."),
    "count_identity_mismatch": (
        "Count identity mismatch",
        "Under the explicitly provided mapping, a reported total differs from the sum of its designated components.",
        "Confirm categories are mutually exclusive and exhaustive; retain original counts."),
    "reported_ratio_mismatch": (
        "Reported ratio differs from recalculation",
        "The reported proportion differs from the result computed with the provided numerator, denominator and scale.",
        "Check denominator definitions, scale, rounding and formulas."),
    "reported_p_mismatch": (
        "Reported P value differs from configured test",
        "A P value differs from the specified test on the mapped observations. Pairing, experimental units and multiplicity corrections still require confirmation.",
        "Confirm test design, direction, assumptions, corrections and plotting-software settings."),
}

UI_ZH = {
    "empty_table": "原表格为空，暂无单元格预览。", "no_coordinates": "此规则未提供可定位的 A1 单元格坐标，请查看上方位置说明。",
    "invalid_coordinates": "没有可解析的 A1 单元格。", "out_of_bounds": "坐标超出当前工作表可读取的行范围；请按上方地址在原始文件中核对。",
    "row_col": "行 / 列", "rows_skipped": "省略中间未显示的行", "excerpt_truncated": "为保证页面可读，仅展示部分行列；完整位置参见上方单元格列表。",
    "statistics_mismatch": "报告统计量与复算结果不符", "statistics_reason": "在用户提供的样本映射和舍入假设下，报告值与程序复算不一致。",
    "statistics_action": "重新核对原始观察值、样本量、标准差定义及报告精度。", "test_unchecked": "统计检验暂未完成",
    "test_reason": "部分统计检验无法读取或计算，不能视为已验证通过。", "test_action": "确认 scipy、样本设计及有效数值后再次复核。",
    "generic_title": "待人工确认的数值提示", "generic_reason": "本条命中的是自动筛查规则，无法脱离原始实验背景独立定性。",
    "generic_action": "核对对应图注、样本 ID 和原始记录。", "P1": "优先核对", "P2": "常规复核", "P3": "低特异性 / 信息",
    "manual_review": "人工确认", "why": "为什么会出现这条提示？", "source_location": "原始位置：", "matched_location": "匹配位置：",
    "sheet_summary": "查看工作表汇总", "evidence": "检测依据：", "action": "建议怎样复核：", "original_excerpt": "原始表格片段",
    "matched_excerpt": "第二处原始表格片段：", "yellow_note": "黄色为匹配单元格", "yellow_flag": "黄色为本条命中的单元格",
    "original_rule": "程序原始规则：", "severity": "严重度：", "checked_label": "此条已完成人工复核",
    "checked_text": "我已在原始实验记录中核实此条（仅当前页面勾选）", "no_findings": "未发现自动规则命中。请注意：这不代表数据已经通过真实性验证，也不表示原始记录、图表统计或样本独立性已经完成复核。",
    "items": " 条", "no_sheets": "本次没有需要定位的工作表。", "all_sheets": "全部工作表",
    "expand": "展开 ▾", "collapse": "收起 ▴", "visible_prefix": "当前显示 ", "of": " / ",
    "title": "Source Data 数值完整性自查", "page_title": "Source Data 完整性审查", "file": "文件：",
    "subtitle": "自动检测结果可逐条展开查看原表中的具体数值与单元格坐标。此页面为离线报告，不向任何服务器发送数据。",
    "sheet_index": "工作表索引", "quick_jump": "异常快速跳转", "sheets_scanned": "已扫描工作表", "findings": "自动规则提示",
    "p1_count": "P1 优先核查", "p2_count": "P2 常规复核", "notice_heading": "阅读前请注意：",
    "notice": "自动筛查提示不是造假结论。即使没有提示，也不代表数据已经通过真实性验证。统计均值、SEM、n、P 值等只有在提供明确单元格映射时才可复算；本报告不校验图像或原始实验身份。",
    "filters": "筛选异常", "search": "搜索 Figure、规则或 A1 单元格…", "filter_sheet": "按工作表筛选", "filter_priority": "按复核级别筛选",
    "all_priorities": "全部优先级", "p3_count": "P3 低特异性/信息", "total_items": "共", "footer": "数值筛查与人工复核辅助 · 原始 Source Data 文件未修改 · 不得据此推断科研不端",
}

UI_EN = {
    "empty_table": "The source table is empty; no cell preview is available.", "no_coordinates": "This check has no parseable A1 cell address. See the location above.",
    "invalid_coordinates": "No valid A1 cell coordinates were found.", "out_of_bounds": "The location is outside the readable worksheet; verify the original address in the workbook.",
    "row_col": "Row / Col", "rows_skipped": "Intermediate rows omitted", "excerpt_truncated": "Only selected rows and columns are shown for readability; use the addresses above for the complete locations.",
    "statistics_mismatch": "Reported statistic differs from recalculation", "statistics_reason": "The reported value differs under the provided sample mapping and rounding assumptions.",
    "statistics_action": "Check original observations, sample size, SD definition and reporting precision.", "test_unchecked": "Statistical test not verified",
    "test_reason": "Some statistical tests could not be loaded or computed and must not be treated as verified.", "test_action": "Check scipy, study design and valid observations before retrying.",
    "generic_title": "Numeric finding requiring manual review", "generic_reason": "An automated pattern check was triggered; its meaning depends on the experimental context.",
    "generic_action": "Check the figure legend, sample IDs and original records.", "P1": "High-priority review", "P2": "Manual review", "P3": "Low specificity / info",
    "manual_review": "Manual review", "why": "Why was this finding flagged?", "source_location": "Source location:", "matched_location": "Matched source location:",
    "sheet_summary": "View worksheet summary", "evidence": "Detection evidence:", "action": "Suggested verification:", "original_excerpt": "Original worksheet excerpt",
    "matched_excerpt": "Second worksheet excerpt:", "yellow_note": "Yellow indicates matched cells", "yellow_flag": "Yellow indicates flagged cells",
    "original_rule": "Detector rule:", "severity": "Severity:", "checked_label": "Manual review complete",
    "checked_text": "I verified this finding against original records (checkbox applies only to this page)", "no_findings": "No automated rules were triggered. This does not establish authenticity or verify the original records, plotted statistics or sample independence.",
    "items": " findings", "no_sheets": "No worksheet locations require attention.", "all_sheets": "All worksheets",
    "expand": "Expand ▾", "collapse": "Collapse ▴", "visible_prefix": "Showing ", "of": " / ",
    "title": "Source Data Integrity Audit", "page_title": "Source Data Integrity Audit", "file": "File:",
    "subtitle": "Expand individual findings to inspect source values and exact cell addresses. This offline report sends no data to any server.",
    "sheet_index": "Worksheet index", "quick_jump": "Jump to findings", "sheets_scanned": "Worksheets scanned", "findings": "Automated findings",
    "p1_count": "P1 priority", "p2_count": "P2 review", "notice_heading": "Before interpreting results:",
    "notice": "Automated findings are not evidence of fabrication. No findings does not establish authenticity. Mean/SEM/n/P values are recalculated only when explicit cell mappings are provided. Image integrity and experimental provenance are outside this audit.",
    "filters": "Filter findings", "search": "Search figure, rule or A1 address…", "filter_sheet": "Filter by worksheet", "filter_priority": "Filter by review level",
    "all_priorities": "All priorities", "p3_count": "P3 informational", "total_items": "Total:", "footer": "Read-only numeric screening with manual verification · Original Source Data not modified · No misconduct inference",
}


def esc(value):
    return html.escape(str(value), quote=True)


def safe_id(sheet):
    return "sheet-" + sha1(sheet.encode("utf-8")).hexdigest()[:12]


def _cells_from_issue(issue):
    coords = set()
    for col, row in re.findall(r"(?<![A-Za-z0-9])([A-Z]{1,3})([1-9][0-9]*)(?![A-Za-z0-9])",
                                str(issue.get("cells", "")).upper()):
        index = 0
        for char in col:
            index = index * 26 + ord(char) - ord("A") + 1
        coords.add((int(row), index))
    return coords


def _col_label(n):
    chars = ""
    while n:
        n, rem = divmod(n - 1, 26)
        chars = chr(65 + rem) + chars
    return chars


def _preview(rows, issue, lang="en"):
    t = UI_ZH if lang == "zh" else UI_EN
    coords = _cells_from_issue(issue)
    if not rows:
        return f'<p class="muted">{t["empty_table"]}</p>'
    if not coords:
        return f'<p class="muted">{t["no_coordinates"]}</p>'

    # Keep the actual flagged columns and their adjacent labels without rendering huge areas.
    selected_rows = sorted({r for r, _ in coords if r >= 1})
    selected_cols = sorted({c for _, c in coords if c >= 1})
    if not selected_rows or not selected_cols:
        return f'<p class="muted">{t["invalid_coordinates"]}</p>'
    if selected_rows[0] > len(rows):
        return f'<p class="muted">{t["out_of_bounds"]}</p>'
    first_row = max(1, selected_rows[0] - 1)
    last_row = min(len(rows), selected_rows[-1] + 1)
    shown_rows = list(range(first_row, min(last_row, first_row + 15) + 1))
    if selected_rows[-1] > shown_rows[-1]:
        shown_rows = shown_rows[:11] + [r for r in range(max(1, selected_rows[-1] - 2),
                                                        min(len(rows), selected_rows[-1] + 1) + 1)
                                        if r not in shown_rows[:11]]

    nearest = set()
    for col in selected_cols[:20]:
        nearest.update([max(1, col - 1), col, col + 1])
    if min(selected_cols) > 3:
        nearest.add(1)  # panel names often occupy the first column
    shown_cols = sorted(nearest)
    if len(shown_cols) > 13:
        shown_cols = sorted(set(selected_cols[:9]) | {1, *selected_cols[-3:]})

    parts = ['<div class="table-scroll"><table class="excerpt"><thead><tr>',
             f'<th class="corner">{t["row_col"]}</th>']
    parts += [f'<th>{esc(_col_label(c))}</th>' for c in shown_cols]
    parts.append('</tr></thead><tbody>')
    prev = None
    for ridx in shown_rows:
        if prev is not None and ridx > prev + 1:
            parts.append(f'<tr><th>…</th><td colspan="{len(shown_cols)}" class="omitted">{t["rows_skipped"]}</td></tr>')
        parts.append(f'<tr><th class="rownum">{ridx}</th>')
        row = rows[ridx - 1] if ridx <= len(rows) else []
        for col in shown_cols:
            val = row[col - 1] if col <= len(row) else ""
            raw = "" if val is None else str(val)
            if len(raw) > 120:
                raw = raw[:117] + "…"
            a1 = f"{_col_label(col)}{ridx}"
            marker = ' class="flagged"' if (ridx, col) in coords else ""
            parts.append(f'<td{marker} title="{esc(a1)}"><small>{esc(a1)}</small><span>{esc(raw)}</span></td>')
        parts.append('</tr>')
        prev = ridx
    parts.append('</tbody></table></div>')
    if len(selected_rows) > len(shown_rows) or len(selected_cols) > len(shown_cols):
        parts.append(f'<p class="muted">{t["excerpt_truncated"]}</p>')
    return "".join(parts)


def _meaning(issue, lang="en"):
    rules = RULES_ZH if lang == "zh" else RULES_EN
    t = UI_ZH if lang == "zh" else UI_EN
    rule = str(issue.get("rule", ""))
    if rule in rules:
        return rules[rule]
    if rule.startswith("reported_") and rule.endswith("_mismatch"):
        return (t["statistics_mismatch"], t["statistics_reason"], t["statistics_action"])
    if rule.startswith("ttest_"):
        return (t["test_unchecked"], t["test_reason"], t["test_action"])
    return (t["generic_title"], t["generic_reason"], t["generic_action"])


def render_html(input_name, findings, grids, coverage, lang="en"):
    """Render complete standalone UTF-8 HTML, one expanded/linked issue per finding."""
    if lang not in {"en", "zh"}:
        raise ValueError("Supported report languages: en, zh")
    t = UI_ZH if lang == "zh" else UI_EN
    counts = Counter(str(f.get("priority", "P2")) for f in findings)
    by_sheet = Counter(str(f.get("sheet", "")) for f in findings)
    by_sheet.update(str(f["related_sheet"]) for f in findings if f.get("related_sheet")
                    and f["related_sheet"] != f.get("sheet"))
    cards = []
    shortcuts = []
    first_p1 = next((i for i, item in enumerate(findings, 1)
                     if item.get("priority") == "P1"), 1)
    for n, issue in enumerate(findings, 1):
        sheet = str(issue.get("sheet", ""))
        related_sheet = str(issue.get("related_sheet", ""))
        related_coords = str(issue.get("related_cells", ""))
        priority = str(issue.get("priority", "P2"))
        priority_label = t.get(priority, t["manual_review"])
        title, explanation, action = _meaning(issue, lang)
        coords = str(issue.get("cells", ""))
        shortcuts.append(f'<a href="#issue-{n}"><span class="dot {esc(priority)}"></span>'
                         f'<strong>{n:03d}</strong> {esc(title)} <em>{esc(sheet)}'
                         + (f' ↔ {esc(related_sheet)}' if related_sheet else '') + '</em></a>')
        preview = _preview(grids.get(sheet, []), issue, lang)
        paired_preview = (f'<h4>{t["matched_excerpt"]} {esc(related_sheet)} <span>{t["yellow_note"]}</span></h4>'
                          + _preview(grids.get(related_sheet, []), {"cells": related_coords}, lang)
                          if related_sheet else "")
        paired_position = (f'<span><b>{t["matched_location"]}</b><code>{esc(related_sheet)}!{esc(related_coords)}</code></span>'
                           if related_sheet else "")
        opened = " open" if n == first_p1 else ""
        cards.append(f'''<details id="issue-{n}" class="issue" data-sheet="{esc(sheet)}"
          data-related-sheet="{esc(related_sheet)}"
          data-priority="{esc(priority)}" data-search="{esc((sheet + ' ' + related_sheet + ' ' + title + ' ' + coords + ' ' + related_coords).lower())}"{opened}>
          <summary><span class="number">#{n:03d}</span><span class="pill {esc(priority)}">{esc(priority)} · {esc(priority_label)}</span>
          <span class="heading">{esc(title)}</span><span class="sheet">{esc(sheet)}</span></summary>
          <div class="issue-body"><div class="callout"><strong>{t["why"]}</strong><p>{esc(explanation)}</p></div>
          <div class="meta"><span><b>{t["source_location"]}</b><code>{esc(sheet)}!{esc(coords)}</code></span>
          {paired_position}
          <a href="#{safe_id(sheet)}">{t["sheet_summary"]}</a></div>
          <p class="fine"><strong>{t["evidence"]}</strong>{esc(issue.get('evidence',''))}</p>
          <p class="action"><strong>{t["action"]}</strong>{esc(action)}</p>
          <h4>{t["original_excerpt"]} <span>{t["yellow_flag"]}</span></h4>{preview}{paired_preview}
          <div class="tiny">{t["original_rule"]}<code>{esc(issue.get('rule',''))}</code> · {t["severity"]}{esc(issue.get('severity','review'))}</div>
          <label class="done"><input type="checkbox" aria-label="{t["checked_label"]}"> {t["checked_text"]}</label>
          </div></details>''')

    if not findings:
        cards.append(f'<section class="empty">{t["no_findings"]}</section>')
    sheet_navigation = []
    for sheet, total in sorted(by_sheet.items(), key=lambda item: (-item[1], item[0])):
        sid = safe_id(sheet)
        sheet_navigation.append(f'<a class="sheet-link" id="{sid}" href="#issue-'
                                f'{next(i for i, f in enumerate(findings, 1) if f.get("sheet") == sheet or f.get("related_sheet") == sheet)}">'
                                f'{esc(sheet)}<span>{total}{t["items"]}</span></a>')
    if not sheet_navigation:
        sheet_navigation = [f'<span class="muted">{t["no_sheets"]}</span>']
    options = [f'<option value="">{t["all_sheets"]}</option>'] + [
        f'<option value="{esc(sheet)}">{esc(sheet)}（{total}）</option>'
        for sheet, total in sorted(by_sheet.items())]
    sheets_scanned = coverage.get("sheets_scanned", len(grids))

    style = r'''
    :root{font-family:ui-sans-serif,system-ui,-apple-system,"Microsoft YaHei",sans-serif;color:#253044;background:#f5f7fb}
    *{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0}a{color:#3157a3;text-decoration:none}
    a:hover{text-decoration:underline}button,select,input{font:inherit}header{background:linear-gradient(122deg,#102b4f,#315a8a);color:white;padding:42px max(24px,calc((100vw - 1230px)/2)) 34px}
    .eyebrow{font-size:12px;letter-spacing:.15em;text-transform:uppercase;color:#bed8ff}.main-title{font-size:clamp(26px,3.4vw,40px);font-weight:750;margin:10px 0}
    .subtitle{color:#dce8fa;max-width:900px;line-height:1.9;margin:0}.layout{max-width:1280px;margin:auto;padding:28px 24px;display:grid;grid-template-columns:260px minmax(0,1fr);gap:24px}
    aside{align-self:start;position:sticky;top:18px;max-height:calc(100vh - 35px);overflow:auto}.panel{background:white;border:1px solid #e5e9f2;box-shadow:0 5px 20px #112f5c0c;border-radius:14px;padding:18px}
    h2{font-size:19px;margin:0 0 14px}h3{font-size:16px}.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:18px}
    .stat{background:#fff;border:1px solid #e1e6f1;border-radius:12px;padding:17px}.stat strong{font-size:29px;display:block;color:#203a63}.stat small{color:#596a84}
    .note{background:#fff9e9;border-left:4px solid #e9b33d;padding:13px 16px;border-radius:8px;font-size:14px;line-height:1.85;margin:18px 0}
    .filters{display:flex;flex-wrap:wrap;gap:10px;margin:18px 0} .filters input,.filters select{background:white;padding:11px 12px;border:1px solid #cbd5e4;border-radius:10px;min-width:150px;max-width:100%}
    .filters input{flex:1;min-width:230px}.issue{background:white;border:1px solid #dfe5ef;border-radius:12px;margin:12px 0;scroll-margin-top:16px;overflow:hidden}
    .issue[open]{border-color:#91abd6;box-shadow:0 8px 26px #143c7b12}.issue>summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:11px;padding:17px;flex-wrap:wrap}
    .issue>summary::-webkit-details-marker{display:none}.issue>summary:after{content:'__EXPAND__';font-size:12px;color:#55729c;margin-left:auto}
    .issue[open]>summary:after{content:'__COLLAPSE__'}.issue-body{padding:0 19px 20px}.number{color:#718199;font-size:12px;font-weight:700}.heading{font-weight:700;flex:1;min-width:160px}.sheet{color:#6a7790;font-size:13px}
    .pill{font-size:12px;font-weight:700;padding:4px 9px;border-radius:999px;white-space:nowrap}.pill.P1{background:#ffe4e4;color:#9d1f2e}.pill.P2{background:#fff0d6;color:#946120}.pill.P3{background:#e9f0fb;color:#335a8d}
    .callout{background:#edf4ff;border-radius:10px;padding:14px 16px;line-height:1.8}.callout p{margin:4px 0 0}.meta{display:flex;flex-wrap:wrap;justify-content:space-between;gap:12px;margin:14px 0;line-height:1.6}
    code{font-family:ui-monospace,SFMono-Regular,Consolas,monospace;background:#f0f3f8;padding:2px 4px;border-radius:4px;overflow-wrap:anywhere}
    .fine,.action{font-size:14px;line-height:1.8}.action{color:#265c51}.issue h4{font-size:15px;margin:19px 0 9px}.issue h4 span{font-weight:400;font-size:12px;color:#7b879c}
    .table-scroll{overflow-x:auto;border:1px solid #dfe6f2;border-radius:9px;max-width:100%}table.excerpt{border-collapse:separate;border-spacing:0;width:100%;font-variant-numeric:tabular-nums}
    table.excerpt th,table.excerpt td{padding:9px 10px;border-bottom:1px solid #e7ebf2;border-right:1px solid #edf0f5;white-space:nowrap;text-align:right;font-size:12px}
    table.excerpt thead th{background:#eaf0fa;color:#496587;position:sticky;top:0;text-align:center}table.excerpt .rownum{color:#6d7b91;background:#f3f6fb;text-align:center}
    table.excerpt td small{display:block;color:#9ca8b8;font-size:10px}table.excerpt td span{display:block;min-height:1em}table.excerpt td.flagged{background:#fff1a8;box-shadow:inset 0 0 0 1px #dfb849;color:#513c04;font-weight:800}
    .omitted{text-align:center!important;color:#8593a6}.tiny,.muted{color:#78869d;font-size:12px}.tiny{margin-top:13px}.done{display:block;margin-top:18px;font-size:13px;color:#5a687c}
    .nav{display:grid;gap:5px}.nav a,.sheet-link{display:block;padding:8px 9px;border-radius:7px;color:#405471;font-size:12px}.nav a:hover,.sheet-link:hover{background:#f0f4fb;text-decoration:none}
    .nav em{display:block;font-style:normal;color:#91a0b3;margin-top:2px}.dot{display:inline-block;height:7px;width:7px;border-radius:99px;background:#8da5c3;margin-right:6px}.dot.P1{background:#c53b4c}.dot.P2{background:#df9d34}
    .sheet-link{display:flex;justify-content:space-between}.sheet-link span{color:#8b99ad}.empty{padding:40px 25px;background:white;border-radius:12px;line-height:1.8}
    footer{text-align:center;padding:28px;color:#8290a6;font-size:12px}@media(max-width:960px){.layout{grid-template-columns:1fr}aside{position:static;max-height:none}.stats{grid-template-columns:repeat(2,1fr)}.panel .nav{max-height:200px;overflow:auto}}
    @media(max-width:540px){header{padding:30px 20px}.layout{padding:18px 12px}.stats{gap:8px}.stat{padding:12px}.stat strong{font-size:24px}}
    '''
    style = style.replace("__EXPAND__", t["expand"]).replace("__COLLAPSE__", t["collapse"])
    script = r'''
    (function(){
      const search = document.getElementById('search');
      const sheet = document.getElementById('sheet-filter');
      const priority = document.getElementById('priority-filter');
      const cards = Array.from(document.querySelectorAll('.issue'));
      const counter = document.getElementById('visible-count');
      function apply(){
        let count=0;
        cards.forEach(card=>{
          const ok=(sheet.value==='' || card.dataset.sheet===sheet.value || card.dataset.relatedSheet===sheet.value)
            && (priority.value==='' || card.dataset.priority===priority.value)
            && (!search.value.trim() || card.dataset.search.includes(search.value.trim().toLowerCase()));
          card.style.display=ok?'':'none'; if(ok)count++;
        });
        counter.textContent='__VISIBLE_PREFIX__'+count+' / '+cards.length+'__ITEMS__';
      }
      [search,sheet,priority].forEach(el=>el.addEventListener('input',apply));
      document.querySelectorAll('.nav a,.sheet-link').forEach(link=>link.addEventListener('click',()=>{
        const node=document.querySelector(link.getAttribute('href'));
        if(node && node.tagName==='DETAILS'){
          if(node.style.display==='none'){search.value=''; sheet.value=''; priority.value=''; apply();}
          node.open=true;
        }
      }));
      apply();
    })();
    '''
    script = script.replace("__VISIBLE_PREFIX__", t["visible_prefix"]).replace("__ITEMS__", t["items"])
    doc_lang = "zh-CN" if lang == "zh" else "en"
    return f'''<!doctype html><html lang="{doc_lang}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src 'none'; connect-src 'none';">
    <title>{t["page_title"]} — {esc(input_name)}</title><style>{style}</style></head><body>
    <header><div class="eyebrow">SOURCE DATA · READ-ONLY AUDIT</div><h1 class="main-title">{t["title"]}</h1>
    <p class="subtitle">{t["file"]} {esc(input_name)} · {t["subtitle"]}</p></header>
    <div class="layout"><aside><div class="panel"><h2>{t["sheet_index"]}</h2>{''.join(sheet_navigation)}
    <h2 style="margin-top:22px">{t["quick_jump"]}</h2><nav class="nav">{''.join(shortcuts)}</nav></div></aside>
    <main><section class="stats"><div class="stat"><strong>{sheets_scanned}</strong><small>{t["sheets_scanned"]}</small></div>
    <div class="stat"><strong>{len(findings)}</strong><small>{t["findings"]}</small></div>
    <div class="stat"><strong>{counts.get('P1',0)}</strong><small>{t["p1_count"]}</small></div>
    <div class="stat"><strong>{counts.get('P2',0)}</strong><small>{t["p2_count"]}</small></div></section>
    <div class="note"><strong>{t["notice_heading"]}</strong>{t["notice"]}</div>
    <section class="filters" aria-label="{t["filters"]}"><input id="search" type="search" placeholder="{t["search"]}">
    <select id="sheet-filter" aria-label="{t["filter_sheet"]}">{''.join(options)}</select>
    <select id="priority-filter" aria-label="{t["filter_priority"]}"><option value="">{t["all_priorities"]}</option><option value="P1">{t["p1_count"]}</option><option value="P2">{t["p2_count"]}</option><option value="P3">{t["p3_count"]}</option></select></section>
    <p class="muted" id="visible-count">{t["total_items"]} {len(findings)}{t["items"]}</p><section>{''.join(cards)}</section></main></div>
    <footer>{t["footer"]}</footer>
    <script>{script}</script></body></html>'''
