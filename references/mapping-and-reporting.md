# Explicit cell mapping and report interpretation

The audit script reads spreadsheets; it cannot infer which numeric rows represent biologically independent replicates. Add explicit mapping only when the source Figure/legend or author confirms it.

## Example configuration

For a hypothetical `Figure 5d` sheet with individual raw measurements in B2:B5 and reported statistics in B7:B10:

```json
{
  "aggregates": [
    {"sheet": "Figure 5d", "range": "B2:B5", "mean": "B7", "sd": "B8", "sem": "B9", "n": "B10", "decimals": 3}
  ],
  "identities": [
    {"sheet": "Figure 5d", "total": "C2", "parts": ["D2", "E2"]}
  ],
  "ratios": [
    {"sheet": "Figure 5d", "numerator": "D2", "denominator": "C2", "reported": "F2", "scale": 100, "decimals": 2}
  ],
  "ttests": [
    {"sheet": "Figure 5d", "a": "H2:H6", "b": "I2:I6", "design": "welch", "reported_p": "J2", "decimals": 4}
  ]
}
```

Use **actual cell positions**; this example is illustrative, not a template to apply blindly. Accepted designs for `ttests`: `welch`, `student`, `paired` (all two-sided). A paired test assumes entries in row order represent matched pairs; confirm the matching IDs. For Excel percent-format cells whose underlying stored number is 0.854 (displayed 85.4%), set `scale` to `1`, not `100`. For a numeric cell storing `85.4`, use `100`.

If a reported statistic is written as text (for example `p < 0.001` or `mean ± SEM` in one cell), the current script cannot directly verify that cell. Note it as unchecked; inspect the source values and statistics tool manually.

## Severity labels

- `confirmed`: exact arithmetic mismatch under explicit config assumptions. Confirm the config and scientific meaning first.
- `review`: numeric pattern, missing value, or empty formula cache requiring manual examination. Never call it fraud.
- no finding: not an endorsement of validity or journal readiness.

## Manuscript Source Data manual review

- Check that relevant charts in Main and Supplementary Figures have individual source values, not only group summaries.
- Maintain sample IDs and units; for nested observations, distinguish animals, independent cell preparations, individual cells and events.
- Review any exclusion criteria and verify that original measurements are retained.
- Cross-check graph marker counts, axis scale, legends, exact n, tests, data normalization, and analysis scripts.
- Keep raw sequence datasets in their appropriate repositories and reference the accessions in Data Availability, rather than flattening a complete multi-gigabyte sequencing dataset into Excel.
- Keep uncropped raw gels/blots where applicable; that is an image-source obligation outside this numerical auditing skill.

Consult the target journal's current Source Data requirements separately; this numeric audit is not a journal compliance assessment.
