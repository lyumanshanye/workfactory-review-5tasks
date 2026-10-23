# GPT DA output review: evidence, not a score

Run: `20261022T204433Z-gpt-5.6-sol-DA-1023-001`

Snapshot: `2186e947e8a88063`. The companion observation JSON binds the input and output hashes, and this note's hash. No judge was called for this review.

## Verified scope

`audit_da_gpt_output.py` independently reads the four frozen originals. It does not import the reference builder or use its saved results. It recomputes target eligibility, endpoint values, both growth measures and rankings, exclusions, ordered top-20 lists, income component changes/contributions, and source row numbers in the inspected sheets. All 62,651 checks passed.

This covers the target dataset and selected contribution/provenance fields. It does not cover every cell in every summary sheet, all Word claims, historical geographic compatibility, visual presentation, or the full rubric.

## Substantive observations

The memo explicitly restricts publication to 3,082 directly comparable regions. It warns against calling this all US counties, excludes 62 regions, and explains why the provided files cannot justify fabricated splits or geographic substitutions. Unlike the reviewed Qwen memo, the inspected text does not broaden the population equality claim to all 3,144 target records.

The memo's uncertainty section is a Word table, not an empty section: paragraph-only extraction would miss its substantive publication decisions. The table distinguishes population-source robustness, geography restrictions, small-population interpretation, price-level limitations, and accounting versus causal conclusions.

Additional workbook readbacks confirm the memo's specific counts: all 3,082 eligible rows have identical BEA/Census population endpoints; 14 of the top 20 have declining population; 11 have 2024 population under 10,000; the largest income-change component is resident net earnings in 15 cases and dividends/interest/rent in five. Rank 21 is San Miguel County, Colorado, with unrounded nominal per-capita growth of 0.6040474207758311. These readbacks are consistent with the reviewed memo, but do not independently approve every other narrative assertion.

Income contributions use fractional growth values formatted as percentages. Differences from a reference storing percentage-point numbers must be interpreted with the displayed formats and definitions; representation alone is not a substantive error.

## Implication for task calibration

No score, Gold approval, weakness label, or A difficulty level is assigned here. Within the inspected scope, this GPT result appears materially stronger than the reviewed Qwen result. A strict difficulty requirement cannot be met by inventing failures or penalizing equivalent mathematics. Await full criterion-level external scoring and the third model; if the measured mean fails the difficulty requirement, revise the task transparently as a new version and rerun the required evidence rather than carrying old scores forward.
