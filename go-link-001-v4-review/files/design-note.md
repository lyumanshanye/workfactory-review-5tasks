# GO-LINK-001 v4 draft

Status: draft; not independently solved or model-calibrated. Frozen v3 and its blind runs are not updated or rescored.

## Implemented scope

- FY2025/FY2026/FY2027 matched-rate comparison; actual reimbursement remains FY2027. Historical scenarios hold market quotes and trip structure fixed and flag historical cap violations rather than inventing bookable capped prices.
- Documented discrete city/airport/flights/hotel/date candidates, airfare switching boundaries, missing-cost ceilings, and piecewise conditional recommendations.
- Individual airport access compared from a standard headquarters base; no presumed agency charter or invented residences.
- Mandatory original-template approval PDF; five criteria cover template/known fields, purpose/mission, attendance/remote justification, cost representation, and missing fields/attachments. Legacy reporting-label applicability is assessed with institutional policy under GO025.
- Five configured deliverables; 32 synchronized binary criteria. Positive weights total 238; negative weights total -7. GO037–040 total 40. PDF-specific criteria total 23 (9.7% of positive weight), compared with 52/306 (17.0%) before consolidation.
- Hotel quality screening requires two concrete searched hotels per city, review score scale/count/platform/date and relevant recent negative feedback. Apply one justified quality standard consistently, then the existing headcount/cost ordering. Documented exclusions may explain insufficient eligible hotels. GO023 evaluates quality evidence and selection with weight 10; fee-inclusion checks belong to GO022.

## Mathematical interpretation for review

For each documented full candidate j at its stated headcount, K_j is known base cost including baseline group airfare A_j; rho is additional airfare inflation, r is reserve, and U_j is missing base cost applicable to that candidate. Candidates may share a city-level U only where procurement scope is genuinely identical. Different hotel/venue packages or headcounts can require different missing-cost parameters.

    C_j(rho, U_j) = (1+r) * (K_j + rho*A_j + U_j)
    rho_ij = (K_j-K_i+U_j-U_i) / (A_i-A_j)
    Umax_j(rho) = B/(1+r) - K_j - rho*A_j

For equal airfare slopes, compare intercepts; do not divide by zero. A pairwise crossing is a recommendation boundary only if both candidates satisfy the relevant constraints and no higher-priority or cheaper feasible candidate supersedes them. A negative Umax means the known costs already exceed budget. Unknown costs are not zero. Without defensible upper bounds, state conditional regions rather than universal robustness.

Discrete choose-one selection can be represented with binary x_j and sum(x_j)=1; enumeration is acceptable. The two scalar formulas alone do not constitute mixed-integer optimization. Apply headcount priority before same-headcount cost minimization. Distinguish a conditional planning recommendation from authority to approve pending costs or compliance evidence.

GO002 evaluates matched rate inputs; GO009 evaluates historical impacts. GO021 evaluates airport access tradeoffs; GO037 evaluates joint candidate scope and itinerary feasibility. GO038/039 evaluate threshold arithmetic; GO040 evaluates actionable conditions around the recommendation and effective alternatives, not an exhaustive partition of all unknown costs. Other documented candidates must be checked so an arbitrary pairwise crossing cannot masquerade as a recommendation boundary.

GO012 evaluates cross-document agreement, not absolute arithmetic. GO029/031/032/034/035 cover the PDF. The judge prompt specifies error ownership: do not fail downstream consistency or decision logic merely because an already-scored arithmetic error propagated; independent downstream mistakes and missing required content still fail.

Date changes recalculate budget components but do not manufacture new quotes. Outside a quote's applicable dates, show stale/needs-requote status and conditional estimates; no requirement for Excel to fetch live prices. Dynamic recommendation must respect this validity status.

## Consolidation map

| Removed ID | Retained ID | Scope |
|---|---|---|
| GO005 | GO004 | Initial supplied-meal and breakfast policy |
| GO016 | GO011 | Reopening integrity as basic delivery |
| GO024 | GO001 | Source traceability |
| GO027 | GO011 | Static snapshot handoff |
| GO028 | GO010 | CSV format and actionable register |
| GO030 | GO029 | Known PDF fields and original template |
| GO033 | GO035 | Unknown PDF fields and attachment status |
| GO036 | GO025 | Legacy labels and current policy applicability |

Stable IDs are intentionally not renumbered. GO014 remains independent at weight 7 because executing the breakfast toggle tests linkage beyond the initial calculation. GO017/020 become weight 3; GO009/023 become weight 10. PDF narrative and disclosure items become weight 3, with template completion and cost representation at weight 7.

## Verification and remaining work

JSON and TOML must parse and agree on all 32 IDs, descriptions, binary types and signed weights. Artifact and metadata lists retain five filenames. SHA-256 comparisons must confirm all seven input files match v3 byte for byte. The machine-readable consolidation audit records the checks and source hashes.

No paid model runs were launched. A fresh independent solution, threshold/region calculation review, rendered deliverable inspection, and difficulty calibration remain necessary before deployment. Existing golden/reference outputs have not been upgraded and cannot establish v4 correctness. The previous v3 explicitly exempted filling the PDF; old outputs must not retroactively lose points under the new requirement.

Additional time/cost, cancellation, or uncertainty models discussed after this draft are proposals only and have not been added to its task or rubric. Verified external conference sourcing remains separate; no unverified event has been inserted into the fixed-date training scenario.

Before the hotel-quality edit, the complete v4 task and rubric generator were copied to `backups/go-v4-before-hotel-quality-20261023T180342551690Z/`; all 15 task files were verified byte-for-byte. Frozen v3 remains in its original iteration directory.

Before consolidation, the complete 40-criterion v4 task, generator and this note were backed up to `backups/go-v4-before-rubric-consolidation-20261023T181131202578Z/`; task bytes were checked after copying. The HTML review is a separate publication and does not overwrite the previous five-task preview or its model outputs.

The user's shortened instruction was preserved as the basis of the subsequent edit. Before execution/quote/screenshot fixes, the complete task and generator were saved to `backups/go-v4-before-executability-fixes-20261024T013422359259Z/`. GO003 now explicitly excludes unauthorized above-cap hotels rather than truncating quotes. GO022 checks room/breakfast demand and group-estimate status; GO037 checks compatible travel, hotel and venue dates/capacity; GO040 checks fixed/per-person/tiered cost behavior when reducing headcount. GO019 requires comprehensive adopted-quote screenshots including applicability and necessary fees, decision-relevant cancellation terms and hotel-review evidence; multiple linked captures are acceptable and irrelevant website content is not required. Criterion count and weights remain 32 and 238 positive / -7 negative. Audit assertions now check substantive requirements rather than superseded exact sentence wording.
