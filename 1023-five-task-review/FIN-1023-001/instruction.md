# 银行风险变化诊断与尽调决策包

你将作为我们的高级财务分析师，完成一项明确为假设性的北新英格兰地区银行组合评估。管理层需要了解截至2025年6月30日，样本内银行的经营表现和财务风险如何演变，哪些机构值得优先跟进，以及有限的复核资源应投向何处。本次结论用于内部管理决策，不应外推为对全国银行业的判断。

请交付regional_bank_risk_model.xlsx。工作簿应提供可复核的银行季度分析，清楚展示样本覆盖、机构对应关系、经营与风险指标、时间趋势、同类比较、组合变化归因、优先级依据和支持管理层复核的明细。模型应能说明组合结果的变化在多大程度上来自持续机构自身表现，又在多大程度上来自样本构成或规模权重变化。

请交付operating_performance_diagnosis.docx。报告应形成完整判断，而不是简单排列若干比率或列出表现最差的银行。请综合盈利能力、资本承受力、信用质量、资金状况、贷款结构、增长和业务特征，指出主要变化驱动与持续落后机构。执行摘要请明确呈现样本整体资产回报率、非应计资产率、净核销率和总风险资本比率。对于重要变化，请提出并比较合理的竞争性解释，说明哪些解释获得现有证据支持、哪些仍需验证，不要把观察性归因表述为已经证明的因果关系。报告应给出可执行的组合层面和机构层面改进建议，并明确推荐管理层采用的优先行动方案。

请交付priority_follow_up_brief.docx。该简报应选择一个有证据支持的优先跟进范围，逐家说明关注理由、需要管理层回答的问题、下一步资料请求、建议行动及后续监测重点。优先范围不设固定银行名单或固定数量，但选择必须能够从模型和报告中核验。

评估期间为2022年3月31日至2025年6月30日。跟进方案应支持未来九十天的管理安排。请在全部交付物中清楚陈述结果所依赖的假设和限制。

## 输入原件

- `FDIC_RMS_Manual_3.2_Loans.pdf`
- `OCC_Rating_Credit_Risk.pdf`
- `fdic_api_documentation.html`
- `fdic_financials_VT_NH_ME_2022_2025.csv`
- `fdic_financials_supplement_official.csv`（同一范围的官方补充快照，含缺失的分母与指标；使用前核对与原快照的对应及差异）
- `fdic_institutions_VT_NH_ME.csv`
- `financial_definitions.yaml`
- `institution_definitions.yaml`

本题为新拟分析委托；所有随题输入均为真实公开原件。这是基于所给快照的回顾分析，不是分析期末当时可得信息回测。输入路径为 `/app/input_files/`，交付路径为 `/app/output/`。

## 源文件清单

| `/app/input_files/` 下的相对路径 | 内容与口径 |
| --- | --- |
| `FDIC_RMS_Manual_3.2_Loans.pdf` | FDIC贷款审查手册原件，按文件版本用于研究指导。 |
| `OCC_Rating_Credit_Risk.pdf` | OCC信用风险评级手册原件。 |
| `fdic_api_documentation.html` | FDIC BankFind接口官方文档快照。 |
| `fdic_financials_VT_NH_ME_2022_2025.csv` | FDIC财务快照；CERT、REPDTE及财务字段，定义和单位见随附官方字段说明。 |
| `fdic_financials_supplement_official.csv` | 同范围FDIC补充财务快照；含原快照缺少的分子分母，使用前核对键和重合字段。 |
| `fdic_institutions_VT_NH_ME.csv` | FDIC机构较晚快照；CERT、名称、州、ACTIVE及报告期等，不是历史状态记录。 |
| `financial_definitions.yaml` | FDIC财务字段、公式和单位官方定义。 |
| `institution_definitions.yaml` | FDIC机构字段官方定义。 |
