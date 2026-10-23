# 五题审阅入口：先建立业务依赖，再合成任务

**给被测模型：每题的 `instruction.md` 作为提示词，以及该题 `environment/input_files/` 的全部原始文件（映射到只读 `/app/input_files/`）。交付写入 `/app/output/`。参考答案、评分标准、tests、provenance 和 research 不提供给盲测模型。**

五题均有实际网页引用关系；源网页、链接文字/上下文、原始href、目标URL与文件SHA256都已冻结。引用关系用于确定下游文件的业务作用，并不自动保证资料格式/版本一致。

| 题号 | 业务任务 | 题面 | 候选参考交付 | 网页引用证据 |
|---|---|---|---|---|
| GO-LINK-001 | 跨财政年度培训会：把报销规则落到可批准的会务预算 | [instruction](tasks/GO-LINK-001/instruction.md) | [文件夹](index.html?domain=GO#reference) | [原始链接及哈希](tasks/GO-LINK-001/provenance/web_relationship.json) |
| FIN-LINK-001 | 新英格兰三州银行：季度信用复核与八个核查名额分配 | [instruction](tasks/FIN-LINK-001/instruction.md) | [文件夹](index.html?domain=FIN#reference) | [原始链接及哈希](tasks/FIN-LINK-001/provenance/web_relationship.json) |
| LAW-LINK-001 | 4800名消费者的小型金融机构：泄露通报与安全整改决策 | [instruction](tasks/LAW-LINK-001/instruction.md) | [文件夹](index.html?domain=LAW#reference) | [原始链接及哈希](tasks/LAW-LINK-001/provenance/web_relationship.json) |
| DA-LINK-001 | 亚利桑那三县户外活动：把AQI历史记录转成可复核的排期政策 | [instruction](tasks/DA-LINK-001/instruction.md) | [文件夹](index.html?domain=DA#reference) | [原始链接及哈希](tasks/DA-LINK-001/provenance/web_relationship.json) |
| GA-LINK-001 | Dota 2平衡复核：将历史对局信号转成六项受控试玩假设 | [instruction](tasks/GA-LINK-001/instruction.md) | [文件夹](index.html?domain=GA#reference) | [原始链接及哈希](tasks/GA-LINK-001/provenance/web_relationship.json) |

## 每题为什么是同一条业务链

### GO-LINK-001 · 通用办公

GSA下载网页引用两年费率表和餐费拆分表 → 按逐日行程适用FTR → 比较人数/城市预算 → 形成有条件批准与补证清单。

网页：[gsa_per_diem_files.html](https://www.gsa.gov/travel/plan-a-trip/per-diem-rates/per-diem-files) 实际引用 [FY2026_PerDiemMasterRatesFile.xlsx](https://www.gsa.gov/system/files/FY2026_PerDiemMasterRatesFile.xlsx)。离线原件位于 [输入目录](index.html?domain=GO#inputs)，不是手工添加到题面的虚构关联。

模型收到的文件：

- [FY2025_PerDiemMasterRatesFile.xlsx](tasks/GO-LINK-001/environment/input_files/FY2025_PerDiemMasterRatesFile.xlsx)：GSA FY2025费率原件，Master第2行为表头；含州、目的地、县、季节、住宿上限和M&IE。
- [FY2026_PerDiemMasterRatesFile.xlsx](tasks/GO-LINK-001/environment/input_files/FY2026_PerDiemMasterRatesFile.xlsx)：GSA FY2026同结构原件；财政年度自10月1日开始。
- [GSA_FY2025_MIE_Breakdown.docx](tasks/GO-LINK-001/environment/input_files/GSA_FY2025_MIE_Breakdown.docx)：FY2025起适用的M&IE分餐与杂费拆分表。
- [eCFR_41_CFR_301-11.xml](tasks/GO-LINK-001/environment/input_files/eCFR_41_CFR_301-11.xml)：2025-01-01冻结的差旅报销规则，供本练习规则适用。
- [eCFR_41_CFR_301-74.xml](tasks/GO-LINK-001/environment/input_files/eCFR_41_CFR_301-74.xml)：同日会议规则，含政府供餐和茶歇处理。
- [gsa_per_diem_files.html](tasks/GO-LINK-001/environment/input_files/gsa_per_diem_files.html)：GSA下载页原始网页快照；其中实际链接建立费率与拆分表的引用关系。

供人工审阅的候选交付（不给模型）：

- [conference_budget.xlsx](tasks/GO-LINK-001/reference_output/conference_budget.xlsx)
- [approval_brief.docx](tasks/GO-LINK-001/reference_output/approval_brief.docx) · [PDF预览](tasks/GO-LINK-001/reference_pdf/approval_brief.pdf)
- [approval_register.csv](tasks/GO-LINK-001/reference_output/approval_register.csv)

候选解复算摘要：`{"recommendation": "Phoenix", "attendees": 24, "total": 51174.48, "stress_total": 53744.88000000001, "options": 6}`。

### FIN-LINK-001 · 金融

FDIC文档页引用Financial/Institution字段定义 → 以CERT和季度核对银行范围 → 计算可比信用代理指标 → 把八个核查名额落实到证据请求。

网页：[fdic_api_documentation.html](https://api.fdic.gov/banks/docs/) 实际引用 [financial_definitions.yaml](https://api.fdic.gov/banks/docs/risview_properties.yaml)。离线原件位于 [输入目录](index.html?domain=FIN#inputs)，不是手工添加到题面的虚构关联。

模型收到的文件：

- [fdic_financials_VT_NH_ME_2022_2025.csv](tasks/FIN-LINK-001/environment/input_files/fdic_financials_VT_NH_ME_2022_2025.csv)：746条原始银行季度记录；REPDTE为YYYYMMDD，CERT为证书号，STALP为州。
- [fdic_institutions_VT_NH_ME.csv](tasks/FIN-LINK-001/environment/input_files/fdic_institutions_VT_NH_ME.csv)：机构快照；含CERT、名称、ACTIVE和资料日期，可能含历史机构。
- [FDIC_RMS_Manual_3.2_Loans.pdf](tasks/FIN-LINK-001/environment/input_files/FDIC_RMS_Manual_3.2_Loans.pdf)：FDIC贷款检查手册原件，用于解释复核证据与局限，不提供本题银行评级。
- [fdic_api_documentation.html](tasks/FIN-LINK-001/environment/input_files/fdic_api_documentation.html)：FDIC API文档页快照，含两类字段定义的实际相对链接。
- [financial_definitions.yaml](tasks/FIN-LINK-001/environment/input_files/financial_definitions.yaml)：文档页所引用的金融字段定义快照。
- [institution_definitions.yaml](tasks/FIN-LINK-001/environment/input_files/institution_definitions.yaml)：文档页所引用的机构字段定义快照。

供人工审阅的候选交付（不给模型）：

- [credit_review.xlsx](tasks/FIN-LINK-001/reference_output/credit_review.xlsx)
- [credit_review_memo.docx](tasks/FIN-LINK-001/reference_output/credit_review_memo.docx) · [PDF预览](tasks/FIN-LINK-001/reference_pdf/credit_review_memo.pdf)
- [follow_up_register.csv](tasks/FIN-LINK-001/reference_output/follow_up_register.csv)

候选解复算摘要：`{"rows": 746, "banks": 54, "quarters": 14, "selected_CERT": [6280, 59289, 29636, 59086, 30878, 19836, 14158, 19690], "regional_observed_change_pp": 0.06962667306964987, "regional_fixed_change_pp": 0.06303109123861816}`。

### LAW-LINK-001 · 法律

FTC业务指南网页引用eCFR具体条款 → 冻结规则核对小机构例外和事件门槛 → 确定通报时间与内容 → 关联控制缺口和整改。

网页：[ftc_safeguards_guide.html](https://search.ftc.gov/business-guidance/resources/ftc-safeguards-rule-what-your-business-needs-know) 实际引用 [ecfr_314_reference_page.html](https://www.ecfr.gov/current/title-16/chapter-I/subchapter-C/part-314)。离线原件位于 [输入目录](index.html?domain=LAW#inputs)，不是手工添加到题面的虚构关联。

模型收到的文件：

- [eCFR_16_CFR_314_2025-01-01.xml](tasks/LAW-LINK-001/environment/input_files/eCFR_16_CFR_314_2025-01-01.xml)：本练习适用的冻结规则全文，含314.2、314.4、314.5、314.6。
- [eCFR_16_CFR_314_2021-12-01.xml](tasks/LAW-LINK-001/environment/input_files/eCFR_16_CFR_314_2021-12-01.xml)：2021旧版本，仅作规则变化对照。
- [FR_2021-25736_Safeguards_Rule_Final.pdf](tasks/LAW-LINK-001/environment/input_files/FR_2021-25736_Safeguards_Rule_Final.pdf)：2021最终规则及立法说明原件。
- [FR_2023-24412_Safeguards_Breach_Notification.pdf](tasks/LAW-LINK-001/environment/input_files/FR_2023-24412_Safeguards_Breach_Notification.pdf)：2023通报修订及生效安排原件。
- [fr_notification_publication.html](tasks/LAW-LINK-001/environment/input_files/fr_notification_publication.html)：同一2023修订的GovInfo HTML原始网页。
- [ftc_safeguards_guide.html](tasks/LAW-LINK-001/environment/input_files/ftc_safeguards_guide.html)：FTC业务指南网页快照，实际链接至eCFR条款。
- [ecfr_314_reference_page.html](tasks/LAW-LINK-001/environment/input_files/ecfr_314_reference_page.html)：当前eCFR页面快照，供验证引用，不替代2025冻结规则。

供人工审阅的候选交付（不给模型）：

- [safeguards_gap_register.xlsx](tasks/LAW-LINK-001/reference_output/safeguards_gap_register.xlsx)
- [legal_action_memo.docx](tasks/LAW-LINK-001/reference_output/legal_action_memo.docx) · [PDF预览](tasks/LAW-LINK-001/reference_pdf/legal_action_memo.pdf)
- [ftc_notice_draft.txt](tasks/LAW-LINK-001/reference_output/ftc_notice_draft.txt)

候选解复算摘要：`{"discovery": "2025-03-03", "deadline": "2025-04-02", "FTC_notice_required": true, "exemptions": ["314.4(b)(1)", "314.4(d)(2)", "314.4(h)", "314.4(i)"]}`。

### DA-LINK-001 · 数据分析

EPA下载网页引用格式说明 → 日AQI按FIPS对齐年度汇总和县表 → 比较两种室内切换政策 → 交付日期级决策日历。

网页：[epa_download_files.html](https://aqs.epa.gov/aqsweb/airdata/download_files.html) 实际引用 [epa_file_formats.html](https://aqs.epa.gov/aqsweb/airdata/FileFormats.html)。离线原件位于 [输入目录](index.html?domain=DA#inputs)，不是手工添加到题面的虚构关联。

模型收到的文件：

- [daily_aqi_by_county_2024.csv](tasks/DA-LINK-001/environment/input_files/daily_aqi_by_county_2024.csv)：EPA原始日表；State Name、county Name、State Code、County Code、Date、AQI、Category、Defining Parameter、Defining Site、Number of Sites Reporting。
- [annual_aqi_by_county_2024.csv](tasks/DA-LINK-001/environment/input_files/annual_aqi_by_county_2024.csv)：原始年汇总；State、County、Year及各AQI分档天数、Days with AQI、Max AQI等。
- [national_county2020.txt](tasks/DA-LINK-001/environment/input_files/national_county2020.txt)：Census县身份原表，竖线分隔，STATEFP/COUNTYFP拼FIPS。
- [AQI_Technical_Assistance_Document.pdf](tasks/DA-LINK-001/environment/input_files/AQI_Technical_Assistance_Document.pdf)：EPA AQI技术说明原件，供类别和指标边界解释。
- [epa_download_files.html](tasks/DA-LINK-001/environment/input_files/epa_download_files.html)：EPA数据下载网页快照，实际引用格式说明及日/年文件。
- [epa_file_formats.html](tasks/DA-LINK-001/environment/input_files/epa_file_formats.html)：EPA字段和数据格式网页快照，反向链接下载页。

供人工审阅的候选交付（不给模型）：

- [outdoor_activity_model.xlsx](tasks/DA-LINK-001/reference_output/outdoor_activity_model.xlsx)
- [activity_policy_brief.docx](tasks/DA-LINK-001/reference_output/activity_policy_brief.docx) · [PDF预览](tasks/DA-LINK-001/reference_pdf/activity_policy_brief.pdf)
- [review_calendar.csv](tasks/DA-LINK-001/reference_output/review_calendar.csv)

候选解复算摘要：`{"daily_rows": 1097, "calendar_rows": 786, "extra_indoor": 67, "extra_hours": 4020, "annual_nonzero_differences": 0, "missing_dates": {"Maricopa": "", "Pima": "", "Yuma": "2024-11-26"}}`。

### GA-LINK-001 · 游戏

UCI数据页明确引用采集代码及英雄映射 → 核实真实117列编码与引用材料差异 → 只用训练集选假设 → 测试集复核后安排受控试玩。

网页：[uci_dota2_dataset.html](https://archive.ics.uci.edu/dataset/367/dota2+games+results) 实际引用 [dota2_collector_page.html](https://gist.github.com/da-steve101/1a7ae319448db431715bd75391a66e1b)。离线原件位于 [输入目录](index.html?domain=GA#inputs)，不是手工添加到题面的虚构关联。

模型收到的文件：

- [dota2Train.csv](tasks/GA-LINK-001/environment/input_files/dota2Train.csv)：UCI原始训练split，无表头；列1胜方，列2cluster，列3mode，列4type，列5起113个英雄指示列。
- [dota2Test.csv](tasks/GA-LINK-001/environment/input_files/dota2Test.csv)：UCI原始保留split，同结构；只用于冻结清单后的复核。
- [uci_dota2_dataset.html](tasks/GA-LINK-001/environment/input_files/uci_dota2_dataset.html)：UCI说明网页快照，正文实际引用采集代码与映射文件，并给原始下载。
- [dota2_collector_page.html](tasks/GA-LINK-001/environment/input_files/dota2_collector_page.html)：UCI所引GitHub gist原始网页快照；核实其格式而非默认等同CSV生成器。
- [heroes_mapping.json](tasks/GA-LINK-001/environment/input_files/heroes_mapping.json)：UCI所引仓库英雄映射文件快照；ID与指示列的关系须独立核实。

供人工审阅的候选交付（不给模型）：

- [balance_review.xlsx](tasks/GA-LINK-001/reference_output/balance_review.xlsx)
- [balance_review_brief.docx](tasks/GA-LINK-001/reference_output/balance_review_brief.docx) · [PDF预览](tasks/GA-LINK-001/reference_pdf/balance_review_brief.pdf)
- [playtest_queue.csv](tasks/GA-LINK-001/reference_output/playtest_queue.csv)

候选解复算摘要：`{"train_rows": 92650, "test_rows": 10294, "cross_exact_rows": 0, "selected": ["H057", "H103", "H036", "H091", "H061", "H080"], "replicated": ["H057", "H103", "H036", "H091", "H061", "H080"]}`。

## 验证状态

192/192 项来源、结构、模板及独立业务重算检查通过，详见 [audit.json](validation/audit.json)。五题使用字节相同的Dockerfile和依赖清单；共享已实际构建的镜像、分别挂载各题只读原始输入，离线Oracle及工具/依赖/用户权限检查均通过。未为同配方重复构建五个镜像。

参考交付为作者候选解，尚未获得独立专家Gold确认。本轮作者候选参考交付的qwen3.7-plus原生判分：5/5完成，0题运行中。已完成主分：GO=1.000000；FIN=0.933333；LAW=1.000000；DA=1.000000；GA=1.000000。详情见 [JUDGE_STATUS.md](JUDGE_STATUS.md)。此分数不是GPT/Qwen独立作答成绩，尚未进行三模型难度校准。原生结果与冻结文件的50项完整性核验全部通过，见 [核验结果](validation/judge_integrity.json)。FIN008扣分涉及对比期间不符，独立CSV重算证据见 [期间复核](validation/FIN008_scope_audit.md)；原始主分0.933333保留，未覆盖评分。 不能用旧题分数充当新题分数。A2/C2为设计标签，未声明正式验收达标。

## 复现和打包

`research/build_specs.py` 从冻结材料生成题面、评分和配置；各题 `solution/reference.py` 可离线重建交付。`research/audit.py` 独立验证主要业务结果。`research/source_manifest.json` 保留输入来源；`research/verified_relationships.json` 汇总五条真实引用。

交付压缩包只含五题、交付说明和验证证据，不含2.4GB缓存、凭据或其他session的文件。详细环境见 [交付文档](交付文档.md)。
