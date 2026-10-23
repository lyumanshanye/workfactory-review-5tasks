# 亚利桑那三县户外活动：把AQI历史记录转成可复核的排期政策

你是社区活动项目的数据分析师。运营团队想用2024年历史资料复盘三县室内备用场地需求，并选择下一轮计划的排期原则。这不是天气预测或个体医疗建议。项目条件为本练习假设，不是真实活动签到记录。

范围为Arizona州Maricopa、Pima、Yuma三县，2024全年。每县每个周一至周五均拟安排一场90分钟、40人的户外活动，节假日也计入。每人每场相当于1.5小时；不能把这些潜在参与人时当作真实出勤。比较：政策A，日AQI>=101就转室内；政策B，AQI>=151转室内，101至150仍户外但降低强度并提供室内选项。两种政策在缺失当日AQI时一律“待测，室内备用”，不将缺失填0。其他日期按政策运行。本题用最终日AQI做回顾性规则模拟，不能声称活动开始前即已知道该值。

工作要求：
1. 在EPA下载页/格式页找到真实网页引用，解释日表、年表和字段定义如何衔接；三县按州县代码构造5位FIPS，与县名原表校验，不靠模糊县名强拼。
2. 保留原始CSV行号，检查县日期键、重复冲突、缺失、非数值AQI以及站点数。为每县补齐2024的366日索引，单列缺测日；不得凭年度总数补造逐日值。
3. 对每县全年可观察日，按EPA AQI分档核算六档天数（Good、Moderate、USG、Unhealthy、Very Unhealthy、Hazardous）、Days with AQI和Max AQI，与年表对应字段逐项对账。差异如实记录，优先追溯记录和口径，不改原件强行对齐。县表是身份核对，不是健康结果数据。
4. 对全年工作日逐县逐日比较A/B：输出AQI、观测状态、两种动作、两种因超阈值转室内的标记、缺测备用标记、A比B多转室内标记及相应潜在参与人时。缺测备用与有观测的超标转室内分开统计。
5. 汇总每县及全项目、按月份的场次和潜在人时，量化两政策差异、差异集中的县/月；根据证据提出可执行的备用场地和活动安排建议，区分资源取舍、敏感人群保护、监测覆盖和历史数据不能支持的结论。不得推断哮喘因果或小时暴露、对下一年外推精确天数。

## 可用材料

输入仅位于只读的 `/app/input_files/`，下列全部文件均已提供。不得修改原件；联网不是完成任务的前提，不需要重新下载或更新数据。

| 文件 | 说明 |
|---|---|
| `daily_aqi_by_county_2024.csv` | EPA原始日表；State Name、county Name、State Code、County Code、Date、AQI、Category、Defining Parameter、Defining Site、Number of Sites Reporting。 |
| `annual_aqi_by_county_2024.csv` | 原始年汇总；State、County、Year及各AQI分档天数、Days with AQI、Max AQI等。 |
| `national_county2020.txt` | Census县身份原表，竖线分隔，STATEFP/COUNTYFP拼FIPS。 |
| `AQI_Technical_Assistance_Document.pdf` | EPA AQI技术说明原件，供类别和指标边界解释。 |
| `epa_download_files.html` | EPA数据下载网页快照，实际引用格式说明及日/年文件。 |
| `epa_file_formats.html` | EPA字段和数据格式网页快照，反向链接下载页。 |

## 交付

仅将最终交付文件写入 `/app/output/`。文件名精确如下，不需要另交代码、JSON或过程日志。

`outdoor_activity_model.xlsx`：至少含DataAudit、AnnualReconciliation、Monthly、CountySummary、Sources工作表；列出关键数字、缺测、对账差异、口径和原始行定位。
`activity_policy_brief.docx`：运营建议，说明A/B对资源需求的影响、重点县月、缺测处理和解释边界。
`review_calendar.csv`：所有三县2024工作日，每县日期一行，含FIPS、县、日期、原始行号（缺测为空）、AQI、站点数、观测状态、A动作、B动作、A/B超标室内标记、缺测备用标记、差异标记及差异潜在人时。
