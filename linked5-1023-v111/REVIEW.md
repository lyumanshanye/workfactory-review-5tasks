# 五题加难版 v1.1.1

给模型：该题instruction.md和environment/input_files全部材料，包含明确假设的decision_constraints.json。参考解、rubric、tests、其他模型答案与本会话不提供给作答模型；外网出口禁用。

校准目标为GPT约85/100、Qwen约70/100，先按五题平均试测，同时逐题复核。目标不冒充实得分。作者候选解尽可能正确，其预检得分与模型盲解分开。独立专家Gold和三模型正式验收均未完成。

| 题目 | 本版新增的业务决策约束 | 题面 | 模型输入 | 作者候选交付 |
|---|---|---|---|---|
| GO | 分年度资金上限41000/12000、待批调拨最多1500、压力情景储备及最终批准条件。 | [题面](tasks/GO-LINK-001/instruction.md) | [目录](index.html?domain=GO#inputs) | [目录](index.html?domain=GO#reference) |
| FIN | 两种替换顺序的对称变化分解、13人日可行性、追加人日后完整核查分配。 | [题面](tasks/FIN-LINK-001/instruction.md) | [目录](index.html?domain=FIN#inputs) | [目录](index.html?domain=FIN#reference) |
| LAW | 825导出记录/600不同消费者/505通知事件人数的口径区分，以及密钥缺证的305–505范围。 | [题面](tasks/LAW-LINK-001/instruction.md) | [目录](index.html?domain=LAW#inputs) | [目录](index.html?domain=LAW#reference) |
| DA | 按日期档期报价和县权重，在月预算/每日最多一县约束下选择最优追加补贴。 | [题面](tasks/DA-LINK-001/instruction.md) | [目录](index.html?domain=DA#inputs) | [目录](index.html?domain=DA#reference) |
| GA | 21个训练cluster逐一剔除的稳健性核验，训练内选三个完整试玩名额。 | [题面](tasks/GA-LINK-001/instruction.md) | [目录](index.html?domain=GA#inputs) | [目录](index.html?domain=GA#reference) |

## 独立作答实分（0–100）

| 题目 | GPT | Qwen |
|---|---:|---:|
| GO | 100.0 | running |
| FIN | running | running |
| LAW | running | 100.0 |
| DA | running | running |
| GA | running | running |

实分覆盖：GPT 1/5，Qwen 1/5。当前完成部分均分：{"gpt": 100.0, "qwen": 100.0}；只有各自五题全出分后才作本版目标判定。状态见 [CALIBRATION_STATUS.json](CALIBRATION_STATUS.json)。

## 验证与版本

原始30份公共输入与固定test.sh/finalize.py/prompt.md字节均保留；新增5份明确假设的决策约束输入，共35份输入、15份业务交付。51项独立结构/数据重算核验通过，15份Docker/本地交付语义一致；五份作者Word已另导出PDF。

[独立核验](validation/difficulty_audit.json) · [跨环境一致性](validation/runtime_output_comparison.json) · [盲解状态与失败项](validation/blind_summary.json) · [盲解可见材料与捕获证据](validation/blind_capture_evidence.json) · [作者解预检](validation/author_judge_summary.json) · [冻结任务文件SHA256](FROZEN_TASKS.json)

基础数据、口径和新决策约束服务于同一个资金/人力/法律行动/排期/试玩决定。新增约束及未审内部主张明确属于本练习假设，不冒充一手公共资料。

GPT网关曾重复发送完整工具参数；技术失败原始记录保留，并只对无有效捕获的样本作新鲜重试。不同有效答案之间不按分数挑选；不将认证、传输失败当能力低分。新版不改写旧版原始分数。

GO007措辞澄清版：题面、输入和已捕获答案字节不变，仅澄清原评分项不要求物理行排序。GO原始分数保留，新评分单独存储。[版本证明](validation/wording_patch_proof.json)；[答案选择来源](validation/selection_map.json)。
