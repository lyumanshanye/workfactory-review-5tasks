# 新五题 Judge 状态

本轮判分对象为冻结的作者候选参考交付，不是GPT或Qwen的独立作答。裁判为qwen3.7-plus，原生RewardKit，按题并行、题内逐条串行。

| 题目 | 状态 | 主分（0–1） | 通过条目 |
|---|---|---:|---|
| GO-LINK-001 | completed | 1.0 | 12 / 12 |
| FIN-LINK-001 | completed | 0.933333 | 12 / 13 |
| LAW-LINK-001 | completed | 1.0 | 12 / 12 |
| DA-LINK-001 | completed | 1.0 | 11 / 11 |
| GA-LINK-001 | completed | 1.0 | 12 / 12 |

只有评分器无错误且冻结文件未改变的最终结果才展示主分。API调用次数不等于完成评分项数。

更新时间：2026-10-23T03:17:56.497506+00:00

FIN008 独立复核：裁判认定的三家新增银行在2024-06-30已存在，与题面要求的对比期间不符。原始主分0.933333保留；见 [期间复核](validation/FIN008_scope_audit.md) 和 [可复算证据](validation/FIN008_scope_audit.json)。

## FIN-LINK-001 未通过项

```json
[
  {
    "id": "FIN008",
    "name": "FIN008",
    "value": 0.0,
    "raw": "no",
    "weight": 7.0,
    "description": "进入/退出银行与资产规模可核验，残差被解释为范围/权重差异而非纯粹因果效果。\nInspect output/credit_review.xlsx, output/credit_review_memo.docx, output/follow_up_register.csv.",
    "reasoning": "Criterion FIN008 has two parts: (1) entering/exiting banks and asset scale must be verifiable, and (2) the residual must be explained as scope/weighting differences rather than pure causal effect.\n\nPart 2 is satisfied: the memo explicitly states '共同银行以2024-06-30资产固定加权，区域变化+0.0630个百分点；与观察变化的差+0.0066个百分点同时含范围和权重变化，不能说纯为银行恶化或单一成因', correctly attributing the residual to scope/weighting changes.\n\nPart 1 is NOT satisfied: The EndpointBridge sheet lists exiting banks (CERT 14240, 17748) but shows entering_CERT=None and entering_end_assets=0 for all states. However, the actual data shows 3 entering banks (CERT 59289, 59298, 35206) with assets 145991, 269483, 72604 respectively. The memo states '进入/退出CERT及金额见EndpointBridge', but this sheet does not contain the entering banks. While one entering bank (59289) appears in follow_up_register.csv, the primary reference point (EndpointBridge) is incomplete, making the entering banks and their asset scale not fully verifiable as claimed."
  }
]
```
