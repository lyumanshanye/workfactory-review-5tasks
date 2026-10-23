# 本轮接口连接记录

密钥不写入本文档、HTML或题包。以下结论来自真实请求验证。

| 服务 | 模型 | 请求地址 | 本轮验证 |
|---|---|---|---|
| GPT | gpt-5.6-sol | https://gapi.completion.opensii.ai/v1/chat/completions | 流式文本、工具调用及工具结果往返均成功，返回OK；未运行独立GPT作答 |
| Qwen（原生判分采用） | qwen3.7-plus | https://apicz.boyuerichdata.com/v1/messages | Anthropic兼容流式及工具调用成功 |
| Qwen（另一可用协议） | qwen3.7-plus | https://apicz.boyuerichdata.com/v1/chat/completions | OpenAI兼容流式及工具调用成功 |

GPT网关需要 `stream: true`；此次非流式请求出现上游Responses事件转换错误，不能据此判断密钥失效。`/models`在该GPT网关未实现，不作为连通性依据。

Qwen思考模式下不要强制required/object形式的tool_choice；自动工具选择已验证。原生判分使用`/messages`与Claude Code工具循环，传输桥只转发协议，不代替模型推理。

本轮五题并行，题内逐项串行，裁判为qwen3.7-plus。被评对象是作者候选参考交付；不是GPT/Qwen盲测模型输出。状态与分数见[JUDGE_STATUS.md](JUDGE_STATUS.md)。
