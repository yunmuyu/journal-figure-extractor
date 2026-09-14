科技期刊图件提取与 AI 校对工具 v1.2
====================================

本版基于已经跑通的 v0.8“Word 对象优先”提取逻辑，移除 Gemini，改为多 API：
1. 阿里云百炼 Qwen（默认主服务）
2. Groq Qwen（备用）
3. OpenRouter Free（最后兜底）

推荐配置
--------
只配置阿里云百炼即可正常使用。若同时配置 Groq / OpenRouter，并把“校对模式”设成“自动模式”，则百炼发生网络错误、429、5xx、模型不可用等 API 失败时，程序会自动尝试后面的已配置服务。

自动降级顺序：百炼 → Groq → OpenRouter Free
注意：模型已经正常返回 PASS / REVIEW / FAIL 时，不会再额外调用第二个模型。

如何启动
--------
完整解压后双击：启动.cmd 或 START.cmd
页面标题必须显示 v1.2。旧版本仍在后台时，本版会自动换端口，不会串台。

阿里云百炼 API Key（最推荐）
---------------------------
1. 登录阿里云百炼控制台，开通模型服务。
2. 进入 API Key / 密钥管理页面，创建按量付费 API Key；默认业务空间即可。
3. 中国大陆版 Base URL 已由程序固定为：
   https://dashscope.aliyuncs.com/compatible-mode/v1
4. 将 Key 粘到网页“阿里云百炼 Qwen”卡片，点击“测试并保存”。
5. 默认模型 qwen3-vl-plus；如想省额度可选 qwen3-vl-flash。

官方帮助：
https://help.aliyun.com/zh/model-studio/get-api-key

Groq API Key（备用）
--------------------
在 GroqCloud Console 创建 API Key，然后粘到 Groq 卡片测试并保存。
默认模型：qwen/qwen3.8-27b
备用：qwen/qwen3.6-27b

控制台：
https://console.groq.com/keys

OpenRouter API Key（最后兜底）
-----------------------------
在 OpenRouter 创建 API Key，粘到 OpenRouter 卡片测试并保存。
模型固定为 openrouter/free。它会路由到当时可用且满足图片输入条件的免费模型，因此结果尺度可能比固定模型更波动，只建议当兜底。

密钥页面：
https://openrouter.ai/settings/keys

隐私与 Key
-----------
- Word 原稿、作者识别、图号识别、图片提取与配对全部在本机完成。
- 只有点击 AI 校对时，当前配对的两张图片才会发送到实际使用的 AI 服务。
- API Key 不写进 HTML/JS。勾选“保存到本机”后，Windows 下优先存入 Windows 凭据管理器。
- 也支持环境变量：DASHSCOPE_API_KEY / GROQ_API_KEY / OPENROUTER_API_KEY。

AI 判错规则
-----------
默认忽略字体、字号、线宽、布局重排、黑白转换等纯样式变化。
重点核对：文字、数字、公式、上下标、单位、节点、层级、箭头、连接关系、方向、流程顺序、图例、坐标/刻度、数据关系、增删内容。
结果：PASS / REVIEW / FAIL，并给出具体差异位置。

导出
----
“下载结果 ZIP”中包含：原稿图、重制图、AI校对结果.json、AI校对结果.csv、三栏 HTML 报告。
报告记录实际使用的 AI 服务、模型与 Token（若服务返回 usage）。

说明
----
本版代码已做 Python / JavaScript 语法检查，但 API Key 与 Windows + Office 环境仍需在用户电脑上做真实端到端测试。


v1.2 关键变化：默认启用“精确双阶段”校对。第一阶段逐字机械抄录并建立文字库存，第二阶段再比较结构/箭头/层级；若文字库存未完全对齐，后端禁止自动 PASS。
