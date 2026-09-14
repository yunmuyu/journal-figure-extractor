# 科技期刊图件提取与 AI 校对工具

Windows 本地工具：从作者 Word 原稿提取图件，与重制后的 PDF/JPG 自动配对，并调用视觉模型做内容差异校对。

当前稳定基线：**v1.2**。

## 主要能力

- 支持 `.doc` / `.docx`
- 从正文提取第一作者、文章标题、图号、图题
- 优先直接提取 Word 中真实图片/矢量对象，必要时回退到 PDF 裁图
- 重制图支持 PDF / JPG
- 按“第一作者 + 图号”自动配对
- 三栏查看：作者原稿｜重制图｜AI 结论
- AI 校对结果：`PASS / REVIEW / FAIL`
- 纯样式变化默认忽略；文字、数字、公式、单位、节点、箭头、连接关系等按实质差异检查
- v1.2 默认使用双阶段严格校对，避免“语义一致但漏字”的误判

## 本机运行

环境：Windows + Microsoft Word。

完整克隆/下载仓库后，双击：

```text
启动.cmd
```

或：

```text
START.cmd
```

程序会启动本地 Flask 服务并自动打开浏览器。

## AI 服务

当前代码支持：

- 阿里云百炼 Qwen（主服务）
- Groq Qwen（备用）
- OpenRouter Free（兜底）

API Key 不应提交到 GitHub。请保存在 Windows 凭据管理器或本机环境变量中。

## 我们后续怎么改

这个仓库采用**极简工作流**：

1. 你直接描述问题/需求，并附截图、样本或结果 ZIP。
2. ChatGPT / Agent 直接读取当前 `main` 代码并修改。
3. 普通 bug 和小功能直接提交到 `main`，不强制分支、PR、Issue。
4. 你本机执行 `git pull`（或让本地 Agent 同步），然后运行验证。
5. 如果还有问题，继续在下一次提交修复。

只有大规模重构或高风险修改时，才单独开分支/PR。

Agent 规则见 [`AGENTS.md`](./AGENTS.md)。

## 安全注意

公开仓库中不要提交：

- API Key / `.env`
- 作者未发表 Word 原稿
- 本地日志
- `runtime` 运行产物
- AI 校对结果中涉及不宜公开的稿件内容

## 项目性质

这是 Flask + Python 的本地 Web 应用，不是纯静态网站，因此 **GitHub Pages 不能直接运行主程序**。GitHub 用于代码托管与协作，真正的 Word/Office 处理仍在 Windows 本机执行。
