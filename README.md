# 抓虾 Office 插件 · DeepSeek Harness（DSH）Word / PowerPoint 文档

**简体中文** | [English](README.en.md)

`dsh-crawshrimp-office` 是一个独立的 DeepSeek Harness（DSH）Office 文档插件，面向本地、可复核的小型 Word 和 PowerPoint 工作流。它可以生成真实的 `.docx` 和 `.pptx` 文件，将有界 Markdown 转换为 DOCX，读取 Word 与逐页 PowerPoint 文本，检查 OOXML ZIP 结构，并按预期文本或数量校验文档。

它适合生成会议纪要、周报、简短方案、简单演示文稿和本地文档质量检查。文件留在配置的工作区内，不调用抓虾 API、共享 connector、模型服务或外部网络。

## 快速开始

需要 Node.js `^22.19.0 || >=24.0.0`，以及已验证的 DSH `0.2.1-alpha.4`。当前发布标签是 `v0.2.0`，通过 DSH 官方插件加载器从 GitHub 安装：

```bash
dsh plugin --profile demo add -w github:howtimeschange/dsh-crawshrimp-office#v0.2.0
```

安装后，可以直接让 DSH 完成本地任务，例如：“用 `crawshrimp_office_word_from_markdown` 把这份周报保存为 `reports/weekly.docx`，再用 `crawshrimp_office_validate` 检查它是否包含‘本周总结’。”也可以直接调用工具；示例参数见 [`examples/word-from-markdown.json`](./examples/word-from-markdown.json)。

插件默认将 `./crawshrimp-office-output` 作为工作区。启动 profile 前可以指定专用目录：

```bash
export CRAWSHRIMP_OFFICE_WORKSPACE="$HOME/crawshrimp-office-output"
```

`-w` 让 DSH CLI 在 profile workspace 中添加插件。安装器会读取 `package.json` 中的 `dsh.bundle.patch`，并自动加载 `cordis.patch.yml`；需要手动加载 patch 的宿主可以使用 [`examples/cordis.patch.yml`](./examples/cordis.patch.yml)。如果宿主需要直接嵌入插件，请加载完整模块命名空间，让 Cordis 同时注册 `apply` 和工具 schema：

```js
import * as office from 'dsh-crawshrimp-office'

await ctx.plugin(office, { workspaceDir: './crawshrimp-office-output' })
```

## 能力与工具

插件实际注册以下 7 个工具：

| 工具 | 用途 | 审批 |
| --- | --- | --- |
| `crawshrimp_office_word_create` | 用标题、段落和简单字符串表格生成 DOCX。 | 遵循 DSH 原生权限；完全权限自动执行。 |
| `crawshrimp_office_word_read` | 从现有 DOCX 提取有上限的纯文本。 | 只读，不需要审批。 |
| `crawshrimp_office_slides_create` | 生成标题页以及标题/项目符号页组成的 PPTX。 | 遵循 DSH 原生权限；完全权限自动执行。 |
| `crawshrimp_office_inspect` | 查看 DOCX/PPTX 的 OOXML ZIP 结构、条目元数据和数量。 | 只读，不需要审批。 |
| `crawshrimp_office_word_from_markdown` | 将有上限的 Markdown 标题、段落、列表和简单表格转换为 DOCX。 | 遵循 DSH 原生权限；完全权限自动执行。 |
| `crawshrimp_office_slides_read` | 读取 PPTX 每一页的有上限纯文本，不解码媒体。 | 只读，不需要审批。 |
| `crawshrimp_office_validate` | 校验 DOCX/PPTX 必需部件及可选的预期文本、表格数或幻灯片数。 | 只读，不需要审批。 |

所有文件名都必须是工作区内的相对路径。插件会解析真实路径，拒绝目录穿越和逃逸到工作区外的符号链接，不覆盖已有输出，并先写临时文件，再以不可覆盖的原子方式落盘。

## 使用示例

生成 Word 文档：

```json
{
  "filename": "reports/brief.docx",
  "title": "Weekly brief",
  "paragraphs": ["A short, reviewable summary."],
  "tables": [{
    "headers": ["Metric", "Value"],
    "rows": [["Orders", "12"]]
  }]
}
```

用 Markdown 生成 Word 文档：

```json
{
  "filename": "reports/weekly.docx",
  "title": "Weekly brief",
  "markdown": "# Weekly brief\n\n- Orders: 12\n- Returns: 1\n\n| Metric | Value |\n| --- | --- |\n| Conversion | 4.2% |"
}
```

生成 PPTX：

```json
{
  "filename": "reports/launch-plan.pptx",
  "title": "Launch plan",
  "slides": [{
    "heading": "Milestones",
    "bullets": ["Design", "Review"]
  }]
}
```

更多参数见 [`examples/word-from-markdown.json`](./examples/word-from-markdown.json)、[`examples/slides-read.json`](./examples/slides-read.json) 和 [`examples/validate.json`](./examples/validate.json)。

Markdown 转换支持标题、段落、有序/无序列表、引用、作为纯文本保留的围栏代码和简单管道表格。行内强调和链接会展平为文本。HTML 会被剥离并返回 warning；图片和媒体语法会保留 alt 文本并返回 warning。插件不会抓取或嵌入外部媒体，也不会执行 HTML；不支持的 Markdown 会在安全范围内保留为纯文本。

`crawshrimp_office_slides_read` 返回 `slideCount` 和每页有上限的文本，只读取 OOXML 文本节点。`crawshrimp_office_validate` 在预期文本或数量不满足时返回 `valid: false`；损坏 ZIP、缺失 OOXML 必需部件、无效 XML 及 DTD/entity 声明会真实返回工具错误。

## 权限与上限

写调用使用宿主的 DSH 原生权限信号。会话同时处于 `danger-full-access` 且原生审批策略为 `never` 时，已有的完全权限授予允许写入，无需重复弹出逐次审批；`ask`、受限模式，或审批/沙箱服务缺失、读取异常时，仍走原生审批路径，只有收到 `allowed-once` 才执行。插件没有单独绕过原生审批的配置项。

默认限制单个输入文件 16 MiB、单个生成文件 16 MiB、Markdown 输入 512 KiB、文本提取 128 KiB；段落、表格、幻灯片、项目符号、Markdown block 和 ZIP 条目也有数量上限。可以通过插件配置中的 `workspaceDir`、`maxInputBytes`、`maxOutputBytes`、`maxTextBytes`、`maxMarkdownBytes` 在限定范围内调整。

## 实际边界

实现范围是 OOXML 包的生成、读取和结构校验。插件不转 PDF/PNG，不运行 LibreOffice，也不执行视觉审查。生成成功结果会返回 `visualReview: "not_rendered"`；只有在另行完成渲染并经过人工或模型检查后，才能描述为视觉已验证。`office_run`、Office COM、宏、公式、图表、媒体解码和任意 Python 执行不属于本包。

示例和测试只使用本地临时文件，不调用模型、商家账号或外部网络服务。该包不编辑已有 Office 文档的任意复杂排版，也不渲染图表或重新计算公式。

## 与其他功能包的关系

当前维护三个可独立安装的功能包：[Browser](https://github.com/howtimeschange/dsh-crawshrimp-browser) 负责网页操作、登录态和页面脚本，[Data](https://github.com/howtimeschange/dsh-crawshrimp-data) 负责记录与表格，Office 负责文档。PDF、标签、模板和渲染归 Office 路线图，0.2.0 尚未实现。

0.2.0 新增 Markdown 转 DOCX、PPTX 逐页文本读取、结构/内容校验，并对齐 DSH 原生完全权限映射。
