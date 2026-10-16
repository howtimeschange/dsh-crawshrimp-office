# dsh-crawshrimp-office

`dsh-crawshrimp-office` 是一个独立的 DeepSeek Harness 插件，用于小型本地 OOXML 文档任务。它可以生成真实的 `.docx`、`.pptx` 文件，把有上限的 Markdown 子集转换为 DOCX，读取 Word 文本和逐页 PowerPoint 文本，并校验 OOXML 结构及预期内容。它不调用抓虾 API，也不依赖共享 connector。

## 安装

插件面向 DSH `0.2.1-alpha.4`，使用 `@deepseek-ai/dsh-mcp-client@0.2.1-alpha.4` 和 `@deepseek-ai/schemastery`。把 GitHub 版本安装到 DSH profile：

```bash
dsh plugin --profile demo add -w github:howtimeschange/dsh-crawshrimp-office#v0.2.0
```

`-w` 让 DSH CLI 在 profile workspace 中添加 GitHub 插件。安装成功后，包中声明的 bundle patch 会自动激活 `cordis.patch.yml`。需要手动加载 patch 的宿主也可以使用 [`examples/cordis.patch.yml`](./examples/cordis.patch.yml)。

启动 profile 前可以指定工作区：

```bash
export CRAWSHRIMP_OFFICE_WORKSPACE="$HOME/crawshrimp-office-output"
```

如果宿主需要直接嵌入插件，应加载完整模块命名空间，让 Cordis 同时注册 `apply` 和工具 schema：

```js
import * as office from 'dsh-crawshrimp-office'

await ctx.plugin(office, { workspaceDir: './crawshrimp-office-output' })
```

默认工作区是 DSH 进程当前目录下的 `./crawshrimp-office-output`。工具传入的文件名必须是该目录下的相对路径。插件会解析真实路径，拒绝目录穿越和逃逸到工作区外的符号链接，不覆盖已有输出，并先写临时文件，再以不可覆盖的原子方式落盘。

## 工具

| 工具 | 用途 | 审批 |
| --- | --- | --- |
| `crawshrimp_office_word_create` | 用标题、段落和简单字符串表格生成 DOCX。 | 必须经过 DSH 原生审批。 |
| `crawshrimp_office_word_read` | 从现有 DOCX 提取有上限的纯文本。 | 只读，不需要审批。 |
| `crawshrimp_office_slides_create` | 生成 PPTX 标题页和标题/项目符号页。 | 必须经过 DSH 原生审批。 |
| `crawshrimp_office_inspect` | 输出有上限的 OOXML ZIP 条目元数据和数量。 | 只读，不需要审批。 |
| `crawshrimp_office_word_from_markdown` | 将有上限的 Markdown 标题、段落、列表和简单表格转换为 DOCX。 | 必须经过 DSH 原生审批。 |
| `crawshrimp_office_slides_read` | 读取 PPTX 每一页的有上限纯文本。 | 只读，不需要审批。 |
| `crawshrimp_office_validate` | 校验 DOCX/PPTX 必需部件及预期文本/数量。 | 只读，不需要审批。 |

写调用使用宿主的原生权限信号。会话同时报告沙箱模式 `danger-full-access` 和原生审批策略 `never` 时，已有的完全权限授予允许写入，无需重复弹出逐次审批；`ask`、受限模式，或审批/沙箱服务缺失、读取异常时，仍走原生审批路径，只有收到 `allowed-once` 才执行。插件没有单独绕过原生审批的配置项。

示例参数：

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

```json
{
  "filename": "reports/weekly.docx",
  "title": "Weekly brief",
  "markdown": "# Weekly brief\n\n- Orders: 12\n- Returns: 1\n\n| Metric | Value |\n| --- | --- |\n| Conversion | 4.2% |"
}
```

相同示例也放在 [`examples/word-from-markdown.json`](./examples/word-from-markdown.json)、[`examples/slides-read.json`](./examples/slides-read.json) 和 [`examples/validate.json`](./examples/validate.json)。

Markdown 转换支持标题、段落、有序/无序列表、引用、作为纯文本保留的围栏代码和简单管道表格。行内强调和链接会展平为文本。HTML 会被剥离并返回 warning；图片和媒体语法会保留 alt 文本并返回 warning。插件不会抓取或嵌入外部媒体，也不会执行 HTML；不支持的 Markdown 会在安全范围内保留为纯文本。

`crawshrimp_office_slides_read` 返回 `slideCount` 和每页有上限的文本，只读取 OOXML 文本节点，不解码媒体。`crawshrimp_office_validate` 在预期文本/数量不满足时返回 `valid: false`；损坏 ZIP、缺失 OOXML 必需部件、无效 XML 及 DTD/entity 声明会真实返回工具错误。

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

## 上限与实际能力

默认限制单个输入文件 16 MiB、单个生成文件 16 MiB、Markdown 输入 512 KiB、文本提取 128 KiB；段落、表格、幻灯片、项目符号、Markdown block 和 ZIP 条目也有数量上限。可以通过插件配置中的 `workspaceDir`、`maxInputBytes`、`maxOutputBytes`、`maxTextBytes`、`maxMarkdownBytes` 在限定范围内调整。

实现范围是 OOXML 包的生成、读取和结构校验。插件不转 PDF/PNG，不运行 LibreOffice，也不执行视觉审查。生成成功结果会返回 `visualReview: "not_rendered"`；只有在另行完成渲染并经过人工或模型检查后，才能描述为视觉已验证。`office_run`、Office COM、宏、公式、图表、媒体解码和任意 Python 执行不属于本包。

示例和测试只使用本地临时文件，不调用模型、商家账号或外部网络服务。

**English:** [README.md](./README.md)

## 功能归属

当前维护三个独立包：[浏览器](https://github.com/howtimeschange/dsh-crawshrimp-browser)负责全部网页操作、登录态和页面脚本，[数据](https://github.com/howtimeschange/dsh-crawshrimp-data)负责记录和表格，Office 负责文档。各包均可独立使用。PDF、标签、模板和渲染归 Office 路线图，0.2.0 尚未实现。

0.2.0 新增 Markdown 转 DOCX、PPTX 逐页文本读取、结构/内容校验，并对齐 DSH 原生完全权限映射。
