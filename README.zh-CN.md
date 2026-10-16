# dsh-crawshrimp-office

`dsh-crawshrimp-office` 是一个独立的 DeepSeek Harness 插件，用于小型本地 OOXML 文档任务。它可以生成真实的 `.docx`、`.pptx` 文件，读取有上限的 `.docx` 文本，并输出 `.docx`/`.pptx` ZIP 包的结构元数据。它不调用抓虾 API，也不依赖共享 connector。

## 安装

插件面向 DSH `0.2.1-alpha.4`，使用 `@deepseek-ai/dsh-mcp-client@0.2.1-alpha.4` 和 `@deepseek-ai/schemastery`。把 GitHub 版本安装到 DSH profile：

```bash
dsh plugin --profile demo add -w github:howtimeschange/dsh-crawshrimp-office#v0.1.0
```

`-w` 让 DSH CLI 在 profile workspace 中添加 GitHub 插件。安装成功后，包中声明的 bundle patch 会自动激活 `cordis.patch.yml`。需要手动加载 patch 的宿主也可以使用 [`examples/cordis.patch.yml`](./examples/cordis.patch.yml)。

启动 profile 前可以指定工作区：

```bash
export CRAWSHRIMP_OFFICE_WORKSPACE="$HOME/crawshrimp-office-output"
```

默认工作区是 DSH 进程当前目录下的 `./crawshrimp-office-output`。工具传入的文件名必须是该目录下的相对路径。插件会解析真实路径，拒绝目录穿越和逃逸到工作区外的符号链接，不覆盖已有输出，并先写临时文件，再以不可覆盖的原子方式落盘。

## 工具

| 工具 | 用途 | 审批 |
| --- | --- | --- |
| `crawshrimp_office_word_create` | 用标题、段落和简单字符串表格生成 DOCX。 | 必须经过 DSH 原生审批。 |
| `crawshrimp_office_word_read` | 从现有 DOCX 提取有上限的纯文本。 | 只读，不需要审批。 |
| `crawshrimp_office_slides_create` | 生成 PPTX 标题页和标题/项目符号页。 | 必须经过 DSH 原生审批。 |
| `crawshrimp_office_inspect` | 输出有上限的 OOXML ZIP 条目元数据和数量。 | 只读，不需要审批。 |

没有审批服务、原生策略为 `never`，或审批结果不是 `allowed-once` 时，写调用都会拒绝执行。插件没有单独绕过原生审批的配置项。

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
  "filename": "reports/launch-plan.pptx",
  "title": "Launch plan",
  "slides": [{
    "heading": "Milestones",
    "bullets": ["Design", "Review"]
  }]
}
```

## 上限与实际能力

默认限制单个输入文件 16 MiB、单个生成文件 16 MiB、DOCX 提取文本 128 KiB；段落、表格、幻灯片、项目符号和 ZIP 条目也有数量上限。可以通过插件配置中的 `workspaceDir`、`maxInputBytes`、`maxOutputBytes`、`maxTextBytes` 在限定范围内调整。

实现范围是 OOXML 包的生成和结构检查。插件不转 PDF/PNG，不运行 LibreOffice，也不执行视觉审查。生成成功结果会返回 `visualReview: "not_rendered"`；只有在另行完成渲染并经过人工或模型检查后，才能描述为视觉已验证。`office_run`、Office COM、宏、公式、图表和任意 Python 执行不属于本包。

示例和测试只使用本地临时文件，不调用模型、商家账号或外部网络服务。

**English:** [README.md](./README.md)
