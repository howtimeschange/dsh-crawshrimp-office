# dsh-crawshrimp-office

`dsh-crawshrimp-office` is a standalone DeepSeek Harness plugin for small, local OOXML document jobs. It creates real `.docx` and `.pptx` files, reads capped text from `.docx`, and reports structural metadata from `.docx`/`.pptx` ZIP packages. It does not use the Crawshrimp API or the shared connector.

## Install

The package targets DSH `0.2.1-alpha.4` and uses `@deepseek-ai/dsh-mcp-client@0.2.1-alpha.4` plus `@deepseek-ai/schemastery`. Install the GitHub release into a DSH profile:

```bash
dsh plugin --profile demo add -w github:howtimeschange/dsh-crawshrimp-office#v0.1.0
```

The `-w` flag lets the DSH CLI add the plugin from the profile workspace. The package declares its bundle patch, so a successful install activates `cordis.patch.yml`. The same patch is available at [`examples/cordis.patch.yml`](./examples/cordis.patch.yml) for a host that loads patches manually.

Set an optional workspace before starting the profile:

```bash
export CRAWSHRIMP_OFFICE_WORKSPACE="$HOME/crawshrimp-office-output"
```

The default workspace is `./crawshrimp-office-output` relative to the DSH process. Every filename supplied to a tool must be relative to that directory. The plugin resolves real paths, rejects traversal and symlink escapes, never overwrites an existing output, and writes through a temporary file before an atomic no-overwrite link.

## Tools

| Tool | Purpose | Approval |
| --- | --- | --- |
| `crawshrimp_office_word_create` | Create a DOCX from a title, paragraphs, and simple string tables. | Native DSH approval required. |
| `crawshrimp_office_word_read` | Extract bounded plain text from an existing DOCX. | Read-only; no approval. |
| `crawshrimp_office_slides_create` | Create a PPTX title slide and heading/bullet slides. | Native DSH approval required. |
| `crawshrimp_office_inspect` | List bounded OOXML ZIP entry metadata and counts. | Read-only; no approval. |

Write calls fail closed when the host has no approval service, when the native policy is `never`, or when the decision is not `allowed-once`. There is no plugin setting that bypasses native approval.

Example arguments:

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

## Bounds and honest limits

The defaults cap each input file at 16 MiB, each generated output at 16 MiB, and extracted DOCX text at 128 KiB. The plugin also caps paragraph, table, slide, bullet, and ZIP-entry counts. These values can be lowered or raised within bounded limits through `workspaceDir`, `maxInputBytes`, `maxOutputBytes`, and `maxTextBytes` in the plugin config.

The implementation generates and inspects OOXML packages. It does not render to PDF or PNG, run LibreOffice, or perform visual review. A successful create result reports `visualReview: "not_rendered"`; it must not be described as visually verified until a separate renderer and human or model review has actually run. `office_run`, Office COM, macros, formulas, charts, and arbitrary Python execution are outside this package.

The examples and tests use local temporary files only. They do not call a model, merchant account, or external network service.

**中文说明：** [README.zh-CN.md](./README.zh-CN.md)
