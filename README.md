# dsh-crawshrimp-office

`dsh-crawshrimp-office` is a standalone DeepSeek Harness plugin for small, local OOXML document jobs. It creates real `.docx` and `.pptx` files, converts a bounded Markdown subset into DOCX, reads capped Word and per-slide PowerPoint text, and validates OOXML structure plus expected content. It does not use the Crawshrimp API or the shared connector.

## Install

The package targets DSH `0.2.1-alpha.4` and uses `@deepseek-ai/dsh-mcp-client@0.2.1-alpha.4` plus `@deepseek-ai/schemastery`. Install the GitHub release into a DSH profile:

```bash
dsh plugin --profile demo add -w github:howtimeschange/dsh-crawshrimp-office#v0.2.0
```

The `-w` flag lets the DSH CLI add the plugin from the profile workspace. The package declares its bundle patch, so a successful install activates `cordis.patch.yml`. The same patch is available at [`examples/cordis.patch.yml`](./examples/cordis.patch.yml) for a host that loads patches manually.

Set an optional workspace before starting the profile:

```bash
export CRAWSHRIMP_OFFICE_WORKSPACE="$HOME/crawshrimp-office-output"
```

When embedding the plugin directly, load the module namespace so its Cordis `apply` and schemas are registered together:

```js
import * as office from 'dsh-crawshrimp-office'

await ctx.plugin(office, { workspaceDir: './crawshrimp-office-output' })
```

The default workspace is `./crawshrimp-office-output` relative to the DSH process. Every filename supplied to a tool must be relative to that directory. The plugin resolves real paths, rejects traversal and symlink escapes, never overwrites an existing output, and writes through a temporary file before an atomic no-overwrite link.

## Tools

| Tool | Purpose | Approval |
| --- | --- | --- |
| `crawshrimp_office_word_create` | Create a DOCX from a title, paragraphs, and simple string tables. | Native DSH approval required. |
| `crawshrimp_office_word_read` | Extract bounded plain text from an existing DOCX. | Read-only; no approval. |
| `crawshrimp_office_slides_create` | Create a PPTX title slide and heading/bullet slides. | Native DSH approval required. |
| `crawshrimp_office_inspect` | List bounded OOXML ZIP entry metadata and counts. | Read-only; no approval. |
| `crawshrimp_office_word_from_markdown` | Convert bounded Markdown headings, paragraphs, lists, and simple tables to DOCX. | Native DSH approval required. |
| `crawshrimp_office_slides_read` | Read bounded plain text from each PPTX slide. | Read-only; no approval. |
| `crawshrimp_office_validate` | Validate DOCX/PPTX required parts and optional expected text/counts. | Read-only; no approval. |

Write calls use the host's native permission signals. When the same session reports both sandbox mode `danger-full-access` and native approval policy `never`, the standing Full Access grant lets the write proceed without a duplicate per-call prompt. In `ask`, with a restricted mode, or when the approval/sandbox service is missing or cannot be read, the plugin keeps the native approval path and fails closed unless it receives `allowed-once`. There is no plugin setting that bypasses native approval.

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
  "filename": "reports/weekly.docx",
  "title": "Weekly brief",
  "markdown": "# Weekly brief\n\n- Orders: 12\n- Returns: 1\n\n| Metric | Value |\n| --- | --- |\n| Conversion | 4.2% |"
}
```

The same payloads are available as [`examples/word-from-markdown.json`](./examples/word-from-markdown.json), [`examples/slides-read.json`](./examples/slides-read.json), and [`examples/validate.json`](./examples/validate.json).

Markdown conversion accepts headings, paragraphs, ordered and unordered lists, blockquotes, fenced code as plain text, and simple pipe tables. Inline emphasis and links are flattened to text. HTML is stripped with warnings; image and media syntax is kept as alt text with warnings. It never fetches or embeds external media and never executes HTML. Unsupported Markdown is retained as safe plain text when possible.

`crawshrimp_office_slides_read` returns `slideCount` plus each slide's bounded text. It reads OOXML text nodes only and does not decode media. `crawshrimp_office_validate` returns `valid: false` for unmet expected text/count checks, while malformed ZIPs, missing required OOXML parts, invalid XML, and DTD/entity declarations return a truthful tool error.

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

The defaults cap each input file at 16 MiB, each generated output at 16 MiB, Markdown input at 512 KiB, and extracted text at 128 KiB. The plugin also caps paragraph, table, slide, bullet, Markdown block, and ZIP-entry counts. These values can be lowered or raised within bounded limits through `workspaceDir`, `maxInputBytes`, `maxOutputBytes`, `maxTextBytes`, and `maxMarkdownBytes` in the plugin config.

The implementation generates, reads, and structurally validates OOXML packages. It does not render to PDF or PNG, run LibreOffice, or perform visual review. A successful create result reports `visualReview: "not_rendered"`; it must not be described as visually verified until a separate renderer and human or model review has actually run. `office_run`, Office COM, macros, formulas, charts, media decoding, and arbitrary Python execution are outside this package.

The examples and tests use local temporary files only. They do not call a model, merchant account, or external network service.

**中文说明：** [README.zh-CN.md](./README.zh-CN.md)

## Functional scope

The maintained packages are [Browser](https://github.com/howtimeschange/dsh-crawshrimp-browser) (all webpage operations, login state and page scripts), [Data](https://github.com/howtimeschange/dsh-crawshrimp-data) (records and tables), and Office (documents). Each works independently. PDF, labels, templates and rendering belong to the Office roadmap; they are not implemented in 0.2.0.

Version 0.2.0 adds Markdown-to-DOCX, per-slide PPTX reading, structural/content validation, and native Full Access permission mapping.
