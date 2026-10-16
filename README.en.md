# Crawshrimp Office Plugin for DeepSeek Harness (DSH)

[简体中文](README.md) · [中文完整说明](README.zh-CN.md)

`dsh-crawshrimp-office` is a standalone DeepSeek Harness (DSH) Office plugin for small, local, reviewable Word and PowerPoint workflows. It creates real `.docx` and `.pptx` files, converts a bounded Markdown subset into DOCX, reads bounded Word and per-slide PowerPoint text, inspects OOXML ZIP structure, and validates expected text or counts.

It is useful for meeting notes, weekly reports, short plans, simple slide decks, and local document quality checks. Files stay inside the configured workspace; the plugin does not call the Crawshrimp API, a shared connector, a model service, or the external network.

## Quick start

Requires Node.js `^22.19.0 || >=24.0.0`, and DSH `0.2.1-alpha.4` (the tested API version). The current release tag is `v0.2.0`; install it from GitHub with the official DSH plugin loader:

```bash
dsh plugin --profile demo add -w github:howtimeschange/dsh-crawshrimp-office#v0.2.0
```

The `-w` flag lets the DSH CLI add the plugin from the profile workspace. After a successful install, the package's bundle patch activates `cordis.patch.yml`; hosts that load patches manually can use [`examples/cordis.patch.yml`](./examples/cordis.patch.yml).

After installation, ask DSH to complete a local task such as: “Use `crawshrimp_office_word_from_markdown` to save this weekly report as `reports/weekly.docx`, then use `crawshrimp_office_validate` to check that it contains ‘This week’.” You can also call the tools directly; see [`examples/word-from-markdown.json`](./examples/word-from-markdown.json) for an argument example.

The default workspace is `./crawshrimp-office-output` relative to the DSH process. Set a dedicated directory before starting the profile:

```bash
export CRAWSHRIMP_OFFICE_WORKSPACE="$HOME/crawshrimp-office-output"
```

The loader reads `package.json`'s `dsh.bundle.patch` metadata and applies `cordis.patch.yml`. When embedding the plugin directly, load the complete module namespace so Cordis registers both `apply` and the tool schemas:

```js
import * as office from 'dsh-crawshrimp-office'

await ctx.plugin(office, { workspaceDir: './crawshrimp-office-output' })
```

## Capabilities and tools

The plugin registers these seven tools:

| Tool | Purpose | Approval |
| --- | --- | --- |
| `crawshrimp_office_word_create` | Create a DOCX from a title, paragraphs, and simple string tables. | Native DSH approval required. |
| `crawshrimp_office_word_read` | Extract bounded plain text from an existing DOCX. | Read-only; no approval. |
| `crawshrimp_office_slides_create` | Create a PPTX with a title slide and heading/bullet slides. | Native DSH approval required. |
| `crawshrimp_office_inspect` | Inspect OOXML ZIP structure, entry metadata, and counts for a DOCX or PPTX. | Read-only; no approval. |
| `crawshrimp_office_word_from_markdown` | Convert bounded Markdown headings, paragraphs, lists, and simple tables to DOCX. | Native DSH approval required. |
| `crawshrimp_office_slides_read` | Read bounded plain text from each PPTX slide without decoding media. | Read-only; no approval. |
| `crawshrimp_office_validate` | Validate DOCX/PPTX required parts and optional expected text, table counts, or slide counts. | Read-only; no approval. |

Every filename must be relative to the workspace. The plugin resolves real paths, rejects traversal and symlink escapes, never overwrites an existing output, and writes through a temporary file before an atomic no-overwrite link.

## Examples

Create a Word document:

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

Create a Word document from Markdown:

```json
{
  "filename": "reports/weekly.docx",
  "title": "Weekly brief",
  "markdown": "# Weekly brief\n\n- Orders: 12\n- Returns: 1\n\n| Metric | Value |\n| --- | --- |\n| Conversion | 4.2% |"
}
```

Create a PPTX:

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

More arguments are in [`examples/word-from-markdown.json`](./examples/word-from-markdown.json), [`examples/slides-read.json`](./examples/slides-read.json), and [`examples/validate.json`](./examples/validate.json).

Markdown conversion accepts headings, paragraphs, ordered and unordered lists, blockquotes, fenced code as plain text, and simple pipe tables. Inline emphasis and links are flattened to text. HTML is stripped with warnings; image and media syntax is kept as alt text with warnings. The plugin never fetches or embeds external media and never executes HTML. Unsupported Markdown is retained as safe plain text when possible.

`crawshrimp_office_slides_read` returns `slideCount` plus bounded text for each slide and reads OOXML text nodes only. `crawshrimp_office_validate` returns `valid: false` when expected text or counts are not met; malformed ZIPs, missing required OOXML parts, invalid XML, and DTD/entity declarations return a truthful tool error.

## Permissions and bounds

Write calls use the host's native DSH permission signals. When a session reports both `danger-full-access` and native approval policy `never`, the existing Full Access grant allows the write without a duplicate per-call prompt. In `ask`, restricted modes, or when the approval/sandbox service is missing or cannot be read, the native approval path remains in place and execution proceeds only after `allowed-once`. There is no plugin setting that bypasses native approval.

The defaults cap each input file at 16 MiB, each generated output at 16 MiB, Markdown input at 512 KiB, and extracted text at 128 KiB. Paragraph, table, slide, bullet, Markdown block, and ZIP-entry counts are capped too. The plugin config can adjust `workspaceDir`, `maxInputBytes`, `maxOutputBytes`, `maxTextBytes`, and `maxMarkdownBytes` within bounded limits.

## Honest limits

The implementation generates, reads, and structurally validates OOXML packages. It does not render to PDF or PNG, run LibreOffice, or perform visual review. A successful create result reports `visualReview: "not_rendered"`; it must not be described as visually verified until a separate renderer and human or model review have actually run. `office_run`, Office COM, macros, formulas, charts, media decoding, and arbitrary Python execution are outside this package.

The examples and tests use local temporary files only. They do not call a model, merchant account, or external network service. This package does not edit arbitrary complex Office layouts, render charts, or recalculate formulas in existing documents.

## Related functional packages

The three maintained packages install independently: [Browser](https://github.com/howtimeschange/dsh-crawshrimp-browser) handles webpage operations, login state, and page scripts; [Data](https://github.com/howtimeschange/dsh-crawshrimp-data) handles records and tables; Office handles document jobs. PDF, labels, templates, and rendering remain on the Office roadmap and are not implemented in 0.2.0.

Version 0.2.0 adds Markdown-to-DOCX, per-slide PPTX reading, structural/content validation, and native DSH Full Access permission mapping.
