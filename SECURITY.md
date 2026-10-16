# Security

## Exposed input and output surface

The plugin exposes four bounded tools:

- `crawshrimp_office_word_create` accepts a title, paragraphs, and simple string tables and writes a new `.docx`.
- `crawshrimp_office_slides_create` accepts a title plus heading and bullet strings and writes a new `.pptx`.
- `crawshrimp_office_word_read` extracts capped plain text from an existing `.docx`.
- `crawshrimp_office_inspect` reads bounded ZIP entry metadata from an existing `.docx` or `.pptx`.

Create and read arguments reject unknown fields, including `image`, `media`, `path`, and `url`. The plugin does not fetch remote resources, accept image or media bytes, invoke a media decoder, run a macro, or invoke the Mammoth CLI. Existing Office packages may contain media entries, but inspection reports their metadata and counts without decoding them.

Writes require the host's native DSH approval service to return `allowed-once`. The plugin enforces relative workspace paths, rejects traversal and symlink escapes, caps input and output sizes, refuses overwrites, and writes output atomically.

## Current npm audit findings

As of 2026-10-16, `npm audit --omit=dev` reports four vulnerable package nodes across these upstream advisories:

- [`GHSA-5p2g-fcmc-qvqq`](https://github.com/advisories/GHSA-5p2g-fcmc-qvqq) and [`GHSA-w3rx-r6r6-pgpr`](https://github.com/advisories/GHSA-w3rx-r6r6-pgpr) affect `image-size` versions `0.6.3` through `2.0.2`. The installed `image-size@1.2.1` is brought in by `pptxgenjs@4.0.0`. The exported slides tool accepts text only and this plugin does not call the image APIs, so those parsers are not reachable through the exposed tool surface.
- [`GHSA-hp3w-g68c-fv3c`](https://github.com/advisories/GHSA-hp3w-g68c-fv3c) affects `sprintf-js` versions through `1.1.3`. The reported chain is `mammoth@1.13.0` → `argparse@1.0.10` → `sprintf-js@1.0.3`; `argparse` is declared for Mammoth's command-line entry point. This plugin imports Mammoth's library entry point and calls `extractRawText` with an in-memory DOCX buffer; it does not invoke that CLI or parse command-line input.

These findings remain present in the dependency graph. This document records the demonstrated reachability boundary; it does not claim that `npm audit` is clean or that the dependencies are vulnerability-free. Keep the versions pinned until compatible upstream releases are available, then rerun the package tests and audit before updating them.
