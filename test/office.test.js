import test from 'node:test'
import assert from 'node:assert/strict'
import { access, mkdtemp, readFile, realpath, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Context } from '@deepseek-ai/cordis'
import { ToolRuntime } from '@deepseek-ai/dsh-tools'
import AdmZip from 'adm-zip'

import * as officeModule from '../lib/index.js'
import { apply, createMarkdownWordBuffer, createSlidesBuffer, createWordBuffer, inspectOfficeFile, readPptxSlides, readWordText, validateOfficeFile } from '../lib/index.js'

async function makeHarness(t, workspaceDir, approval = undefined, sandboxPolicy = undefined) {
  const ctx = new Context()
  const promptDispose = ctx.provide('systemPrompt', { tools: () => () => {}, section: () => () => {} })
  const approvalDispose = approval ? ctx.provide('approval', approval) : undefined
  const sandboxDispose = sandboxPolicy ? ctx.provide('sandboxPolicy', sandboxPolicy) : undefined
  const toolsFiber = ctx.plugin(ToolRuntime)
  await toolsFiber
  const pluginFiber = ctx.plugin({
    inject: ['tools'],
    async apply(pluginContext) {
      await apply(pluginContext, { workspaceDir })
    },
  })
  await pluginFiber
  t.after(async () => {
    await pluginFiber.dispose()
    await toolsFiber.dispose()
    if (approvalDispose) await approvalDispose()
    if (sandboxDispose) await sandboxDispose()
    await promptDispose()
  })
  return ctx
}

function execution(name, callId = name) {
  return {
    name,
    callId,
    agent: { id: 'office-test-agent', session: { id: 'office-test-session' } },
    signal: new AbortController().signal,
  }
}

async function invoke(ctx, name, args, callId = name) {
  return ctx.tools.execute({ name, callId, arguments: args, ...execution(name, callId) })
}

async function expectToolError(ctx, name, args, pattern, callId) {
  const response = await invoke(ctx, name, args, callId)
  assert.equal(response.isError, true)
  assert.match(response.error?.message || response.content?.[0]?.text || '', pattern)
  return response
}

function updateZipText(buffer, entryName, update) {
  const zip = new AdmZip(buffer)
  const source = zip.readAsText(entryName)
  zip.updateFile(entryName, Buffer.from(update(source), 'utf8'))
  return zip
}

test('direct Cordis module lifecycle registers and disposes all seven Office tools', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-lifecycle-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  assert.deepEqual(ctx.tools.schemas().map(schema => schema.name), [
    'crawshrimp_office_word_create',
    'crawshrimp_office_word_read',
    'crawshrimp_office_slides_create',
    'crawshrimp_office_inspect',
    'crawshrimp_office_word_from_markdown',
    'crawshrimp_office_slides_read',
    'crawshrimp_office_validate',
  ])
  assert.equal(ctx.tools.get('crawshrimp_office_word_create').name, 'crawshrimp_office_word_create')
})

test('direct ctx.plugin(module, config) loads and unloads the module namespace', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-direct-module-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = new Context()
  const promptDispose = ctx.provide('systemPrompt', { tools: () => () => {}, section: () => () => {} })
  const toolsFiber = ctx.plugin(ToolRuntime)
  await toolsFiber
  const pluginFiber = ctx.plugin(officeModule, { workspaceDir: root })
  await pluginFiber
  assert.equal(ctx.tools.schemas().length, 7)
  await pluginFiber.dispose()
  assert.equal(ctx.tools.schemas().length, 0)
  await toolsFiber.dispose()
  await promptDispose()
})

test('writes and reads a real DOCX, then inspects its OOXML structure', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-docx-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const approvals = []
  const ctx = await makeHarness(t, root, {
    effectivePolicy: () => 'ask',
    async request(request) { approvals.push(request); return 'allowed-once' },
  })
  const created = await invoke(ctx, 'crawshrimp_office_word_create', {
    filename: 'reports/brief.docx',
    title: 'Weekly brief',
    paragraphs: ['A bounded paragraph.'],
    tables: [{ headers: ['Metric', 'Value'], rows: [['Orders', '12']] }],
  }, 'doc-create')
  assert.equal(created.isError, false)
  const output = await realpath(join(root, 'reports/brief.docx'))
  assert.ok((await readFile(output)).subarray(0, 2).toString() === 'PK')

  const read = await invoke(ctx, 'crawshrimp_office_word_read', { filename: 'reports/brief.docx' }, 'doc-read')
  assert.equal(read.isError, false)
  assert.match(read.value.structuredContent.text, /Weekly brief/)
  assert.match(read.value.structuredContent.text, /Orders/)
  assert.equal(approvals.length, 1, 'pure read must not request approval')

  const inspected = await invoke(ctx, 'crawshrimp_office_inspect', { filename: 'reports/brief.docx' }, 'doc-inspect')
  assert.equal(inspected.value.structuredContent.type, 'docx')
  assert.ok(inspected.value.structuredContent.entryCount > 0)
  assert.equal(inspected.value.structuredContent.counts.tables, 1)
})

test('writes fixed-width DOCX tables with readable mixed-script content', async () => {
  const fixtures = [
    {
      label: 'word_create',
      buffer: await createWordBuffer({
        filename: 'table.docx',
        title: 'Table fixture',
        tables: [{ headers: ['Metric 指标', 'Value 数值'], rows: [['Orders 订单', '12']] }],
      }),
    },
    {
      label: 'markdown',
      buffer: (await createMarkdownWordBuffer({
        filename: 'table.docx',
        markdown: '| Metric 指标 | Value 数值 |\n| --- | --- |\n| Orders 订单 | 12 |',
      })).buffer,
    },
  ]
  for (const fixture of fixtures) {
    const xml = new AdmZip(fixture.buffer).readAsText('word/document.xml')
    const table = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/u)?.[0]
    assert.ok(table, `${fixture.label} should contain a table`)
    assert.match(table, /<w:tblW\b[^>]*w:type="dxa"[^>]*w:w="9000"/u)
    assert.match(table, /<w:tblLayout\b[^>]*w:type="fixed"/u)
    const grid = table.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/u)?.[0]
    assert.ok(grid, `${fixture.label} should contain a table grid`)
    const widths = [...grid.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"/gu)].map(match => Number(match[1]))
    assert.equal(widths.length, 2)
    assert.equal(widths.reduce((total, width) => total + width, 0), 9000)
    assert.ok(widths.every(width => width > 100))
    assert.match(table, /Metric 指标/u)
    assert.match(table, /Value 数值/u)
    assert.match(table, /Orders 订单/u)
    assert.match(table, />12<\/w:t>/u)
  }
})

test('caps Office text at UTF-8 boundaries and keeps tab semantics consistent', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-text-semantics-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const options = { maxInputBytes: 2 * 1024 * 1024, maxTextBytes: 5 }
  const docxZip = updateZipText(
    await createWordBuffer({ filename: 'tab.docx', title: 'Source' }),
    'word/document.xml',
    xml => xml.replace(/(<w:t[^>]*>)Source(<\/w:t>)/u, '$1A$2<w:tab/><w:t>B</w:t>'),
  )
  await writeFile(join(root, 'tab.docx'), docxZip.toBuffer())
  const tabRead = await readWordText(root, { filename: 'tab.docx' }, { ...options, maxTextBytes: 100 })
  assert.match(tabRead.text, /A\tB/u)
  const tabValidation = await validateOfficeFile(root, { filename: 'tab.docx', expectedText: 'A\tB' }, { ...options, maxTextBytes: 100 })
  assert.equal(tabValidation.valid, true)
  assert.match(tabValidation.text, /A\tB/u)

  const unicodeDocx = await createWordBuffer({ filename: 'unicode.docx', title: 'A😀中' })
  await writeFile(join(root, 'unicode.docx'), unicodeDocx)
  const boundedDoc = await readWordText(root, { filename: 'unicode.docx' }, options)
  assert.ok(Buffer.byteLength(boundedDoc.text, 'utf8') <= options.maxTextBytes)
  assert.doesNotMatch(boundedDoc.text, /�/u)
  assert.equal(boundedDoc.text, 'A😀')

  const unicodePptx = await createSlidesBuffer({ filename: 'unicode.pptx', title: 'A😀中', slides: [] })
  await writeFile(join(root, 'unicode.pptx'), unicodePptx)
  const boundedSlides = await readPptxSlides(root, { filename: 'unicode.pptx' }, options)
  assert.ok(Buffer.byteLength(boundedSlides.slides[0].text, 'utf8') <= options.maxTextBytes)
  assert.doesNotMatch(boundedSlides.slides[0].text, /�/u)
  assert.equal(boundedSlides.slides[0].text, 'A😀')
})

test('converts bounded Markdown into a real DOCX with warnings for media and HTML', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-markdown-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  const markdown = [
    '# 周报 📦',
    '',
    '这是 **Unicode** 正文，包含 [链接文字](https://example.com)。',
    '',
    '- 设计',
    '- 评审',
    '',
    '1. 发布',
    '2. 复盘',
    '',
    '| 指标 | 数值 |',
    '| --- | --- |',
    '| 订单 | 12 |',
    '',
    '![商品图](https://example.com/product.png) <strong>安全纯文本</strong>',
  ].join('\n')
  const created = await invoke(ctx, 'crawshrimp_office_word_from_markdown', {
    filename: 'reports/weekly.docx',
    title: 'Markdown 导出',
    markdown,
  }, 'markdown-create')
  assert.equal(created.isError, false)
  assert.ok(created.value.structuredContent.warnings.some(warning => warning.code === 'UNSUPPORTED_MEDIA'))
  assert.ok(created.value.structuredContent.warnings.some(warning => warning.code === 'HTML_STRIPPED'))
  assert.equal((await readFile(join(root, 'reports/weekly.docx'))).subarray(0, 2).toString(), 'PK')
  const read = await invoke(ctx, 'crawshrimp_office_word_read', { filename: 'reports/weekly.docx' }, 'markdown-read')
  assert.match(read.value.structuredContent.text, /周报 📦/u)
  assert.match(read.value.structuredContent.text, /设计/u)
  assert.match(read.value.structuredContent.text, /订单/u)
  assert.match(read.value.structuredContent.text, /商品图/u)
  assert.doesNotMatch(read.value.structuredContent.text, /example\.com\/product\.png/u)
  const valid = await invoke(ctx, 'crawshrimp_office_validate', {
    filename: 'reports/weekly.docx',
    expectedText: '订单',
    expectedTableCount: 1,
  }, 'markdown-validate')
  assert.equal(valid.value.structuredContent.valid, true)
  const invalid = await invoke(ctx, 'crawshrimp_office_validate', {
    filename: 'reports/weekly.docx',
    expectedText: '不存在的内容',
    expectedTableCount: 2,
  }, 'markdown-validate-negative')
  assert.equal(invalid.isError, false)
  assert.equal(invalid.value.structuredContent.valid, false)
  assert.deepEqual(invalid.value.structuredContent.issues.map(issue => issue.code), ['EXPECTED_TEXT_NOT_FOUND', 'TABLE_COUNT_MISMATCH'])
  await assert.rejects(
    createMarkdownWordBuffer({ filename: 'too-large.docx', markdown: '中文' }, { maxMarkdownBytes: 2 }),
    error => error?.code === 'INPUT_TOO_LARGE',
  )
})

test('writes a real PPTX and exposes slide text in the OOXML package', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-pptx-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  const created = await invoke(ctx, 'crawshrimp_office_slides_create', {
    filename: 'deck.pptx',
    title: 'Launch plan',
    slides: [{ heading: 'Milestones', bullets: ['Design', 'Review'] }],
  }, 'slides-create')
  assert.equal(created.isError, false)
  const zip = new AdmZip(join(root, 'deck.pptx'))
  const slideXml = zip.readAsText('ppt/slides/slide2.xml')
  assert.match(slideXml, /Milestones/)
  assert.match(slideXml, /Design/)
  const inspected = await invoke(ctx, 'crawshrimp_office_inspect', { filename: 'deck.pptx' }, 'slides-inspect')
  assert.equal(inspected.value.structuredContent.type, 'pptx')
  assert.equal(inspected.value.structuredContent.counts.slides, 2)
})

test('reads PPTX slide text and validates positive and negative expectations', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-slides-read-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  const created = await invoke(ctx, 'crawshrimp_office_slides_create', {
    filename: 'unicode.pptx',
    title: '发布计划 🚀',
    slides: [{ heading: '里程碑', bullets: ['设计', '评审'] }],
  }, 'slides-read-source')
  assert.equal(created.isError, false)
  const read = await invoke(ctx, 'crawshrimp_office_slides_read', { filename: 'unicode.pptx' }, 'slides-read')
  assert.equal(read.value.structuredContent.slideCount, 2)
  assert.equal(read.value.structuredContent.slides[0].text, '发布计划 🚀')
  assert.match(read.value.structuredContent.slides[1].text, /里程碑\n设计\n评审/u)
  const valid = await invoke(ctx, 'crawshrimp_office_validate', {
    filename: 'unicode.pptx',
    expectedText: '评审',
    expectedSlideCount: 2,
  }, 'slides-validate')
  assert.equal(valid.value.structuredContent.valid, true)
  const invalid = await invoke(ctx, 'crawshrimp_office_validate', {
    filename: 'unicode.pptx',
    expectedText: '缺失文本',
    expectedSlideCount: 3,
  }, 'slides-validate-negative')
  assert.equal(invalid.isError, false)
  assert.equal(invalid.value.structuredContent.valid, false)
  assert.deepEqual(invalid.value.structuredContent.issues.map(issue => issue.code), ['EXPECTED_TEXT_NOT_FOUND', 'SLIDE_COUNT_MISMATCH'])
})

test('rejects corrupt ZIPs, missing OOXML parts, and DTD/entity declarations', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-malformed-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const source = await createSlidesBuffer({ filename: 'source.pptx', title: 'Source', slides: [{ heading: 'One' }] })
  await writeFile(join(root, 'broken.pptx'), Buffer.from('this is not an Office ZIP'))
  const ctx = await makeHarness(t, root)
  await expectToolError(ctx, 'crawshrimp_office_slides_read', { filename: 'broken.pptx' }, /INVALID_DOCUMENT/, 'broken-zip')

  const missing = new AdmZip(source)
  missing.deleteFile('ppt/presentation.xml')
  await writeFile(join(root, 'missing.pptx'), missing.toBuffer())
  await expectToolError(ctx, 'crawshrimp_office_validate', { filename: 'missing.pptx' }, /required part: ppt\/presentation\.xml/, 'missing-presentation')

  const unsafe = new AdmZip(source)
  const slideXml = unsafe.readAsText('ppt/slides/slide1.xml')
  unsafe.updateFile('ppt/slides/slide1.xml', Buffer.from(slideXml.replace(/(<\?xml[^>]+>)/u, '$1\n<!DOCTYPE sld [<!ENTITY x "expanded">]>')))
  await writeFile(join(root, 'unsafe.pptx'), unsafe.toBuffer())
  await expectToolError(ctx, 'crawshrimp_office_slides_read', { filename: 'unsafe.pptx' }, /UNSAFE_XML/, 'unsafe-xml')

  const docx = new AdmZip(await createWordBuffer({ filename: 'source.docx', title: 'Source' }))
  docx.deleteFile('word/document.xml')
  await writeFile(join(root, 'missing.docx'), docx.toBuffer())
  await expectToolError(ctx, 'crawshrimp_office_validate', { filename: 'missing.docx' }, /required part: word\/document\.xml/, 'missing-document')
  const unsafeDocx = new AdmZip(await createWordBuffer({ filename: 'unsafe.docx', title: 'Source' }))
  const documentXml = unsafeDocx.readAsText('word/document.xml')
  unsafeDocx.updateFile('word/document.xml', Buffer.from(documentXml.replace(/(<\?xml[^>]+>)/u, '$1\n<!DOCTYPE document [<!ENTITY x "expanded">]>')))
  await writeFile(join(root, 'unsafe.docx'), unsafeDocx.toBuffer())
  await expectToolError(ctx, 'crawshrimp_office_word_read', { filename: 'unsafe.docx' }, /UNSAFE_XML/, 'unsafe-docx')
})

test('rejects malformed XML and incomplete OOXML package metadata', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-ooxml-structure-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const source = await createWordBuffer({ filename: 'source.docx', title: 'Source' })
  const options = { maxInputBytes: 2 * 1024 * 1024, maxTextBytes: 128 * 1024 }
  const cases = [
    {
      filename: 'control.docx',
      update: zip => updateZipText(zip, 'word/document.xml', xml => xml.replace('Source', 'S\u0001ource')),
      operation: 'read',
    },
    {
      filename: 'unbound-prefix.docx',
      update: zip => updateZipText(zip, 'word/document.xml', xml => xml.replace('<w:body>', '<bad:body>').replace('</w:body>', '</bad:body>')),
      operation: 'read',
    },
    {
      filename: 'missing-override.docx',
      update: zip => updateZipText(zip, '[Content_Types].xml', xml => xml.replace(/<Override ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document\.main\+xml" PartName="\/word\/document\.xml"\/>/u, '')),
      operation: 'validate',
    },
    {
      filename: 'wrong-main-relation.docx',
      update: zip => updateZipText(zip, '_rels/.rels', xml => xml.replace(/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/officeDocument"/u, 'Type="urn:wrong-main"')),
      operation: 'validate',
    },
    {
      filename: 'wrong-content-namespace.docx',
      update: zip => updateZipText(zip, '[Content_Types].xml', xml => xml.replace('xmlns="http://schemas.openxmlformats.org/package/2006/content-types"', 'xmlns="urn:wrong-content-types"')),
      operation: 'validate',
    },
    {
      filename: 'wrong-rels-namespace.docx',
      update: zip => updateZipText(zip, '_rels/.rels', xml => xml.replace('xmlns="http://schemas.openxmlformats.org/package/2006/relationships"', 'xmlns="urn:wrong-package-rels"')),
      operation: 'validate',
    },
  ]
  for (const fixture of cases) {
    const zip = fixture.update(new AdmZip(source))
    await writeFile(join(root, fixture.filename), zip.toBuffer())
    if (fixture.operation === 'read') {
      await assert.rejects(
        readWordText(root, { filename: fixture.filename }, options),
        error => error?.code === 'INVALID_DOCUMENT',
        fixture.filename,
      )
    } else {
      await assert.rejects(
        validateOfficeFile(root, { filename: fixture.filename }, options),
        error => error?.code === 'INVALID_DOCUMENT',
        fixture.filename,
      )
    }
  }
})

test('inspect counts parsed tables and relationship-linked slides only', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-inspect-counts-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const options = { maxInputBytes: 2 * 1024 * 1024, maxTextBytes: 128 * 1024 }

  const docx = updateZipText(
    await createWordBuffer({ filename: 'no-table.docx', title: 'No table' }),
    'word/document.xml',
    xml => xml.replace('</w:body>', '<!-- <w:tbl><w:tr/></w:tbl> --></w:body>'),
  )
  await writeFile(join(root, 'no-table.docx'), docx.toBuffer())
  const inspectedDocx = await inspectOfficeFile(root, { filename: 'no-table.docx' }, options, new AbortController().signal)
  assert.equal(inspectedDocx.counts.tables, 0)

  const pptx = new AdmZip(await createSlidesBuffer({ filename: 'orphan.pptx', title: 'Deck', slides: [{ heading: 'Linked' }] }))
  const presentationXml = pptx.readAsText('ppt/presentation.xml')
  const slideIds = [...presentationXml.matchAll(/<p:sldId\b[^>]*\/>/gu)]
  assert.equal(slideIds.length, 2)
  pptx.updateFile('ppt/presentation.xml', Buffer.from(presentationXml.replace(slideIds[1][0], ''), 'utf8'))
  await writeFile(join(root, 'orphan.pptx'), pptx.toBuffer())
  const inspectedPptx = await inspectOfficeFile(root, { filename: 'orphan.pptx' }, options, new AbortController().signal)
  assert.equal(inspectedPptx.counts.slides, 1)
})

test('fails closed when native approval rejects and when approval is absent', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-denied-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const requests = []
  const denied = await makeHarness(t, root, {
    // The native ApprovalService owns policy enforcement. This mock models its
    // `never` response and confirms the plugin still asks it for a decision.
    effectivePolicy: () => 'never',
    request: async request => { requests.push(request); return 'denied' },
  })
  await expectToolError(denied, 'crawshrimp_office_word_create', { filename: 'policy.docx', title: 'Native policy' }, /APPROVAL_REQUIRED/, 'policy-never')
  assert.equal(requests.length, 1)
  assert.equal(requests[0].toolName, 'crawshrimp_office_word_create')
  const absent = await makeHarness(t, root)
  await expectToolError(absent, 'crawshrimp_office_slides_create', { filename: 'denied.pptx', title: 'Nope' }, /APPROVAL_REQUIRED/, 'denied-absent')
})

test('maps native Full Access to one standing grant and keeps all other states fail closed', async t => {
  const fullRoot = await mkdtemp(join(tmpdir(), 'dsh-office-full-access-'))
  t.after(() => rm(fullRoot, { recursive: true, force: true }))
  let fullRequests = 0
  const full = await makeHarness(t, fullRoot, {
    effectivePolicy: () => 'never',
    request: async () => { fullRequests += 1; return 'rejected' },
  }, { resolve: ({ session }) => ({ mode: 'danger-full-access', sessionId: session.id }) })
  const fullResult = await invoke(full, 'crawshrimp_office_word_create', {
    filename: 'full-access.docx', title: 'Full Access',
  }, 'full-access')
  assert.equal(fullResult.isError, false, JSON.stringify(fullResult))
  assert.equal(fullRequests, 0, 'Full Access must not open a second approval request')
  await access(join(fullRoot, 'full-access.docx'))

  const askRoot = await mkdtemp(join(tmpdir(), 'dsh-office-full-ask-'))
  t.after(() => rm(askRoot, { recursive: true, force: true }))
  let askRequests = 0
  const ask = await makeHarness(t, askRoot, {
    effectivePolicy: () => 'ask',
    request: async () => { askRequests += 1; return 'allowed-once' },
  }, { resolve: () => ({ mode: 'danger-full-access' }) })
  const askResult = await invoke(ask, 'crawshrimp_office_word_create', {
    filename: 'ask.docx', title: 'Ask policy',
  }, 'full-ask')
  assert.equal(askResult.isError, false, JSON.stringify(askResult))
  assert.equal(askRequests, 1, 'danger-full-access with ask still uses native approval')

  const restrictedRoot = await mkdtemp(join(tmpdir(), 'dsh-office-restricted-never-'))
  t.after(() => rm(restrictedRoot, { recursive: true, force: true }))
  let restrictedRequests = 0
  const restricted = await makeHarness(t, restrictedRoot, {
    effectivePolicy: () => 'never',
    request: async () => { restrictedRequests += 1; return 'rejected' },
  }, { resolve: () => ({ mode: 'workspace-write' }) })
  await expectToolError(restricted, 'crawshrimp_office_word_create', {
    filename: 'restricted.docx', title: 'Restricted',
  }, /APPROVAL_REQUIRED/, 'restricted-never')
  assert.equal(restrictedRequests, 1, 'never with a restricted sandbox must not bypass the approval path')

  const missingRoot = await mkdtemp(join(tmpdir(), 'dsh-office-missing-approval-'))
  t.after(() => rm(missingRoot, { recursive: true, force: true }))
  const missing = await makeHarness(t, missingRoot, undefined, { resolve: () => ({ mode: 'danger-full-access' }) })
  await expectToolError(missing, 'crawshrimp_office_word_create', {
    filename: 'missing.docx', title: 'Missing approval',
  }, /APPROVAL_REQUIRED/, 'missing-approval')

  const errorRoot = await mkdtemp(join(tmpdir(), 'dsh-office-broken-sandbox-'))
  t.after(() => rm(errorRoot, { recursive: true, force: true }))
  let errorRequests = 0
  const broken = await makeHarness(t, errorRoot, {
    effectivePolicy: () => 'never',
    request: async () => { errorRequests += 1; return 'rejected' },
  }, { resolve: () => { throw new Error('sandbox unavailable') } })
  await expectToolError(broken, 'crawshrimp_office_word_create', {
    filename: 'broken.docx', title: 'Broken sandbox',
  }, /APPROVAL_REQUIRED/, 'broken-sandbox')
  assert.equal(errorRequests, 1, 'permission detection errors must fall back to approval, never to an automatic grant')
})

test('rejects unsupported image fields at the tool boundary', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-schema-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  await expectToolError(ctx, 'crawshrimp_office_slides_create', {
    filename: 'image.pptx',
    title: 'Text only',
    image: { path: '/tmp/secret.png' },
  }, /unsupported field: image/, 'unsupported-image')
})

test('rejects traversal, symlink escape, and overwrite attempts', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-paths-'))
  const outside = await mkdtemp(join(tmpdir(), 'dsh-office-outside-'))
  t.after(() => Promise.all([
    rm(root, { recursive: true, force: true }),
    rm(outside, { recursive: true, force: true }),
  ]))
  await symlink(outside, join(root, 'linked'))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  await expectToolError(ctx, 'crawshrimp_office_word_create', { filename: '../escape.docx', title: 'Escape' }, /PATH_OUTSIDE_WORKSPACE/, 'traversal')
  await expectToolError(ctx, 'crawshrimp_office_word_create', { filename: 'linked/escape.docx', title: 'Escape' }, /PATH_OUTSIDE_WORKSPACE/, 'symlink')
  await invoke(ctx, 'crawshrimp_office_word_create', { filename: 'same.docx', title: 'First' }, 'first')
  await expectToolError(ctx, 'crawshrimp_office_word_create', { filename: 'same.docx', title: 'Second' }, /OUTPUT_EXISTS/, 'second')
  await assert.rejects(access(join(outside, 'escape.docx')), { code: 'ENOENT' })
})

test('helper APIs enforce the configured input and output boundaries', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-helpers-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const options = { maxInputBytes: 2 * 1024 * 1024, maxTextBytes: 64 }
  const doc = await createWordBuffer({ filename: 'source.docx', title: 'Source', paragraphs: ['text'] })
  const slides = await createSlidesBuffer({ filename: 'source.pptx', title: 'Source', slides: [{ heading: 'One', bullets: ['text'] }] })
  await (await import('node:fs/promises')).writeFile(join(root, 'source.docx'), doc)
  await (await import('node:fs/promises')).writeFile(join(root, 'source.pptx'), slides)
  assert.match((await readWordText(root, { filename: 'source.docx' }, options)).text, /Source/)
  assert.equal((await readPptxSlides(root, { filename: 'source.pptx' }, options)).slideCount, 2)
  assert.equal((await validateOfficeFile(root, { filename: 'source.pptx', expectedSlideCount: 2 }, options)).valid, true)
  assert.equal((await inspectOfficeFile(root, { filename: 'source.pptx' }, options, new AbortController().signal)).type, 'pptx')
})
