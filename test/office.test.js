import test from 'node:test'
import assert from 'node:assert/strict'
import { access, mkdtemp, readFile, realpath, rm, symlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Context } from '@deepseek-ai/cordis'
import { ToolRuntime } from '@deepseek-ai/dsh-tools'
import AdmZip from 'adm-zip'

import { apply, createSlidesBuffer, createWordBuffer, inspectOfficeFile, readWordText } from '../lib/index.js'

async function makeHarness(t, workspaceDir, approval = undefined) {
  const ctx = new Context()
  const promptDispose = ctx.provide('systemPrompt', { tools: () => () => {}, section: () => () => {} })
  const approvalDispose = approval ? ctx.provide('approval', approval) : undefined
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

test('real Cordis lifecycle registers and disposes the four Office tools', async t => {
  const root = await mkdtemp(join(tmpdir(), 'dsh-office-lifecycle-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  const ctx = await makeHarness(t, root, { effectivePolicy: () => 'ask', request: async () => 'allowed-once' })
  assert.deepEqual(ctx.tools.schemas().map(schema => schema.name), [
    'crawshrimp_office_word_create',
    'crawshrimp_office_word_read',
    'crawshrimp_office_slides_create',
    'crawshrimp_office_inspect',
  ])
  assert.equal(ctx.tools.get('crawshrimp_office_word_create').name, 'crawshrimp_office_word_create')
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
  assert.equal((await inspectOfficeFile(root, { filename: 'source.pptx' }, options, new AbortController().signal)).type, 'pptx')
})
