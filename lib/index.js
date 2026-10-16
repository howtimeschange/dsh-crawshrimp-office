import { randomUUID } from 'node:crypto'
import { mkdir, readFile, realpath, stat, unlink, writeFile, link, lstat } from 'node:fs/promises'
import { dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'

import z from '@deepseek-ai/schemastery'
import { createMcpToolDefinition } from '@deepseek-ai/dsh-mcp-client'
import AdmZip from 'adm-zip'
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx'
import mammoth from 'mammoth'
import pptxgen from 'pptxgenjs'

export const name = 'crawshrimp-office'
export const inject = ['tools']

const DEFAULT_MAX_INPUT_BYTES = 16 * 1024 * 1024
const DEFAULT_MAX_OUTPUT_BYTES = 16 * 1024 * 1024
const DEFAULT_MAX_TEXT_BYTES = 128 * 1024
const MAX_FILENAME_LENGTH = 240
const MAX_TITLE_LENGTH = 240
const MAX_PARAGRAPHS = 100
const MAX_PARAGRAPH_LENGTH = 12_000
const MAX_TABLES = 20
const MAX_TABLE_ROWS = 100
const MAX_TABLE_COLUMNS = 30
const MAX_CELL_LENGTH = 2_000
const MAX_SLIDES = 100
const MAX_BULLETS = 30
const MAX_BULLET_LENGTH = 2_000
const MAX_ZIP_ENTRIES = 2_000
const MAX_RETURNED_ENTRIES = 500
const MAX_ARGUMENT_BYTES = 2 * 1024 * 1024

export const Config = z.object({
  workspaceDir: z.string().default('./crawshrimp-office-output'),
  maxInputBytes: z.natural().default(DEFAULT_MAX_INPUT_BYTES),
  maxOutputBytes: z.natural().default(DEFAULT_MAX_OUTPUT_BYTES),
  maxTextBytes: z.natural().default(DEFAULT_MAX_TEXT_BYTES),
})

export class OfficePluginError extends Error {
  constructor(code, message) {
    super(message)
    this.name = 'CrawshrimpOfficeError'
    this.code = code
  }
}

function fail(code, message) {
  throw new OfficePluginError(code, message)
}

function checkAbort(signal) {
  signal?.throwIfAborted?.()
}

function asObject(value, label) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) fail('INVALID_ARGUMENTS', `${label} must be an object`)
  return value
}

function assertKnownKeys(value, keys, label) {
  const allowed = new Set(keys)
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) fail('INVALID_ARGUMENTS', `${label} contains unsupported field: ${key}`)
  }
}

function asString(value, label, { max = 2_000, required = true } = {}) {
  if (typeof value !== 'string' || (required && !value.trim()) || value.length > max || /[\u0000-\u001f]/u.test(value)) {
    fail('INVALID_ARGUMENTS', `${label} must be a string of at most ${max} characters`)
  }
  return value
}

function asArray(value, label, max) {
  if (!Array.isArray(value) || value.length > max) fail('INVALID_ARGUMENTS', `${label} must contain at most ${max} items`)
  return value
}

function asSafeInteger(value, label, { min = 0, max = Number.MAX_SAFE_INTEGER } = {}) {
  if (!Number.isSafeInteger(value) || value < min || value > max) fail('INVALID_ARGUMENTS', `${label} is outside its allowed range`)
  return value
}

function normaliseOptions(config = {}) {
  const options = {
    workspaceDir: typeof config.workspaceDir === 'string' && config.workspaceDir.trim()
      ? config.workspaceDir
      : './crawshrimp-office-output',
    maxInputBytes: config.maxInputBytes ?? DEFAULT_MAX_INPUT_BYTES,
    maxOutputBytes: config.maxOutputBytes ?? DEFAULT_MAX_OUTPUT_BYTES,
    maxTextBytes: config.maxTextBytes ?? DEFAULT_MAX_TEXT_BYTES,
  }
  asSafeInteger(options.maxInputBytes, 'maxInputBytes', { min: 1, max: 128 * 1024 * 1024 })
  asSafeInteger(options.maxOutputBytes, 'maxOutputBytes', { min: 1, max: 128 * 1024 * 1024 })
  asSafeInteger(options.maxTextBytes, 'maxTextBytes', { min: 1, max: 8 * 1024 * 1024 })
  return options
}

async function ensureWorkspace(workspaceDir) {
  const absolute = resolve(workspaceDir)
  await mkdir(absolute, { recursive: true })
  return realpath(absolute)
}

function assertRelativeFilename(filename) {
  asString(filename, 'filename', { max: MAX_FILENAME_LENGTH })
  if (isAbsolute(filename) || filename.split(/[\\/]/u).some(part => part === '..')) {
    fail('PATH_OUTSIDE_WORKSPACE', 'filename must be a relative path inside workspaceDir')
  }
  const normalised = resolve('/safe-root', filename)
  if (normalised === '/safe-root' || normalised.startsWith(`/safe-root${sep}`) === false) {
    fail('PATH_OUTSIDE_WORKSPACE', 'filename must identify a file inside workspaceDir')
  }
  return filename
}

function assertContained(root, candidate, { allowRoot = false } = {}) {
  const rel = relative(root, candidate)
  if ((!allowRoot && rel === '') || rel === '..' || rel.startsWith(`..${sep}`) || isAbsolute(rel)) {
    fail('PATH_OUTSIDE_WORKSPACE', 'path is outside workspaceDir')
  }
}

async function existingAncestor(pathname) {
  let current = pathname
  while (true) {
    try {
      return await realpath(current)
    } catch (error) {
      if (error?.code !== 'ENOENT') throw error
      const parent = dirname(current)
      if (parent === current) throw error
      current = parent
    }
  }
}

export async function resolveWorkspacePath(workspaceRoot, filename, { forWrite = false } = {}) {
  const canonicalRoot = await realpath(workspaceRoot)
  assertRelativeFilename(filename)
  const candidate = resolve(canonicalRoot, filename)
  assertContained(canonicalRoot, candidate)
  if (forWrite) {
    const ancestor = await existingAncestor(dirname(candidate))
    assertContained(canonicalRoot, ancestor, { allowRoot: true })
    await mkdir(dirname(candidate), { recursive: true })
    const parent = await realpath(dirname(candidate))
    assertContained(canonicalRoot, parent, { allowRoot: true })
    return { candidate, parent }
  }
  let resolved
  try {
    resolved = await realpath(candidate)
  } catch (error) {
    if (error?.code === 'ENOENT') fail('FILE_NOT_FOUND', `file does not exist: ${filename}`)
    throw error
  }
  assertContained(canonicalRoot, resolved)
  const info = await stat(resolved)
  if (!info.isFile()) fail('INVALID_FILE', `not a regular file: ${filename}`)
  return { candidate: resolved, parent: dirname(resolved), info }
}

async function assertNewOutput(target, root) {
  assertContained(root, target)
  try {
    await lstat(target)
    fail('OUTPUT_EXISTS', `refusing to overwrite existing output: ${relative(root, target)}`)
  } catch (error) {
    if (error?.code !== 'ENOENT') throw error
  }
}

async function atomicWrite(root, filename, bytes, signal, maxOutputBytes) {
  checkAbort(signal)
  if (!Buffer.isBuffer(bytes)) bytes = Buffer.from(bytes)
  if (bytes.byteLength > maxOutputBytes) fail('OUTPUT_TOO_LARGE', `output exceeds ${maxOutputBytes} bytes`)
  const { candidate, parent } = await resolveWorkspacePath(root, filename, { forWrite: true })
  await assertNewOutput(candidate, root)
  checkAbort(signal)
  const temporary = join(parent, `.${candidate.split(sep).pop()}.${randomUUID()}.tmp`)
  try {
    await writeFile(temporary, bytes, { flag: 'wx', mode: 0o600 })
    checkAbort(signal)
    // link() creates the destination atomically and fails if another writer won
    // the name between the preflight and this point. It avoids rename-overwrite.
    await link(temporary, candidate)
    await unlink(temporary)
    return candidate
  } finally {
    await unlink(temporary).catch(error => {
      if (error?.code !== 'ENOENT') throw error
    })
  }
}

function truncateUtf8(value, maxBytes) {
  const bytes = Buffer.from(value, 'utf8')
  if (bytes.byteLength <= maxBytes) return { value, truncated: false }
  return { value: bytes.subarray(0, maxBytes).toString('utf8'), truncated: true }
}

function cleanText(value, label, max) {
  return asString(String(value ?? ''), label, { max, required: false })
}

function assertPayloadSize(value) {
  let bytes
  try {
    bytes = Buffer.byteLength(JSON.stringify(value), 'utf8')
  } catch {
    fail('INVALID_ARGUMENTS', 'arguments must be JSON serializable')
  }
  if (bytes > MAX_ARGUMENT_BYTES) fail('INPUT_TOO_LARGE', `arguments exceed ${MAX_ARGUMENT_BYTES} bytes`)
}

function validateWordCreate(input) {
  const value = asObject(input, 'arguments')
  assertPayloadSize(value)
  assertKnownKeys(value, ['filename', 'title', 'paragraphs', 'tables'], 'arguments')
  const filename = asString(value.filename, 'filename', { max: MAX_FILENAME_LENGTH })
  if (extname(filename).toLowerCase() !== '.docx') fail('INVALID_ARGUMENTS', 'filename must end in .docx')
  const title = cleanText(value.title, 'title', MAX_TITLE_LENGTH)
  const paragraphs = asArray(value.paragraphs ?? [], 'paragraphs', MAX_PARAGRAPHS).map((item, index) => cleanText(item, `paragraphs[${index}]`, MAX_PARAGRAPH_LENGTH))
  const tables = asArray(value.tables ?? [], 'tables', MAX_TABLES).map((table, tableIndex) => {
    const object = asObject(table, `tables[${tableIndex}]`)
    assertKnownKeys(object, ['headers', 'rows'], `tables[${tableIndex}]`)
    const headers = object.headers === undefined ? [] : asArray(object.headers, `tables[${tableIndex}].headers`, MAX_TABLE_COLUMNS).map((cell, index) => cleanText(cell, `tables[${tableIndex}].headers[${index}]`, MAX_CELL_LENGTH))
    const rows = asArray(object.rows ?? [], `tables[${tableIndex}].rows`, MAX_TABLE_ROWS).map((row, rowIndex) => {
      const cells = asArray(row, `tables[${tableIndex}].rows[${rowIndex}]`, MAX_TABLE_COLUMNS).map((cell, cellIndex) => cleanText(cell, `tables[${tableIndex}].rows[${rowIndex}][${cellIndex}]`, MAX_CELL_LENGTH))
      if (headers.length && cells.length !== headers.length) fail('INVALID_ARGUMENTS', `table ${tableIndex} row ${rowIndex} does not match headers`)
      return cells
    })
    if (!headers.length && rows.some(row => row.length === 0)) fail('INVALID_ARGUMENTS', `table ${tableIndex} contains an empty row`)
    if (headers.length === 0 && rows.length === 0) fail('INVALID_ARGUMENTS', `table ${tableIndex} is empty`)
    return { headers, rows }
  })
  if (!title && !paragraphs.length && !tables.length) fail('INVALID_ARGUMENTS', 'document must contain a title, paragraph, or table')
  return { filename, title, paragraphs, tables }
}

function validateSlidesCreate(input) {
  const value = asObject(input, 'arguments')
  assertPayloadSize(value)
  assertKnownKeys(value, ['filename', 'title', 'slides'], 'arguments')
  const filename = asString(value.filename, 'filename', { max: MAX_FILENAME_LENGTH })
  if (extname(filename).toLowerCase() !== '.pptx') fail('INVALID_ARGUMENTS', 'filename must end in .pptx')
  const title = asString(value.title, 'title', { max: MAX_TITLE_LENGTH })
  const slides = asArray(value.slides ?? [], 'slides', MAX_SLIDES).map((slide, index) => {
    const object = asObject(slide, `slides[${index}]`)
    assertKnownKeys(object, ['heading', 'bullets'], `slides[${index}]`)
    const heading = asString(object.heading, `slides[${index}].heading`, { max: MAX_TITLE_LENGTH })
    const bullets = asArray(object.bullets ?? [], `slides[${index}].bullets`, MAX_BULLETS).map((bullet, bulletIndex) => cleanText(bullet, `slides[${index}].bullets[${bulletIndex}]`, MAX_BULLET_LENGTH))
    return { heading, bullets }
  })
  return { filename, title, slides }
}

function validateRead(input) {
  const value = asObject(input, 'arguments')
  assertKnownKeys(value, ['filename'], 'arguments')
  return { filename: asString(value.filename, 'filename', { max: MAX_FILENAME_LENGTH }) }
}

function createParagraph(text, heading = false) {
  return new Paragraph({
    children: [new TextRun({ text })],
    ...(heading ? { heading: HeadingLevel.TITLE } : {}),
  })
}

export async function createWordBuffer(input) {
  const value = validateWordCreate(input)
  const children = []
  if (value.title) children.push(createParagraph(value.title, true))
  for (const paragraph of value.paragraphs) children.push(createParagraph(paragraph))
  for (const table of value.tables) {
    const rows = []
    if (table.headers.length) rows.push(table.headers)
    rows.push(...table.rows)
    children.push(new Table({
      rows: rows.map(row => new TableRow({
        children: row.map(cell => new TableCell({ children: [createParagraph(cell)] })),
      })),
    }))
  }
  return Packer.toBuffer(new Document({ sections: [{ children }] }))
}

export async function createSlidesBuffer(input) {
  const value = validateSlidesCreate(input)
  const pptx = new pptxgen()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'dsh-crawshrimp-office'
  pptx.subject = 'OOXML presentation generated by DeepSeek Harness'
  pptx.title = value.title
  const titleSlide = pptx.addSlide()
  titleSlide.addText(value.title, { x: 0.7, y: 2.2, w: 12, h: 0.8, fontSize: 30, bold: true, color: '1F2937', margin: 0 })
  for (const item of value.slides) {
    const slide = pptx.addSlide()
    slide.addText(item.heading, { x: 0.7, y: 0.45, w: 12, h: 0.55, fontSize: 24, bold: true, color: '1F2937', margin: 0 })
    if (item.bullets.length) {
      slide.addText(item.bullets.map(text => ({ text, options: { bullet: { indent: 16 } } })), {
        x: 0.95, y: 1.35, w: 11.2, h: 5.2, fontSize: 20, breakLine: true, color: '374151', margin: 0.05,
      })
    }
  }
  return pptx.write({ outputType: 'nodebuffer' })
}

async function readOfficeBuffer(root, filename, maxInputBytes) {
  const { candidate, info } = await resolveWorkspacePath(root, filename)
  if (info.size > maxInputBytes) fail('INPUT_TOO_LARGE', `input exceeds ${maxInputBytes} bytes`)
  const canonicalRoot = await realpath(root)
  return { buffer: await readFile(candidate), relativePath: relative(canonicalRoot, candidate), bytes: info.size }
}

function readZipEntries(buffer, maxUncompressedBytes) {
  let entries
  try {
    entries = new AdmZip(buffer).getEntries()
  } catch (error) {
    fail('INVALID_DOCUMENT', `could not inspect Office ZIP: ${error.message}`)
  }
  if (entries.length > MAX_ZIP_ENTRIES) fail('ZIP_TOO_LARGE', `Office package contains more than ${MAX_ZIP_ENTRIES} entries`)
  const uncompressedBytes = entries.reduce((total, entry) => total + Number(entry.header?.size || 0), 0)
  if (uncompressedBytes > maxUncompressedBytes) fail('INPUT_TOO_LARGE', `uncompressed Office content exceeds ${maxUncompressedBytes} bytes`)
  return entries
}

export async function readWordText(root, input, options, signal) {
  checkAbort(signal)
  const value = validateRead(input)
  if (extname(value.filename).toLowerCase() !== '.docx') fail('INVALID_ARGUMENTS', 'filename must end in .docx')
  const loaded = await readOfficeBuffer(root, value.filename, options.maxInputBytes)
  checkAbort(signal)
  const entries = readZipEntries(loaded.buffer, options.maxInputBytes)
  officeType(entries)
  let extracted
  try {
    extracted = await mammoth.extractRawText({ buffer: loaded.buffer })
  } catch (error) {
    fail('INVALID_DOCUMENT', `could not read DOCX: ${error.message}`)
  }
  const text = truncateUtf8(extracted.value || '', options.maxTextBytes)
  return {
    filename: loaded.relativePath,
    bytes: loaded.bytes,
    text: text.value,
    truncated: text.truncated,
    messageCount: Array.isArray(extracted.messages) ? extracted.messages.length : 0,
  }
}

function officeType(entries) {
  const names = new Set(entries.map(entry => entry.entryName))
  if (names.has('word/document.xml')) return 'docx'
  if (names.has('ppt/presentation.xml')) return 'pptx'
  fail('UNSUPPORTED_FORMAT', 'file is not a recognized DOCX or PPTX package')
}

export async function inspectOfficeFile(root, input, options, signal) {
  checkAbort(signal)
  const value = validateRead(input)
  const loaded = await readOfficeBuffer(root, value.filename, options.maxInputBytes)
  checkAbort(signal)
  const entries = readZipEntries(loaded.buffer, options.maxInputBytes)
  const type = officeType(entries)
  const names = entries.map(entry => entry.entryName)
  const files = entries.filter(entry => !entry.isDirectory)
  const xmlCount = files.filter(entry => entry.entryName.toLowerCase().endsWith('.xml')).length
  const relationshipCount = files.filter(entry => entry.entryName.toLowerCase().endsWith('.rels')).length
  const mediaCount = files.filter(entry => /^(word|ppt)\/media\//u.test(entry.entryName)).length
  const slideCount = files.filter(entry => /^ppt\/slides\/slide\d+\.xml$/u.test(entry.entryName)).length
  let tableCount = 0
  if (type === 'docx') {
    const documentEntry = entries.find(entry => entry.entryName === 'word/document.xml')
    if (documentEntry) tableCount = (documentEntry.getData().toString('utf8').match(/<w:tbl(?:\s|>)/gu) || []).length
  }
  const returned = entries.slice(0, MAX_RETURNED_ENTRIES).map(entry => ({
    name: entry.entryName,
    directory: Boolean(entry.isDirectory),
    compressedBytes: entry.header.compressedSize,
    uncompressedBytes: entry.header.size,
  }))
  return {
    filename: loaded.relativePath,
    bytes: loaded.bytes,
    type,
    entryCount: entries.length,
    counts: { files: files.length, directories: entries.length - files.length, xml: xmlCount, relationships: relationshipCount, media: mediaCount, slides: slideCount, tables: tableCount },
    entries: returned,
    entriesTruncated: entries.length > returned.length,
    hasContentTypes: names.includes('[Content_Types].xml'),
  }
}

function result(value) {
  return { isError: false, content: [{ type: 'text', text: JSON.stringify(value) }], structuredContent: value }
}

function errorResult(error) {
  const code = error?.code || 'OFFICE_PLUGIN_ERROR'
  const message = error?.message || String(error)
  return { isError: true, content: [{ type: 'text', text: `${code}: ${message}` }], structuredContent: { error: { code, message } } }
}

async function nativeApproval(ctx, execution, summary) {
  const approval = ctx?.get?.('approval')
  if (!approval || !execution?.agent || typeof approval.request !== 'function') return false
  const decision = await approval.request({
    agent: execution.agent,
    toolName: execution.name,
    callId: execution.callId,
    reason: `Crawshrimp Office write: ${summary.operation}`,
    arguments: summary,
    signal: execution.signal,
  })
  return decision === 'allowed-once'
}

function toolSchemas() {
  return {
    wordCreate: {
      type: 'object', additionalProperties: false,
      properties: {
        filename: { type: 'string', description: 'Relative .docx output path inside workspaceDir' },
        title: { type: 'string' },
        paragraphs: { type: 'array', items: { type: 'string' }, maxItems: MAX_PARAGRAPHS },
        tables: { type: 'array', maxItems: MAX_TABLES, items: { type: 'object', additionalProperties: false, properties: { headers: { type: 'array', items: { type: 'string' } }, rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } } } } },
      },
      required: ['filename'],
    },
    wordRead: {
      type: 'object', additionalProperties: false,
      properties: { filename: { type: 'string', description: 'Relative .docx path inside workspaceDir' } },
      required: ['filename'],
    },
    slidesCreate: {
      type: 'object', additionalProperties: false,
      properties: {
        filename: { type: 'string', description: 'Relative .pptx output path inside workspaceDir' },
        title: { type: 'string' },
        slides: { type: 'array', maxItems: MAX_SLIDES, items: { type: 'object', additionalProperties: false, properties: { heading: { type: 'string' }, bullets: { type: 'array', items: { type: 'string' } } }, required: ['heading'] } },
      },
      required: ['filename', 'title'],
    },
    inspect: {
      type: 'object', additionalProperties: false,
      properties: { filename: { type: 'string', description: 'Relative .docx or .pptx path inside workspaceDir' } },
      required: ['filename'],
    },
  }
}

function combinedSignal(execution, options) {
  const signals = []
  if (options.lifetime?.signal) signals.push(options.lifetime.signal)
  if (execution?.signal) signals.push(execution.signal)
  if (signals.length === 0) return undefined
  if (signals.length === 1) return signals[0]
  return AbortSignal.any(signals)
}

function trackInFlight(options, operation) {
  const promise = Promise.resolve().then(operation)
  if (!options.inFlight) return promise
  options.inFlight.add(promise)
  return promise.finally(() => options.inFlight.delete(promise))
}

export async function createToolDefinitions(ctx, options) {
  const schemas = toolSchemas()
  const write = (toolName, description, inputSchema, operation, execute) => createMcpToolDefinition(ctx, {
    name: toolName,
    rawName: toolName,
    description,
    inputSchema,
    call: (args, execution) => trackInFlight(options, async () => {
      const signal = combinedSignal(execution, options)
      const scopedExecution = { ...execution, signal }
      try {
        checkAbort(signal)
        if (!await nativeApproval(ctx, scopedExecution, { operation, filename: args?.filename })) return errorResult(new OfficePluginError('APPROVAL_REQUIRED', 'native DSH approval was not granted'))
        checkAbort(signal)
        return result(await execute(args, scopedExecution))
      } catch (error) {
        if (error?.name === 'AbortError' || signal?.aborted) throw error
        return errorResult(error)
      }
    }),
  })
  const read = (toolName, description, inputSchema, execute) => createMcpToolDefinition(ctx, {
    name: toolName,
    rawName: toolName,
    description,
    inputSchema,
    call: (args, execution) => trackInFlight(options, async () => {
      const signal = combinedSignal(execution, options)
      const scopedExecution = { ...execution, signal }
      try {
        checkAbort(signal)
        return result(await execute(args, scopedExecution))
      } catch (error) {
        if (error?.name === 'AbortError' || signal?.aborted) throw error
        return errorResult(error)
      }
    }),
  })
  return [
    write('crawshrimp_office_word_create', 'Create a DOCX from a title, paragraphs, and simple tables.', schemas.wordCreate, 'word_create', async (args, execution) => {
      const validated = validateWordCreate(args)
      const bytes = await createWordBuffer(validated)
      const target = await atomicWrite(options.workspaceRoot, validated.filename, bytes, execution?.signal, options.maxOutputBytes)
      return { filename: relative(options.workspaceRoot, target), type: 'docx', bytes: bytes.byteLength, visualReview: 'not_rendered' }
    }),
    read('crawshrimp_office_word_read', 'Extract capped plain text from a DOCX.', schemas.wordRead, (args, execution) => readWordText(options.workspaceRoot, args, options, execution?.signal)),
    write('crawshrimp_office_slides_create', 'Create a PPTX with a title slide and simple heading/bullet slides.', schemas.slidesCreate, 'slides_create', async (args, execution) => {
      const validated = validateSlidesCreate(args)
      const bytes = await createSlidesBuffer(validated)
      const target = await atomicWrite(options.workspaceRoot, validated.filename, bytes, execution?.signal, options.maxOutputBytes)
      return { filename: relative(options.workspaceRoot, target), type: 'pptx', bytes: bytes.byteLength, visualReview: 'not_rendered' }
    }),
    read('crawshrimp_office_inspect', 'Inspect OOXML ZIP structure and bounded entry metadata for a DOCX or PPTX.', schemas.inspect, (args, execution) => inspectOfficeFile(options.workspaceRoot, args, options, execution?.signal)),
  ]
}

export async function apply(ctx, config = {}) {
  const options = normaliseOptions(config)
  options.workspaceRoot = await ensureWorkspace(options.workspaceDir)
  options.lifetime = new AbortController()
  options.inFlight = new Set()
  const definitions = await createToolDefinitions(ctx, options)
  ctx.effect(() => {
    const disposers = []
    try {
      for (const definition of definitions) disposers.push(ctx.tools.register(definition))
    } catch (error) {
      for (const dispose of disposers.reverse()) dispose()
      throw error
    }
    return async () => {
      options.lifetime.abort(new Error('Office plugin unloaded'))
      await Promise.allSettled([...options.inFlight])
      for (const dispose of disposers.reverse()) dispose()
    }
  })
}
