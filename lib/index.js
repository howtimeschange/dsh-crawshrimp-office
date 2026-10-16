import { randomUUID } from 'node:crypto'
import { mkdir, readFile, realpath, stat, unlink, writeFile, link, lstat } from 'node:fs/promises'
import { dirname, extname, isAbsolute, join, posix as posixPath, relative, resolve, sep } from 'node:path'

import z from '@deepseek-ai/schemastery'
import { createMcpToolDefinition } from '@deepseek-ai/dsh-mcp-client'
import { DOMParser } from '@xmldom/xmldom'
import AdmZip from 'adm-zip'
import {
  AlignmentType,
  Document,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import mammoth from 'mammoth'
import pptxgen from 'pptxgenjs'
import { nativeFullAccess } from './permissions.js'

export const name = 'crawshrimp-office'
export const inject = ['tools']

const DEFAULT_MAX_INPUT_BYTES = 16 * 1024 * 1024
const DEFAULT_MAX_OUTPUT_BYTES = 16 * 1024 * 1024
const DEFAULT_MAX_TEXT_BYTES = 128 * 1024
const DEFAULT_MAX_MARKDOWN_BYTES = 512 * 1024
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
const MAX_MARKDOWN_LINES = 5_000
const MAX_MARKDOWN_BLOCKS = 500
const MAX_MARKDOWN_LINE_LENGTH = 12_000
const MAX_MARKDOWN_TABLE_COLUMNS = 30
const MAX_MARKDOWN_TABLE_ROWS = 100
const MAX_READ_SLIDES = MAX_SLIDES + 1
const DEFAULT_TABLE_WIDTH_TWIPS = 9_000
const DRAWINGML_NAMESPACE = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const STRICT_DRAWINGML_NAMESPACE = 'http://purl.oclc.org/ooxml/drawingml/main'
const DRAWINGML_NAMESPACES = new Set([DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE])
const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const STRICT_WORD_NAMESPACE = 'http://purl.oclc.org/ooxml/wordprocessingml/main'
const WORD_NAMESPACES = new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE])
const PRESENTATION_NAMESPACE = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const STRICT_PRESENTATION_NAMESPACE = 'http://purl.oclc.org/ooxml/presentationml/main'
const PRESENTATION_NAMESPACES = new Set([PRESENTATION_NAMESPACE, STRICT_PRESENTATION_NAMESPACE])
const RELATIONSHIPS_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const STRICT_RELATIONSHIPS_NAMESPACE = 'http://purl.oclc.org/ooxml/officeDocument/relationships'
const RELATIONSHIP_NAMESPACES = new Set([RELATIONSHIPS_NAMESPACE, STRICT_RELATIONSHIPS_NAMESPACE])
const PACKAGE_RELATIONSHIPS_NAMESPACE = 'http://schemas.openxmlformats.org/package/2006/relationships'
const STRICT_PACKAGE_RELATIONSHIPS_NAMESPACE = 'http://purl.oclc.org/ooxml/package/relationships'
const PACKAGE_RELATIONSHIP_NAMESPACES = new Set([PACKAGE_RELATIONSHIPS_NAMESPACE, STRICT_PACKAGE_RELATIONSHIPS_NAMESPACE])
const CONTENT_TYPES_NAMESPACE = 'http://schemas.openxmlformats.org/package/2006/content-types'
const DOCX_MAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const PPTX_MAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
const OFFICE_DOCUMENT_RELATIONSHIP_TYPES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
])
const PPTX_SLIDE_RELATIONSHIP_TYPES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  'http://purl.oclc.org/ooxml/officeDocument/relationships/slide',
])
const XML_FORBIDDEN_DECLARATION = /<!DOCTYPE\b|<!ENTITY\b/iu

export const Config = z.object({
  workspaceDir: z.string().default('./crawshrimp-office-output'),
  maxInputBytes: z.natural().default(DEFAULT_MAX_INPUT_BYTES),
  maxOutputBytes: z.natural().default(DEFAULT_MAX_OUTPUT_BYTES),
  maxTextBytes: z.natural().default(DEFAULT_MAX_TEXT_BYTES),
  maxMarkdownBytes: z.natural().default(DEFAULT_MAX_MARKDOWN_BYTES),
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

function asMultilineString(value, label, { maxBytes = DEFAULT_MAX_MARKDOWN_BYTES, required = true } = {}) {
  if (typeof value !== 'string' || (required && !value.trim()) || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(value)) {
    fail('INVALID_ARGUMENTS', `${label} must be a safe UTF-8 string`)
  }
  const bytes = Buffer.byteLength(value, 'utf8')
  if (bytes > maxBytes) fail('INPUT_TOO_LARGE', `${label} exceeds ${maxBytes} bytes`)
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
    maxMarkdownBytes: config.maxMarkdownBytes ?? DEFAULT_MAX_MARKDOWN_BYTES,
  }
  asSafeInteger(options.maxInputBytes, 'maxInputBytes', { min: 1, max: 128 * 1024 * 1024 })
  asSafeInteger(options.maxOutputBytes, 'maxOutputBytes', { min: 1, max: 128 * 1024 * 1024 })
  asSafeInteger(options.maxTextBytes, 'maxTextBytes', { min: 1, max: 8 * 1024 * 1024 })
  asSafeInteger(options.maxMarkdownBytes, 'maxMarkdownBytes', { min: 1, max: MAX_ARGUMENT_BYTES })
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
  let end = Math.max(0, Math.min(maxBytes, bytes.byteLength))
  // Never split a UTF-8 sequence. Buffer#toString would otherwise insert U+FFFD,
  // whose three bytes can also make the returned value exceed maxBytes.
  while (end > 0) {
    let start = end - 1
    while (start > 0 && (bytes[start] & 0xc0) === 0x80) start -= 1
    const lead = bytes[start]
    const sequenceLength = lead < 0x80 ? 1 : lead < 0xe0 ? 2 : lead < 0xf0 ? 3 : 4
    if (start + sequenceLength <= end) break
    end = start
  }
  return { value: bytes.subarray(0, end).toString('utf8'), truncated: true }
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

function validateMarkdownCreate(input, options = {}) {
  const value = asObject(input, 'arguments')
  assertPayloadSize(value)
  assertKnownKeys(value, ['filename', 'markdown', 'title'], 'arguments')
  const filename = asString(value.filename, 'filename', { max: MAX_FILENAME_LENGTH })
  if (extname(filename).toLowerCase() !== '.docx') fail('INVALID_ARGUMENTS', 'filename must end in .docx')
  const markdown = asMultilineString(value.markdown, 'markdown', {
    maxBytes: options.maxMarkdownBytes ?? DEFAULT_MAX_MARKDOWN_BYTES,
  })
  const title = value.title === undefined ? '' : cleanText(value.title, 'title', MAX_TITLE_LENGTH)
  if (!markdown.trim() && !title) fail('INVALID_ARGUMENTS', 'markdown or title is required')
  return { filename, markdown, title }
}

function validateOfficeValidation(input) {
  const value = asObject(input, 'arguments')
  assertPayloadSize(value)
  assertKnownKeys(value, ['filename', 'expectedText', 'expectedSlideCount', 'expectedTableCount'], 'arguments')
  const filename = asString(value.filename, 'filename', { max: MAX_FILENAME_LENGTH })
  const extension = extname(filename).toLowerCase()
  if (extension !== '.docx' && extension !== '.pptx') fail('INVALID_ARGUMENTS', 'filename must end in .docx or .pptx')
  const expectedText = value.expectedText === undefined
    ? undefined
    : asMultilineString(value.expectedText, 'expectedText', { maxBytes: 8 * 1024 * 1024, required: false })
  const expectedSlideCount = value.expectedSlideCount === undefined
    ? undefined
    : asSafeInteger(value.expectedSlideCount, 'expectedSlideCount', { min: 0, max: MAX_READ_SLIDES })
  const expectedTableCount = value.expectedTableCount === undefined
    ? undefined
    : asSafeInteger(value.expectedTableCount, 'expectedTableCount', { min: 0, max: MAX_TABLES * MAX_TABLE_ROWS })
  return { filename, expectedText, expectedSlideCount, expectedTableCount }
}

function createParagraph(text, heading = false) {
  return new Paragraph({
    children: [new TextRun({ text })],
    ...(heading ? { heading: HeadingLevel.TITLE } : {}),
  })
}

function tableColumnWidths(columnCount) {
  const count = Math.max(1, columnCount)
  const base = Math.floor(DEFAULT_TABLE_WIDTH_TWIPS / count)
  const remainder = DEFAULT_TABLE_WIDTH_TWIPS - base * count
  return Array.from({ length: count }, (_unused, index) => base + (index < remainder ? 1 : 0))
}

function createOfficeTable(rows) {
  const columnCount = Math.max(1, ...rows.map(row => row.length))
  const columnWidths = tableColumnWidths(columnCount)
  return new Table({
    width: { size: DEFAULT_TABLE_WIDTH_TWIPS, type: WidthType.DXA },
    columnWidths,
    layout: TableLayoutType.FIXED,
    rows: rows.map(row => new TableRow({
      children: row.map((cell, columnIndex) => new TableCell({
        width: { size: columnWidths[columnIndex] ?? columnWidths[columnWidths.length - 1], type: WidthType.DXA },
        children: [createParagraph(cell)],
      })),
    })),
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
    children.push(createOfficeTable(rows))
  }
  return Packer.toBuffer(new Document({ sections: [{ children }] }))
}

function markdownWarning(warnings, code, message, line) {
  warnings.push({ code, message, ...(line === undefined ? {} : { line }) })
}

function splitMarkdownTableRow(line) {
  let value = line.trim()
  if (value.startsWith('|')) value = value.slice(1)
  if (value.endsWith('|') && !value.endsWith('\\|')) value = value.slice(0, -1)
  const cells = []
  let current = ''
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]
    if (character === '\\' && value[index + 1] === '|') {
      current += '|'
      index += 1
    } else if (character === '|') {
      cells.push(current.trim())
      current = ''
    } else {
      current += character
    }
  }
  cells.push(current.trim())
  return cells
}

function isMarkdownTableDelimiter(line) {
  const cells = splitMarkdownTableRow(line)
  return cells.length >= 2 && cells.every(cell => /^:?-{3,}:?$/u.test(cell))
}

function sanitiseMarkdownInline(value, warnings, line) {
  let text = value
  text = text.replace(/!\[([^\]]*)\]\(([^)]*)\)/gu, (_match, alt) => {
    markdownWarning(warnings, 'UNSUPPORTED_MEDIA', 'Markdown image syntax was kept as alt text; no image was fetched or embedded.', line)
    return alt
  })
  text = text.replace(/<img\b([^>]*)>/giu, (_match, attributes) => {
    const altMatch = attributes.match(/\balt\s*=\s*(["'])(.*?)\1/iu)
    markdownWarning(warnings, 'UNSUPPORTED_MEDIA', 'HTML image markup was kept as alt text; no image was fetched or embedded.', line)
    return altMatch?.[2] || ''
  })
  text = text.replace(/<!--[\s\S]*?-->/gu, () => {
    markdownWarning(warnings, 'HTML_STRIPPED', 'HTML comments were stripped and never executed.', line)
    return ''
  })
  text = text.replace(/<((?:https?|mailto):[^>\s]+)>/giu, '$1')
  text = text.replace(/<\/?[A-Za-z][^>]*>/gu, () => {
    markdownWarning(warnings, 'HTML_STRIPPED', 'HTML tags were stripped and never executed.', line)
    return ''
  })
  text = text.replace(/\[([^\]]+)\]\((?:[^()]|\([^)]*\))*\)/gu, '$1')
  text = text.replace(/\[([^\]]+)\]\[[^\]]*\]/gu, '$1')
  text = text.replace(/`([^`]+)`/gu, '$1')
  text = text.replace(/(\*\*|__)(.*?)\1/gu, '$2')
  text = text.replace(/(\*|_)(.*?)\1/gu, '$2')
  text = text.replace(/\\([\\`*_{}\[\]()#+.!|>~-])/gu, '$1')
  return text.replace(/[ \t]+/gu, ' ').trim()
}

function isMarkdownListLine(line) {
  return /^(\s*)([-+*]|\d+[.)])\s+(.+)$/u.test(line)
}

function isMarkdownBlockStart(lines, index) {
  const line = lines[index] || ''
  if (/^\s*$|^ {0,3}#{1,6}\s+/u.test(line) || isMarkdownListLine(line) || /^\s*>/u.test(line) || /^\s*(`{3,}|~{3,})/u.test(line)) return true
  return index + 1 < lines.length && isMarkdownTableDelimiter(lines[index + 1]) && line.includes('|')
}

function parseMarkdown(markdown) {
  const lines = markdown.replace(/\r\n?/gu, '\n').split('\n')
  if (lines.length > MAX_MARKDOWN_LINES) fail('INPUT_TOO_LARGE', `markdown contains more than ${MAX_MARKDOWN_LINES} lines`)
  for (const [index, line] of lines.entries()) {
    if (line.length > MAX_MARKDOWN_LINE_LENGTH) fail('INPUT_TOO_LARGE', `markdown line ${index + 1} exceeds ${MAX_MARKDOWN_LINE_LENGTH} characters`)
  }
  const warnings = []
  const blocks = []
  const push = block => {
    if (blocks.length >= MAX_MARKDOWN_BLOCKS) fail('INPUT_TOO_LARGE', `markdown contains more than ${MAX_MARKDOWN_BLOCKS} blocks`)
    blocks.push(block)
  }
  let index = 0
  while (index < lines.length) {
    const line = lines[index]
    if (!line.trim()) {
      index += 1
      continue
    }
    const fence = line.match(/^ {0,3}(`{3,}|~{3,})\s*(.*)$/u)
    if (fence) {
      const marker = fence[1][0]
      const markerLength = fence[1].length
      const codeLines = []
      const startLine = index + 1
      index += 1
      let closed = false
      while (index < lines.length) {
        if (new RegExp(`^ {0,3}${marker}{${markerLength},}\\s*$`, 'u').test(lines[index])) {
          closed = true
          index += 1
          break
        }
        codeLines.push(lines[index])
        index += 1
      }
      if (!closed) markdownWarning(warnings, 'UNTERMINATED_CODE_FENCE', 'Unterminated code fence was kept as plain text.', startLine)
      push({ type: 'paragraph', text: codeLines.join('\n') })
      continue
    }
    const heading = line.match(/^ {0,3}(#{1,6})\s+(.+?)\s*$/u)
    if (heading) {
      const headingText = heading[2].replace(/\s+#+\s*$/u, '')
      push({ type: 'heading', level: heading[1].length, text: sanitiseMarkdownInline(headingText, warnings, index + 1) })
      index += 1
      continue
    }
    if (index + 1 < lines.length && line.includes('|') && isMarkdownTableDelimiter(lines[index + 1])) {
      const headers = splitMarkdownTableRow(line)
      if (headers.length > MAX_MARKDOWN_TABLE_COLUMNS) fail('INPUT_TOO_LARGE', `markdown table exceeds ${MAX_MARKDOWN_TABLE_COLUMNS} columns`)
      const rows = []
      index += 2
      while (index < lines.length && lines[index].trim() && lines[index].includes('|') && !isMarkdownBlockStart(lines, index)) {
        const row = splitMarkdownTableRow(lines[index])
        if (row.length > MAX_MARKDOWN_TABLE_COLUMNS) fail('INPUT_TOO_LARGE', `markdown table exceeds ${MAX_MARKDOWN_TABLE_COLUMNS} columns`)
        if (row.length !== headers.length) {
          markdownWarning(warnings, 'TABLE_SHAPE_NORMALISED', `Table row ${index + 1} was padded or truncated to ${headers.length} columns.`, index + 1)
        }
        const normalised = Array.from({ length: headers.length }, (_unused, cellIndex) => row[cellIndex] || '')
        rows.push(normalised.map(cell => sanitiseMarkdownInline(cell, warnings, index + 1)))
        if (rows.length >= MAX_MARKDOWN_TABLE_ROWS) fail('INPUT_TOO_LARGE', `markdown table exceeds ${MAX_MARKDOWN_TABLE_ROWS} rows`)
        index += 1
      }
      push({
        type: 'table',
        headers: headers.map(cell => sanitiseMarkdownInline(cell, warnings, index - rows.length - 1)),
        rows,
      })
      continue
    }
    const list = line.match(/^(\s*)([-+*]|\d+[.)])\s+(.+)$/u)
    if (list) {
      while (index < lines.length) {
        const item = lines[index].match(/^(\s*)([-+*]|\d+[.)])\s+(.+)$/u)
        if (!item) break
        const indent = Math.min(8, Math.floor(item[1].replace(/\t/gu, '  ').length / 2))
        push({
          type: 'list',
          ordered: /^\d/u.test(item[2]),
          level: indent,
          text: sanitiseMarkdownInline(item[3], warnings, index + 1),
        })
        index += 1
      }
      continue
    }
    if (/^\s*>/u.test(line)) {
      const quoteLines = []
      const startLine = index + 1
      while (index < lines.length && /^\s*>/u.test(lines[index])) {
        quoteLines.push(sanitiseMarkdownInline(lines[index].replace(/^\s*>\s?/u, ''), warnings, index + 1))
        index += 1
      }
      push({ type: 'paragraph', text: quoteLines.join('\n'), quote: true, line: startLine })
      continue
    }
    if (/^\s{0,3}(([-*_])\s*){3,}$/u.test(line)) {
      index += 1
      continue
    }
    const paragraphLines = []
    const startLine = index + 1
    while (index < lines.length && lines[index].trim() && !isMarkdownBlockStart(lines, index)) {
      paragraphLines.push(sanitiseMarkdownInline(lines[index], warnings, index + 1))
      index += 1
    }
    if (paragraphLines.length) push({ type: 'paragraph', text: paragraphLines.join(' '), line: startLine })
    else index += 1
  }
  return { blocks, warnings }
}

function markdownNumberingLevels(format) {
  return Array.from({ length: 9 }, (_unused, level) => ({
    level,
    format,
    text: format === LevelFormat.BULLET ? (level % 2 ? '○' : '●') : `%${level + 1}.`,
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
  }))
}

function markdownHeading(level) {
  return [
    HeadingLevel.HEADING_1,
    HeadingLevel.HEADING_2,
    HeadingLevel.HEADING_3,
    HeadingLevel.HEADING_4,
    HeadingLevel.HEADING_5,
    HeadingLevel.HEADING_6,
  ][Math.max(0, Math.min(5, level - 1))]
}

function markdownBlockToDocx(block) {
  if (block.type === 'heading') return new Paragraph({ heading: markdownHeading(block.level), children: [new TextRun({ text: block.text })] })
  if (block.type === 'list') {
    return new Paragraph({
      numbering: { reference: block.ordered ? 'markdown-ordered' : 'markdown-bullet', level: block.level },
      children: [new TextRun({ text: block.text })],
    })
  }
  if (block.type === 'table') {
    const rows = [block.headers, ...block.rows]
    return createOfficeTable(rows)
  }
  return new Paragraph({ children: [new TextRun({ text: block.text })] })
}

export async function createMarkdownWordBuffer(input, options = {}) {
  const value = validateMarkdownCreate(input, options)
  const parsed = parseMarkdown(value.markdown)
  const children = []
  if (value.title) children.push(createParagraph(value.title, true))
  children.push(...parsed.blocks.map(markdownBlockToDocx))
  const document = new Document({
    numbering: {
      config: [
        { reference: 'markdown-bullet', levels: markdownNumberingLevels(LevelFormat.BULLET) },
        { reference: 'markdown-ordered', levels: markdownNumberingLevels(LevelFormat.DECIMAL) },
      ],
    },
    sections: [{ children }],
  })
  return { buffer: await Packer.toBuffer(document), warnings: parsed.warnings }
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

function requiredZipEntry(entries, entryName) {
  const entry = entries.find(candidate => candidate.entryName === entryName)
  if (!entry || entry.isDirectory) fail('INVALID_DOCUMENT', `Office package is missing required part: ${entryName}`)
  return entry
}

function xmlLocalName(node) {
  return node?.localName || node?.nodeName?.split(':').pop() || ''
}

function namespaceMatches(namespace, allowed) {
  if (!allowed) return true
  return allowed instanceof Set ? allowed.has(namespace) : namespace === allowed
}

function walkXmlElements(node, callback) {
  if (!node) return
  if (node.nodeType === 1) callback(node)
  for (let child = node.firstChild; child; child = child.nextSibling) walkXmlElements(child, callback)
}

function findXmlElements(document, localName, namespace) {
  const matches = []
  walkXmlElements(document, element => {
    if (xmlLocalName(element) === localName && namespaceMatches(element.namespaceURI, namespace)) matches.push(element)
  })
  return matches
}

function xmlAttribute(element, localName, namespace) {
  if (namespace) {
    const namespaces = namespace instanceof Set ? namespace : new Set([namespace])
    if (typeof element.getAttributeNS === 'function') {
      for (const candidate of namespaces) {
        const value = element.getAttributeNS(candidate, localName)
        if (value) return value
      }
    }
    return ''
  }
  for (let index = 0; element.attributes && index < element.attributes.length; index += 1) {
    const attribute = element.attributes.item(index)
    if ((attribute.localName || attribute.name?.split(':').pop()) === localName) return attribute.value
  }
  return ''
}

function assertXmlNamespaceBindings(document, label) {
  walkXmlElements(document, element => {
    if (element.prefix && !element.namespaceURI) fail('INVALID_DOCUMENT', `${label} contains an unbound XML prefix: ${element.prefix}`)
    for (let index = 0; element.attributes && index < element.attributes.length; index += 1) {
      const attribute = element.attributes.item(index)
      if (attribute.prefix && attribute.prefix !== 'xmlns' && !attribute.namespaceURI) {
        fail('INVALID_DOCUMENT', `${label} contains an unbound XML attribute prefix: ${attribute.prefix}`)
      }
    }
  })
}

function assertXmlRoot(document, localName, namespaces, label) {
  const root = document?.documentElement
  if (!root || xmlLocalName(root) !== localName || !namespaceMatches(root.namespaceURI, namespaces)) {
    fail('INVALID_DOCUMENT', `${label} has an unexpected root element or namespace`)
  }
}

function assertXmlCharacters(source, label) {
  for (let index = 0; index < source.length; index += 1) {
    const code = source.charCodeAt(index)
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = source.charCodeAt(index + 1)
      if (next >= 0xdc00 && next <= 0xdfff) {
        index += 1
        continue
      }
      fail('INVALID_DOCUMENT', `${label} contains an invalid XML 1.0 character`)
    }
    if (code >= 0xdc00 && code <= 0xdfff
      || (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d)
      || code === 0xfffe || code === 0xffff) {
      fail('INVALID_DOCUMENT', `${label} contains an invalid XML 1.0 character`)
    }
  }
}

function parseOoxmlXml(bytes, label) {
  let source
  try {
    source = Buffer.isBuffer(bytes) ? new TextDecoder('utf-8', { fatal: true }).decode(bytes) : String(bytes)
  } catch {
    fail('INVALID_DOCUMENT', `${label} is not valid UTF-8`)
  }
  assertXmlCharacters(source, label)
  if (XML_FORBIDDEN_DECLARATION.test(source)) fail('UNSAFE_XML', `${label} contains a DTD or entity declaration`)
  const errors = []
  let document
  try {
    document = new DOMParser({
      errorHandler: {
        warning: message => errors.push(`warning: ${message}`),
        error: message => errors.push(String(message)),
        fatalError: message => errors.push(String(message)),
      },
    }).parseFromString(source, 'application/xml')
  } catch (error) {
    fail('INVALID_DOCUMENT', `could not parse ${label}: ${error.message}`)
  }
  if (errors.length || !document?.documentElement || xmlLocalName(document.documentElement) === 'parsererror') {
    fail('INVALID_DOCUMENT', `could not parse ${label}${errors[0] ? `: ${errors[0]}` : ''}`)
  }
  assertXmlNamespaceBindings(document, label)
  return document
}

function parseZipXmlEntry(entry, label) {
  try {
    return parseOoxmlXml(entry.getData(), label)
  } catch (error) {
    if (error?.code) throw error
    fail('INVALID_DOCUMENT', `could not read ${label}: ${error.message}`)
  }
}

function collectXmlText(node) {
  let text = ''
  if (!node) return text
  if (node.nodeType === 3 || node.nodeType === 4) return node.data || ''
  if (node.nodeType !== 1 && node.nodeType !== 9) return text
  if (xmlLocalName(node) === 'br') text += '\n'
  if (xmlLocalName(node) === 'tab') text += '\t'
  for (let child = node.firstChild; child; child = child.nextSibling) text += collectXmlText(child)
  return text
}

async function extractWordText(buffer) {
  let extracted
  try {
    extracted = await mammoth.extractRawText({ buffer })
  } catch (error) {
    fail('INVALID_DOCUMENT', `could not read DOCX: ${error.message}`)
  }
  return {
    text: extracted.value || '',
    messageCount: Array.isArray(extracted.messages) ? extracted.messages.length : 0,
  }
}

function resolvePackageTarget(baseDirectory, target) {
  let decoded
  try {
    decoded = decodeURIComponent(String(target).split('#', 1)[0])
  } catch {
    fail('INVALID_DOCUMENT', `invalid OOXML relationship target: ${target}`)
  }
  if (/^[A-Za-z][A-Za-z0-9+.-]*:/u.test(decoded)) fail('INVALID_DOCUMENT', `external OOXML relationship target is not supported: ${target}`)
  const normalized = decoded.startsWith('/')
    ? posixPath.normalize(decoded.slice(1))
    : posixPath.normalize(posixPath.join(baseDirectory, decoded))
  if (!normalized || normalized === '..' || normalized.startsWith('../') || posixPath.isAbsolute(normalized)) {
    fail('PATH_OUTSIDE_WORKSPACE', `OOXML relationship target escapes its package: ${target}`)
  }
  return normalized
}

function readPptxSlideTargets(entries) {
  const presentationEntry = requiredZipEntry(entries, 'ppt/presentation.xml')
  const relationshipsEntry = requiredZipEntry(entries, 'ppt/_rels/presentation.xml.rels')
  const presentation = parseZipXmlEntry(presentationEntry, 'ppt/presentation.xml')
  const relationships = parseZipXmlEntry(relationshipsEntry, 'ppt/_rels/presentation.xml.rels')
  assertXmlRoot(presentation, 'presentation', PRESENTATION_NAMESPACES, 'ppt/presentation.xml')
  assertXmlRoot(relationships, 'Relationships', PACKAGE_RELATIONSHIP_NAMESPACES, 'ppt/presentation.xml.rels')
  const relationshipMap = new Map()
  for (const relationship of findXmlElements(relationships, 'Relationship', PACKAGE_RELATIONSHIP_NAMESPACES)) {
    const id = xmlAttribute(relationship, 'Id')
    const target = xmlAttribute(relationship, 'Target')
    const type = xmlAttribute(relationship, 'Type')
    if (id && target && PPTX_SLIDE_RELATIONSHIP_TYPES.has(type)) relationshipMap.set(id, { target, type })
  }
  const slideIds = findXmlElements(presentation, 'sldId', PRESENTATION_NAMESPACES)
  if (slideIds.length > MAX_READ_SLIDES) fail('INPUT_TOO_LARGE', `presentation contains more than ${MAX_READ_SLIDES} slides`)
  return slideIds.map((slideId, index) => {
    const relationshipId = xmlAttribute(slideId, 'id', RELATIONSHIP_NAMESPACES)
    const relationship = relationshipMap.get(relationshipId)
    if (!relationship) fail('INVALID_DOCUMENT', `presentation slide ${index + 1} has no relationship target`)
    const target = resolvePackageTarget('ppt', relationship.target)
    if (!target.startsWith('ppt/slides/')) fail('INVALID_DOCUMENT', `presentation slide target is outside ppt/slides: ${relationship.target}`)
    const entry = requiredZipEntry(entries, target)
    return { index: index + 1, target, entry }
  })
}

function extractPptxSlideText(entry, index) {
  const document = parseZipXmlEntry(entry, entry.entryName)
  assertXmlRoot(document, 'sld', PRESENTATION_NAMESPACES, entry.entryName)
  const paragraphs = findXmlElements(document, 'p')
    .filter(paragraph => DRAWINGML_NAMESPACES.has(paragraph.namespaceURI))
    .map(paragraph => collectXmlText(paragraph).replace(/[ \t]+$/gu, '').trim())
    .filter(Boolean)
  const text = paragraphs.join('\n')
  return { index, filename: entry.entryName, text }
}

export async function readWordText(root, input, options, signal) {
  checkAbort(signal)
  const value = validateRead(input)
  if (extname(value.filename).toLowerCase() !== '.docx') fail('INVALID_ARGUMENTS', 'filename must end in .docx')
  const loaded = await readOfficeBuffer(root, value.filename, options.maxInputBytes)
  checkAbort(signal)
  const entries = readZipEntries(loaded.buffer, options.maxInputBytes)
  ensureOfficePackage(entries, value.filename)
  validatePackageRoot(entries, 'docx')
  const documentEntry = requiredZipEntry(entries, 'word/document.xml')
  const document = parseZipXmlEntry(documentEntry, 'word/document.xml')
  assertXmlRoot(document, 'document', WORD_NAMESPACES, 'word/document.xml')
  const extracted = await extractWordText(loaded.buffer)
  const text = truncateUtf8(extracted.text, options.maxTextBytes)
  return {
    filename: loaded.relativePath,
    bytes: loaded.bytes,
    text: text.value,
    truncated: text.truncated,
    messageCount: extracted.messageCount,
  }
}

export async function readPptxSlides(root, input, options, signal) {
  checkAbort(signal)
  const value = validateRead(input)
  if (extname(value.filename).toLowerCase() !== '.pptx') fail('INVALID_ARGUMENTS', 'filename must end in .pptx')
  const loaded = await readOfficeBuffer(root, value.filename, options.maxInputBytes)
  checkAbort(signal)
  const entries = readZipEntries(loaded.buffer, options.maxInputBytes)
  ensureOfficePackage(entries, value.filename)
  validatePackageRoot(entries, 'pptx')
  const slides = readPptxSlideTargets(entries).map(slide => {
    const extracted = extractPptxSlideText(slide.entry, slide.index)
    const text = truncateUtf8(extracted.text, options.maxTextBytes)
    return { ...extracted, text: text.value, truncated: text.truncated }
  })
  return { filename: loaded.relativePath, bytes: loaded.bytes, type: 'pptx', slideCount: slides.length, slides }
}

function officeType(entries) {
  const names = new Set(entries.map(entry => entry.entryName))
  if (names.has('word/document.xml')) return 'docx'
  if (names.has('ppt/presentation.xml')) return 'pptx'
  fail('UNSUPPORTED_FORMAT', 'file is not a recognized DOCX or PPTX package')
}

function ensureOfficePackage(entries, filename) {
  const extension = extname(filename).toLowerCase()
  const type = extension === '.docx' ? 'docx' : extension === '.pptx' ? 'pptx' : undefined
  if (!type) fail('INVALID_ARGUMENTS', 'filename must end in .docx or .pptx')
  requiredZipEntry(entries, type === 'docx' ? 'word/document.xml' : 'ppt/presentation.xml')
  if (officeType(entries) !== type) fail('INVALID_DOCUMENT', `file is not a ${type.toUpperCase()} package`)
  return type
}

export async function inspectOfficeFile(root, input, options, signal) {
  checkAbort(signal)
  const value = validateRead(input)
  const loaded = await readOfficeBuffer(root, value.filename, options.maxInputBytes)
  checkAbort(signal)
  const entries = readZipEntries(loaded.buffer, options.maxInputBytes)
  const type = ensureOfficePackage(entries, value.filename)
  validatePackageRoot(entries, type)
  const names = entries.map(entry => entry.entryName)
  const files = entries.filter(entry => !entry.isDirectory)
  const xmlCount = files.filter(entry => entry.entryName.toLowerCase().endsWith('.xml')).length
  const relationshipCount = files.filter(entry => entry.entryName.toLowerCase().endsWith('.rels')).length
  const mediaCount = files.filter(entry => /^(word|ppt)\/media\//u.test(entry.entryName)).length
  const slideCount = type === 'pptx' ? readPptxSlideTargets(entries).length : 0
  let tableCount = 0
  if (type === 'docx') {
    const documentEntry = requiredZipEntry(entries, 'word/document.xml')
    const document = parseZipXmlEntry(documentEntry, 'word/document.xml')
    assertXmlRoot(document, 'document', WORD_NAMESPACES, 'word/document.xml')
    tableCount = findXmlElements(document, 'tbl', WORD_NAMESPACES).length
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

function validationIssue(code, message, expected, actual) {
  return { code, message, ...(expected === undefined ? {} : { expected }), ...(actual === undefined ? {} : { actual }) }
}

function validatePackageRoot(entries, type) {
  const contentTypes = parseZipXmlEntry(requiredZipEntry(entries, '[Content_Types].xml'), '[Content_Types].xml')
  assertXmlRoot(contentTypes, 'Types', CONTENT_TYPES_NAMESPACE, '[Content_Types].xml')
  const mainPart = type === 'docx' ? 'word/document.xml' : 'ppt/presentation.xml'
  const expectedContentType = type === 'docx' ? DOCX_MAIN_CONTENT_TYPE : PPTX_MAIN_CONTENT_TYPE
  const hasMainOverride = findXmlElements(contentTypes, 'Override', CONTENT_TYPES_NAMESPACE).some(override => {
    return xmlAttribute(override, 'PartName') === `/${mainPart}`
      && xmlAttribute(override, 'ContentType') === expectedContentType
  })
  if (!hasMainOverride) fail('INVALID_DOCUMENT', `[Content_Types].xml does not declare ${mainPart} with the expected content type`)
  const packageRelationships = parseZipXmlEntry(requiredZipEntry(entries, '_rels/.rels'), '_rels/.rels')
  assertXmlRoot(packageRelationships, 'Relationships', PACKAGE_RELATIONSHIP_NAMESPACES, '_rels/.rels')
  const hasMainRelationship = findXmlElements(packageRelationships, 'Relationship', PACKAGE_RELATIONSHIP_NAMESPACES).some(relationship => {
    const relationshipType = xmlAttribute(relationship, 'Type')
    if (!OFFICE_DOCUMENT_RELATIONSHIP_TYPES.has(relationshipType)) return false
    const target = xmlAttribute(relationship, 'Target')
    return resolvePackageTarget('', target) === mainPart
  })
  if (!hasMainRelationship) fail('INVALID_DOCUMENT', `_rels/.rels does not point to ${mainPart}`)
  const mainDocument = parseZipXmlEntry(requiredZipEntry(entries, mainPart), mainPart)
  assertXmlRoot(
    mainDocument,
    type === 'docx' ? 'document' : 'presentation',
    type === 'docx' ? WORD_NAMESPACES : PRESENTATION_NAMESPACES,
    mainPart,
  )
}

export async function validateOfficeFile(root, input, options, signal) {
  checkAbort(signal)
  const value = validateOfficeValidation(input)
  const loaded = await readOfficeBuffer(root, value.filename, options.maxInputBytes)
  checkAbort(signal)
  const entries = readZipEntries(loaded.buffer, options.maxInputBytes)
  const type = ensureOfficePackage(entries, value.filename)
  validatePackageRoot(entries, type)
  const issues = []
  let text = ''
  let slideCount = 0
  let tableCount = 0
  let slides
  if (type === 'docx') {
    const document = parseZipXmlEntry(requiredZipEntry(entries, 'word/document.xml'), 'word/document.xml')
    assertXmlRoot(document, 'document', WORD_NAMESPACES, 'word/document.xml')
    const extracted = await extractWordText(loaded.buffer)
    text = extracted.text
    tableCount = findXmlElements(document, 'tbl', WORD_NAMESPACES).length
  } else {
    slides = readPptxSlideTargets(entries).map(slide => extractPptxSlideText(slide.entry, slide.index))
    slideCount = slides.length
    text = slides.map(slide => slide.text).filter(Boolean).join('\n')
  }
  const boundedText = truncateUtf8(text, options.maxTextBytes)
  if (value.expectedText !== undefined && !text.includes(value.expectedText)) {
    issues.push(validationIssue('EXPECTED_TEXT_NOT_FOUND', 'expectedText was not found in the document text', value.expectedText, boundedText.value))
  }
  if (value.expectedSlideCount !== undefined) {
    if (type !== 'pptx') issues.push(validationIssue('SLIDE_COUNT_NOT_APPLICABLE', 'expectedSlideCount applies only to PPTX files', value.expectedSlideCount, undefined))
    else if (slideCount !== value.expectedSlideCount) issues.push(validationIssue('SLIDE_COUNT_MISMATCH', 'PPTX slide count does not match expectedSlideCount', value.expectedSlideCount, slideCount))
  }
  if (value.expectedTableCount !== undefined) {
    if (type !== 'docx') issues.push(validationIssue('TABLE_COUNT_NOT_APPLICABLE', 'expectedTableCount applies only to DOCX files', value.expectedTableCount, undefined))
    else if (tableCount !== value.expectedTableCount) issues.push(validationIssue('TABLE_COUNT_MISMATCH', 'DOCX table count does not match expectedTableCount', value.expectedTableCount, tableCount))
  }
  const response = {
    filename: loaded.relativePath,
    bytes: loaded.bytes,
    type,
    valid: issues.length === 0,
    issues,
    text: boundedText.value,
    textTruncated: boundedText.truncated,
    structure: {
      hasContentTypes: true,
      hasPackageRelationships: true,
      requiredMainPart: type === 'docx' ? 'word/document.xml' : 'ppt/presentation.xml',
      entryCount: entries.length,
    },
  }
  response.slideCount = slideCount
  response.tableCount = tableCount
  if (slides) response.slides = slides.map(slide => ({ index: slide.index, filename: slide.filename, text: truncateUtf8(slide.text, options.maxTextBytes).value }))
  return response
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
  const approval = typeof ctx?.get === 'function' ? ctx.get('approval', false) : ctx?.approval
  if (!approval || !execution?.agent) return false
  if (nativeFullAccess(ctx, execution.agent)) return true
  if (typeof approval.request !== 'function') return false
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
    wordFromMarkdown: {
      type: 'object', additionalProperties: false,
      properties: {
        filename: { type: 'string', description: 'Relative .docx output path inside workspaceDir' },
        markdown: { type: 'string', description: 'Bounded Markdown text; headings, paragraphs, lists, simple tables, and safe plain-text fallbacks are supported.' },
        title: { type: 'string', description: 'Optional document title' },
      },
      required: ['filename', 'markdown'],
    },
    slidesRead: {
      type: 'object', additionalProperties: false,
      properties: { filename: { type: 'string', description: 'Relative .pptx path inside workspaceDir' } },
      required: ['filename'],
    },
    validate: {
      type: 'object', additionalProperties: false,
      properties: {
        filename: { type: 'string', description: 'Relative .docx or .pptx path inside workspaceDir' },
        expectedText: { type: 'string', description: 'Optional text that must occur in the extracted document text' },
        expectedSlideCount: { type: 'integer', minimum: 0, description: 'Optional PPTX slide count check' },
        expectedTableCount: { type: 'integer', minimum: 0, description: 'Optional DOCX table count check' },
      },
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
    write('crawshrimp_office_word_from_markdown', 'Create a DOCX from bounded Markdown without fetching or embedding external media.', schemas.wordFromMarkdown, 'word_from_markdown', async (args, execution) => {
      const validated = validateMarkdownCreate(args, options)
      const generated = await createMarkdownWordBuffer(validated, options)
      const target = await atomicWrite(options.workspaceRoot, validated.filename, generated.buffer, execution?.signal, options.maxOutputBytes)
      return { filename: relative(options.workspaceRoot, target), type: 'docx', bytes: generated.buffer.byteLength, warnings: generated.warnings, visualReview: 'not_rendered' }
    }),
    read('crawshrimp_office_slides_read', 'Read bounded plain text from each PPTX slide without decoding media.', schemas.slidesRead, (args, execution) => readPptxSlides(options.workspaceRoot, args, options, execution?.signal)),
    read('crawshrimp_office_validate', 'Validate DOCX or PPTX structure and optional expected text/counts.', schemas.validate, (args, execution) => validateOfficeFile(options.workspaceRoot, args, options, execution?.signal)),
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
