require('dotenv').config()
const express = require('express')
const cors = require('cors')
const { MongoClient, ObjectId } = require('mongodb')
const { BlobServiceClient } = require('@azure/storage-blob')
const pdfParse = require('pdf-parse')
const mammoth = require('mammoth')
const XLSX = require('xlsx')
const { parse: htmlParse } = require('node-html-parser')
const yaml = require('js-yaml')
const Papa = require('papaparse')
const { simpleParser } = require('mailparser')
const { parseOffice } = require('officeparser')
const crypto = require('crypto')
const app = express()
const allowedOrigins = [
  'http://localhost:8080',
  'http://localhost:3000',
  'https://app.powerbi.com',
  'https://msit.powerbi.com',
  'https://anuritchat.vercel.app',
  'https://ragadminpanel.vercel.app',
  'https://df.powerbi.com',
  'https://api.powerbi.com',
]
function originAllowed(origin) {
  if (!origin) return true
  if (origin === 'null') return true
  if (allowedOrigins.includes(origin)) return true
  if (/\.(powerbi|microsoft|office)\.com$/.test(origin)) return true
  return false
}
app.use(cors({
  origin: (origin, callback) => callback(null, originAllowed(origin)),
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'x-session-id'],
  credentials: true,
}))
app.options('*', cors({
  origin: (origin, callback) => callback(null, true),
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'x-session-id'],
  credentials: true,
}))
app.use(express.json())
const MONGODB_URI = process.env.MONGODB_URI
const MONGODB_DB = process.env.MONGODB_DB || 'clientcreds'
const CHAT_HISTORY_URI = process.env.CHAT_HISTORY_URI
const CHAT_HISTORY_DB = process.env.CHAT_HISTORY_DB || 'chathistory'
const AZURE_CONNECTION_STRING = process.env.AZURE_CONNECTION_STRING || ''
const AZURE_CONTAINER_NAME = process.env.AZURE_CONTAINER_NAME || 'vectordbforrag'
const ADMIN_API_KEY = process.env.ADMIN_API_KEY
const KEY_CHECK_INTERVAL_MS = parseInt(process.env.KEY_CHECK_INTERVAL_MS || '300000', 10)
const PHI4_ENDPOINT = process.env.PHI4_ENDPOINT
const PHI4_API_KEY = process.env.PHI4_API_KEY
const PHI4_MODEL = process.env.PHI4_MODEL || 'Phi-4'
const PHI4_TIMEOUT_MS = parseInt(process.env.PHI4_TIMEOUT_MS || '30000', 10)
const AZURE_EMBED_ENDPOINT = process.env.AZURE_EMBED_ENDPOINT || ''
const AZURE_EMBED_KEY = process.env.AZURE_EMBED_KEY || ''
const AZURE_EMBED_MODEL = process.env.AZURE_EMBED_MODEL || 'text-embedding-ada-002'
const EMBED_TIMEOUT_MS = parseInt(process.env.EMBED_TIMEOUT_MS || '10000', 10)
const EMBED_POOL_LIMIT = parseInt(process.env.EMBED_POOL_LIMIT || '20', 10)
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || '60000', 10)
const WARMUP_CLIENT_IDS = (process.env.WARMUP_CLIENT_IDS || '').split(',').map(s => s.trim()).filter(Boolean)
const RAW_PREFIX = 'raw'
const CHUNK_SIZE = 1200
const CHUNK_OVERLAP = 2
const BLOB_CONCURRENCY = parseInt(process.env.BLOB_CONCURRENCY || '8', 10)
const CHUNK_CACHE_TTL = parseInt(process.env.CHUNK_CACHE_TTL_MS || '300000', 10)
const blobServiceClient = AZURE_CONNECTION_STRING
  ? BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING)
  : null
const SUPPORTED_EXTENSIONS = new Set([
  '.pdf', '.docx', '.doc', '.txt', '.rtf', '.odt',
  '.xlsx', '.xls', '.ods', '.csv', '.tsv',
  '.pptx', '.ppt',
  '.html', '.htm', '.xml', '.md', '.markdown', '.rst',
  '.json', '.jsonl', '.yaml', '.yml', '.toml',
  '.py', '.js', '.ts', '.jsx', '.tsx',
  '.java', '.cpp', '.c', '.h', '.cs',
  '.go', '.rb', '.php', '.swift', '.kt',
  '.r', '.sql', '.sh', '.bash', '.ps1',
  '.epub', '.eml',
])
const RESPONSE_CACHE = new Map()
const RESPONSE_CACHE_TTL = 10 * 60 * 1000
const RESPONSE_CACHE_MAX = 1000
function responseCacheGet(key) {
  const entry = RESPONSE_CACHE.get(key)
  if (!entry) return null
  if (Date.now() - entry.ts > RESPONSE_CACHE_TTL) { RESPONSE_CACHE.delete(key); return null }
  return entry.value
}
function responseCacheSet(key, value) {
  if (RESPONSE_CACHE.size >= RESPONSE_CACHE_MAX) {
    RESPONSE_CACHE.delete(RESPONSE_CACHE.keys().next().value)
  }
  RESPONSE_CACHE.set(key, { value, ts: Date.now() })
}
function getCacheKey(clientId, query) {
  return `${clientId}:${query.toLowerCase().trim()}`
}
const IN_FLIGHT = new Map()
let phiActiveCount = 0
const PHI_MAX_CONCURRENT = 3
const phiQueue = []
function runWithPhiLimit(fn) {
  return new Promise((resolve, reject) => {
    function tryRun() {
      if (phiActiveCount < PHI_MAX_CONCURRENT) {
        phiActiveCount++
        Promise.resolve().then(fn).then(
          result => { phiActiveCount--; drainPhiQueue(); resolve(result) },
          err => { phiActiveCount--; drainPhiQueue(); reject(err) }
        )
      } else {
        phiQueue.push(tryRun)
      }
    }
    tryRun()
  })
}
function drainPhiQueue() {
  if (phiQueue.length > 0 && phiActiveCount < PHI_MAX_CONCURRENT) {
    phiQueue.shift()()
  }
}
let phiFailures = 0
let phiBlockedUntil = 0
function phiCircuitOpen() {
  if (Date.now() < phiBlockedUntil) return true
  if (phiBlockedUntil > 0) { phiBlockedUntil = 0; phiFailures = 0 }
  return false
}
function phiRecordSuccess() { phiFailures = 0; phiBlockedUntil = 0 }
function phiRecordFailure() {
  phiFailures++
  if (phiFailures >= 3) {
    phiBlockedUntil = Date.now() + 30000
    console.error(`[phi4] Circuit breaker OPEN for 30s`)
  }
}
async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    return await fetch(url, { ...options, signal: controller.signal })
  } catch (err) {
    if (err.name === 'AbortError') throw new Error(`Request timed out after ${timeoutMs}ms`)
    throw err
  } finally {
    clearTimeout(timer)
  }
}
function withRequestTimeout(fn, timeoutMs = REQUEST_TIMEOUT_MS) {
  return async (req, res, next) => {
    let settled = false
    const timer = setTimeout(() => {
      if (!settled) {
        settled = true
        if (!res.headersSent) res.status(503).json({ error: 'Request timed out. Please try again.' })
      }
    }, timeoutMs)
    try {
      await fn(req, res, next)
    } catch (err) {
      if (!settled) next(err)
    } finally {
      settled = true
      clearTimeout(timer)
    }
  }
}
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
function cosineSim(a, b) {
  let dot = 0, normA = 0, normB = 0
  for (let i = 0; i < a.length; i++) {
    dot += a[i] * b[i]
    normA += a[i] * a[i]
    normB += b[i] * b[i]
  }
  return dot / (Math.sqrt(normA) * Math.sqrt(normB) + 1e-9)
}
function buildInvertedIndex(chunks) {
  const index = new Map()
  for (let i = 0; i < chunks.length; i++) {
    const words = (chunks[i].text || '').toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)
    for (const w of words) {
      if (w.length < 2) continue
      if (!index.has(w)) index.set(w, new Set())
      index.get(w).add(i)
    }
  }
  return index
}
function validateQuery(query) {
  if (!query || typeof query !== 'string') return { valid: false, message: 'Please enter a complete question to get an accurate answer.' }
  const trimmed = query.trim()
  if (trimmed.length <= 1) return { valid: false, message: 'Please enter a complete question to get an accurate answer.' }
  const words = trimmed.split(/\s+/).filter(w => w.length > 0)
  if (words.length < 2) return { valid: false, message: 'Please enter a more detailed question so I can provide an accurate answer.' }
  return { valid: true }
}
function detectQueryIntent(query) {
  const q = query.toLowerCase().trim()
  if (/^(hi|hello|hey|howdy|greetings|good\s+(morning|afternoon|evening)|how\s+are\s+you)\b/.test(q)) return 'greeting'
  if (/\b(url|link|dashboard|power\s*bi|report\s+url)\b/.test(q)) return 'url_lookup'
  if (/how\s+(is|are|was|were)\s+.+\s+(calculated|computed|determined|derived)|what\s+is\s+the\s+(formula|calculation)\s+for|how\s+do\s+you\s+(calculate|compute)/.test(q)) return 'calculation'
  if (/^(what\s+(is|are|does)|define|explain|meaning\s+of|tell\s+me\s+about|describe)\s+/.test(q) || /\b(definition|meaning)\b/.test(q)) return 'definition'
  if (/\b(vs|versus|difference|compare|between)\b/.test(q)) return 'comparison'
  if (/^(show|list|find|get|fetch|give)\s+(me\s+)?|^how\s+many\s+/.test(q)) return 'lookup'
  return 'general'
}
function extractSubject(query) {
  const q = query.toLowerCase().trim().replace(/[?!.]+$/, '')
  const patterns = [
    /^what\s+is\s+(?:an?\s+|the\s+)?(.+)$/,
    /^what\s+are\s+(.+)$/,
    /^what\s+does\s+(.+?)\s+mean$/,
    /^define\s+(?:an?\s+|the\s+)?(.+)$/,
    /^explain\s+(?:an?\s+|the\s+)?how\s+(.+?)\s+(?:is\s+)?calculated$/,
    /^explain\s+(?:an?\s+|the\s+)?(.+)$/,
    /^tell\s+me\s+about\s+(?:an?\s+|the\s+)?(.+)$/,
    /^meaning\s+of\s+(?:an?\s+|the\s+)?(.+)$/,
    /^describe\s+(?:an?\s+|the\s+)?(.+)$/,
    /^how\s+is\s+(.+?)\s+(calculated|defined|measured|computed)$/,
    /^how\s+are\s+(.+?)\s+(calculated|defined|measured|computed)$/,
    /^what\s+is\s+the\s+formula\s+for\s+(?:calculating\s+)?(?:an?\s+|the\s+)?(.+)$/,
    /^how\s+(?:do\s+you\s+)?calculate\s+(?:an?\s+|the\s+)?(.+)$/,
    /^what\s+does\s+(.+?)\s+represent/,
    /^what\s+is\s+the\s+purpose\s+of\s+(?:the\s+)?(.+?)\s+(?:attribute|measure|field|column)$/,
    /^compare\s+(.+?)\s+(?:vs|versus)\s+(.+)$/,
    /^difference\s+between\s+(.+?)\s+and\s+(.+)$/,
  ]
  for (const p of patterns) {
    const m = q.match(p)
    if (m) return m[1].trim()
  }
  return q
}
function extractUrlKeywords(query) {
  const stopWords = new Set(['power', 'bi', 'report', 'url', 'link', 'for', 'the', 'a', 'an', 'of', 'in', 'get', 'me', 'show', 'give', 'find', 'fetch'])
  return query.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1 && !stopWords.has(w))
}
function fixBrokenUrls(text) {
  return text.replace(/https:\/\/[^\s]+(\s+[^\s]+)/g, (match) => match.replace(/\s/g, ''))
}
function normalizeTerms(term) {
  const t = term.toLowerCase().trim()
  const variants = new Set([t])
  if (t.endsWith('s')) variants.add(t.slice(0, -1))
  else variants.add(t + 's')
  if (t.endsWith('ies')) variants.add(t.slice(0, -3) + 'y')
  if (t.endsWith('y')) variants.add(t.slice(0, -1) + 'ies')
  return [...variants]
}
async function callPhi4(systemPrompt, userMessage, maxTokens = 1024) {
  if (!PHI4_ENDPOINT || !PHI4_API_KEY) throw new Error('PHI4_ENDPOINT and PHI4_API_KEY are required')
  if (phiCircuitOpen()) throw new Error('Model temporarily unavailable')
  return runWithPhiLimit(async () => {
    try {
      const response = await fetchWithTimeout(
        PHI4_ENDPOINT,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${PHI4_API_KEY}` },
          body: JSON.stringify({
            model: PHI4_MODEL,
            messages: [
              { role: 'system', content: systemPrompt },
              { role: 'user', content: userMessage },
            ],
            temperature: 0.1,
            max_tokens: maxTokens,
          }),
        },
        PHI4_TIMEOUT_MS
      )
      if (!response.ok) {
        const errText = await response.text()
        throw new Error(`Phi-4 API error ${response.status}: ${errText}`)
      }
      const data = await response.json()
      phiRecordSuccess()
      return data.choices?.[0]?.message?.content || ''
    } catch (err) {
      phiRecordFailure()
      throw err
    }
  })
}
async function embedQueryAzure(query) {
  if (!AZURE_EMBED_ENDPOINT || !AZURE_EMBED_KEY) return null
  try {
    const response = await fetchWithTimeout(
      AZURE_EMBED_ENDPOINT,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': AZURE_EMBED_KEY },
        body: JSON.stringify({ input: query, model: AZURE_EMBED_MODEL }),
      },
      EMBED_TIMEOUT_MS
    )
    if (!response.ok) return null
    const data = await response.json()
    return data.data?.[0]?.embedding || null
  } catch { return null }
}
async function embedBatch(texts) {
  if (!AZURE_EMBED_ENDPOINT || !AZURE_EMBED_KEY || !texts.length) return []
  try {
    const response = await fetchWithTimeout(
      AZURE_EMBED_ENDPOINT,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': AZURE_EMBED_KEY },
        body: JSON.stringify({ input: texts, model: AZURE_EMBED_MODEL }),
      },
      EMBED_TIMEOUT_MS
    )
    if (!response.ok) return []
    const data = await response.json()
    return (data.data || []).sort((a, b) => a.index - b.index).map(d => d.embedding)
  } catch { return [] }
}
function scoreHeaderMatch(header, patterns) {
  const h = header.toLowerCase().trim()
  for (const [regex, score] of patterns) {
    if (regex.test(h)) return score
  }
  return 0
}

function detectColumns(headers) {
  const NAME_PATTERNS = [
    [/\b(measure|attribute|field|metric|kpi)\s*name\b/, 100],
    [/^name$/, 90],
    [/\bname\b/, 70],
    [/\btitle\b/, 50],
  ]
  const TABLE_PATTERNS = [
    [/\b(table|module|category|group|domain|section)\b/, 100],
    [/^table$/, 90],
  ]
  const DESC_PATTERNS = [
    [/\b(description|desc|definition|about|summary|detail)\b/, 100],
    [/^desc$/, 90],
  ]
  const FORMULA_PATTERNS = [
    [/\b(formula|calculation|calc|how\s+calculated|computed\s+as)\b/, 100],
    [/^formula$/, 90],
  ]
  const URL_PATTERNS = [
    [/\b(url|link|href|report\s+link|dashboard)\b/, 100],
  ]
  const ADDITIONAL_PATTERNS = [
    [/\b(additional|extra|notes?|info|configuration|config|mdm)\b/, 100],
  ]

  const colIdx = {}
  const scored = headers.map((h, i) => ({
    i,
    table: scoreHeaderMatch(h, TABLE_PATTERNS),
    name: scoreHeaderMatch(h, NAME_PATTERNS),
    description: scoreHeaderMatch(h, DESC_PATTERNS),
    formula: scoreHeaderMatch(h, FORMULA_PATTERNS),
    url: scoreHeaderMatch(h, URL_PATTERNS),
    additional: scoreHeaderMatch(h, ADDITIONAL_PATTERNS),
  }))

  for (const field of ['table', 'name', 'description', 'formula', 'url', 'additional']) {
    const best = scored.filter(c => c[field] > 0).sort((a, b) => b[field] - a[field])[0]
    if (best) colIdx[field] = best.i
  }
  return colIdx
}

function extractSpreadsheet(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellNF: true })
  const parts = []

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1, raw: false })
    if (!rawRows.length) continue
    let headerRowIdx = -1
    for (let i = 0; i < Math.min(15, rawRows.length); i++) {
      const cells = rawRows[i].map(c => String(c).trim()).filter(Boolean)
      if (cells.length < 2) continue
      // Skip rows that are clearly copyright/title blurbs (single long string)
      if (cells.length === 1 && cells[0].length > 60) continue
      // A good header row has short cell values (column names are rarely > 50 chars)
      const shortCells = cells.filter(c => c.length <= 60)
      if (shortCells.length >= 2) {
        headerRowIdx = i
        break
      }
    }
    if (headerRowIdx === -1) headerRowIdx = 0
    const rawHeaders = rawRows[headerRowIdx].map(h => String(h).trim())
    const headers = []
    let lastNonBlank = ''
    for (const h of rawHeaders) {
      if (h !== '') { lastNonBlank = h; headers.push(h) }
      else headers.push(lastNonBlank || `Col${headers.length + 1}`)
    }
    const colIdx = detectColumns(headers)
    parts.push(`=== Sheet: ${sheetName} ===`)
    let rowsEmitted = 0
    for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
      const row = rawRows[i]
      if (!row.some(cell => String(cell).trim() !== '')) continue
      const cells = row.map(cell => String(cell || '').replace(/\r?\n/g, ' ').trim())
      const pairs = []
      for (let j = 0; j < Math.max(headers.length, cells.length); j++) {
        const val = cells[j] || ''
        if (!val) continue
        pairs.push(`${headers[j] || `Col${j + 1}`}: ${val}`)
      }
      if (pairs.length) parts.push(pairs.join(' | '))
      const nameVal = colIdx.name !== undefined ? (cells[colIdx.name] || '').trim() : ''
      const tableVal = colIdx.table !== undefined ? (cells[colIdx.table] || '').trim() : sheetName
      const descVal = colIdx.description !== undefined ? (cells[colIdx.description] || '').trim() : ''
      const formulaVal = colIdx.formula !== undefined ? (cells[colIdx.formula] || '').trim() : ''
      const additionalVal = colIdx.additional !== undefined ? (cells[colIdx.additional] || '').trim() : ''
      const urlVal = colIdx.url !== undefined ? (cells[colIdx.url] || '').trim() : ''
      if (nameVal) {
        let synthesis = `${nameVal}`
        if (tableVal && tableVal !== sheetName) synthesis += ` (${tableVal})`
        if (descVal) synthesis += ` is defined as: ${descVal}`
        if (formulaVal) synthesis += ` Formula: ${formulaVal}`
        if (additionalVal) synthesis += ` Additional Info: ${additionalVal}`
        if (urlVal) synthesis += ` URL: ${urlVal}`
        parts.push(synthesis)
      } else if (descVal) {
        parts.push(descVal)
      }
      if (nameVal && formulaVal) {
        parts.push(`How to calculate ${nameVal}: ${formulaVal}`)
      }

      parts.push('')
      rowsEmitted++
    }
    if (rowsEmitted === 0) {
      for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
        const row = rawRows[i]
        const cells = row.map(c => String(c || '').trim()).filter(Boolean)
        if (cells.length) parts.push(cells.join(' | '))
      }
    }
  }
  return parts.join('\n')
}
function keywordSearch(query, chunks, topK, intent, invertedIndex) {
  const subject = extractSubject(query)
  const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
  const queryLower = query.toLowerCase()
  const isMultiWord = subjectWords.length > 1
  const subjectPhraseRegex = isMultiWord
    ? new RegExp(escapeRegex(subject.toLowerCase()), 'i')
    : new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b`, 'i')
  let candidateIndices
  if (invertedIndex && subjectWords.length > 0) {
    const wordsToIndex = intent === 'url_lookup' ? extractUrlKeywords(query) : subjectWords
    const union = new Set()
    for (const w of wordsToIndex) {
      for (const idx of (invertedIndex.get(w) || new Set())) union.add(idx)
      for (const variant of normalizeTerms(w)) {
        for (const idx of (invertedIndex.get(variant) || new Set())) union.add(idx)
      }
    }
    if (intent === 'url_lookup') {
      for (const w of ['url', 'link', 'http']) {
        for (const idx of (invertedIndex.get(w) || new Set())) union.add(idx)
      }
    }
    candidateIndices = union
  }
  const source = candidateIndices
    ? [...candidateIndices].map(i => chunks[i]).filter(Boolean)
    : chunks.slice(0, 200)
  return source
    .map(c => {
      const text = (c.text || '').toLowerCase()
      let score = 0
      if (intent === 'url_lookup') {
        if (!text.includes('http')) return { ...c, _score: 0 }
        const kws = extractUrlKeywords(query)
        const matched = kws.filter(w => text.includes(w)).length
        if (matched === 0) return { ...c, _score: 0 }
        score += matched * 10
        if (text.includes(kws.join(' '))) score += 15
      } else {
        const phraseFound = subjectPhraseRegex.test(c.text || '')
        if (phraseFound) {
          score += subjectWords.length * 6
          const measurePattern = new RegExp(`\\|\\s*${escapeRegex(subject.toLowerCase())}\\s*\\|`, 'i')
          if (measurePattern.test(c.text || '')) score += subjectWords.length * 4
          const defPattern = new RegExp(`${escapeRegex(subject.toLowerCase())}\\s*(is defined as|is calculated as|formula:)`, 'i')
          if (defPattern.test(c.text || '')) score += subjectWords.length * 8
        }
        const wordCoverage = subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(c.text || '')).length
        score += wordCoverage * 2
        if (text.includes(queryLower)) score += 3
      }
      return { ...c, _score: score }
    })
    .filter(c => c._score > 0)
    .sort((a, b) => b._score - a._score)
    .slice(0, topK)
}
function relaxedKeywordSearch(query, chunks, topK, invertedIndex) {
  const allWords = query.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 2)
  const union = new Set()
  if (invertedIndex) {
    for (const w of allWords) {
      for (const idx of (invertedIndex.get(w) || new Set())) union.add(idx)
      for (const variant of normalizeTerms(w)) {
        for (const idx of (invertedIndex.get(variant) || new Set())) union.add(idx)
      }
    }
  }
  const source = union.size > 0
    ? [...union].map(i => chunks[i]).filter(Boolean)
    : chunks.slice(0, 300)
  return source
    .map(c => {
      const text = (c.text || '').toLowerCase()
      const matched = allWords.filter(w => text.includes(w)).length
      return { ...c, _score: matched }
    })
    .filter(c => c._score > 0)
    .sort((a, b) => b._score - a._score)
    .slice(0, topK)
}
async function retrieveChunks(query, chunks, topK, invertedIndex) {
  const intent = detectQueryIntent(query)
  const normalizedQuery = query.toLowerCase().trim().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ')
  const candidates = keywordSearch(normalizedQuery, chunks, Math.min(150, chunks.length), intent, invertedIndex)
  const pool = candidates.length > 0 ? candidates : chunks.slice(0, 150)
  const topScore = pool[0]?._score || 0
  if (topScore >= 6) {
    return pool.slice(0, Math.min(topK, 8))
  }
  if ((intent === 'definition' || intent === 'calculation') && topScore >= 3) {
    return pool.slice(0, Math.min(topK, 8))
  }
  if (intent === 'url_lookup' && pool.length > 0) {
    return pool.slice(0, Math.min(topK, 6))
  }
  if (AZURE_EMBED_ENDPOINT && AZURE_EMBED_KEY) {
    try {
      const queryVec = await embedQueryAzure(normalizedQuery)
      if (queryVec) {
        const poolSlice = pool.slice(0, EMBED_POOL_LIMIT)
        const embeddings = await embedBatch(poolSlice.map(c => (c.text || '').slice(0, 512)))
        const maxKeyword = pool[0]?._score || 1
        const weight = (intent === 'definition' || intent === 'calculation')
          ? { semantic: 0.35, keyword: 0.65 }
          : { semantic: 0.70, keyword: 0.30 }
        const scored = poolSlice.map((c, i) => {
          if (!embeddings[i]) return c
          const semanticScore = cosineSim(queryVec, embeddings[i])
          const keywordNorm = typeof c._score === 'number' ? c._score / maxKeyword : 0
          return { ...c, _score: semanticScore * weight.semantic + keywordNorm * weight.keyword }
        })
        const remainder = pool.slice(EMBED_POOL_LIMIT).map(c => ({
          ...c,
          _score: (typeof c._score === 'number' ? c._score / maxKeyword : 0) * weight.keyword,
        }))
        const result = [...scored, ...remainder].sort((a, b) => b._score - a._score).slice(0, Math.min(topK, 8))
        if (result.length > 0) return result
      }
    } catch (err) {
      console.warn('[retrieveChunks] embed failed, using keyword fallback:', err.message)
    }
  }
  if (pool.length > 0) return pool.slice(0, Math.min(topK, 8))
  const relaxed = relaxedKeywordSearch(normalizedQuery, chunks, Math.min(topK * 2, 16), invertedIndex)
  return relaxed.slice(0, Math.min(topK, 8))
}
function buildContext(hits) {
  const seen = new Set()
  const deduped = []
  for (const h of hits) {
    const fp = (h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fp)) {
      seen.add(fp)
      deduped.push(h)
    }
    if (deduped.length >= 6) break
  }
  return deduped.map((h, i) => {
    const limit = i === 0 ? 1200 : 900
    return `[Source ${i + 1}]\n${(h.text || '').trim().slice(0, limit)}`
  }).join('\n\n---\n\n')
}
function buildSystemPrompt(intent) {
  const base = `You are a helpful data dictionary assistant for a real estate analytics platform.
Answer ONLY using the provided context. Follow these STRICT formatting rules:

FORMATTING RULES:
1. Always use **bold** for field/measure names and labels like **Definition:** and **Formula:**.
2. Never output raw pipe-separated data rows (like "Name: X | Description: Y"). Convert to prose.
3. Write in complete English sentences. Never cut off mid-sentence. End with a period.
4. For definitions: use the pattern — "**[Name]** is [description]. **Formula:** [formula]."
5. For comparisons: use a small table or two bullet sections, one per item.
6. For URL lookups: output the full URL on its own line, nothing before or after.
7. For lists: use bullet points (- item).
8. Keep answers concise: 3–6 sentences for definitions, more only if formula is complex.
9. Do NOT include source references like [1], [2] or sheet names in your answer.
10. If the context does not contain the answer, say exactly: "I could not find this in your documents."`

  const intentGuide = {
    definition: `
RESPONSE FORMAT for definition:
**[Measure/Attribute Name]** is [one sentence description].
**Formula:** [formula in plain English, if available].
[Optional: one sentence of additional context if present in source.]`,
    calculation: `
RESPONSE FORMAT for formula/calculation:
**How [Name] is Calculated**
[Step-by-step explanation in plain English.]
**Formula:** [exact formula text].`,
    comparison: `
RESPONSE FORMAT for comparison:
Compare the two items clearly. Use this structure:
**[Item A]:** [description + formula if any]
**[Item B]:** [description + formula if any]
**Key Difference:** [one sentence summary].`,
    lookup: `
RESPONSE FORMAT for lookup:
State the exact value or list. Use bullet points if multiple items.`,
    url_lookup: `
RESPONSE FORMAT for URL:
Return ONLY the exact full URL. No other text.`,
    general: `
RESPONSE FORMAT: Answer directly in 2–5 complete sentences using bold labels where helpful.`,
  }
  return base + (intentGuide[intent] || intentGuide.general)
}

function buildUserMessage(query, hits, intent) {
  const context = buildContext(hits)
  const subject = extractSubject(query)
  let instruction = ''
  if (intent === 'definition') {
    instruction = `\n\nUsing ONLY the context above, write a well-formatted definition of "${subject}". Bold the name. Include definition and formula if present. Do NOT copy raw data rows. End with a period.`
  } else if (intent === 'calculation') {
    instruction = `\n\nUsing ONLY the context above, explain exactly how "${subject}" is calculated. Bold the formula label. State all steps. Do NOT copy raw data rows. End with a period.`
  } else if (intent === 'url_lookup') {
    instruction = `\n\nUsing ONLY the context above, return the full URL related to "${extractUrlKeywords(query).join(' ')}". Return ONLY the URL, nothing else.`
  } else if (intent === 'comparison') {
    instruction = `\n\nUsing ONLY the context above, compare these items clearly: ${query}. Use bold labels for each item. End with a period.`
  } else {
    instruction = `\n\nUsing ONLY the context above, answer this question completely and in a well-formatted way: ${query} End with a period.`
  }
  return `CONTEXT:\n${context}${instruction}`
}
function buildFallbackAnswer(query, hits) {
  if (!hits || hits.length === 0) {
    return "I could not find relevant information about this in your documents."
  }
  const intent = detectQueryIntent(query)
  const subject = extractSubject(query)
  const subjectLower = subject.toLowerCase()

  if (intent === 'url_lookup') {
    const urlKeywords = extractUrlKeywords(query)
    for (const h of hits) {
      for (const line of (h.text || '').split('\n')) {
        if (!line.toLowerCase().includes('http')) continue
        const matched = urlKeywords.filter(w => line.toLowerCase().includes(w)).length
        if (matched > 0) {
          const urlMatch = line.match(/https?:\/\/\S+/)
          if (urlMatch) return urlMatch[0].replace(/\s/g, '')
        }
      }
    }
    return "I could not find a matching URL in your documents."
  }
  const synthesisPattern = new RegExp(
    `${escapeRegex(subjectLower)}[^\\n]*is defined as:\\s*([^\\n]+?)(?:\\s+Formula:\\s*([^\\n]+))?(?:\\s+Additional Info:\\s*([^\\n]+))?$`,
    'im'
  )
  for (const h of hits) {
    const m = (h.text || '').match(synthesisPattern)
    if (m) {
      const desc = (m[1] || '').trim().slice(0, 600)
      const formula = (m[2] || '').trim().slice(0, 400)
      const additional = (m[3] || '').trim().slice(0, 300)
      const cap = subject.charAt(0).toUpperCase() + subject.slice(1)
      let answer = `**${cap}** is ${desc}.`
      if (formula) answer += `\n\n**Formula:** ${formula}.`
      if (additional) answer += `\n\n**Additional Info:** ${additional}.`
      return answer
    }
  }

  // Secondary: look for "How to calculate X: formula" lines
  if (intent === 'calculation') {
    const calcPattern = new RegExp(`how to calculate ${escapeRegex(subjectLower)}:\\s*([^\\n]+)`, 'im')
    for (const h of hits) {
      const m = (h.text || '').match(calcPattern)
      if (m) {
        const cap = subject.charAt(0).toUpperCase() + subject.slice(1)
        return `**How ${cap} is Calculated**\n\n**Formula:** ${m[1].trim().slice(0, 500)}.`
      }
    }
  }

  // Tertiary: extract meaningful lines that mention the subject
  const matchingLines = []
  for (const h of hits) {
    for (const line of (h.text || '').split('\n')) {
      const ll = line.toLowerCase()
      if (!ll.includes(subjectLower)) continue
      if (line.trim().length <= 20) continue
      // Skip raw pipe-heavy rows (more than 3 pipes = raw spreadsheet row)
      if ((line.match(/\|/g) || []).length > 3) continue
      const cleaned = line.trim()
        .replace(/^===\s*Sheet:.*===\s*$/, '')
        .replace(/\(from\s+[A-Za-z\s]+\)/g, '')
        .trim()
      if (cleaned.length > 15) matchingLines.push(cleaned)
    }
  }
  if (matchingLines.length > 0) {
    const unique = [...new Set(matchingLines)].slice(0, 3)
    const cap = subject.charAt(0).toUpperCase() + subject.slice(1)
    return `**${cap}:** ${unique.join(' ').slice(0, 600)}.`
  }

  return "I could not find that specific information in your documents."
}
// ───────────────────────────────────────────────────────────────────────────────

function cleanAnswer(rawAnswer) {
  if (!rawAnswer) return ''
  let cleaned = fixBrokenUrls(rawAnswer)
    .replace(/^\s*\[Source\s*\d+\]\s*/gm, '')
    // Remove pure pipe-delimited rows (4+ pipes = raw spreadsheet row)
    .replace(/^[^\n]*\|[^\n]*\|[^\n]*\|[^\n]*\|[^\n]*$/gm, '')
    .replace(/=== .+ ===\s*/gm, '')
    .replace(/\(from\s+[A-Za-z\s]+\)\s*/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
  if (cleaned.length > 0 && !/[.!?]$/.test(cleaned)) cleaned += '.'
  return cleaned
}
async function answerWithPhi4(query, hits, intent) {
  const systemPrompt = buildSystemPrompt(intent)
  const userMessage = buildUserMessage(query, hits, intent)
  return callPhi4(systemPrompt, userMessage, 1024)
}
async function extractPdf(buffer) {
  const r = await pdfParse(buffer)
  return r.text || ''
}
async function extractWord(buffer) {
  const r = await mammoth.extractRawText({ buffer })
  return r.value || ''
}
function extractCsv(buffer, delimiter) {
  const text = buffer.toString('utf-8')
  const result = Papa.parse(text, { header: true, skipEmptyLines: true, delimiter })
  if (!result.data?.length) return text
  return result.data.map((row, i) => `Row ${i + 1}: ` + Object.entries(row).map(([k, v]) => `${k}=${v}`).join(' | ')).join('\n')
}
async function extractOffice(buffer) {
  return new Promise((resolve, reject) => {
    parseOffice(buffer, (text, err) => {
      if (err) reject(err)
      else resolve(text || '')
    }, { outputErrorToConsole: false })
  })
}
function extractHtml(buffer) {
  const root = htmlParse(buffer.toString('utf-8'))
  root.querySelectorAll('script, style').forEach(n => n.remove())
  return root.structuredText || root.innerText || root.rawText || ''
}
function extractXml(buffer) { return buffer.toString('utf-8').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() }
function extractJson(buffer) { try { return JSON.stringify(JSON.parse(buffer.toString('utf-8')), null, 2) } catch { return buffer.toString('utf-8') } }
function extractJsonl(buffer) {
  return buffer.toString('utf-8').split('\n').filter(Boolean)
    .map(line => { try { return JSON.stringify(JSON.parse(line)) } catch { return line } })
    .join('\n')
}
function extractYaml(buffer) { try { return JSON.stringify(yaml.load(buffer.toString('utf-8')), null, 2) } catch { return buffer.toString('utf-8') } }
async function extractEml(buffer) {
  const parsed = await simpleParser(buffer)
  const parts = []
  if (parsed.subject) parts.push(`Subject: ${parsed.subject}`)
  if (parsed.from) parts.push(`From: ${parsed.from.text}`)
  if (parsed.to) parts.push(`To: ${parsed.to.text}`)
  if (parsed.date) parts.push(`Date: ${parsed.date}`)
  if (parsed.text) parts.push(`\n${parsed.text}`)
  else if (parsed.html) parts.push(`\n${extractHtml(Buffer.from(parsed.html))}`)
  return parts.join('\n')
}
async function extractEpub(buffer) {
  return new Promise(resolve => {
    parseOffice(buffer, (text, err) => {
      resolve(err || !text ? '[EPUB: convert to PDF for best results]' : text)
    }, { outputErrorToConsole: false })
  })
}
async function extractTextFromBuffer(buffer, fileName) {
  const ext = ('.' + fileName.split('.').pop()).toLowerCase()
  if (ext === '.pdf') return extractPdf(buffer)
  if (ext === '.docx' || ext === '.doc') return extractWord(buffer)
  if (ext === '.odt' || ext === '.rtf') return extractOffice(buffer)
  if (['.xlsx', '.xls', '.ods'].includes(ext)) return extractSpreadsheet(buffer)
  if (ext === '.csv') return extractCsv(buffer, ',')
  if (ext === '.tsv') return extractCsv(buffer, '\t')
  if (ext === '.pptx' || ext === '.ppt') return extractOffice(buffer)
  if (ext === '.html' || ext === '.htm') return extractHtml(buffer)
  if (ext === '.xml') return extractXml(buffer)
  if (['.md', '.markdown', '.rst'].includes(ext)) return buffer.toString('utf-8')
  if (ext === '.json') return extractJson(buffer)
  if (ext === '.jsonl') return extractJsonl(buffer)
  if (ext === '.yaml' || ext === '.yml') return extractYaml(buffer)
  if (ext === '.toml') return buffer.toString('utf-8')
  if (ext === '.epub') return extractEpub(buffer)
  if (ext === '.eml') return extractEml(buffer)
  const plainText = new Set(['.txt', '.py', '.js', '.ts', '.jsx', '.tsx', '.java', '.cpp', '.c', '.h', '.cs', '.go', '.rb', '.php', '.swift', '.kt', '.r', '.sql', '.sh', '.bash', '.ps1'])
  if (plainText.has(ext)) return buffer.toString('utf-8')
  return ''
}

// ─── IMPROVED CHUNKER ──────────────────────────────────────────────────────────
// Spreadsheet synthesis lines ("X is defined as: ...") are short (100-300 chars).
// We group them in PAIRS so each chunk contains both the raw row AND its synthesis,
// keeping context together. This prevents the model from seeing the synthesis line
// without the name, or the name without the description.

function chunkText(text, sourceFile) {
  const chunks = []
  let chunkIndex = 0
  const blocks = text.replace(/\r\n/g, '\n').split(/\n{2,}/).map(b => b.trim()).filter(b => b.length > 0)
  let buffer = []
  let bufferLength = 0

  function flush() {
    const chunkStr = buffer.join('\n\n')
    if (chunkStr.length >= 30) {
      chunks.push({ text: chunkStr, source_file: sourceFile, chunk_index: chunkIndex++, embedding: [] })
    }
    buffer = []
    bufferLength = 0
  }

  for (let bi = 0; bi < blocks.length; bi++) {
    const block = blocks[bi]
    if (block.length > CHUNK_SIZE * 1.5) {
      if (buffer.length > 0) flush()
      const lines = block.split('\n').filter(l => l.trim())
      let lineBuffer = []
      let lineLength = 0
      for (const line of lines) {
        const projected = lineLength + (lineBuffer.length ? 1 : 0) + line.length
        if (lineBuffer.length > 0 && projected > CHUNK_SIZE) {
          const s = lineBuffer.join('\n')
          if (s.length >= 30) chunks.push({ text: s, source_file: sourceFile, chunk_index: chunkIndex++, embedding: [] })
          lineBuffer = lineBuffer.slice(-CHUNK_OVERLAP)
          lineLength = lineBuffer.join('\n').length
        }
        lineBuffer.push(line)
        lineLength += (lineLength ? 1 : 0) + line.length
      }
      if (lineBuffer.length) {
        const s = lineBuffer.join('\n')
        if (s.length >= 30) chunks.push({ text: s, source_file: sourceFile, chunk_index: chunkIndex++, embedding: [] })
      }
      continue
    }
    const projected = bufferLength + (bufferLength ? 2 : 0) + block.length
    if (buffer.length > 0 && projected > CHUNK_SIZE) {
      const lastBlock = buffer[buffer.length - 1] || ''
      flush()
      if (lastBlock) { buffer.push(lastBlock); bufferLength = lastBlock.length }
    }
    buffer.push(block)
    bufferLength += (bufferLength ? 2 : 0) + block.length
  }
  if (buffer.length > 0) flush()
  return chunks
}
// ───────────────────────────────────────────────────────────────────────────────

async function downloadBlobAsBuffer(containerClient, blobName) {
  const download = await containerClient.getBlobClient(blobName).download()
  const parts = []
  for await (const chunk of download.readableStreamBody) {
    parts.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
  }
  return Buffer.concat(parts)
}
async function _doLoadChunks(clientId) {
  if (!blobServiceClient) throw new Error('AZURE_CONNECTION_STRING not set')
  const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
  const prefix = `${RAW_PREFIX}/${clientId}/`
  const blobNames = []
  for await (const blob of containerClient.listBlobsFlat({ prefix })) {
    const fileName = blob.name.split('/').pop()
    const ext = ('.' + fileName.split('.').pop()).toLowerCase()
    if (SUPPORTED_EXTENSIONS.has(ext)) blobNames.push(blob.name)
  }
  const allChunks = []
  for (let i = 0; i < blobNames.length; i += BLOB_CONCURRENCY) {
    const batch = blobNames.slice(i, i + BLOB_CONCURRENCY)
    const results = await Promise.allSettled(
      batch.map(async (blobName) => {
        const fileName = blobName.split('/').pop()
        const buffer = await downloadBlobAsBuffer(containerClient, blobName)
        const text = await extractTextFromBuffer(buffer, fileName)
        if (!text?.trim()) return []
        return chunkText(text, fileName)
      })
    )
    for (const result of results) {
      if (result.status === 'fulfilled') allChunks.push(...result.value)
      else console.warn('[loadChunks] blob failed:', result.reason?.message)
    }
  }
  return allChunks
}
const CHUNK_CACHE = new Map()
async function loadChunksForClient(clientId) {
  const now = Date.now()
  const cached = CHUNK_CACHE.get(clientId)
  if (cached && cached.chunks) {
    if (now - cached.ts <= CHUNK_CACHE_TTL) {
      return { chunks: cached.chunks, invertedIndex: cached.invertedIndex }
    }
    if (!cached.loading) {
      const refreshPromise = _doLoadChunks(clientId)
        .then(chunks => {
          const invertedIndex = buildInvertedIndex(chunks)
          CHUNK_CACHE.set(clientId, { chunks, invertedIndex, ts: Date.now(), loading: null })
          console.log(`[chunkCache] Background refresh done for ${clientId}: ${chunks.length} chunks`)
        })
        .catch(err => {
          const existing = CHUNK_CACHE.get(clientId)
          if (existing) CHUNK_CACHE.set(clientId, { ...existing, loading: null })
          console.warn(`[chunkCache] Background refresh failed for ${clientId}: ${err.message}`)
        })
      CHUNK_CACHE.set(clientId, { ...cached, loading: refreshPromise })
    }
    return { chunks: cached.chunks, invertedIndex: cached.invertedIndex }
  }
  if (cached && cached.loading) {
    await cached.loading
    const entry = CHUNK_CACHE.get(clientId)
    return { chunks: entry?.chunks || [], invertedIndex: entry?.invertedIndex || null }
  }
  const loadPromise = _doLoadChunks(clientId)
    .then(chunks => {
      const invertedIndex = buildInvertedIndex(chunks)
      CHUNK_CACHE.set(clientId, { chunks, invertedIndex, ts: Date.now(), loading: null })
      console.log(`[chunkCache] Loaded ${chunks.length} chunks for ${clientId}`)
      return chunks
    })
    .catch(err => {
      CHUNK_CACHE.set(clientId, { chunks: null, invertedIndex: null, ts: 0, loading: null })
      throw err
    })
  CHUNK_CACHE.set(clientId, { chunks: null, invertedIndex: null, ts: 0, loading: loadPromise })
  await loadPromise
  const entry = CHUNK_CACHE.get(clientId)
  return { chunks: entry?.chunks || [], invertedIndex: entry?.invertedIndex || null }
}
function invalidateChunkCache(clientId) {
  CHUNK_CACHE.delete(clientId)
  console.log(`[chunkCache] Invalidated cache for client: ${clientId}`)
}
function warmupChunkCaches() {
  if (!WARMUP_CLIENT_IDS.length || !blobServiceClient) return
  console.log(`[warmup] Pre-loading chunks for: ${WARMUP_CLIENT_IDS.join(', ')}`)
  for (const clientId of WARMUP_CLIENT_IDS) {
    loadChunksForClient(clientId)
      .then(({ chunks }) => console.log(`[warmup] ${clientId} — ${chunks.length} chunks ready`))
      .catch(err => console.warn(`[warmup] ${clientId} — ${err.message}`))
  }
}
let db = null
async function getDb() {
  if (db) return db
  const client = new MongoClient(MONGODB_URI)
  await client.connect()
  db = client.db(MONGODB_DB)
  await db.collection('clients').createIndex({ apiKey: 1 }, { unique: true, sparse: true })
  return db
}
let chatDb = null
async function getChatDb() {
  if (chatDb) return chatDb
  const uri = CHAT_HISTORY_URI || MONGODB_URI
  const client = new MongoClient(uri)
  await client.connect()
  chatDb = client.db(CHAT_HISTORY_DB)
  return chatDb
}
const CLIENT_CACHE = new Map()
const CACHE_TTL_MS = 5 * 60 * 1000
function getCached(apiKey) {
  const entry = CLIENT_CACHE.get(apiKey)
  if (!entry) return null
  if (Date.now() - entry.cachedAt > CACHE_TTL_MS) { CLIENT_CACHE.delete(apiKey); return null }
  return entry
}
function setCache(apiKey, data) { CLIENT_CACHE.set(apiKey, { ...data, cachedAt: Date.now() }) }
function evictCache(apiKey) { if (apiKey) CLIENT_CACHE.delete(apiKey) }
async function verifyApiKey(apiKey) {
  if (!apiKey || !apiKey.startsWith('rak_')) return null
  const cached = getCached(apiKey)
  if (cached) return { clientId: cached.clientId, name: cached.name }
  const database = await getDb()
  const client = await database.collection('clients').findOne({ apiKey }, { projection: { clientId: 1, name: 1, _id: 0 } })
  if (!client) return null
  setCache(apiKey, { clientId: client.clientId, name: client.name })
  return { clientId: client.clientId, name: client.name }
}
function startApiKeyHealthChecker() {
  if (!MONGODB_URI) return
  setInterval(async () => {
    const keys = [...CLIENT_CACHE.keys()]
    if (!keys.length) return
    try {
      const database = await getDb()
      const validDocs = await database.collection('clients').find({ apiKey: { $in: keys } }, { projection: { apiKey: 1, _id: 0 } }).toArray()
      const validSet = new Set(validDocs.map(d => d.apiKey))
      for (const key of keys) if (!validSet.has(key)) evictCache(key)
    } catch { }
  }, KEY_CHECK_INTERVAL_MS)
}
function extractApiKey(req) {
  const header = req.headers['authorization'] || ''
  return header.startsWith('Bearer ') ? header.slice(7).trim() : null
}
async function requireClientKey(req, res, next) {
  const apiKey = extractApiKey(req) || req.body?.apiKey
  if (!apiKey) return res.status(401).json({ error: 'Missing API key' })
  const client = await verifyApiKey(apiKey)
  if (!client) return res.status(401).json({ error: 'Invalid or expired API key' })
  req.client = client
  next()
}
function requireAdminKey(req, res, next) {
  const key = extractApiKey(req)
  if (!key || key !== ADMIN_API_KEY) return res.status(401).json({ error: 'Unauthorized' })
  next()
}
function generateApiKey() {
  return `rak_${crypto.randomBytes(32).toString('hex')}`
}
function generateTitle(query) {
  const cleaned = query.trim().replace(/[?!.]+$/, '')
  return cleaned.length > 50 ? cleaned.slice(0, 50) + '…' : cleaned
}
app.get('/health', (req, res) => res.json({
  ok: true,
  service: 'ask-data',
  model: PHI4_MODEL,
  embeddings: AZURE_EMBED_ENDPOINT ? 'azure' : 'keyword-only',
  chunkCacheSize: CHUNK_CACHE.size,
  responseCacheSize: RESPONSE_CACHE.size,
  phiCircuitOpen: phiCircuitOpen(),
  phiFailures,
}))
app.post('/client/verify', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ valid: false, error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ valid: false, error: 'Invalid or expired API key' })
    res.json({ valid: true, client })
  } catch (err) { res.status(500).json({ valid: false, error: err.message }) }
})
app.post('/admin/clients', requireAdminKey, async (req, res) => {
  try {
    let { name, clientId, apiKey } = req.body
    if (!name || !clientId) return res.status(400).json({ error: 'name and clientId are required' })
    if (!apiKey) {
      apiKey = generateApiKey()
    } else if (!apiKey.startsWith('rak_')) {
      return res.status(400).json({ error: 'apiKey must start with "rak_"' })
    }
    const database = await getDb()
    const col = database.collection('clients')
    const existing = await col.findOne({ $or: [{ clientId }, { apiKey }] })
    if (existing) {
      const field = existing.clientId === clientId ? 'clientId' : 'apiKey'
      return res.status(409).json({ error: `A client with this ${field} already exists` })
    }
    const now = new Date().toISOString()
    const doc = { name: name.trim(), clientId: clientId.trim().toLowerCase(), apiKey, apiKeyRotatedAt: now, folderLink: '', sourceType: 'google-drive', status: 'idle', documentsCount: 0, autoSync: false, watchIntervalMs: 300000, lastRunAt: null, lastError: null, createdAt: now, updatedAt: now }
    const result = await col.insertOne(doc)
    res.status(201).json({ ...doc, _id: result.insertedId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.get('/admin/clients', requireAdminKey, async (req, res) => {
  try {
    const database = await getDb()
    const clients = await database.collection('clients').find({}, { projection: { apiKey: 0 } }).sort({ createdAt: -1 }).toArray()
    res.json({ clients })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.get('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const database = await getDb()
    const client = await database.collection('clients').findOne({ clientId: req.params.clientId })
    if (!client) return res.status(404).json({ error: 'Client not found' })
    res.json(client)
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/admin/clients/:clientId/regenerate-key', requireAdminKey, async (req, res) => {
  try {
    const database = await getDb()
    const col = database.collection('clients')
    const oldClient = await col.findOne({ clientId: req.params.clientId }, { projection: { apiKey: 1 } })
    if (!oldClient) return res.status(404).json({ error: 'Client not found' })
    const newApiKey = generateApiKey()
    const now = new Date().toISOString()
    if (oldClient.apiKey) evictCache(oldClient.apiKey)
    await col.findOneAndUpdate({ clientId: req.params.clientId }, { $set: { apiKey: newApiKey, apiKeyRotatedAt: now, updatedAt: now } }, { returnDocument: 'after' })
    res.json({ success: true, clientId: req.params.clientId, newApiKey, apiKeyRotatedAt: now })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.patch('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const database = await getDb()
    const updates = { ...req.body, updatedAt: new Date().toISOString() }
    if (updates.apiKey !== undefined) {
      if (!updates.apiKey.startsWith('rak_')) return res.status(400).json({ error: 'apiKey must start with "rak_"' })
      const old = await database.collection('clients').findOne({ clientId: req.params.clientId }, { projection: { apiKey: 1 } })
      if (old?.apiKey) evictCache(old.apiKey)
      updates.apiKeyRotatedAt = new Date().toISOString()
    }
    const result = await database.collection('clients').findOneAndUpdate({ clientId: req.params.clientId }, { $set: updates }, { returnDocument: 'after' })
    if (!result) return res.status(404).json({ error: 'Client not found' })
    res.json(result)
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.delete('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const { clientId } = req.params
    const database = await getDb()
    const client = await database.collection('clients').findOne({ clientId })
    if (!client) return res.status(404).json({ error: 'Client not found' })
    if (client.apiKey) evictCache(client.apiKey)
    await database.collection('clients').deleteOne({ clientId })
    invalidateChunkCache(clientId)
    const blobsDeleted = [], blobsFailed = []
    if (blobServiceClient) {
      try {
        const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
        for (const prefix of [`raw/${clientId}/`, `meta/${clientId}/`]) {
          for await (const blob of containerClient.listBlobsFlat({ prefix })) {
            try { await containerClient.deleteBlob(blob.name); blobsDeleted.push(blob.name) }
            catch (e) { blobsFailed.push({ name: blob.name, error: e.message }) }
          }
        }
      } catch (azureErr) { blobsFailed.push({ name: 'azure-connection', error: azureErr.message }) }
    }
    res.json({ ok: true, deleted: clientId, blobsDeleted: blobsDeleted.length, blobsFailed: blobsFailed.length > 0 ? blobsFailed : undefined })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/admin/clients/:clientId/invalidate-cache', requireAdminKey, (req, res) => {
  invalidateChunkCache(req.params.clientId)
  RESPONSE_CACHE.clear()
  res.json({ ok: true, clientId: req.params.clientId, message: 'Chunk + response cache invalidated' })
})
app.post('/client/login', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ error: 'Invalid API key' })
    if (blobServiceClient) {
      loadChunksForClient(client.clientId).catch(err => console.warn(`[login warmup] ${client.clientId}: ${err.message}`))
    }
    res.json({ ok: true, client })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/login', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ error: 'Invalid API key' })
    if (blobServiceClient) {
      loadChunksForClient(client.clientId).catch(err => console.warn(`[chat/login warmup] ${client.clientId}: ${err.message}`))
    }
    res.json({ ok: true, client })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.get('/client/me', requireClientKey, async (req, res) => {
  try {
    const database = await getDb()
    const client = await database.collection('clients').findOne({ clientId: req.client.clientId }, { projection: { apiKey: 0 } })
    if (!client) return res.status(404).json({ error: 'Client not found' })
    res.json(client)
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/conversations', requireClientKey, async (req, res) => {
  try {
    const { title } = req.body
    const database = await getChatDb()
    const now = new Date()
    const conversation = { clientId: req.client.clientId, title: title || 'New Conversation', messages: [], createdAt: now, updatedAt: now }
    const result = await database.collection('conversations').insertOne(conversation)
    res.status(201).json({ ...conversation, _id: result.insertedId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/conversations/list', requireClientKey, async (req, res) => {
  try {
    const database = await getChatDb()
    const conversations = await database.collection('conversations').find({ clientId: req.client.clientId }, { projection: { messages: 0 } }).sort({ updatedAt: -1 }).toArray()
    res.json({ conversations })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/conversations/get', requireClientKey, async (req, res) => {
  try {
    const { conversationId } = req.body
    if (!conversationId) return res.status(400).json({ error: 'conversationId is required' })
    const database = await getChatDb()
    const conversation = await database.collection('conversations').findOne({ _id: new ObjectId(conversationId), clientId: req.client.clientId })
    if (!conversation) return res.status(404).json({ error: 'Conversation not found' })
    res.json(conversation)
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/conversations/rename', requireClientKey, async (req, res) => {
  try {
    const { conversationId, title } = req.body
    if (!conversationId || !title) return res.status(400).json({ error: 'conversationId and title are required' })
    const database = await getChatDb()
    const result = await database.collection('conversations').findOneAndUpdate({ _id: new ObjectId(conversationId), clientId: req.client.clientId }, { $set: { title: title.trim(), updatedAt: new Date() } }, { returnDocument: 'after', projection: { messages: 0 } })
    if (!result) return res.status(404).json({ error: 'Conversation not found' })
    res.json(result)
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/conversations/delete', requireClientKey, async (req, res) => {
  try {
    const { conversationId } = req.body
    if (!conversationId) return res.status(400).json({ error: 'conversationId is required' })
    const database = await getChatDb()
    const result = await database.collection('conversations').deleteOne({ _id: new ObjectId(conversationId), clientId: req.client.clientId })
    if (result.deletedCount === 0) return res.status(404).json({ error: 'Conversation not found' })
    res.json({ ok: true, deleted: conversationId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/message', requireClientKey, withRequestTimeout(async (req, res) => {
  try {
    const { query, topK = 6, conversationId } = req.body
    if (!query?.trim()) return res.status(400).json({ error: 'query is required' })
    const validation = validateQuery(query)
    if (!validation.valid) {
      return res.json({
        answer: validation.message,
        sources: [],
        conversationId: conversationId || null,
        client: req.client,
      })
    }
    const { clientId, name } = req.client
    const intent = detectQueryIntent(query.trim())
    if (intent === 'greeting') {
      return res.json({
        answer: "Hello! I'm your data dictionary assistant. Ask me anything about your measures, attributes, or reports.",
        sources: [],
        conversationId: conversationId || null,
        client: { clientId, name },
      })
    }
    const cacheKey = getCacheKey(clientId, query)
    const cached = responseCacheGet(cacheKey)
    if (cached) {
      return res.json({ ...cached, cached: true, conversationId: conversationId || cached.conversationId })
    }
    if (IN_FLIGHT.has(cacheKey)) {
      try {
        const result = await IN_FLIGHT.get(cacheKey)
        return res.json({ ...result, conversationId: conversationId || result.conversationId })
      } catch { }
    }
    const requestPromise = (async () => {
      const { chunks, invertedIndex } = await loadChunksForClient(clientId)
      if (chunks.length === 0) {
        return {
          answer: 'No documents found for your account. Please ensure your documents have been ingested first.',
          sources: [],
          conversationId: conversationId || null,
          client: { clientId, name },
        }
      }
      let hits = await retrieveChunks(query.trim(), chunks, Math.min(topK, 8), invertedIndex)
      if (hits.length === 0) {
        hits = relaxedKeywordSearch(query.trim(), chunks, 12, invertedIndex)
      }
      console.log(`[chat/message] "${query.slice(0, 60)}" → intent=${intent}, hits=${hits.length}, topScore=${hits[0]?._score?.toFixed(2) || 0}`)
      if (hits.length === 0) {
        return {
          answer: "I could not find relevant information about this in your documents. Try rephrasing your question.",
          sources: [],
          conversationId: conversationId || null,
          client: { clientId, name },
        }
      }
      let rawAnswer = ''
      try {
        rawAnswer = await Promise.race([
          answerWithPhi4(query.trim(), hits, intent),
          new Promise((_, reject) => setTimeout(() => reject(new Error('Model timeout')), 25000)),
        ])
      } catch (err) {
        console.warn(`[phi4] Using fallback: ${err.message}`)
      }
      const isBlank = !rawAnswer || rawAnswer.trim().length < 15
      const answer = isBlank
        ? buildFallbackAnswer(query.trim(), hits)
        : cleanAnswer(rawAnswer)
      if (isBlank) console.warn(`[phi4] Blank response, used fallback for: "${query.slice(0, 60)}"`)
      const sources = hits.map(h => ({
        source_file: h.source_file || 'unknown',
        chunk_index: h.chunk_index ?? 0,
        score: typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
        preview: (h.text || '').slice(0, 200),
      }))
      let activeConversationId = conversationId || null
      try {
        const chatDatabase = await getChatDb()
        const col = chatDatabase.collection('conversations')
        const now = new Date()
        const userMsg = { role: 'user', content: query.trim(), timestamp: now }
        const assistantMsg = { role: 'assistant', content: answer, sources: sources.map(s => ({ source_file: s.source_file, score: s.score })), timestamp: now }
        if (activeConversationId) {
          const updated = await col.findOneAndUpdate(
            { _id: new ObjectId(activeConversationId), clientId },
            { $push: { messages: { $each: [userMsg, assistantMsg] } }, $set: { updatedAt: now } },
            { returnDocument: 'after', projection: { _id: 1 } }
          )
          if (!updated) activeConversationId = null
        }
        if (!activeConversationId) {
          const result = await col.insertOne({ clientId, title: generateTitle(query.trim()), messages: [userMsg, assistantMsg], createdAt: now, updatedAt: now })
          activeConversationId = result.insertedId.toString()
        }
      } catch (saveErr) {
        console.warn('[chat/message] Failed to save conversation:', saveErr.message)
      }
      return { answer, sources, conversationId: activeConversationId, client: { clientId, name } }
    })()
    IN_FLIGHT.set(cacheKey, requestPromise)
    let result
    try {
      result = await requestPromise
    } finally {
      IN_FLIGHT.delete(cacheKey)
    }
    if (result.answer && result.answer.length > 15) {
      responseCacheSet(cacheKey, result)
    }
    res.json(result)
  } catch (err) {
    console.error('[chat/message] Error:', err.message)
    if (!res.headersSent) res.status(500).json({ error: err.message })
  }
}))
app.use((err, req, res, next) => {
  console.error('[global error handler]', err)
  if (!res.headersSent) res.status(500).json({ error: 'An unexpected error occurred. Please try again.' })
})
const PORT = process.env.PORT || 4000
app.listen(PORT, () => {
  console.log(`rag-client-auth running on port ${PORT}`)
  console.log(`Model: ${PHI4_MODEL} | Endpoint: ${PHI4_ENDPOINT ? 'configured' : 'MISSING'}`)
  startApiKeyHealthChecker()
  warmupChunkCaches()
})
module.exports = app
