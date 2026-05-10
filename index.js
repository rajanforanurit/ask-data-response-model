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
const KEYWORD_SHORTCIRCUIT_SCORE = parseInt(process.env.KEYWORD_SHORTCIRCUIT_SCORE || '6', 10)
const WARMUP_CLIENT_IDS = (process.env.WARMUP_CLIENT_IDS || '').split(',').map(s => s.trim()).filter(Boolean)

// ─── OPTIMIZATION 1: Larger chunks to preserve multi-line formulas ─────────────
const RAW_PREFIX    = 'raw'
const CHUNK_SIZE    = 900   // increased from 500 — prevents formula/description splits
const CHUNK_OVERLAP = 3     // slightly more overlap for context continuity

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

const DOC_TYPE = {
  SPREADSHEET:  'spreadsheet',
  PDF:          'pdf',
  WORD:         'word',
  PRESENTATION: 'presentation',
  CODE:         'code',
  DATA:         'data',
  TEXT:         'text',
  EMAIL:        'email',
  WEB:          'web',
  UNKNOWN:      'unknown',
}

// ─── Response cache ───────────────────────────────────────────────────────────
const RESPONSE_CACHE     = new Map()
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
    const firstKey = RESPONSE_CACHE.keys().next().value
    RESPONSE_CACHE.delete(firstKey)
  }
  RESPONSE_CACHE.set(key, { value, ts: Date.now() })
}

function getCacheKey(clientId, query) {
  return `${clientId}:${query.toLowerCase().trim()}`
}

// ─── Phi-4 concurrency / circuit breaker ─────────────────────────────────────
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
          err    => { phiActiveCount--; drainPhiQueue(); reject(err) }
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
    const next = phiQueue.shift()
    next()
  }
}

let phiFailures     = 0
let phiBlockedUntil = 0

function phiCircuitOpen() {
  if (Date.now() < phiBlockedUntil) return true
  if (phiBlockedUntil > 0) {
    phiBlockedUntil = 0
    phiFailures     = 0
    console.log('[phi4] Circuit breaker reset (timeout elapsed)')
  }
  return false
}

function phiRecordSuccess() { phiFailures = 0; phiBlockedUntil = 0 }

function phiRecordFailure() {
  phiFailures++
  if (phiFailures >= 3) {
    phiBlockedUntil = Date.now() + 30000
    console.error(`[phi4] Circuit breaker OPEN for 30s after ${phiFailures} failures`)
  }
}

// ─── Utilities ────────────────────────────────────────────────────────────────
async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    const response = await fetch(url, { ...options, signal: controller.signal })
    return response
  } catch (err) {
    if (err.name === 'AbortError') throw new Error(`Request timed out after ${timeoutMs}ms: ${url}`)
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
        console.error(`[timeout] Request to ${req.path} exceeded ${timeoutMs}ms`)
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

function classifyExtension(fileName) {
  const ext = ('.' + fileName.split('.').pop()).toLowerCase()
  if (['.xlsx', '.xls', '.ods'].includes(ext))                               return DOC_TYPE.SPREADSHEET
  if (ext === '.pdf')                                                         return DOC_TYPE.PDF
  if (['.docx', '.doc', '.odt', '.rtf'].includes(ext))                       return DOC_TYPE.WORD
  if (['.pptx', '.ppt'].includes(ext))                                       return DOC_TYPE.PRESENTATION
  if (['.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h','.cs',
       '.go','.rb','.php','.swift','.kt','.r','.sql','.sh','.bash','.ps1']
      .includes(ext))                                                         return DOC_TYPE.CODE
  if (['.json','.jsonl','.yaml','.yml','.toml','.csv','.tsv'].includes(ext)) return DOC_TYPE.DATA
  if (['.txt','.md','.markdown','.rst'].includes(ext))                       return DOC_TYPE.TEXT
  if (ext === '.eml')                                                         return DOC_TYPE.EMAIL
  if (['.html','.htm','.xml'].includes(ext))                                 return DOC_TYPE.WEB
  return DOC_TYPE.UNKNOWN
}

function detectQueryIntent(query) {
  const q = query.toLowerCase().trim()
  const DEFINITION_PATTERNS  = [/^what\s+(is|are|does)\s+/,/^define\s+/,/^explain\s+/,/^meaning\s+of\s+/,/^tell\s+me\s+about\s+/,/^describe\s+/,/^how\s+is\s+.+\s+(calculated|defined|measured|computed)/,/\bmeaning\b/,/\bdefinition\b/,/\bwhat\s+does\b/]
  const CALCULATION_PATTERNS = [/how\s+(is|are|was|were)\s+.+\s+(calculated|computed|determined|derived)/,/what\s+is\s+the\s+(formula|calculation|computation)\s+for/,/how\s+do\s+you\s+(calculate|compute|determine)/]
  const LOOKUP_PATTERNS      = [/^(show|list|find|get|fetch|give)\s+(me\s+)?/,/^how\s+many\s+/,/^what\s+(is\s+the\s+)?(value|number|count|total|sum|amount)/]
  const COMPARISON_PATTERNS  = [/\bvs\b|\bversus\b|\bdifference\b|\bcompare\b|\bbetween\b/]
  const URL_PATTERNS         = [/\burl\b/,/\blink\b/,/\breport\b.*\burl\b/,/\burl\b.*\breport\b/,/\bpower\s*bi\b/,/\bdashboard\b/]
  const GREETING_PATTERNS    = [/^(hi|hello|hey|howdy|greetings|good\s+(morning|afternoon|evening)|how\s+are\s+you|what's\s+up|sup)\b/]
  if (GREETING_PATTERNS.some(p => p.test(q)))    return 'greeting'
  if (URL_PATTERNS.some(p => p.test(q)))         return 'url_lookup'
  if (CALCULATION_PATTERNS.some(p => p.test(q))) return 'calculation'
  if (DEFINITION_PATTERNS.some(p => p.test(q)))  return 'definition'
  if (COMPARISON_PATTERNS.some(p => p.test(q)))  return 'comparison'
  if (LOOKUP_PATTERNS.some(p => p.test(q)))      return 'lookup'
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
  ]
  for (const p of patterns) {
    const m = q.match(p)
    if (m) return m[1].trim()
  }
  return q
}

function extractUrlKeywords(query) {
  const q = query.toLowerCase()
  const stopWords = new Set(['power', 'bi', 'report', 'url', 'link', 'for', 'the', 'a', 'an', 'of', 'in', 'get', 'me', 'show', 'give', 'find', 'fetch'])
  const words = q.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1 && !stopWords.has(w))
  return words
}

function fixBrokenUrls(text) {
  return text.replace(/https:\/\/[^\s]+(\s+[^\s]+)/g, (match) => match.replace(/\s/g, ''))
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function cosineSim(a, b) {
  let dot = 0, normA = 0, normB = 0
  for (let i = 0; i < a.length; i++) {
    dot   += a[i] * b[i]
    normA += a[i] * a[i]
    normB += b[i] * b[i]
  }
  return dot / (Math.sqrt(normA) * Math.sqrt(normB) + 1e-9)
}

function buildInvertedIndex(chunks) {
  const index = new Map()
  for (let i = 0; i < chunks.length; i++) {
    const words = (chunks[i].text || '').toLowerCase()
      .replace(/[^\w\s]/g, ' ')
      .split(/\s+/)
    for (const w of words) {
      if (w.length < 2) continue
      if (!index.has(w)) index.set(w, new Set())
      index.get(w).add(i)
    }
  }
  return index
}

// ─── OPTIMIZATION 2: Query rewriting for multi-word and synonym expansion ─────
function generateQueryVariants(query) {
  const q = query.toLowerCase().trim()
  const subject = extractSubject(q)
  const variants = new Set([q, subject])

  // Add common abbreviations / synonyms relevant to real estate analytics
  const synonymMap = {
    'net applications': ['net apps', 'net application count', 'applications net'],
    'gross applications': ['gross apps', 'total applications'],
    'effective rent': ['eff rent', 'net rent', 'effective rental rate'],
    'competitor asking rent': ['comp asking rent', 'competitor rent', 'competitive asking rent'],
    'occupancy': ['occupancy rate', 'occ rate', 'occupancy status', 'occupied'],
    'psf': ['per square foot', 'per sq ft'],
    'lease application': ['application lease', 'leasing application'],
    'sort order': ['sort sequence', 'ordering', 'gl sort'],
    'gl activity': ['gl account', 'general ledger activity', 'gl account code'],
  }

  for (const [key, synonyms] of Object.entries(synonymMap)) {
    if (subject.includes(key) || q.includes(key)) {
      synonyms.forEach(s => variants.add(s))
    }
  }

  // Also add subject without stop words
  const stopWords = new Set(['the', 'a', 'an', 'is', 'are', 'how', 'what', 'does', 'do', 'calculated', 'defined', 'determine'])
  const strippedSubject = subject.split(/\s+/).filter(w => !stopWords.has(w)).join(' ')
  if (strippedSubject && strippedSubject !== subject) variants.add(strippedSubject)

  return [...variants].filter(v => v.length > 1)
}

// ─── Phi-4 call ───────────────────────────────────────────────────────────────
async function callPhi4(systemPrompt, userMessage, maxTokens = 768) {
  if (!PHI4_ENDPOINT || !PHI4_API_KEY) throw new Error('PHI4_ENDPOINT and PHI4_API_KEY environment variables are required')
  if (phiCircuitOpen()) throw new Error('Model temporarily unavailable (circuit breaker open)')

  return runWithPhiLimit(async () => {
    try {
      const response = await fetchWithTimeout(
        PHI4_ENDPOINT,
        {
          method:  'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${PHI4_API_KEY}` },
          body: JSON.stringify({
            model:       PHI4_MODEL,
            messages:    [{ role: 'system', content: systemPrompt }, { role: 'user', content: userMessage }],
            temperature: 0.1,
            max_tokens:  maxTokens,
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

// ─── Embeddings ───────────────────────────────────────────────────────────────
async function embedQueryAzure(query) {
  if (!AZURE_EMBED_ENDPOINT || !AZURE_EMBED_KEY) return null
  try {
    const response = await fetchWithTimeout(
      AZURE_EMBED_ENDPOINT,
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': AZURE_EMBED_KEY },
        body:    JSON.stringify({ input: query, model: AZURE_EMBED_MODEL }),
      },
      EMBED_TIMEOUT_MS
    )
    if (!response.ok) return null
    const data = await response.json()
    return data.data?.[0]?.embedding || null
  } catch (err) {
    console.warn('[embedQueryAzure] failed:', err.message)
    return null
  }
}

async function embedBatch(texts) {
  if (!AZURE_EMBED_ENDPOINT || !AZURE_EMBED_KEY || !texts.length) return []
  try {
    const response = await fetchWithTimeout(
      AZURE_EMBED_ENDPOINT,
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': AZURE_EMBED_KEY },
        body:    JSON.stringify({ input: texts, model: AZURE_EMBED_MODEL }),
      },
      EMBED_TIMEOUT_MS
    )
    if (!response.ok) return []
    const data = await response.json()
    return (data.data || []).sort((a, b) => a.index - b.index).map(d => d.embedding)
  } catch (err) {
    console.warn('[embedBatch] failed:', err.message)
    return []
  }
}

// ─── OPTIMIZATION 3: Improved keyword search with multi-word phrase scoring ───
function keywordSearch(query, chunks, topK, intent = 'general', invertedIndex = null) {
  const subject      = (intent === 'definition' || intent === 'calculation') ? extractSubject(query) : query.toLowerCase()
  const queryLower   = query.toLowerCase()
  const subjectLower = subject.toLowerCase()
  const subjectWords = subjectLower.split(/\s+/).filter(w => w.length > 1)
  const queryWords   = queryLower.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1)

  const urlKeywords = intent === 'url_lookup' ? extractUrlKeywords(query) : []

  // FIXED: for multi-word subjects, prefer phrase match over word-boundary-only match
  const isMultiWord       = subjectWords.length > 1
  const subjectPhraseRegex = isMultiWord
    ? new RegExp(escapeRegex(subjectLower), 'i')
    : new RegExp(`\\b${escapeRegex(subjectLower)}\\b`, 'i')

  // Gather candidate chunk indices using inverted index
  let candidateIndices
  if (invertedIndex && subjectWords.length > 0) {
    const wordsToIndex = intent === 'url_lookup' ? urlKeywords : subjectWords
    const sets = wordsToIndex.map(w => invertedIndex.get(w) || new Set())
    if (intent === 'url_lookup') {
      const urlSet  = invertedIndex.get('url')  || new Set()
      const linkSet = invertedIndex.get('link') || new Set()
      const httpSet = invertedIndex.get('http') || new Set()
      sets.push(urlSet, linkSet, httpSet)
    }
    const union = new Set()
    for (const s of sets) for (const idx of s) union.add(idx)
    candidateIndices = union
  }

  const source = candidateIndices
    ? [...candidateIndices].map(i => chunks[i]).filter(Boolean)
    : chunks.slice(0, 200)

  return source
    .map(c => {
      const text    = (c.text || '').toLowerCase()
      const docType = classifyExtension(c.source_file || '')
      let score = 0

      if (intent === 'url_lookup') {
        if (!text.includes('http')) return { ...c, _score: 0 }
        const topicMatches = urlKeywords.filter(w => text.includes(w)).length
        score += topicMatches * 10
        if (topicMatches === 0) return { ...c, _score: 0 }
        const topicPhrase = urlKeywords.join(' ')
        if (text.includes(topicPhrase)) score += 15
      } else {
        // OPTIMIZATION: phrase match scores highest — prevents false negatives on multi-word terms
        const phraseFound = subjectPhraseRegex.test(c.text || '')
        if (phraseFound) {
          score += subjectWords.length * 5  // boosted from 4
          // Extra boost if it appears in a "Measure Name" or "Attribute Name" position
          const measurePattern = new RegExp(`\\|\\s*${escapeRegex(subjectLower)}\\s*\\|`, 'i')
          if (measurePattern.test(c.text || '')) score += subjectWords.length * 4

          // Boost for definition/formula lines
          const defPattern = new RegExp(`${escapeRegex(subjectLower)}\\s*(is|are)\\s*(defined|described|calculated|measured|computed|the|a)`, 'i')
          if (defPattern.test(c.text || '')) score += subjectWords.length * 6
        }

        // Individual word coverage scoring
        const wordCoverage = subjectWords.filter(w => {
          const wPattern = new RegExp(`\\b${escapeRegex(w)}\\b`, 'i')
          return wPattern.test(c.text || '')
        }).length
        score += wordCoverage * 2

        // Spreadsheet-specific scoring
        if (docType === DOC_TYPE.SPREADSHEET || docType === DOC_TYPE.DATA) {
          if (phraseFound) score += subjectWords.length * 5  // spreadsheet bonus
          // "is described as" or "Formula:" pattern
          const descPattern = new RegExp(`${escapeRegex(subjectLower)}\\s*(is described as|formula):`, 'i')
          if (descPattern.test(c.text || '')) score += subjectWords.length * 8

          for (const w of subjectWords) {
            const kvPattern = new RegExp(`:\\s*${escapeRegex(w)}\\b|\\|\\s*${escapeRegex(w)}\\b`, 'i')
            if (kvPattern.test(c.text || '')) score += 2
          }
        }

        if (text.includes(queryLower)) score += queryWords.length * 2
      }

      return { ...c, _score: score }
    })
    .filter(c => c._score > 0)
    .sort((a, b) => b._score - a._score)
    .slice(0, topK)
}

// ─── OPTIMIZATION 4: Multi-query retrieval with automatic fallback expansion ──
async function retrieveChunks(query, chunks, topK = 6, invertedIndex = null) {
  const intent          = detectQueryIntent(query)
  const normalizedQuery = query.toLowerCase().trim().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ')
  const keywordTopK     = Math.min(200, chunks.length)

  // Primary retrieval
  let candidates = keywordSearch(normalizedQuery, chunks, keywordTopK, intent, invertedIndex)

  // OPTIMIZATION: if top score is weak, try query variants before giving up
  const topScore = candidates[0]?._score || 0
  if (topScore < KEYWORD_SHORTCIRCUIT_SCORE) {
    const variants = generateQueryVariants(query)
    console.log(`[retrieveChunks] Weak primary score (${topScore}), trying ${variants.length} query variants`)

    const variantResults = []
    for (const variant of variants) {
      if (variant === normalizedQuery) continue
      const variantNorm = variant.toLowerCase().trim()
      const variantHits = keywordSearch(variantNorm, chunks, Math.min(50, chunks.length), intent, invertedIndex)
      variantResults.push(...variantHits)
    }

    // Merge and deduplicate by chunk fingerprint, keeping highest score
    const byFingerprint = new Map()
    for (const c of [...candidates, ...variantResults]) {
      const fp = (c.text || '').trim().slice(0, 60)
      const existing = byFingerprint.get(fp)
      if (!existing || c._score > existing._score) byFingerprint.set(fp, c)
    }
    candidates = [...byFingerprint.values()].sort((a, b) => b._score - a._score)
    console.log(`[retrieveChunks] After variant expansion: ${candidates.length} candidates, top score: ${candidates[0]?._score || 0}`)
  }

  const pool = candidates.length > 0 ? candidates : chunks.slice(0, 150)

  if (pool.length > 0 && pool[0]._score >= KEYWORD_SHORTCIRCUIT_SCORE) {
    console.log(`[retrieveChunks] keyword short-circuit (score=${pool[0]._score})`)
    return pool.slice(0, Math.min(topK, 10))
  }

  if ((intent === 'definition' || intent === 'calculation') && pool.length > 0 && pool[0]._score >= 3) {
    return pool.slice(0, Math.min(topK, 10))
  }

  if (intent === 'url_lookup' && pool.length > 0) {
    return pool.slice(0, Math.min(topK, 6))
  }

  // Semantic (vector) retrieval
  if (AZURE_EMBED_ENDPOINT && AZURE_EMBED_KEY) {
    try {
      const queryVec = await embedQueryAzure(normalizedQuery)
      if (queryVec) {
        const poolSlice  = pool.slice(0, EMBED_POOL_LIMIT)
        const chunkTexts = poolSlice.map(c => (c.text || '').toLowerCase().slice(0, 512))
        const embeddings = await embedBatch(chunkTexts)

        const maxKeyword = pool[0]._score || 1
        const weight     = (intent === 'definition' || intent === 'calculation')
          ? { semantic: 0.35, keyword: 0.65 }
          : { semantic: 0.70, keyword: 0.30 }

        const scored = poolSlice.map((c, i) => {
          const chunkVec = embeddings[i]
          if (!chunkVec) return c
          const semanticScore = cosineSim(queryVec, chunkVec)
          const keywordNorm   = typeof c._score === 'number' ? c._score / maxKeyword : 0
          return { ...c, _score: semanticScore * weight.semantic + keywordNorm * weight.keyword }
        })
        const remainder = pool.slice(EMBED_POOL_LIMIT).map(c => ({
          ...c,
          _score: (typeof c._score === 'number' ? c._score / maxKeyword : 0) * weight.keyword,
        }))

        return [...scored, ...remainder]
          .sort((a, b) => b._score - a._score)
          .slice(0, Math.min(topK, 12))
      }
    } catch (err) {
      console.warn('[retrieveChunks] Azure embed failed, keyword fallback:', err.message)
    }
  }

  return pool.slice(0, Math.min(topK, 12))
}

// ─── OPTIMIZATION 5: Wider context window for answer generation ───────────────
function buildContext(hits) {
  const seen    = new Set()
  const deduped = []
  for (const h of hits) {
    const fingerprint = (h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fingerprint)) {
      seen.add(fingerprint)
      deduped.push(h)
    }
    if (deduped.length >= 8) break  // increased from 6 — more context for 14B model
  }

  return deduped.map((h, i) => {
    // Larger windows — Phi-4 14B can handle it
    const charLimit = i === 0 ? 900 : i <= 2 ? 700 : 500
    const text = (h.text || '').trim().slice(0, charLimit)
    return `[${i + 1}] ${text}`
  }).join('\n\n')
}

const SYSTEM_PROMPT_CACHE     = new Map()
const SYSTEM_PROMPT_CACHE_MAX = 100

// ─── OPTIMIZATION 6: Improved intent strategy for calculation queries ──────────
const INTENT_STRATEGY = {
  definition:  `TASK: Define the exact term asked. Write 2–4 natural English sentences. State what it measures or represents, then how it is calculated if a formula is present. Combine multiple excerpts if needed. Never copy raw pipe-delimited data.`,
  calculation: `TASK: Explain exactly how the measure is calculated. State the formula clearly in plain English (e.g. "calculated by dividing X by Y"). Include all steps if the formula has multiple components. If multiple related formulas exist, describe each. Never say "not found" if a formula-like pattern exists in the context.`,
  lookup:      `TASK: Find and report the exact value or list asked for. State clearly what it represents.`,
  comparison:  `TASK: Compare the items asked. Write a structured but natural summary — similarities and differences.`,
  url_lookup:  `TASK: Return the full URL that matches the topic, on its own line, unmodified. If none found, say so.`,
  general:     `TASK: Find information that directly answers the question and write a clear, complete summary.`,
}

const TYPE_NOTES = {
  [DOC_TYPE.SPREADSHEET]:  `Spreadsheet data is serialized with headers like "Measure Name | Description | Formula". Read all columns — Description contains the definition, Formula contains the calculation. Never reproduce raw pipe-delimited rows verbatim; synthesize into natural prose.`,
  [DOC_TYPE.DATA]:         `JSON/YAML/CSV fields may be nested ("parent.child: value"). Read for meaning; never dump raw key:value pairs.`,
  [DOC_TYPE.PDF]:          `PDF text may have minor extraction artefacts. Read numbers and dates exactly as shown; ignore headers, footers, and page numbers.`,
  [DOC_TYPE.WORD]:         `Word docs contain prose, lists, and tables. Headings show section structure. Quote policy/definition statements accurately.`,
  [DOC_TYPE.PRESENTATION]: `Slide titles are section headers; bullets are detail. Do not infer beyond what the slide states.`,
  [DOC_TYPE.CODE]:         `Read code literally — function/variable names and comments all matter. Describe what code does in plain English unless code output is requested.`,
  [DOC_TYPE.TEXT]:         `Markdown formatting (##, **, -) indicates structure. Lists represent discrete facts or steps.`,
  [DOC_TYPE.EMAIL]:        `Attribute statements to their sender. Dates and times are as stated in the email header.`,
  [DOC_TYPE.WEB]:          `Focus on main body content; ignore repetitive navigation text. Include URLs exactly as written.`,
}

const BASE_SYSTEM_PROMPT = `You are a document assistant for a real estate analytics platform. Answer ONLY from the provided context.

RULES:
- Write in natural English sentences. Never output raw pipe-delimited rows or dump key:value pairs.
- Use only information present in the context. Never invent or assume.
- Match terms case-insensitively.
- No citation markers ([1], [2]…). No padding or caveats.
- If a URL appears, return it exactly as-is on its own line.
- If a formula or calculation is in the context, state it clearly in plain English.
- Only say "I couldn't find that in your documents" if there is genuinely NO relevant information in the context after reading all excerpts carefully.
- Prefer a partial answer with what IS known over refusing to answer.
- Keep answers focused — 2–5 sentences for definitions, more if a formula has multiple steps.`

function buildDynamicSystemPrompt(hits, intent = 'general') {
  const types     = [...new Set(hits.map(h => classifyExtension(h.source_file || '')))]
  const typeNotes = types.map(t => TYPE_NOTES[t]).filter(Boolean)

  const cacheKey = `${intent}::${types.sort().join(',')}`
  const cached   = SYSTEM_PROMPT_CACHE.get(cacheKey)
  if (cached) return cached

  const parts = [
    BASE_SYSTEM_PROMPT,
    INTENT_STRATEGY[intent] || INTENT_STRATEGY.general,
  ]
  if (typeNotes.length) parts.push('SOURCE NOTES:\n' + typeNotes.map((n, i) => `${i + 1}. ${n}`).join('\n'))
  if (types.length > 1) parts.push(`Context spans ${types.length} source types — apply the relevant note per excerpt.`)

  const prompt = parts.join('\n\n')

  if (SYSTEM_PROMPT_CACHE.size >= SYSTEM_PROMPT_CACHE_MAX)
    SYSTEM_PROMPT_CACHE.delete(SYSTEM_PROMPT_CACHE.keys().next().value)
  SYSTEM_PROMPT_CACHE.set(cacheKey, prompt)
  return prompt
}

// ─── MongoDB helpers ──────────────────────────────────────────────────────────
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
  const uri    = CHAT_HISTORY_URI || MONGODB_URI
  const client = new MongoClient(uri)
  await client.connect()
  chatDb = client.db(CHAT_HISTORY_DB)
  return chatDb
}

// ─── Client cache / auth ──────────────────────────────────────────────────────
const CLIENT_CACHE = new Map()
const CACHE_TTL_MS = 5 * 60 * 1000

function getCached(apiKey) {
  const entry = CLIENT_CACHE.get(apiKey)
  if (!entry) return null
  if (Date.now() - entry.cachedAt > CACHE_TTL_MS) { CLIENT_CACHE.delete(apiKey); return null }
  return entry
}
function setCache(apiKey, data) { CLIENT_CACHE.set(apiKey, { ...data, cachedAt: Date.now() }) }
function evictCache(apiKey)     { if (apiKey) CLIENT_CACHE.delete(apiKey) }

async function verifyApiKey(apiKey) {
  if (!apiKey || !apiKey.startsWith('rak_')) return null
  const cached = getCached(apiKey)
  if (cached) return { clientId: cached.clientId, name: cached.name }
  const database = await getDb()
  const client   = await database.collection('clients').findOne({ apiKey }, { projection: { clientId: 1, name: 1, _id: 0 } })
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
      const database  = await getDb()
      const validDocs = await database.collection('clients').find({ apiKey: { $in: keys } }, { projection: { apiKey: 1, _id: 0 } }).toArray()
      const validSet  = new Set(validDocs.map(d => d.apiKey))
      for (const key of keys) if (!validSet.has(key)) evictCache(key)
    } catch {}
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

// ─── Document extractors ──────────────────────────────────────────────────────
async function extractPdf(buffer)  { const r = await pdfParse(buffer); return r.text || '' }
async function extractWord(buffer) { const r = await mammoth.extractRawText({ buffer }); return r.value || '' }

// ─── OPTIMIZATION 7: Completely rewritten extractSpreadsheet ─────────────────
//
// Key improvements over original:
//  1. Skips the copyright/header row (row 0) properly
//  2. Handles blank column headers (previously labeled "Field1", "Field2"…)
//     — now falls back to the nearest non-blank header to the left, which
//     is what Excel merged-cell data usually represents.
//  3. Each data row generates ONE rich semantic chunk line that contains:
//     - Table (sheet name)
//     - All field:value pairs with proper headers
//     - A natural-language "X is defined as / formula:" synthesis line
//  4. Generates a sheet-level summary at the top of each sheet block
//     so that the sheet name itself is searchable context.
//  5. Formula column is explicitly labelled and included in the synthesis
//     so that "how is X calculated" queries can find the formula.
//  6. Preserves multi-line cell content (newlines → spaces) without truncation.
//
function extractSpreadsheet(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellNF: true, cellFormula: true })
  const parts    = []

  for (const sheetName of workbook.SheetNames) {
    const sheet   = workbook.Sheets[sheetName]
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1, raw: false })
    if (!rawRows.length) continue

    // Find the actual header row (first row where > 2 cells have content)
    let headerRowIdx = 0
    for (let i = 0; i < Math.min(5, rawRows.length); i++) {
      const nonEmpty = rawRows[i].filter(cell => String(cell).trim() !== '').length
      if (nonEmpty >= 2) { headerRowIdx = i; break }
    }

    // Build header array, resolving blank cells by carrying forward the last non-blank
    const rawHeaders = rawRows[headerRowIdx].map(h => String(h).trim())
    const headers    = []
    let lastNonBlank = ''
    for (const h of rawHeaders) {
      if (h !== '') { lastNonBlank = h; headers.push(h) }
      else          { headers.push(lastNonBlank || `Col${headers.length + 1}`) }
    }

    // Identify key column indices by name pattern
    const colIdx = {}
    headers.forEach((h, i) => {
      const hLower = h.toLowerCase()
      if (/table|module|category/.test(hLower) && !colIdx.table)       colIdx.table       = i
      if (/measure\s*name|attribute\s*name|name/.test(hLower) && !colIdx.name) colIdx.name = i
      if (/description|desc/.test(hLower) && !colIdx.description)      colIdx.description = i
      if (/formula|calculation|calc/.test(hLower) && !colIdx.formula)  colIdx.formula     = i
    })

    // Sheet-level summary header — makes the sheet name itself searchable
    parts.push(`=== Sheet: ${sheetName} ===`)
    parts.push(`This sheet contains ${rawRows.length - headerRowIdx - 1} entries from the "${sheetName}" data dictionary.`)
    if (colIdx.name !== undefined)        parts.push(`Columns include: ${headers.filter(h => h).join(', ')}.`)
    parts.push('')

    // Process each data row
    for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
      const row = rawRows[i]

      // Skip completely blank rows
      if (!row.some(cell => String(cell).trim() !== '')) continue

      // Normalize each cell value (collapse internal newlines to spaces)
      const cells = row.map(cell => String(cell || '').replace(/\r?\n/g, ' ').trim())

      // Build key-value pairs using the resolved headers
      const pairs = []
      for (let j = 0; j < Math.max(headers.length, cells.length); j++) {
        const val = cells[j] || ''
        if (!val) continue
        const key = headers[j] || `Col${j + 1}`
        pairs.push({ key, val })
      }
      if (!pairs.length) continue

      // ── Line 1: pipe-delimited record (for keyword search) ──────────────
      parts.push(pairs.map(p => `${p.key}: ${p.val}`).join(' | '))

      // ── Line 2: Natural-language synthesis ────────────────────────────────
      const nameVal  = colIdx.name        !== undefined ? cells[colIdx.name]        : ''
      const tableVal = colIdx.table       !== undefined ? cells[colIdx.table]       : sheetName
      const descVal  = colIdx.description !== undefined ? cells[colIdx.description] : ''
      const formulaVal = colIdx.formula   !== undefined ? cells[colIdx.formula]     : ''

      if (nameVal && descVal) {
        let synthesis = `${nameVal} (from ${tableVal}) is defined as: ${descVal}`
        if (formulaVal) synthesis += ` Formula: ${formulaVal}`
        parts.push(synthesis)
      } else if (nameVal && formulaVal) {
        parts.push(`${nameVal} formula: ${formulaVal}`)
      }

      parts.push('')
    }
  }

  return parts.join('\n')
}

function extractCsv(buffer, delimiter = ',') {
  const text   = buffer.toString('utf-8')
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

function extractXml(buffer)  { return buffer.toString('utf-8').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() }
function extractJson(buffer) { try { return JSON.stringify(JSON.parse(buffer.toString('utf-8')), null, 2) } catch { return buffer.toString('utf-8') } }
function extractJsonl(buffer) {
  return buffer.toString('utf-8').split('\n').filter(Boolean)
    .map(line => { try { return JSON.stringify(JSON.parse(line)) } catch { return line } })
    .join('\n')
}
function extractYaml(buffer) { try { return JSON.stringify(yaml.load(buffer.toString('utf-8')), null, 2) } catch { return buffer.toString('utf-8') } }

async function extractEml(buffer) {
  const parsed = await simpleParser(buffer)
  const parts  = []
  if (parsed.subject) parts.push(`Subject: ${parsed.subject}`)
  if (parsed.from)    parts.push(`From: ${parsed.from.text}`)
  if (parsed.to)      parts.push(`To: ${parsed.to.text}`)
  if (parsed.date)    parts.push(`Date: ${parsed.date}`)
  if (parsed.text)    parts.push(`\n${parsed.text}`)
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
  if (ext === '.pdf')                           return extractPdf(buffer)
  if (ext === '.docx' || ext === '.doc')        return extractWord(buffer)
  if (ext === '.odt'  || ext === '.rtf')        return extractOffice(buffer)
  if (['.xlsx','.xls','.ods'].includes(ext))    return extractSpreadsheet(buffer)
  if (ext === '.csv')                           return extractCsv(buffer, ',')
  if (ext === '.tsv')                           return extractCsv(buffer, '\t')
  if (ext === '.pptx' || ext === '.ppt')        return extractOffice(buffer)
  if (ext === '.html' || ext === '.htm')        return extractHtml(buffer)
  if (ext === '.xml')                           return extractXml(buffer)
  if (['.md','.markdown','.rst'].includes(ext)) return buffer.toString('utf-8')
  if (ext === '.json')                          return extractJson(buffer)
  if (ext === '.jsonl')                         return extractJsonl(buffer)
  if (ext === '.yaml' || ext === '.yml')        return extractYaml(buffer)
  if (ext === '.toml')                          return buffer.toString('utf-8')
  const plainText = new Set(['.txt','.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h','.cs','.go','.rb','.php','.swift','.kt','.r','.sql','.sh','.bash','.ps1'])
  if (plainText.has(ext))                       return buffer.toString('utf-8')
  if (ext === '.epub')                          return extractEpub(buffer)
  if (ext === '.eml')                           return extractEml(buffer)
  return ''
}

// ─── OPTIMIZATION 8: Smarter chunking that keeps semantic units together ──────
//
// Original: splits purely on character count, breaking multi-line formulas.
// New: prefers splitting on blank lines (semantic boundaries in spreadsheet output)
// and only hard-splits if a single block exceeds CHUNK_SIZE.
//
function chunkText(text, sourceFile) {
  const chunks   = []
  let chunkIndex = 0

  // Split text into "blocks" separated by blank lines
  const blocks = text.replace(/\r\n/g, '\n').split(/\n{2,}/).map(b => b.trim()).filter(b => b.length > 0)

  let buffer        = []
  let bufferLength  = 0

  function flush() {
    const chunkStr = buffer.join('\n\n')
    if (chunkStr.length >= 30) {
      chunks.push({ text: chunkStr, source_file: sourceFile, chunk_index: chunkIndex++, embedding: [] })
    }
    buffer       = []
    bufferLength = 0
  }

  for (const block of blocks) {
    // If this single block is enormous, hard-split it on newlines
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
          // Overlap: keep last CHUNK_OVERLAP lines
          lineBuffer  = lineBuffer.slice(-CHUNK_OVERLAP)
          lineLength  = lineBuffer.join('\n').length
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
      flush()
      // Overlap: carry last block into next chunk
      buffer.push(blocks[blocks.indexOf(block) - 1] || '')
      bufferLength = buffer[0]?.length || 0
    }
    buffer.push(block)
    bufferLength += (bufferLength ? 2 : 0) + block.length
  }

  if (buffer.length > 0) flush()
  return chunks
}

// ─── Blob / chunk loading ─────────────────────────────────────────────────────
async function downloadBlobAsBuffer(containerClient, blobName) {
  const download = await containerClient.getBlobClient(blobName).download()
  const parts    = []
  for await (const chunk of download.readableStreamBody)
    parts.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
  return Buffer.concat(parts)
}

const BLOB_CONCURRENCY = parseInt(process.env.BLOB_CONCURRENCY || '8', 10)

async function _doLoadChunks(clientId) {
  if (!blobServiceClient) throw new Error('AZURE_CONNECTION_STRING not set')
  const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
  const prefix          = `${RAW_PREFIX}/${clientId}/`

  const blobNames = []
  for await (const blob of containerClient.listBlobsFlat({ prefix })) {
    const fileName = blob.name.split('/').pop()
    const ext      = ('.' + fileName.split('.').pop()).toLowerCase()
    if (SUPPORTED_EXTENSIONS.has(ext)) blobNames.push(blob.name)
  }

  const allChunks = []
  for (let i = 0; i < blobNames.length; i += BLOB_CONCURRENCY) {
    const batch   = blobNames.slice(i, i + BLOB_CONCURRENCY)
    const results = await Promise.allSettled(
      batch.map(async (blobName) => {
        const fileName = blobName.split('/').pop()
        const buffer   = await downloadBlobAsBuffer(containerClient, blobName)
        const text     = await extractTextFromBuffer(buffer, fileName)
        if (!text?.trim()) return []
        return chunkText(text, fileName)
      })
    )
    for (const result of results) {
      if (result.status === 'fulfilled') allChunks.push(...result.value)
      else console.warn(`[loadChunks] blob failed:`, result.reason?.message)
    }
  }
  return allChunks
}

const CHUNK_CACHE     = new Map()
const CHUNK_CACHE_TTL = parseInt(process.env.CHUNK_CACHE_TTL_MS || '300000', 10)

async function loadChunksForClient(clientId) {
  const now    = Date.now()
  const cached = CHUNK_CACHE.get(clientId)

  if (cached && cached.chunks) {
    const isStale = now - cached.ts > CHUNK_CACHE_TTL
    if (!isStale) {
      return { chunks: cached.chunks, invertedIndex: cached.invertedIndex }
    }
    if (!cached.loading) {
      console.log(`[chunkCache] Serving stale cache for ${clientId}, refreshing in background`)
      const refreshPromise = _doLoadChunks(clientId)
        .then(chunks => {
          const invertedIndex = buildInvertedIndex(chunks)
          CHUNK_CACHE.set(clientId, { chunks, invertedIndex, ts: Date.now(), loading: null })
          console.log(`[chunkCache] Background refresh done for ${clientId}: ${chunks.length} chunks`)
          return chunks
        })
        .catch(err => {
          const existing = CHUNK_CACHE.get(clientId)
          CHUNK_CACHE.set(clientId, { ...existing, loading: null })
          console.warn(`[chunkCache] Background refresh failed for ${clientId}: ${err.message}`)
        })
      CHUNK_CACHE.set(clientId, { ...cached, loading: refreshPromise })
      return { chunks: cached.chunks, invertedIndex: cached.invertedIndex }
    }
    return { chunks: cached.chunks, invertedIndex: cached.invertedIndex }
  }

  if (cached && cached.loading) {
    const chunks = await cached.loading
    const entry  = CHUNK_CACHE.get(clientId)
    return { chunks: chunks || entry?.chunks || [], invertedIndex: entry?.invertedIndex || null }
  }

  const loadPromise = _doLoadChunks(clientId)
    .then(chunks => {
      const invertedIndex = buildInvertedIndex(chunks)
      CHUNK_CACHE.set(clientId, { chunks, invertedIndex, ts: Date.now(), loading: null })
      console.log(`[chunkCache] Loaded ${chunks.length} chunks + built inverted index for ${clientId}`)
      return chunks
    })
    .catch(err => {
      CHUNK_CACHE.set(clientId, { chunks: null, invertedIndex: null, ts: 0, loading: null })
      throw err
    })

  CHUNK_CACHE.set(clientId, { chunks: null, invertedIndex: null, ts: 0, loading: loadPromise })
  const chunks = await loadPromise
  const entry  = CHUNK_CACHE.get(clientId)
  return { chunks, invertedIndex: entry?.invertedIndex || null }
}

function invalidateChunkCache(clientId) {
  CHUNK_CACHE.delete(clientId)
  console.log(`[chunkCache] Invalidated cache for client: ${clientId}`)
}

function warmupChunkCaches() {
  if (!WARMUP_CLIENT_IDS.length || !blobServiceClient) return
  console.log(`[warmup] Pre-loading chunks for ${WARMUP_CLIENT_IDS.length} client(s): ${WARMUP_CLIENT_IDS.join(', ')}`)
  for (const clientId of WARMUP_CLIENT_IDS) {
    loadChunksForClient(clientId)
      .then(({ chunks }) => console.log(`[warmup] ${clientId} — ${chunks.length} chunks ready`))
      .catch(err => console.warn(`[warmup] ${clientId} — ${err.message}`))
  }
}

// ─── Answer generation ────────────────────────────────────────────────────────
async function answerWithPhi4(originalQuery, hits, intent = 'general') {
  const systemPrompt = buildDynamicSystemPrompt(hits, intent)
  const context      = buildContext(hits)

  let taskHint = ''
  if (intent === 'definition') {
    taskHint = `\nYou must define ONLY this exact term: "${extractSubject(originalQuery)}". Write 2–4 clean sentences: what it measures, and how it is calculated if a formula exists. Do NOT copy raw data rows.`
  } else if (intent === 'calculation') {
    taskHint = `\nExplain exactly how "${extractSubject(originalQuery)}" is calculated. State the formula clearly. If multiple formula steps exist, describe each step. Never say "not found" if formula-like text exists.`
  } else if (intent === 'url_lookup') {
    taskHint = `\nReturn the full URL for: "${extractUrlKeywords(originalQuery).join(' ')}". Put it on its own line. No disclaimers.`
  }

  const userMessage = `CONTEXT:\n${context}${taskHint}\n\nQuestion: ${originalQuery}`
  // Use more tokens for calculation/definition since formulas can be long
  const maxTokens = (intent === 'calculation' || intent === 'definition') ? 1024 : 768
  return callPhi4(systemPrompt, userMessage, maxTokens)
}

// ─── OPTIMIZATION 9: Improved fallback answer — handles blank Phi-4 responses ─
function buildFallbackAnswer(query, hits) {
  if (!hits || hits.length === 0) {
    return "I couldn't find relevant information in your documents for this query."
  }
  const intent  = detectQueryIntent(query)
  const subject = extractSubject(query).toLowerCase()

  if (intent === 'url_lookup') {
    const urlKeywords = extractUrlKeywords(query)
    for (const h of hits) {
      const lines = (h.text || '').split('\n')
      for (const line of lines) {
        if (!line.toLowerCase().includes('http')) continue
        const lineLower = line.toLowerCase()
        const matches = urlKeywords.filter(w => lineLower.includes(w)).length
        if (matches > 0) {
          const urlMatch = line.match(/https?:\/\/\S+/)
          if (urlMatch) return urlMatch[0].replace(/\s/g, '')
        }
      }
    }
  }

  // Look for the natural-language synthesis lines generated by extractSpreadsheet
  const synthesisPattern = new RegExp(`${escapeRegex(subject)}\\s*\\(from .+?\\)\\s*is defined as:\\s*(.+?)(?:\\s+Formula:\\s*(.+))?$`, 'im')
  for (const h of hits) {
    const m = (h.text || '').match(synthesisPattern)
    if (m) {
      const desc    = m[1]?.trim().slice(0, 400) || ''
      const formula = m[2]?.trim().slice(0, 300) || ''
      if (desc && formula) return `${subject.charAt(0).toUpperCase() + subject.slice(1)} is ${desc} It is calculated as: ${formula}`
      if (desc)            return `${subject.charAt(0).toUpperCase() + subject.slice(1)} is ${desc}`
    }
  }

  // Look for "Formula:" lines
  if (intent === 'calculation') {
    for (const h of hits) {
      const lines = (h.text || '').split('\n')
      for (const line of lines) {
        const ll = line.toLowerCase()
        if (ll.includes(subject) && /formula:/i.test(line)) {
          const formulaPart = line.replace(/^.*formula:\s*/i, '').trim()
          if (formulaPart.length > 5) return `${subject.charAt(0).toUpperCase() + subject.slice(1)} is calculated as: ${formulaPart.slice(0, 400)}`
        }
      }
    }
  }

  // Original "is described as" fallback
  for (const h of hits) {
    const lines = (h.text || '').split('\n')
    for (const line of lines) {
      if (line.toLowerCase().includes(subject) && /is described as:/i.test(line)) {
        const parts = line.split(/is described as:/i)
        if (parts[1]) {
          const raw       = parts[1].trim().slice(0, 300)
          const descParts = raw.split('|').map(p => p.trim()).filter(Boolean)
          if (descParts.length === 1) return `${subject.charAt(0).toUpperCase() + subject.slice(1)} is ${descParts[0]}.`
          return `${subject.charAt(0).toUpperCase() + subject.slice(1)}: ${descParts.join('. ')}`
        }
      }
    }
  }

  // Generic matching lines
  const matchingLines = []
  for (const h of hits) {
    const lines = (h.text || '').split('\n')
    for (const line of lines) {
      if (line.toLowerCase().includes(subject) && line.trim().length > 20) {
        if ((line.match(/\|/g) || []).length > 4) continue  // skip over-pipe-y lines
        if (/is described as:/i.test(line)) continue
        matchingLines.push(line.trim())
      }
    }
  }

  if (matchingLines.length > 0) {
    return [...new Set(matchingLines)].slice(0, 3).join(' ').slice(0, 400)
  }

  return "I couldn't find that specific information in your documents."
}

// ─── OPTIMIZATION 10: Stricter answer cleanup — don't strip valid content ─────
function cleanAnswer(rawAnswer) {
  return fixBrokenUrls(rawAnswer)
    // Remove raw "FieldN:" artifacts from old chunker
    .replace(/\bField\d+\s*:\s*/gi, '')
    .replace(/\|\s*Field\d+\b/gi, '')
    // Remove raw "X is described as: pipe | pipe | pipe" dump lines
    // ONLY if the line is purely a data dump (4+ pipe segments, no natural prose)
    .replace(/^[^\n]*\|[^\n]*\|[^\n]*\|[^\n]*\|[^\n]*\|[^\n]*$/gm, '')
    // Remove lines that are just "X is described as: …" verbatim copy (more than 3 pipe segments)
    .replace(/^.+\s+is described as:\s*.+\|.+\|.+$/gmi, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function generateTitle(query) {
  const cleaned = query.trim().replace(/[?!.]+$/, '')
  return cleaned.length > 50 ? cleaned.slice(0, 50) + '…' : cleaned
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/health', (req, res) => res.json({
  ok:                true,
  service:           'ask-data',
  model:             PHI4_MODEL,
  embeddings:        AZURE_EMBED_ENDPOINT ? 'azure-anurit' : 'keyword-only',
  chunkSize:         CHUNK_SIZE,
  chunkCacheSize:    CHUNK_CACHE.size,
  promptCacheSize:   SYSTEM_PROMPT_CACHE.size,
  responseCacheSize: RESPONSE_CACHE.size,
  phiCircuitOpen:    phiCircuitOpen(),
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
    const col      = database.collection('clients')
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
    const clients  = await database.collection('clients').find({}, { projection: { apiKey: 0 } }).sort({ createdAt: -1 }).toArray()
    res.json({ clients })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.get('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const database = await getDb()
    const client   = await database.collection('clients').findOne({ clientId: req.params.clientId })
    if (!client) return res.status(404).json({ error: 'Client not found' })
    res.json(client)
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/admin/clients/:clientId/regenerate-key', requireAdminKey, async (req, res) => {
  try {
    const database  = await getDb()
    const col       = database.collection('clients')
    const oldClient = await col.findOne({ clientId: req.params.clientId }, { projection: { apiKey: 1 } })
    if (!oldClient) return res.status(404).json({ error: 'Client not found' })
    const newApiKey = generateApiKey()
    const now       = new Date().toISOString()
    if (oldClient.apiKey) evictCache(oldClient.apiKey)
    await col.findOneAndUpdate({ clientId: req.params.clientId }, { $set: { apiKey: newApiKey, apiKeyRotatedAt: now, updatedAt: now } }, { returnDocument: 'after' })
    res.json({ success: true, clientId: req.params.clientId, newApiKey, apiKeyRotatedAt: now })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.patch('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const database = await getDb()
    const updates  = { ...req.body, updatedAt: new Date().toISOString() }
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
    const database     = await getDb()
    const client       = await database.collection('clients').findOne({ clientId })
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
  SYSTEM_PROMPT_CACHE.clear()
  RESPONSE_CACHE.clear()
  res.json({ ok: true, clientId: req.params.clientId, message: 'Chunk + prompt + response cache invalidated' })
})

app.post('/client/login', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ error: 'Invalid API key' })
    if (blobServiceClient) {
      loadChunksForClient(client.clientId).catch(err =>
        console.warn(`[login warmup] ${client.clientId}: ${err.message}`)
      )
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
      loadChunksForClient(client.clientId).catch(err =>
        console.warn(`[chat/login warmup] ${client.clientId}: ${err.message}`)
      )
    }
    res.json({ ok: true, client })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.get('/client/me', requireClientKey, async (req, res) => {
  try {
    const database = await getDb()
    const client   = await database.collection('clients').findOne({ clientId: req.client.clientId }, { projection: { apiKey: 0 } })
    if (!client) return res.status(404).json({ error: 'Client not found' })
    res.json(client)
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/conversations', requireClientKey, async (req, res) => {
  try {
    const { title }    = req.body
    const database     = await getChatDb()
    const now          = new Date()
    const conversation = { clientId: req.client.clientId, title: title || 'New Conversation', messages: [], createdAt: now, updatedAt: now }
    const result       = await database.collection('conversations').insertOne(conversation)
    res.status(201).json({ ...conversation, _id: result.insertedId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/conversations/list', requireClientKey, async (req, res) => {
  try {
    const database      = await getChatDb()
    const conversations = await database.collection('conversations').find({ clientId: req.client.clientId }, { projection: { messages: 0 } }).sort({ updatedAt: -1 }).toArray()
    res.json({ conversations })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/conversations/get', requireClientKey, async (req, res) => {
  try {
    const { conversationId } = req.body
    if (!conversationId) return res.status(400).json({ error: 'conversationId is required' })
    const database     = await getChatDb()
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
    const result   = await database.collection('conversations').findOneAndUpdate({ _id: new ObjectId(conversationId), clientId: req.client.clientId }, { $set: { title: title.trim(), updatedAt: new Date() } }, { returnDocument: 'after', projection: { messages: 0 } })
    if (!result) return res.status(404).json({ error: 'Conversation not found' })
    res.json(result)
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/conversations/delete', requireClientKey, async (req, res) => {
  try {
    const { conversationId } = req.body
    if (!conversationId) return res.status(400).json({ error: 'conversationId is required' })
    const database = await getChatDb()
    const result   = await database.collection('conversations').deleteOne({ _id: new ObjectId(conversationId), clientId: req.client.clientId })
    if (result.deletedCount === 0) return res.status(404).json({ error: 'Conversation not found' })
    res.json({ ok: true, deleted: conversationId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/message', requireClientKey, withRequestTimeout(async (req, res) => {
  try {
    const { query, topK = 8, conversationId } = req.body  // default topK raised to 8
    if (!query?.trim()) return res.status(400).json({ error: 'query is required' })
    const { clientId, name } = req.client

    const intent = detectQueryIntent(query.trim())
    if (intent === 'greeting') {
      return res.json({
        answer: "Hello! I'm your document assistant. Ask me anything about your data.",
        sources: [],
        conversationId: conversationId || null,
        client: { clientId, name },
      })
    }

    const cacheKey = getCacheKey(clientId, query)
    const cached   = responseCacheGet(cacheKey)
    if (cached) {
      console.log(`[cache] HIT for "${query.slice(0, 50)}"`)
      return res.json({ ...cached, cached: true, conversationId: conversationId || cached.conversationId })
    }

    if (IN_FLIGHT.has(cacheKey)) {
      console.log(`[dedup] Waiting for in-flight request: "${query.slice(0, 50)}"`)
      try {
        const result = await IN_FLIGHT.get(cacheKey)
        return res.json({ ...result, conversationId: conversationId || result.conversationId })
      } catch {}
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

      const hits = await retrieveChunks(query.trim(), chunks, Math.min(topK, 20), invertedIndex)
      console.log(`[chat/message] "${query.slice(0, 60)}" → intent=${intent}, hits=${hits.length}, topScore=${hits[0]?._score?.toFixed(2) || 0}`)

      if (hits.length === 0) {
        return {
          answer: "I couldn't find that in your documents. Try rephrasing your question.",
          sources: [],
          conversationId: conversationId || null,
          client: { clientId, name },
        }
      }

      let rawAnswer = ''
      try {
        rawAnswer = await Promise.race([
          answerWithPhi4(query.trim(), hits, intent),
          new Promise((_, reject) =>
            setTimeout(() => reject(new Error('Model response timeout (25s)')), 25000)
          ),
        ])
      } catch (err) {
        console.warn(`[phi4] Using fallback answer: ${err.message}`)
      }

      // OPTIMIZATION: if Phi-4 returns blank/very short, use fallback
      const isBlankAnswer = !rawAnswer || rawAnswer.trim().length < 10
      const answer = isBlankAnswer
        ? buildFallbackAnswer(query.trim(), hits)
        : cleanAnswer(rawAnswer)

      if (isBlankAnswer) {
        console.warn(`[phi4] Blank/short response detected, used fallback for: "${query.slice(0, 60)}"`)
      }

      const sources = hits.map(h => ({
        source_file: h.source_file  || 'unknown',
        chunk_index: h.chunk_index  ?? 0,
        score:       typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
        preview:     (h.text || '').slice(0, 300),
      }))

      let activeConversationId = conversationId || null
      try {
        const chatDatabase = await getChatDb()
        const col          = chatDatabase.collection('conversations')
        const now          = new Date()
        const userMsg      = { role: 'user',      content: query.trim(), timestamp: now }
        const assistantMsg = {
          role:      'assistant',
          content:   answer,
          sources:   sources.map(s => ({ source_file: s.source_file, score: s.score })),
          timestamp: now,
        }

        if (activeConversationId) {
          const updated = await col.findOneAndUpdate(
            { _id: new ObjectId(activeConversationId), clientId },
            { $push: { messages: { $each: [userMsg, assistantMsg] } }, $set: { updatedAt: now } },
            { returnDocument: 'after', projection: { _id: 1 } }
          )
          if (!updated) {
            console.warn(`[chat/message] conversationId ${activeConversationId} not found for ${clientId}, creating new`)
            activeConversationId = null
          }
        }

        if (!activeConversationId) {
          const title  = generateTitle(query.trim())
          const result = await col.insertOne({ clientId, title, messages: [userMsg, assistantMsg], createdAt: now, updatedAt: now })
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
    if (result.answer && result.answer.length > 10) {
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
  console.log(`Chunk size: ${CHUNK_SIZE} | Overlap: ${CHUNK_OVERLAP} lines`)
  console.log(`Phi-4 timeout: ${PHI4_TIMEOUT_MS}ms | Embed timeout: ${EMBED_TIMEOUT_MS}ms`)
  console.log(`Chunk cache TTL: ${CHUNK_CACHE_TTL}ms | Embed pool limit: ${EMBED_POOL_LIMIT}`)
  console.log(`Request timeout: ${REQUEST_TIMEOUT_MS}ms | Keyword short-circuit score: ${KEYWORD_SHORTCIRCUIT_SCORE}`)
  console.log(`Blob concurrency: ${BLOB_CONCURRENCY}`)
  console.log(`Azure blob client: ${blobServiceClient ? 'singleton ready' : 'MISSING connection string'}`)
  startApiKeyHealthChecker()
  warmupChunkCaches()
})

module.exports = app
