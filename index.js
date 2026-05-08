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
const MONGODB_DB  = process.env.MONGODB_DB || 'clientcreds'
const CHAT_HISTORY_URI = process.env.CHAT_HISTORY_URI
const CHAT_HISTORY_DB = process.env.CHAT_HISTORY_DB || 'chathistory'
const AZURE_CONNECTION_STRING = process.env.AZURE_CONNECTION_STRING || ''
const AZURE_CONTAINER_NAME = process.env.AZURE_CONTAINER_NAME || 'vectordbforrag'
const ADMIN_API_KEY  = process.env.ADMIN_API_KEY
const KEY_CHECK_INTERVAL_MS   = parseInt(process.env.KEY_CHECK_INTERVAL_MS || '300000', 10)
const PHI4_ENDPOINT = process.env.PHI4_ENDPOINT
const PHI4_API_KEY = process.env.PHI4_API_KEY
const PHI4_MODEL  = process.env.PHI4_MODEL || 'Phi-4-mini-instruct'
const PHI4_TIMEOUT_MS  = parseInt(process.env.PHI4_TIMEOUT_MS || '30000', 10)
const AZURE_EMBED_ENDPOINT = process.env.AZURE_EMBED_ENDPOINT || ''
const AZURE_EMBED_KEY = process.env.AZURE_EMBED_KEY || ''
const AZURE_EMBED_MODEL = process.env.AZURE_EMBED_MODEL || 'text-embedding-ada-002'
const EMBED_TIMEOUT_MS = parseInt(process.env.EMBED_TIMEOUT_MS || '10000', 10)
const EMBED_POOL_LIMIT = parseInt(process.env.EMBED_POOL_LIMIT || '20', 10)
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || '60000', 10)
const KEYWORD_SHORTCIRCUIT_SCORE = parseInt(process.env.KEYWORD_SHORTCIRCUIT_SCORE || '6', 10)
const WARMUP_CLIENT_IDS = (process.env.WARMUP_CLIENT_IDS || '').split(',').map(s => s.trim()).filter(Boolean)
const RAW_PREFIX = 'raw'
const CHUNK_SIZE    = 220  
const CHUNK_OVERLAP = 1     
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

// ─── Response Cache ───────────────────────────────────────────────────────────
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
let phiFailures  = 0
let phiBlockedUntil  = 0
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
    console.error(`🚨 [phi4] Circuit breaker OPEN for 30s after ${phiFailures} failures`)
  }
}
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

// ─── Helpers ──────────────────────────────────────────────────────────────────
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
  const DEFINITION_PATTERNS = [/^what\s+(is|are|does)\s+/,/^define\s+/,/^explain\s+/,/^meaning\s+of\s+/,/^tell\s+me\s+about\s+/,/^describe\s+/,/^how\s+is\s+.+\s+(calculated|defined|measured|computed)/,/\bmeaning\b/,/\bdefinition\b/,/\bwhat\s+does\b/]
  const LOOKUP_PATTERNS     = [/^(show|list|find|get|fetch|give)\s+(me\s+)?/,/^how\s+many\s+/,/^what\s+(is\s+the\s+)?(value|number|count|total|sum|amount)/]
  const COMPARISON_PATTERNS = [/\bvs\b|\bversus\b|\bdifference\b|\bcompare\b|\bbetween\b/]
  const URL_PATTERNS        = [/\burl\b/,/\blink\b/,/\breport\b.*\burl\b/,/\burl\b.*\breport\b/,/\bpower\s*bi\b/,/\bdashboard\b/]
  const GREETING_PATTERNS   = [/^(hi|hello|hey|howdy|greetings|good\s+(morning|afternoon|evening)|how\s+are\s+you|what's\s+up|sup)\b/]
  if (GREETING_PATTERNS.some(p => p.test(q))) return 'greeting'
  if (URL_PATTERNS.some(p => p.test(q)))        return 'url_lookup'
  if (DEFINITION_PATTERNS.some(p => p.test(q))) return 'definition'
  if (COMPARISON_PATTERNS.some(p => p.test(q))) return 'comparison'
  if (LOOKUP_PATTERNS.some(p => p.test(q)))     return 'lookup'
  return 'general'
}

function extractSubject(query) {
  const q = query.toLowerCase().trim().replace(/[?!.]+$/, '')
  const patterns = [
    /^what\s+is\s+(?:an?\s+|the\s+)?(.+)$/,
    /^what\s+are\s+(.+)$/,
    /^what\s+does\s+(.+?)\s+mean$/,
    /^define\s+(?:an?\s+|the\s+)?(.+)$/,
    /^explain\s+(?:an?\s+|the\s+)?(.+)$/,
    /^tell\s+me\s+about\s+(?:an?\s+|the\s+)?(.+)$/,
    /^meaning\s+of\s+(?:an?\s+|the\s+)?(.+)$/,
    /^describe\s+(?:an?\s+|the\s+)?(.+)$/,
    /^how\s+is\s+(.+?)\s+(calculated|defined|measured|computed)$/,
    /^definition\s+of\s+(?:an?\s+|the\s+)?(.+)$/,
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
      if (w.length < 3) continue
      if (!index.has(w)) index.set(w, new Set())
      index.get(w).add(i)
    }
  }
  return index
}

// ─── FIX 1: Exact term validation ────────────────────────────────────────────
// Checks that the exact subject term (not just partial words) exists in hits.
// Prevents "region" from matching "resolution", "occupancy" from matching "occupied", etc.
function containsExactTerm(query, hits) {
  const subject = extractSubject(query).toLowerCase().trim()
  const normalizedSubject = subject.replace(/[^\w\s]/g, '').replace(/\s+/g, ' ').trim()

  return hits.some(h => {
    const text = (h.text || '')
      .toLowerCase()
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')

    // Full phrase exact match
    if (text.includes(normalizedSubject)) return true

    // All words must appear as whole words (word boundary check)
    const words = normalizedSubject.split(' ').filter(w => w.length > 1)
    return words.length > 0 && words.every(w =>
      new RegExp(`\\b${escapeRegex(w)}\\b`).test(text)
    )
  })
}

// ─── Phi-4 call ───────────────────────────────────────────────────────────────
async function callPhi4(systemPrompt, userMessage) {
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
            temperature: 0.0,   // zero temperature — deterministic, no creative drift
            max_tokens:  400,   // reduced from 512 — keep answers tight
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

// ─── FIX 4+5: Keyword search — exact word boundaries, no partial matching ────
// REMOVED: partial word scoring (`subjectWords.filter(w => text.includes(w)).length * 2`)
// REPLACED: text.includes(subjectLower) → word-boundary regex (\\bterm\\b)
// This stops "region" hitting "resolution", "occupancy" hitting "occupied", etc.
function keywordSearch(query, chunks, topK, intent = 'general', invertedIndex = null) {
  const subject      = intent === 'definition' ? extractSubject(query) : query.toLowerCase()
  const queryLower   = query.toLowerCase()
  const subjectLower = subject.toLowerCase().trim()
  const subjectWords = subjectLower.split(/\s+/).filter(w => w.length > 1)
  const queryWords   = queryLower.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1)

  const urlKeywords = intent === 'url_lookup' ? extractUrlKeywords(query) : []

  // Build exact-match pattern (full phrase with word boundaries)
  const exactPattern = new RegExp(`\\b${escapeRegex(subjectLower)}\\b`, 'i')

  // ── FIX 6: No random fallback candidates ────────────────────────────────────
  // Old code: `chunks.slice(0, 100)` as fallback — this injected random chunks.
  // New code: if no index candidates, return empty — refuse to guess.
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

  // CRITICAL FIX: No candidates → return empty, not random chunks
  if (!candidateIndices || candidateIndices.size === 0) {
    return []
  }

  const source = [...candidateIndices].map(i => chunks[i]).filter(Boolean)

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
        // FIX 5: Exact word-boundary match (replaces text.includes which caused partial hits)
        if (exactPattern.test(text)) {
          score += 15  // strong signal: exact term present

          const defPattern = new RegExp(`${escapeRegex(subjectLower)}\\s*(is|are)\\s*(defined|described|calculated|measured|computed)`, 'i')
          if (defPattern.test(c.text || '')) score += subjectWords.length * 6
        }

        // NOTE: Removed `subjectWords.filter(w => text.includes(w)).length * 2`
        // That partial-word scoring caused region→resolution, occupancy→occupied contamination.

        if ((docType === DOC_TYPE.SPREADSHEET || docType === DOC_TYPE.DATA) && intent === 'definition') {
          const descPattern = new RegExp(`${escapeRegex(subjectLower)}\\s*(is described as|is defined as):`, 'i')
          if (descPattern.test(c.text || '')) score += subjectWords.length * 8
        }

        if (docType === DOC_TYPE.SPREADSHEET || docType === DOC_TYPE.DATA) {
          for (const w of subjectWords) {
            // Use word boundary here too, not loose includes
            const kvPattern = new RegExp(`:\\s*\\b${escapeRegex(w)}\\b|\\|\\s*\\b${escapeRegex(w)}\\b`, 'i')
            if (kvPattern.test(c.text || '')) score += 2
          }
        }

        // Full query match (secondary signal)
        if (exactPattern.test(queryLower) && text.includes(queryLower)) score += queryWords.length * 2
      }

      return { ...c, _score: score }
    })
    .filter(c => c._score > 0)
    .sort((a, b) => b._score - a._score)
    .slice(0, topK)
}

async function retrieveChunks(query, chunks, topK = 3, invertedIndex = null) {
  const intent          = detectQueryIntent(query)
  const normalizedQuery = query.toLowerCase().trim().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ')
  const subjectLower    = extractSubject(query).toLowerCase().trim()

  // ── FIX 8: Exact-match-first for definitions ─────────────────────────────
  // For definition queries, scan ALL chunks for exact term before any embedding.
  // This bypasses embedding entirely when we have a solid exact match.
  if (intent === 'definition') {
    const exactHits = chunks
      .map(c => {
        const text = (c.text || '').toLowerCase()
        const exact = new RegExp(`\\b${escapeRegex(subjectLower)}\\b`, 'i').test(text)
        return { ...c, _score: exact ? 20 : 0 }
      })
      .filter(c => c._score > 0)
      .slice(0, topK)  // FIX 10: topK=2 for definitions (set at call site)

    if (exactHits.length > 0) {
      console.log(`[retrieveChunks] exact-match-first: found ${exactHits.length} hits for "${subjectLower}"`)
      return exactHits
    }
    // No exact hits for definition — return empty, refuse to guess
    console.log(`[retrieveChunks] no exact hits for definition: "${subjectLower}"`)
    return []
  }

  const keywordTopK = Math.min(100, chunks.length)
  const candidates  = keywordSearch(normalizedQuery, chunks, keywordTopK, intent, invertedIndex)
  const pool        = candidates  // no fallback to random chunks

  if (pool.length === 0) return []

  // ── FIX 7: Safe keyword shortcut — only if exact term confirmed ───────────
  if (pool[0]._score >= KEYWORD_SHORTCIRCUIT_SCORE && containsExactTerm(query, pool)) {
    console.log(`[retrieveChunks] keyword short-circuit (score=${pool[0]._score}) — exact term confirmed`)
    return pool.slice(0, Math.min(topK, 5))
  }

  if (intent === 'url_lookup' && pool.length > 0) {
    return pool.slice(0, Math.min(topK, 6))
  }

  if (AZURE_EMBED_ENDPOINT && AZURE_EMBED_KEY) {
    try {
      const queryVec = await embedQueryAzure(normalizedQuery)
      if (queryVec) {
        const poolSlice  = pool.slice(0, EMBED_POOL_LIMIT)
        const chunkTexts = poolSlice.map(c => (c.text || '').toLowerCase().slice(0, 512))
        const embeddings = await embedBatch(chunkTexts)

        const maxKeyword = pool[0]._score || 1
        const weight     = { semantic: 0.50, keyword: 0.50 }  // balanced — keyword anchors precision

        const scored = poolSlice.map((c, i) => {
          const chunkVec = embeddings[i]
          if (!chunkVec) return c
          const semanticScore = cosineSim(queryVec, chunkVec)
          const keywordNorm   = typeof c._score === 'number' ? c._score / maxKeyword : 0
          return { ...c, _score: semanticScore * weight.semantic + keywordNorm * weight.keyword }
        })

        // ── FIX 9: Filter out low-confidence chunks ──────────────────────────
        const filtered = scored.filter(c => c._score >= 0.25)

        return filtered
          .sort((a, b) => b._score - a._score)
          .slice(0, Math.min(topK, 5))
      }
    } catch (err) {
      console.warn('[retrieveChunks] Azure embed failed, keyword fallback:', err.message)
    }
  }

  return pool.slice(0, Math.min(topK, 5))
}

// ── FIX 14: buildContext — compact, deduplicated, noise-filtered ──────────────
function buildContext(hits) {
  const seen = new Set()
  const deduped = []
  for (const h of hits) {
    const fingerprint = (h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fingerprint)) {
      seen.add(fingerprint)
      deduped.push(h)
    }
    if (deduped.length >= 3) break  // max 3 chunks for Phi-4-mini
  }

  return deduped.map((h, i) => {
    // Tighter char limits — keep context lean for small model
    const charLimit = i === 0 ? 400 : 280
    const lines = (h.text || '')
      .split('\n')
      // Remove noisy metadata lines
      .filter(l => !/(copyright|all rights reserved|page \d+|proprietary|redistribution|www\.|http)/i.test(l) || l.toLowerCase().includes('http'))
      // Remove empty or very short lines
      .filter(l => l.trim().length > 4)
      .join('\n')
      .trim()
      .slice(0, charLimit)

    return `[${i + 1}] ${lines}`
  }).join('\n\n')
}

const SYSTEM_PROMPT_CACHE     = new Map()
const SYSTEM_PROMPT_CACHE_MAX = 100

// ── FIX 12: Strict system prompt for Phi-4-mini ───────────────────────────────
// Old prompt allowed model to infer/substitute. New prompt enforces strict grounding.
function buildDynamicSystemPrompt(hits, intent = 'general') {
  const hasSpreadsheet = hits.some(h => classifyExtension(h.source_file || '') === DOC_TYPE.SPREADSHEET)
  const hasPdf         = hits.some(h => classifyExtension(h.source_file || '') === DOC_TYPE.PDF)
  const cacheKey       = `${intent}::${hasSpreadsheet}::${hasPdf}`
  const cached         = SYSTEM_PROMPT_CACHE.get(cacheKey)
  if (cached) return cached

  const base = `You are a STRICT document QA assistant.

Answer ONLY using the retrieved context below.

ABSOLUTE RULES:
- NEVER guess or infer
- NEVER use outside knowledge  
- NEVER substitute a similar word or concept
- NEVER answer if the exact term is not in context
- If a term is absent from context: say "I couldn't find that exact information in your documents."
- If retrieved context is weak or unrelated: say "I couldn't find an accurate answer for that in your documents."
- Do NOT add citation markers like [1], [2]
- Do NOT mention file names
- Keep answers concise and precise`

  const intentGuide = {
    definition: `

TASK — DEFINITION:
Find lines where the EXACT term asked is defined, described, or explained.
Look for: "X is defined as", "X is described as", "X means", "X: <explanation>".
Return the definition clearly. If the exact term is absent, say not found.`,

    lookup: `

TASK — LOOKUP:
Return the exact value, count, or list from context.
Report precise numbers only. If not found, say so.`,

    comparison: `

TASK — COMPARISON:
Find data for EACH item and compare clearly.
Use parallel structure. Note gaps explicitly if one side has less data.`,

    url_lookup: `

TASK — URL:
Find and return the full URL matching the topic.
Return the URL unbroken on its own line. If none match, say not found.`,

    general: `

TASK — GENERAL:
Answer directly from context. Stay focused on what was asked.`,
  }[intent] || `

TASK — GENERAL:
Answer directly from context. Stay focused on what was asked.`

  const spreadsheetGuide = hasSpreadsheet
    ? `

SPREADSHEET DATA: Rows are pipe-separated key:value pairs.
Lines ending "is described as: ..." are definition summaries — prioritise these.
Write a natural sentence from data. Never output raw pipe-delimited rows verbatim.`
    : ''

  const pdfGuide = hasPdf
    ? `

PDF DATA: Ignore headers, footers, page numbers. Focus on substantive content.`
    : ''

  const prompt = `${base}${intentGuide}${spreadsheetGuide}${pdfGuide}`

  if (SYSTEM_PROMPT_CACHE.size >= SYSTEM_PROMPT_CACHE_MAX) {
    SYSTEM_PROMPT_CACHE.delete(SYSTEM_PROMPT_CACHE.keys().next().value)
  }
  SYSTEM_PROMPT_CACHE.set(cacheKey, prompt)
  return prompt
}

// ─── Mongo helpers ────────────────────────────────────────────────────────────
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

// ─── Document extraction ──────────────────────────────────────────────────────
async function extractPdf(buffer)  { const r = await pdfParse(buffer); return r.text || '' }
async function extractWord(buffer) { const r = await mammoth.extractRawText({ buffer }); return r.value || '' }

function extractSpreadsheet(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' })
  const parts    = []
  for (const sheetName of workbook.SheetNames) {
    const sheet   = workbook.Sheets[sheetName]
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 })
    if (!rawRows.length) continue
    parts.push(`=== Sheet: ${sheetName} ===`)
    let headerRowIdx = 0
    for (let i = 0; i < Math.min(10, rawRows.length); i++) {
      if (rawRows[i].some(cell => String(cell).trim() !== '')) { headerRowIdx = i; break }
    }
    const headers = rawRows[headerRowIdx].map(h => String(h).trim())
    for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
      const row = rawRows[i]
      if (!row.some(cell => String(cell).trim() !== '')) continue
      const pairs = []
      for (let j = 0; j < Math.max(headers.length, row.length); j++) {
        const val = String(row[j] || '').trim()
        if (!val) continue
        const key = headers[j] && headers[j] !== '' ? headers[j] : `Field${j + 1}`
        pairs.push(`${key}: ${val}`)
      }
      if (pairs.length > 0) {
        parts.push(pairs.join(' | '))
        // FIX 3: REMOVED synthetic "X is described as: ..." generation.
        // This was creating false semantic relationships that contaminated embeddings.
        // e.g. "Lease Resolution is described as: ..." caused occupancy→lease resolution hits.
      }
    }
    parts.push('')
    parts.push('[All values in this sheet:]')
    for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
      const row = rawRows[i]
      const rowValues = row.map((cell, j) => {
        const val = String(cell || '').trim()
        if (!val) return ''
        const key = headers[j] && headers[j] !== '' ? headers[j] : `Field${j + 1}`
        return `${val} (${key})`
      }).filter(Boolean)
      if (rowValues.length) parts.push(rowValues.join(', '))
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

function chunkText(text, sourceFile) {
  const chunks = []
  let index    = 0
  const lines  = text.replace(/\r\n/g, '\n').split('\n').map(l => l.trim()).filter(l => l.length > 0)
  let buffer   = []
  for (const line of lines) {
    const projectedLength = buffer.join('\n').length + (buffer.length ? 1 : 0) + line.length
    if (buffer.length > 0 && projectedLength > CHUNK_SIZE) {
      const chunkTextStr = buffer.join('\n')
      if (chunkTextStr.length > 30) chunks.push({ text: chunkTextStr, source_file: sourceFile, chunk_index: index++, embedding: [] })
      buffer = buffer.slice(-CHUNK_OVERLAP)
    }
    buffer.push(line)
  }
  if (buffer.length > 0) {
    const chunkTextStr = buffer.join('\n')
    if (chunkTextStr.length > 30) chunks.push({ text: chunkTextStr, source_file: sourceFile, chunk_index: index++, embedding: [] })
  }
  return chunks
}

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
      .then(({ chunks }) => console.log(`[warmup] ✓ ${clientId} — ${chunks.length} chunks ready`))
      .catch(err => console.warn(`[warmup] ✗ ${clientId} — ${err.message}`))
  }
}

// ─── Answer builder ────────────────────────────────────────────────────────────
async function answerWithPhi4(originalQuery, hits, intent = 'general') {
  const systemPrompt = buildDynamicSystemPrompt(hits, intent)
  const context      = buildContext(hits)

  const subjectHint  = intent === 'definition'
    ? `\nDefine ONLY the exact term: "${extractSubject(originalQuery)}". Do not define any other term. If not found in context, say "I couldn't find that exact information in your documents."`
    : intent === 'url_lookup'
    ? `\nReturn URL(s) for: "${extractUrlKeywords(originalQuery).join(' ')}".`
    : ''

  const userMessage = `CONTEXT:\n${context}${subjectHint}\n\nQuestion: ${originalQuery}`
  return callPhi4(systemPrompt, userMessage)
}

// ── FIX 13: Strict fallback — exact match only, no guessing ──────────────────
function buildFallbackAnswer(query, hits) {
  if (!hits?.length) {
    return "I couldn't find that in your documents."
  }

  const subject = extractSubject(query).toLowerCase().trim()

  // Only return a fallback if the exact term actually appears in the chunk
  const exact = hits.find(h => {
    const text = (h.text || '').toLowerCase()
    return new RegExp(`\\b${escapeRegex(subject)}\\b`, 'i').test(text)
  })

  if (!exact) {
    return "I couldn't find an exact match for that term in your documents."
  }

  // Return a clean snippet from the matching chunk
  return exact.text
    .replace(/Field\d+:\s*/gi, '')
    .replace(/\|\s*Field\d+\b/gi, '')
    .split('\n')
    .filter(l => !/(copyright|all rights reserved|proprietary|confidential)/i.test(l))
    .join('\n')
    .slice(0, 300)
    .trim()
}

function generateTitle(query) {
  const cleaned = query.trim().replace(/[?!.]+$/, '')
  return cleaned.length > 50 ? cleaned.slice(0, 50) + '…' : cleaned
}

// ─── Routes ───────────────────────────────────────────────────────────────────
app.get('/health', (req, res) => res.json({
  ok:               true,
  service:          'ask-data',
  model:            'ask-data-response-model',
  embeddings:       AZURE_EMBED_ENDPOINT ? 'azure-anurit' : 'keyword-only',
  chunkCacheSize:   CHUNK_CACHE.size,
  promptCacheSize:  SYSTEM_PROMPT_CACHE.size,
  responseCacheSize: RESPONSE_CACHE.size,
  phiCircuitOpen:   phiCircuitOpen(),
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

// ─── MAIN CHAT ENDPOINT ───────────────────────────────────────────────────────
app.post('/chat/message', requireClientKey, withRequestTimeout(async (req, res) => {
  try {
    const { query, topK = 3, conversationId } = req.body  // FIX 10: default topK=3
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

      // FIX 10: Use intent-appropriate topK — small for definitions, slightly larger for general
      const effectiveTopK = intent === 'definition' ? 2 : Math.min(topK, 5)
      const hits = await retrieveChunks(query.trim(), chunks, effectiveTopK, invertedIndex)

      // ── FIX 15: Retrieval debug logging ──────────────────────────────────────
      console.log({
        query: query.trim(),
        intent,
        hitsCount: hits.length,
        topHits: hits.slice(0, 5).map(h => ({
          score:   h._score,
          source:  h.source_file,
          preview: h.text?.slice(0, 120),
        })),
      })

      // ── FIX 2: Strict confidence filter — never send weak hits to Phi-4 ─────
      const bestScore = hits[0]?._score || 0
      if (hits.length === 0 || bestScore < 3 || !containsExactTerm(query, hits)) {
        console.log(`[confidence] REFUSED — bestScore=${bestScore}, exactMatch=${containsExactTerm(query, hits)}`)
        return {
          answer: "I couldn't find an accurate answer for that in your documents.",
          sources: [],
          conversationId: conversationId || null,
          client: { clientId, name },
        }
      }

      let rawAnswer
      try {
        rawAnswer = await Promise.race([
          answerWithPhi4(query.trim(), hits, intent),
          new Promise((_, reject) =>
            setTimeout(() => reject(new Error('Model response timeout (8s)')), 8000)
          ),
        ])
      } catch (err) {
        console.warn(`⚠️ [phi4] Using fallback answer: ${err.message}`)
        rawAnswer = buildFallbackAnswer(query.trim(), hits)
      }

      const cleanAnswer = fixBrokenUrls(rawAnswer)
        .replace(/\bField\d+\s*:\s*/gi, '')
        .replace(/\|\s*Field\d+\b/gi, '')
        .trim()

      const answer  = cleanAnswer
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
  console.log(`Phi-4 timeout: ${PHI4_TIMEOUT_MS}ms | Embed timeout: ${EMBED_TIMEOUT_MS}ms`)
  console.log(`Chunk cache TTL: ${CHUNK_CACHE_TTL}ms | Embed pool limit: ${EMBED_POOL_LIMIT}`)
  console.log(`Request timeout: ${REQUEST_TIMEOUT_MS}ms | Keyword short-circuit score: ${KEYWORD_SHORTCIRCUIT_SCORE}`)
  console.log(`Blob concurrency: ${BLOB_CONCURRENCY}`)
  console.log(`Chunk size: ${CHUNK_SIZE} | Chunk overlap: ${CHUNK_OVERLAP}`)
  console.log(`Azure blob client: ${blobServiceClient ? 'singleton ready' : 'MISSING connection string'}`)
  startApiKeyHealthChecker()
  warmupChunkCaches()
})
module.exports = app
