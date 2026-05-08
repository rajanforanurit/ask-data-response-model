require('dotenv').config()
const express = require('express')
const cors = require('cors')
const { MongoClient, ObjectId } = require('mongodb')
const { BlobServiceClient } = require('@azure/storage-blob')
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
  origin:         (origin, cb) => cb(null, originAllowed(origin)),
  methods:        ['GET','POST','PUT','PATCH','DELETE','OPTIONS'],
  allowedHeaders: ['Content-Type','Authorization','x-session-id'],
  credentials:    true,
}))
app.options('*', cors({
  origin:         (origin, cb) => cb(null, true),
  methods:        ['GET','POST','PUT','PATCH','DELETE','OPTIONS'],
  allowedHeaders: ['Content-Type','Authorization','x-session-id'],
  credentials:    true,
}))
app.use(express.json())

// ─── ENV ───────────────────────────────────────────────────────────────────────
const MONGODB_URI          = process.env.MONGODB_URI
const MONGODB_DB           = process.env.MONGODB_DB           || 'clientcreds'
const CHAT_HISTORY_URI     = process.env.CHAT_HISTORY_URI
const CHAT_HISTORY_DB      = process.env.CHAT_HISTORY_DB      || 'chathistory'
const AZURE_CONNECTION_STRING = process.env.AZURE_CONNECTION_STRING || ''
const AZURE_CONTAINER_NAME = process.env.AZURE_CONTAINER_NAME || 'vectordbforrag'
const ADMIN_API_KEY        = process.env.ADMIN_API_KEY
const KEY_CHECK_INTERVAL_MS = parseInt(process.env.KEY_CHECK_INTERVAL_MS || '300000', 10)

const PHI4_ENDPOINT    = process.env.PHI4_ENDPOINT
const PHI4_API_KEY     = process.env.PHI4_API_KEY
const PHI4_MODEL       = process.env.PHI4_MODEL       || 'Phi-4-mini-instruct'
const PHI4_TIMEOUT_MS  = parseInt(process.env.PHI4_TIMEOUT_MS  || '30000', 10)
const MINILM_SIDECAR_URL   = process.env.MINILM_SIDECAR_URL   || 'http://localhost:5001'
const EMBED_TIMEOUT_MS      = parseInt(process.env.EMBED_TIMEOUT_MS      || '5000',  10)

const REQUEST_TIMEOUT_MS         = parseInt(process.env.REQUEST_TIMEOUT_MS         || '60000', 10)
const KEYWORD_SHORTCIRCUIT_SCORE = parseInt(process.env.KEYWORD_SHORTCIRCUIT_SCORE || '6',     10)
const WARMUP_CLIENT_IDS          = (process.env.WARMUP_CLIENT_IDS || '').split(',').map(s => s.trim()).filter(Boolean)
const BLOB_CONCURRENCY           = parseInt(process.env.BLOB_CONCURRENCY || '8', 10)
const VECTORS_PREFIX = (clientId) => `meta/${clientId}/vectors/`

const blobServiceClient = AZURE_CONNECTION_STRING
  ? BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING)
  : null

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
    RESPONSE_CACHE.delete(RESPONSE_CACHE.keys().next().value)
  }
  RESPONSE_CACHE.set(key, { value, ts: Date.now() })
}
function getCacheKey(clientId, query) {
  return `${clientId}:${query.toLowerCase().trim()}`
}

// ─── Phi-4 concurrency + circuit breaker ──────────────────────────────────────
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
          r  => { phiActiveCount--; drainPhiQueue(); resolve(r) },
          e  => { phiActiveCount--; drainPhiQueue(); reject(e) }
        )
      } else { phiQueue.push(tryRun) }
    }
    tryRun()
  })
}
function drainPhiQueue() {
  if (phiQueue.length > 0 && phiActiveCount < PHI_MAX_CONCURRENT) phiQueue.shift()()
}
let phiFailures = 0, phiBlockedUntil = 0
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
    console.error(`🚨 [phi4] Circuit breaker OPEN for 30s after ${phiFailures} failures`)
  }
}

// ─── Fetch helpers ────────────────────────────────────────────────────────────
async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    return await fetch(url, { ...options, signal: controller.signal })
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
    try { await fn(req, res, next) } catch (err) { if (!settled) next(err) } finally {
      settled = true; clearTimeout(timer)
    }
  }
}
async function embedQuery(query) {
  try {
    const response = await fetchWithTimeout(
      `${MINILM_SIDECAR_URL}/embed`,
      {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ text: query }),
      },
      EMBED_TIMEOUT_MS
    )
    if (!response.ok) return null
    const data = await response.json()
    return data.embedding || null          // float[]
  } catch (err) {
    console.warn('[embedQuery] MiniLM sidecar unavailable, keyword-only mode:', err.message)
    return null
  }
}

// ─── Pure-JS cosine similarity ────────────────────────────────────────────────
function cosineSim(a, b) {
  let dot = 0, na = 0, nb = 0
  for (let i = 0; i < a.length; i++) {
    dot += a[i] * b[i]
    na  += a[i] * a[i]
    nb  += b[i] * b[i]
  }
  return dot / (Math.sqrt(na) * Math.sqrt(nb) + 1e-9)
}
const DOC_TYPE = {
  SPREADSHEET: 'spreadsheet', PDF: 'pdf', WORD: 'word',
  PRESENTATION: 'presentation', CODE: 'code', DATA: 'data',
  TEXT: 'text', EMAIL: 'email', WEB: 'web', UNKNOWN: 'unknown',
}

function classifyExtension(fileName) {
  const ext = ('.' + (fileName || '').split('.').pop()).toLowerCase()
  if (['.xlsx','.xls','.ods'].includes(ext))                         return DOC_TYPE.SPREADSHEET
  if (ext === '.pdf')                                                 return DOC_TYPE.PDF
  if (['.docx','.doc','.odt','.rtf'].includes(ext))                 return DOC_TYPE.WORD
  if (['.pptx','.ppt'].includes(ext))                               return DOC_TYPE.PRESENTATION
  if (['.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h',
       '.cs','.go','.rb','.php','.swift','.kt','.r','.sql','.sh',
       '.bash','.ps1'].includes(ext))                                return DOC_TYPE.CODE
  if (['.json','.jsonl','.yaml','.yml','.toml','.csv','.tsv'].includes(ext)) return DOC_TYPE.DATA
  if (['.txt','.md','.markdown','.rst'].includes(ext))              return DOC_TYPE.TEXT
  if (ext === '.eml')                                               return DOC_TYPE.EMAIL
  if (['.html','.htm','.xml'].includes(ext))                       return DOC_TYPE.WEB
  return DOC_TYPE.UNKNOWN
}

function detectQueryIntent(query) {
  const q = query.toLowerCase().trim()
  const GREETING_PATTERNS   = [/^(hi|hello|hey|howdy|greetings|good\s+(morning|afternoon|evening)|how\s+are\s+you|what's\s+up|sup)\b/]
  const URL_PATTERNS        = [/\burl\b/,/\blink\b/,/\breport\b.*\burl\b/,/\burl\b.*\breport\b/,/\bpower\s*bi\b/,/\bdashboard\b/]
  // lookup BEFORE definition so "what is application count" → lookup not definition
  const LOOKUP_PATTERNS     = [
    /^(show|list|find|get|fetch|give)\s+(me\s+)?/,
    /^how\s+many\s+/,
    /\bcount\b/,
    /\btotal\b/,
    /\bsum\b/,
    /\bnumber\s+of\b/,
    /^what\s+(is\s+the\s+)?(value|number|count|total|sum|amount)\b/,
  ]
  const DEFINITION_PATTERNS = [
    /^what\s+(is|are|does)\s+/,
    /^define\s+/,
    /^explain\s+/,
    /^meaning\s+of\s+/,
    /^tell\s+me\s+about\s+/,
    /^describe\s+/,
    /^how\s+is\s+.+\s+(calculated|defined|measured|computed)/,
    /\bmeaning\b/,
    /\bdefinition\b/,
    /\bwhat\s+does\b/,
  ]
  const COMPARISON_PATTERNS = [/\bvs\b|\bversus\b|\bdifference\b|\bcompare\b|\bbetween\b/]
  if (GREETING_PATTERNS.some(p => p.test(q)))   return 'greeting'
  if (URL_PATTERNS.some(p => p.test(q)))         return 'url_lookup'
  if (LOOKUP_PATTERNS.some(p => p.test(q)))      return 'lookup'
  if (DEFINITION_PATTERNS.some(p => p.test(q))) return 'definition'
  if (COMPARISON_PATTERNS.some(p => p.test(q))) return 'comparison'
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
  ]
  for (const p of patterns) { const m = q.match(p); if (m) return m[1].trim() }
  return q
}

function extractUrlKeywords(query) {
  const q = query.toLowerCase()
  const stop = new Set(['power','bi','report','url','link','for','the','a','an','of','in','get','me','show','give','find','fetch'])
  return q.replace(/[^\w\s]/g,' ').split(/\s+/).filter(w => w.length > 1 && !stop.has(w))
}

function fixBrokenUrls(text) {
  return text.replace(/https:\/\/[^\s]+(\s+[^\s]+)/g, m => m.replace(/\s/g,''))
}

function escapeRegex(str) { return str.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') }

function buildInvertedIndex(chunks) {
  const index = new Map()
  for (let i = 0; i < chunks.length; i++) {
    const words = (chunks[i].text || '').toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
    for (const w of words) {
      if (w.length < 3) continue
      if (!index.has(w)) index.set(w, new Set())
      index.get(w).add(i)
    }
  }
  return index
}

// ─── Phi-4 ────────────────────────────────────────────────────────────────────
async function callPhi4(systemPrompt, userMessage) {
  if (!PHI4_ENDPOINT || !PHI4_API_KEY) throw new Error('PHI4_ENDPOINT and PHI4_API_KEY are required')
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
            max_tokens:  512,
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
async function _doLoadVectors(clientId) {
  if (!blobServiceClient) throw new Error('AZURE_CONNECTION_STRING not set')

  const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
  const prefix = VECTORS_PREFIX(clientId)
  const blobNames = []
  for await (const blob of containerClient.listBlobsFlat({ prefix })) {
    if (blob.name.endsWith('.json')) blobNames.push(blob.name)
  }

  if (blobNames.length === 0) {
    console.warn(`[vectorLoader] No vector files found at prefix: ${prefix}`)
    return []
  }

  const allChunks = []

  for (let i = 0; i < blobNames.length; i += BLOB_CONCURRENCY) {
    const batch   = blobNames.slice(i, i + BLOB_CONCURRENCY)
    const results = await Promise.allSettled(
      batch.map(async (blobName) => {
        const downloadResponse = await containerClient.getBlobClient(blobName).download()
        const parts = []
        for await (const chunk of downloadResponse.readableStreamBody)
          parts.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
        const raw  = Buffer.concat(parts).toString('utf-8')
        const records = JSON.parse(raw)
        if (!Array.isArray(records)) {
          console.warn(`[vectorLoader] Unexpected format in ${blobName} — expected array`)
          return []
        }
        // Validate each record has an embedding
        return records.filter(r => r && r.text && Array.isArray(r.embedding) && r.embedding.length > 0)
      })
    )

    for (const result of results) {
      if (result.status === 'fulfilled') allChunks.push(...result.value)
      else console.warn(`[vectorLoader] Failed to load blob:`, result.reason?.message)
    }
  }

  console.log(`[vectorLoader] Loaded ${allChunks.length} pre-embedded chunks for client '${clientId}' from ${blobNames.length} vector files`)
  return allChunks
}

// ─── Chunk / vector cache ─────────────────────────────────────────────────────
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
      console.log(`[vectorCache] Stale for ${clientId}, refreshing in background`)
      const refreshPromise = _doLoadVectors(clientId)
        .then(chunks => {
          const invertedIndex = buildInvertedIndex(chunks)
          CHUNK_CACHE.set(clientId, { chunks, invertedIndex, ts: Date.now(), loading: null })
          console.log(`[vectorCache] Background refresh done for ${clientId}: ${chunks.length} chunks`)
          return chunks
        })
        .catch(err => {
          const existing = CHUNK_CACHE.get(clientId)
          CHUNK_CACHE.set(clientId, { ...existing, loading: null })
          console.warn(`[vectorCache] Background refresh failed for ${clientId}: ${err.message}`)
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

  const loadPromise = _doLoadVectors(clientId)
    .then(chunks => {
      const invertedIndex = buildInvertedIndex(chunks)
      CHUNK_CACHE.set(clientId, { chunks, invertedIndex, ts: Date.now(), loading: null })
      console.log(`[vectorCache] Loaded ${chunks.length} vectors + built inverted index for ${clientId}`)
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
  console.log(`[vectorCache] Invalidated cache for client: ${clientId}`)
}

function warmupChunkCaches() {
  if (!WARMUP_CLIENT_IDS.length || !blobServiceClient) return
  console.log(`[warmup] Pre-loading vectors for ${WARMUP_CLIENT_IDS.length} client(s): ${WARMUP_CLIENT_IDS.join(', ')}`)
  for (const clientId of WARMUP_CLIENT_IDS) {
    loadChunksForClient(clientId)
      .then(({ chunks }) => console.log(`[warmup] ✓ ${clientId} — ${chunks.length} chunks ready`))
      .catch(err => console.warn(`[warmup] ✗ ${clientId} — ${err.message}`))
  }
}

// ─── Keyword search (secondary signal / fallback) ─────────────────────────────
function keywordSearch(query, chunks, topK, intent = 'general', invertedIndex = null) {
  const subject      = intent === 'definition' ? extractSubject(query) : query.toLowerCase()
  const queryLower   = query.toLowerCase()
  const subjectLower = subject.toLowerCase()
  const subjectWords = subjectLower.split(/\s+/).filter(w => w.length > 1)
  const queryWords   = queryLower.replace(/[^\w\s]/g,' ').split(/\s+/).filter(w => w.length > 1)
  const urlKeywords  = intent === 'url_lookup' ? extractUrlKeywords(query) : []

  let candidateIndices
  if (invertedIndex && subjectWords.length > 0) {
    const wordsToIndex = intent === 'url_lookup' ? urlKeywords : subjectWords
    const sets = wordsToIndex.map(w => invertedIndex.get(w) || new Set())
    if (intent === 'url_lookup') {
      sets.push(invertedIndex.get('url') || new Set())
      sets.push(invertedIndex.get('link') || new Set())
      sets.push(invertedIndex.get('http') || new Set())
    }
    const union = new Set()
    for (const s of sets) for (const idx of s) union.add(idx)
    candidateIndices = union
  }

  const source = candidateIndices
    ? [...candidateIndices].map(i => chunks[i]).filter(Boolean)
    : chunks.slice(0, 200)

  return source.map(c => {
    const text = (c.text || '').toLowerCase()
    let score  = 0

    if (intent === 'url_lookup') {
      if (!text.includes('http')) return { ...c, _kwScore: 0 }
      const topicMatches = urlKeywords.filter(w => text.includes(w)).length
      score += topicMatches * 10
      if (topicMatches === 0) return { ...c, _kwScore: 0 }
      if (text.includes(urlKeywords.join(' '))) score += 15
    } else {
      if (text.includes(subjectLower)) {
        score += subjectWords.length * 4
        const defPat = new RegExp(`${escapeRegex(subjectLower)}\\s*(is|are)\\s*(defined|described|calculated|measured|computed)`,'i')
        if (defPat.test(c.text || '')) score += subjectWords.length * 6
      }
      score += subjectWords.filter(w => text.includes(w)).length * 2

      const docType = classifyExtension(c.source_file || '')
      if ((docType === DOC_TYPE.SPREADSHEET || docType === DOC_TYPE.DATA) && intent === 'definition') {
        const descPat = new RegExp(`${escapeRegex(subjectLower)}\\s*(is described as|is defined as):`,'i')
        if (descPat.test(c.text || '')) score += subjectWords.length * 8
      }
      if (text.includes(queryLower)) score += queryWords.length * 2
    }

    return { ...c, _kwScore: score }
  })
    .filter(c => c._kwScore > 0)
    .sort((a, b) => b._kwScore - a._kwScore)
    .slice(0, topK)
}
async function retrieveChunks(query, chunks, topK = 6, invertedIndex = null) {
  const intent = detectQueryIntent(query)
  if (intent === 'url_lookup') {
    const kwResults = keywordSearch(query, chunks, Math.min(topK + 4, 20), intent, invertedIndex)
    return kwResults.slice(0, topK).map(c => ({ ...c, _score: c._kwScore || 0 }))
  }
  const kwTopK      = intent === 'definition' ? Math.min(300, chunks.length) : Math.min(200, chunks.length)
  const kwCandidates = keywordSearch(query, chunks, kwTopK, intent, invertedIndex)

  // Strong keyword short-circuit (very specific term match — no embedding needed)
  if (kwCandidates.length > 0 && kwCandidates[0]._kwScore >= KEYWORD_SHORTCIRCUIT_SCORE && intent === 'definition') {
    console.log(`[retrieve] keyword short-circuit score=${kwCandidates[0]._kwScore}`)
    return kwCandidates.slice(0, Math.min(topK, 10)).map(c => ({ ...c, _score: c._kwScore || 0 }))
  }

  // Working set: use keyword candidates if available, else all chunks (capped)
  const workingSet = kwCandidates.length > 0 ? kwCandidates : chunks.slice(0, 500)
  const maxKw      = (kwCandidates[0]?._kwScore) || 1

  // ── Embed query with MiniLM sidecar ───────────────────────────────────────
  const queryVec = await embedQuery(query)

  if (queryVec) {
    // Cosine similarity against every chunk in working set
    const COSINE_WEIGHT  = 0.85
    const KEYWORD_WEIGHT = 0.15

    const scored = workingSet.map(c => {
      const cosine  = (c.embedding && c.embedding.length > 0)
        ? cosineSim(queryVec, c.embedding)
        : 0
      const kwNorm  = typeof c._kwScore === 'number' ? c._kwScore / maxKw : 0
      const blended = cosine * COSINE_WEIGHT + kwNorm * KEYWORD_WEIGHT
      return { ...c, _score: blended, _cosine: cosine, _kwNorm: kwNorm }
    })
    let fullScored = scored
    if (kwCandidates.length > 0 && kwCandidates.length < chunks.length) {
      const kwSet = new Set(kwCandidates.map(c => c.chunk_id || c.text?.slice(0,40)))
      const rest  = chunks
        .filter(c => !kwSet.has(c.chunk_id || c.text?.slice(0,40)))
        .slice(0, 200)
        .map(c => {
          const cosine = (c.embedding && c.embedding.length > 0)
            ? cosineSim(queryVec, c.embedding)
            : 0
          return { ...c, _score: cosine * COSINE_WEIGHT, _cosine: cosine }
        })
      fullScored = [...scored, ...rest]
    }

    return fullScored
      .sort((a, b) => b._score - a._score)
      .slice(0, Math.min(topK, 12))
  }

  // ── Fallback: keyword-only (sidecar unavailable) ──────────────────────────
  console.warn('[retrieve] MiniLM sidecar unavailable — keyword-only mode')
  return kwCandidates.slice(0, Math.min(topK, 12)).map(c => ({ ...c, _score: c._kwScore || 0 }))
}

// ─── Context builder ──────────────────────────────────────────────────────────
function buildContext(hits) {
  const seen   = new Set()
  const deduped = []
  for (const h of hits) {
    const fp = (h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
    if (deduped.length >= 6) break
  }
  return deduped.map((h, i) => {
    const limit = i === 0 ? 600 : i <= 2 ? 450 : 350
    return `[${i + 1}] ${(h.text || '').trim().slice(0, limit)}`
  }).join('\n\n')
}

// ─── System prompt builder ────────────────────────────────────────────────────
const SYSTEM_PROMPT_CACHE     = new Map()
const SYSTEM_PROMPT_CACHE_MAX = 100

function buildDynamicSystemPrompt(hits, intent = 'general') {
  const hasSpreadsheet = hits.some(h => classifyExtension(h.source_file || '') === DOC_TYPE.SPREADSHEET)
  const hasPdf         = hits.some(h => classifyExtension(h.source_file || '') === DOC_TYPE.PDF)
  const cacheKey       = `${intent}::${hasSpreadsheet}::${hasPdf}`
  const cached         = SYSTEM_PROMPT_CACHE.get(cacheKey)
  if (cached) return cached

  const base = `You are a highly precise document QA assistant.

You MUST follow this reasoning process:
1. Read ALL chunks carefully
2. Identify which chunk directly answers the question
3. Extract ONLY relevant lines
4. Combine them into a clear answer

STRICT RULES:
- Do NOT guess
- Do NOT use outside knowledge
- If multiple chunks conflict → mention it
- If no answer → say: "I couldn't find that in your documents."`

  const intentGuide = {
    definition: `
DEFINITION TASK: Find lines containing "is described as", "is defined as", or a clear explanation of the exact term asked. Synthesise into a clean 1-2 sentence definition. If not found, say "I couldn't find a definition for that term in your documents." Do NOT define a different term.`,
    lookup: `
LOOKUP TASK: Return the exact value, count, or list from CONTEXT. Report precise numbers — no approximations.`,
    comparison: `
COMPARISON TASK: Find information for EACH item and compare them clearly. Note any gaps.`,
    url_lookup: `
URL TASK: Find and return the full URL that matches the topic. Return the URL on its own line, unbroken. If multiple match, list each. If none, say not found.`,
    general: `
GENERAL TASK: Synthesise a clear, accurate answer from CONTEXT. Stay focused on what was asked.`,
  }[intent] || '\nGENERAL TASK: Synthesise a clear, accurate answer from CONTEXT.'

  const spreadsheetGuide = hasSpreadsheet
    ? `\nSPREADSHEET DATA: Rows appear as pipe-separated key:value pairs. Lines ending in "is described as: ..." are definition summaries — prioritise these. Write a natural sentence — never output raw pipe-delimited rows verbatim.`
    : ''

  const pdfGuide = hasPdf
    ? `\nPDF DATA: Ignore repeated headers, footers, copyright lines, and page numbers.`
    : ''

  const prompt = `${base}${intentGuide}${spreadsheetGuide}${pdfGuide}`

  if (SYSTEM_PROMPT_CACHE.size >= SYSTEM_PROMPT_CACHE_MAX)
    SYSTEM_PROMPT_CACHE.delete(SYSTEM_PROMPT_CACHE.keys().next().value)
  SYSTEM_PROMPT_CACHE.set(cacheKey, prompt)
  return prompt
}

// ─── Answer builders ──────────────────────────────────────────────────────────
async function answerWithPhi4(originalQuery, hits, intent = 'general') {
  const systemPrompt = buildDynamicSystemPrompt(hits, intent)
  const context      = buildContext(hits)
  const subjectHint  =
    intent === 'definition'
      ? `\nDefine ONLY this exact term: "${extractSubject(originalQuery)}". Do not define any other term.`
      : intent === 'url_lookup'
      ? `\nReturn URL(s) for: "${extractUrlKeywords(originalQuery).join(' ')}".`
      : ''
  const userMessage = `CONTEXT:\n${context}${subjectHint}\n\nQuestion: ${originalQuery}`
  return callPhi4(systemPrompt, userMessage)
}

function buildFallbackAnswer(query, hits) {
  if (!hits || hits.length === 0) return "I couldn't find relevant information in your documents for this query."
  const intent = detectQueryIntent(query)

  if (intent === 'url_lookup') {
    const urlKeywords = extractUrlKeywords(query)
    for (const h of hits) {
      for (const line of (h.text || '').split('\n')) {
        if (!line.toLowerCase().includes('http')) continue
        if (urlKeywords.filter(w => line.toLowerCase().includes(w)).length > 0) {
          const m = line.match(/https?:\/\/\S+/)
          if (m) return m[0].replace(/\s/g,'')
        }
      }
    }
  }

  const subject  = extractSubject(query).toLowerCase()
  const descLine = hits.find(h => {
    const t = (h.text || '').toLowerCase()
    return t.includes('is described as') && t.includes(subject)
  })
  if (descLine) {
    const lines = (descLine.text || '').split('\n')
    const rel   = lines.find(l => l.toLowerCase().includes(subject) && l.toLowerCase().includes('is described as'))
    if (rel) {
      const parts = rel.split(/is described as:/i)
      if (parts[1]) return `${subject.charAt(0).toUpperCase() + subject.slice(1)} is ${parts[1].trim().slice(0, 300)}`
    }
  }

  const best    = hits.find(h => (h.text || '').toLowerCase().includes(subject)) || hits[0]
  const snippet = (best.text || '')
    .replace(/Field\d+:\s*/gi,'').replace(/\|\s*Field\d+\b/gi,'')
    .split('\n')
    .filter(l => !/(copyright|all rights reserved|proprietary|confidential|redistribution)/i.test(l))
    .join('\n').slice(0, 350).trim()
  return snippet || "I couldn't find that specific information in your documents."
}

function generateTitle(query) {
  const cleaned = query.trim().replace(/[?!.]+$/,'')
  return cleaned.length > 50 ? cleaned.slice(0,50) + '…' : cleaned
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

const CLIENT_CACHE = new Map()
const CACHE_TTL_MS = 5 * 60 * 1000

function getCached(apiKey) {
  const entry = CLIENT_CACHE.get(apiKey)
  if (!entry) return null
  if (Date.now() - entry.cachedAt > CACHE_TTL_MS) { CLIENT_CACHE.delete(apiKey); return null }
  return entry
}
function setCache(apiKey, data)  { CLIENT_CACHE.set(apiKey, { ...data, cachedAt: Date.now() }) }
function evictCache(apiKey)      { if (apiKey) CLIENT_CACHE.delete(apiKey) }

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

function generateApiKey() { return `rak_${crypto.randomBytes(32).toString('hex')}` }

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/health', (req, res) => res.json({
  ok:               true,
  service:          'ask-data',
  retrievalMode:    'minilm-vector-cosine',
  embeddingModel:   'all-MiniLM-L6-v2',
  sidecarUrl:       MINILM_SIDECAR_URL,
  vectorCacheSize:  CHUNK_CACHE.size,
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
    if (!apiKey) { apiKey = generateApiKey() }
    else if (!apiKey.startsWith('rak_')) return res.status(400).json({ error: 'apiKey must start with "rak_"' })
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
  res.json({ ok: true, clientId: req.params.clientId, message: 'Vector + prompt + response cache invalidated' })
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
    const result   = await database.collection('conversations').findOneAndUpdate(
      { _id: new ObjectId(conversationId), clientId: req.client.clientId },
      { $set: { title: title.trim(), updatedAt: new Date() } },
      { returnDocument: 'after', projection: { messages: 0 } }
    )
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
    const { query, topK = 6, conversationId } = req.body
    if (!query?.trim()) return res.status(400).json({ error: 'query is required' })
    const { clientId, name } = req.client
    const intent = detectQueryIntent(query.trim())

    if (intent === 'greeting') {
      return res.json({
        answer:         "Hello! I'm your document assistant. Ask me anything about your data.",
        sources:        [],
        conversationId: conversationId || null,
        client:         { clientId, name },
      })
    }

    // ── Response cache ──────────────────────────────────────────────────────
    const cacheKey = getCacheKey(clientId, query)
    const cached   = responseCacheGet(cacheKey)
    if (cached) {
      console.log(`[cache] HIT for "${query.slice(0, 50)}"`)
      return res.json({ ...cached, cached: true, conversationId: conversationId || cached.conversationId })
    }

    // ── In-flight dedup ─────────────────────────────────────────────────────
    if (IN_FLIGHT.has(cacheKey)) {
      console.log(`[dedup] Waiting for in-flight: "${query.slice(0, 50)}"`)
      try {
        const result = await IN_FLIGHT.get(cacheKey)
        return res.json({ ...result, conversationId: conversationId || result.conversationId })
      } catch {}
    }

    const requestPromise = (async () => {
      // Load pre-embedded vectors from blob (cached after first call)
      const { chunks, invertedIndex } = await loadChunksForClient(clientId)

      if (chunks.length === 0) {
        return {
          answer:         'No documents found for your account. Please ensure your documents have been ingested first.',
          sources:        [],
          conversationId: conversationId || null,
          client:         { clientId, name },
        }
      }

      // Cosine similarity retrieval (primary) + keyword (secondary)
      const hits = await retrieveChunks(query.trim(), chunks, Math.min(topK, 20), invertedIndex)

      if (hits.length === 0) {
        return {
          answer:         "I couldn't find that in your documents. Try rephrasing your question.",
          sources:        [],
          conversationId: conversationId || null,
          client:         { clientId, name },
        }
      }

      // Phi-4 answer generation
      let rawAnswer
      try {
        rawAnswer = await Promise.race([
          answerWithPhi4(query.trim(), hits, intent),
          new Promise((_, reject) =>
            setTimeout(() => reject(new Error('Model response timeout (8s)')), 8000)
          ),
        ])
      } catch (err) {
        console.warn(`⚠️ [phi4] Fallback answer: ${err.message}`)
        rawAnswer = buildFallbackAnswer(query.trim(), hits)
      }

      const cleanAnswer = fixBrokenUrls(rawAnswer)
        .replace(/\bField\d+\s*:\s*/gi, '')
        .replace(/\|\s*Field\d+\b/gi, '')
        .trim()

      const sources = hits.map(h => ({
        source_file: h.source_file  || 'unknown',
        chunk_index: h.chunk_index  ?? 0,
        score:       typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
        cosine:      typeof h._cosine === 'number' ? parseFloat(h._cosine.toFixed(4)) : null,
        preview:     (h.text || '').slice(0, 300),
      }))

      // Persist conversation to MongoDB
      let activeConversationId = conversationId || null
      try {
        const chatDatabase = await getChatDb()
        const col          = chatDatabase.collection('conversations')
        const now          = new Date()
        const userMsg      = { role: 'user',      content: query.trim(), timestamp: now }
        const assistantMsg = {
          role:      'assistant',
          content:   cleanAnswer,
          sources:   sources.map(s => ({ source_file: s.source_file, score: s.score })),
          timestamp: now,
        }
        if (activeConversationId) {
          const updated = await col.findOneAndUpdate(
            { _id: new ObjectId(activeConversationId), clientId },
            { $push: { messages: { $each: [userMsg, assistantMsg] } }, $set: { updatedAt: now } },
            { returnDocument: 'after', projection: { _id: 1 } }
          )
          if (!updated) activeConversationId = null
        }
        if (!activeConversationId) {
          const result = await col.insertOne({
            clientId,
            title:     generateTitle(query.trim()),
            messages:  [userMsg, assistantMsg],
            createdAt: now,
            updatedAt: now,
          })
          activeConversationId = result.insertedId.toString()
        }
      } catch (saveErr) {
        console.warn('[chat/message] Failed to save conversation:', saveErr.message)
      }

      return { answer: cleanAnswer, sources, conversationId: activeConversationId, client: { clientId, name } }
    })()

    IN_FLIGHT.set(cacheKey, requestPromise)
    let result
    try { result = await requestPromise }
    finally { IN_FLIGHT.delete(cacheKey) }

    if (result.answer && result.answer.length > 10) responseCacheSet(cacheKey, result)

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
  console.log(`Retrieval mode : MiniLM vector cosine search`)
  console.log(`MiniLM sidecar : ${MINILM_SIDECAR_URL}`)
  console.log(`Phi-4 model    : ${PHI4_MODEL} | Endpoint: ${PHI4_ENDPOINT ? 'configured' : 'MISSING'}`)
  console.log(`Phi-4 timeout  : ${PHI4_TIMEOUT_MS}ms`)
  console.log(`Embed timeout  : ${EMBED_TIMEOUT_MS}ms`)
  console.log(`Vector cache TTL: ${CHUNK_CACHE_TTL}ms`)
  console.log(`Blob concurrency: ${BLOB_CONCURRENCY}`)
  console.log(`Azure blob     : ${blobServiceClient ? 'singleton ready' : 'MISSING connection string'}`)
  startApiKeyHealthChecker()
  warmupChunkCaches()
})
module.exports = app
