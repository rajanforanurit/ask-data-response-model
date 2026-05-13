require('dotenv').config()
const express = require('express')
const cors = require('cors')
const { MongoClient, ObjectId } = require('mongodb')
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
  'https://www.anuritinnovation.com/',
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

// ── Environment variables ──────────────────────────────────────────────────────
const MONGODB_URI             = process.env.MONGODB_URI
const MONGODB_DB              = process.env.MONGODB_DB || 'clientcreds'
const CHAT_HISTORY_URI        = process.env.CHAT_HISTORY_URI
const CHAT_HISTORY_DB         = process.env.CHAT_HISTORY_DB || 'chathistory'
const ADMIN_API_KEY           = process.env.ADMIN_API_KEY
const KEY_CHECK_INTERVAL_MS   = parseInt(process.env.KEY_CHECK_INTERVAL_MS || '300000', 10)
const REQUEST_TIMEOUT_MS      = parseInt(process.env.REQUEST_TIMEOUT_MS || '60000', 10)
const PHI4_ENDPOINT           = process.env.PHI4_ENDPOINT
const PHI4_API_KEY            = process.env.PHI4_API_KEY
const PHI4_MODEL              = process.env.PHI4_MODEL || 'Phi-4'
const PHI4_TIMEOUT_MS         = parseInt(process.env.PHI4_TIMEOUT_MS || '30000', 10)
const EMBED_TIMEOUT_MS        = parseInt(process.env.EMBED_TIMEOUT_MS || '10000', 10)

// ── NEW: Python embed endpoint (replaces Azure OpenAI embeddings) ──────────────
const PYTHON_EMBED_ENDPOINT   = process.env.PYTHON_EMBED_ENDPOINT || ''
const PYTHON_EMBED_API_KEY    = process.env.PYTHON_EMBED_API_KEY || ''

// ── Azure AI Search ────────────────────────────────────────────────────────────
const AZURE_SEARCH_ENDPOINT   = process.env.AZURE_SEARCH_ENDPOINT || ''
const AZURE_SEARCH_KEY        = process.env.AZURE_SEARCH_KEY || ''
const AZURE_SEARCH_INDEX      = process.env.AZURE_SEARCH_INDEX || 'rag-chunks'
const SEARCH_API_VERSION      = '2024-11-01-preview'

// ── Response cache ─────────────────────────────────────────────────────────────
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

// ── Phi-4 concurrency limiter ──────────────────────────────────────────────────
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
      } else { phiQueue.push(tryRun) }
    }
    tryRun()
  })
}
function drainPhiQueue() {
  if (phiQueue.length > 0 && phiActiveCount < PHI_MAX_CONCURRENT) phiQueue.shift()()
}

// ── Phi-4 circuit breaker ──────────────────────────────────────────────────────
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
  if (phiFailures >= 3) { phiBlockedUntil = Date.now() + 30000; console.error('[phi4] Circuit breaker OPEN for 30s') }
}

// ── Utilities ──────────────────────────────────────────────────────────────────
async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    return await fetch(url, { ...options, signal: controller.signal })
  } catch (err) {
    if (err.name === 'AbortError') throw new Error(`Request timed out after ${timeoutMs}ms`)
    throw err
  } finally { clearTimeout(timer) }
}

function withRequestTimeout(fn, timeoutMs = REQUEST_TIMEOUT_MS) {
  return async (req, res, next) => {
    let settled = false
    const timer = setTimeout(() => {
      if (!settled) { settled = true; if (!res.headersSent) res.status(503).json({ error: 'Request timed out. Please try again.' }) }
    }, timeoutMs)
    try { await fn(req, res, next) } catch (err) { if (!settled) next(err) } finally { settled = true; clearTimeout(timer) }
  }
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

function fixBrokenUrls(text) {
  return text.replace(/https:\/\/[^\s]+(\s+[^\s]+)/g, (match) => match.replace(/\s/g, ''))
}

function ensureSinglePeriod(text) {
  if (!text) return ''
  return text.replace(/\.{2,}/g, '.').replace(/\.\s*\./g, '.').trim()
}

function capFirst(str) {
  if (!str) return ''
  return str.charAt(0).toUpperCase() + str.slice(1)
}

function generateTitle(query) {
  const cleaned = query.trim().replace(/[?!.]+$/, '')
  return cleaned.length > 50 ? cleaned.slice(0, 50) + '…' : cleaned
}

// ── Embeddings: calls Python ingestion API (384-dim MiniLM, matches index) ─────
async function embedQueryAzure(query) {
  if (!PYTHON_EMBED_ENDPOINT) return null
  try {
    const response = await fetchWithTimeout(
      PYTHON_EMBED_ENDPOINT,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${PYTHON_EMBED_API_KEY}`,
        },
        body: JSON.stringify({ text: query }),
      },
      EMBED_TIMEOUT_MS
    )
    if (!response.ok) {
      console.error('[embed] Python embed API error:', response.status)
      return null
    }
    const data = await response.json()
    return data.embedding || null
  } catch (err) {
    console.error('[embed] Exception:', err.message)
    return null
  }
}

// ── Azure AI Search ────────────────────────────────────────────────────────────
async function searchChunks(clientId, query, topK = 6) {
  if (!AZURE_SEARCH_ENDPOINT || !AZURE_SEARCH_KEY) {
    throw new Error('AZURE_SEARCH_ENDPOINT and AZURE_SEARCH_KEY are not configured')
  }
  const url = `${AZURE_SEARCH_ENDPOINT}/indexes/${AZURE_SEARCH_INDEX}/docs/search?api-version=${SEARCH_API_VERSION}`
  const headers = {
    'Content-Type': 'application/json',
    'api-key': AZURE_SEARCH_KEY,
  }
  const queryVec = await embedQueryAzure(query)
  const searchBody = {
    search: query,
    filter: `client_id eq '${clientId}'`,
    select: 'text,source_file,chunk_index,page,doc_id',
    top: topK,
    queryType: 'semantic',
    semanticConfiguration: 'semantic-config',
  }
  if (queryVec) {
    searchBody.vectorQueries = [{
      kind: 'vector',
      vector: queryVec,
      fields: 'embedding',
      k: topK * 3,
      exhaustive: false,
    }]
  }
  const resp = await fetchWithTimeout(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(searchBody),
  }, 15000)
  if (!resp.ok) {
    const errText = await resp.text()
    throw new Error(`Azure AI Search error ${resp.status}: ${errText}`)
  }
  const data = await resp.json()
  return (data.value || []).map(doc => ({
    text:        doc.text || '',
    source_file: doc.source_file || 'unknown',
    chunk_index: doc.chunk_index ?? 0,
    page:        doc.page ?? 1,
    doc_id:      doc.doc_id || '',
    _score:      doc['@search.rerankerScore'] ?? doc['@search.score'] ?? 0,
  }))
}

// ── Phi-4 LLM ─────────────────────────────────────────────────────────────────
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
              { role: 'user',   content: userMessage },
            ],
            temperature: 0.1,
            max_tokens: maxTokens,
          }),
        },
        PHI4_TIMEOUT_MS
      )
      if (!response.ok) { const errText = await response.text(); throw new Error(`Phi-4 API error ${response.status}: ${errText}`) }
      const data = await response.json()
      phiRecordSuccess()
      return data.choices?.[0]?.message?.content || ''
    } catch (err) { phiRecordFailure(); throw err }
  })
}

// ── Context & prompt builders ──────────────────────────────────────────────────
function buildContext(hits) {
  const seen = new Set()
  const deduped = []
  for (const h of hits) {
    const fp = `${h.source_file || ''}::${h.chunk_index ?? h.text?.slice(0, 40)}`
    if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
    if (deduped.length >= 6) break
  }
  return deduped.map((h, i) => {
    const limit = i === 0 ? 1200 : 900
    return `[Source ${i + 1}]\n${(h.text || '').trim().slice(0, limit)}`
  }).join('\n\n---\n\n')
}

function buildSystemPrompt() {
  return `You are a helpful, versatile data assistant. You can answer questions about any topic found in the provided documents — real estate, finance, healthcare, logistics, retail, or any other domain.
Answer ONLY using the provided context. Follow these STRICT rules:
FORMATTING RULES:
1. Bold the main subject using **Name**.
2. Write in complete, natural English sentences only. Never output raw pipe-separated data.
3. Never end a sentence with double periods. Use only one period.
4. Remove all sheet references like "=== Sheet: X ===" from your output.
5. Do NOT include source references like [1], [2].
6. For definitions: bold the name, state the definition clearly.
7. For formulas: bold "Formula:" and write it plainly.
8. For URL lookups: output only the full URL on its own line.
9. Keep answers concise: 2-5 sentences unless the data requires more.
10. If the context does not contain the answer, say exactly: "I could not find this in your documents."
TABULAR DATA RULES:
11. When context contains rows of structured data (e.g. "City: X, Country: Y" or CSV-like rows), treat each as a data record.
12. For lookup questions ("what country is Karachi in"), find the matching row and state the value directly.
13. For count questions ("how many cities"), count distinct rows and state the number.
14. For ranking questions ("which city has highest visitors"), compare values across all rows and name the winner with its value.
15. For aggregation ("total visitors in India"), sum values from all matching rows and state the total.
16. NEVER say you cannot find information if matching rows are clearly present in the context.
17. For ranking or comparison, scan ALL provided sources before answering — the best match may not be in the first source.`
}

function buildUserMessage(query, hits) {
  const context = buildContext(hits)
  return `CONTEXT:\n${context}\n\nUsing ONLY the context above, answer this question clearly and directly: ${query}`
}

function buildFallbackAnswer(query, hits) {
  if (!hits || hits.length === 0) return "I could not find relevant information about this in your documents."
  const intent = detectQueryIntent(query)
  const qLower = query.toLowerCase()
  if (intent === 'url_lookup') {
    const urlRegex = /https?:\/\/[^\s"'<>]+/
    for (const h of hits) {
      const urlMatch = (h.text || '').match(urlRegex)
      if (urlMatch) return urlMatch[0].replace(/[.,;)]+$/, '').trim()
    }
    return "I could not find a matching URL in your documents."
  }
  const rowPattern = /([A-Za-z\s#.()\/]+):\s*([^,\n]+)/g
  const allRows = []
  for (const h of hits) {
    const lines = (h.text || '').split('\n')
    for (const line of lines) {
      if (line.trim().length < 10) continue
      const fields = {}
      let m
      while ((m = rowPattern.exec(line)) !== null) {
        fields[m[1].trim().toLowerCase()] = m[2].trim()
      }
      rowPattern.lastIndex = 0
      if (Object.keys(fields).length >= 2) allRows.push({ line, fields })
    }
  }
  if (allRows.length > 0) {
    const queryWords = qLower.replace(/[?!.,]/g, '').split(/\s+/).filter(w => w.length > 2)
    const matchingRows = allRows.filter(r =>
      queryWords.some(w => r.line.toLowerCase().includes(w))
    )
    const sourceRows = matchingRows.length > 0 ? matchingRows : allRows.slice(0, 3)
    for (const { fields } of sourceRows) {
      if (/country/.test(qLower) && fields['country']) {
        const city = fields['city'] || ''
        return ensureSinglePeriod(`**${city}** is located in **${fields['country']}**.`)
      }
      if (/continent/.test(qLower) && fields['continent']) {
        const city = fields['city'] || ''
        return ensureSinglePeriod(`**${city}** is in **${fields['continent']}**.`)
      }
      if (/region/.test(qLower) && fields['region']) {
        const city = fields['city'] || ''
        return ensureSinglePeriod(`**${city}** belongs to the **${fields['region']}** region.`)
      }
      if (/visitor/.test(qLower) && (fields['number of unique visitors'] || fields['number of visitors (rounded)'])) {
        const city = fields['city'] || ''
        const val = fields['number of unique visitors'] || fields['number of visitors (rounded)']
        return ensureSinglePeriod(`**${city}** has **${val}** unique visitors.`)
      }
    }
    if (matchingRows.length > 0) {
      const best = matchingRows[0]
      const readable = Object.entries(best.fields)
        .map(([k, v]) => `${capFirst(k)}: ${v}`)
        .join(', ')
      return ensureSinglePeriod(readable + '.')
    }
  }
  return "I could not find that specific information in your documents."
}

function cleanAnswer(rawAnswer) {
  if (!rawAnswer) return ''
  let cleaned = fixBrokenUrls(rawAnswer)
    .replace(/^\s*\[Source\s*\d+\]\s*/gm, '')
    .replace(/^[^\n]*(\|[^\n]*){3,}$/gm, '')
    .replace(/=== .+ ===\s*/gm, '')
    .replace(/\(from\s+[A-Za-z\s]+\)\s*/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\.{2,}/g, '.')
    .replace(/\.\s*\./g, '.')
    .trim()
  if (cleaned.length > 0 && !/[.!?]$/.test(cleaned)) cleaned += '.'
  return ensureSinglePeriod(cleaned)
}

// ── MongoDB helpers ────────────────────────────────────────────────────────────
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

// ── API key cache ──────────────────────────────────────────────────────────────
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

// ── Routes ─────────────────────────────────────────────────────────────────────

// FIX: health now correctly reports embedding status based on PYTHON_EMBED_ENDPOINT
app.get('/health', (req, res) => res.json({
  ok: true,
  service: 'ask-data',
  model: PHI4_MODEL,
  vectorStore: 'Azure AI Search',
  searchConfigured: !!(AZURE_SEARCH_ENDPOINT && AZURE_SEARCH_KEY),
  embeddings: PYTHON_EMBED_ENDPOINT ? 'python-minilm' : 'none',
  embeddingEndpoint: PYTHON_EMBED_ENDPOINT ? PYTHON_EMBED_ENDPOINT : 'not configured',
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

app.post('/client/login', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ error: 'Invalid API key' })
    res.json({ ok: true, client })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/login', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ error: 'Invalid API key' })
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

app.post('/admin/clients', requireAdminKey, async (req, res) => {
  try {
    let { name, clientId, apiKey } = req.body
    if (!name || !clientId) return res.status(400).json({ error: 'name and clientId are required' })
    if (!apiKey) { apiKey = generateApiKey() }
    else if (!apiKey.startsWith('rak_')) { return res.status(400).json({ error: 'apiKey must start with "rak_"' }) }
    const database = await getDb()
    const col = database.collection('clients')
    const existing = await col.findOne({ $or: [{ clientId }, { apiKey }] })
    if (existing) {
      const field = existing.clientId === clientId ? 'clientId' : 'apiKey'
      return res.status(409).json({ error: `A client with this ${field} already exists` })
    }
    const now = new Date().toISOString()
    const doc = { name: name.trim(), clientId: clientId.trim().toLowerCase(), apiKey, apiKeyRotatedAt: now, status: 'idle', createdAt: now, updatedAt: now }
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
    RESPONSE_CACHE.forEach((_, key) => { if (key.startsWith(clientId + ':')) RESPONSE_CACHE.delete(key) })
    res.json({ ok: true, deleted: clientId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/admin/clients/:clientId/invalidate-cache', requireAdminKey, (req, res) => {
  const { clientId } = req.params
  RESPONSE_CACHE.forEach((_, key) => { if (key.startsWith(clientId + ':')) RESPONSE_CACHE.delete(key) })
  res.json({ ok: true, clientId, message: 'Response cache invalidated' })
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
    const result = await database.collection('conversations').findOneAndUpdate(
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
      return res.json({ answer: validation.message, sources: [], conversationId: conversationId || null, client: req.client })
    }
    const { clientId, name } = req.client
    const intent = detectQueryIntent(query.trim())
    if (intent === 'greeting') {
      return res.json({
        answer: "Hello! I'm your data assistant. Ask me anything about the information in your documents.",
        sources: [],
        conversationId: conversationId || null,
        client: { clientId, name },
      })
    }
    const cacheKey = getCacheKey(clientId, query)
    const cached = responseCacheGet(cacheKey)
    if (cached) return res.json({ ...cached, cached: true, conversationId: conversationId || cached.conversationId })
    if (IN_FLIGHT.has(cacheKey)) {
      try {
        const result = await IN_FLIGHT.get(cacheKey)
        return res.json({ ...result, conversationId: conversationId || result.conversationId })
      } catch { }
    }
    const requestPromise = (async () => {
      let hits = []
      try {
        hits = await searchChunks(clientId, query.trim(), Math.min(topK, 8))
        console.log(`[chat/message] "${query.slice(0, 60)}" → intent=${intent}, hits=${hits.length}, topScore=${hits[0]?._score?.toFixed(3) || 0}`)
      } catch (searchErr) {
        console.error('[chat/message] Azure AI Search failed:', searchErr.message)
        const isDev = process.env.NODE_ENV !== 'production'
        return {
          answer: isDev
            ? `Search error: ${searchErr.message}`
            : 'The search service is temporarily unavailable. Please try again in a moment.',
          sources: [],
          conversationId: conversationId || null,
          client: { clientId, name },
        }
      }
      if (hits.length === 0) {
        return {
          answer: "No relevant information was found for your question. Please make sure your documents have been uploaded and indexed, or try rephrasing your question.",
          sources: [],
          conversationId: conversationId || null,
          client: { clientId, name },
        }
      }
      let rawAnswer = ''
      try {
        rawAnswer = await Promise.race([
          callPhi4(buildSystemPrompt(), buildUserMessage(query.trim(), hits), 1024),
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
        score:       typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
        preview:     (h.text || '').slice(0, 200),
      }))
      let activeConversationId = conversationId || null
      try {
        const chatDatabase = await getChatDb()
        const col = chatDatabase.collection('conversations')
        const now = new Date()
        const userMsg      = { role: 'user',      content: query.trim(), timestamp: now }
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
      } catch (saveErr) { console.warn('[chat/message] Failed to save conversation:', saveErr.message) }
      return { answer, sources, conversationId: activeConversationId, client: { clientId, name } }
    })()
    IN_FLIGHT.set(cacheKey, requestPromise)
    let result
    try { result = await requestPromise } finally { IN_FLIGHT.delete(cacheKey) }
    if (result.answer && result.answer.length > 15) responseCacheSet(cacheKey, result)
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
  console.log(`Model: ${PHI4_MODEL}`)
  console.log(`Vector store: Azure AI Search — ${AZURE_SEARCH_ENDPOINT || 'NOT CONFIGURED'}`)
  console.log(`Search index: ${AZURE_SEARCH_INDEX}`)
  console.log(`Embeddings: ${PYTHON_EMBED_ENDPOINT ? `Python MiniLM — ${PYTHON_EMBED_ENDPOINT}` : 'DISABLED — keyword+BM25 only'}`)
  startApiKeyHealthChecker()
})

module.exports = app
