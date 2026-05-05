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
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
}))

app.options('*', cors({
  origin: (origin, callback) => callback(null, true),
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
}))

app.use(express.json())

// ─── ENV ──────────────────────────────────────────────────────────────────────
const MONGODB_URI            = process.env.MONGODB_URI
const MONGODB_DB             = process.env.MONGODB_DB || 'clientcreds'
const CHAT_HISTORY_URI       = process.env.CHAT_HISTORY_URI
const CHAT_HISTORY_DB        = process.env.CHAT_HISTORY_DB || 'chathistory'
const AZURE_CONNECTION_STRING = process.env.AZURE_CONNECTION_STRING || ''
const AZURE_CONTAINER_NAME   = process.env.AZURE_CONTAINER_NAME || 'vectordbforrag'
const ADMIN_API_KEY          = process.env.ADMIN_API_KEY
const KEY_CHECK_INTERVAL_MS  = parseInt(process.env.KEY_CHECK_INTERVAL_MS || '300000', 10)

// Azure AI Foundry / Phi-4 config
const PHI4_ENDPOINT  = process.env.PHI4_ENDPOINT  // e.g. https://xxx.services.ai.azure.com/models/chat/completions?api-version=2024-05-01-preview
const PHI4_API_KEY   = process.env.PHI4_API_KEY   // the key from Azure AI Foundry
const PHI4_MODEL     = process.env.PHI4_MODEL || 'Phi-4-mini-instruct'

// Optional: Azure OpenAI embeddings (text-embedding-ada-002 or similar)
// If not set, falls back to keyword-only search
const AZURE_EMBED_ENDPOINT = process.env.AZURE_EMBED_ENDPOINT || ''
const AZURE_EMBED_KEY      = process.env.AZURE_EMBED_KEY || ''
const AZURE_EMBED_MODEL    = process.env.AZURE_EMBED_MODEL || 'text-embedding-ada-002'

const RAW_PREFIX   = 'raw'
const CHUNK_SIZE   = 500
const CHUNK_OVERLAP = 2

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

// ─── Helpers ──────────────────────────────────────────────────────────────────
function classifyExtension(fileName) {
  const ext = ('.' + fileName.split('.').pop()).toLowerCase()
  if (['.xlsx', '.xls', '.ods'].includes(ext))                              return DOC_TYPE.SPREADSHEET
  if (ext === '.pdf')                                                        return DOC_TYPE.PDF
  if (['.docx', '.doc', '.odt', '.rtf'].includes(ext))                      return DOC_TYPE.WORD
  if (['.pptx', '.ppt'].includes(ext))                                      return DOC_TYPE.PRESENTATION
  if (['.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h','.cs',
       '.go','.rb','.php','.swift','.kt','.r','.sql','.sh','.bash','.ps1']
      .includes(ext))                                                        return DOC_TYPE.CODE
  if (['.json','.jsonl','.yaml','.yml','.toml','.csv','.tsv'].includes(ext)) return DOC_TYPE.DATA
  if (['.txt','.md','.markdown','.rst'].includes(ext))                      return DOC_TYPE.TEXT
  if (ext === '.eml')                                                        return DOC_TYPE.EMAIL
  if (['.html','.htm','.xml'].includes(ext))                                return DOC_TYPE.WEB
  return DOC_TYPE.UNKNOWN
}

function inferSchema(fileName, textSamples) {
  const type = classifyExtension(fileName)
  const schema = { type, fileName, columns: [], sampleValues: [], topics: [] }
  if (type === DOC_TYPE.SPREADSHEET || type === DOC_TYPE.DATA) {
    const columnSet = new Set(), valueSet = new Set()
    for (const sample of textSamples.slice(0, 60)) {
      const pairs = sample.split('|').map(s => s.trim())
      for (const pair of pairs) {
        const colonIdx = pair.indexOf(':')
        if (colonIdx > 0) {
          const key = pair.slice(0, colonIdx).trim()
          const val = pair.slice(colonIdx + 1).trim()
          if (key && key.length < 80)  columnSet.add(key)
          if (val && val.length < 120) valueSet.add(val)
        }
      }
    }
    schema.columns      = [...columnSet].slice(0, 30)
    schema.sampleValues = [...valueSet].slice(0, 20)
  } else {
    const freq = {}
    for (const sample of textSamples.slice(0, 30)) {
      for (const word of sample.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)) {
        if (word.length > 5) freq[word] = (freq[word] || 0) + 1
      }
    }
    schema.topics = Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 15).map(([w]) => w)
  }
  return schema
}

function detectQueryIntent(query) {
  const q = query.toLowerCase().trim()
  const DEFINITION_PATTERNS = [/^what\s+(is|are|does)\s+/,/^define\s+/,/^explain\s+/,/^meaning\s+of\s+/,/^tell\s+me\s+about\s+/,/^describe\s+/,/^how\s+is\s+.+\s+(calculated|defined|measured|computed)/,/\bmeaning\b/,/\bdefinition\b/,/\bwhat\s+does\b/]
  const LOOKUP_PATTERNS     = [/^(show|list|find|get|fetch|give)\s+(me\s+)?/,/^how\s+many\s+/,/^what\s+(is\s+the\s+)?(value|number|count|total|sum|amount)/]
  const COMPARISON_PATTERNS = [/\bvs\b|\bversus\b|\bdifference\b|\bcompare\b|\bbetween\b/]
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
  ]
  for (const p of patterns) {
    const m = q.match(p)
    if (m) return m[1].trim()
  }
  return q
}
function fixBrokenUrls(text) {
  return text
    .replace(/https:\/\/[^\s]+(\s+[^\s]+)/g, (match) => match.replace(/\s/g, ''))
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

// ─── Azure AI Foundry: Phi-4 chat completion ──────────────────────────────────
async function callPhi4(systemPrompt, userMessage) {
  if (!PHI4_ENDPOINT || !PHI4_API_KEY) {
    throw new Error('PHI4_ENDPOINT and PHI4_API_KEY environment variables are required')
  }

  const body = {
    model: PHI4_MODEL,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userMessage  },
    ],
    temperature: 0.2,
    max_tokens:  1024,
  }

  const response = await fetch(PHI4_ENDPOINT, {
    method:  'POST',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': `Bearer ${PHI4_API_KEY}`,
    },
    body: JSON.stringify(body),
  })

  if (!response.ok) {
    const errText = await response.text()
    throw new Error(`Phi-4 API error ${response.status}: ${errText}`)
  }

  const data = await response.json()
  return data.choices?.[0]?.message?.content || ''
}

// ─── Azure OpenAI Embeddings (optional) ──────────────────────────────────────
async function embedQueryAzure(query) {
  if (!AZURE_EMBED_ENDPOINT || !AZURE_EMBED_KEY) return null

  const response = await fetch(AZURE_EMBED_ENDPOINT, {
    method:  'POST',
    headers: {
      'Content-Type':  'application/json',
      'api-key':       AZURE_EMBED_KEY,
    },
    body: JSON.stringify({ input: query, model: AZURE_EMBED_MODEL }),
  })

  if (!response.ok) return null
  const data = await response.json()
  return data.data?.[0]?.embedding || null
}

// ─── Keyword search ───────────────────────────────────────────────────────────
function keywordSearch(query, chunks, topK, intent = 'general') {
  const subject      = intent === 'definition' ? extractSubject(query) : query.toLowerCase()
  const queryLower   = query.toLowerCase()
  const subjectLower = subject.toLowerCase()
  const subjectWords = subjectLower.split(/\s+/).filter(w => w.length > 1)
  const queryWords   = queryLower.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1)

  return chunks
    .map(c => {
      const text    = (c.text || '').toLowerCase()
      const docType = classifyExtension(c.source_file || '')
      let score = 0

      if (text.includes(subjectLower)) {
        score += subjectWords.length * 4
        const defPattern = new RegExp(`${escapeRegex(subjectLower)}\\s*(is|are)\\s*(defined|described|calculated|measured|computed)`, 'i')
        if (defPattern.test(c.text || '')) score += subjectWords.length * 6
      }

      const wordHits = subjectWords.filter(w => text.includes(w)).length
      score += wordHits * 2

      if ((docType === DOC_TYPE.SPREADSHEET || docType === DOC_TYPE.DATA) && intent === 'definition') {
        const descPattern = new RegExp(`${escapeRegex(subjectLower)}\\s*(is described as|is defined as):`, 'i')
        if (descPattern.test(c.text || '')) score += subjectWords.length * 8
      }

      if (docType === DOC_TYPE.SPREADSHEET || docType === DOC_TYPE.DATA) {
        for (const w of subjectWords) {
          const kvPattern = new RegExp(`:\\s*${escapeRegex(w)}\\b|\\|\\s*${escapeRegex(w)}\\b`, 'i')
          if (kvPattern.test(c.text || '')) score += 2
        }
      }

      if (text.includes(queryLower)) score += queryWords.length * 2
      return { ...c, _score: score }
    })
    .filter(c => c._score > 0)
    .sort((a, b) => b._score - a._score)
    .slice(0, topK)
}

// ─── Retrieval ────────────────────────────────────────────────────────────────
async function retrieveChunks(query, chunks, topK = 6) {
  const intent           = detectQueryIntent(query)
  const normalizedQuery  = query.toLowerCase().trim().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ')
  const keywordTopK      = intent === 'definition' ? Math.min(150, chunks.length) : Math.min(100, chunks.length)
  const candidates       = keywordSearch(normalizedQuery, chunks, keywordTopK, intent)
  const pool             = candidates.length > 0 ? candidates : chunks.slice(0, 100)

  // Fast path: strong keyword match on definition queries
  if (intent === 'definition' && pool.length > 0 && pool[0]._score >= 8) {
    return pool.slice(0, Math.min(topK, 10))
  }

  // Try semantic re-ranking via Azure embeddings
  if (AZURE_EMBED_ENDPOINT && AZURE_EMBED_KEY) {
    try {
      const queryVec = await embedQueryAzure(normalizedQuery)
      if (queryVec) {
        const scored = []
        for (const c of pool) {
          try {
            const chunkVec = await embedQueryAzure((c.text || '').toLowerCase().slice(0, 512))
            if (chunkVec) {
              const semanticScore = cosineSim(queryVec, chunkVec)
              const maxKeyword    = pool[0]._score || 1
              const keywordNorm   = typeof c._score === 'number' ? c._score / maxKeyword : 0
              const weight        = intent === 'definition'
                ? { semantic: 0.35, keyword: 0.65 }
                : { semantic: 0.70, keyword: 0.30 }
              scored.push({ ...c, _score: semanticScore * weight.semantic + keywordNorm * weight.keyword })
            } else {
              scored.push(c)
            }
          } catch {
            scored.push(c)
          }
        }
        return scored.sort((a, b) => b._score - a._score).slice(0, Math.min(topK, 12))
      }
    } catch (err) {
      console.warn('[retrieveChunks] Azure embed failed, keyword fallback:', err.message)
    }
  }

  return pool.slice(0, Math.min(topK, 12))
}

// ─── Context & prompt builders ────────────────────────────────────────────────
function buildContext(hits) {
  return hits.map((h, i) => {
    const src      = h.source_file || 'document'
    const docType  = classifyExtension(src)
    const typeLabel = {
      [DOC_TYPE.SPREADSHEET]:  'spreadsheet — pipe-separated key:value pairs, each line is one record',
      [DOC_TYPE.DATA]:         'structured data (JSON/YAML/CSV)',
      [DOC_TYPE.PDF]:          'PDF document',
      [DOC_TYPE.WORD]:         'Word document',
      [DOC_TYPE.PRESENTATION]: 'presentation slide',
      [DOC_TYPE.CODE]:         'source code',
      [DOC_TYPE.TEXT]:         'text/markdown document',
      [DOC_TYPE.EMAIL]:        'email',
      [DOC_TYPE.WEB]:          'web/HTML page',
      [DOC_TYPE.UNKNOWN]:      'document',
    }[docType] || 'document'
    return `[Excerpt ${i + 1} | type: ${typeLabel} | relevance: ${typeof h._score === 'number' ? h._score.toFixed(3) : 'n/a'}]\n${(h.text || '').trim()}`
  }).join('\n\n---\n\n')
}

function buildDynamicSystemPrompt(hits, intent = 'general') {
  const fileMap = new Map()
  for (const h of hits) {
    const src = h.source_file || 'unknown'
    if (!fileMap.has(src)) fileMap.set(src, [])
    fileMap.get(src).push((h.text || '').trim())
  }

  const schemas       = []
  for (const [fileName, samples] of fileMap) schemas.push(inferSchema(fileName, samples))
  const spreadsheets  = schemas.filter(s => s.type === DOC_TYPE.SPREADSHEET)
  const dataFiles     = schemas.filter(s => s.type === DOC_TYPE.DATA)
  const pdfDocs       = schemas.filter(s => s.type === DOC_TYPE.PDF)
  const wordDocs      = schemas.filter(s => s.type === DOC_TYPE.WORD)
  const presentations = schemas.filter(s => s.type === DOC_TYPE.PRESENTATION)
  const codeFiles     = schemas.filter(s => s.type === DOC_TYPE.CODE)
  const textFiles     = schemas.filter(s => s.type === DOC_TYPE.TEXT)
  const emailFiles    = schemas.filter(s => s.type === DOC_TYPE.EMAIL)
  const webFiles      = schemas.filter(s => s.type === DOC_TYPE.WEB)

  const intentInstructions = {
    definition: `ANSWER STRATEGY — DEFINITION QUERY: The user is asking for the definition or meaning of a specific term or metric. 1. Scan ALL excerpts for: the EXACT term, "is defined as", "is described as", "is calculated as", or any sentence that explains what the term IS. 2. If found, state the definition clearly and completely. Include calculation logic, filters, or conditions if mentioned. 3. If the term appears as a column name or value in structured data, explain its role in that context. 4. Do NOT describe tangentially related metrics — focus on the EXACT term asked about. 5. If multiple excerpts define the same term differently, reconcile them or present both definitions.`,
    lookup:     `ANSWER STRATEGY — LOOKUP QUERY: The user wants a specific value, count, or list from the data. 1. Find the exact records, rows, or values that match the query. 2. Report the precise values — do not approximate. 3. For spreadsheet data, scan all rows; the answer may span multiple records. 4. State clearly where the data comes from (metric name, column name, etc.).`,
    comparison: `ANSWER STRATEGY — COMPARISON QUERY: The user wants to compare two or more items. 1. Find all relevant information for EACH item being compared. 2. Structure the answer as a clear comparison — similarities and differences. 3. Use parallel structure so the comparison is easy to follow. 4. If one side has more data than the other, note the gap explicitly.`,
    general:    `ANSWER STRATEGY — GENERAL QUERY: Scan all excerpts carefully. Find information that directly answers the question. Synthesise a clear, complete answer. If the information spans multiple excerpts, combine it coherently.`,
  }[intent] || ''

  const base = `You are a knowledgeable document assistant. Answer questions ONLY using the document context provided.
UNIVERSAL RULES:
1. Answer ONLY from the context. Never invent or assume information.
2. Search the ENTIRE context — every excerpt — before concluding something is absent.
3. Case-insensitive matching: "applicant count", "Applicant Count", "APPLICANT COUNT" are identical.
4. If a term appears ANYWHERE in the context — as a label, value, heading, or inline text — treat it as present.
5. If information is truly absent after thorough search, say: "I couldn't find specific information about that in your documents."
6. Do NOT add citation markers like [1], [2], [3].
7. Do NOT mention file names or source document names in your answer.
8. Write clearly, concisely, and directly — like a knowledgeable colleague.
9. Answer only what was asked. No padding or filler.
10. NEVER say "the context does not define" or "not mentioned" if the term appears anywhere.
11. If a URL is present, ALWAYS return the FULL URL exactly as-is. Never truncate or shorten URLs.
12. URLs must be returned in one continuous line with no line breaks or spaces.
${intentInstructions}`

  const typeBlocks = []
  if (spreadsheets.length > 0) {
    const colSummary = spreadsheets.filter(s => s.columns.length > 0).map(s => ` • ${s.fileName}: [${s.columns.join(', ')}]`).join('\n')
    typeBlocks.push(`SPREADSHEET RULES: - Data is serialised as pipe-delimited key:value rows. Each line = one record. - Lines like "X is described as: ..." are definition summaries — prioritise them for definition queries. - For definition queries: a field name matching the subject IS a definition. Explain it from surrounding values. - Scan ALL rows — the answer may not be in the first matching row. ${colSummary ? `- Detected columns:\n${colSummary}` : ''}`)
  }
  if (dataFiles.length > 0)     typeBlocks.push(`STRUCTURED DATA RULES (JSON/YAML/CSV): - Fields and values may be nested. Treat "parent.child: value" as a nested attribute. - Every key and every value is meaningful data — no prose definition required.`)
  if (pdfDocs.length > 0)       typeBlocks.push(`PDF RULES: - Content is extracted from PDF pages. Minor formatting artefacts may exist. - Read numbers, dates, and figures exactly as they appear.`)
  if (wordDocs.length > 0)      typeBlocks.push(`WORD DOCUMENT RULES: - Context contains prose, lists, and tables. Headings indicate section structure. - Quote definitions or policy statements accurately.`)
  if (presentations.length > 0) typeBlocks.push(`PRESENTATION RULES: - Slide titles are section headers; bullets are supporting detail. - Do not infer beyond what the slide explicitly states.`)
  if (codeFiles.length > 0)     typeBlocks.push(`CODE RULES: - Read code literally. Function/variable names and comments are all meaningful. - Describe what code does in plain English unless code output is requested.`)
  if (textFiles.length > 0)     typeBlocks.push(`TEXT/MARKDOWN RULES: - Markdown formatting (##, **, -) indicates structure. Interpret accordingly. - Lists represent discrete facts or steps.`)
  if (emailFiles.length > 0)    typeBlocks.push(`EMAIL RULES: - Attribute statements to their sender. Do not mix up correspondents. - Dates and times are as stated in the email header.`)
  if (webFiles.length > 0)      typeBlocks.push(`WEB/HTML RULES: - Focus on main body content. Ignore repetitive navigation text. - Include URLs exactly if mentioned.`)

  const uniqueTypes = [...new Set(schemas.map(s => s.type))]
  const mixedNote   = uniqueTypes.length > 1 ? `\nMIXED DOCUMENT SET: Context contains ${uniqueTypes.length} document types (${uniqueTypes.join(', ')}). Apply the relevant rules above for each excerpt.` : ''
  return [base, ...typeBlocks, mixedNote].filter(Boolean).join('\n') + '\n'
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

const CLIENT_CACHE   = new Map()
const CACHE_TTL_MS   = 5 * 60 * 1000

function getCached(apiKey) {
  const entry = CLIENT_CACHE.get(apiKey)
  if (!entry) return null
  if (Date.now() - entry.cachedAt > CACHE_TTL_MS) { CLIENT_CACHE.delete(apiKey); return null }
  return entry
}
function setCache(apiKey, data) { CLIENT_CACHE.set(apiKey, { ...data, cachedAt: Date.now() }) }
function evictCache(apiKey)    { if (apiKey) CLIENT_CACHE.delete(apiKey) }

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
      const pairs = [], values = []
      for (let j = 0; j < Math.max(headers.length, row.length); j++) {
        const val = String(row[j] || '').trim()
        if (!val) continue
        const key = headers[j] && headers[j] !== '' ? headers[j] : `Field${j + 1}`
        pairs.push(`${key}: ${val}`)
        values.push(val)
      }
      if (pairs.length > 0) {
        parts.push(pairs.join(' | '))
        if (values.length >= 2) parts.push(`${values[0]} is described as: ${pairs.slice(1).join(', ')}`)
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
  if (ext === '.pdf')                          return extractPdf(buffer)
  if (ext === '.docx' || ext === '.doc')       return extractWord(buffer)
  if (ext === '.odt'  || ext === '.rtf')       return extractOffice(buffer)
  if (['.xlsx','.xls','.ods'].includes(ext))   return extractSpreadsheet(buffer)
  if (ext === '.csv')                          return extractCsv(buffer, ',')
  if (ext === '.tsv')                          return extractCsv(buffer, '\t')
  if (ext === '.pptx' || ext === '.ppt')       return extractOffice(buffer)
  if (ext === '.html' || ext === '.htm')       return extractHtml(buffer)
  if (ext === '.xml')                          return extractXml(buffer)
  if (['.md','.markdown','.rst'].includes(ext)) return buffer.toString('utf-8')
  if (ext === '.json')                         return extractJson(buffer)
  if (ext === '.jsonl')                        return extractJsonl(buffer)
  if (ext === '.yaml' || ext === '.yml')       return extractYaml(buffer)
  if (ext === '.toml')                         return buffer.toString('utf-8')
  const plainText = new Set(['.txt','.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h','.cs','.go','.rb','.php','.swift','.kt','.r','.sql','.sh','.bash','.ps1'])
  if (plainText.has(ext))                      return buffer.toString('utf-8')
  if (ext === '.epub')                         return extractEpub(buffer)
  if (ext === '.eml')                          return extractEml(buffer)
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

async function loadChunksForClient(clientId) {
  if (!AZURE_CONNECTION_STRING) throw new Error('AZURE_CONNECTION_STRING not set')
  const containerClient = BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING).getContainerClient(AZURE_CONTAINER_NAME)
  const prefix          = `${RAW_PREFIX}/${clientId}/`
  const allChunks       = []
  for await (const blob of containerClient.listBlobsFlat({ prefix })) {
    const fileName = blob.name.split('/').pop()
    const ext      = ('.' + fileName.split('.').pop()).toLowerCase()
    if (!SUPPORTED_EXTENSIONS.has(ext)) continue
    try {
      const buffer = await downloadBlobAsBuffer(containerClient, blob.name)
      const text   = await extractTextFromBuffer(buffer, fileName)
      if (!text?.trim()) continue
      allChunks.push(...chunkText(text, fileName))
    } catch (err) {
      console.warn(`[loadChunks] Failed to process ${blob.name}:`, err.message)
    }
  }
  return allChunks
}

// ─── Answer generation using Phi-4 ───────────────────────────────────────────
async function answerWithPhi4(originalQuery, hits, intent = 'general') {
  const systemPrompt = buildDynamicSystemPrompt(hits, intent)
  const context      = buildContext(hits)
  const subjectHint  = intent === 'definition'
    ? `\nThe user is specifically asking for the DEFINITION of: "${extractSubject(originalQuery)}". Focus entirely on that term.`
    : ''

  const userMessage = `---DOCUMENT CONTEXT START---
${context}
---DOCUMENT CONTEXT END---
${subjectHint}
User question: "${originalQuery}"
Instructions: Scan the entire context above. Apply the document-type rules. Give a direct, complete answer. Do not say a term is missing if it appears anywhere in the context.`

  return callPhi4(systemPrompt, userMessage)
}

function generateTitle(query) {
  const cleaned = query.trim().replace(/[?!.]+$/, '')
  return cleaned.length > 50 ? cleaned.slice(0, 50) + '…' : cleaned
}

// ─── Routes ───────────────────────────────────────────────────────────────────

// Health
app.get('/health', (req, res) => res.json({
  ok: true,
  service: 'ask-data',
  model:   'ask-data-response-model',
  embeddings: AZURE_EMBED_ENDPOINT ? 'azure-openai' : 'keyword-only',
}))

// Client verify
app.post('/client/verify', async (req, res) => {
  try {
    const apiKey = extractApiKey(req) || req.body?.apiKey
    if (!apiKey) return res.status(400).json({ valid: false, error: 'apiKey is required' })
    const client = await verifyApiKey(apiKey)
    if (!client) return res.status(401).json({ valid: false, error: 'Invalid or expired API key' })
    res.json({ valid: true, client })
  } catch (err) {
    res.status(500).json({ valid: false, error: err.message })
  }
})

// Admin Routes
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
    const blobsDeleted = [], blobsFailed = []
    if (AZURE_CONNECTION_STRING) {
      try {
        const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING)
        const containerClient   = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
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

// Auth
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
    const client   = await database.collection('clients').findOne({ clientId: req.client.clientId }, { projection: { apiKey: 0 } })
    if (!client) return res.status(404).json({ error: 'Client not found' })
    res.json(client)
  } catch (err) { res.status(500).json({ error: err.message }) }
})

// Conversations
app.post('/chat/conversations', requireClientKey, async (req, res) => {
  try {
    const { title }  = req.body
    const database   = await getChatDb()
    const now        = new Date()
    const conversation = { clientId: req.client.clientId, title: title || 'New Conversation', messages: [], createdAt: now, updatedAt: now }
    const result     = await database.collection('conversations').insertOne(conversation)
    res.status(201).json({ ...conversation, _id: result.insertedId })
  } catch (err) { res.status(500).json({ error: err.message }) }
})

app.post('/chat/conversations/list', requireClientKey, async (req, res) => {
  try {
    const database       = await getChatDb()
    const conversations  = await database.collection('conversations').find({ clientId: req.client.clientId }, { projection: { messages: 0 } }).sort({ updatedAt: -1 }).toArray()
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

// ─── Chat message ─────────────────────────────────────────────────────────────
app.post('/chat/message', requireClientKey, async (req, res) => {
  try {
    const { query, topK = 6, conversationId } = req.body
    if (!query?.trim()) return res.status(400).json({ error: 'query is required' })

    const { clientId, name } = req.client
    const chunks = await loadChunksForClient(clientId)

    if (chunks.length === 0) {
      return res.json({ answer: 'No documents found for your account. Please ensure your documents have been ingested first.', sources: [], client: { clientId, name } })
    }

    const intent = detectQueryIntent(query.trim())
    const hits   = await retrieveChunks(query.trim(), chunks, Math.min(topK, 20))

    if (hits.length === 0) {
      return res.json({ answer: "I couldn't find that in your documents. Try rephrasing your question or asking about it differently.", sources: [], client: { clientId, name } })
    }

    const rawAnswer = await answerWithPhi4(query.trim(), hits, intent)
    const answer = fixBrokenUrls(rawAnswer)

    const sources = hits.map(h => ({
      source_file:  h.source_file  || 'unknown',
      chunk_index:  h.chunk_index  ?? 0,
      score:        typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
      preview:      (h.text || '').slice(0, 300),
    }))

    try {
      const chatDatabase = await getChatDb()
      const col          = chatDatabase.collection('conversations')
      const now          = new Date()
      const userMsg      = { role: 'user',      content: query.trim(), timestamp: now }
      const assistantMsg = { role: 'assistant', content: answer, sources: sources.map(s => ({ source_file: s.source_file, score: s.score })), timestamp: now }

      if (conversationId) {
        await col.updateOne({ _id: new ObjectId(conversationId), clientId }, { $push: { messages: { $each: [userMsg, assistantMsg] } }, $set: { updatedAt: now } })
        res.json({ answer, sources, client: { clientId, name }, conversationId })
      } else {
        const title  = generateTitle(query.trim())
        const result = await col.insertOne({ clientId, title, messages: [userMsg, assistantMsg], createdAt: now, updatedAt: now })
        res.json({ answer, sources, client: { clientId, name }, conversationId: result.insertedId.toString() })
      }
    } catch (histErr) {
      res.json({ answer, sources, client: { clientId, name }, conversationId: conversationId || null })
    }
  } catch (err) {
    res.status(500).json({ error: err.message })
  }
})

// ─── Start ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 4000
app.listen(PORT, () => {
  console.log(`rag-client-auth running on port ${PORT}`)
  console.log(`Model: ${PHI4_MODEL} | Endpoint: ${PHI4_ENDPOINT ? 'configured' : 'MISSING'}`)
  startApiKeyHealthChecker()
})

module.exports = app
