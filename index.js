require('dotenv').config()
const express = require('express')
const cors = require('cors')
const { MongoClient, ObjectId } = require('mongodb')
const { BlobServiceClient } = require('@azure/storage-blob')
const pdfParse = require('pdf-parse')
const mammoth = require('mammoth')
const XLSX = require('xlsx')
const Papa = require('papaparse')
const stringSimilarity = require('string-similarity')
const crypto = require('crypto')
const { resolveIntent, isOutOfScope } = require('./src/ed')

const app = express()

const allowedOrigins = [
  'http://localhost:8080',
  'http://localhost:3000',
  'https://app.powerbi.com',
  'https://msit.powerbi.com',
  'https://anuritchat.vercel.app',
  'https://askdatatest.vercel.app',
  'https://ragadminpanel.vercel.app',
  'https://df.powerbi.com',
  'https://www.anuritinnovation.com/',
  'https://api.powerbi.com',
]

const originAllowed = o => !o || o === 'null' || allowedOrigins.includes(o) || /\.(powerbi|microsoft|office)\.com$/.test(o)
const corsOpts = {
  origin: (o, cb) => cb(null, originAllowed(o)),
  methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'x-session-id'],
  credentials: true,
}
app.use(cors(corsOpts))
app.options('*', cors({ ...corsOpts, origin: (o, cb) => cb(null, true) }))
app.use(express.json())

const MONGODB_URI = process.env.MONGODB_URI
const MONGODB_DB = process.env.MONGODB_DB || 'clientcreds'
const CHAT_HISTORY_URI = process.env.CHAT_HISTORY_URI
const CHAT_HISTORY_DB = process.env.CHAT_HISTORY_DB || 'chathistory'
const AZURE_CONNECTION_STRING = process.env.AZURE_CONNECTION_STRING || ''
const AZURE_CONTAINER_NAME = process.env.AZURE_CONTAINER_NAME || 'vectordbforrag'
const ADMIN_API_KEY = process.env.ADMIN_API_KEY
const KEY_CHECK_INTERVAL_MS = parseInt(process.env.KEY_CHECK_INTERVAL_MS || '300000', 10)
const ASKDATA_ENDPOINT = process.env.ASKDATA_ENDPOINT || ''
const ASKDATA_KEY = process.env.ASKDATA_KEY || ''
const ASKDATA_MODEL = process.env.ASKDATA_MODEL || 'ASKDATA'
const ASKDATA_TIMEOUT_MS = parseInt(process.env.ASKDATA_TIMEOUT_MS || '30000', 10)
const ASKDATA2_ENDPOINT = process.env.ASKDATA2_ENDPOINT || ''
const ASKDATA2_KEY = process.env.ASKDATA2_KEY || ''
const ASKDATA2_MODEL = process.env.ASKDATA2_MODEL || 'ASKDATA2'
const ASKDATA2_TIMEOUT_MS = parseInt(process.env.ASKDATA2_TIMEOUT_MS || '30000', 10)
const ASKDATA2_REWRITE_TIMEOUT_MS = parseInt(process.env.ASKDATA2_REWRITE_TIMEOUT_MS || '8000', 10)
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || '60000', 10)
const WARMUP_CLIENT_IDS = (process.env.WARMUP_CLIENT_IDS || '').split(',').map(s => s.trim()).filter(Boolean)

const RAW_PREFIX = 'raw'
const CHUNK_SIZE = 900
const CHUNK_OVERLAP = 100
const POLICY_CHUNK_SIZE = 700
const POLICY_CHUNK_OVERLAP = 100
const RESEARCH_CHUNK_SIZE = 500
const RESEARCH_CHUNK_OVERLAP = 80
const BLOB_CONCURRENCY = parseInt(process.env.BLOB_CONCURRENCY || '8', 10)
const CHUNK_CACHE_TTL = parseInt(process.env.CHUNK_CACHE_TTL_MS || '300000', 10)
const MAX_HITS_GLOBAL = 20
const CONTEXT_CHAR_LIMIT = 2800
const RELATED_KEYWORDS_COUNT = 5
const RELATED_KEYWORDS_MIN_SCORE = 1
const SENTENCE_WINDOW_SIZE = 2
const MIN_HIT_SCORE_DICT = 3
const MIN_HIT_SCORE_POLICY = 2
const MIN_HIT_SCORE_DEFAULT = 4

const blobServiceClient = AZURE_CONNECTION_STRING ? BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING) : null
const SUPPORTED_EXTENSIONS = new Set(['.pdf', '.docx', '.xlsx', '.json', '.txt', '.csv'])

const RESPONSE_CACHE = new Map()
const RESPONSE_CACHE_TTL = 10 * 60 * 1000
const RESPONSE_CACHE_MAX = 1000

function responseCacheGet(key) {
  const e = RESPONSE_CACHE.get(key)
  if (!e) return null
  if (Date.now() - e.ts > RESPONSE_CACHE_TTL) { RESPONSE_CACHE.delete(key); return null }
  return e.value
}

function responseCacheSet(key, value) {
  if (RESPONSE_CACHE.size >= RESPONSE_CACHE_MAX) RESPONSE_CACHE.delete(RESPONSE_CACHE.keys().next().value)
  RESPONSE_CACHE.set(key, { value, ts: Date.now() })
}

const SYNONYM_PAIRS = [
  [/\bapp(lication)?s?\s+(count|volume|number)\b/i, 'application count'],
  [/\b(total|submitted)\s+app(lication)?s?\b/i, 'application count'],
  [/\bnumber\s+of\s+app(lication)?s\b/i, 'application count'],
  [/\bocc(upancy)?\s+(rate|formula)\b/i, m => `occupancy ${m.match(/formula/i) ? 'formula' : 'rate'}`],
  [/\blead\s+(acq\w*\s+)?cost\b/i, 'lead acquisition cost'],
  [/\bsec\.?\s*(dep\w*)?\b/i, 'security deposit'],
  [/\brent\s+inc\b/i, 'rent increase'],
  [/\bnotice\s+(per|req)\w*\b/i, m => m.match(/req/i) ? 'notice requirement' : 'notice period'],
  [/\blate\s+fee\b/i, 'late payment fee'],
  [/\bpenalty\s+clause\b/i, 'penalty clause'],
  [/\bterm\w*\s+clause\b/i, 'termination clause'],
  [/\beviction\s+proc\w*\b/i, 'eviction procedure'],
  [/\bmaint\w*\s+resp\w*\b/i, 'maintenance responsibility'],
]

function applySynonyms(q) {
  for (const [pat, rep] of SYNONYM_PAIRS) {
    if (typeof rep === 'function') q = q.replace(pat, rep)
    else q = q.replace(pat, rep)
  }
  return q
}

const TYPO_MAP = {
  ehat: 'what', waht: 'what', whta: 'what', whar: 'what', hwo: 'how', hoe: 'how',
  difine: 'define', definr: 'define', defien: 'define', defne: 'define',
  expain: 'explain', expalin: 'explain', explian: 'explain',
  wht: 'what', shwo: 'show', lsit: 'list', lits: 'list',
  polcy: 'policy', policiy: 'policy', poilcy: 'policy',
  tennant: 'tenant', tennat: 'tenant', tentant: 'tenant',
  lanlord: 'landlord', landord: 'landlord',
  rentel: 'rental', rentl: 'rental', leas: 'lease', laese: 'lease',
  deposite: 'deposit', depoist: 'deposit', notise: 'notice', noice: 'notice',
  terminaton: 'termination', termiantion: 'termination',
  maintenence: 'maintenance', maintanence: 'maintenance',
}

function applyTypos(q) {
  return q.split(/\s+/).map(w => TYPO_MAP[w.toLowerCase()] || w).join(' ')
}

function levenshteinSimilarity(a, b) {
  if (!a && !b) return 1
  if (!a || !b) return 0
  a = a.toLowerCase(); b = b.toLowerCase()
  const m = a.length, n = b.length
  const dp = Array.from({ length: m + 1 }, (_, i) => Array.from({ length: n + 1 }, (_, j) => i === 0 ? j : j === 0 ? i : 0))
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1])
  return 1 - dp[m][n] / Math.max(m, n)
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') }
function capFirst(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : '' }

function normalizeQuery(q) {
  return applySynonyms(q).toLowerCase().trim()
    .replace(/\bweek\s+(\d)\b/g, (_, n) => `week 0${n}`)
    .replace(/[?!.]+$/, '').replace(/\s+/g, ' ')
}

function normalizeQueryForCache(q) {
  return normalizeQuery(q)
    .replace(/^(what\s+is\s+(the\s+)?(definition|meaning)\s+(of|for|to)\s+)/i, '')
    .replace(/^(define|explain|tell\s+me\s+about|what\s+are|what\s+is|how\s+(do\s+you\s+|is\s+|are\s+)?calculate|describe|meaning\s+of)\s+(the\s+)?/i, '')
    .replace(/[?!.]+$/, '').replace(/\s+/g, ' ').trim()
}

function getCacheKey(clientId, q) { return `${clientId}:${normalizeQueryForCache(q)}` }

function validateQuery(q) {
  if (!q || typeof q !== 'string') return { valid: false, message: 'Please enter a complete question.' }
  const words = q.trim().split(/\s+/).filter(Boolean)
  if (words.length < 2) return { valid: false, message: 'Please enter a more detailed question.' }
  return { valid: true }
}

function detectDocumentType(chunks) {
  if (!chunks?.length) return 'mixed'
  let policy = 0, dict = 0, research = 0
  for (const c of chunks.slice(0, 50)) {
    const t = (c.text || '').toLowerCase()
    if (c.metadata?.measure || c.metadata?.formula !== undefined) dict += 3
    if (/\b(shall|must|tenant|landlord|lessee|lessor|clause|policy|agreement|pursuant|notwithstanding|whereas|hereby)\b/.test(t)) policy++
    if (/\b(rent|lease|deposit|notice|termination|eviction|maintenance|penalty|breach|obligation)\b/.test(t)) policy++
    if (/\b(is defined as|formula|calculated as|measure|attribute|kpi|metric)\b/.test(t)) dict++
    if (/^(section|article|clause|\d+\.\d+)/im.test(c.text || '')) policy += 2
    if (/\b(abstract|introduction|methodology|conclusion|accuracy|precision|recall|epoch|neural|dataset|training|classification|algorithm)\b/.test(t)) research++
    if (/\b(figure\s+\d|table\s+\d|et\s+al|doi:|references?|ieee|arxiv)\b/.test(t)) research += 2
    if (/\b(employee|employer|conduct|leave|attendance|harassment|compensation|payroll|remote\s+work|performance|handbook|hr)\b/.test(t)) policy += 2
  }
  if (research > policy * 2 && research > dict * 2) return 'research'
  if (policy > dict * 1.5 && policy > research * 1.5) return 'policy'
  if (dict > policy * 1.5 && dict > research * 1.5) return 'dictionary'
  return 'mixed'
}

function detectQueryIntent(q) {
  const n = normalizeQuery(q)
  if (/^(hi|hello|hey|howdy|greetings|good\s+(morning|afternoon|evening)|how\s+are\s+you)\b/.test(n)) return 'greeting'
  if (/\b(url|link|dashboard|power\s*bi|report\s+url)\b/.test(n)) return 'url_lookup'
  if (
    /\b(formula|equation|calculate|calculation|calculated|computed|derived)\b/i.test(n) ||
    /how\s+(is|are|was|were)\s+.+\s+(calculated|computed|determined|derived)/i.test(n) ||
    /what\s+is\s+the\s+(formula|calculation)\s+for/i.test(n) ||
    /how\s+do\s+you\s+(calculate|compute)/i.test(n)
  ) return 'calculation'
  if (/\b(what\s+(happens|is\s+the\s+penalty|are\s+the\s+consequences)|penalty|consequence|breach|violation|non.compliance)\b/i.test(n)) return 'policy_consequence'
  if (/\b(allowed|permitted|can\s+(tenant|landlord|i)|is\s+it\s+allowed|may\s+(tenant|landlord)|right\s+to|entitled\s+to)\b/i.test(n)) return 'policy_permission'
  if (
    /\b(how\s+(many|much|long|often)|duration|period|days|months|amount|limit|maximum|minimum|deadline)\b/i.test(n) &&
    /\b(notice|deposit|rent|fee|penalty|maintenance|payment)\b/i.test(n)
  ) return 'policy_numeric'
  if (/\b(policy|clause|rule|requirement|condition|obligation|responsibility|procedure)\b/i.test(n)) return 'policy_lookup'
  if (/\b(how\s+(should|do|can|must)\s+(employee|staff|worker|i)|what\s+(should|must|do)\s+(employee|staff|worker|i))\b/i.test(n)) return 'policy_lookup'
  if (/\b(report|reporting|escalate|escalation|notify|notification|submit|raise|file)\b/i.test(n)) return 'policy_lookup'
  if (
    /^(what\s+is\s+(the\s+)?(definition|meaning)|define|what\s+(is|are)|explain|tell\s+me\s+about|describe|meaning\s+of)/i.test(n) ||
    /\b(definition|meaning)\b/i.test(n)
  ) return 'definition'
  if (/\b(vs|versus|difference|compare|between)\b/.test(n)) return 'comparison'
  return 'general'
}

function detectMultiTopicQuery(q) {
  const stops = new Set(['is', 'are', 'was', 'were', 'it', 'this', 'that', 'its', 'my', 'your'])
  const diffPats = [
    /^(?:what\s+is\s+the\s+)?difference\s+between\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
    /^compare\s+(.+?)\s+(?:vs\.?|versus|and)\s+(.+?)[\s?]*$/i,
    /^(.+?)\s+vs\.?\s+(.+?)[\s?]*$/i,
  ]
  const andPats = [
    /^what\s+(?:is|are)\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
    /^(?:define|explain|tell\s+me\s+about)\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
    /^(.+?)\s+and\s+(.+?)[\s?]*$/i,
  ]
  for (const p of diffPats) {
    const m = q.match(p)
    if (m) {
      const [a, b] = [m[1], m[2]].map(s => s.trim().replace(/^(what\s+is\s+|the\s+)/i, '').trim())
      if (a.length > 1 && b.length > 1) return { isMulti: true, topics: [a, b], mode: 'comparison' }
    }
  }
  for (const p of andPats) {
    const m = q.match(p)
    if (m) {
      const [a, b] = [m[1], m[2]].map(s => s.trim().replace(/^(what\s+is\s+|what\s+are\s+|define\s+|the\s+)/i, '').trim())
      if (a.length > 1 && b.length > 1 && !stops.has(a.toLowerCase()) && !stops.has(b.toLowerCase()))
        return { isMulti: true, topics: [a, b], mode: 'multi_definition' }
    }
  }
  return { isMulti: false, topics: [], mode: null }
}

function extractSubject(q) {
  const raw = q.trim().replace(/[?!.]+$/, '').trim()
  const n = normalizeQuery(applySynonyms(raw))
  const strippedPats = [
    /^what\s+(?:is|are)\s+(?:the\s+)?(?:definition|meaning)\s+(?:of|for|to)\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^define\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^explain\s+(?:an?\s+|the\s+)?how\s+(.+?)\s+(?:is\s+)?calculated$/i,
    /^explain\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^tell\s+me\s+about\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^describe\s+(?:me\s+)?(?:an?\s+|the\s+)?(.+)$/i,
    /^meaning\s+of\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^how\s+(?:is|are)\s+(.+?)\s+(?:calculated|defined|measured|computed)$/i,
    /^what\s+is\s+the\s+formula\s+for\s+(?:calculating\s+)?(?:an?\s+|the\s+)?(.+)$/i,
    /^how\s+(?:do\s+you\s+)?calculate\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^(.+?)\s+(?:formula|equation|calculation)$/i,
    /^what\s+(?:is|are)\s+(?:an?\s+|the\s+)?(.+)$/i,
    /^(?:what\s+is\s+)?(.+)$/i,
  ]
  for (const p of strippedPats) {
    const m = n.match(p)
    if (m) {
      const s = m[1].trim().replace(/[?!.]+$/, '').trim()
      if (s.length > 0) return s
    }
  }
  return n.replace(/[?!.]+$/, '').trim()
}

function extractMeasureSubject(q) {
  const raw = q.trim().replace(/[?!]+$/, '').trim()
  const prefixes = [
    /^what\s+is\s+(?:the\s+)?/i,
    /^define\s+/i,
    /^explain\s+/i,
    /^tell\s+me\s+about\s+/i,
    /^describe\s+/i,
    /^how\s+(?:is|are|do\s+you\s+calculate)\s+/i,
    /^what\s+(?:is|are)\s+(?:the\s+)?(?:formula|definition|meaning)\s+(?:for|of)\s+/i,
  ]
  let subject = raw
  for (const p of prefixes) {
    const m = raw.match(new RegExp('^' + p.source, 'i'))
    if (m) { subject = raw.slice(m[0].length).trim(); break }
  }
  return subject.replace(/[?!.]+$/, '').trim()
}

function scoreMeasureMatch(querySubject, measureName) {
  if (!measureName) return 0
  const ql = querySubject.toLowerCase().trim()
  const ml = measureName.toLowerCase().trim()
  if (ml === ql) return 1000
  if (ml.includes(ql) && ql.length > 4) return 800
  if (ql.includes(ml) && ml.length > 4) return 700
  const qlNoParens = ql.replace(/\s*\([^)]*\)\s*/g, ' ').replace(/\s+/g, ' ').trim()
  const mlNoParens = ml.replace(/\s*\([^)]*\)\s*/g, ' ').replace(/\s+/g, ' ').trim()
  if (mlNoParens === qlNoParens && qlNoParens.length > 3) return 750
  const parensQ = (ql.match(/\(([^)]+)\)/g) || []).map(x => x.toLowerCase())
  const parensM = (ml.match(/\(([^)]+)\)/g) || []).map(x => x.toLowerCase())
  const matchingParens = parensQ.filter(p => parensM.includes(p)).length
  if (matchingParens > 0 && qlNoParens.length > 3 && mlNoParens.includes(qlNoParens)) return 850
  const qWords = ql.replace(/[()]/g, ' ').split(/\s+/).filter(w => w.length > 1)
  const mWords = ml.replace(/[()]/g, ' ').split(/\s+/).filter(w => w.length > 1)
  const matchedWords = qWords.filter(w => mWords.includes(w)).length
  if (matchedWords === qWords.length && qWords.length >= 3) return 600
  const lev = levenshteinSimilarity(ql, ml)
  if (lev >= 0.9) return 500
  if (lev >= 0.8) return 300
  if (matchedWords >= Math.ceil(qWords.length * 0.7) && qWords.length >= 2) return 200 + matchedWords * 10
  return 0
}

function findBestMeasureChunks(querySubject, chunks, topN = 5) {
  const scored = []
  for (const c of chunks) {
    if (!c.metadata?.measure || c.metadata._expansionRow) continue
    const score = scoreMeasureMatch(querySubject, c.metadata.measure)
    if (score > 0) scored.push({ chunk: c, score })
  }
  scored.sort((a, b) => b.score - a.score)
  return scored.slice(0, topN).map(x => ({ ...x.chunk, _measureScore: x.score }))
}

function extractUrlKeywords(q) {
  const stops = new Set(['power', 'bi', 'report', 'url', 'link', 'for', 'the', 'a', 'an', 'of', 'in', 'get', 'me', 'show', 'give', 'find', 'fetch'])
  return q.toLowerCase().replace(/[^\w\s-]/g, ' ').split(/\s+/).filter(w => w.length > 1 && !stops.has(w))
}

function fixBrokenUrls(t) { return t.replace(/https:\/\/[^\s]+(\s+[^\s]+)/g, m => m.replace(/\s/g, '')) }

function normalizeTerms(t) {
  const l = t.toLowerCase().trim(), v = new Set([l])
  if (l.endsWith('s')) v.add(l.slice(0, -1)); else v.add(l + 's')
  if (l.endsWith('ies')) v.add(l.slice(0, -3) + 'y')
  if (l.endsWith('y')) v.add(l.slice(0, -1) + 'ies')
  return [...v]
}

function trimToCompleteSentence(text, maxLen = 800) {
  if (!text || text.length <= maxLen) return text
  const t = text.slice(0, maxLen)
  const last = Math.max(t.lastIndexOf('. '), t.lastIndexOf('.\n'), t.lastIndexOf('.'))
  if (last > maxLen * 0.5) return t.slice(0, last + 1).trim()
  const ls = t.lastIndexOf(' ')
  return (ls > maxLen * 0.7 ? t.slice(0, ls) : t).trim()
}

function trimPreviewToSentence(text, maxLen = 200) {
  if (!text || text.length <= maxLen) return text.trim()
  const t = text.slice(0, maxLen)
  const end = Math.max(t.lastIndexOf('. '), t.lastIndexOf('.\n'), t.lastIndexOf('! '), t.lastIndexOf('? '))
  if (end > maxLen * 0.4) return t.slice(0, end + 1).trim()
  const ls = t.lastIndexOf(' ')
  return (ls > maxLen * 0.6 ? t.slice(0, ls).trim() + '…' : t.trim() + '…')
}

function ensureSinglePeriod(t) { return t ? t.replace(/\.{2,}/g, '.').replace(/\.\s*\./g, '.').trim() : '' }

function extractFormulaFromText(text) {
  if (!text) return ''
  for (const p of [
    /formula\s*:\s*([^\n.]+)/i, /calculated\s+as\s+([^\n.]+)/i, /computed\s+as\s+([^\n.]+)/i,
    /([a-z0-9\s%()#]+\s*\/\s*[a-z0-9\s%()#]+)/i, /([a-z0-9\s%()#]+\s*=\s*[a-z0-9\s%()#+\-*/]+)/i,
  ]) {
    const m = text.match(p)
    if (m?.[1]?.trim().length > 3) return m[1].trim()
  }
  return ''
}

const NEGATIVE_PAIRS = [
  ['non-recurring', 'recurring'], ['non recurring', 'recurring'], ['denied', 'approved'],
  ['inactive', 'active'], ['rejected', 'accepted'], ['unpaid', 'paid'], ['cancelled', 'active'],
  ['delinquent', 'current'], ['non-', ''],
]

function computeNegativePenalty(subj, text) {
  const qs = subj.toLowerCase(), ct = text.toLowerCase()
  let penalty = 0
  for (const [neg, pos] of NEGATIVE_PAIRS) {
    if (!pos) continue
    const qHasPos = pos.length > 0 && new RegExp(`\\b${escapeRegex(pos)}\\b`, 'i').test(qs)
    const qHasNeg = new RegExp(`\\b${escapeRegex(neg)}\\b`, 'i').test(qs)
    if (qHasPos && !qHasNeg && new RegExp(`\\b${escapeRegex(neg)}\\b`, 'i').test(ct)) penalty += 30
    if (qHasNeg && pos.length > 0 && !new RegExp(`\\b${escapeRegex(neg)}\\b`, 'i').test(ct) && new RegExp(`\\b${escapeRegex(pos)}\\b`, 'i').test(ct)) penalty += 20
  }
  return penalty
}

function buildVocabulary(chunks) {
  const vocab = new Set()
  const stops = new Set(['is', 'the', 'a', 'an', 'of', 'in', 'for', 'to', 'at', 'by', 'as', 'on', 'or', 'and', 'be', 'it', 'its', 'with', 'that', 'this', 'from', 'are', 'was', 'were'])
  for (const c of chunks) {
    const words = (c.text || '').toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)
    for (const w of words) if (w.length >= 3 && !stops.has(w)) vocab.add(w)
    if (c.metadata?.measure) {
      for (const w of c.metadata.measure.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/))
        if (w.length >= 3 && !stops.has(w)) vocab.add(w)
    }
  }
  return [...vocab]
}

const DOMAIN_SHORT_SAFELIST = new Set([
  'count', 'rate', 'rent', 'cost', 'date', 'type', 'name', 'unit', 'term', 'area', 'base', 'gross', 'net',
  'avg', 'sum', 'min', 'max', 'ytd', 'mtd', 'per', 'fee', 'tax', 'due', 'paid', 'void', 'open', 'loss', 'gain',
  'flow', 'days', 'beds', 'bath', 'sqft', 'tier', 'band', 'code', 'flag', 'rank', 'sort', 'key', 'ref',
  'clause', 'rule', 'policy', 'lease', 'notice', 'deposit', 'penalty', 'breach',
  'cnn', 'rnn', 'lstm', 'gru', 'svm', 'mlp', 'knn', 'pca', 'gan', 'vgg',
])

function fuzzyCorrectQuery(q, chunks) {
  if (!chunks?.length) return q
  const vocab = buildVocabulary(chunks)
  if (!vocab.length) return q
  const stops = new Set(['what', 'is', 'are', 'how', 'the', 'a', 'an', 'of', 'in', 'for', 'to', 'at', 'by', 'as', 'on', 'or', 'and', 'define', 'explain', 'show', 'find', 'get', 'list', 'give'])
  return q.split(/\s+/).map(w => {
    const l = w.toLowerCase()
    if (stops.has(l) || DOMAIN_SHORT_SAFELIST.has(l) || l.length < 6 || vocab.includes(l)) return w
    const { bestMatch } = stringSimilarity.findBestMatch(l, vocab)
    const score = bestMatch.rating * 0.6 + levenshteinSimilarity(l, bestMatch.target) * 0.4
    if (score >= 0.72 && bestMatch.target !== l) return bestMatch.target
    return w
  }).join(' ')
}

function needsQueryRewrite(q) {
  const words = q.trim().split(/\s+/).filter(Boolean)
  if (words.length <= 2) return true
  if (/[^\x00-\x7F]/.test(q) && words.length < 5) return true
  if (/(.)\1{3,}/.test(q)) return true
  if (words.length < 4 && !/\b(what|how|define|explain|formula|calculate|list|show|find|url|link)\b/i.test(q)) return true
  if (!/\b(is|are|was|were|what|how|why|when|where|who|define|explain|calculate|show|list|find|get|give|tell)\b/i.test(q) && words.length < 6) return true
  return false
}

async function rewriteQueryWithAskdata2(q) {
  if (!ASKDATA2_ENDPOINT || !ASKDATA2_KEY) return q
  try {
    const r = await fetchWithTimeout(ASKDATA2_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${ASKDATA2_KEY}`, 'Accept': 'application/json' },
      body: JSON.stringify({
        model: ASKDATA2_MODEL,
        messages: [
          { role: 'system', content: 'Fix spelling, grammar, and structure of this RAG query. Expand abbreviations. Return ONLY the rewritten query. If already correct, return unchanged.' },
          { role: 'user', content: q },
        ],
        max_tokens: 80, temperature: 0.0, top_p: 1.0, stream: false,
      }),
    }, ASKDATA2_REWRITE_TIMEOUT_MS)
    if (!r.ok) return q
    const data = await r.json()
    const out = (data.choices?.[0]?.message?.content || '').trim().replace(/^["']|["']$/g, '').trim()
    return (!out || out.length < 3 || out.length > q.length * 4) ? q : out
  } catch { return q }
}

async function preprocessQuery(q) {
  return needsQueryRewrite(q) ? rewriteQueryWithAskdata2(q) : q
}

function expandQueryForPolicy(q) {
  const l = q.toLowerCase()
  const exp = []
  if (/\bsecurity\s+deposit\b/.test(l)) exp.push('security deposit refund return conditions deduction')
  if (/\bnotice\s+(period|to\s+(vacate|quit|terminate))\b/.test(l)) exp.push('notice period days written termination vacate')
  if (/\blate\s+(fee|payment|rent)\b/.test(l)) exp.push('late fee penalty grace period overdue')
  if (/\b(termination|end\s+of\s+lease)\b/.test(l)) exp.push('termination clause early termination penalty break lease')
  if (/\b(maintenance|repair)\b/.test(l)) exp.push('maintenance repair responsibility landlord tenant')
  if (/\beviction\b/.test(l)) exp.push('eviction process procedure notice breach non-payment')
  if (/\b(rent\s+increase|escalation)\b/.test(l)) exp.push('rent increase escalation annual percentage notice')
  if (/\bpet\b/.test(l)) exp.push('pet policy allowed permitted deposit fee')
  if (/\b(sublease|sublet)\b/.test(l)) exp.push('sublease sublet permission consent landlord')
  if (/\brenewal\b/.test(l)) exp.push('lease renewal term extension option notice')
  if (/\b(report|reporting|unethical|misconduct|ethics)\b/.test(l)) exp.push('report unethical behavior reporting channels HR escalation ethics committee anonymous')
  if (/\b(leave|absence|attendance|absent)\b/.test(l)) exp.push('leave attendance absence casual sick earned maternity paternity')
  if (/\b(remote|hybrid|work\s+from\s+home|wfh)\b/.test(l)) exp.push('remote work hybrid flexible arrangement available business hours')
  if (/\b(harassment|discrimination|equal\s+opportunity|inclusive)\b/.test(l)) exp.push('harassment discrimination equal opportunity inclusive complaint report')
  if (/\b(performance|evaluation|review|appraisal)\b/.test(l)) exp.push('performance evaluation review appraisal objectives assessment')
  if (/\b(salary|compensation|payroll|benefits|insurance)\b/.test(l)) exp.push('compensation salary payroll benefits medical insurance')
  if (/\b(safety|emergency|evacuation|incident)\b/.test(l)) exp.push('safety emergency evacuation incident workplace health')
  if (/\b(data|security|privacy|confidential|password)\b/.test(l)) exp.push('data privacy security confidential password information')
  if (/\b(disciplin|warning|suspension|terminat)\b/.test(l)) exp.push('disciplinary action warning suspension termination conduct violation')
  return exp.length ? q + ' ' + exp.join(' ') : q
}

function computeBM25Score(terms, text, avgLen, k1 = 1.5, b = 0.75) {
  const words = text.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)
  const dl = words.length
  const tf = {}
  for (const w of words) tf[w] = (tf[w] || 0) + 1
  let score = 0
  for (const t of terms) {
    const f = tf[t] || 0
    if (!f) continue
    score += 1.5 * (f * (k1 + 1)) / (f + k1 * (1 - b + b * dl / avgLen))
  }
  return score
}

function computePolicyRelevanceScore(q, text, intent) {
  const t = text.toLowerCase()
  let score = 0
  const signals = ['shall', 'must', 'may', 'tenant', 'landlord', 'lessee', 'lessor', 'pursuant', 'hereby', 'thereof', 'herein', 'notwithstanding', 'whereas', 'obligation', 'liability', 'clause', 'section', 'article', 'employee', 'employer', 'encouraged', 'required', 'prohibited', 'expected']
  score += signals.filter(s => t.includes(s)).length * 2
  if (intent === 'policy_consequence' && /\b(penalty|consequence|liable|breach|default|eviction|forfeit|charge|fine|disciplinary|warning|suspension|termination)\b/.test(t)) score += 20
  if (intent === 'policy_permission' && /\b(permitted|allowed|may|shall\s+not|must\s+not|prohibited|forbidden|cannot|restricted|encouraged)\b/.test(t)) score += 20
  if (intent === 'policy_numeric' && /\b\d+\s*(days?|months?|years?|percent|%)\b/.test(t)) score += 25
  if (/^(section|article|clause|\d+\.\d+)/im.test(text)) score += 10
  return score
}

function lightweightRerank(q, chunks, intent, docType) {
  if (!chunks.length) return chunks
  const ql = q.toLowerCase()
  const terms = ql.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 2)
  const totalLen = chunks.reduce((s, c) => s + (c.text || '').split(/\s+/).length, 0)
  const avgLen = totalLen / chunks.length || 100
  const isPolicy = ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  return chunks.map(c => {
    const text = c.text || ''
    const focusSentence = c.metadata?.focus_sentence || ''
    let score = computeBM25Score(terms, text, avgLen) * 10
    if (isPolicy || docType === 'policy') score += computePolicyRelevanceScore(q, text, intent)
    if (focusSentence) {
      const focusBM25 = computeBM25Score(terms, focusSentence, avgLen) * 20
      score += focusBM25
      if (isPolicy && /\b(encouraged|required|must|shall|prohibited|expected|may)\b/i.test(focusSentence)) score += 15
    }
    if (c.metadata?.section_heading) {
      const hl = (c.metadata.section_heading || '').toLowerCase()
      score += terms.filter(t => hl.includes(t)).length * 8
    }
    if (c.metadata?.is_definition_chunk && intent === 'definition') score += 12
    if (c.metadata?.measure) {
      const ml = (c.metadata.measure || '').toLowerCase().trim()
      if (ml === terms.join(' ').trim()) score += 80
    }
    if (c.metadata?.is_clause_chunk && isPolicy) score += 12
    if (new RegExp(escapeRegex(ql.slice(0, 30)), 'i').test(text)) score += 8
    if (c._measureScore) score += c._measureScore * 0.5
    return { ...c, _rerankScore: score, _score: (c._score || 0) + score * 0.3 }
  }).sort((a, b) => (b._score - a._score) || (b._rerankScore - a._rerankScore))
}

function buildInvertedIndex(chunks) {
  const idx = new Map()
  for (let i = 0; i < chunks.length; i++) {
    const focusSentence = chunks[i].metadata?.focus_sentence || ''
    const measureName = chunks[i].metadata?.measure || ''
    const fullText = (chunks[i].text || '') + ' ' + focusSentence + ' ' + measureName
    const words = fullText.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)
    for (const w of words) {
      if (w.length < 2) continue
      if (!idx.has(w)) idx.set(w, new Set())
      idx.get(w).add(i)
    }
  }
  return idx
}

function keywordSearch(q, chunks, topK, intent, invertedIndex) {
  const subject = extractSubject(q)
  const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
  const qLower = normalizeQuery(q)
  const subjectRegex = subjectWords.length > 1
    ? new RegExp(escapeRegex(subject.toLowerCase()), 'i')
    : new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b`, 'i')
  let candidateSet
  if (invertedIndex) {
    const union = new Set()
    const words = intent === 'url_lookup' ? extractUrlKeywords(q) : subjectWords
    for (const w of words) {
      for (const i of (invertedIndex.get(w) || new Set())) union.add(i)
      for (const v of normalizeTerms(w)) for (const i of (invertedIndex.get(v) || new Set())) union.add(i)
    }
    if (intent === 'url_lookup') for (const w of ['url', 'link', 'https', 'powerbi']) for (const i of (invertedIndex.get(w) || new Set())) union.add(i)
    candidateSet = union
  }
  const source = candidateSet ? [...candidateSet].map(i => chunks[i]).filter(Boolean) : chunks.slice(0, 200)
  return source.map(c => {
    const text = (c.text || '').toLowerCase()
    const focusSentence = (c.metadata?.focus_sentence || '').toLowerCase()
    let score = 0
    if (intent === 'url_lookup') {
      if (!text.includes('http')) return { ...c, _score: 0 }
      const kws = extractUrlKeywords(q)
      const matched = kws.filter(w => text.includes(w)).length
      if (!matched) return { ...c, _score: 0 }
      score += matched * 10
    } else {
      if (subjectRegex.test(c.text || '')) {
        score += subjectWords.length * 6
        if (/\b(is defined as|is calculated as|formula:|shall|must|means|encouraged|required|expected|prohibited)\b/i.test((c.text || '').slice(0, (c.text || '').toLowerCase().indexOf(subject.toLowerCase()) + 200))) score += subjectWords.length * 8
      }
      score += subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(c.text || '')).length * 2
      if (focusSentence && subjectRegex.test(focusSentence)) score += subjectWords.length * 10
      if (focusSentence) score += subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(focusSentence)).length * 4
      if (new RegExp(`\\b${escapeRegex(qLower)}\\b`, 'i').test(c.text || '')) score += 3
      if (intent === 'calculation') {
        if (/\b(formula|calculated\s+as|computed\s+as|how\s+to\s+calculate|formula\s+for)\b/i.test(text)) score += 15
        if (text.includes('=') || text.includes('/')) score += 5
      }
      if (intent === 'policy_consequence' && /\b(penalty|consequence|liable|breach|default|eviction|forfeit|disciplinary|warning|suspension|termination)\b/i.test(text)) score += 20
      if (intent === 'policy_permission' && /\b(permitted|allowed|may\s+(not)?|shall\s+not|must\s+not|prohibited|forbidden|encouraged|required)\b/i.test(text)) score += 20
      if (intent === 'policy_numeric' && /\b\d+\s*(days?|months?|years?|percent|%)\b/i.test(text)) score += 20
      if (c.metadata?.section_heading && subjectWords.some(w => (c.metadata.section_heading || '').toLowerCase().includes(w))) score += 25
      if (c.metadata?.measure) {
        const ml = (c.metadata.measure || '').toLowerCase().trim()
        const measureScore = scoreMeasureMatch(subject, c.metadata.measure)
        if (measureScore >= 1000) score += 200
        else if (measureScore >= 800) score += 150
        else if (measureScore >= 600) score += 100
        else if (measureScore >= 300) score += 50
        else if (ml === subject.toLowerCase().trim()) score += 100
        else if (subjectWords.some(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(ml))) score += 10
      }
      score -= computeNegativePenalty(subject, c.text || '')
    }
    return { ...c, _score: score }
  }).filter(c => c._score > 0).sort((a, b) => b._score - a._score).slice(0, topK)
}

function relaxedKeywordSearch(q, chunks, topK, invertedIndex) {
  const subject = extractSubject(q)
  const words = [...new Set([
    ...subject.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1),
    ...q.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 2),
  ])]
  const union = new Set()
  if (invertedIndex) {
    for (const w of words) {
      for (const i of (invertedIndex.get(w) || new Set())) union.add(i)
      for (const v of normalizeTerms(w)) for (const i of (invertedIndex.get(v) || new Set())) union.add(i)
    }
  }
  const source = union.size ? [...union].map(i => chunks[i]).filter(Boolean) : chunks.slice(0, 300)
  return source.map(c => {
    const text = (c.text || '').toLowerCase()
    const focusSentence = (c.metadata?.focus_sentence || '').toLowerCase()
    const matched = words.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(text)).length
    const focusMatched = focusSentence ? words.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(focusSentence)).length * 2 : 0
    const subjectMatch = subject.length > 2 && new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b`, 'i').test(text) ? 5 : 0
    let meta = 0
    if (c.metadata?.measure) {
      const measureScore = scoreMeasureMatch(subject, c.metadata.measure)
      if (measureScore >= 700) meta += 100
      else if (measureScore >= 500) meta += 60
      else if (measureScore >= 300) meta += 30
      else {
        const ml = c.metadata.measure.toLowerCase()
        if (ml === subject.toLowerCase().trim()) meta += 50
        else meta += words.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(ml)).length * 3
      }
    }
    if (c.metadata?.section_heading) meta += words.filter(w => (c.metadata.section_heading || '').toLowerCase().includes(w)).length * 5
    const penalty = computeNegativePenalty(subject, c.text || '')
    return { ...c, _score: Math.max(0, matched + focusMatched + subjectMatch + meta - penalty) }
  }).filter(c => c._score > 0).sort((a, b) => b._score - a._score).slice(0, topK)
}

async function retrieveChunks(q, chunks, topK, invertedIndex, docType, _retry = false) {
  const intent = detectQueryIntent(q)
  const isPolicy = ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  let sq = normalizeQuery(q).replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ')
  if (isPolicy || docType === 'policy' || docType === 'mixed') sq = expandQueryForPolicy(sq)
  if (intent === 'all_urls') return chunks.filter(c => /https?:\/\/\S+/.test(c.text || '')).slice(0, 100)
  const hasMeasureData = chunks.some(c => c.metadata?.measure)
  if (hasMeasureData && docType !== 'policy') {
    const measureSubject = extractMeasureSubject(q)
    const measureHits = findBestMeasureChunks(measureSubject, chunks, 5)
    if (measureHits.length > 0 && measureHits[0]._measureScore >= 600) {
      return measureHits.slice(0, Math.min(topK, MAX_HITS_GLOBAL))
    }
    if (measureHits.length > 0 && measureHits[0]._measureScore >= 300) {
      const kbHits = keywordSearch(sq, chunks, Math.min(40, chunks.length), intent, invertedIndex)
      const combined = [...measureHits, ...kbHits]
      const seen = new Set()
      const deduped = []
      for (const c of combined) {
        const key = (c.metadata?.measure || c.text || '').slice(0, 60)
        if (!seen.has(key)) { seen.add(key); deduped.push(c) }
      }
      return deduped.slice(0, Math.min(topK, MAX_HITS_GLOBAL))
    }
  }
  const candidates = keywordSearch(sq, chunks, Math.min(80, chunks.length), intent, invertedIndex)
  const pool = candidates.length ? candidates : chunks.slice(0, 80)
  const topScore = pool[0]?._score || 0
  let top = []
  if (topScore >= 6) top = pool.slice(0, Math.min(MAX_HITS_GLOBAL, pool.length))
  else if ((intent === 'definition' || intent === 'calculation') && topScore >= 3) top = pool.slice(0, Math.min(MAX_HITS_GLOBAL, pool.length))
  else if (isPolicy && topScore >= 2) top = pool.slice(0, Math.min(MAX_HITS_GLOBAL, pool.length))
  else if (topScore >= 2) top = pool.slice(0, Math.min(10, pool.length))
  if (!top.length && !_retry) {
    const corrected = fuzzyCorrectQuery(q, chunks)
    if (corrected.toLowerCase() !== q.toLowerCase()) return retrieveChunks(corrected, chunks, topK, invertedIndex, docType, true)
  }
  if (!top.length) top = relaxedKeywordSearch(sq, chunks, Math.min(topK * 2, 32), invertedIndex).slice(0, Math.min(topK, MAX_HITS_GLOBAL))
  if (top.length > 1) top = lightweightRerank(q, top, intent, docType)
  const effectiveTopK = intent === 'definition' ? 3 : intent === 'calculation' ? 3 : isPolicy ? 5 : 4
  return top.slice(0, Math.min(effectiveTopK, MAX_HITS_GLOBAL))
}

function getMinRelevanceThreshold(docType, intent) {
  if (docType === 'dictionary') return MIN_HIT_SCORE_DICT
  if (docType === 'policy') return MIN_HIT_SCORE_POLICY
  const isPolicy = ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  if (isPolicy) return MIN_HIT_SCORE_POLICY
  return MIN_HIT_SCORE_DEFAULT
}

function buildContext(hits, intent, docType) {
  const seen = new Set()
  const deduped = []
  for (const h of hits) {
    if (h.metadata?._expansionRow) continue
    const fp = (h.metadata?.focus_sentence || h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
    if (deduped.length >= 5) break
  }
  const isPolicy = docType === 'policy' || ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  let total = 0
  const parts = []
  for (let i = 0; i < deduped.length; i++) {
    const limit = i === 0 ? Math.min(800, CONTEXT_CHAR_LIMIT - total) : Math.min(600, CONTEXT_CHAR_LIMIT - total)
    if (limit < 50) break
    let header = `[S${i + 1}]`
    if (deduped[i].metadata?.section_heading) header += `[${deduped[i].metadata.section_heading.slice(0, 50)}]`
    const focusSentence = deduped[i].metadata?.focus_sentence
    if (isPolicy && focusSentence) {
      const focusBlock = `FOCUS: ${focusSentence}\nCONTEXT: ${(deduped[i].text || '').slice(0, limit - focusSentence.length - 20)}`
      parts.push(`${header}\n${focusBlock}`)
      total += focusBlock.length + 20
    } else if (deduped[i].metadata?.measure) {
      const measureBlock = buildMeasureContextBlock(deduped[i], limit)
      parts.push(`${header}\n${measureBlock}`)
      total += measureBlock.length + 20
    } else {
      const windowText = (deduped[i].text || '').trim()
      let contextBlock = windowText.slice(0, limit)
      if (focusSentence && !contextBlock.includes(focusSentence.slice(0, 30))) {
        contextBlock = focusSentence + '\n' + contextBlock
      }
      contextBlock = contextBlock.slice(0, limit)
      parts.push(`${header}\n${contextBlock}`)
      total += contextBlock.length + 20
    }
  }
  return parts.join('\n---\n')
}

function buildMeasureContextBlock(chunk, maxLen) {
  const m = chunk.metadata
  if (!m) return (chunk.text || '').slice(0, maxLen)
  const parts = []
  if (m.measure) parts.push(`Measure: ${m.measure}`)
  if (m.table) parts.push(`Table: ${m.table}`)
  if (m.description) parts.push(`Description: ${m.description}`)
  if (m.formula) parts.push(`Formula: ${m.formula}`)
  return parts.join('\n').slice(0, maxLen)
}

function buildSystemPrompt(intent, docType) {
  const isPolicy = docType === 'policy' || ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  const isMeasure = docType === 'dictionary' || (docType === 'mixed' && intent !== 'policy_lookup')
  if (isMeasure) {
    if (intent === 'calculation') {
      return `You are a data dictionary assistant. Answer ONLY from the provided context.
RULES:
1. The context contains "Measure:", "Description:", and "Formula:" fields.
2. Return exactly: "**Formula for [Measure Name]:** [formula]."
3. If formula field is empty, extract from description.
4. Never say "I could not find" if the measure context is provided — build the answer from it.
5. Do not add extra sentences or caveats.`
    }
    return `You are a data dictionary assistant. Answer ONLY from the provided context.
RULES:
1. The context contains structured fields: Measure, Table, Description, Formula.
2. Give a clear 1-3 sentence answer using these fields.
3. Bold the measure name.
4. If description is present, use it as the primary definition.
5. Never say "I could not find" if measure context is provided — synthesize from the fields given.
6. Do not fabricate information not present in the context.`
  }
  if (isPolicy) {
    const rule = intent === 'policy_consequence' ? 'State exact penalty, amount, timeframe, or procedure. Be specific.' :
      intent === 'policy_permission' ? 'State clearly if permitted or prohibited and any conditions.' :
      intent === 'policy_numeric' ? 'State the exact number (days/months/amount/%). Do not approximate.' :
      'Find and state the specific rule, procedure, or guideline that directly answers the question.'
    return `You are a precise HR and policy document assistant. Answer ONLY from the provided context.
CRITICAL RULES:
1. Each context block has a "FOCUS:" sentence — this is the most relevant sentence. Prioritize it.
2. Answer the EXACT question asked in 1-3 complete sentences.
3. Do NOT mix answer from multiple sections unless the question requires it.
4. Do not cite source labels like [S1].
5. If the FOCUS sentence directly answers the question, use it as your primary answer.
6. Match your answer to what the question is asking: reporting process vs consequences are different topics.
${rule}`
  }
  if (docType === 'research') return `Answer from context only. Factual and precise. State exact numbers for metrics. 2-3 sentences. No source references.`
  const rule = intent === 'definition' ? 'One sentence definition only. Bold name. No formula.' :
    intent === 'calculation' ? 'Output only: "**Formula for [Name]:** [formula]." Nothing else.' :
    intent === 'comparison' ? 'Bold each name. One definition each. End with "**Key Difference:**" sentence from context.' :
    'Answer directly in 2-3 sentences.'
  return `Answer from context only. Bold subject. Complete sentences. No source refs. If not found say so.\n${rule}`
}

function buildUserMessage(q, hits, intent, docType) {
  const context = buildContext(hits, intent, docType)
  const subject = extractSubject(q)
  const measureSubject = extractMeasureSubject(q)
  const isPolicy = ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  const isMeasure = hits.some(h => h.metadata?.measure)
  let inst = ''
  if (isMeasure && intent === 'calculation') {
    inst = `\nReturn only: "**Formula for ${capFirst(measureSubject)}:** [formula from context]."`
  } else if (isMeasure) {
    inst = `\nQuestion: "${q}"\nThe context contains structured data for the measure "${measureSubject}". Use the Description and Formula fields to answer directly. Bold the measure name.`
  } else if (intent === 'definition' && docType !== 'policy' && docType !== 'research') {
    inst = `\nOne-sentence definition of "${subject}". Bold name. No formula.`
  } else if (intent === 'calculation') {
    inst = `\nReturn only: "**Formula for ${capFirst(subject)}:** [formula]."`
  } else if (intent === 'url_lookup') {
    inst = `\nReturn only the URL for "${extractUrlKeywords(q).join(' ')}".`
  } else if (intent === 'comparison') {
    inst = `\nCompare: ${q}. Bold each. Short definition each. "**Key Difference:**" at end.`
  } else if (isPolicy || docType === 'policy') {
    inst = `\nQUESTION: "${q}"\n\nThe FOCUS sentence in each context block is the most relevant sentence. Read it carefully. Answer ONLY what the question asks about "${subject}". Use the FOCUS sentence as your primary source. Answer in 1-3 direct sentences.`
  } else if (docType === 'research') {
    inst = `\nAnswer from the research context: "${q}". Precise. State exact numbers.`
  } else {
    inst = `\nAnswer: ${q}. Only what was asked. 2-3 sentences.`
  }
  return `CONTEXT:\n${context}${inst}`
}

function isWeakAnswer(a) {
  if (!a || a.trim().length < 15) return true
  const weak = ['i could not find', 'no relevant information', 'not found in', 'i don\'t have', 'i don\'t see', 'unable to find', 'not mentioned', 'not present in', 'no information about', 'cannot find', 'does not contain', 'not available in', 'i couldn\'t find']
  const l = a.toLowerCase().trim()
  return weak.some(p => l.startsWith(p) || (l.length < 80 && l.includes(p)))
}

function buildDirectMeasureAnswer(q, hits, intent) {
  const hit = hits.find(h => h.metadata?.measure && !h.metadata._expansionRow)
  if (!hit) return null
  const m = hit.metadata
  const measureName = m.measure || extractMeasureSubject(q)
  const capName = capFirst(measureName)
  if (intent === 'calculation') {
    if (m.formula) return ensureSinglePeriod(`**Formula for ${capName}:** ${m.formula}.`)
    if (m.description) {
      const f = extractFormulaFromText(m.description)
      if (f) return ensureSinglePeriod(`**Formula for ${capName}:** ${f}.`)
    }
    if (m.description) return ensureSinglePeriod(`**${capName}** is calculated as follows: ${trimToCompleteSentence(m.description, 400)}.`)
    return null
  }
  if (m.description) {
    let ans = `**${capName}** is defined as: ${trimToCompleteSentence(m.description, 400)}`
    if (!ans.endsWith('.')) ans += '.'
    if (m.formula && intent !== 'definition') ans += `\n\n**Formula:** ${m.formula}.`
    return ensureSinglePeriod(ans)
  }
  if (hit.text && hit.text.includes('is defined as')) {
    const match = hit.text.match(/is defined as:?\s*([^.]{10,}\.)/i)
    if (match) return ensureSinglePeriod(`**${capName}** is defined as: ${match[1].trim()}`)
  }
  return null
}

function buildDirectPolicyAnswer(q, hits, intent) {
  const subject = extractSubject(q)
  const esc = escapeRegex(subject.toLowerCase())
  const focusSentences = hits
    .map(h => h.metadata?.focus_sentence)
    .filter(Boolean)
    .filter(s => new RegExp(`\\b${esc}\\b`, 'i').test(s))
  if (focusSentences.length) {
    return ensureSinglePeriod(trimToCompleteSentence([...new Set(focusSentences)].slice(0, 2).join(' '), 400))
  }
  const lines = []
  for (const h of hits) {
    for (const line of (h.text || '').split(/\n+/)) {
      if (line.trim().length < 20) continue
      const isRelevant = /\b(shall|must|may|employee|employer|days?|months?|\d+|notice|deposit|rent|fee|penalty|encouraged|required|expected|prohibited)\b/i.test(line)
      if (isRelevant || new RegExp(`\\b${esc}\\b`, 'i').test(line)) lines.push(line.trim())
    }
  }
  if (lines.length) return ensureSinglePeriod(trimToCompleteSentence([...new Set(lines)].slice(0, 2).join(' '), 400))
  return null
}

function extractAllUrlsFromChunks(chunks) {
  const results = [], seen = new Set()
  const urlRe = /https?:\/\/[^\s"'<>]+/g
  for (const c of chunks) {
    for (const line of (c.text || '').split('\n')) {
      const urls = line.match(urlRe)
      if (!urls) continue
      for (const url of urls) {
        const clean = url.replace(/[.,;)]+$/, '').trim()
        if (!clean.startsWith('http') || seen.has(clean)) continue
        seen.add(clean)
        let name = 'Report'
        const m = line.match(/^(?:Report URL|Power BI link)\s+for\s+(.+?)(?:\s*\([^)]+\))?\s*:\s*https?:/i)
        if (m) name = m[1].trim()
        else {
          const before = line.slice(0, line.indexOf('http')).trim().replace(/\.\s*URL\s*:?\s*$/i, '').replace(/\s*:\s*$/, '').replace(/^(URL|Link|Dashboard|Report)\s*:?\s*/i, '').trim()
          if (before.length > 1 && before.length < 120) name = before
        }
        results.push({ name, url: clean })
      }
    }
  }
  return results
}

function buildFallbackAnswer(q, hits, intent, docType) {
  if (!hits?.length) return "I could not find relevant information about this in your documents."
  const subject = extractSubject(q)
  const esc = escapeRegex(subject.toLowerCase())
  const isPolicy = docType === 'policy' || ['policy_lookup', 'policy_consequence', 'policy_permission', 'policy_numeric'].includes(intent)
  if (intent === 'all_urls') {
    const entries = extractAllUrlsFromChunks(hits)
    return entries.length ? entries.map(e => `**${e.name}:** ${e.url}`).join('\n') : "No URLs found."
  }
  if (intent === 'url_lookup') {
    const urlRe = /https?:\/\/[^\s"'<>]+/
    const kws = extractUrlKeywords(q)
    for (const h of hits) for (const line of (h.text || '').split('\n')) {
      if (!urlRe.test(line)) continue
      if (kws.some(w => line.toLowerCase().includes(w))) {
        const m = line.match(urlRe)
        if (m) return m[0].replace(/[.,;)]+$/, '').trim()
      }
    }
    for (const h of hits) for (const line of (h.text || '').split('\n')) {
      const m = line.match(urlRe)
      if (m) return m[0].replace(/[.,;)]+$/, '').trim()
    }
    return "No matching URL found."
  }
  const directMeasureAns = buildDirectMeasureAnswer(q, hits, intent)
  if (directMeasureAns) return directMeasureAns
  if (isPolicy) {
    const policyAns = buildDirectPolicyAnswer(q, hits, intent)
    if (policyAns) return policyAns
    const excerpt = trimToCompleteSentence((hits[0]?.text || '').trim(), 300)
    return excerpt.length > 30 ? ensureSinglePeriod(excerpt) : "I could not find specific information in your documents."
  }
  if (docType === 'research') {
    const lines = []
    for (const h of hits) {
      for (const line of (h.text || '').split(/\n+/)) {
        if (line.trim().length < 20) continue
        if (/\b(\d+\.?\d*\s*%?|accuracy|precision|recall|model|result)\b/i.test(line) || new RegExp(`\\b${esc}\\b`, 'i').test(line)) lines.push(line.trim())
      }
    }
    if (lines.length) return ensureSinglePeriod(trimToCompleteSentence([...new Set(lines)].slice(0, 2).join(' '), 400))
    const excerpt = trimToCompleteSentence((hits[0]?.text || '').trim(), 300)
    return excerpt.length > 30 ? ensureSinglePeriod(excerpt) : "I could not find specific information in your documents."
  }
  if (intent === 'calculation') {
    for (const h of hits) {
      if (h.metadata?.formula && new RegExp(`\\b${esc}\\b`, 'i').test(h.metadata.measure || ''))
        return ensureSinglePeriod(`**Formula for ${capFirst(h.metadata.measure)}:** ${h.metadata.formula}.`)
      for (const pat of [`how to calculate ${esc}:\\s*([^\\n]+)`, `formula for ${esc}:\\s*([^\\n]+)`]) {
        const m = (h.text || '').match(new RegExp(pat, 'im'))
        if (m) return ensureSinglePeriod(`**Formula for ${capFirst(subject)}:** ${trimToCompleteSentence(m[1].trim(), 300)}.`)
      }
    }
    for (const h of hits) {
      if (!new RegExp(`\\b${esc}\\b`, 'i').test(h.text || '')) continue
      const f = extractFormulaFromText(h.text || '')
      if (f) return ensureSinglePeriod(`**Formula for ${capFirst(subject)}:** ${f}.`)
    }
    return `I could not find a formula for ${capFirst(subject)} in your documents.`
  }
  for (const h of hits) {
    if (!h.metadata?.measure) continue
    const ml = (h.metadata.measure || '').toLowerCase().trim()
    if (ml === subject.toLowerCase() || new RegExp(`\\b${esc}\\b`, 'i').test(ml)) {
      const cap = capFirst(h.metadata.measure)
      if (h.metadata.description) return ensureSinglePeriod(`**${cap}** is defined as: ${h.metadata.description}.`)
    }
  }
  const synPat = new RegExp(`${esc}[^\\n]*is defined as:\\s*([^.\\n]+(?:\\.[^.\\n]+)?)(?:\\.\\s*Formula:\\s*([^.\\n]+))?`, 'im')
  for (const h of hits) {
    const m = (h.text || '').match(synPat)
    if (m) {
      const desc = trimToCompleteSentence((m[1] || '').trim(), 400)
      let ans = `**${capFirst(subject)}** is ${desc}`
      if (!ans.endsWith('.')) ans += '.'
      if (intent !== 'definition' && m[2]) ans += `\n\n**Formula:** ${m[2].trim().slice(0, 300)}.`
      return ensureSinglePeriod(ans)
    }
  }
  const lines = []
  for (const h of hits) for (const line of (h.text || '').split('\n')) {
    if (!new RegExp(`\\b${esc}\\b`, 'i').test(line) || line.trim().length <= 20) continue
    if ((line.match(/\|/g) || []).length > 2) continue
    lines.push(line.trim().replace(/\(from\s+[A-Za-z\s]+\)/g, '').trim())
  }
  if (lines.length) return ensureSinglePeriod(`**${capFirst(subject)}:** ${trimToCompleteSentence([...new Set(lines)].slice(0, 2).join(' '), 400)}.`)
  return "I could not find that specific information in your documents."
}

function cleanAnswer(raw) {
  if (!raw) return ''
  let c = fixBrokenUrls(raw)
    .replace(/^\s*\[S?\s*\d+\][^\n]*\n?/gm, '')
    .replace(/^[^\n]*(\|[^\n]*){3,}$/gm, '')
    .replace(/=== .+ ===\s*/gm, '')
    .replace(/\(from\s+[A-Za-z\s]+\)\s*/g, '')
    .replace(/FOCUS:\s*/g, '')
    .replace(/CONTEXT:\s*/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\.{2,}/g, '.').replace(/\.\s*\./g, '.').trim()
  const lastIdx = Math.max(c.lastIndexOf('. '), c.lastIndexOf('.\n'), c.lastIndexOf('! '), c.lastIndexOf('? '))
  if (lastIdx > c.length * 0.5) { const t = c.slice(0, lastIdx + 1).trim(); if (t.length > 20) c = t }
  if (c.length > 0 && !/[.!?]$/.test(c)) c += '.'
  return ensureSinglePeriod(c)
}

const IN_FLIGHT = new Map()
let askedataActiveCount = 0
const ASKDATA_MAX_CONCURRENT = 3
const askedataQueue = []

function runWithAskedataLimit(fn) {
  return new Promise((res, rej) => {
    function tryRun() {
      if (askedataActiveCount < ASKDATA_MAX_CONCURRENT) {
        askedataActiveCount++
        Promise.resolve().then(fn).then(
          r => { askedataActiveCount--; drainAskedataQueue(); res(r) },
          e => { askedataActiveCount--; drainAskedataQueue(); rej(e) }
        )
      } else askedataQueue.push(tryRun)
    }
    tryRun()
  })
}

function drainAskedataQueue() {
  if (askedataQueue.length > 0 && askedataActiveCount < ASKDATA_MAX_CONCURRENT) askedataQueue.shift()()
}

let askedataFailures = 0, askedataBlockedUntil = 0

function askedataCircuitOpen() {
  if (Date.now() < askedataBlockedUntil) return true
  if (askedataBlockedUntil > 0) { askedataBlockedUntil = 0; askedataFailures = 0 }
  return false
}

function askedataRecordSuccess() { askedataFailures = 0; askedataBlockedUntil = 0 }
function askedataRecordFailure() { if (++askedataFailures >= 3) { askedataBlockedUntil = Date.now() + 30000 } }

async function fetchWithTimeout(url, opts, ms) {
  const ctrl = new AbortController()
  const t = setTimeout(() => ctrl.abort(), ms)
  try { return await fetch(url, { ...opts, signal: ctrl.signal }) }
  catch (e) { if (e.name === 'AbortError') throw new Error(`Timed out after ${ms}ms`); throw e }
  finally { clearTimeout(t) }
}

function withRequestTimeout(fn, ms = REQUEST_TIMEOUT_MS) {
  return async (req, res, next) => {
    let done = false
    const t = setTimeout(() => { if (!done) { done = true; if (!res.headersSent) res.status(503).json({ error: 'Request timed out.' }) } }, ms)
    try { await fn(req, res, next) } catch (e) { if (!done) next(e) } finally { done = true; clearTimeout(t) }
  }
}

async function callASKDATA(sys, user, maxTokens = 512) {
  if (!ASKDATA_ENDPOINT || !ASKDATA_KEY) throw new Error('ASKDATA not configured')
  if (askedataCircuitOpen()) throw new Error('ASKDATA circuit open')
  return runWithAskedataLimit(async () => {
    try {
      const r = await fetchWithTimeout(ASKDATA_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${ASKDATA_KEY}` },
        body: JSON.stringify({ model: ASKDATA_MODEL, messages: [{ role: 'system', content: sys }, { role: 'user', content: user }], temperature: 0.1, max_tokens: maxTokens }),
      }, ASKDATA_TIMEOUT_MS)
      if (!r.ok) { const t = await r.text(); throw new Error(`ASKDATA ${r.status}: ${t}`) }
      const d = await r.json(); askedataRecordSuccess()
      return d.choices?.[0]?.message?.content || ''
    } catch (e) { askedataRecordFailure(); throw e }
  })
}

async function callASKDATA2(sys, user, maxTokens = 512) {
  if (!ASKDATA2_ENDPOINT || !ASKDATA2_KEY) throw new Error('ASKDATA2 not configured')
  const r = await fetchWithTimeout(ASKDATA2_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${ASKDATA2_KEY}`, 'Accept': 'application/json' },
    body: JSON.stringify({ model: ASKDATA2_MODEL, messages: [{ role: 'system', content: sys }, { role: 'user', content: user }], max_tokens: maxTokens, temperature: 0.1, top_p: 1.0, stream: false }),
  }, ASKDATA2_TIMEOUT_MS)
  if (!r.ok) { const t = await r.text(); throw new Error(`ASKDATA2 ${r.status}: ${t}`) }
  const d = await r.json()
  return d.choices?.[0]?.message?.content || ''
}

async function callBestAvailableEngine(sys, user, maxTokens = 512) {
  if (ASKDATA_ENDPOINT && ASKDATA_KEY && !askedataCircuitOpen()) {
    try { const r = await callASKDATA(sys, user, maxTokens); if (r?.trim().length >= 15) return r } catch (e) { console.warn(`[ASKDATA] ${e.message}`) }
  }
  if (ASKDATA2_ENDPOINT && ASKDATA2_KEY) {
    try { const r = await callASKDATA2(sys, user, maxTokens); if (r?.trim().length >= 15) return r } catch (e) { console.error(`[ASKDATA2] ${e.message}`) }
  }
  return ''
}

async function generateAnswerWithFallback(q, hits, intent, docType, chunks, invertedIndex, topK) {
  const isMeasureQuery = hits.some(h => h.metadata?.measure && !h.metadata._expansionRow)
  if (isMeasureQuery) {
    const directAns = buildDirectMeasureAnswer(q, hits, intent)
    if (directAns) {
      const sys = buildSystemPrompt(intent, docType)
      const user = buildUserMessage(q, hits, intent, docType)
      let llmRaw = ''
      try { llmRaw = await Promise.race([callBestAvailableEngine(sys, user, 300), new Promise((_, r) => setTimeout(() => r(new Error('timeout')), 20000))]) }
      catch (e) { console.warn(`[genAnswer measure LLM] ${e.message}`) }
      if (!isWeakAnswer(llmRaw)) return cleanAnswer(llmRaw)
      return directAns
    }
  }
  const sys = buildSystemPrompt(intent, docType)
  const user = buildUserMessage(q, hits, intent, docType)
  let raw = ''
  try { raw = await Promise.race([callBestAvailableEngine(sys, user, 400), new Promise((_, r) => setTimeout(() => r(new Error('timeout')), 40000))]) }
  catch (e) { console.warn(`[genAnswer] ${e.message}`) }
  if (!isWeakAnswer(raw)) return cleanAnswer(raw)
  const exQ = docType === 'policy' ? expandQueryForPolicy(q) : q
  let fbHits = await retrieveChunks(exQ, chunks, Math.min(topK * 2, 12), invertedIndex, docType)
  if (!fbHits.length) fbHits = relaxedKeywordSearch(exQ, chunks, 20, invertedIndex)
  if (!fbHits.length) fbHits = hits
  if (fbHits.some(h => h.metadata?.measure && !h.metadata._expansionRow)) {
    const directAns2 = buildDirectMeasureAnswer(q, fbHits, intent)
    if (directAns2) return directAns2
  }
  const fbSys = sys + '\nRECOVERY: Use semantically related terms. Synthesize from any relevant context.'
  let fbRaw = ''
  try { fbRaw = await Promise.race([callBestAvailableEngine(fbSys, buildUserMessage(q, fbHits, intent, docType), 400), new Promise((_, r) => setTimeout(() => r(new Error('timeout')), 25000))]) }
  catch (e) { console.warn(`[genAnswer fallback] ${e.message}`) }
  if (!isWeakAnswer(fbRaw)) return cleanAnswer(fbRaw)
  const rule = buildFallbackAnswer(q, fbHits, intent, docType)
  if (rule && !rule.toLowerCase().includes('could not find')) return rule
  if (!isWeakAnswer(raw)) return cleanAnswer(raw)
  return buildFallbackAnswer(q, hits, intent, docType)
}

async function generateAnswerForTopic(topic, chunks, topK, invertedIndex, docType) {
  const tq = `what is ${topic}`
  let hits = await retrieveChunks(tq, chunks, topK, invertedIndex, docType)
  if (!hits.length) hits = relaxedKeywordSearch(tq, chunks, 20, invertedIndex)
  if (!hits.length) return null
  const ans = await generateAnswerWithFallback(tq, hits, 'definition', docType, chunks, invertedIndex, topK)
  return ans && !/[.!?]$/.test(ans) ? ans + '.' : ans
}

async function generateComparisonAnswer(a, b, chunks, topK, invertedIndex, docType) {
  const [hA, hB] = await Promise.all([
    retrieveChunks(`what is ${a}`, chunks, topK, invertedIndex, docType),
    retrieveChunks(`what is ${b}`, chunks, topK, invertedIndex, docType),
  ])
  const seen = new Set(), deduped = []
  for (const h of [...hA, ...hB]) { const fp = (h.text || '').trim().slice(0, 80).toLowerCase(); if (!seen.has(fp)) { seen.add(fp); deduped.push(h) } }
  if (!deduped.length) return null
  const ans = await generateAnswerWithFallback(`difference between ${a} and ${b}`, deduped, 'comparison', docType, chunks, invertedIndex, topK)
  if (ans?.trim().length >= 15) return ans
  const [ansA, ansB] = await Promise.all([generateAnswerForTopic(a, chunks, topK, invertedIndex, docType), generateAnswerForTopic(b, chunks, topK, invertedIndex, docType)])
  return [
    `**${capFirst(a)}:** ${ansA && !ansA.includes('could not find') ? ansA : 'Not found.'}`,
    `**${capFirst(b)}:** ${ansB && !ansB.includes('could not find') ? ansB : 'Not found.'}`,
  ].join('\n\n')
}

async function handleMultiTopicQuery(topics, mode, chunks, topK, invertedIndex, docType) {
  if (mode === 'comparison' && topics.length === 2) {
    const ans = await generateComparisonAnswer(topics[0], topics[1], chunks, topK, invertedIndex, docType)
    if (ans) return ans
  }
  const results = await Promise.all(topics.map(async t => ({ t, ans: await generateAnswerForTopic(t, chunks, topK, invertedIndex, docType) })))
  return results.map(({ t, ans }) => {
    const cap = capFirst(t)
    return `**${cap}:**\n${(!ans || ans.includes('could not find')) ? `Not found for "${cap}".` : ans}`
  }).join('\n\n')
}

async function extractPdf(buffer) { const r = await pdfParse(buffer); return r.text || '' }

function splitIntoSentences(text) {
  if (!text) return []
  const normalized = text.replace(/\r\n/g, '\n').replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim()
  const raw = []
  const abbrevSafePattern = /(?<![A-Z][a-z]?\.|Mr\.|Ms\.|Dr\.|Sr\.|Jr\.|vs\.|etc\.|i\.e\.|e\.g\.)([.!?])\s+(?=[A-Z"'])/g
  let lastIndex = 0
  let match
  while ((match = abbrevSafePattern.exec(normalized)) !== null) {
    const sentence = normalized.slice(lastIndex, match.index + 1).trim()
    if (sentence.length > 15) raw.push(sentence)
    lastIndex = match.index + match[0].length - (match[0].length - match[0].trimStart().length)
    abbrevSafePattern.lastIndex = match.index + 1
  }
  const remaining = normalized.slice(lastIndex).trim()
  if (remaining.length > 15) raw.push(remaining)
  if (raw.length === 0 && normalized.length > 15) return [normalized]
  return raw
}

function chunkDocxBySentenceWindow(text, sourceFile, sectionHeading, headingLevel, windowSize) {
  const sentences = splitIntoSentences(text)
  if (!sentences.length) return []
  const chunks = []
  for (let i = 0; i < sentences.length; i++) {
    const start = Math.max(0, i - windowSize)
    const end = Math.min(sentences.length - 1, i + windowSize)
    const window = sentences.slice(start, end + 1).join(' ')
    const focusSentence = sentences[i]
    const isDefinitionLike = /\b(is defined as|means|refers to|encouraged to|required to|must|shall|prohibited|expected to)\b/i.test(focusSentence)
    chunks.push({
      text: window,
      source_file: sourceFile,
      chunk_index: 0,
      embedding: [],
      metadata: {
        section_heading: sectionHeading || '',
        heading_level: headingLevel || 0,
        focus_sentence: focusSentence,
        sentence_index: i,
        sentence_total: sentences.length,
        window_start: start,
        window_end: end,
        is_definition_chunk: isDefinitionLike,
        chunk_position: i < 2 ? 'early' : i > sentences.length - 3 ? 'late' : 'middle',
      },
    })
  }
  return chunks
}

async function extractWordWithHeadings(buffer) {
  const styleMap = [
    "p[style-name='Heading 1'] => h1:fresh",
    "p[style-name='Heading 2'] => h2:fresh",
    "p[style-name='Heading 3'] => h3:fresh",
    "p[style-name='Heading 4'] => h4:fresh",
    "p[style-name='Title'] => h1:fresh",
    "p[style-name='Subtitle'] => h2:fresh",
    "p[style-name='heading 1'] => h1:fresh",
    "p[style-name='heading 2'] => h2:fresh",
  ]
  try {
    const r = await mammoth.convertToHtml({ buffer, styleMap })
    return { html: r.value || '', hasHeadings: /<h[1-4]>/i.test(r.value) }
  } catch {
    const r = await mammoth.extractRawText({ buffer })
    return { html: r.value || '', hasHeadings: false }
  }
}

function htmlToSentenceWindowChunks(html, sourceFile) {
  const chunks = []
  let globalChunkIndex = 0
  const tagStripRe = /<[^>]+>/g
  const decode = s => s
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ').replace(/&#(\d+);/g, (_, n) => String.fromCharCode(n))
    .replace(/&quot;/g, '"').replace(/&#x2019;/g, "'").replace(/&#x2018;/g, "'")
    .replace(/&#x201C;/g, '"').replace(/&#x201D;/g, '"')
  const headingRe = /<(h[1-4])>(.*?)<\/h[1-4]>/gi
  let lastIndex = 0, currentHeading = '', currentLevel = 0
  const sections = []
  let matchH
  headingRe.lastIndex = 0
  while ((matchH = headingRe.exec(html)) !== null) {
    if (lastIndex < matchH.index) {
      sections.push({ heading: currentHeading, level: currentLevel, content: html.slice(lastIndex, matchH.index) })
    }
    currentHeading = decode(matchH[2].replace(tagStripRe, '').trim())
    currentLevel = parseInt(matchH[1][1])
    lastIndex = matchH.index + matchH[0].length
  }
  if (lastIndex < html.length) {
    sections.push({ heading: currentHeading, level: currentLevel, content: html.slice(lastIndex) })
  }
  for (const sec of sections) {
    const rawText = decode(
      sec.content
        .replace(/<\/?(p|li|br|div|ul|ol|td|tr)[^>]*>/gi, '\n')
        .replace(tagStripRe, '')
        .replace(/\n{3,}/g, '\n\n')
        .trim()
    )
    if (rawText.length < 20) continue
    const paragraphs = rawText.split(/\n{2,}/).map(p => p.trim()).filter(p => p.length > 20)
    for (const para of paragraphs) {
      const sentenceChunks = chunkDocxBySentenceWindow(para, sourceFile, sec.heading, sec.level, SENTENCE_WINDOW_SIZE)
      for (const sc of sentenceChunks) {
        sc.chunk_index = globalChunkIndex++
        chunks.push(sc)
      }
    }
  }
  return chunks
}

function chunkPlainText(text, sourceFile, chunkSize, overlap, isPolicy) {
  const chunks = [], blocks = text.replace(/\r\n/g, '\n').split(/\n{2,}/).map(b => b.trim()).filter(b => b.length > 0)
  let buf = [], bufLen = 0, idx = 0
  function flush() {
    const s = buf.join('\n\n')
    if (s.length >= 30) chunks.push({ text: s, source_file: sourceFile, chunk_index: idx++, embedding: [], metadata: { is_definition_chunk: /\b(is defined as|means|refers to)\b/i.test(s), chunk_position: chunks.length < 3 ? 'early' : 'middle' } })
    buf = []; bufLen = 0
  }
  for (const block of blocks) {
    if (block.length > chunkSize * 1.5) {
      if (buf.length) flush()
      const lines = block.split('\n').filter(l => l.trim())
      let lb = [], ll = 0
      for (const line of lines) {
        if (lb.length && ll + line.length > chunkSize) {
          const s = lb.join('\n')
          if (s.length >= 30) chunks.push({ text: s, source_file: sourceFile, chunk_index: idx++, embedding: [], metadata: { chunk_position: 'middle' } })
          lb = lb.slice(-overlap); ll = lb.join('\n').length
        }
        lb.push(line); ll += line.length
      }
      if (lb.length) { const s = lb.join('\n'); if (s.length >= 30) chunks.push({ text: s, source_file: sourceFile, chunk_index: idx++, embedding: [], metadata: { chunk_position: 'middle' } }) }
      continue
    }
    const proj = bufLen + (bufLen ? 2 : 0) + block.length
    if (buf.length && proj > chunkSize) { const last = buf[buf.length - 1] || ''; flush(); if (last) { buf.push(last); bufLen = last.length } }
    buf.push(block); bufLen += (bufLen ? 2 : 0) + block.length
  }
  if (buf.length) flush()
  return chunks
}

function splitWithOverlap(text, maxSize, overlap) {
  const sents = text.match(/[^.!?]+[.!?]+/g) || [text]
  const out = []
  let cur = [], curLen = 0
  for (const s of sents) {
    if (curLen + s.length > maxSize && cur.length) {
      out.push(cur.join(' '))
      let ol = [], olen = 0
      for (let i = cur.length - 1; i >= 0; i--) {
        if (olen + cur[i].length <= overlap) { ol.unshift(cur[i]); olen += cur[i].length } else break
      }
      cur = [...ol]; curLen = olen
    }
    cur.push(s); curLen += s.length
  }
  if (cur.length) out.push(cur.join(' '))
  return out.filter(s => s.trim().length > 30)
}

function isResearchDocument(text, fileName) {
  const name = (fileName || '').toLowerCase()
  if (/research|paper|study|survey|journal|conference|thesis|dissertation|preprint/i.test(name)) return true
  const s = text.slice(0, 4000).toLowerCase()
  let sig = 0
  if (/\b(abstract|introduction|methodology|related\s+work|literature\s+review)\b/.test(s)) sig += 3
  if (/\b(accuracy|precision|recall|f1.score|auc|roc|confusion\s+matrix)\b/.test(s)) sig += 3
  if (/\b(neural\s+network|deep\s+learning|machine\s+learning|convolutional|classification|detection)\b/.test(s)) sig += 2
  if (/\b(dataset|training\s+set|test\s+set|validation|epoch|batch\s+size)\b/.test(s)) sig += 2
  if (/\b(et\s+al|doi:|arxiv|ieee|figure\s+\d|table\s+\d|references)\b/.test(s)) sig += 3
  return sig >= 6
}

function isPolicyDocument(text, fileName) {
  const name = (fileName || '').toLowerCase()
  if (/policy|lease|agreement|contract|terms|conditions|rules|manual|handbook|sop|compliance|procedure|offer|letter/i.test(name)) return true
  const s = text.slice(0, 3000).toLowerCase()
  let sig = 0
  if (/\b(shall|must|hereby|pursuant|notwithstanding|whereas|thereof|herein)\b/.test(s)) sig += 3
  if (/\b(tenant|landlord|lessee|lessor|party|parties|employee|employer)\b/.test(s)) sig += 2
  if (/\b(clause|exhibit|addendum|schedule|section|article)\b/.test(s)) sig += 2
  if (/\b(agreement|contract|policy|lease|terms|offer|salary|compensation|joining)\b/.test(s)) sig += 2
  if (/\b(security deposit|notice period|termination|eviction|maintenance|late fee|probation|benefits)\b/.test(s)) sig += 3
  if (/^(section|article|clause|\d+\.\d+)\s/im.test(text.slice(0, 5000))) sig += 3
  if (/\b(employee|employer|conduct|leave|attendance|harassment|compensation|payroll|remote\s+work|performance|handbook|hr)\b/.test(s)) sig += 3
  if (/\b(abstract|methodology|conclusion|accuracy|precision|recall|epoch|neural|figure|table\s+\d)\b/.test(s)) sig -= 3
  return sig >= 5
}

function chunkResearchDocument(text, sourceFile) {
  const secPat = /^(?:(?:Abstract|Introduction|Background|Related\s+Work|Literature\s+Review|Methodology|Methods?|Proposed\s+(?:Method|Model|Approach|Framework)|(?:Experimental\s+)?(?:Results?|Evaluation|Discussion)|Conclusion|References?|Acknowledgements?|Appendix)\s*\n|(?:\d+\.?\s+[A-Z][A-Za-z\s]{3,})\n)/gm
  const matches = []
  let m
  while ((m = secPat.exec(text)) !== null) matches.push({ index: m.index, heading: m[0].trim() })
  if (matches.length < 2) return chunkPlainText(text, sourceFile, RESEARCH_CHUNK_SIZE, RESEARCH_CHUNK_OVERLAP, false)
  const chunks = []
  let idx = 0
  for (let i = 0; i < matches.length; i++) {
    const start = matches[i].index, end = i + 1 < matches.length ? matches[i + 1].index : text.length
    const sec = text.slice(start, end).trim()
    if (sec.length < 30) continue
    const pos = i < 2 ? 'early' : i > matches.length - 2 ? 'late' : 'middle'
    if (sec.length <= RESEARCH_CHUNK_SIZE) {
      chunks.push({ text: sec, source_file: sourceFile, chunk_index: idx++, embedding: [], metadata: { section_heading: matches[i].heading, is_research_section: true, chunk_position: pos } })
    } else {
      for (const sub of splitWithOverlap(sec, RESEARCH_CHUNK_SIZE, RESEARCH_CHUNK_OVERLAP))
        chunks.push({ text: sub, source_file: sourceFile, chunk_index: idx++, embedding: [], metadata: { section_heading: matches[i].heading, is_research_section: true, chunk_position: pos } })
    }
  }
  return chunks.length ? chunks : chunkPlainText(text, sourceFile, RESEARCH_CHUNK_SIZE, RESEARCH_CHUNK_OVERLAP, false)
}

function extractSpreadsheet(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellNF: true })
  const rows = []
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName]
    const raw = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1, raw: false })
    if (!raw.length) continue
    let hIdx = -1
    for (let i = 0; i < Math.min(15, raw.length); i++) {
      const cells = raw[i].map(c => String(c).trim()).filter(Boolean)
      if (cells.length >= 2 && cells.filter(c => c.length <= 60).length >= 2) { hIdx = i; break }
    }
    if (hIdx === -1) hIdx = 0
    const rawH = raw[hIdx].map(h => String(h).trim())
    const headers = []
    let lastNB = ''
    for (const h of rawH) { if (h !== '') { lastNB = h; headers.push(h) } else headers.push(lastNB || `Col${headers.length + 1}`) }
    const scoreH = (h, pats) => { const l = h.toLowerCase().trim(); for (const [r, s] of pats) if (r.test(l)) return s; return 0 }
    const hPats = {
      name: [[/\b(measure|attribute|field|metric|kpi)\s*name\b/, 100], [/^name$/, 90], [/\bname\b/, 70]],
      table: [[/\b(table|module|category|group|domain|section)\b/, 100], [/^table$/, 90]],
      description: [[/\b(description|desc|definition|about|summary)\b/, 100]],
      formula: [[/\b(formula|calculation|calc|how\s+calculated|computed\s+as)\b/, 100]],
      url: [[/\b(url|link|href|report\s+link|dashboard)\b/, 100]],
      additional: [[/\b(additional|extra|notes?|info|configuration)\b/, 100]],
    }
    const colIdx = {}
    for (const [field, pats] of Object.entries(hPats)) {
      const best = headers.map((h, i) => ({ i, s: scoreH(h, pats) })).filter(x => x.s > 0).sort((a, b) => b.s - a.s)[0]
      if (best) colIdx[field] = best.i
    }
    let emitted = 0
    for (let i = hIdx + 1; i < raw.length; i++) {
      const row = raw[i]
      if (!row.some(c => String(c).trim() !== '')) continue
      const cells = row.map(c => String(c || '').replace(/\r?\n/g, ' ').trim())
      const name = colIdx.name !== undefined ? (cells[colIdx.name] || '').trim() : ''
      const table = colIdx.table !== undefined ? (cells[colIdx.table] || '').trim() : sheetName
      const desc = colIdx.description !== undefined ? (cells[colIdx.description] || '').trim() : ''
      const url = colIdx.url !== undefined ? (cells[colIdx.url] || '').trim() : ''
      const add = colIdx.additional !== undefined ? (cells[colIdx.additional] || '').trim() : ''
      let formula = colIdx.formula !== undefined ? (cells[colIdx.formula] || '').trim() : ''
      if (!formula && desc) {
        for (const p of [/(.*?\/.*?)/i, /(=.*?)/i, /(calculated\s+as.*)/i, /(divided\s+by.*)/i, /(sum\s+of.*)/i]) {
          const m = desc.match(p); if (m?.[0]?.trim().length > 3) { formula = m[0].trim(); break }
        }
      }
      if (name) {
        let syn = `${name}`
        if (table && table !== sheetName) syn += ` (${table})`
        if (desc) syn += ` is defined as: ${desc}`
        if (formula && !desc.toLowerCase().includes(formula.toLowerCase())) syn += `. Formula: ${formula}`
        if (add) syn += `. Additional Info: ${add}`
        if (url) syn += `. URL: ${url}`
        rows.push({ text: syn, metadata: { measure: name, table: table || sheetName, formula: formula || '', description: desc || '', url: url || '', sourceSheet: sheetName, _expansionRow: false } })
        if (formula) {
          rows.push({ text: `How to calculate ${name}: ${formula}`, metadata: { measure: name, table: table || sheetName, formula, description: desc || '', url: '', sourceSheet: sheetName, _expansionRow: true } })
          rows.push({ text: `Formula for ${name}: ${formula}`, metadata: { measure: name, table: table || sheetName, formula, description: desc || '', url: '', sourceSheet: sheetName, _expansionRow: true } })
        }
        if (url) {
          rows.push({ text: `Report URL for ${name}: ${url}`, metadata: { measure: name, table: table || sheetName, formula: '', description: '', url, sourceSheet: sheetName, _expansionRow: true } })
          rows.push({ text: `Power BI link for ${name}: ${url}`, metadata: { measure: name, table: table || sheetName, formula: '', description: '', url, sourceSheet: sheetName, _expansionRow: true } })
        }
        emitted++
      } else if (desc) {
        rows.push({ text: desc, metadata: { measure: '', table: table || sheetName, formula: '', description: desc, url: '', sourceSheet: sheetName, _expansionRow: false } })
      }
    }
    if (!emitted) {
      for (let i = hIdx + 1; i < raw.length; i++) {
        const cells = raw[i].map(c => String(c || '').trim()).filter(Boolean)
        if (cells.length) rows.push({ text: cells.join(' | '), metadata: { measure: '', table: sheetName, formula: '', description: '', url: '', sourceSheet: sheetName, _expansionRow: false } })
      }
    }
  }
  return rows
}

async function extractTextFromBuffer(buffer, fileName) {
  const ext = ('.' + fileName.split('.').pop()).toLowerCase()
  if (ext === '.pdf') return extractPdf(buffer)
  if (ext === '.docx') return null
  if (ext === '.csv') { const t = buffer.toString('utf-8'); const r = Papa.parse(t, { header: true, skipEmptyLines: true }); return r.data?.length ? r.data.map((row, i) => `Row ${i + 1}: ` + Object.entries(row).map(([k, v]) => `${k}=${v}`).join(' | ')).join('\n') : t }
  if (ext === '.json') { try { return JSON.stringify(JSON.parse(buffer.toString('utf-8')), null, 2) } catch { return buffer.toString('utf-8') } }
  if (ext === '.txt') return buffer.toString('utf-8')
  return ''
}

async function downloadBlobAsBuffer(containerClient, blobName) {
  const dl = await containerClient.getBlobClient(blobName).download()
  const parts = []
  for await (const c of dl.readableStreamBody) parts.push(Buffer.isBuffer(c) ? c : Buffer.from(c))
  return Buffer.concat(parts)
}

async function _doLoadChunks(clientId) {
  if (!blobServiceClient) throw new Error('AZURE_CONNECTION_STRING not set')
  const container = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
  const prefix = `${RAW_PREFIX}/${clientId}/`
  const blobNames = []
  for await (const blob of container.listBlobsFlat({ prefix })) {
    const ext = ('.' + blob.name.split('.').pop()).toLowerCase()
    if (SUPPORTED_EXTENSIONS.has(ext)) blobNames.push(blob.name)
  }
  const allChunks = []
  let offset = 0
  for (let i = 0; i < blobNames.length; i += BLOB_CONCURRENCY) {
    const batch = blobNames.slice(i, i + BLOB_CONCURRENCY)
    const results = await Promise.allSettled(batch.map(async blobName => {
      const fileName = blobName.split('/').pop()
      const ext = ('.' + fileName.split('.').pop()).toLowerCase()
      const buffer = await downloadBlobAsBuffer(container, blobName)
      if (ext === '.xlsx') {
        return extractSpreadsheet(buffer).map((r, idx) => ({ text: r.text, source_file: fileName, chunk_index: idx, embedding: [], metadata: r.metadata || null }))
      }
      if (ext === '.docx') {
        const { html, hasHeadings } = await extractWordWithHeadings(buffer)
        if (!html?.trim()) return []
        if (hasHeadings) {
          const chunks = htmlToSentenceWindowChunks(html, fileName)
          if (chunks.length > 0) {
            console.log(`[chunkLoader] sentence-window docx (headed): ${fileName} (${chunks.length} chunks)`)
            return chunks
          }
        }
        const plainText = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
        if (!plainText) return []
        const paragraphs = plainText.split(/\s{3,}|\n\n+/).map(p => p.trim()).filter(p => p.length > 20)
        const chunks = []
        let idx = 0
        for (const para of paragraphs) {
          const sentChunks = chunkDocxBySentenceWindow(para, fileName, '', 0, SENTENCE_WINDOW_SIZE)
          for (const sc of sentChunks) { sc.chunk_index = idx++; chunks.push(sc) }
        }
        if (chunks.length > 0) {
          console.log(`[chunkLoader] sentence-window docx (plain): ${fileName} (${chunks.length} chunks)`)
          return chunks
        }
        return chunkPlainText(plainText, fileName, POLICY_CHUNK_SIZE, POLICY_CHUNK_OVERLAP, true)
      }
      const text = await extractTextFromBuffer(buffer, fileName)
      if (!text?.trim()) return []
      if (isResearchDocument(text, fileName)) { console.log(`[chunkLoader] research: ${fileName}`); return chunkResearchDocument(text, fileName) }
      if (isPolicyDocument(text, fileName)) { console.log(`[chunkLoader] policy: ${fileName}`); return chunkPlainText(text, fileName, POLICY_CHUNK_SIZE, POLICY_CHUNK_OVERLAP, true) }
      return chunkPlainText(text, fileName, CHUNK_SIZE, CHUNK_OVERLAP, false)
    }))
    for (const r of results) {
      if (r.status === 'fulfilled') { r.value.forEach((c, i) => { c.chunk_index = offset + i }); offset += r.value.length; allChunks.push(...r.value) }
      else console.warn('[loadChunks] blob failed:', r.reason?.message)
    }
  }
  return allChunks
}

const CHUNK_CACHE = new Map()

async function loadChunksForClient(clientId) {
  const now = Date.now()
  const cached = CHUNK_CACHE.get(clientId)
  if (cached?.chunks) {
    if (now - cached.ts <= CHUNK_CACHE_TTL) return cached
    if (!cached.loading) {
      const p = _doLoadChunks(clientId).then(chunks => {
        const invertedIndex = buildInvertedIndex(chunks), docType = detectDocumentType(chunks)
        CHUNK_CACHE.set(clientId, { chunks, invertedIndex, docType, ts: Date.now(), loading: null })
      }).catch(e => { const ex = CHUNK_CACHE.get(clientId); if (ex) CHUNK_CACHE.set(clientId, { ...ex, loading: null }); console.warn(`[cache refresh] ${clientId}: ${e.message}`) })
      CHUNK_CACHE.set(clientId, { ...cached, loading: p })
    }
    return cached
  }
  if (cached?.loading) { await cached.loading; return CHUNK_CACHE.get(clientId) }
  const p = _doLoadChunks(clientId).then(chunks => {
    const invertedIndex = buildInvertedIndex(chunks), docType = detectDocumentType(chunks)
    CHUNK_CACHE.set(clientId, { chunks, invertedIndex, docType, ts: Date.now(), loading: null })
    return chunks
  }).catch(e => { CHUNK_CACHE.set(clientId, { chunks: null, invertedIndex: null, docType: 'mixed', ts: 0, loading: null }); throw e })
  CHUNK_CACHE.set(clientId, { chunks: null, invertedIndex: null, docType: 'mixed', ts: 0, loading: p })
  await p
  return CHUNK_CACHE.get(clientId)
}

function invalidateChunkCache(clientId) { CHUNK_CACHE.delete(clientId) }

function warmupChunkCaches() {
  if (!WARMUP_CLIENT_IDS.length || !blobServiceClient) return
  for (const id of WARMUP_CLIENT_IDS) loadChunksForClient(id).then(({ chunks }) => console.log(`[warmup] ${id}: ${chunks.length} chunks`)).catch(e => console.warn(`[warmup] ${id}: ${e.message}`))
}

function computeKeywordRelevanceScore(subjectWords, chunk) {
  const measure = (chunk.metadata?.measure || '').toLowerCase()
  const desc = (chunk.metadata?.description || '').toLowerCase()
  const text = (chunk.text || '').toLowerCase()
  const sp = subjectWords.join(' ')
  if (measure === sp) return 0
  let s = 0
  s += subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(measure)).length * 15
  if (subjectWords.every(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(measure)) && measure !== sp) s += 20
  s += subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(desc)).length * 3
  s += subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(text)).length
  return s
}

function buildRelatedKeywords(subject, hits, chunks, invertedIndex, topN = RELATED_KEYWORDS_COUNT) {
  const sl = subject.toLowerCase().trim()
  const sw = sl.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 2)
  if (!sw.length) return []
  const primaryKeys = new Set(hits.map(h => h.metadata?.measure?.toLowerCase().trim()).filter(Boolean))
  const cands = new Map(), seen = new Set([sl])
  for (const c of chunks) {
    if (!c.metadata?.measure || c.metadata._expansionRow) continue
    const ml = c.metadata.measure.trim().toLowerCase()
    if (seen.has(ml)) continue
    seen.add(ml)
    const score = computeKeywordRelevanceScore(sw, c)
    if (score < RELATED_KEYWORDS_MIN_SCORE) continue
    const ex = cands.get(ml)
    if (!ex || score > ex.score) cands.set(ml, { keyword: c.metadata.measure, score, table: c.metadata.table || '', description: c.metadata.description || '', formula: c.metadata.formula || '', isPrimary: primaryKeys.has(ml) })
  }
  const sorted = [...cands.values()].sort((a, b) => b.isPrimary !== a.isPrimary ? (b.isPrimary ? 1 : -1) : b.score - a.score).slice(0, topN)
  const max = sorted[0]?.score || 1
  return sorted.map(item => ({ keyword: item.keyword, table: item.table, description: item.description ? trimPreviewToSentence(item.description, 120) : '', formula: item.formula ? trimPreviewToSentence(item.formula, 100) : '', confidenceScore: Math.min(100, Math.round(item.score / Math.max(max, 1) * 100)), isPrimaryHit: item.isPrimary }))
}

function buildRelatedMetrics(subject, hits, chunks, topN = 6) {
  const hasMeasureData = chunks.some(c => c.metadata?.measure && !c.metadata._expansionRow)
  if (!hasMeasureData) return []
  const sl = subject.toLowerCase().trim()
  const sw = sl.replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 2)
  if (!sw.length) return []
  const primaryMeasures = new Set(
    hits.filter(h => h.metadata?.measure && !h.metadata._expansionRow)
      .map(h => h.metadata.measure.toLowerCase().trim())
  )
  const cands = new Map()
  for (const c of chunks) {
    if (!c.metadata?.measure || c.metadata._expansionRow) continue
    const ml = c.metadata.measure.trim().toLowerCase()
    if (primaryMeasures.has(ml)) continue
    const desc = (c.metadata.description || '').toLowerCase()
    let score = 0
    for (const w of sw) {
      const re = new RegExp(`\\b${escapeRegex(w)}\\b`, 'i')
      if (re.test(ml)) score += 20
      if (re.test(desc)) score += 5
    }
    if (score < 5) continue
    const existing = cands.get(ml)
    if (!existing || score > existing.score) {
      const rawDesc = c.metadata.description || ''
      const firstSentence = rawDesc.split(/\.\s+/)[0].trim()
      const shortDescription = firstSentence.length > 130 ? firstSentence.slice(0, 127) + '...' : (firstSentence ? (firstSentence.endsWith('.') ? firstSentence : firstSentence + '.') : '')
      cands.set(ml, { measure: c.metadata.measure, table: c.metadata.table || '', shortDescription, formula: c.metadata.formula || '', score })
    }
  }
  const sorted = [...cands.values()].sort((a, b) => b.score - a.score).slice(0, topN)
  if (!sorted.length) return []
  const maxScore = sorted[0].score || 1
  return sorted.map(item => ({
    measure: item.measure,
    table: item.table,
    shortDescription: item.shortDescription,
    formula: item.formula,
    confidenceScore: Math.min(100, Math.round((item.score / maxScore) * 100)),
  }))
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
  const client = new MongoClient(CHAT_HISTORY_URI || MONGODB_URI)
  await client.connect()
  chatDb = client.db(CHAT_HISTORY_DB)
  return chatDb
}

const CLIENT_CACHE = new Map(), CACHE_TTL_MS = 5 * 60 * 1000
function getCached(k) { const e = CLIENT_CACHE.get(k); if (!e || Date.now() - e.cachedAt > CACHE_TTL_MS) { if (e) CLIENT_CACHE.delete(k); return null }; return e }
function setCache(k, d) { CLIENT_CACHE.set(k, { ...d, cachedAt: Date.now() }) }
function evictCache(k) { if (k) CLIENT_CACHE.delete(k) }

async function verifyApiKey(apiKey) {
  if (!apiKey?.startsWith('rak_')) return null
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
      const valid = new Set((await database.collection('clients').find({ apiKey: { $in: keys } }, { projection: { apiKey: 1, _id: 0 } }).toArray()).map(d => d.apiKey))
      for (const k of keys) if (!valid.has(k)) evictCache(k)
    } catch {}
  }, KEY_CHECK_INTERVAL_MS)
}

function extractApiKey(req) { const h = req.headers['authorization'] || ''; return h.startsWith('Bearer ') ? h.slice(7).trim() : null }

async function requireClientKey(req, res, next) {
  const k = extractApiKey(req) || req.body?.apiKey
  if (!k) return res.status(401).json({ error: 'Missing API key' })
  const client = await verifyApiKey(k)
  if (!client) return res.status(401).json({ error: 'Invalid or expired API key' })
  req.client = client; next()
}

function requireAdminKey(req, res, next) {
  const k = extractApiKey(req)
  if (!k || k !== ADMIN_API_KEY) return res.status(401).json({ error: 'Unauthorized' })
  next()
}

function generateApiKey() { return `rak_${crypto.randomBytes(32).toString('hex')}` }
function generateTitle(q) { const c = q.trim().replace(/[?!.]+$/, ''); return c.length > 50 ? c.slice(0, 50) + '...' : c }

async function saveConversationMessage(clientId, conversationId, q, answer, sources) {
  try {
    const db2 = await getChatDb(), col = db2.collection('conversations'), now = new Date()
    const userMsg = { role: 'user', content: q, timestamp: now }
    const aMsg = { role: 'assistant', content: answer, sources: sources.map(s => ({ source_file: s.source_file, score: s.score })), timestamp: now }
    let activeId = conversationId || null
    if (activeId) {
      const upd = await col.findOneAndUpdate({ _id: new ObjectId(activeId), clientId }, { $push: { messages: { $each: [userMsg, aMsg] } }, $set: { updatedAt: now } }, { returnDocument: 'after', projection: { _id: 1 } })
      if (!upd) activeId = null
    }
    if (!activeId) {
      const r = await col.insertOne({ clientId, title: generateTitle(q), messages: [userMsg, aMsg], createdAt: now, updatedAt: now })
      activeId = r.insertedId.toString()
    }
    return activeId
  } catch (e) { console.warn('[saveConversationMessage]', e.message); return conversationId || null }
}

function buildDedupedSources(hits) {
  const seen = new Set(), out = []
  for (const h of hits) {
    if (h.metadata?._expansionRow) continue
    const key = h.metadata?.measure ? `measure:${h.metadata.measure.toLowerCase().trim()}` :
      h.metadata?.url ? `url:${h.metadata.url.toLowerCase().trim()}` :
      `text:${(h.metadata?.focus_sentence || h.text || '').trim().slice(0, 80).toLowerCase()}`
    if (seen.has(key)) continue
    seen.add(key)
    const preview = h.metadata?.focus_sentence
      ? trimPreviewToSentence(h.metadata.focus_sentence, 200)
      : trimPreviewToSentence(h.text || '', 200)
    out.push({
      source_file: h.source_file || 'unknown',
      chunk_index: h.chunk_index ?? 0,
      score: typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
      measure: h.metadata?.measure || null,
      table: h.metadata?.table || null,
      section: h.metadata?.section_heading || null,
      preview,
    })
  }
  return out
}

app.get('/health', (req, res) => res.json({ ok: true, service: 'ask-data', chunkCacheSize: CHUNK_CACHE.size, responseCacheSize: RESPONSE_CACHE.size, circuitOpen: askedataCircuitOpen() }))

app.post('/client/verify', async (req, res) => {
  try {
    const k = extractApiKey(req) || req.body?.apiKey
    if (!k) return res.status(400).json({ valid: false, error: 'apiKey required' })
    const client = await verifyApiKey(k)
    if (!client) return res.status(401).json({ valid: false, error: 'Invalid key' })
    res.json({ valid: true, client })
  } catch (e) { res.status(500).json({ valid: false, error: e.message }) }
})

app.post('/admin/clients', requireAdminKey, async (req, res) => {
  try {
    let { name, clientId, apiKey } = req.body
    if (!name || !clientId) return res.status(400).json({ error: 'name and clientId required' })
    if (!apiKey) apiKey = generateApiKey()
    else if (!apiKey.startsWith('rak_')) return res.status(400).json({ error: 'apiKey must start with "rak_"' })
    const database = await getDb(), col = database.collection('clients')
    const existing = await col.findOne({ $or: [{ clientId }, { apiKey }] })
    if (existing) return res.status(409).json({ error: `Conflict on ${existing.clientId === clientId ? 'clientId' : 'apiKey'}` })
    const now = new Date().toISOString()
    const doc = { name: name.trim(), clientId: clientId.trim().toLowerCase(), apiKey, apiKeyRotatedAt: now, folderLink: '', sourceType: 'google-drive', status: 'idle', documentsCount: 0, autoSync: false, watchIntervalMs: 300000, lastRunAt: null, lastError: null, createdAt: now, updatedAt: now }
    const r = await col.insertOne(doc)
    res.status(201).json({ ...doc, _id: r.insertedId })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.get('/admin/clients', requireAdminKey, async (req, res) => {
  try { const db2 = await getDb(); res.json({ clients: await db2.collection('clients').find({}, { projection: { apiKey: 0 } }).sort({ createdAt: -1 }).toArray() }) }
  catch (e) { res.status(500).json({ error: e.message }) }
})

app.get('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const db2 = await getDb(), client = await db2.collection('clients').findOne({ clientId: req.params.clientId })
    if (!client) return res.status(404).json({ error: 'Not found' })
    res.json(client)
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/admin/clients/:clientId/regenerate-key', requireAdminKey, async (req, res) => {
  try {
    const db2 = await getDb(), col = db2.collection('clients')
    const old = await col.findOne({ clientId: req.params.clientId }, { projection: { apiKey: 1 } })
    if (!old) return res.status(404).json({ error: 'Not found' })
    const newKey = generateApiKey(), now = new Date().toISOString()
    if (old.apiKey) evictCache(old.apiKey)
    await col.findOneAndUpdate({ clientId: req.params.clientId }, { $set: { apiKey: newKey, apiKeyRotatedAt: now, updatedAt: now } })
    res.json({ success: true, clientId: req.params.clientId, newApiKey: newKey, apiKeyRotatedAt: now })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.patch('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const db2 = await getDb()
    const updates = { ...req.body, updatedAt: new Date().toISOString() }
    if (updates.apiKey !== undefined) {
      if (!updates.apiKey.startsWith('rak_')) return res.status(400).json({ error: 'apiKey must start with "rak_"' })
      const old = await db2.collection('clients').findOne({ clientId: req.params.clientId }, { projection: { apiKey: 1 } })
      if (old?.apiKey) evictCache(old.apiKey)
      updates.apiKeyRotatedAt = new Date().toISOString()
    }
    const r = await db2.collection('clients').findOneAndUpdate({ clientId: req.params.clientId }, { $set: updates }, { returnDocument: 'after' })
    if (!r) return res.status(404).json({ error: 'Not found' })
    res.json(r)
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.delete('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
  try {
    const { clientId } = req.params, db2 = await getDb()
    const client = await db2.collection('clients').findOne({ clientId })
    if (!client) return res.status(404).json({ error: 'Not found' })
    if (client.apiKey) evictCache(client.apiKey)
    await db2.collection('clients').deleteOne({ clientId })
    invalidateChunkCache(clientId)
    const deleted = [], failed = []
    if (blobServiceClient) {
      try {
        const container = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
        for (const pfx of [`raw/${clientId}/`, `meta/${clientId}/`]) {
          for await (const blob of container.listBlobsFlat({ prefix: pfx })) {
            try { await container.deleteBlob(blob.name); deleted.push(blob.name) }
            catch (e) { failed.push({ name: blob.name, error: e.message }) }
          }
        }
      } catch (e) { failed.push({ name: 'azure', error: e.message }) }
    }
    res.json({ ok: true, deleted: clientId, blobsDeleted: deleted.length, blobsFailed: failed.length ? failed : undefined })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/admin/clients/:clientId/invalidate-cache', requireAdminKey, (req, res) => {
  invalidateChunkCache(req.params.clientId)
  RESPONSE_CACHE.clear()
  res.json({ ok: true, clientId: req.params.clientId })
})

app.post('/client/login', async (req, res) => {
  try {
    const k = extractApiKey(req) || req.body?.apiKey
    if (!k) return res.status(400).json({ error: 'apiKey required' })
    const client = await verifyApiKey(k)
    if (!client) return res.status(401).json({ error: 'Invalid key' })
    if (blobServiceClient) loadChunksForClient(client.clientId).catch(e => console.warn(`[login warmup] ${e.message}`))
    res.json({ ok: true, client })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/login', async (req, res) => {
  try {
    const k = extractApiKey(req) || req.body?.apiKey
    if (!k) return res.status(400).json({ error: 'apiKey required' })
    const client = await verifyApiKey(k)
    if (!client) return res.status(401).json({ error: 'Invalid key' })
    if (blobServiceClient) loadChunksForClient(client.clientId).catch(e => console.warn(`[chat/login warmup] ${e.message}`))
    res.json({ ok: true, client })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.get('/client/me', requireClientKey, async (req, res) => {
  try {
    const db2 = await getDb(), client = await db2.collection('clients').findOne({ clientId: req.client.clientId }, { projection: { apiKey: 0 } })
    if (!client) return res.status(404).json({ error: 'Not found' })
    res.json(client)
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/conversations', requireClientKey, async (req, res) => {
  try {
    const db2 = await getChatDb(), now = new Date()
    const conv = { clientId: req.client.clientId, title: req.body.title || 'New Conversation', messages: [], createdAt: now, updatedAt: now }
    const r = await db2.collection('conversations').insertOne(conv)
    res.status(201).json({ ...conv, _id: r.insertedId })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/conversations/list', requireClientKey, async (req, res) => {
  try {
    const db2 = await getChatDb()
    res.json({ conversations: await db2.collection('conversations').find({ clientId: req.client.clientId }, { projection: { messages: 0 } }).sort({ updatedAt: -1 }).toArray() })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/conversations/get', requireClientKey, async (req, res) => {
  try {
    const { conversationId } = req.body
    if (!conversationId) return res.status(400).json({ error: 'conversationId required' })
    const db2 = await getChatDb()
    const conv = await db2.collection('conversations').findOne({ _id: new ObjectId(conversationId), clientId: req.client.clientId })
    if (!conv) return res.status(404).json({ error: 'Not found' })
    res.json(conv)
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/conversations/rename', requireClientKey, async (req, res) => {
  try {
    const { conversationId, title } = req.body
    if (!conversationId || !title) return res.status(400).json({ error: 'conversationId and title required' })
    const db2 = await getChatDb()
    const r = await db2.collection('conversations').findOneAndUpdate({ _id: new ObjectId(conversationId), clientId: req.client.clientId }, { $set: { title: title.trim(), updatedAt: new Date() } }, { returnDocument: 'after', projection: { messages: 0 } })
    if (!r) return res.status(404).json({ error: 'Not found' })
    res.json(r)
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/conversations/delete', requireClientKey, async (req, res) => {
  try {
    const { conversationId } = req.body
    if (!conversationId) return res.status(400).json({ error: 'conversationId required' })
    const db2 = await getChatDb()
    const r = await db2.collection('conversations').deleteOne({ _id: new ObjectId(conversationId), clientId: req.client.clientId })
    if (r.deletedCount === 0) return res.status(404).json({ error: 'Not found' })
    res.json({ ok: true, deleted: conversationId })
  } catch (e) { res.status(500).json({ error: e.message }) }
})

app.post('/chat/message', requireClientKey, withRequestTimeout(async (req, res) => {
  try {
    const { query, topK = 5, conversationId } = req.body
    if (!query?.trim()) return res.status(400).json({ error: 'query required' })
    const { clientId, name } = req.client

    const intentResult = resolveIntent(query.trim())
    if (intentResult) {
      const cid = await saveConversationMessage(clientId, conversationId || null, query.trim(), intentResult.response, [])
      return res.json({ answer: intentResult.response, sources: [], relatedKeywords: [], relatedMetrics: [], conversationId: cid, client: { clientId, name } })
    }

    const val = validateQuery(query)
    if (!val.valid) return res.json({ answer: val.message, sources: [], relatedKeywords: [], relatedMetrics: [], conversationId: conversationId || null, client: { clientId, name } })

    const cacheKey = getCacheKey(clientId, query)
    const cached = responseCacheGet(cacheKey)
    if (cached) {
      const cid = await saveConversationMessage(clientId, conversationId || null, query.trim(), cached.answer, cached.sources || [])
      return res.json({ ...cached, cached: true, conversationId: cid })
    }

    if (IN_FLIGHT.has(cacheKey)) {
      try { const r = await IN_FLIGHT.get(cacheKey); const cid = await saveConversationMessage(clientId, conversationId || null, query.trim(), r.answer, r.sources || []); return res.json({ ...r, conversationId: cid }) } catch {}
    }

    const reqPromise = (async () => {
      const { chunks, invertedIndex, docType } = await loadChunksForClient(clientId)
      if (!chunks?.length) return { answer: 'No documents found. Please ingest documents first.', sources: [], relatedKeywords: [], relatedMetrics: [], client: { clientId, name } }

      let pq = applyTypos(query.trim())
      pq = applySynonyms(pq)
      pq = fuzzyCorrectQuery(pq, chunks)
      pq = await preprocessQuery(pq)

      const queryTerms = pq.toLowerCase().replace(/^(define|what\s+is|explain|tell\s+me\s+about)\s+/i, '').trim().split(/\s+/).filter(w => w.length > 2)
      const termFoundInChunks = queryTerms.some(term =>
        chunks.some(c => (c.text || '').toLowerCase().includes(term) || (c.metadata?.measure || '').toLowerCase().includes(term))
      )

      if (!termFoundInChunks && isOutOfScope(pq)) {
        return { answer: 'I can only answer questions based on your uploaded documents. That topic isn\'t covered in them. Please ask something related to your enterprise data, KPIs, or reports.', sources: [], relatedKeywords: [], relatedMetrics: [], client: { clientId, name } }
      }

      const eDocType = docType || 'mixed'
      const eIntent = detectQueryIntent(pq)

      if (eIntent === 'all_urls') {
        const uc = chunks.filter(c => /https?:\/\/\S+/.test(c.text || ''))
        const entries = extractAllUrlsFromChunks(uc)
        return { answer: entries.length ? entries.map(e => `**${e.name}:** ${e.url}`).join('\n') : 'No URLs found.', sources: buildDedupedSources(uc.slice(0, 5)), relatedKeywords: [], relatedMetrics: [], client: { clientId, name } }
      }

      const multi = detectMultiTopicQuery(pq)
      if (multi.isMulti) {
        const answer = await handleMultiTopicQuery(multi.topics, multi.mode, chunks, Math.min(topK, MAX_HITS_GLOBAL), invertedIndex, eDocType)
        return { answer, sources: [], relatedKeywords: [], relatedMetrics: [], client: { clientId, name } }
      }

      let hits = await retrieveChunks(pq, chunks, Math.min(topK, MAX_HITS_GLOBAL), invertedIndex, eDocType)
      if (!hits.length) hits = relaxedKeywordSearch(pq, chunks, 32, invertedIndex)
      if (!hits.length) return { answer: 'I could not find relevant information about this in your documents. Please try rephrasing your question.', sources: [], relatedKeywords: [], relatedMetrics: [], client: { clientId, name } }

      const topHitScore = hits[0]?._score ?? 0
      const minRelevance = getMinRelevanceThreshold(eDocType, eIntent)
      if (topHitScore < minRelevance) {
        return { answer: 'I could not find relevant information about this in your documents. Please try rephrasing your question.', sources: [], relatedKeywords: [], relatedMetrics: [], client: { clientId, name } }
      }

      const answer = await generateAnswerWithFallback(pq, hits, eIntent, eDocType, chunks, invertedIndex, Math.min(topK, MAX_HITS_GLOBAL))
      const sources = buildDedupedSources(hits)
      const subject = extractSubject(pq)
      const related = buildRelatedKeywords(subject, hits, chunks, invertedIndex, RELATED_KEYWORDS_COUNT)
      const hasMeasureData = chunks.some(c => c.metadata?.measure && !c.metadata._expansionRow)
      const isDefinitionOrGeneral = ['definition', 'general', 'calculation'].includes(eIntent)
      const relatedMetrics = (hasMeasureData && isDefinitionOrGeneral) ? buildRelatedMetrics(subject, hits, chunks, 6) : []

      return { answer, sources, relatedKeywords: related, relatedMetrics, client: { clientId, name } }
    })()

    IN_FLIGHT.set(cacheKey, reqPromise)
    let result
    try { result = await reqPromise } finally { IN_FLIGHT.delete(cacheKey) }

    if (result.answer?.length > 15) responseCacheSet(cacheKey, result)
    const cid = await saveConversationMessage(clientId, conversationId || null, query.trim(), result.answer, result.sources || [])
    res.json({ ...result, conversationId: cid })
  } catch (e) { console.error('[chat/message]', e.message); if (!res.headersSent) res.status(500).json({ error: e.message }) }
}))

app.use((err, req, res, next) => { console.error('[global error]', err); if (!res.headersSent) res.status(500).json({ error: 'Unexpected error.' }) })

if (process.env.VERCEL !== '1') {
  const PORT = process.env.PORT || 4000
  app.listen(PORT, () => {
    console.log(`Service on port ${PORT}`)
    startApiKeyHealthChecker()
    warmupChunkCaches()
  })
} else {
  console.log('Running on Vercel')
}

module.exports = app
