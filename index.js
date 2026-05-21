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
const stringSimilarity = require('string-similarity')
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
const AZURE_EMBED_ENDPOINT = process.env.AZURE_EMBED_ENDPOINT || ''
const AZURE_EMBED_KEY = process.env.AZURE_EMBED_KEY || ''
const AZURE_EMBED_MODEL = process.env.AZURE_EMBED_MODEL || 'text-embedding-3-small'
const EMBED_TIMEOUT_MS = parseInt(process.env.EMBED_TIMEOUT_MS || '10000', 10)
const EMBED_POOL_LIMIT = parseInt(process.env.EMBED_POOL_LIMIT || '20', 10)
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || '60000', 10)
const WARMUP_CLIENT_IDS = (process.env.WARMUP_CLIENT_IDS || '').split(',').map(s => s.trim()).filter(Boolean)
const RERANKER_ENDPOINT = process.env.RERANKER_ENDPOINT || ''
const RERANKER_KEY = process.env.RERANKER_KEY || ''
const RERANKER_TIMEOUT_MS = parseInt(process.env.RERANKER_TIMEOUT_MS || '8000', 10)
const RAW_PREFIX = 'raw'
const CHUNK_SIZE = 1200
const CHUNK_OVERLAP = 2
const BLOB_CONCURRENCY = parseInt(process.env.BLOB_CONCURRENCY || '8', 10)
const CHUNK_CACHE_TTL = parseInt(process.env.CHUNK_CACHE_TTL_MS || '300000', 10)
const MAX_HITS_GLOBAL = 50
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
const SYNONYM_MAP = [
{ pattern: /\bapp(lication)?\s+count\b/i, canonical: 'application count' },
{ pattern: /\btotal\s+app(lication)?\s+count\b/i, canonical: 'application count' },
{ pattern: /\bnumber\s+of\s+app(lication)?s\b/i, canonical: 'application count' },
{ pattern: /\bapp(lication)?\s+volume\b/i, canonical: 'application count' },
{ pattern: /\btotal\s+submitted\s+app(lication)?s\b/i, canonical: 'application count' },
{ pattern: /\btotal\s+app(lication)?s\b/i, canonical: 'application count' },
{ pattern: /\bsubmitted\s+app(lication)?s\b/i, canonical: 'application count' },
{ pattern: /\bocc(upancy)?\s+rate\b/i, canonical: 'occupancy rate' },
{ pattern: /\bocc(upancy)?\s+formula\b/i, canonical: 'occupancy formula' },
{ pattern: /\blead\s+acq(uisition)?\s+cost\b/i, canonical: 'lead acquisition cost' },
{ pattern: /\blead\s+cost\b/i, canonical: 'lead acquisition cost' },
]
function applySynonyms(query) {
let q = query
for (const { pattern, canonical } of SYNONYM_MAP) {
q = q.replace(pattern, canonical)
}
return q
}
const TYPO_MAP = {
ehat: 'what',
waht: 'what',
whta: 'what',
whar: 'what',
hwo: 'how',
hoe: 'how',
difine: 'define',
definr: 'define',
defien: 'define',
defne: 'define',
deifne: 'define',
expain: 'explain',
expalin: 'explain',
explian: 'explain',
wht: 'what',
shwo: 'show',
lsit: 'list',
lits: 'list',
}
function applyTypos(query) {
return query
.split(/\s+/)
.map(w => {
const lower = w.toLowerCase()
return TYPO_MAP[lower] !== undefined ? TYPO_MAP[lower] : w
})
.join(' ')
}
function levenshteinDistance(a, b) {
const m = a.length
const n = b.length
const dp = Array.from({ length: m + 1 }, (_, i) => Array.from({ length: n + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0)))
for (let i = 1; i <= m; i++) {
for (let j = 1; j <= n; j++) {
if (a[i - 1] === b[j - 1]) dp[i][j] = dp[i - 1][j - 1]
else dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1])
}
}
return dp[m][n]
}
function levenshteinSimilarity(a, b) {
if (!a && !b) return 1
if (!a || !b) return 0
const dist = levenshteinDistance(a.toLowerCase(), b.toLowerCase())
return 1 - dist / Math.max(a.length, b.length)
}
function normalizeQueryForCache(query) {
return applySynonyms(query)
.toLowerCase()
.trim()
.replace(/\bweek\s+(\d)\b/g, (_, n) => `week 0${n}`)
.replace(/^(what\s+is\s+(the\s+)?(definition|meaning)\s+(of|for|to)\s+)/i, '')
.replace(/^(define\s+(the\s+)?)/i, '')
.replace(/^(explain\s+(the\s+)?)/i, '')
.replace(/^(tell\s+me\s+about\s+(the\s+)?)/i, '')
.replace(/^(what\s+are\s+(the\s+)?)/i, '')
.replace(/^(what\s+is\s+)/i, '')
.replace(/^(how\s+(do\s+you\s+|is\s+|are\s+)?calculate\s+(the\s+)?)/i, '')
.replace(/^(describe\s+(the\s+|me\s+)?)/i, '')
.replace(/^(meaning\s+of\s+(the\s+)?)/i, '')
.replace(/[?!.]+$/, '')
.replace(/\s+/g, ' ')
.trim()
}
function getCacheKey(clientId, query) {
return `${clientId}:${normalizeQueryForCache(query)}`
}
const IN_FLIGHT = new Map()
let askedataActiveCount = 0
const ASKDATA_MAX_CONCURRENT = 3
const askedataQueue = []
function runWithAskedataLimit(fn) {
return new Promise((resolve, reject) => {
function tryRun() {
if (askedataActiveCount < ASKDATA_MAX_CONCURRENT) {
askedataActiveCount++
Promise.resolve().then(fn).then(
result => { askedataActiveCount--; drainAskedataQueue(); resolve(result) },
err => { askedataActiveCount--; drainAskedataQueue(); reject(err) }
)
} else {
askedataQueue.push(tryRun)
}
}
tryRun()
})
}
function drainAskedataQueue() {
if (askedataQueue.length > 0 && askedataActiveCount < ASKDATA_MAX_CONCURRENT) {
askedataQueue.shift()()
}
}
let askedataFailures = 0
let askedataBlockedUntil = 0
function askedataCircuitOpen() {
if (Date.now() < askedataBlockedUntil) return true
if (askedataBlockedUntil > 0) { askedataBlockedUntil = 0; askedataFailures = 0 }
return false
}
function askedataRecordSuccess() { askedataFailures = 0; askedataBlockedUntil = 0 }
function askedataRecordFailure() {
askedataFailures++
if (askedataFailures >= 3) {
askedataBlockedUntil = Date.now() + 30000
console.error('[ASKDATA] Circuit breaker OPEN for 30s')
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
function normalizeQuery(query) {
return applySynonyms(query)
.toLowerCase()
.trim()
.replace(/\bweek\s+(\d)\b/g, (_, n) => `week 0${n}`)
.replace(/[?!.]+$/, '')
.replace(/\s+/g, ' ')
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
const q = normalizeQuery(query)
if (/^(hi|hello|hey|howdy|greetings|good\s+(morning|afternoon|evening)|how\s+are\s+you)\b/.test(q)) return 'greeting'
if (/\b(url|link|dashboard|power\s*bi|report\s+url)\b/.test(q)) return 'url_lookup'
if (
/\b(formula|equation|calculate|calculation|calculated|computed|derived|how\s+is\s+.+\s+calculated|how\s+are\s+.+\s+calculated)\b/i.test(q) ||
/how\s+(is|are|was|were)\s+.+\s+(calculated|computed|determined|derived)/i.test(q) ||
/what\s+is\s+the\s+(formula|calculation)\s+for/i.test(q) ||
/how\s+do\s+you\s+(calculate|compute)/i.test(q)
) return 'calculation'
if (
/^(what\s+is\s+(the\s+)?(definition|meaning)\s+(of|for|to)\s+)/i.test(q) ||
/^(define\s+(the\s+)?)/i.test(q) ||
/^(what\s+(is|are)\s+(an?\s+|the\s+)?)/i.test(q) ||
/^(explain\s+(the\s+)?)/i.test(q) ||
/^(tell\s+me\s+about\s+(the\s+)?)/i.test(q) ||
/^(describe\s+(the\s+|me\s+)?)/i.test(q) ||
/^(meaning\s+of\s+)/i.test(q) ||
/\b(definition|meaning)\b/i.test(q)
) return 'definition'
if (/\b(vs|versus|difference|compare|between)\b/.test(q)) return 'comparison'
if (/^(show|list|find|get|fetch|give)\s+(me\s+)?|^how\s+many\s+/.test(q)) return 'lookup'
return 'general'
}
function detectMultiTopicQuery(query) {
const q = query.trim()
const diffPatterns = [
/^(?:what\s+is\s+the\s+)?difference\s+between\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
/^compare\s+(.+?)\s+(?:vs\.?|versus|and)\s+(.+?)[\s?]*$/i,
/^(.+?)\s+vs\.?\s+(.+?)[\s?]*$/i,
]
const andSplitPatterns = [
/^what\s+is\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
/^what\s+are\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
/^define\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
/^explain\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
/^tell\s+me\s+about\s+(.+?)\s+and\s+(.+?)[\s?]*$/i,
/^(.+?)\s+and\s+(.+?)[\s?]*$/i,
]
for (const p of diffPatterns) {
const m = q.match(p)
if (m) {
const a = m[1].trim().replace(/^(what\s+is\s+|the\s+)/i, '').trim()
const b = m[2].trim().replace(/^(what\s+is\s+|the\s+)/i, '').trim()
if (a.length > 1 && b.length > 1) return { isMulti: true, topics: [a, b], mode: 'comparison' }
}
}
for (const p of andSplitPatterns) {
const m = q.match(p)
if (m) {
const a = m[1].trim().replace(/^(what\s+is\s+|what\s+are\s+|define\s+|the\s+)/i, '').trim()
const b = m[2].trim().replace(/^(what\s+is\s+|what\s+are\s+|define\s+|the\s+)/i, '').trim()
const stopWords = new Set(['is', 'are', 'was', 'were', 'it', 'this', 'that', 'its', 'my', 'your'])
if (a.length > 1 && b.length > 1 && !stopWords.has(a.toLowerCase()) && !stopWords.has(b.toLowerCase())) {
return { isMulti: true, topics: [a, b], mode: 'multi_definition' }
}
}
}
return { isMulti: false, topics: [], mode: null }
}
function extractSubject(query) {
const normalized = applySynonyms(query)
const q = normalizeQuery(normalized)
const patterns = [
/^what\s+is\s+(?:the\s+)?(?:definition|meaning)\s+(?:of|for|to)\s+(?:an?\s+|the\s+)?(.+)$/i,
/^what\s+(?:is|are)\s+(?:the\s+)?(?:definition|meaning)\s+(?:of|for|to)\s+(?:an?\s+|the\s+)?(.+)$/i,
/^define\s+(?:an?\s+|the\s+)?(.+)$/i,
/^explain\s+(?:an?\s+|the\s+)?how\s+(.+?)\s+(?:is\s+)?calculated$/i,
/^explain\s+(?:an?\s+|the\s+)?(.+)$/i,
/^tell\s+me\s+about\s+(?:an?\s+|the\s+)?(.+)$/i,
/^describe\s+(?:me\s+)?(?:an?\s+|the\s+)?(.+)$/i,
/^meaning\s+of\s+(?:an?\s+|the\s+)?(.+)$/i,
/^(?:what\s+is\s+the\s+)?meaning\s+of\s+(?:an?\s+|the\s+)?(.+)$/i,
/^describe\s+(?:an?\s+|the\s+)?(.+)$/i,
/^how\s+is\s+(.+?)\s+(?:calculated|defined|measured|computed)$/i,
/^how\s+are\s+(.+?)\s+(?:calculated|defined|measured|computed)$/i,
/^what\s+is\s+the\s+formula\s+for\s+(?:calculating\s+)?(?:an?\s+|the\s+)?(.+)$/i,
/^how\s+(?:do\s+you\s+)?calculate\s+(?:an?\s+|the\s+)?(.+)$/i,
/^(?:formula|equation)\s+(?:for|of)\s+(?:an?\s+|the\s+)?(.+)$/i,
/^(.+?)\s+(?:formula|equation|calculation)$/i,
/^what\s+does\s+(.+?)\s+represent/i,
/^what\s+is\s+the\s+purpose\s+of\s+(?:the\s+)?(.+?)\s+(?:attribute|measure|field|column)$/i,
/^compare\s+(.+?)\s+(?:vs|versus)\s+(.+)$/i,
/^difference\s+between\s+(.+?)\s+and\s+(.+)$/i,
/^what\s+(?:is|are)\s+(?:an?\s+|the\s+)?(.+)$/i,
/^(?:what\s+is\s+)?(.+)$/i,
]
for (const p of patterns) {
const m = q.match(p)
if (m) {
const subject = m[1].trim().replace(/[?!.]+$/, '').trim()
if (subject.length > 0) return subject
}
}
return q.replace(/[?!.]+$/, '').trim()
}
function extractUrlKeywords(query) {
const stopWords = new Set(['power', 'bi', 'report', 'url', 'link', 'for', 'the', 'a', 'an', 'of', 'in', 'get', 'me', 'show', 'give', 'find', 'fetch'])
return query.toLowerCase().replace(/[^\w\s-]/g, ' ').split(/\s+/).filter(w => w.length > 1 && !stopWords.has(w))
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
function trimToCompleteSentence(text, maxLen = 1200) {
if (!text || text.length <= maxLen) return text
const truncated = text.slice(0, maxLen)
const lastPeriod = Math.max(truncated.lastIndexOf('. '), truncated.lastIndexOf('.\n'), truncated.lastIndexOf('.'))
if (lastPeriod > maxLen * 0.5) return truncated.slice(0, lastPeriod + 1).trim()
return truncated.trim()
}
function ensureSinglePeriod(text) {
if (!text) return ''
return text.replace(/\.{2,}/g, '.').replace(/\.\s*\./g, '.').trim()
}
function capFirst(str) {
if (!str) return ''
return str.charAt(0).toUpperCase() + str.slice(1)
}
function extractFormulaFromText(text) {
if (!text) return ''
const patterns = [
/formula\s*:\s*([^\n.]+)/i,
/calculated\s+as\s+([^\n.]+)/i,
/computed\s+as\s+([^\n.]+)/i,
/how\s+to\s+calculate\s+[^:]+:\s*([^\n.]+)/i,
/([a-z0-9\s%()#]+\s*\/\s*[a-z0-9\s%()#]+)/i,
/([a-z0-9\s%()#]+\s*=\s*[a-z0-9\s%()#+\-*/]+)/i,
/divided\s+by\s+([^\n.]+)/i,
/sum\s+of\s+([^\n.]+)/i,
/multiplied\s+by\s+([^\n.]+)/i,
]
for (const p of patterns) {
const m = text.match(p)
if (m && m[1] && m[1].trim().length > 3) return m[1].trim()
}
return ''
}
const NEGATIVE_PAIRS = [
['non-recurring', 'recurring'],
['non recurring', 'recurring'],
['denied', 'approved'],
['inactive', 'active'],
['rejected', 'accepted'],
['unapproved', 'approved'],
['unpaid', 'paid'],
['cancelled', 'active'],
['canceled', 'active'],
['delinquent', 'current'],
['non-', ''],
]
function computeNegativePenalty(querySubject, chunkText) {
const qs = querySubject.toLowerCase()
const ct = chunkText.toLowerCase()
let penalty = 0
for (const [negTerm, posTerm] of NEGATIVE_PAIRS) {
if (posTerm === null || posTerm === undefined) continue
const queryHasPositive = posTerm.length > 0 && new RegExp(`\\b${escapeRegex(posTerm)}\\b`, 'i').test(qs)
const queryHasNegative = new RegExp(`\\b${escapeRegex(negTerm)}\\b`, 'i').test(qs)
if (queryHasPositive && !queryHasNegative) {
if (new RegExp(`\\b${escapeRegex(negTerm)}\\b`, 'i').test(ct)) {
penalty += 30
}
}
if (queryHasNegative) {
if (posTerm.length > 0 && !new RegExp(`\\b${escapeRegex(negTerm)}\\b`, 'i').test(ct) && new RegExp(`\\b${escapeRegex(posTerm)}\\b`, 'i').test(ct)) {
penalty += 20
}
}
}
return penalty
}
function buildVocabulary(chunks) {
const vocab = new Set()
const stopWords = new Set(['is', 'the', 'a', 'an', 'of', 'in', 'for', 'to', 'at', 'by', 'as', 'on', 'or', 'and', 'be', 'it', 'its', 'with', 'that', 'this', 'from', 'are', 'was', 'were'])
for (const chunk of chunks) {
const words = (chunk.text || '').toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)
for (const w of words) {
if (w.length >= 3 && !stopWords.has(w)) vocab.add(w)
}
if (chunk.metadata && chunk.metadata.measure) {
const measureWords = chunk.metadata.measure.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/)
for (const w of measureWords) {
if (w.length >= 3 && !stopWords.has(w)) vocab.add(w)
}
}
}
return [...vocab]
}
const DOMAIN_SHORT_SAFELIST = new Set([
'count', 'rate', 'rent', 'cost', 'date', 'type', 'name', 'unit', 'term', 'area',
'base', 'gross', 'net', 'avg', 'sum', 'min', 'max', 'ytd', 'mtd', 'per', 'fee',
'tax', 'due', 'paid', 'void', 'open', 'loss', 'gain', 'flow', 'days', 'beds',
'bath', 'sqft', 'tier', 'band', 'code', 'flag', 'rank', 'sort', 'key', 'ref',
])
function fuzzyCorrectQuery(query, chunks) {
if (!chunks || chunks.length === 0) return query
const vocabulary = buildVocabulary(chunks)
if (vocabulary.length === 0) return query
const stopWords = new Set(['what', 'is', 'are', 'how', 'the', 'a', 'an', 'of', 'in', 'for', 'to', 'at', 'by', 'as', 'on', 'or', 'and', 'define', 'explain', 'show', 'find', 'get', 'list', 'give'])
const words = query.split(/\s+/)
const corrected = words.map(word => {
const wordLower = word.toLowerCase()
if (stopWords.has(wordLower)) return word
if (DOMAIN_SHORT_SAFELIST.has(wordLower)) return word
if (wordLower.length < 6) return word
if (vocabulary.includes(wordLower)) return word
const { bestMatch } = stringSimilarity.findBestMatch(wordLower, vocabulary)
const levSim = levenshteinSimilarity(wordLower, bestMatch.target)
const combinedScore = bestMatch.rating * 0.6 + levSim * 0.4
if (combinedScore >= 0.72 && bestMatch.target !== wordLower) {
console.log(`[fuzzyCorrect] "${word}" → "${bestMatch.target}" (combined: ${combinedScore.toFixed(3)}, trigram: ${bestMatch.rating.toFixed(3)}, lev: ${levSim.toFixed(3)})`)
return bestMatch.target
}
return word
})
return corrected.join(' ')
}
function needsQueryRewrite(query) {
const trimmed = query.trim()
const words = trimmed.split(/\s+/).filter(Boolean)
if (words.length <= 2) return true
if (/[^\x00-\x7F]/.test(trimmed) && words.length < 5) return true
if (/(.)\1{3,}/.test(trimmed)) return true
const commonTypos = /\b(ocupan|occup[ae]ncy|occupncy|appl?ic|applicton|applcation|leas[ei]|tennat|tentant|vacnt|vacanc|porperty|proprty|reveneu|revenu[^e]|anuall|anual|montly|mounthly|efftive|efective|efftiv|ehat|difine|definr)\b/i
if (commonTypos.test(trimmed)) return true
if (words.length < 4 && !/\b(what|how|define|explain|formula|calculate|list|show|find|url|link)\b/i.test(trimmed)) return true
const hasNoVerb = !/\b(is|are|was|were|what|how|why|when|where|who|define|explain|calculate|show|list|find|get|give|tell)\b/i.test(trimmed)
if (hasNoVerb && words.length < 6) return true
return false
}
async function rewriteQueryWithAskdata2(query) {
if (!ASKDATA2_ENDPOINT || !ASKDATA2_KEY) return query
try {
const response = await fetchWithTimeout(
ASKDATA2_ENDPOINT,
{
method: 'POST',
headers: {
'Content-Type': 'application/json',
'Authorization': `Bearer ${ASKDATA2_KEY}`,
'Accept': 'application/json',
},
body: JSON.stringify({
model: ASKDATA2_MODEL,
messages: [
{
role: 'system',
content: 'You are a query rewriter for a real estate analytics data dictionary RAG system. Fix spelling, grammar, ambiguity, and structure. Expand abbreviations (app→application, occ→occupancy). Normalize week numbers (week 5→week 05). Return ONLY the rewritten query as plain text. No explanation. If already correct, return unchanged.',
},
{ role: 'user', content: query },
],
max_tokens: 120,
temperature: 0.0,
top_p: 1.0,
stream: false,
}),
},
ASKDATA2_REWRITE_TIMEOUT_MS
)
if (!response.ok) return query
const data = await response.json()
const rewritten = (data.choices?.[0]?.message?.content || '').trim().replace(/^["']|["']$/g, '').trim()
if (!rewritten || rewritten.length < 3 || rewritten.length > query.length * 4) return query
if (rewritten.toLowerCase() !== query.toLowerCase()) {
console.log(`[QueryPipeline] After rewrite: "${rewritten}"`)
}
return rewritten
} catch (err) {
console.warn(`[rewriteQueryWithAskdata2] Bypassed (${err.message}), using original query`)
return query
}
}
async function preprocessQuery(query) {
if (!needsQueryRewrite(query)) return query
return rewriteQueryWithAskdata2(query)
}
async function callASKDATA(systemPrompt, userMessage, maxTokens = 1024) {
if (!ASKDATA_ENDPOINT || !ASKDATA_KEY) throw new Error('ASKDATA_ENDPOINT and ASKDATA_KEY are required')
if (askedataCircuitOpen()) throw new Error('ASKDATA temporarily unavailable')
return runWithAskedataLimit(async () => {
try {
const response = await fetchWithTimeout(
ASKDATA_ENDPOINT,
{
method: 'POST',
headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${ASKDATA_KEY}` },
body: JSON.stringify({
model: ASKDATA_MODEL,
messages: [
{ role: 'system', content: systemPrompt },
{ role: 'user', content: userMessage },
],
temperature: 0.1,
max_tokens: maxTokens,
}),
},
ASKDATA_TIMEOUT_MS
)
if (!response.ok) {
const errText = await response.text()
throw new Error(`ASKDATA error ${response.status}: ${errText}`)
}
const data = await response.json()
askedataRecordSuccess()
return data.choices?.[0]?.message?.content || ''
} catch (err) {
askedataRecordFailure()
throw err
}
})
}
async function callASKDATA2(systemPrompt, userMessage, maxTokens = 1024) {
if (!ASKDATA2_ENDPOINT || !ASKDATA2_KEY) throw new Error('ASKDATA2_ENDPOINT and ASKDATA2_KEY are required')
try {
const response = await fetchWithTimeout(
ASKDATA2_ENDPOINT,
{
method: 'POST',
headers: {
'Content-Type': 'application/json',
'Authorization': `Bearer ${ASKDATA2_KEY}`,
'Accept': 'application/json',
},
body: JSON.stringify({
model: ASKDATA2_MODEL,
messages: [
{ role: 'system', content: systemPrompt },
{ role: 'user', content: userMessage },
],
max_tokens: maxTokens,
temperature: 0.1,
top_p: 1.0,
stream: false,
}),
},
ASKDATA2_TIMEOUT_MS
)
if (!response.ok) {
const errText = await response.text()
throw new Error(`ASKDATA2 error ${response.status}: ${errText}`)
}
const data = await response.json()
return data.choices?.[0]?.message?.content || ''
} catch (err) {
console.error(`[ASKDATA2] Failed: ${err.message}`)
throw err
}
}
async function callBestAvailableEngine(systemPrompt, userMessage, maxTokens = 1024) {
let primaryError = null
if (ASKDATA_ENDPOINT && ASKDATA_KEY && !askedataCircuitOpen()) {
try {
const result = await callASKDATA(systemPrompt, userMessage, maxTokens)
if (result && result.trim().length >= 15) return result
primaryError = new Error('ASKDATA returned blank response')
} catch (err) {
primaryError = err
console.warn(`[ASKDATA] Failed, switching to ASKDATA2: ${err.message}`)
}
} else {
primaryError = new Error('ASKDATA unavailable (circuit open or not configured)')
}
if (ASKDATA2_ENDPOINT && ASKDATA2_KEY) {
try {
console.log(`[ASKDATA2] Activating (Reason: ${primaryError?.message})`)
const result = await callASKDATA2(systemPrompt, userMessage, maxTokens)
if (result && result.trim().length >= 15) return result
} catch (err) {
console.error(`[ASKDATA2] Also failed: ${err.message}`)
}
}
return ''
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
async function rerankerScore(query, texts) {
if (!RERANKER_ENDPOINT) return null
try {
const response = await fetchWithTimeout(
RERANKER_ENDPOINT,
{
method: 'POST',
headers: {
'Content-Type': 'application/json',
...(RERANKER_KEY ? { 'Authorization': `Bearer ${RERANKER_KEY}` } : {}),
},
body: JSON.stringify({ query, texts }),
},
RERANKER_TIMEOUT_MS
)
if (!response.ok) return null
const data = await response.json()
if (Array.isArray(data)) return data.map(d => (typeof d === 'number' ? d : d.score ?? 0))
if (Array.isArray(data.scores)) return data.scores
return null
} catch (err) {
console.warn(`[reranker] Failed: ${err.message}`)
return null
}
}
async function rerankChunks(query, chunks) {
if (!RERANKER_ENDPOINT || chunks.length === 0) return chunks
try {
const texts = chunks.map(c => (c.text || '').slice(0, 512))
const scores = await rerankerScore(query, texts)
if (!scores || scores.length !== chunks.length) return chunks
return chunks
.map((c, i) => ({ ...c, _rerankerScore: scores[i] }))
.sort((a, b) => b._rerankerScore - a._rerankerScore)
} catch (err) {
console.warn(`[rerankChunks] Reranker error, skipping: ${err.message}`)
return chunks
}
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
[/report.?name/i, 85],
[/\bname\b/, 70],
[/\btitle\b/, 50],
]
const TABLE_PATTERNS = [
[/\b(table|module|category|group|domain|section|workspace)\b/, 100],
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
if (colIdx.url !== undefined && colIdx.name === undefined) {
const usedIdx = new Set(Object.values(colIdx))
for (let i = 0; i < headers.length; i++) {
if (!usedIdx.has(i) && headers[i].trim()) { colIdx.name = i; break }
}
}
return colIdx
}
function extractSpreadsheet(buffer) {
const workbook = XLSX.read(buffer, { type: 'buffer', cellNF: true })
const rows = []
const headerParts = []
for (const sheetName of workbook.SheetNames) {
const sheet = workbook.Sheets[sheetName]
const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1, raw: false })
if (!rawRows.length) continue
let headerRowIdx = -1
for (let i = 0; i < Math.min(15, rawRows.length); i++) {
const cells = rawRows[i].map(c => String(c).trim()).filter(Boolean)
if (cells.length < 2) continue
if (cells.length === 1 && cells[0].length > 60) continue
const shortCells = cells.filter(c => c.length <= 60)
if (shortCells.length >= 2) { headerRowIdx = i; break }
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
headerParts.push(`=== Sheet: ${sheetName} ===`)
let rowsEmitted = 0
for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
const row = rawRows[i]
if (!row.some(cell => String(cell).trim() !== '')) continue
const cells = row.map(cell => String(cell || '').replace(/\r?\n/g, ' ').trim())
const nameVal = colIdx.name !== undefined ? (cells[colIdx.name] || '').trim() : ''
const tableVal = colIdx.table !== undefined ? (cells[colIdx.table] || '').trim() : sheetName
const descVal = colIdx.description !== undefined ? (cells[colIdx.description] || '').trim() : ''
const urlVal = colIdx.url !== undefined ? (cells[colIdx.url] || '').trim() : ''
const additionalVal = colIdx.additional !== undefined ? (cells[colIdx.additional] || '').trim() : ''
let formulaVal = colIdx.formula !== undefined ? (cells[colIdx.formula] || '').trim() : ''
if (!formulaVal && descVal) {
const formulaPatterns = [
/(.*?\/.*?)/i,
/(=.*?)/i,
/(calculated\s+as.*)/i,
/(computed\s+as.*)/i,
/(divided\s+by.*)/i,
/(multiplied\s+by.*)/i,
/(sum\s+of.*)/i,
]
for (const pattern of formulaPatterns) {
const match = descVal.match(pattern)
if (match && match[0].trim().length > 3) {
formulaVal = match[0].trim()
break
}
}
}
if (nameVal) {
let synthesis = `${nameVal}`
if (tableVal && tableVal !== sheetName) synthesis += ` (${tableVal})`
if (descVal) synthesis += ` is defined as: ${descVal}`
if (formulaVal && !descVal.toLowerCase().includes(formulaVal.toLowerCase())) {
synthesis += `. Formula: ${formulaVal}`
}
if (additionalVal) synthesis += `. Additional Info: ${additionalVal}`
if (urlVal) synthesis += `. URL: ${urlVal}`
rows.push({
text: synthesis,
metadata: {
measure: nameVal,
table: tableVal || sheetName,
formula: formulaVal || '',
description: descVal || '',
url: urlVal || '',
sourceSheet: sheetName,
}
})
if (formulaVal) {
rows.push({
text: `How to calculate ${nameVal}: ${formulaVal}`,
metadata: { measure: nameVal, table: tableVal || sheetName, formula: formulaVal, description: descVal || '', url: '', sourceSheet: sheetName }
})
rows.push({
text: `Formula for ${nameVal}: ${formulaVal}`,
metadata: { measure: nameVal, table: tableVal || sheetName, formula: formulaVal, description: descVal || '', url: '', sourceSheet: sheetName }
})
}
if (urlVal) {
rows.push({
text: `Report URL for ${nameVal}: ${urlVal}`,
metadata: { measure: nameVal, table: tableVal || sheetName, formula: '', description: '', url: urlVal, sourceSheet: sheetName }
})
rows.push({
text: `Power BI link for ${nameVal}: ${urlVal}`,
metadata: { measure: nameVal, table: tableVal || sheetName, formula: '', description: '', url: urlVal, sourceSheet: sheetName }
})
if (tableVal && tableVal !== sheetName) {
rows.push({
text: `Report URL for ${nameVal} (${tableVal}): ${urlVal}`,
metadata: { measure: nameVal, table: tableVal, formula: '', description: '', url: urlVal, sourceSheet: sheetName }
})
}
}
rowsEmitted++
} else if (descVal) {
rows.push({
text: descVal,
metadata: { measure: '', table: tableVal || sheetName, formula: '', description: descVal, url: '', sourceSheet: sheetName }
})
}
}
if (rowsEmitted === 0) {
for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
const row = rawRows[i]
const cells = row.map(c => String(c || '').trim()).filter(Boolean)
if (cells.length) {
rows.push({
text: cells.join(' | '),
metadata: { measure: '', table: sheetName, formula: '', description: '', url: '', sourceSheet: sheetName }
})
}
}
}
}
return rows
}
function keywordSearch(query, chunks, topK, intent, invertedIndex) {
const subject = extractSubject(query)
const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
const queryLower = normalizeQuery(query)
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
for (const w of ['url', 'link', 'https', 'http', 'powerbi', 'app']) {
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
if (intent === 'all_urls') {
if (!text.includes('http')) return { ...c, _score: 0 }
return { ...c, _score: 10 }
}
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
if (new RegExp(`\\|\\s*${escapeRegex(subject.toLowerCase())}\\s*\\|`, 'i').test(c.text || '')) score += subjectWords.length * 4
if (new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b[\\s\\S]{0,30}(is defined as|is calculated as|formula:)`, 'i').test(c.text || '')) score += subjectWords.length * 8
}
const wordCoverage = subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(c.text || '')).length
score += wordCoverage * 2
if (new RegExp(`\\b${escapeRegex(queryLower)}\\b`, 'i').test(c.text || '')) score += 3
if (subjectPhraseRegex.test(c.text || '')) score += 4
if (intent === 'calculation') {
if (/\bformula\b/i.test(text)) score += 15
if (/\bcalculated as\b/i.test(text)) score += 10
if (/\bcomputed as\b/i.test(text)) score += 10
if (text.includes('=')) score += 8
if (text.includes('/')) score += 5
if (/\bhow to calculate\b/i.test(text)) score += 12
if (/\bformula for\b/i.test(text)) score += 12
}
if (c.metadata && c.metadata.measure) {
const measureLower = (c.metadata.measure || '').toLowerCase().trim()
if (measureLower === subject.toLowerCase().trim()) {
score += 100
} else if (subjectWords.some(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(measureLower))) {
score += 10
}
}
const penalty = computeNegativePenalty(subject, c.text || '')
score -= penalty
}
return { ...c, _score: score }
})
.filter(c => c._score > 0)
.sort((a, b) => b._score - a._score)
.slice(0, topK)
}
function relaxedKeywordSearch(query, chunks, topK, invertedIndex) {
const subject = extractSubject(query)
const allWords = [
...subject.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 1),
...query.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(w => w.length > 2),
]
const uniqueWords = [...new Set(allWords)]
const union = new Set()
if (invertedIndex) {
for (const w of uniqueWords) {
for (const idx of (invertedIndex.get(w) || new Set())) union.add(idx)
for (const variant of normalizeTerms(w)) {
for (const idx of (invertedIndex.get(variant) || new Set())) union.add(idx)
}
}
}
const source = union.size > 0 ? [...union].map(i => chunks[i]).filter(Boolean) : chunks.slice(0, 300)
return source
.map(c => {
const text = (c.text || '').toLowerCase()
const matched = uniqueWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(text)).length
const subjectMatch = subject.length > 2 && new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b`, 'i').test(text) ? 5 : 0
let metaBoost = 0
if (c.metadata && c.metadata.measure) {
const ml = c.metadata.measure.toLowerCase()
if (ml === subject.toLowerCase().trim()) {
metaBoost += 50
} else {
const subjectMatched = uniqueWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`, 'i').test(ml)).length
metaBoost = subjectMatched * 3
}
}
const penalty = computeNegativePenalty(subject, c.text || '')
return { ...c, _score: Math.max(0, matched + subjectMatch + metaBoost - penalty) }
})
.filter(c => c._score > 0)
.sort((a, b) => b._score - a._score)
.slice(0, topK)
}
async function retrieveChunks(query, chunks, topK, invertedIndex, _isRetry = false) {
const intent = detectQueryIntent(query)
const normalizedQuery = normalizeQuery(query).replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ')
const MAX_HITS = MAX_HITS_GLOBAL
if (intent === 'all_urls') {
return chunks.filter(c => /https?:\/\/\S+/.test(c.text || '')).slice(0, 100)
}
const candidates = keywordSearch(normalizedQuery, chunks, Math.min(150, chunks.length), intent, invertedIndex)
const pool = candidates.length > 0 ? candidates : chunks.slice(0, 150)
const topScore = pool[0]?._score || 0
let topCandidates = []
if (topScore >= 6) topCandidates = pool.slice(0, Math.min(MAX_HITS, pool.length))
else if ((intent === 'definition' || intent === 'calculation') && topScore >= 3) topCandidates = pool.slice(0, Math.min(MAX_HITS, pool.length))
else if (intent === 'url_lookup' && pool.length > 0) return pool.slice(0, Math.min(topK, 6))
if (topCandidates.length === 0 && AZURE_EMBED_ENDPOINT && AZURE_EMBED_KEY) {
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
const result = [...scored, ...remainder].sort((a, b) => b._score - a._score).slice(0, Math.min(MAX_HITS, MAX_HITS_GLOBAL))
if (result.length > 0) topCandidates = result
}
} catch (err) {
console.warn('[retrieveChunks] Embedding failed, using keyword fallback:', err.message)
}
}
if (topCandidates.length === 0 && pool.length > 0) {
topCandidates = pool.slice(0, Math.min(MAX_HITS, pool.length))
}
if (topCandidates.length === 0 && !_isRetry) {
const corrected = fuzzyCorrectQuery(query, chunks)
if (corrected.toLowerCase() !== query.toLowerCase()) {
console.log(`[QueryPipeline] Self-healing retry with fuzzy-corrected query: "${corrected}"`)
return retrieveChunks(corrected, chunks, topK, invertedIndex, true)
}
}
if (topCandidates.length === 0) {
topCandidates = relaxedKeywordSearch(normalizedQuery, chunks, Math.min(topK * 2, 64), invertedIndex).slice(0, Math.min(topK, MAX_HITS))
}
if (RERANKER_ENDPOINT && topCandidates.length > 1) {
const reranked = await rerankChunks(normalizedQuery, topCandidates)
return reranked.slice(0, Math.min(topK, 5))
}
return topCandidates.slice(0, Math.min(topK, MAX_HITS))
}
function buildContext(hits) {
const seen = new Set()
const deduped = []
for (const h of hits) {
const fp = (h.text || '').trim().slice(0, 80).toLowerCase()
if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
if (deduped.length >= 8) break
}
return deduped.map((h, i) => {
const limit = i === 0 ? 1200 : 900
return `[Source ${i + 1}]\n${(h.text || '').trim().slice(0, limit)}`
}).join('\n\n---\n\n')
}
function buildSystemPrompt(intent) {
const intentRule = intent === 'definition'
? `Definition: Bold measure name, one sentence definition only. No formula or calculation details.`
: intent === 'calculation'
? `Calculation: Output ONLY "**Formula for [Name]:** [formula]." No definition or description.`
: intent === 'comparison'
? `Comparison: Bold each name. Write a concise definition for each. End with a "**Key Difference:**" sentence derived strictly from the context. Do not invent differences.`
: `General: Answer directly in 2-4 sentences. Do not volunteer formulas or definitions unprompted.`
return `You are a data dictionary assistant for a real estate analytics platform. Answer ONLY from context.
Rules: Bold the subject with **Name**. Write complete sentences only. No pipe-delimited data. No double periods. No source references like [1]. No sheet references. Keep answers concise.
If context lacks the answer, say: "I could not find this in your documents."
Intent rule (highest priority): ${intentRule}
Formats: Definition: "**[Name]** is [desc]." | Calculation: "**Formula for [Name]:** [formula]." | URL: return URL only. | Comparison: "**[A]:** [desc]. **[B]:** [desc]. **Key Difference:** [one sentence derived from context]."`
}
function buildUserMessage(query, hits, intent) {
const context = buildContext(hits)
const subject = extractSubject(query)
let instruction = ''
if (intent === 'definition') {
instruction = `\n\nFrom the context, write a one-sentence definition of "${subject}". Bold the name. No formula or calculation. End with one period.`
} else if (intent === 'calculation') {
instruction = `\n\nFrom the context, return only: "**Formula for ${capFirst(subject)}:** [formula]." No definition or description.`
} else if (intent === 'url_lookup') {
instruction = `\n\nFrom the context, return only the full URL for "${extractUrlKeywords(query).join(' ')}".`
} else if (intent === 'all_urls') {
instruction = `\n\nFrom the context, list ALL URLs. Format: name: URL. One per line.`
} else if (intent === 'comparison') {
instruction = `\n\nFrom the context, compare: ${query}. Bold each name. Provide a concise definition for each based strictly on the context. End with a "**Key Difference:**" sentence that accurately reflects what the context says about these two items. Do not fabricate or assume differences. One period at end.`
} else {
instruction = `\n\nFrom the context, answer in clear sentences: ${query}. Answer only what was asked. No pipe characters. One period at end.`
}
return `CONTEXT:\n${context}${instruction}`
}
function extractAllUrlsFromChunks(chunks) {
const results = []
const seen = new Set()
const urlRegex = /https?:\/\/[^\s"'<>]+/g
for (const chunk of chunks) {
const lines = (chunk.text || '').split('\n')
for (const line of lines) {
const urls = line.match(urlRegex)
if (!urls) continue
for (const url of urls) {
const cleanUrl = url.replace(/[.,;)]+$/, '').trim()
if (!cleanUrl.startsWith('http') || seen.has(cleanUrl)) continue
seen.add(cleanUrl)
let name = 'Report'
const reportUrlMatch = line.match(/^(?:Report URL|Power BI link)\s+for\s+(.+?)(?:\s*\([^)]+\))?\s*:\s*https?:/i)
if (reportUrlMatch) {
name = reportUrlMatch[1].trim()
} else {
const beforeUrl = line.slice(0, line.indexOf('http')).trim()
if (beforeUrl) {
const cleaned = beforeUrl
.replace(/\.\s*URL\s*:?\s*$/i, '')
.replace(/\s*:\s*$/, '')
.replace(/^(URL|Link|Dashboard|Report)\s*:?\s*/i, '')
.trim()
if (cleaned.length > 1 && cleaned.length < 120) name = cleaned
}
}
results.push({ name, url: cleanUrl })
}
}
}
return results
}
function buildFallbackAnswer(query, hits, intent) {
if (!hits || hits.length === 0) return "I could not find relevant information about this in your documents."
const resolvedIntent = intent || detectQueryIntent(query)
const subject = extractSubject(query)
const subjectLower = subject.toLowerCase()
const escapedSubject = escapeRegex(subjectLower)
if (resolvedIntent === 'all_urls') {
const urlEntries = extractAllUrlsFromChunks(hits)
if (urlEntries.length === 0) return "I could not find any URLs in your documents."
return urlEntries.map(e => `**${e.name}:** ${e.url}`).join('\n')
}
if (resolvedIntent === 'url_lookup') {
const urlKeywords = extractUrlKeywords(query)
const urlRegex = /https?:\/\/[^\s"'<>]+/
for (const h of hits) {
for (const line of (h.text || '').split('\n')) {
if (!urlRegex.test(line)) continue
const matched = urlKeywords.filter(w => line.toLowerCase().includes(w)).length
if (matched > 0) {
const urlMatch = line.match(urlRegex)
if (urlMatch) return urlMatch[0].replace(/[.,;)]+$/, '').trim()
}
}
}
for (const h of hits) {
for (const line of (h.text || '').split('\n')) {
const urlMatch = line.match(urlRegex)
if (urlMatch) return urlMatch[0].replace(/[.,;)]+$/, '').trim()
}
}
return "I could not find a matching URL in your documents."
}
if (resolvedIntent === 'calculation') {
for (const h of hits) {
if (h.metadata && h.metadata.formula && new RegExp(`\\b${escapedSubject}\\b`, 'i').test(h.metadata.measure || '')) {
const cap = capFirst(h.metadata.measure)
return ensureSinglePeriod(`**Formula for ${cap}:** ${h.metadata.formula}.`)
}
}
const calcPattern = new RegExp(`how to calculate ${escapedSubject}:\\s*([^\\n]+)`, 'im')
for (const h of hits) {
const m = (h.text || '').match(calcPattern)
if (m) {
const cap = capFirst(subject)
return ensureSinglePeriod(`**Formula for ${cap}:** ${trimToCompleteSentence(m[1].trim(), 500)}.`)
}
}
const formulaPattern = new RegExp(`formula for ${escapedSubject}:\\s*([^\\n]+)`, 'im')
for (const h of hits) {
const m = (h.text || '').match(formulaPattern)
if (m) {
const cap = capFirst(subject)
return ensureSinglePeriod(`**Formula for ${cap}:** ${trimToCompleteSentence(m[1].trim(), 500)}.`)
}
}
for (const h of hits) {
const text = h.text || ''
if (!new RegExp(`\\b${escapedSubject}\\b`, 'i').test(text)) continue
const extracted = extractFormulaFromText(text)
if (extracted) {
const cap = capFirst(subject)
return ensureSinglePeriod(`**Formula for ${cap}:** ${extracted}.`)
}
}
return `I could not find a formula for ${capFirst(subject)} in your documents.`
}
for (const h of hits) {
if (h.metadata && h.metadata.measure) {
const measureLower = (h.metadata.measure || '').toLowerCase().trim()
if (measureLower === subjectLower || new RegExp(`\\b${escapedSubject}\\b`, 'i').test(measureLower) || new RegExp(`\\b${escapeRegex(measureLower)}\\b`, 'i').test(subjectLower)) {
const cap = capFirst(h.metadata.measure)
if (resolvedIntent === 'definition' && h.metadata.description) {
return ensureSinglePeriod(`**${cap}** is defined as: ${h.metadata.description}.`)
}
let answer = `**${cap}**`
if (h.metadata.description) answer += ` is defined as: ${h.metadata.description}`
if (!answer.endsWith('.')) answer += '.'
return ensureSinglePeriod(answer)
}
}
}
const synthesisPattern = new RegExp(
`${escapedSubject}[^\\n]*is defined as:\\s*([^.\\n]+(?:\\.[^.\\n]+)?)(?:\\.\\s*Formula:\\s*([^.\\n]+(?:\\.[^.\\n]+)?))?(?:\\.\\s*Additional Info:\\s*([^.\\n]+))?`,
'im'
)
for (const h of hits) {
const m = (h.text || '').match(synthesisPattern)
if (m) {
const desc = trimToCompleteSentence((m[1] || '').trim(), 600)
const additional = (m[3] || '').trim().slice(0, 300)
const cap = capFirst(subject)
let answer = `**${cap}** is ${desc}`
if (!answer.endsWith('.')) answer += '.'
if (resolvedIntent !== 'definition' && m[2]) {
const formula = (m[2] || '').trim().slice(0, 400)
if (formula) { answer += `\n\n**Formula:** ${formula}`; if (!answer.endsWith('.')) answer += '.' }
}
if (additional) { answer += `\n\n**Additional Info:** ${additional}`; if (!answer.endsWith('.')) answer += '.' }
return ensureSinglePeriod(answer)
}
}
const matchingLines = []
for (const h of hits) {
for (const line of (h.text || '').split('\n')) {
if (!new RegExp(`\\b${escapedSubject}\\b`, 'i').test(line)) continue
if (line.trim().length <= 20) continue
if ((line.match(/\|/g) || []).length > 2) continue
if (/^===\s*Sheet:/.test(line.trim())) continue
if (resolvedIntent === 'definition' && /formula|calculated as|computed as/i.test(line)) continue
const cleaned = line.trim().replace(/\(from\s+[A-Za-z\s]+\)/g, '').trim()
if (cleaned.length > 15) matchingLines.push(cleaned)
}
}
if (matchingLines.length > 0) {
const cap = capFirst(subject)
const joined = trimToCompleteSentence([...new Set(matchingLines)].slice(0, 3).join(' '), 600)
return ensureSinglePeriod(`**${cap}:** ${joined}.`)
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
cleaned = trimToLastCompleteSentence(cleaned)
if (cleaned.length > 0 && !/[.!?]$/.test(cleaned)) cleaned += '.'
return ensureSinglePeriod(cleaned)
}
function trimToLastCompleteSentence(text) {
if (!text) return ''
const lastIdx = Math.max(text.lastIndexOf('. '), text.lastIndexOf('.\n'), text.lastIndexOf('! '), text.lastIndexOf('? '))
if (lastIdx > text.length * 0.5) {
const trimmed = text.slice(0, lastIdx + 1).trim()
if (trimmed.length > 20) return trimmed
}
if (/[.!?]$/.test(text)) return text
const periodIdx = text.lastIndexOf('.')
if (periodIdx > text.length * 0.5) return text.slice(0, periodIdx + 1).trim()
return text
}
async function generateAnswer(query, hits, intent) {
return callBestAvailableEngine(buildSystemPrompt(intent), buildUserMessage(query, hits, intent), 1024)
}
async function generateAnswerForTopic(topic, chunks, topK, invertedIndex) {
const topicQuery = `what is ${topic}`
let hits = await retrieveChunks(topicQuery, chunks, topK, invertedIndex)
if (hits.length === 0) hits = relaxedKeywordSearch(topicQuery, chunks, 32, invertedIndex)
if (hits.length === 0) return null
const intent = 'definition'
let rawAnswer = ''
try {
rawAnswer = await Promise.race([
generateAnswer(topicQuery, hits, intent),
new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 25000)),
])
} catch (err) {
console.warn(`[generateAnswerForTopic] Engine failed for "${topic}": ${err.message}`)
}
const isBlank = !rawAnswer || rawAnswer.trim().length < 15
let answer = isBlank ? buildFallbackAnswer(topicQuery, hits, intent) : cleanAnswer(rawAnswer)
answer = answer.replace(/^\*\*[^*]+\*\*\s*(is defined as:?\s*)?/i, '').trim()
if (answer && !/[.!?]$/.test(answer)) answer += '.'
return answer
}
async function generateComparisonAnswer(topicA, topicB, chunks, topK, invertedIndex) {
const comparisonQuery = `difference between ${topicA} and ${topicB}`
const hitsA = await retrieveChunks(`what is ${topicA}`, chunks, topK, invertedIndex)
const hitsB = await retrieveChunks(`what is ${topicB}`, chunks, topK, invertedIndex)
const allHits = [...hitsA, ...hitsB]
const seen = new Set()
const deduped = []
for (const h of allHits) {
const fp = (h.text || '').trim().slice(0, 80).toLowerCase()
if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
}
if (deduped.length === 0) return null
let rawAnswer = ''
try {
rawAnswer = await Promise.race([
generateAnswer(comparisonQuery, deduped, 'comparison'),
new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 25000)),
])
} catch (err) {
console.warn(`[generateComparisonAnswer] Engine failed for "${topicA}" vs "${topicB}": ${err.message}`)
}
if (rawAnswer && rawAnswer.trim().length >= 15) return cleanAnswer(rawAnswer)
const answerA = await generateAnswerForTopic(topicA, chunks, topK, invertedIndex)
const answerB = await generateAnswerForTopic(topicB, chunks, topK, invertedIndex)
const parts = []
if (answerA && !answerA.includes('could not find')) parts.push(`**${capFirst(topicA)}:** ${answerA}`)
else parts.push(`**${capFirst(topicA)}:** I could not find information about "${capFirst(topicA)}" in your documents.`)
if (answerB && !answerB.includes('could not find')) parts.push(`**${capFirst(topicB)}:** ${answerB}`)
else parts.push(`**${capFirst(topicB)}:** I could not find information about "${capFirst(topicB)}" in your documents.`)
return parts.join('\n\n')
}

async function handleMultiTopicQuery(topics, mode, chunks, topK, invertedIndex) {
  const results = await Promise.all(
    topics.map(async (topic) => {
      const answer = await generateAnswerForTopic(topic, chunks, topK, invertedIndex)
      return { topic, answer }
    })
  )
  const parts = results.map(({ topic, answer }) => {
    const cap = capFirst(topic)
    if (!answer || answer.includes('could not find') || answer.includes('not present')) {
      return `**${cap}:**\nI could not find information about "${cap}" in your documents.`
    }
    return `**${cap}:**\n${answer}`
  })
  if (mode === 'comparison' && results.length === 2) {
    const [a, b] = results
    const bothFound =
      a.answer && !a.answer.includes('could not find') &&
      b.answer && !b.answer.includes('could not find')

    if (bothFound) {
      const comparisonAnswer = await generateComparisonAnswer(
        a.topic, b.topic, chunks, topK, invertedIndex
      )
      if (comparisonAnswer && !comparisonAnswer.includes('could not find')) {
        return comparisonAnswer
      }
    }
    return parts.join('\n\n')
  }
  return parts.join('\n\n')
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
parseOffice(buffer, (text, err) => { if (err) reject(err); else resolve(text || '') }, { outputErrorToConsole: false })
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
.map(line => { try { return JSON.stringify(JSON.parse(line)) } catch { return line } }).join('\n')
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
parseOffice(buffer, (text, err) => { resolve(err || !text ? '[EPUB: convert to PDF for best results]' : text) }, { outputErrorToConsole: false })
})
}
async function extractTextFromBuffer(buffer, fileName) {
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
if (ext === '.pdf') return extractPdf(buffer)
if (ext === '.docx' || ext === '.doc') return extractWord(buffer)
if (ext === '.odt' || ext === '.rtf') return extractOffice(buffer)
if (['.xlsx', '.xls', '.ods'].includes(ext)) return null
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
function chunkText(text, sourceFile) {
const chunks = []
let chunkIndex = 0
const blocks = text.replace(/\r\n/g, '\n').split(/\n{2,}/).map(b => b.trim()).filter(b => b.length > 0)
let buffer = []
let bufferLength = 0
function flush() {
const chunkStr = buffer.join('\n\n')
if (chunkStr.length >= 30) chunks.push({ text: chunkStr, source_file: sourceFile, chunk_index: chunkIndex++, embedding: [] })
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
let chunkIndexOffset = 0
for (let i = 0; i < blobNames.length; i += BLOB_CONCURRENCY) {
const batch = blobNames.slice(i, i + BLOB_CONCURRENCY)
const results = await Promise.allSettled(
batch.map(async (blobName) => {
const fileName = blobName.split('/').pop()
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
const buffer = await downloadBlobAsBuffer(containerClient, blobName)
if (['.xlsx', '.xls', '.ods'].includes(ext)) {
const structuredRows = extractSpreadsheet(buffer)
return structuredRows.map((row, idx) => ({
text: row.text,
source_file: fileName,
chunk_index: idx,
embedding: [],
metadata: row.metadata || null,
}))
}
const text = await extractTextFromBuffer(buffer, fileName)
if (!text?.trim()) return []
return chunkText(text, fileName)
})
)
for (const result of results) {
if (result.status === 'fulfilled') {
const fileChunks = result.value
fileChunks.forEach((c, idx) => { c.chunk_index = chunkIndexOffset + idx })
chunkIndexOffset += fileChunks.length
allChunks.push(...fileChunks)
} else {
console.warn('[loadChunks] Blob failed:', result.reason?.message)
}
}
}
return allChunks
}
const CHUNK_CACHE = new Map()
async function loadChunksForClient(clientId) {
const now = Date.now()
const cached = CHUNK_CACHE.get(clientId)
if (cached && cached.chunks) {
if (now - cached.ts <= CHUNK_CACHE_TTL) return { chunks: cached.chunks, invertedIndex: cached.invertedIndex }
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
function generateApiKey() { return `rak_${crypto.randomBytes(32).toString('hex')}` }
function generateTitle(query) {
const cleaned = query.trim().replace(/[?!.]+$/, '')
return cleaned.length > 50 ? cleaned.slice(0, 50) + '…' : cleaned
}
async function saveConversationMessage(clientId, conversationId, query, answer, sources) {
try {
const chatDatabase = await getChatDb()
const col = chatDatabase.collection('conversations')
const now = new Date()
const userMsg = { role: 'user', content: query, timestamp: now }
const assistantMsg = { role: 'assistant', content: answer, sources: sources.map(s => ({ source_file: s.source_file, score: s.score })), timestamp: now }
let activeConversationId = conversationId || null
if (activeConversationId) {
const updated = await col.findOneAndUpdate(
{ _id: new ObjectId(activeConversationId), clientId },
{ $push: { messages: { $each: [userMsg, assistantMsg] } }, $set: { updatedAt: now } },
{ returnDocument: 'after', projection: { _id: 1 } }
)
if (!updated) activeConversationId = null
}
if (!activeConversationId) {
const result = await col.insertOne({ clientId, title: generateTitle(query), messages: [userMsg, assistantMsg], createdAt: now, updatedAt: now })
activeConversationId = result.insertedId.toString()
}
return activeConversationId
} catch (saveErr) {
console.warn('[saveConversationMessage] Failed:', saveErr.message)
return conversationId || null
}
}
app.get('/health', (req, res) => res.json({
ok: true,
service: 'ask-data',
engines: { primary: ASKDATA_ENDPOINT ? 'configured' : 'missing', fallback: ASKDATA2_ENDPOINT ? 'configured' : 'missing' },
embeddings: AZURE_EMBED_ENDPOINT ? `azure (${AZURE_EMBED_MODEL})` : 'keyword-only',
reranker: RERANKER_ENDPOINT ? 'configured' : 'disabled',
chunkCacheSize: CHUNK_CACHE.size,
responseCacheSize: RESPONSE_CACHE.size,
primaryCircuitOpen: askedataCircuitOpen(),
primaryFailures: askedataFailures,
maxHits: MAX_HITS_GLOBAL,
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
if (blobServiceClient) loadChunksForClient(client.clientId).catch(err => console.warn(`[login warmup] ${client.clientId}: ${err.message}`))
res.json({ ok: true, client })
} catch (err) { res.status(500).json({ error: err.message }) }
})
app.post('/chat/login', async (req, res) => {
try {
const apiKey = extractApiKey(req) || req.body?.apiKey
if (!apiKey) return res.status(400).json({ error: 'apiKey is required' })
const client = await verifyApiKey(apiKey)
if (!client) return res.status(401).json({ error: 'Invalid API key' })
if (blobServiceClient) loadChunksForClient(client.clientId).catch(err => console.warn(`[chat/login warmup] ${client.clientId}: ${err.message}`))
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
if (!validation.valid) return res.json({ answer: validation.message, sources: [], conversationId: conversationId || null, client: req.client })
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
const activeConversationId = await saveConversationMessage(clientId, conversationId || null, query.trim(), cached.answer, cached.sources || [])
return res.json({ ...cached, cached: true, conversationId: activeConversationId })
}
if (IN_FLIGHT.has(cacheKey)) {
try {
const result = await IN_FLIGHT.get(cacheKey)
const activeConversationId = await saveConversationMessage(clientId, conversationId || null, query.trim(), result.answer, result.sources || [])
return res.json({ ...result, conversationId: activeConversationId })
} catch { }
}
const requestPromise = (async () => {
const { chunks, invertedIndex } = await loadChunksForClient(clientId)
if (chunks.length === 0) {
return { answer: 'No documents found for your account. Please ensure your documents have been ingested first.', sources: [], client: { clientId, name } }
}
let processedQuery = applyTypos(query.trim())
console.log(`[QueryPipeline] Original: "${query.trim()}"`)
if (processedQuery !== query.trim()) {
console.log(`[QueryPipeline] After typos: "${processedQuery}"`)
}
processedQuery = applySynonyms(processedQuery)
const fuzzyResult = fuzzyCorrectQuery(processedQuery, chunks)
if (fuzzyResult !== processedQuery) {
console.log(`[QueryPipeline] After fuzzy: "${fuzzyResult}"`)
}
processedQuery = fuzzyResult
const rewritten = await preprocessQuery(processedQuery)
if (rewritten !== processedQuery) {
console.log(`[QueryPipeline] After rewrite: "${rewritten}"`)
}
processedQuery = rewritten
if (intent === 'all_urls') {
const urlChunks = chunks.filter(c => /https?:\/\/\S+/.test(c.text || ''))
const urlEntries = extractAllUrlsFromChunks(urlChunks)
const answer = urlEntries.length > 0 ? urlEntries.map(e => `**${e.name}:** ${e.url}`).join('\n') : "I could not find any URLs in your documents."
const sources = urlChunks.slice(0, 6).map(h => ({ source_file: h.source_file || 'unknown', chunk_index: h.chunk_index ?? 0, score: null, preview: (h.text || '').slice(0, 200) }))
return { answer, sources, client: { clientId, name } }
}
const multiTopicCheck = detectMultiTopicQuery(processedQuery)
if (multiTopicCheck.isMulti) {
console.log(`[chat/message] Multi-topic detected: ${JSON.stringify(multiTopicCheck.topics)} mode=${multiTopicCheck.mode}`)
const answer = await handleMultiTopicQuery(multiTopicCheck.topics, multiTopicCheck.mode, chunks, Math.min(topK, MAX_HITS_GLOBAL), invertedIndex)
return { answer, sources: [], client: { clientId, name } }
}
let hits = await retrieveChunks(processedQuery, chunks, Math.min(topK, MAX_HITS_GLOBAL), invertedIndex)
if (hits.length === 0) hits = relaxedKeywordSearch(processedQuery, chunks, 64, invertedIndex)
console.log(`[chat/message] "${query.slice(0, 60)}" → intent=${intent}, subject="${extractSubject(processedQuery)}", hits=${hits.length}, topScore=${hits[0]?._score?.toFixed(2) || 0}`)
if (hits.length === 0) {
return { answer: "I could not find relevant information about this in your documents. Try rephrasing your question.", sources: [], client: { clientId, name } }
}
let rawAnswer = ''
if (intent !== 'url_lookup') {
try {
rawAnswer = await Promise.race([
generateAnswer(processedQuery, hits, intent),
new Promise((_, reject) => setTimeout(() => reject(new Error('All engines timed out')), 55000)),
])
} catch (err) {
console.warn(`[chat/message] All engines failed, using rule-based fallback: ${err.message}`)
}
}
const isBlank = !rawAnswer || rawAnswer.trim().length < 15
const answer = isBlank ? buildFallbackAnswer(processedQuery, hits, intent) : cleanAnswer(rawAnswer)
if (isBlank) console.warn(`[chat/message] Blank from all engines, used rule-based fallback for: "${query.slice(0, 60)}"`)
const sources = hits.map(h => ({
source_file: h.source_file || 'unknown',
chunk_index: h.chunk_index ?? 0,
score: typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
preview: (h.text || '').slice(0, 200),
}))
return { answer, sources, client: { clientId, name } }
})()
IN_FLIGHT.set(cacheKey, requestPromise)
let result
try { result = await requestPromise } finally { IN_FLIGHT.delete(cacheKey) }
if (result.answer && result.answer.length > 15) responseCacheSet(cacheKey, result)
const activeConversationId = await saveConversationMessage(clientId, conversationId || null, query.trim(), result.answer, result.sources || [])
res.json({ ...result, conversationId: activeConversationId })
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
console.log(`Service running on port ${PORT}`)
console.log(`ASKDATA: ${ASKDATA_ENDPOINT ? 'configured' : 'MISSING'} | ASKDATA2: ${ASKDATA2_ENDPOINT ? 'configured' : 'missing'}`)
console.log(`Embedding model: ${AZURE_EMBED_MODEL} | Reranker: ${RERANKER_ENDPOINT ? 'configured' : 'disabled'} | MAX_HITS: ${MAX_HITS_GLOBAL}`)
startApiKeyHealthChecker()
warmupChunkCaches()
})
module.exports = app
