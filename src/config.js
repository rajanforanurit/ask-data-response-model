require('dotenv').config()
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
const FAISS_PREFIX = 'faiss'
const BM25_PREFIX = 'bm25'
const CHUNK_SIZE = 600
const CHUNK_OVERLAP = 120
const BLOB_CONCURRENCY = parseInt(process.env.BLOB_CONCURRENCY || '8', 10)
const CHUNK_CACHE_TTL = parseInt(process.env.CHUNK_CACHE_TTL_MS || '300000', 10)
const MAX_HITS_GLOBAL = 50
const FAISS_TOP_K = 40
const BM25_TOP_K = 40
const RERANK_TOP_K = 10
const SUPPORTED_EXTENSIONS = new Set(['.pdf','.docx','.doc','.xlsx','.xls','.csv'])
const SYNONYM_MAP = [
{pattern:/\bapp(lication)?\s+count\b/i,canonical:'application count'},
{pattern:/\btotal\s+app(lication)?\s+count\b/i,canonical:'application count'},
{pattern:/\bnumber\s+of\s+app(lication)?s\b/i,canonical:'application count'},
{pattern:/\bapp(lication)?\s+volume\b/i,canonical:'application count'},
{pattern:/\btotal\s+submitted\s+app(lication)?s\b/i,canonical:'application count'},
{pattern:/\btotal\s+app(lication)?s\b/i,canonical:'application count'},
{pattern:/\bsubmitted\s+app(lication)?s\b/i,canonical:'application count'},
{pattern:/\bocc(upancy)?\s+rate\b/i,canonical:'occupancy rate'},
{pattern:/\bocc(upancy)?\s+formula\b/i,canonical:'occupancy formula'},
{pattern:/\blead\s+acq(uisition)?\s+cost\b/i,canonical:'lead acquisition cost'},
{pattern:/\blead\s+cost\b/i,canonical:'lead acquisition cost'},
{pattern:/\b(salary\s+package|pay\s+package|ctc\s+package|compensation\s+package)\b/i,canonical:'compensation breakup'},
{pattern:/\b(total\s+ctc|total\s+cost\s+to\s+company|cost\s+to\s+company)\b/i,canonical:'total ctc'},
{pattern:/\b(gross\s+salary|gross\s+pay|gross\s+wage)\b/i,canonical:'gross salary'},
{pattern:/\b(basic\s+salary|base\s+salary|basic\s+wage|base\s+pay)\b/i,canonical:'basic pay'},
{pattern:/\b(hra|house\s+rent\s+allowance)\b/i,canonical:'hra'},
{pattern:/\b(special\s+allowance|sp\s+allowance)\b/i,canonical:'special allowance'},
{pattern:/\b(performance\s+bonus|annual\s+bonus|yearly\s+bonus|incentive\s+bonus)\b/i,canonical:'performance bonus'},
{pattern:/\b(package\s+of|salary\s+of|pay\s+of|ctc\s+of|compensation\s+of|remuneration\s+of)\b/i,canonical:'compensation breakup of'},
{pattern:/\bwhat\s+is\s+(his|her|their)\s+package\b/i,canonical:'what is compensation breakup'},
{pattern:/\b(joining\s+date|date\s+of\s+joining|start\s+date)\b/i,canonical:'joining date'},
{pattern:/\b(notice\s+period|probation\s+period|probation)\b/i,canonical:'probation'},
{pattern:/\b(data\s+retention\s+policy|retention\s+policy|data\s+retention\s+period)\b/i,canonical:'data retention'},
{pattern:/\b(privacy\s+policy|privacy\s+terms)\b/i,canonical:'privacy policy'},
{pattern:/\b(contact\s+info|contact\s+details|contact\s+us)\b/i,canonical:'contact information'},
{pattern:/\b(children\s+privacy|kids\s+privacy|minors\s+privacy)\b/i,canonical:"children's privacy"},
{pattern:/\b(data\s+security|data\s+protection|security\s+policy)\b/i,canonical:'data security'},
{pattern:/\b(data\s+sharing|share\s+data|third\s+party\s+sharing)\b/i,canonical:'data sharing'},
]
const TYPO_MAP = {
ehat:'what',waht:'what',whta:'what',whar:'what',
hwo:'how',hoe:'how',
difine:'define',definr:'define',defien:'define',defne:'define',deifne:'define',
expain:'explain',expalin:'explain',explian:'explain',
wht:'what',shwo:'show',lsit:'list',lits:'list',
occupany:'occupancy',occpancy:'occupancy',occupncy:'occupancy',
applcant:'applicant',applicnat:'applicant',
retension:'retention',retantion:'retention',
varience:'variance',varianec:'variance',
}
const DOMAIN_SHORT_SAFELIST = new Set([
'count','rate','rent','cost','date','type','name','unit','term','area',
'base','gross','net','avg','sum','min','max','ytd','mtd','per','fee',
'tax','due','paid','void','open','loss','gain','flow','days','beds',
'bath','sqft','tier','band','code','flag','rank','sort','key','ref',
])
const LOCATION_QUALIFIER_PATTERN = /\s+(?:in|at|for|of|within|under|inside|from)\s+(?:the\s+)?(?:[a-z][a-z\s]{1,30})$/i
const PERSON_QUALIFIER_PATTERN = /\s+(?:of|for)\s+(?:[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)$/
const RESPONSE_CACHE = new Map()
const RESPONSE_CACHE_TTL = 10 * 60 * 1000
const RESPONSE_CACHE_MAX = 1000
function responseCacheGet(key) {
const entry = RESPONSE_CACHE.get(key)
if (!entry) return null
if (Date.now() - entry.ts > RESPONSE_CACHE_TTL) {RESPONSE_CACHE.delete(key);return null}
return entry.value
}
function responseCacheSet(key, value) {
if (RESPONSE_CACHE.size >= RESPONSE_CACHE_MAX) RESPONSE_CACHE.delete(RESPONSE_CACHE.keys().next().value)
RESPONSE_CACHE.set(key, {value, ts: Date.now()})
}
function escapeRegex(str) {return str.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}
function capFirst(str) {if (!str) return '';return str.charAt(0).toUpperCase() + str.slice(1)}
function ensureSinglePeriod(text) {if (!text) return '';return text.replace(/\.{2,}/g,'.').replace(/\.\s*\./g,'.').trim()}
function trimToCompleteSentence(text, maxLen = 1200) {
if (!text || text.length <= maxLen) return text
const truncated = text.slice(0, maxLen)
const lastPeriod = Math.max(truncated.lastIndexOf('. '),truncated.lastIndexOf('.\n'),truncated.lastIndexOf('.'))
if (lastPeriod > maxLen * 0.5) return truncated.slice(0, lastPeriod + 1).trim()
return truncated.trim()
}
function trimToLastCompleteSentence(text) {
if (!text) return ''
const lastIdx = Math.max(text.lastIndexOf('. '),text.lastIndexOf('.\n'),text.lastIndexOf('! '),text.lastIndexOf('? '))
if (lastIdx > text.length * 0.5) {
const trimmed = text.slice(0, lastIdx + 1).trim()
if (trimmed.length > 20) return trimmed
}
if (/[.!?]$/.test(text)) return text
const periodIdx = text.lastIndexOf('.')
if (periodIdx > text.length * 0.5) return text.slice(0, periodIdx + 1).trim()
return text
}
function fixBrokenUrls(text) {return text.replace(/https:\/\/[^\s]+(\s+[^\s]+)/g,match => match.replace(/\s/g,''))}
function applySynonyms(query) {
let q = query
for (const {pattern, canonical} of SYNONYM_MAP) q = q.replace(pattern, canonical)
return q
}
function applyTypos(query) {
return query.split(/\s+/).map(w => {
const lower = w.toLowerCase()
return TYPO_MAP[lower] !== undefined ? TYPO_MAP[lower] : w
}).join(' ')
}
function levenshteinDistance(a, b) {
const m = a.length, n = b.length
const dp = Array.from({length: m + 1},(_,i) => Array.from({length: n + 1},(_,j) => (i === 0 ? j : j === 0 ? i : 0)))
for (let i = 1; i <= m; i++) {
for (let j = 1; j <= n; j++) {
if (a[i-1] === b[j-1]) dp[i][j] = dp[i-1][j-1]
else dp[i][j] = 1 + Math.min(dp[i-1][j],dp[i][j-1],dp[i-1][j-1])
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
function normalizeQuery(query) {
return applySynonyms(query)
.toLowerCase().trim()
.replace(/\bweek\s+(\d)\b/g,(_,n) => `week 0${n}`)
.replace(/[?!.]+$/,'').replace(/\s+/g,' ')
}
function normalizeQueryForCache(query) {
return applySynonyms(query)
.toLowerCase().trim()
.replace(/\bweek\s+(\d)\b/g,(_,n) => `week 0${n}`)
.replace(/^(what\s+is\s+(the\s+)?(definition|meaning)\s+(of|for|to)\s+)/i,'')
.replace(/^(define\s+(the\s+)?)/i,'')
.replace(/^(explain\s+(the\s+)?)/i,'')
.replace(/^(tell\s+me\s+about\s+(the\s+)?)/i,'')
.replace(/^(what\s+are\s+(the\s+)?)/i,'')
.replace(/^(what\s+is\s+)/i,'')
.replace(/^(how\s+(do\s+you\s+|is\s+|are\s+)?calculate\s+(the\s+)?)/i,'')
.replace(/^(describe\s+(the\s+|me\s+)?)/i,'')
.replace(/^(meaning\s+of\s+(the\s+)?)/i,'')
.replace(/[?!.]+$/,'').replace(/\s+/g,' ').trim()
}
function getCacheKey(clientId, query) {return `${clientId}:${normalizeQueryForCache(query)}`}
function validateQuery(query) {
if (!query || typeof query !== 'string') return {valid:false,message:'Please enter a complete question to get an accurate answer.'}
const trimmed = query.trim()
if (trimmed.length <= 1) return {valid:false,message:'Please enter a complete question to get an accurate answer.'}
const words = trimmed.split(/\s+/).filter(w => w.length > 0)
if (words.length < 2) return {valid:false,message:'Please enter a more detailed question so I can provide an accurate answer.'}
return {valid:true}
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
const a = m[1].trim().replace(/^(what\s+is\s+|the\s+)/i,'').trim()
const b = m[2].trim().replace(/^(what\s+is\s+|the\s+)/i,'').trim()
if (a.length > 1 && b.length > 1) return {isMulti:true,topics:[a,b],mode:'comparison'}
}
}
for (const p of andSplitPatterns) {
const m = q.match(p)
if (m) {
const a = m[1].trim().replace(/^(what\s+is\s+|what\s+are\s+|define\s+|the\s+)/i,'').trim()
const b = m[2].trim().replace(/^(what\s+is\s+|what\s+are\s+|define\s+|the\s+)/i,'').trim()
const stopWords = new Set(['is','are','was','were','it','this','that','its','my','your'])
if (a.length > 1 && b.length > 1 && !stopWords.has(a.toLowerCase()) && !stopWords.has(b.toLowerCase())) {
return {isMulti:true,topics:[a,b],mode:'multi_definition'}
}
}
}
return {isMulti:false,topics:[],mode:null}
}
function stripLocationQualifier(subject) {
const locationWords = new Set([
'siwan seva','siwan','seva','anurit','anurit innovation','akshay','akshay kumar',
'app','application','system','platform','service','portal','dashboard','module',
])
const stripped = subject.replace(LOCATION_QUALIFIER_PATTERN, '').trim()
if (stripped.length > 2 && stripped !== subject) return stripped
const personStripped = subject.replace(PERSON_QUALIFIER_PATTERN, '').trim()
if (personStripped.length > 2 && personStripped !== subject) return personStripped
return subject
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
/^compare\s+(.+?)\s+(?:vs|versus)\s+(.+)$/i,
/^difference\s+between\s+(.+?)\s+and\s+(.+)$/i,
/^what\s+(?:is|are)\s+(?:an?\s+|the\s+)?(.+)$/i,
/^(?:what\s+is\s+)?(.+)$/i,
]
for (const p of patterns) {
const m = q.match(p)
if (m) {
const raw = m[1].trim().replace(/[?!.]+$/,'').trim()
if (raw.length > 0) return stripLocationQualifier(raw)
}
}
return stripLocationQualifier(q.replace(/[?!.]+$/,'').trim())
}
function extractUrlKeywords(query) {
const stopWords = new Set(['power','bi','report','url','link','for','the','a','an','of','in','get','me','show','give','find','fetch'])
return query.toLowerCase().replace(/[^\w\s-]/g,' ').split(/\s+/).filter(w => w.length > 1 && !stopWords.has(w))
}
function normalizeTerms(term) {
const t = term.toLowerCase().trim()
const variants = new Set([t])
if (t.endsWith('s')) variants.add(t.slice(0,-1))
else variants.add(t + 's')
if (t.endsWith('ies')) variants.add(t.slice(0,-3) + 'y')
if (t.endsWith('y')) variants.add(t.slice(0,-1) + 'ies')
return [...variants]
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
['non-recurring','recurring'],['non recurring','recurring'],
['denied','approved'],['inactive','active'],
['rejected','accepted'],['unapproved','approved'],
['unpaid','paid'],['cancelled','active'],
['canceled','active'],['delinquent','current'],
['non-',''],
]
function computeNegativePenalty(querySubject, chunkText) {
const qs = querySubject.toLowerCase()
const ct = chunkText.toLowerCase()
let penalty = 0
for (const [negTerm, posTerm] of NEGATIVE_PAIRS) {
if (posTerm === null || posTerm === undefined) continue
const queryHasPositive = posTerm.length > 0 && new RegExp(`\\b${escapeRegex(posTerm)}\\b`,'i').test(qs)
const queryHasNegative = new RegExp(`\\b${escapeRegex(negTerm)}\\b`,'i').test(qs)
if (queryHasPositive && !queryHasNegative) {
if (new RegExp(`\\b${escapeRegex(negTerm)}\\b`,'i').test(ct)) penalty += 30
}
if (queryHasNegative) {
if (posTerm.length > 0 && !new RegExp(`\\b${escapeRegex(negTerm)}\\b`,'i').test(ct) && new RegExp(`\\b${escapeRegex(posTerm)}\\b`,'i').test(ct)) penalty += 20
}
}
return penalty
}
function buildInvertedIndex(chunks) {
const index = new Map()
for (let i = 0; i < chunks.length; i++) {
const words = (chunks[i].text || '').toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
for (const w of words) {
if (w.length < 2) continue
if (!index.has(w)) index.set(w, new Set())
index.get(w).add(i)
}
}
return index
}
function selectFocusedHits(hits, topK = 6) {
if (!hits || hits.length === 0) return []
const byFile = {}
for (const h of hits) {
const f = h.source_file || 'unknown'
if (!byFile[f]) byFile[f] = []
byFile[f].push(h)
}
const files = Object.keys(byFile)
if (files.length === 1) return hits.slice(0, topK)
const fileScores = files.map(f => {
const fhits = byFile[f]
const topScore = fhits[0]?._score || 0
const top2avg = fhits.slice(0, 2).reduce((s, h) => s + (h._score || 0), 0) / Math.min(fhits.length, 2)
return { file: f, score: topScore * 0.6 + top2avg * 0.4, hits: fhits }
}).sort((a, b) => b.score - a.score)
const best = fileScores[0]
const runnerUp = fileScores[1]
if (runnerUp && best.score >= runnerUp.score * 1.8) {
return best.hits.slice(0, topK)
}
const fileCounts = {}
const result = []
for (const h of hits) {
const f = h.source_file || 'unknown'
fileCounts[f] = (fileCounts[f] || 0) + 1
if (f === best.file || fileCounts[f] <= Math.floor(topK * 0.3)) result.push(h)
if (result.length >= topK) break
}
return result
}
async function fetchWithTimeout(url, options, timeoutMs) {
const controller = new AbortController()
const timer = setTimeout(() => controller.abort(), timeoutMs)
try {
return await fetch(url, {...options, signal: controller.signal})
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
if (!res.headersSent) res.status(503).json({error:'Request timed out. Please try again.'})
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
function generateApiKey() {
const crypto = require('crypto')
return `rak_${crypto.randomBytes(32).toString('hex')}`
}
function generateTitle(query) {
const cleaned = query.trim().replace(/[?!.]+$/,'')
return cleaned.length > 50 ? cleaned.slice(0,50) + '...' : cleaned
}
module.exports = {
MONGODB_URI,MONGODB_DB,CHAT_HISTORY_URI,CHAT_HISTORY_DB,
AZURE_CONNECTION_STRING,AZURE_CONTAINER_NAME,ADMIN_API_KEY,
KEY_CHECK_INTERVAL_MS,ASKDATA_ENDPOINT,ASKDATA_KEY,ASKDATA_MODEL,
ASKDATA_TIMEOUT_MS,ASKDATA2_ENDPOINT,ASKDATA2_KEY,ASKDATA2_MODEL,
ASKDATA2_TIMEOUT_MS,ASKDATA2_REWRITE_TIMEOUT_MS,REQUEST_TIMEOUT_MS,
WARMUP_CLIENT_IDS,RAW_PREFIX,FAISS_PREFIX,BM25_PREFIX,
CHUNK_SIZE,CHUNK_OVERLAP,BLOB_CONCURRENCY,
CHUNK_CACHE_TTL,MAX_HITS_GLOBAL,FAISS_TOP_K,BM25_TOP_K,RERANK_TOP_K,
SUPPORTED_EXTENSIONS,SYNONYM_MAP,TYPO_MAP,
DOMAIN_SHORT_SAFELIST,RESPONSE_CACHE,
responseCacheGet,responseCacheSet,escapeRegex,capFirst,ensureSinglePeriod,
trimToCompleteSentence,trimToLastCompleteSentence,fixBrokenUrls,
applySynonyms,applyTypos,levenshteinSimilarity,normalizeQuery,
normalizeQueryForCache,getCacheKey,validateQuery,detectQueryIntent,
detectMultiTopicQuery,extractSubject,extractUrlKeywords,normalizeTerms,
extractFormulaFromText,computeNegativePenalty,buildInvertedIndex,
selectFocusedHits,
fetchWithTimeout,withRequestTimeout,generateApiKey,generateTitle,
}
