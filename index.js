require('dotenv').config()
const express = require('express')
const cors = require('cors')
const {MongoClient, ObjectId} = require('mongodb')
const {BlobServiceClient} = require('@azure/storage-blob')
const pdfParse = require('pdf-parse')
const mammoth = require('mammoth')
const {parse: htmlParse} = require('node-html-parser')
const yaml = require('js-yaml')
const Papa = require('papaparse')
const {simpleParser} = require('mailparser')
const {parseOffice} = require('officeparser')
const crypto = require('crypto')
const {
MONGODB_URI,MONGODB_DB,CHAT_HISTORY_URI,CHAT_HISTORY_DB,
AZURE_CONNECTION_STRING,AZURE_CONTAINER_NAME,ADMIN_API_KEY,
KEY_CHECK_INTERVAL_MS,ASKDATA_ENDPOINT,ASKDATA_KEY,ASKDATA_MODEL,
ASKDATA_TIMEOUT_MS,ASKDATA2_ENDPOINT,ASKDATA2_KEY,ASKDATA2_MODEL,
ASKDATA2_TIMEOUT_MS,REQUEST_TIMEOUT_MS,WARMUP_CLIENT_IDS,
RAW_PREFIX,BLOB_CONCURRENCY,CHUNK_CACHE_TTL,MAX_HITS_GLOBAL,
SUPPORTED_EXTENSIONS,
responseCacheGet,responseCacheSet,getCacheKey,escapeRegex,capFirst,
ensureSinglePeriod,fixBrokenUrls,trimToLastCompleteSentence,
applySynonyms,applyTypos,normalizeQuery,validateQuery,
detectQueryIntent,detectMultiTopicQuery,extractSubject,
buildInvertedIndex,fetchWithTimeout,withRequestTimeout,
generateApiKey,generateTitle,
} = require('./src/config')
const {
extractSpreadsheet,retrieveChunksDD,buildSystemPromptDD,buildUserMessageDD,
buildFallbackAnswerDD,fuzzyCorrectQuery,relaxedKeywordSearchDD,extractAllUrlsFromChunks,
} = require('./src/dd')
const {
chunkStructuredDocument,buildInvertedIndexSF,retrieveChunksSF,
buildSystemPromptSF,buildUserMessageSF,buildFallbackAnswerSF,
} = require('./src/sf')
const {
slidingWindowChunk,buildInvertedIndexUD,retrieveChunksUD,preprocessQueryUD,
buildSystemPromptUD,buildUserMessageUD,buildFallbackAnswerUD,cleanOcrNoise,
} = require('./src/ud')
const app = express()
const allowedOrigins = [
'http://localhost:8080','http://localhost:3000',
'https://app.powerbi.com','https://msit.powerbi.com',
'https://anuritchat.vercel.app','https://ragadminpanel.vercel.app',
'https://df.powerbi.com','https://www.anuritinnovation.com/',
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
origin:(origin, callback) => callback(null, originAllowed(origin)),
methods:['GET','POST','PUT','PATCH','DELETE','OPTIONS'],
allowedHeaders:['Content-Type','Authorization','x-session-id'],
credentials:true,
}))
app.options('*', cors({
origin:(origin, callback) => callback(null, true),
methods:['GET','POST','PUT','PATCH','DELETE','OPTIONS'],
allowedHeaders:['Content-Type','Authorization','x-session-id'],
credentials:true,
}))
app.use(express.json())
const blobServiceClient = AZURE_CONNECTION_STRING ? BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING) : null
let askedataActiveCount = 0
const ASKDATA_MAX_CONCURRENT = 3
const askedataQueue = []
function runWithAskedataLimit(fn) {
return new Promise((resolve, reject) => {
function tryRun() {
if (askedataActiveCount < ASKDATA_MAX_CONCURRENT) {
askedataActiveCount++
Promise.resolve().then(fn).then(
result => {askedataActiveCount--;drainAskedataQueue();resolve(result)},
err => {askedataActiveCount--;drainAskedataQueue();reject(err)}
)
} else {
askedataQueue.push(tryRun)
}
}
tryRun()
})
}
function drainAskedataQueue() {
if (askedataQueue.length > 0 && askedataActiveCount < ASKDATA_MAX_CONCURRENT) askedataQueue.shift()()
}
let askedataFailures = 0
let askedataBlockedUntil = 0
function askedataCircuitOpen() {
if (Date.now() < askedataBlockedUntil) return true
if (askedataBlockedUntil > 0) {askedataBlockedUntil = 0;askedataFailures = 0}
return false
}
function askedataRecordSuccess() {askedataFailures = 0;askedataBlockedUntil = 0}
function askedataRecordFailure() {
askedataFailures++
if (askedataFailures >= 3) {
askedataBlockedUntil = Date.now() + 30000
console.error('[ASKDATA] Circuit breaker OPEN for 30s')
}
}
async function callASKDATA(systemPrompt, userMessage, maxTokens = 1024) {
if (!ASKDATA_ENDPOINT || !ASKDATA_KEY) throw new Error('ASKDATA_ENDPOINT and ASKDATA_KEY are required')
if (askedataCircuitOpen()) throw new Error('ASKDATA temporarily unavailable')
return runWithAskedataLimit(async () => {
try {
const response = await fetchWithTimeout(
ASKDATA_ENDPOINT,
{
method:'POST',
headers:{'Content-Type':'application/json','Authorization':`Bearer ${ASKDATA_KEY}`},
body:JSON.stringify({model:ASKDATA_MODEL,messages:[{role:'system',content:systemPrompt},{role:'user',content:userMessage}],temperature:0.1,max_tokens:maxTokens}),
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
method:'POST',
headers:{'Content-Type':'application/json','Authorization':`Bearer ${ASKDATA2_KEY}`,'Accept':'application/json'},
body:JSON.stringify({model:ASKDATA2_MODEL,messages:[{role:'system',content:systemPrompt},{role:'user',content:userMessage}],max_tokens:maxTokens,temperature:0.1,top_p:1.0,stream:false}),
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
function classifyDocumentType(chunks, fileName) {
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
const codeExts = new Set(['.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h','.cs','.go','.rb','.php','.swift','.kt','.r','.sql','.sh','.bash','.ps1'])
const spreadsheetExts = new Set(['.xlsx','.xls','.ods','.csv','.tsv'])
if (codeExts.has(ext)) return 'code'
if (spreadsheetExts.has(ext)) return 'data_dictionary'
if (!chunks || chunks.length === 0) return 'unstructured'
const metaChunks = chunks.filter(c => c.metadata && c.metadata.measure).length
if (metaChunks / Math.max(chunks.length,1) > 0.4) return 'data_dictionary'
const sample = chunks.slice(0,10).map(c => c.text || '').join('\n')
const ddSignals = [/is defined as:/i,/formula\s*:/i,/how to calculate/i,/report url for/i,/power bi link for/i].filter(p => p.test(sample)).length
if (ddSignals >= 2) return 'data_dictionary'
const sfSignals = [/^#{1,6}\s+\S/m,/^(Abstract|Introduction|Background|Methodology|Methods|Results|Discussion|Conclusion|References|Summary|Executive Summary|Key Findings)\b/im,/\|\s*\S+\s*\|/,/^\d+\.\d+\s+[A-Z]/m].filter(p => p.test(sample)).length
if (sfSignals >= 2) return 'structured'
const codeSignals = [/^(function|class|def|import|const|let|var|async)\s/m,/=>\s*{/,/\bpublic\s+(static\s+)?\w+\s+\w+\s*\(/m].filter(p => p.test(sample)).length
if (codeSignals >= 2) return 'code'
const hasLongParagraphs = chunks.filter(c => (c.text || '').length > 300).length > chunks.length * 0.4
if (hasLongParagraphs) return 'unstructured'
return 'unstructured'
}
function detectDocTypeFromChunks(chunks) {
if (!chunks || chunks.length === 0) return 'unstructured'
const ddCount = chunks.filter(c => c.metadata && c.metadata.measure).length
if (ddCount / Math.max(chunks.length,1) > 0.4) return 'data_dictionary'
const sfCount = chunks.filter(c => c.metadata && c.metadata.section).length
if (sfCount / Math.max(chunks.length,1) > 0.3) return 'structured'
return 'unstructured'
}
async function extractPdf(buffer) {const r = await pdfParse(buffer);return r.text || ''}
async function extractWord(buffer) {
  const htmlResult = await mammoth.convertToHtml({buffer})
  const html = htmlResult.value
  const withTables = html.replace(/<table>([\s\S]*?)<\/table>/g, (_, tableInner) => {
    const rows = []
    const trMatches = tableInner.match(/<tr>([\s\S]*?)<\/tr>/g) || []
    let headers = []
    trMatches.forEach((tr, rowIdx) => {
      const cells = (tr.match(/<t[dh]>([\s\S]*?)<\/t[dh]>/g) || [])
        .map(cell => cell.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim())
        .filter(Boolean)
      if (rowIdx === 0) {
        headers = cells
      } else {
        rows.push(cells.map((val, i) => `${headers[i] || 'Col' + (i + 1)}: ${val}`).join(' | '))
      }
    })
    return '\n' + rows.join('\n') + '\n'
  })
  const plain = withTables
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/h[1-6]>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&#\d+;/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
  const lines = plain.split('\n')
  const out = []
  let buf = []

  const isSectionHeading = (line) => {
    const t = line.trim()
    if (!t || t.length > 60 || t.includes('|') || t.includes(':')) return false
    return /^[A-Z][A-Za-z\s]+$/.test(t) || /^(Compensation|Bonus|Other|Notice|Employee|Offer|Letter|Designation|Department|Location|Joining)\b/i.test(t)
  }

  for (const line of lines) {
    if (isSectionHeading(line) && buf.length > 0) {
      const combined = buf.join('\n').trim()
      if (combined.length > 20) out.push(combined)
      buf = [line]
    } else {
      buf.push(line)
    }
  }
  if (buf.length > 0) {
    const combined = buf.join('\n').trim()
    if (combined.length > 20) out.push(combined)
  }
  return out.join('\n\n')
}
async function extractOffice(buffer) {
return new Promise((resolve, reject) => {
parseOffice(buffer, (text, err) => {if (err) reject(err);else resolve(text || '')},{outputErrorToConsole:false})
})
}
function extractHtml(buffer) {
const root = htmlParse(buffer.toString('utf-8'))
root.querySelectorAll('script, style').forEach(n => n.remove())
return root.structuredText || root.innerText || root.rawText || ''
}
function extractXml(buffer) {return buffer.toString('utf-8').replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim()}
function extractJson(buffer) {try {return JSON.stringify(JSON.parse(buffer.toString('utf-8')),null,2)} catch {return buffer.toString('utf-8')}}
function extractJsonl(buffer) {
return buffer.toString('utf-8').split('\n').filter(Boolean).map(line => {try {return JSON.stringify(JSON.parse(line))} catch {return line}}).join('\n')
}
function extractYaml(buffer) {try {return JSON.stringify(yaml.load(buffer.toString('utf-8')),null,2)} catch {return buffer.toString('utf-8')}}
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
parseOffice(buffer, (text, err) => {resolve(err || !text ? '[EPUB: convert to PDF for best results]' : text)},{outputErrorToConsole:false})
})
}
async function extractTextFromBuffer(buffer, fileName) {
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
if (ext === '.pdf') return extractPdf(buffer)
if (ext === '.docx' || ext === '.doc') return extractWord(buffer)
if (ext === '.odt' || ext === '.rtf') return extractOffice(buffer)
if (['.xlsx','.xls','.ods'].includes(ext)) return null
if (ext === '.csv') return extractCsv(buffer,',')
if (ext === '.tsv') return extractCsv(buffer,'\t')
if (ext === '.pptx' || ext === '.ppt') return extractOffice(buffer)
if (ext === '.html' || ext === '.htm') return extractHtml(buffer)
if (ext === '.xml') return extractXml(buffer)
if (['.md','.markdown','.rst'].includes(ext)) return buffer.toString('utf-8')
if (ext === '.json') return extractJson(buffer)
if (ext === '.jsonl') return extractJsonl(buffer)
if (ext === '.yaml' || ext === '.yml') return extractYaml(buffer)
if (ext === '.toml') return buffer.toString('utf-8')
if (ext === '.epub') return extractEpub(buffer)
if (ext === '.eml') return extractEml(buffer)
const plainText = new Set(['.txt','.py','.js','.ts','.jsx','.tsx','.java','.cpp','.c','.h','.cs','.go','.rb','.php','.swift','.kt','.r','.sql','.sh','.bash','.ps1'])
if (plainText.has(ext)) return buffer.toString('utf-8')
return ''
}
function sniffDocumentType(text, fileName) {
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
const sfExts = new Set(['.pdf','.docx','.doc','.odt','.rtf','.pptx','.ppt','.md','.markdown','.rst','.html','.htm'])
if (!sfExts.has(ext)) return 'unstructured'
const lines = text.split('\n').slice(0,50)
let headingCount = 0
let tableCount = 0
let longParaCount = 0
for (const line of lines) {
if (/^#{1,6}\s+\S/.test(line.trim())) headingCount++
if (/^(Abstract|Introduction|Background|Methodology|Methods|Results|Discussion|Conclusion|References|Summary|Executive Summary|Key Findings|Recommendations)\b/i.test(line.trim())) headingCount++
if (/\|.*\|/.test(line)) tableCount++
if (line.length > 200) longParaCount++
}
const sampleLower = text.slice(0,3000).toLowerCase()
const ddSignals = [/is defined as/,/formula:/,/how to calculate/,/measure name/,/attribute name/].filter(p => p.test(sampleLower)).length
if (ddSignals >= 2) return 'data_dictionary'
if (headingCount >= 3 || tableCount >= 2) return 'structured'
if (longParaCount >= 5) return 'unstructured'
return 'unstructured'
}
function chunkTextLegacy(text, sourceFile) {
const CHUNK_SIZE_L = 1200
const CHUNK_OVERLAP_L = 2
const chunks = []
let chunkIndex = 0
const blocks = text.replace(/\r\n/g,'\n').split(/\n{2,}/).map(b => b.trim()).filter(b => b.length > 0)
let buffer = []
let bufferLength = 0
function flush() {
const chunkStr = buffer.join('\n\n')
if (chunkStr.length >= 30) chunks.push({text:chunkStr,source_file:sourceFile,chunk_index:chunkIndex++,embedding:[]})
buffer = []
bufferLength = 0
}
for (let bi = 0; bi < blocks.length; bi++) {
const block = blocks[bi]
if (block.length > CHUNK_SIZE_L * 1.5) {
if (buffer.length > 0) flush()
const lines = block.split('\n').filter(l => l.trim())
let lineBuffer = []
let lineLength = 0
for (const line of lines) {
const projected = lineLength + (lineBuffer.length ? 1 : 0) + line.length
if (lineBuffer.length > 0 && projected > CHUNK_SIZE_L) {
const s = lineBuffer.join('\n')
if (s.length >= 30) chunks.push({text:s,source_file:sourceFile,chunk_index:chunkIndex++,embedding:[]})
lineBuffer = lineBuffer.slice(-CHUNK_OVERLAP_L)
lineLength = lineBuffer.join('\n').length
}
lineBuffer.push(line)
lineLength += (lineLength ? 1 : 0) + line.length
}
if (lineBuffer.length) {
const s = lineBuffer.join('\n')
if (s.length >= 30) chunks.push({text:s,source_file:sourceFile,chunk_index:chunkIndex++,embedding:[]})
}
continue
}
const projected = bufferLength + (bufferLength ? 2 : 0) + block.length
if (buffer.length > 0 && projected > CHUNK_SIZE_L) {
const lastBlock = buffer[buffer.length - 1] || ''
flush()
if (lastBlock) {buffer.push(lastBlock);bufferLength = lastBlock.length}
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
for await (const blob of containerClient.listBlobsFlat({prefix})) {
const fileName = blob.name.split('/').pop()
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
if (SUPPORTED_EXTENSIONS.has(ext)) blobNames.push(blob.name)
}
const allChunks = []
let chunkIndexOffset = 0
for (let i = 0; i < blobNames.length; i += BLOB_CONCURRENCY) {
const batch = blobNames.slice(i, i + BLOB_CONCURRENCY)
const results = await Promise.allSettled(batch.map(async (blobName) => {
const fileName = blobName.split('/').pop()
const ext = ('.' + fileName.split('.').pop()).toLowerCase()
const buffer = await downloadBlobAsBuffer(containerClient, blobName)
if (['.xlsx','.xls','.ods'].includes(ext)) {
const structuredRows = extractSpreadsheet(buffer)
return structuredRows.map((row,idx) => ({
text:row.text,source_file:fileName,chunk_index:idx,embedding:[],
metadata:row.metadata || null,docType:'data_dictionary',
}))
}
const text = await extractTextFromBuffer(buffer, fileName)
if (!text?.trim()) return []
const docTypeHint = sniffDocumentType(text, fileName)
let rawChunks = []
if (docTypeHint === 'data_dictionary') {
rawChunks = chunkTextLegacy(text, fileName)
} else if (docTypeHint === 'structured') {
rawChunks = chunkStructuredDocument(text, fileName)
} else {
rawChunks = slidingWindowChunk(text, fileName)
}
const finalDocType = docTypeHint === 'data_dictionary' ? 'data_dictionary' : classifyDocumentType(rawChunks, fileName)
return rawChunks.map(c => ({...c, docType:finalDocType}))
}))
for (const result of results) {
if (result.status === 'fulfilled') {
const fileChunks = result.value
fileChunks.forEach((c,idx) => {c.chunk_index = chunkIndexOffset + idx})
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
if (now - cached.ts <= CHUNK_CACHE_TTL) return {chunks:cached.chunks,invertedIndexes:cached.invertedIndexes}
if (!cached.loading) {
const refreshPromise = _doLoadChunks(clientId)
.then(chunks => {
const invertedIndexes = buildAllInvertedIndexes(chunks)
CHUNK_CACHE.set(clientId, {chunks,invertedIndexes,ts:Date.now(),loading:null})
console.log(`[chunkCache] Background refresh done for ${clientId}: ${chunks.length} chunks`)
})
.catch(err => {
const existing = CHUNK_CACHE.get(clientId)
if (existing) CHUNK_CACHE.set(clientId, {...existing,loading:null})
console.warn(`[chunkCache] Background refresh failed for ${clientId}: ${err.message}`)
})
CHUNK_CACHE.set(clientId, {...cached,loading:refreshPromise})
}
return {chunks:cached.chunks,invertedIndexes:cached.invertedIndexes}
}
if (cached && cached.loading) {
await cached.loading
const entry = CHUNK_CACHE.get(clientId)
return {chunks:entry?.chunks || [],invertedIndexes:entry?.invertedIndexes || {}}
}
const loadPromise = _doLoadChunks(clientId)
.then(chunks => {
const invertedIndexes = buildAllInvertedIndexes(chunks)
CHUNK_CACHE.set(clientId, {chunks,invertedIndexes,ts:Date.now(),loading:null})
console.log(`[chunkCache] Loaded ${chunks.length} chunks for ${clientId}`)
return chunks
})
.catch(err => {
CHUNK_CACHE.set(clientId, {chunks:null,invertedIndexes:{},ts:0,loading:null})
throw err
})
CHUNK_CACHE.set(clientId, {chunks:null,invertedIndexes:{},ts:0,loading:loadPromise})
await loadPromise
const entry = CHUNK_CACHE.get(clientId)
return {chunks:entry?.chunks || [],invertedIndexes:entry?.invertedIndexes || {}}
}
function buildAllInvertedIndexes(chunks) {
const ddChunks = chunks.filter(c => c.docType === 'data_dictionary')
const sfChunks = chunks.filter(c => c.docType === 'structured')
const udChunks = chunks.filter(c => c.docType !== 'data_dictionary' && c.docType !== 'structured')
return {
dd: ddChunks.length > 0 ? buildInvertedIndex(ddChunks) : buildInvertedIndex(chunks),
sf: sfChunks.length > 0 ? buildInvertedIndexSF(sfChunks) : null,
ud: udChunks.length > 0 ? buildInvertedIndexUD(udChunks) : null,
all: buildInvertedIndex(chunks),
}
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
.then(({chunks}) => console.log(`[warmup] ${clientId} -- ${chunks.length} chunks ready`))
.catch(err => console.warn(`[warmup] ${clientId} -- ${err.message}`))
}
}
let db = null
async function getDb() {
if (db) return db
const client = new MongoClient(MONGODB_URI)
await client.connect()
db = client.db(MONGODB_DB)
await db.collection('clients').createIndex({apiKey:1},{unique:true,sparse:true})
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
if (Date.now() - entry.cachedAt > CACHE_TTL_MS) {CLIENT_CACHE.delete(apiKey);return null}
return entry
}
function setCache(apiKey, data) {CLIENT_CACHE.set(apiKey, {...data,cachedAt:Date.now()})}
function evictCache(apiKey) {if (apiKey) CLIENT_CACHE.delete(apiKey)}
async function verifyApiKey(apiKey) {
if (!apiKey || !apiKey.startsWith('rak_')) return null
const cached = getCached(apiKey)
if (cached) return {clientId:cached.clientId,name:cached.name}
const database = await getDb()
const client = await database.collection('clients').findOne({apiKey},{projection:{clientId:1,name:1,_id:0}})
if (!client) return null
setCache(apiKey, {clientId:client.clientId,name:client.name})
return {clientId:client.clientId,name:client.name}
}
function startApiKeyHealthChecker() {
if (!MONGODB_URI) return
setInterval(async () => {
const keys = [...CLIENT_CACHE.keys()]
if (!keys.length) return
try {
const database = await getDb()
const validDocs = await database.collection('clients').find({apiKey:{$in:keys}},{projection:{apiKey:1,_id:0}}).toArray()
const validSet = new Set(validDocs.map(d => d.apiKey))
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
if (!apiKey) return res.status(401).json({error:'Missing API key'})
const client = await verifyApiKey(apiKey)
if (!client) return res.status(401).json({error:'Invalid or expired API key'})
req.client = client
next()
}
function requireAdminKey(req, res, next) {
const key = extractApiKey(req)
if (!key || key !== ADMIN_API_KEY) return res.status(401).json({error:'Unauthorized'})
next()
}
function cleanAnswer(rawAnswer) {
if (!rawAnswer) return ''
let cleaned = fixBrokenUrls(rawAnswer)
.replace(/^\s*\[Source\s*\d+\]\s*/gm,'')
.replace(/^[^\n]*(\|[^\n]*){3,}$/gm,'')
.replace(/=== .+ ===\s*/gm,'')
.replace(/\(from\s+[A-Za-z\s]+\)\s*/g,'')
.replace(/\n{3,}/g,'\n\n')
.replace(/\.{2,}/g,'.')
.replace(/\.\s*\./g,'.')
.trim()
cleaned = trimToLastCompleteSentence(cleaned)
if (cleaned.length > 0 && !/[.!?]$/.test(cleaned)) cleaned += '.'
return ensureSinglePeriod(cleaned)
}
async function generateAnswer(query, hits, intent, docType) {
let systemPrompt, userMessage
if (docType === 'data_dictionary') {
systemPrompt = buildSystemPromptDD(intent)
userMessage = buildUserMessageDD(query, hits, intent)
} else if (docType === 'structured') {
systemPrompt = buildSystemPromptSF(intent)
userMessage = buildUserMessageSF(query, hits, intent)
} else {
systemPrompt = buildSystemPromptUD(intent)
userMessage = buildUserMessageUD(query, hits, intent)
}
return callBestAvailableEngine(systemPrompt, userMessage, 1024)
}
function buildFallbackAnswer(query, hits, intent, docType) {
if (docType === 'data_dictionary') return buildFallbackAnswerDD(query, hits, intent)
if (docType === 'structured') return buildFallbackAnswerSF(query, hits)
return buildFallbackAnswerUD(query, hits)
}
async function retrieveHitsForDocType(processedQuery, chunks, topK, invertedIndexes, docType, intent) {
if (docType === 'data_dictionary') {
const ddChunks = chunks.filter(c => c.docType === 'data_dictionary')
const idx = invertedIndexes.dd
const hits = await retrieveChunksDD(processedQuery, ddChunks.length > 0 ? ddChunks : chunks, topK, idx)
if (hits.length === 0) return relaxedKeywordSearchDD(processedQuery, ddChunks.length > 0 ? ddChunks : chunks, 64, idx)
return hits
}
if (docType === 'structured') {
const sfChunks = chunks.filter(c => c.docType === 'structured')
const idx = invertedIndexes.sf || invertedIndexes.all
return retrieveChunksSF(processedQuery, sfChunks.length > 0 ? sfChunks : chunks, topK, idx)
}
const udChunks = chunks.filter(c => c.docType !== 'data_dictionary' && c.docType !== 'structured')
const idx = invertedIndexes.ud || invertedIndexes.all
return retrieveChunksUD(processedQuery, udChunks.length > 0 ? udChunks : chunks, topK, idx)
}
function computePoolConfidence(hits, docType) {
if (!hits || hits.length === 0) return 0
const topScore = hits[0]?._score || 0
const secondScore = hits[1]?._score || 0
const gap = topScore - secondScore
const MIN_THRESHOLD = 3
if (topScore < MIN_THRESHOLD) return 0
const topN = hits.slice(0, Math.min(5, hits.length))
const avgScore = topN.reduce((s, h) => s + (h._score || 0), 0) / topN.length
let confidence = avgScore + gap * 0.5
if (docType === 'data_dictionary') confidence *= 1.4
return confidence
}
async function retrieveBestHitsAcrossAllTypes(processedQuery, chunks, topK, invertedIndexes, intent) {
const ddChunks = chunks.filter(c => c.docType === 'data_dictionary')
const sfChunks = chunks.filter(c => c.docType === 'structured')
const udChunks = chunks.filter(c => c.docType !== 'data_dictionary' && c.docType !== 'structured')

const [ddHits, sfHits, udHits] = await Promise.all([
ddChunks.length > 0
? retrieveChunksDD(processedQuery, ddChunks, topK, invertedIndexes.dd).catch(() => [])
: Promise.resolve([]),
sfChunks.length > 0
? Promise.resolve(retrieveChunksSF(processedQuery, sfChunks, topK, invertedIndexes.sf || invertedIndexes.all))
: Promise.resolve([]),
udChunks.length > 0
? Promise.resolve(retrieveChunksUD(processedQuery, udChunks, topK, invertedIndexes.ud || invertedIndexes.all))
: Promise.resolve([]),
])

const pools = [
{hits: ddHits, docType: 'data_dictionary'},
{hits: sfHits, docType: 'structured'},
{hits: udHits, docType: 'unstructured'},
].filter(p => p.hits && p.hits.length > 0)

if (pools.length === 0) return {hits: [], docType: 'unstructured'}

for (const pool of pools) {
pool.confidence = computePoolConfidence(pool.hits, pool.docType)
console.log(`[routing] docType=${pool.docType} topScore=${pool.hits[0]?._score?.toFixed(2)||0} confidence=${pool.confidence.toFixed(2)}`)
}

const validPools = pools.filter(p => p.confidence > 0)
if (validPools.length === 0) {
pools.sort((a, b) => (b.hits[0]?._score || 0) - (a.hits[0]?._score || 0))
return {hits: pools[0].hits, docType: pools[0].docType}
}

validPools.sort((a, b) => b.confidence - a.confidence)
return {hits: validPools[0].hits, docType: validPools[0].docType}
}
async function generateAnswerForTopic(topic, chunks, topK, invertedIndexes) {
const topicQuery = `what is ${topic}`
const docType = detectDocTypeFromChunks(chunks)
let hits = await retrieveHitsForDocType(topicQuery, chunks, topK, invertedIndexes, docType, 'definition')
if (hits.length === 0) hits = relaxedKeywordSearchDD(topicQuery, chunks, 32, invertedIndexes.all)
if (hits.length === 0) return null
let rawAnswer = ''
try {
rawAnswer = await Promise.race([
generateAnswer(topicQuery, hits, 'definition', docType),
new Promise((_,reject) => setTimeout(() => reject(new Error('timeout')),25000)),
])
} catch (err) {
console.warn(`[generateAnswerForTopic] Engine failed for "${topic}": ${err.message}`)
}
const isBlank = !rawAnswer || rawAnswer.trim().length < 15
let answer = isBlank ? buildFallbackAnswer(topicQuery, hits, 'definition', docType) : cleanAnswer(rawAnswer)
answer = answer.replace(/^\*\*[^*]+\*\*\s*(is defined as:?\s*)?/i,'').trim()
if (answer && !/[.!?]$/.test(answer)) answer += '.'
return answer
}
async function generateComparisonAnswer(topicA, topicB, chunks, topK, invertedIndexes) {
const comparisonQuery = `difference between ${topicA} and ${topicB}`
const docType = detectDocTypeFromChunks(chunks)
const hitsA = await retrieveHitsForDocType(`what is ${topicA}`, chunks, topK, invertedIndexes, docType, 'definition')
const hitsB = await retrieveHitsForDocType(`what is ${topicB}`, chunks, topK, invertedIndexes, docType, 'definition')
const allHits = [...hitsA,...hitsB]
const seen = new Set()
const deduped = []
for (const h of allHits) {
const fp = (h.text || '').trim().slice(0,80).toLowerCase()
if (!seen.has(fp)) {seen.add(fp);deduped.push(h)}
}
if (deduped.length === 0) return null
let rawAnswer = ''
try {
rawAnswer = await Promise.race([
generateAnswer(comparisonQuery, deduped, 'comparison', docType),
new Promise((_,reject) => setTimeout(() => reject(new Error('timeout')),25000)),
])
} catch (err) {
console.warn(`[generateComparisonAnswer] Engine failed for "${topicA}" vs "${topicB}": ${err.message}`)
}
if (rawAnswer && rawAnswer.trim().length >= 15) return cleanAnswer(rawAnswer)
const answerA = await generateAnswerForTopic(topicA, chunks, topK, invertedIndexes)
const answerB = await generateAnswerForTopic(topicB, chunks, topK, invertedIndexes)
const parts = []
if (answerA && !answerA.includes('could not find')) parts.push(`**${capFirst(topicA)}:** ${answerA}`)
else parts.push(`**${capFirst(topicA)}:** I could not find information about "${capFirst(topicA)}" in your documents.`)
if (answerB && !answerB.includes('could not find')) parts.push(`**${capFirst(topicB)}:** ${answerB}`)
else parts.push(`**${capFirst(topicB)}:** I could not find information about "${capFirst(topicB)}" in your documents.`)
return parts.join('\n\n')
}
async function handleMultiTopicQuery(topics, mode, chunks, topK, invertedIndexes) {
const results = await Promise.all(topics.map(async (topic) => {
const answer = await generateAnswerForTopic(topic, chunks, topK, invertedIndexes)
return {topic, answer}
}))
const parts = results.map(({topic, answer}) => {
const cap = capFirst(topic)
if (!answer || answer.includes('could not find') || answer.includes('not present')) {
return `**${cap}:**\nI could not find information about "${cap}" in your documents.`
}
return `**${cap}:**\n${answer}`
})
if (mode === 'comparison' && results.length === 2) {
const [a, b] = results
const bothFound = a.answer && !a.answer.includes('could not find') && b.answer && !b.answer.includes('could not find')
if (bothFound) {
const comparisonAnswer = await generateComparisonAnswer(a.topic, b.topic, chunks, topK, invertedIndexes)
if (comparisonAnswer && !comparisonAnswer.includes('could not find')) return comparisonAnswer
}
return parts.join('\n\n')
}
return parts.join('\n\n')
}
async function saveConversationMessage(clientId, conversationId, query, answer, sources) {
try {
const chatDatabase = await getChatDb()
const col = chatDatabase.collection('conversations')
const now = new Date()
const userMsg = {role:'user',content:query,timestamp:now}
const assistantMsg = {role:'assistant',content:answer,sources:sources.map(s => ({source_file:s.source_file,score:s.score})),timestamp:now}
let activeConversationId = conversationId || null
if (activeConversationId) {
const updated = await col.findOneAndUpdate(
{_id:new ObjectId(activeConversationId),clientId},
{$push:{messages:{$each:[userMsg,assistantMsg]}},$set:{updatedAt:now}},
{returnDocument:'after',projection:{_id:1}}
)
if (!updated) activeConversationId = null
}
if (!activeConversationId) {
const result = await col.insertOne({clientId,title:generateTitle(query),messages:[userMsg,assistantMsg],createdAt:now,updatedAt:now})
activeConversationId = result.insertedId.toString()
}
return activeConversationId
} catch (saveErr) {
console.warn('[saveConversationMessage] Failed:', saveErr.message)
return conversationId || null
}
}
const IN_FLIGHT = new Map()
app.get('/health', (req, res) => res.json({
ok:true,service:'ask-data',
engines:{primary:ASKDATA_ENDPOINT ? 'configured' : 'missing',fallback:ASKDATA2_ENDPOINT ? 'configured' : 'missing'},
chunkCacheSize:CHUNK_CACHE.size,
primaryCircuitOpen:askedataCircuitOpen(),
primaryFailures:askedataFailures,
maxHits:MAX_HITS_GLOBAL,
}))
app.post('/client/verify', async (req, res) => {
try {
const apiKey = extractApiKey(req) || req.body?.apiKey
if (!apiKey) return res.status(400).json({valid:false,error:'apiKey is required'})
const client = await verifyApiKey(apiKey)
if (!client) return res.status(401).json({valid:false,error:'Invalid or expired API key'})
res.json({valid:true,client})
} catch (err) {res.status(500).json({valid:false,error:err.message})}
})
app.post('/admin/clients', requireAdminKey, async (req, res) => {
try {
let {name, clientId, apiKey} = req.body
if (!name || !clientId) return res.status(400).json({error:'name and clientId are required'})
if (!apiKey) {apiKey = generateApiKey()}
else if (!apiKey.startsWith('rak_')) return res.status(400).json({error:'apiKey must start with "rak_"'})
const database = await getDb()
const col = database.collection('clients')
const existing = await col.findOne({$or:[{clientId},{apiKey}]})
if (existing) {
const field = existing.clientId === clientId ? 'clientId' : 'apiKey'
return res.status(409).json({error:`A client with this ${field} already exists`})
}
const now = new Date().toISOString()
const doc = {name:name.trim(),clientId:clientId.trim().toLowerCase(),apiKey,apiKeyRotatedAt:now,folderLink:'',sourceType:'google-drive',status:'idle',documentsCount:0,autoSync:false,watchIntervalMs:300000,lastRunAt:null,lastError:null,createdAt:now,updatedAt:now}
const result = await col.insertOne(doc)
res.status(201).json({...doc,_id:result.insertedId})
} catch (err) {res.status(500).json({error:err.message})}
})
app.get('/admin/clients', requireAdminKey, async (req, res) => {
try {
const database = await getDb()
const clients = await database.collection('clients').find({},{projection:{apiKey:0}}).sort({createdAt:-1}).toArray()
res.json({clients})
} catch (err) {res.status(500).json({error:err.message})}
})
app.get('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
try {
const database = await getDb()
const client = await database.collection('clients').findOne({clientId:req.params.clientId})
if (!client) return res.status(404).json({error:'Client not found'})
res.json(client)
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/admin/clients/:clientId/regenerate-key', requireAdminKey, async (req, res) => {
try {
const database = await getDb()
const col = database.collection('clients')
const oldClient = await col.findOne({clientId:req.params.clientId},{projection:{apiKey:1}})
if (!oldClient) return res.status(404).json({error:'Client not found'})
const newApiKey = generateApiKey()
const now = new Date().toISOString()
if (oldClient.apiKey) evictCache(oldClient.apiKey)
await col.findOneAndUpdate({clientId:req.params.clientId},{$set:{apiKey:newApiKey,apiKeyRotatedAt:now,updatedAt:now}},{returnDocument:'after'})
res.json({success:true,clientId:req.params.clientId,newApiKey,apiKeyRotatedAt:now})
} catch (err) {res.status(500).json({error:err.message})}
})
app.patch('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
try {
const database = await getDb()
const updates = {...req.body,updatedAt:new Date().toISOString()}
if (updates.apiKey !== undefined) {
if (!updates.apiKey.startsWith('rak_')) return res.status(400).json({error:'apiKey must start with "rak_"'})
const old = await database.collection('clients').findOne({clientId:req.params.clientId},{projection:{apiKey:1}})
if (old?.apiKey) evictCache(old.apiKey)
updates.apiKeyRotatedAt = new Date().toISOString()
}
const result = await database.collection('clients').findOneAndUpdate({clientId:req.params.clientId},{$set:updates},{returnDocument:'after'})
if (!result) return res.status(404).json({error:'Client not found'})
res.json(result)
} catch (err) {res.status(500).json({error:err.message})}
})
app.delete('/admin/clients/:clientId', requireAdminKey, async (req, res) => {
try {
const {clientId} = req.params
const database = await getDb()
const client = await database.collection('clients').findOne({clientId})
if (!client) return res.status(404).json({error:'Client not found'})
if (client.apiKey) evictCache(client.apiKey)
await database.collection('clients').deleteOne({clientId})
invalidateChunkCache(clientId)
const blobsDeleted = [], blobsFailed = []
if (blobServiceClient) {
try {
const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
for (const prefix of [`raw/${clientId}/`,`meta/${clientId}/`]) {
for await (const blob of containerClient.listBlobsFlat({prefix})) {
try {await containerClient.deleteBlob(blob.name);blobsDeleted.push(blob.name)}
catch (e) {blobsFailed.push({name:blob.name,error:e.message})}
}
}
} catch (azureErr) {blobsFailed.push({name:'azure-connection',error:azureErr.message})}
}
res.json({ok:true,deleted:clientId,blobsDeleted:blobsDeleted.length,blobsFailed:blobsFailed.length > 0 ? blobsFailed : undefined})
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/admin/clients/:clientId/invalidate-cache', requireAdminKey, (req, res) => {
invalidateChunkCache(req.params.clientId)
const {RESPONSE_CACHE} = require('./src/config')
RESPONSE_CACHE.clear()
res.json({ok:true,clientId:req.params.clientId,message:'Chunk + response cache invalidated'})
})
app.post('/client/login', async (req, res) => {
try {
const apiKey = extractApiKey(req) || req.body?.apiKey
if (!apiKey) return res.status(400).json({error:'apiKey is required'})
const client = await verifyApiKey(apiKey)
if (!client) return res.status(401).json({error:'Invalid API key'})
if (blobServiceClient) loadChunksForClient(client.clientId).catch(err => console.warn(`[login warmup] ${client.clientId}: ${err.message}`))
res.json({ok:true,client})
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/login', async (req, res) => {
try {
const apiKey = extractApiKey(req) || req.body?.apiKey
if (!apiKey) return res.status(400).json({error:'apiKey is required'})
const client = await verifyApiKey(apiKey)
if (!client) return res.status(401).json({error:'Invalid API key'})
if (blobServiceClient) loadChunksForClient(client.clientId).catch(err => console.warn(`[chat/login warmup] ${client.clientId}: ${err.message}`))
res.json({ok:true,client})
} catch (err) {res.status(500).json({error:err.message})}
})
app.get('/client/me', requireClientKey, async (req, res) => {
try {
const database = await getDb()
const client = await database.collection('clients').findOne({clientId:req.client.clientId},{projection:{apiKey:0}})
if (!client) return res.status(404).json({error:'Client not found'})
res.json(client)
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/conversations', requireClientKey, async (req, res) => {
try {
const {title} = req.body
const database = await getChatDb()
const now = new Date()
const conversation = {clientId:req.client.clientId,title:title||'New Conversation',messages:[],createdAt:now,updatedAt:now}
const result = await database.collection('conversations').insertOne(conversation)
res.status(201).json({...conversation,_id:result.insertedId})
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/list', requireClientKey, async (req, res) => {
try {
const database = await getChatDb()
const conversations = await database.collection('conversations').find({clientId:req.client.clientId},{projection:{messages:0}}).sort({updatedAt:-1}).toArray()
res.json({conversations})
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/get', requireClientKey, async (req, res) => {
try {
const {conversationId} = req.body
if (!conversationId) return res.status(400).json({error:'conversationId is required'})
const database = await getChatDb()
const conversation = await database.collection('conversations').findOne({_id:new ObjectId(conversationId),clientId:req.client.clientId})
if (!conversation) return res.status(404).json({error:'Conversation not found'})
res.json(conversation)
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/rename', requireClientKey, async (req, res) => {
try {
const {conversationId, title} = req.body
if (!conversationId || !title) return res.status(400).json({error:'conversationId and title are required'})
const database = await getChatDb()
const result = await database.collection('conversations').findOneAndUpdate(
{_id:new ObjectId(conversationId),clientId:req.client.clientId},
{$set:{title:title.trim(),updatedAt:new Date()}},
{returnDocument:'after',projection:{messages:0}}
)
if (!result) return res.status(404).json({error:'Conversation not found'})
res.json(result)
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/delete', requireClientKey, async (req, res) => {
try {
const {conversationId} = req.body
if (!conversationId) return res.status(400).json({error:'conversationId is required'})
const database = await getChatDb()
const result = await database.collection('conversations').deleteOne({_id:new ObjectId(conversationId),clientId:req.client.clientId})
if (result.deletedCount === 0) return res.status(404).json({error:'Conversation not found'})
res.json({ok:true,deleted:conversationId})
} catch (err) {res.status(500).json({error:err.message})}
})
app.post('/chat/message', requireClientKey, withRequestTimeout(async (req, res) => {
try {
const {query, topK = 6, conversationId} = req.body
if (!query?.trim()) return res.status(400).json({error:'query is required'})
const validation = validateQuery(query)
if (!validation.valid) return res.json({answer:validation.message,sources:[],conversationId:conversationId||null,client:req.client})
const {clientId, name} = req.client
const intent = detectQueryIntent(query.trim())
if (intent === 'greeting') {
return res.json({
answer:"Hello! I'm your document assistant. Ask me anything about your data, research papers, or documents.",
sources:[],conversationId:conversationId||null,client:{clientId,name},
})
}
const cacheKey = getCacheKey(clientId, query)
const cached = responseCacheGet(cacheKey)
if (cached) {
const activeConversationId = await saveConversationMessage(clientId, conversationId||null, query.trim(), cached.answer, cached.sources||[])
return res.json({...cached,cached:true,conversationId:activeConversationId})
}
if (IN_FLIGHT.has(cacheKey)) {
try {
const result = await IN_FLIGHT.get(cacheKey)
const activeConversationId = await saveConversationMessage(clientId, conversationId||null, query.trim(), result.answer, result.sources||[])
return res.json({...result,conversationId:activeConversationId})
} catch {}
}
const requestPromise = (async () => {
const {chunks, invertedIndexes} = await loadChunksForClient(clientId)
if (chunks.length === 0) return {answer:'No documents found for your account. Please ensure your documents have been ingested first.',sources:[],client:{clientId,name}}
let processedQuery = applyTypos(query.trim())
console.log(`[QueryPipeline] Original: "${query.trim()}"`)
if (processedQuery !== query.trim()) console.log(`[QueryPipeline] After typos: "${processedQuery}"`)
processedQuery = applySynonyms(processedQuery)

const hasSFChunks = chunks.some(c => c.docType === 'structured')
const hasUDChunks = chunks.some(c => c.docType !== 'data_dictionary' && c.docType !== 'structured')

if (hasSFChunks || hasUDChunks) {
const rewritten = await preprocessQueryUD(processedQuery)
if (rewritten !== processedQuery) console.log(`[QueryPipeline] After UD rewrite: "${rewritten}"`)
processedQuery = rewritten
} else {
const corrected = fuzzyCorrectQuery(processedQuery, chunks.filter(c => c.docType === 'data_dictionary'))
if (corrected !== processedQuery) console.log(`[QueryPipeline] After DD fuzzy: "${corrected}"`)
processedQuery = corrected
}

if (intent === 'all_urls') {
const urlChunks = chunks.filter(c => /https?:\/\/\S+/.test(c.text || ''))
const urlEntries = extractAllUrlsFromChunks(urlChunks)
const answer = urlEntries.length > 0 ? urlEntries.map(e => `**${e.name}:** ${e.url}`).join('\n') : 'I could not find any URLs in your documents.'
const sources = urlChunks.slice(0,6).map(h => ({source_file:h.source_file||'unknown',chunk_index:h.chunk_index??0,score:null,preview:(h.text||'').slice(0,200)}))
return {answer,sources,client:{clientId,name}}
}
const multiTopicCheck = detectMultiTopicQuery(processedQuery)
if (multiTopicCheck.isMulti) {
console.log(`[chat/message] Multi-topic detected: ${JSON.stringify(multiTopicCheck.topics)} mode=${multiTopicCheck.mode}`)
const answer = await handleMultiTopicQuery(multiTopicCheck.topics, multiTopicCheck.mode, chunks, Math.min(topK, MAX_HITS_GLOBAL), invertedIndexes)
return {answer,sources:[],client:{clientId,name}}
}

const {hits, docType: routedDocType} = await retrieveBestHitsAcrossAllTypes(processedQuery, chunks, Math.min(topK, MAX_HITS_GLOBAL), invertedIndexes, intent)

let finalHits = hits
let finalDocType = routedDocType

if (finalHits.length === 0) {
finalHits = relaxedKeywordSearchDD(processedQuery, chunks, 64, invertedIndexes.all)
finalDocType = detectDocTypeFromChunks(finalHits.length > 0 ? finalHits : chunks)
}

console.log(`[chat/message] "${query.slice(0,60)}" -> intent=${intent}, docType=${finalDocType}, hits=${finalHits.length}, topScore=${finalHits[0]?._score?.toFixed(2)||0}`)

if (finalHits.length === 0) return {answer:'I could not find relevant information about this in your documents. Try rephrasing your question.',sources:[],client:{clientId,name}}
let rawAnswer = ''
if (intent !== 'url_lookup') {
try {
rawAnswer = await Promise.race([
generateAnswer(processedQuery, finalHits, intent, finalDocType),
new Promise((_,reject) => setTimeout(() => reject(new Error('All engines timed out')),55000)),
])
} catch (err) {
console.warn(`[chat/message] All engines failed, using rule-based fallback: ${err.message}`)
}
}
const isBlank = !rawAnswer || rawAnswer.trim().length < 15
const answer = isBlank ? buildFallbackAnswer(processedQuery, finalHits, intent, finalDocType) : cleanAnswer(rawAnswer)
if (isBlank) console.warn(`[chat/message] Blank from all engines, used rule-based fallback for: "${query.slice(0,60)}"`)
const sources = finalHits.map(h => ({
source_file:h.source_file||'unknown',
chunk_index:h.chunk_index??0,
score:typeof h._score === 'number' ? parseFloat(h._score.toFixed(4)) : null,
preview:(h.text||'').slice(0,200),
}))
return {answer,sources,client:{clientId,name}}
})()
IN_FLIGHT.set(cacheKey, requestPromise)
let result
try {result = await requestPromise} finally {IN_FLIGHT.delete(cacheKey)}
if (result.answer && result.answer.length > 15) responseCacheSet(cacheKey, result)
const activeConversationId = await saveConversationMessage(clientId, conversationId||null, query.trim(), result.answer, result.sources||[])
res.json({...result,conversationId:activeConversationId})
} catch (err) {
console.error('[chat/message] Error:', err.message)
if (!res.headersSent) res.status(500).json({error:err.message})
}
}))
app.use((err, req, res, next) => {
console.error('[global error handler]', err)
if (!res.headersSent) res.status(500).json({error:'An unexpected error occurred. Please try again.'})
})
const PORT = process.env.PORT || 4000
app.listen(PORT, () => {
console.log(`Service running on port ${PORT}`)
console.log(`ASKDATA: ${ASKDATA_ENDPOINT ? 'configured' : 'MISSING'} | ASKDATA2: ${ASKDATA2_ENDPOINT ? 'configured' : 'missing'}`)
console.log(`MAX_HITS: ${MAX_HITS_GLOBAL}`)
startApiKeyHealthChecker()
warmupChunkCaches()
})
module.exports = app
