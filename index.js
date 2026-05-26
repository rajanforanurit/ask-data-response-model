require('dotenv').config()
const express=require('express')
const cors=require('cors')
const {MongoClient,ObjectId}=require('mongodb')
const {BlobServiceClient}=require('@azure/storage-blob')
const pdfParse=require('pdf-parse')
const mammoth=require('mammoth')
const {parse:htmlParse}=require('node-html-parser')
const crypto=require('crypto')
const {
MONGODB_URI,MONGODB_DB,CHAT_HISTORY_URI,CHAT_HISTORY_DB,
AZURE_CONNECTION_STRING,AZURE_CONTAINER_NAME,ADMIN_API_KEY,
KEY_CHECK_INTERVAL_MS,ASKDATA_ENDPOINT,ASKDATA_KEY,ASKDATA_MODEL,
ASKDATA_TIMEOUT_MS,ASKDATA2_ENDPOINT,ASKDATA2_KEY,ASKDATA2_MODEL,
ASKDATA2_TIMEOUT_MS,REQUEST_TIMEOUT_MS,WARMUP_CLIENT_IDS,
FAISS_PREFIX,BM25_PREFIX,CHUNK_CACHE_TTL,MAX_HITS_GLOBAL,
FAISS_TOP_K,BM25_TOP_K,RERANK_TOP_K,
SUPPORTED_EXTENSIONS,
responseCacheGet,responseCacheSet,getCacheKey,escapeRegex,capFirst,
ensureSinglePeriod,fixBrokenUrls,trimToLastCompleteSentence,
applySynonyms,applyTypos,normalizeQuery,validateQuery,
detectQueryIntent,detectMultiTopicQuery,extractSubject,
buildInvertedIndex,fetchWithTimeout,withRequestTimeout,
generateApiKey,generateTitle,selectFocusedHits,
}=require('./src/config')
const {
extractSpreadsheet,retrieveChunksDD,buildSystemPromptDD,buildUserMessageDD,
buildFallbackAnswerDD,fuzzyCorrectQuery,relaxedKeywordSearchDD,extractAllUrlsFromChunks,
}=require('./src/dd')
const {
chunkStructuredDocument,buildInvertedIndexSF,retrieveChunksSF,
buildSystemPromptSF,buildUserMessageSF,buildFallbackAnswerSF,
}=require('./src/sf')
const {
slidingWindowChunk,buildInvertedIndexUD,retrieveChunksUD,preprocessQueryUD,
buildSystemPromptUD,buildUserMessageUD,buildFallbackAnswerUD,cleanOcrNoise,
}=require('./src/ud')
const STRUCTURED_KEYWORDS=['column','datatype','data type','schema','table','mapping','field','kpi','etl','nullable','primary key','foreign key','varchar','integer','nvarchar','decimal','definition','alias','catalog','dimension','measure','attribute','grain','fact table','lookup','reference']
const STRUCTURED_FILE_TYPES=new Set(['csv','xlsx','xls','tsv'])
const UPPERCASE_FIELD_RE=/[A-Z_]{3,}/g
const ALIASES={
'Scenario Alias':['SCN_ALIAS_CD','scenario alias','scenario alt name','scn alias'],
'Scenario Code':['SCN_CD','scenario code','scenario identifier','scn cd'],
'Scenario Name':['SCN_NM','scenario name','scenario description','scn nm'],
'KPI':['key performance indicator','metric','measure','kpi code'],
'ETL':['extract transform load','data pipeline','etl mapping'],
'Primary Key':['PK','primary key column','unique identifier','id column'],
'Foreign Key':['FK','foreign key column','reference column'],
'Status':['STS_CD','status code','status flag','active flag','sts cd'],
'Description':['DESC','DESCR','description field','definition field'],
'Timestamp':['created_at','updated_at','modified_dt','create_dt'],
}
const STRUCTURED_QUERY_PATTERNS=[
/\bdefine\b/i,/\bdefinition\b/i,/\bcolumn\b/i,/\bfield\b/i,/\btable\b/i,
/\bschema\b/i,/\bmapping\b/i,/\bdatatype\b/i,/\bdata type\b/i,/\bkpi\b/i,
/\betl\b/i,/\bmetric\b/i,/\bmeasure\b/i,/\bdimension\b/i,/\bmeaning of\b/i,
/\balias\b/i,/\bcatalog\b/i,/what does .+ stand for/i,/what is the .+ column/i,/what is .+ field/i,
]
const UPPERCASE_CODE_PATTERN=/[A-Z][A-Z0-9_]{2,}[A-Z0-9]/
function detectDocumentType(text='',metadata={}){
const lowered=text.toLowerCase()
const fileType=(metadata.fileType||metadata.file_type||'').toLowerCase().replace(/^\./,'')
if(STRUCTURED_FILE_TYPES.has(fileType))return 'structured'
const docTypeHint=metadata.doc_type||metadata.docType||''
if(docTypeHint==='tabular'||docTypeHint==='structured')return 'structured'
const keywordHits=STRUCTURED_KEYWORDS.filter(k=>lowered.includes(k)).length
const uppercaseMatches=(text.match(UPPERCASE_FIELD_RE)||[]).length
const lines=text.split('\n').filter(l=>l.trim())
const pipeLines=lines.filter(l=>(l.match(/\|/g)||[]).length>=2).length
const tableDensity=lines.length>0?pipeLines/lines.length:0
let score=0
if(keywordHits>=3)score+=3
else if(keywordHits>=1)score+=1
if(uppercaseMatches>20)score+=3
else if(uppercaseMatches>8)score+=1
if(tableDensity>=0.4)score+=3
else if(tableDensity>=0.2)score+=1
return score>=3?'structured':'semantic'
}
function queryHitsAlias(query=''){
const qLower=query.toLowerCase()
for(const[key,values]of Object.entries(ALIASES)){
if(key.toLowerCase()===qLower)return true
for(const v of values){if(v.toLowerCase()===qLower||qLower.includes(v.toLowerCase()))return true}
}
return false
}
function expandQueryString(query=''){
const expanded=[query]
const qLower=query.toLowerCase()
for(const[key,values]of Object.entries(ALIASES)){
if(key.toLowerCase()===qLower){expanded.push(...values);continue}
for(const v of values){
if(v.toLowerCase()===qLower){expanded.push(key,...values.filter(x=>x!==v));break}
}
if(qLower.includes(key.toLowerCase()))expanded.push(...values)
for(const v of values){
if(qLower.includes(v.toLowerCase())){expanded.push(key,...values.filter(x=>x!==v));break}
}
}
const extras=[...new Set(expanded)].slice(1).join(' ')
return extras?`${query} ${extras}`:query
}
function classifyQuery(query=''){
if(!query.trim())return 'semantic'
if(UPPERCASE_CODE_PATTERN.test(query))return 'structured'
const qLower=query.toLowerCase()
for(const pattern of STRUCTURED_QUERY_PATTERNS){if(pattern.test(qLower))return 'structured'}
if(queryHitsAlias(query))return 'structured'
return 'semantic'
}
function calculateStructuredScore(doc,query,bm25Score=0,bm25Max=1){
const qLower=query.toLowerCase().trim()
const queryTokens=qLower.split(/\s+/).filter(w=>w.length>=2)
const normalizedBm25=bm25Max>0?(bm25Score/bm25Max)*0.5:0
let exactMatch=0
const colLower=(doc.column||doc.metadata?.measure||'').toLowerCase().trim()
if(colLower&&(colLower===qLower||queryTokens.some(t=>t===colLower)))exactMatch+=100
for(const token of queryTokens){
if(token.length>=3&&(doc.text||'').toLowerCase().includes(token))exactMatch+=3
}
let aliasMatch=0
const aliases=doc.aliases||doc.metadata?.aliases||[]
for(const alias of aliases){
if(alias.toLowerCase()===qLower)aliasMatch+=50
else if(queryTokens.some(t=>t.length>=3&&alias.toLowerCase().includes(t)))aliasMatch+=20
}
let tableMatch=0
const tableName=(doc.table||doc.metadata?.table||'').toLowerCase()
if(tableName){
for(const token of queryTokens){if(token.length>=3&&tableName.includes(token))tableMatch+=20}
if(tableName===qLower)tableMatch+=40
}
let columnMatch=0
if(colLower){
for(const token of queryTokens){
if(token.toUpperCase()===colLower.toUpperCase())columnMatch+=60
else if(token.length>=3&&colLower.includes(token))columnMatch+=15
}
}
return normalizedBm25+exactMatch+aliasMatch+tableMatch+columnMatch
}
function buildStructuredChunk(row,sourceFile='',chunkIndex=0){
const text=[
row.table?`Table: ${row.table}`:'',
row.column?`Column: ${row.column}`:'',
row.datatype?`Datatype: ${row.datatype}`:'',
row.definition?`Definition: ${row.definition}`:'',
row.aliases&&row.aliases.length>0?`Aliases: ${row.aliases.join(', ')}`:'',
].filter(Boolean).join('\n')
return{
docType:'structured',
table:row.table||'',
column:row.column||'',
datatype:row.datatype||'',
definition:row.definition||'',
aliases:row.aliases||[],
text,
source_file:sourceFile,
chunk_index:chunkIndex,
embedding:[],
metadata:{docType:'structured',table:row.table||'',column:row.column||'',datatype:row.datatype||'',definition:row.definition||'',aliases:row.aliases||[]},
}
}
function structuredSearchDD(query,chunks,topK=10){
const expandedQuery=expandQueryString(query)
const k1=1.5,b=0.75
const tokenized=chunks.map(c=>(c.text||'').toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/).filter(w=>w.length>1))
const avgdl=tokenized.reduce((s,t)=>s+t.length,0)/Math.max(tokenized.length,1)
const df=new Map()
for(const tokens of tokenized){const seen=new Set(tokens);for(const t of seen)df.set(t,(df.get(t)||0)+1)}
const N=tokenized.length
const qTokens=expandedQuery.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/).filter(w=>w.length>1)
const bm25Scores=tokenized.map(tokens=>{
const tf=new Map()
for(const t of tokens)tf.set(t,(tf.get(t)||0)+1)
let score=0
for(const q of qTokens){
const idf=Math.log((N-(df.get(q)||0)+0.5)/((df.get(q)||0)+0.5)+1)
const tfVal=tf.get(q)||0
const numerator=tfVal*(k1+1)
const denominator=tfVal+k1*(1-b+b*tokens.length/avgdl)
score+=idf*(numerator/denominator)
}
return score
})
const bm25Max=Math.max(...bm25Scores,1)
const scored=chunks.map((chunk,i)=>({...chunk,_score:calculateStructuredScore(chunk,expandedQuery,bm25Scores[i]||0,bm25Max)})).filter(c=>c._score>0)
scored.sort((a,b)=>b._score-a._score)
return scored.slice(0,topK)
}
const app=express()
const allowedOrigins=[
'http://localhost:8080','http://localhost:3000',
'https://app.powerbi.com','https://msit.powerbi.com',
'https://anuritchat.vercel.app','https://ragadminpanel.vercel.app',
'https://df.powerbi.com','https://www.anuritinnovation.com/',
'https://api.powerbi.com',
]
function originAllowed(origin){
if(!origin)return true
if(origin==='null')return true
if(allowedOrigins.includes(origin))return true
if(/\.(powerbi|microsoft|office)\.com$/.test(origin))return true
return false
}
app.use(cors({
origin:(origin,callback)=>callback(null,originAllowed(origin)),
methods:['GET','POST','PUT','PATCH','DELETE','OPTIONS'],
allowedHeaders:['Content-Type','Authorization','x-session-id'],
credentials:true,
}))
app.options('*',cors({
origin:(origin,callback)=>callback(null,true),
methods:['GET','POST','PUT','PATCH','DELETE','OPTIONS'],
allowedHeaders:['Content-Type','Authorization','x-session-id'],
credentials:true,
}))
app.use(express.json())
const blobServiceClient=AZURE_CONNECTION_STRING?BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING):null
let askedataActiveCount=0
const ASKDATA_MAX_CONCURRENT=3
const askedataQueue=[]
function runWithAskedataLimit(fn){
return new Promise((resolve,reject)=>{
function tryRun(){
if(askedataActiveCount<ASKDATA_MAX_CONCURRENT){
askedataActiveCount++
Promise.resolve().then(fn).then(
result=>{askedataActiveCount--;drainAskedataQueue();resolve(result)},
err=>{askedataActiveCount--;drainAskedataQueue();reject(err)}
)
}else{
askedataQueue.push(tryRun)
}
}
tryRun()
})
}
function drainAskedataQueue(){
if(askedataQueue.length>0&&askedataActiveCount<ASKDATA_MAX_CONCURRENT)askedataQueue.shift()()
}
let askedataFailures=0
let askedataBlockedUntil=0
function askedataCircuitOpen(){
if(Date.now()<askedataBlockedUntil)return true
if(askedataBlockedUntil>0){askedataBlockedUntil=0;askedataFailures=0}
return false
}
function askedataRecordSuccess(){askedataFailures=0;askedataBlockedUntil=0}
function askedataRecordFailure(){
askedataFailures++
if(askedataFailures>=3){
askedataBlockedUntil=Date.now()+30000
console.error('[ASKDATA] Circuit breaker OPEN for 30s')
}
}
async function callASKDATA(systemPrompt,userMessage,maxTokens=1024){
if(!ASKDATA_ENDPOINT||!ASKDATA_KEY)throw new Error('ASKDATA not configured')
if(askedataCircuitOpen())throw new Error('ASKDATA temporarily unavailable')
return runWithAskedataLimit(async()=>{
try{
const response=await fetchWithTimeout(
ASKDATA_ENDPOINT,
{
method:'POST',
headers:{'Content-Type':'application/json','Authorization':`Bearer ${ASKDATA_KEY}`},
body:JSON.stringify({model:ASKDATA_MODEL,messages:[{role:'system',content:systemPrompt},{role:'user',content:userMessage}],temperature:0.1,max_tokens:maxTokens}),
},
ASKDATA_TIMEOUT_MS
)
if(!response.ok){
const errText=await response.text()
throw new Error(`ASKDATA error ${response.status}: ${errText}`)
}
const data=await response.json()
askedataRecordSuccess()
return data.choices?.[0]?.message?.content||''
}catch(err){
askedataRecordFailure()
throw err
}
})
}
async function callASKDATA2(systemPrompt,userMessage,maxTokens=1024){
if(!ASKDATA2_ENDPOINT||!ASKDATA2_KEY)throw new Error('ASKDATA2 not configured')
try{
const response=await fetchWithTimeout(
ASKDATA2_ENDPOINT,
{
method:'POST',
headers:{'Content-Type':'application/json','Authorization':`Bearer ${ASKDATA2_KEY}`,'Accept':'application/json'},
body:JSON.stringify({model:ASKDATA2_MODEL,messages:[{role:'system',content:systemPrompt},{role:'user',content:userMessage}],max_tokens:maxTokens,temperature:0.1,top_p:1.0,stream:false}),
},
ASKDATA2_TIMEOUT_MS
)
if(!response.ok){
const errText=await response.text()
throw new Error(`ASKDATA2 error ${response.status}: ${errText}`)
}
const data=await response.json()
return data.choices?.[0]?.message?.content||''
}catch(err){
console.error(`[ASKDATA2] Failed: ${err.message}`)
throw err
}
}
async function callBestAvailableEngine(systemPrompt,userMessage,maxTokens=1024){
let primaryError=null
if(ASKDATA_ENDPOINT&&ASKDATA_KEY&&!askedataCircuitOpen()){
try{
const result=await callASKDATA(systemPrompt,userMessage,maxTokens)
if(result&&result.trim().length>=15)return result
primaryError=new Error('ASKDATA returned blank response')
}catch(err){
primaryError=err
console.warn(`[ASKDATA] Failed, switching to ASKDATA2: ${err.message}`)
}
}else{
primaryError=new Error('ASKDATA unavailable')
}
if(ASKDATA2_ENDPOINT&&ASKDATA2_KEY){
try{
const result=await callASKDATA2(systemPrompt,userMessage,maxTokens)
if(result&&result.trim().length>=15)return result
}catch(err){
console.error(`[ASKDATA2] Also failed: ${err.message}`)
}
}
return ''
}
async function downloadBlobAsBuffer(containerClient,blobName){
const download=await containerClient.getBlobClient(blobName).download()
const parts=[]
for await(const chunk of download.readableStreamBody){
parts.push(Buffer.isBuffer(chunk)?chunk:Buffer.from(chunk))
}
return Buffer.concat(parts)
}
function resolveDocType(c){
const meta=c.metadata||{}
const section=c.section||meta.section||''
const measure=c.measure||meta.measure||''
if(section)return 'structured'
if(measure)return 'data_dictionary'
return 'unstructured'
}
async function loadChunksFromBlob(clientId){
if(!blobServiceClient)throw new Error('AZURE_CONNECTION_STRING not set')
const containerClient=blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
const chunksBlob=`${FAISS_PREFIX}/${clientId}/chunks.json`
let chunks=[]
try{
const buf=await downloadBlobAsBuffer(containerClient,chunksBlob)
const raw=JSON.parse(buf.toString('utf-8'))
chunks=raw.map((c,i)=>({
text:c.text||'',
source_file:c.source_file||'unknown',
doc_id:c.doc_id||c.metadata?.doc_id||'unknown',
chunk_index:i,
embedding:c.embedding||null,
metadata:c.metadata||null,
section:c.section||(c.metadata&&c.metadata.section)||'',
measure:c.measure||(c.metadata&&c.metadata.measure)||'',
docType:resolveDocType(c),
}))
console.log(`[blobLoad] Loaded ${chunks.length} chunks for ${clientId} from ${chunksBlob}`)
const byDoc={}
for(const c of chunks){byDoc[c.doc_id]=(byDoc[c.doc_id]||0)+1}
console.log(`[blobLoad] doc distribution: ${JSON.stringify(byDoc)}`)
}catch(err){
console.warn(`[blobLoad] chunks.json not found for ${clientId}: ${err.message}`)
}
return chunks
}
async function loadBM25FromBlob(clientId){
if(!blobServiceClient)return null
const containerClient=blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
const bm25Blob=`${BM25_PREFIX}/${clientId}/bm25.pkl`
try{
await containerClient.getBlobClient(bm25Blob).getProperties()
console.log(`[blobLoad] BM25 blob exists for ${clientId} (Python-side scoring used)`)
return{clientId,blobName:bm25Blob,available:true}
}catch{
console.warn(`[blobLoad] BM25 blob not found for ${clientId}`)
return null
}
}
function bm25ScoreJS(query,chunks){
const k1=1.5,b=0.75
const queryTokens=query.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/).filter(w=>w.length>1)
if(queryTokens.length===0)return new Array(chunks.length).fill(0)
const tokenized=chunks.map(c=>(c.text||'').toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/).filter(w=>w.length>1))
const avgdl=tokenized.reduce((s,t)=>s+t.length,0)/Math.max(tokenized.length,1)
const df=new Map()
for(const tokens of tokenized){
const seen=new Set(tokens)
for(const t of seen)df.set(t,(df.get(t)||0)+1)
}
const N=tokenized.length
return tokenized.map(tokens=>{
const tf=new Map()
for(const t of tokens)tf.set(t,(tf.get(t)||0)+1)
let score=0
for(const q of queryTokens){
const idf=Math.log((N-(df.get(q)||0)+0.5)/((df.get(q)||0)+0.5)+1)
const tfVal=tf.get(q)||0
const numerator=tfVal*(k1+1)
const denominator=tfVal+k1*(1-b+b*tokens.length/avgdl)
score+=idf*(numerator/denominator)
}
return score
})
}
function hybridSearch(query,chunks,faissTopK=FAISS_TOP_K,bm25TopK=BM25_TOP_K,alpha=0.6,sourceFilter=null){
const pool=sourceFilter?chunks.filter(c=>c.source_file===sourceFilter):chunks
if(pool.length===0)return[]
const invertedIndex=buildInvertedIndex(pool)
const queryWords=normalizeQuery(query).replace(/[^\w\s]/g,' ').split(/\s+/).filter(w=>w.length>1)
const subject=extractSubject(query)
const subjectWords=subject.toLowerCase().split(/\s+/).filter(w=>w.length>1)
const candidateSet=new Set()
for(const w of[...queryWords,...subjectWords]){
for(const idx of(invertedIndex.get(w)||new Set()))candidateSet.add(idx)
}
const source=candidateSet.size>0
?[...candidateSet].map(i=>pool[i]).filter(Boolean)
:pool.slice(0,Math.min(300,pool.length))
const subjectPhraseStr=subject.toLowerCase()
const faissScores=source.map(c=>{
const text=(c.text||'').toLowerCase()
let score=0
const wordCoverage=queryWords.filter(w=>text.includes(w)).length
score+=wordCoverage*3
const subjectCoverage=subjectWords.filter(w=>new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(text)).length
score+=subjectCoverage*5
if(text.includes(subjectPhraseStr))score+=8
if(c.metadata?.section||c.section){
const sl=(c.section||c.metadata?.section||'').toLowerCase()
const sm=subjectWords.filter(w=>sl.includes(w)).length
score+=sm*8
if(sl.includes(subjectPhraseStr))score+=12
}
if(c.metadata?.measure||c.measure){
const ml=(c.measure||c.metadata?.measure||'').toLowerCase()
if(ml===subjectPhraseStr)score+=100
}
return{chunk:c,faissScore:score}
})
const faissMax=Math.max(...faissScores.map(x=>x.faissScore),1)
const bm25Raw=bm25ScoreJS(query,source)
const bm25Max=Math.max(...bm25Raw,1)
const hybrid=faissScores.map((x,i)=>{
const fNorm=x.faissScore/faissMax
const bNorm=bm25Raw[i]/bm25Max
return{chunk:x.chunk,score:alpha*fNorm+(1-alpha)*bNorm}
})
hybrid.sort((a,b)=>b.score-a.score)
const SCORE_THRESHOLD=0.08
const seen=new Set()
const results=[]
for(const{chunk,score}of hybrid){
if(score<SCORE_THRESHOLD)break
const fp=(chunk.text||'').trim().slice(0,60).toLowerCase()
if(!seen.has(fp)){
seen.add(fp)
results.push({...chunk,_score:score})
}
if(results.length>=Math.max(faissTopK,bm25TopK))break
}
return results
}
function classifyDocumentType(chunks,fileName){
if(!chunks||chunks.length===0)return 'unstructured'
const ddCount=chunks.filter(c=>c.docType==='data_dictionary').length
if(ddCount/Math.max(chunks.length,1)>0.4)return 'data_dictionary'
const sfCount=chunks.filter(c=>c.docType==='structured').length
if(sfCount/Math.max(chunks.length,1)>0.3)return 'structured'
return 'unstructured'
}
function detectDocTypeFromChunks(chunks){
if(!chunks||chunks.length===0)return 'unstructured'
const ddCount=chunks.filter(c=>c.docType==='data_dictionary').length
if(ddCount/Math.max(chunks.length,1)>0.4)return 'data_dictionary'
const sfCount=chunks.filter(c=>c.docType==='structured').length
if(sfCount/Math.max(chunks.length,1)>0.3)return 'structured'
return 'unstructured'
}
async function _doLoadChunks(clientId){
const chunks=await loadChunksFromBlob(clientId)
if(chunks.length===0){
console.warn(`[loadChunks] No chunks found in blob for ${clientId}`)
return[]
}
console.log(`[loadChunks] ${clientId} -> ${chunks.length} total chunks`)
return chunks
}
const CHUNK_CACHE=new Map()
async function loadChunksForClient(clientId){
const now=Date.now()
const cached=CHUNK_CACHE.get(clientId)
if(cached&&cached.chunks){
if(now-cached.ts<=CHUNK_CACHE_TTL)return{chunks:cached.chunks,invertedIndexes:cached.invertedIndexes}
if(!cached.loading){
const refreshPromise=_doLoadChunks(clientId)
.then(chunks=>{
const invertedIndexes=buildAllInvertedIndexes(chunks)
CHUNK_CACHE.set(clientId,{chunks,invertedIndexes,ts:Date.now(),loading:null})
console.log(`[chunkCache] Background refresh done for ${clientId}: ${chunks.length} chunks`)
})
.catch(err=>{
const existing=CHUNK_CACHE.get(clientId)
if(existing)CHUNK_CACHE.set(clientId,{...existing,loading:null})
console.warn(`[chunkCache] Background refresh failed for ${clientId}: ${err.message}`)
})
CHUNK_CACHE.set(clientId,{...cached,loading:refreshPromise})
}
return{chunks:cached.chunks,invertedIndexes:cached.invertedIndexes}
}
if(cached&&cached.loading){
await cached.loading
const entry=CHUNK_CACHE.get(clientId)
return{chunks:entry?.chunks||[],invertedIndexes:entry?.invertedIndexes||{}}
}
const loadPromise=_doLoadChunks(clientId)
.then(chunks=>{
const invertedIndexes=buildAllInvertedIndexes(chunks)
CHUNK_CACHE.set(clientId,{chunks,invertedIndexes,ts:Date.now(),loading:null})
return chunks
})
.catch(err=>{
CHUNK_CACHE.set(clientId,{chunks:null,invertedIndexes:{},ts:0,loading:null})
throw err
})
CHUNK_CACHE.set(clientId,{chunks:null,invertedIndexes:{},ts:0,loading:loadPromise})
await loadPromise
const entry=CHUNK_CACHE.get(clientId)
return{chunks:entry?.chunks||[],invertedIndexes:entry?.invertedIndexes||{}}
}
function buildAllInvertedIndexes(chunks){
const ddChunks=chunks.filter(c=>c.docType==='data_dictionary')
const sfChunks=chunks.filter(c=>c.docType==='structured')
const udChunks=chunks.filter(c=>c.docType!=='data_dictionary'&&c.docType!=='structured')
const structuredChunks=[...ddChunks,...sfChunks]
const semanticChunks=udChunks
return{
dd:ddChunks.length>0?buildInvertedIndex(ddChunks):buildInvertedIndex(chunks),
sf:sfChunks.length>0?buildInvertedIndexSF(sfChunks):null,
ud:udChunks.length>0?buildInvertedIndexUD(udChunks):null,
all:buildInvertedIndex(chunks),
structured:structuredChunks.length>0?buildInvertedIndex(structuredChunks):null,
semantic:semanticChunks.length>0?buildInvertedIndex(semanticChunks):null,
}
}
function invalidateChunkCache(clientId){
CHUNK_CACHE.delete(clientId)
console.log(`[chunkCache] Invalidated cache for client: ${clientId}`)
}
function warmupChunkCaches(){
if(!WARMUP_CLIENT_IDS.length||!blobServiceClient)return
console.log(`[warmup] Pre-loading chunks for: ${WARMUP_CLIENT_IDS.join(', ')}`)
for(const clientId of WARMUP_CLIENT_IDS){
loadChunksForClient(clientId)
.then(({chunks})=>console.log(`[warmup] ${clientId} -- ${chunks.length} chunks ready`))
.catch(err=>console.warn(`[warmup] ${clientId} -- ${err.message}`))
}
}
let db=null
async function getDb(){
if(db)return db
const client=new MongoClient(MONGODB_URI)
await client.connect()
db=client.db(MONGODB_DB)
await db.collection('clients').createIndex({apiKey:1},{unique:true,sparse:true})
return db
}
let chatDb=null
async function getChatDb(){
if(chatDb)return chatDb
const uri=CHAT_HISTORY_URI||MONGODB_URI
const client=new MongoClient(uri)
await client.connect()
chatDb=client.db(CHAT_HISTORY_DB)
return chatDb
}
const CLIENT_CACHE=new Map()
const CACHE_TTL_MS=5*60*1000
function getCached(apiKey){
const entry=CLIENT_CACHE.get(apiKey)
if(!entry)return null
if(Date.now()-entry.cachedAt>CACHE_TTL_MS){CLIENT_CACHE.delete(apiKey);return null}
return entry
}
function setCache(apiKey,data){CLIENT_CACHE.set(apiKey,{...data,cachedAt:Date.now()})}
function evictCache(apiKey){if(apiKey)CLIENT_CACHE.delete(apiKey)}
async function verifyApiKey(apiKey){
if(!apiKey||!apiKey.startsWith('rak_'))return null
const cached=getCached(apiKey)
if(cached)return{clientId:cached.clientId,name:cached.name}
const database=await getDb()
const client=await database.collection('clients').findOne({apiKey},{projection:{clientId:1,name:1,_id:0}})
if(!client)return null
setCache(apiKey,{clientId:client.clientId,name:client.name})
return{clientId:client.clientId,name:client.name}
}
function startApiKeyHealthChecker(){
if(!MONGODB_URI)return
setInterval(async()=>{
const keys=[...CLIENT_CACHE.keys()]
if(!keys.length)return
try{
const database=await getDb()
const validDocs=await database.collection('clients').find({apiKey:{$in:keys}},{projection:{apiKey:1,_id:0}}).toArray()
const validSet=new Set(validDocs.map(d=>d.apiKey))
for(const key of keys)if(!validSet.has(key))evictCache(key)
}catch{}
},KEY_CHECK_INTERVAL_MS)
}
function extractApiKey(req){
const header=req.headers['authorization']||''
return header.startsWith('Bearer ')?header.slice(7).trim():null
}
async function requireClientKey(req,res,next){
const apiKey=extractApiKey(req)||req.body?.apiKey
if(!apiKey)return res.status(401).json({error:'Missing API key'})
const client=await verifyApiKey(apiKey)
if(!client)return res.status(401).json({error:'Invalid or expired API key'})
req.client=client
next()
}
function requireAdminKey(req,res,next){
const key=extractApiKey(req)
if(!key||key!==ADMIN_API_KEY)return res.status(401).json({error:'Unauthorized'})
next()
}
function stripMetadata(text){
if(!text)return''
return text
.replace(/\[File:[^\]]*\]\s*/g,'')
.replace(/\[Sheet:[^\]]*\]\s*/g,'')
.replace(/^\s*(Table Name|Measure Name|Attribute Name|Description|Source File|Sheet|Connected Fact Table)[^:\n]*:[^\n]*/gim,'')
.replace(/^[^\n]*(Table Name|Measure Name|Attribute Name)\s*:[^\n]*/gim,'')
.replace(/\[File:[^\]]*\]\s*(?:Table Name|Measure Name|Attribute Name)[^\n]*/gi,'')
.replace(/^\s*Connected Fact Table[^\n]*/gim,'')
.replace(/\s*\|\s*(Table Name|Measure Name|Attribute Name|Description|Source File|Connected Fact Table)[^|]*\|?[^\n]*/gi,'')
.replace(/^\s*›\s*\d+\s+document\s+section[^\n]*/gim,'')
.replace(/^\s*\[Source\s*\d+\]\s*/gm,'')
.replace(/^[^\n]*(\|[^\n]*){3,}$/gm,'')
.replace(/=== .+ ===\s*/gm,'')
.replace(/\(from\s+[A-Za-z\s]+\)\s*/g,'')
.replace(/\n{3,}/g,'\n\n')
.trim()
}
function cleanAnswer(rawAnswer){
if(!rawAnswer)return''
let cleaned=stripMetadata(fixBrokenUrls(rawAnswer))
.replace(/\.{2,}/g,'.')
.replace(/\.\s*\./g,'.')
.trim()
cleaned=trimToLastCompleteSentence(cleaned)
if(cleaned.length>0&&!/[.!?]$/.test(cleaned))cleaned+='.'
return ensureSinglePeriod(cleaned)
}
async function generateAnswer(query,hits,intent,docType){
let systemPrompt,userMessage
if(docType==='data_dictionary'){
systemPrompt=buildSystemPromptDD(intent)
userMessage=buildUserMessageDD(query,hits,intent)
}else if(docType==='structured'){
systemPrompt=buildSystemPromptSF(intent)
userMessage=buildUserMessageSF(query,hits,intent)
}else{
systemPrompt=buildSystemPromptUD(intent)
userMessage=buildUserMessageUD(query,hits,intent)
}
return callBestAvailableEngine(systemPrompt,userMessage,1024)
}
function buildFallbackAnswer(query,hits,intent,docType){
if(docType==='data_dictionary')return buildFallbackAnswerDD(query,hits,intent)
if(docType==='structured')return buildFallbackAnswerSF(query,hits)
return buildFallbackAnswerUD(query,hits)
}
function computePoolConfidence(hits,docType){
if(!hits||hits.length===0)return 0
const topScore=hits[0]?._score||0
const secondScore=hits[1]?._score||0
const gap=topScore-secondScore
const MIN_THRESHOLD=0.08
if(topScore<MIN_THRESHOLD)return 0
const topN=hits.slice(0,Math.min(5,hits.length))
const avgScore=topN.reduce((s,h)=>s+(h._score||0),0)/topN.length
let confidence=avgScore+gap*0.5
if(docType==='data_dictionary')confidence*=1.4
return confidence
}
function pickBestSourceFile(query,chunks){
if(!chunks||chunks.length===0)return null
const files=[...new Set(chunks.map(c=>c.source_file).filter(Boolean))]
if(files.length<=1)return null
const queryWords=normalizeQuery(query).replace(/[^\w\s]/g,' ').split(/\s+/).filter(w=>w.length>1)
const fileScores=files.map(f=>{
const fileChunks=chunks.filter(c=>c.source_file===f)
const bm25Raw=bm25ScoreJS(query,fileChunks)
const bm25Max=Math.max(...bm25Raw,1)
const bm25Avg=bm25Raw.slice().sort((a,b)=>b-a).slice(0,5).reduce((s,v)=>s+v,0)/5
const bm25Norm=bm25Avg/bm25Max
const subjectHits=fileChunks.filter(c=>{
const t=(c.text||'').toLowerCase()
return queryWords.filter(w=>t.includes(w)).length>=Math.ceil(queryWords.length*0.5)
}).length
const coverage=subjectHits/Math.max(fileChunks.length,1)
return{file:f,score:bm25Norm*0.6+coverage*0.4}
}).sort((a,b)=>b.score-a.score)
const best=fileScores[0]
const runner=fileScores[1]
if(runner&&best.score<runner.score*1.5)return null
console.log(`[docRouter] Routed to: ${best.file} (score=${best.score.toFixed(3)})`)
return best.file
}
async function retrieveBestHitsAcrossAllTypes(processedQuery,chunks,topK,invertedIndexes,intent,sourceFilter=null){
const pool=sourceFilter?chunks.filter(c=>c.source_file===sourceFilter):chunks
const ddChunks=pool.filter(c=>c.docType==='data_dictionary')
const sfChunks=pool.filter(c=>c.docType==='structured')
const udChunks=pool.filter(c=>c.docType!=='data_dictionary'&&c.docType!=='structured')
const structuredPool=[...ddChunks,...sfChunks]
const semanticPool=udChunks
const poolInvertedIndexes=sourceFilter?buildAllInvertedIndexes(pool):invertedIndexes
const queryType=classifyQuery(processedQuery)
console.log(`[adaptiveRoute] queryType=${queryType} | structured=${structuredPool.length} | semantic=${semanticPool.length}`)
if(queryType==='structured'&&structuredPool.length>0){
const ddHits=ddChunks.length>0?structuredSearchDD(processedQuery,ddChunks,topK):[]
const sfKeyword=sfChunks.length>0?retrieveChunksSF(processedQuery,sfChunks,topK,poolInvertedIndexes.sf||poolInvertedIndexes.all):[]
const seenS=new Set()
const mergedS=[]
for(const h of[...ddHits,...sfKeyword]){
const fp=(h.text||'').trim().slice(0,60).toLowerCase()
if(!seenS.has(fp)){seenS.add(fp);mergedS.push(h)}
if(mergedS.length>=topK)break
}
if(mergedS.length>0){
const docType=ddHits.length>=sfKeyword.length?'data_dictionary':'structured'
console.log(`[adaptiveRoute] structured retrieval -> ${mergedS.length} hits, docType=${docType}`)
return{hits:mergedS,docType}
}
console.log(`[adaptiveRoute] structured retrieval empty, falling back to semantic`)
}
const hybridHits=hybridSearch(processedQuery,semanticPool.length>0?semanticPool:pool,FAISS_TOP_K,BM25_TOP_K,0.6,null)
const udKeyword=semanticPool.length>0?retrieveChunksUD(processedQuery,semanticPool,topK,poolInvertedIndexes.ud||poolInvertedIndexes.all):[]
const seenU=new Set()
const mergedU=[]
for(const h of[...hybridHits,...udKeyword]){
const fp=(h.text||'').trim().slice(0,60).toLowerCase()
if(!seenU.has(fp)){seenU.add(fp);mergedU.push(h)}
if(mergedU.length>=topK)break
}
const docType=detectDocTypeFromChunks(mergedU.length>0?mergedU:pool)
console.log(`[adaptiveRoute] semantic retrieval -> ${mergedU.length} hits, docType=${docType}`)
return{hits:mergedU,docType}
}
async function generateAnswerForTopic(topic,chunks,topK,invertedIndexes){
const topicQuery=`what is ${topic}`
const docType=detectDocTypeFromChunks(chunks)
let hits=await retrieveBestHitsAcrossAllTypes(topicQuery,chunks,topK,invertedIndexes,'definition')
if(hits.hits)hits=hits.hits
if(hits.length===0)hits=relaxedKeywordSearchDD(topicQuery,chunks,32,invertedIndexes.all)
if(hits.length===0)return null
let rawAnswer=''
try{
rawAnswer=await Promise.race([
generateAnswer(topicQuery,hits,'definition',docType),
new Promise((_,reject)=>setTimeout(()=>reject(new Error('timeout')),25000)),
])
}catch(err){
console.warn(`[generateAnswerForTopic] Failed for "${topic}": ${err.message}`)
}
const isBlank=!rawAnswer||rawAnswer.trim().length<15
const rawFinal=isBlank?buildFallbackAnswer(topicQuery,hits,'definition',docType):rawAnswer
let answer=cleanAnswer(rawFinal)
answer=answer.replace(/^\*\*[^*]+\*\*\s*(is defined as:?\s*)?/i,'').trim()
if(answer&&!/[.!?]$/.test(answer))answer+='.'
return answer
}
async function generateComparisonAnswer(topicA,topicB,chunks,topK,invertedIndexes){
const comparisonQuery=`difference between ${topicA} and ${topicB}`
const docType=detectDocTypeFromChunks(chunks)
const resA=await retrieveBestHitsAcrossAllTypes(`what is ${topicA}`,chunks,topK,invertedIndexes,'definition')
const resB=await retrieveBestHitsAcrossAllTypes(`what is ${topicB}`,chunks,topK,invertedIndexes,'definition')
const hitsA=resA.hits||[]
const hitsB=resB.hits||[]
const allHits=[...hitsA,...hitsB]
const seen=new Set()
const deduped=[]
for(const h of allHits){
const fp=(h.text||'').trim().slice(0,80).toLowerCase()
if(!seen.has(fp)){seen.add(fp);deduped.push(h)}
}
if(deduped.length===0)return null
let rawAnswer=''
try{
rawAnswer=await Promise.race([
generateAnswer(comparisonQuery,deduped,'comparison',docType),
new Promise((_,reject)=>setTimeout(()=>reject(new Error('timeout')),25000)),
])
}catch(err){
console.warn(`[generateComparisonAnswer] Failed: ${err.message}`)
}
if(rawAnswer&&rawAnswer.trim().length>=15)return cleanAnswer(rawAnswer)
const answerA=await generateAnswerForTopic(topicA,chunks,topK,invertedIndexes)
const answerB=await generateAnswerForTopic(topicB,chunks,topK,invertedIndexes)
const parts=[]
if(answerA&&!answerA.includes('could not find'))parts.push(`**${capFirst(topicA)}:** ${answerA}`)
else parts.push(`**${capFirst(topicA)}:** I could not find information about "${capFirst(topicA)}" in your documents.`)
if(answerB&&!answerB.includes('could not find'))parts.push(`**${capFirst(topicB)}:** ${answerB}`)
else parts.push(`**${capFirst(topicB)}:** I could not find information about "${capFirst(topicB)}" in your documents.`)
return parts.join('\n\n')
}
async function handleMultiTopicQuery(topics,mode,chunks,topK,invertedIndexes){
const results=await Promise.all(topics.map(async(topic)=>{
const answer=await generateAnswerForTopic(topic,chunks,topK,invertedIndexes)
return{topic,answer}
}))
const parts=results.map(({topic,answer})=>{
const cap=capFirst(topic)
if(!answer||answer.includes('could not find')||answer.includes('not present')){
return`**${cap}:**\nI could not find information about "${cap}" in your documents.`
}
return`**${cap}:**\n${answer}`
})
if(mode==='comparison'&&results.length===2){
const[a,b]=results
const bothFound=a.answer&&!a.answer.includes('could not find')&&b.answer&&!b.answer.includes('could not find')
if(bothFound){
const comparisonAnswer=await generateComparisonAnswer(a.topic,b.topic,chunks,topK,invertedIndexes)
if(comparisonAnswer&&!comparisonAnswer.includes('could not find'))return comparisonAnswer
}
return parts.join('\n\n')
}
return parts.join('\n\n')
}
async function saveConversationMessage(clientId,conversationId,query,answer,sources){
try{
const chatDatabase=await getChatDb()
const col=chatDatabase.collection('conversations')
const now=new Date()
const userMsg={role:'user',content:query,timestamp:now}
const assistantMsg={role:'assistant',content:answer,sources:sources.map(s=>({source_file:s.source_file,score:s.score})),timestamp:now}
let activeConversationId=conversationId||null
if(activeConversationId){
const updated=await col.findOneAndUpdate(
{_id:new ObjectId(activeConversationId),clientId},
{$push:{messages:{$each:[userMsg,assistantMsg]}},$set:{updatedAt:now}},
{returnDocument:'after',projection:{_id:1}}
)
if(!updated)activeConversationId=null
}
if(!activeConversationId){
const result=await col.insertOne({clientId,title:generateTitle(query),messages:[userMsg,assistantMsg],createdAt:now,updatedAt:now})
activeConversationId=result.insertedId.toString()
}
return activeConversationId
}catch(saveErr){
console.warn('[saveConversationMessage] Failed:',saveErr.message)
return conversationId||null
}
}
const IN_FLIGHT=new Map()
app.get('/health',(req,res)=>res.json({
ok:true,service:'ask-data',
engines:{primary:ASKDATA_ENDPOINT?'configured':'missing',fallback:ASKDATA2_ENDPOINT?'configured':'missing'},
retrieval:'adaptive dual-retrieval: structured BM25+alias+exact | semantic hybrid+rerank',
chunkCacheSize:CHUNK_CACHE.size,
primaryCircuitOpen:askedataCircuitOpen(),
}))
app.post('/client/verify',async(req,res)=>{
try{
const apiKey=extractApiKey(req)||req.body?.apiKey
if(!apiKey)return res.status(400).json({valid:false,error:'apiKey is required'})
const client=await verifyApiKey(apiKey)
if(!client)return res.status(401).json({valid:false,error:'Invalid or expired API key'})
res.json({valid:true,client})
}catch(err){res.status(500).json({valid:false,error:err.message})}
})
app.post('/admin/clients',requireAdminKey,async(req,res)=>{
try{
let{name,clientId,apiKey}=req.body
if(!name||!clientId)return res.status(400).json({error:'name and clientId are required'})
if(!apiKey){apiKey=generateApiKey()}
else if(!apiKey.startsWith('rak_'))return res.status(400).json({error:'apiKey must start with "rak_"'})
const database=await getDb()
const col=database.collection('clients')
const existing=await col.findOne({$or:[{clientId},{apiKey}]})
if(existing){
const field=existing.clientId===clientId?'clientId':'apiKey'
return res.status(409).json({error:`A client with this ${field} already exists`})
}
const now=new Date().toISOString()
const doc={name:name.trim(),clientId:clientId.trim().toLowerCase(),apiKey,apiKeyRotatedAt:now,folderLink:'',sourceType:'google-drive',status:'idle',documentsCount:0,autoSync:false,watchIntervalMs:300000,lastRunAt:null,lastError:null,createdAt:now,updatedAt:now}
const result=await col.insertOne(doc)
res.status(201).json({...doc,_id:result.insertedId})
}catch(err){res.status(500).json({error:err.message})}
})
app.get('/admin/clients',requireAdminKey,async(req,res)=>{
try{
const database=await getDb()
const clients=await database.collection('clients').find({},{projection:{apiKey:0}}).sort({createdAt:-1}).toArray()
res.json({clients})
}catch(err){res.status(500).json({error:err.message})}
})
app.get('/admin/clients/:clientId',requireAdminKey,async(req,res)=>{
try{
const database=await getDb()
const client=await database.collection('clients').findOne({clientId:req.params.clientId})
if(!client)return res.status(404).json({error:'Client not found'})
res.json(client)
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/admin/clients/:clientId/regenerate-key',requireAdminKey,async(req,res)=>{
try{
const database=await getDb()
const col=database.collection('clients')
const oldClient=await col.findOne({clientId:req.params.clientId},{projection:{apiKey:1}})
if(!oldClient)return res.status(404).json({error:'Client not found'})
const newApiKey=generateApiKey()
const now=new Date().toISOString()
if(oldClient.apiKey)evictCache(oldClient.apiKey)
await col.findOneAndUpdate({clientId:req.params.clientId},{$set:{apiKey:newApiKey,apiKeyRotatedAt:now,updatedAt:now}},{returnDocument:'after'})
res.json({success:true,clientId:req.params.clientId,newApiKey,apiKeyRotatedAt:now})
}catch(err){res.status(500).json({error:err.message})}
})
app.patch('/admin/clients/:clientId',requireAdminKey,async(req,res)=>{
try{
const database=await getDb()
const updates={...req.body,updatedAt:new Date().toISOString()}
if(updates.apiKey!==undefined){
if(!updates.apiKey.startsWith('rak_'))return res.status(400).json({error:'apiKey must start with "rak_"'})
const old=await database.collection('clients').findOne({clientId:req.params.clientId},{projection:{apiKey:1}})
if(old?.apiKey)evictCache(old.apiKey)
updates.apiKeyRotatedAt=new Date().toISOString()
}
const result=await database.collection('clients').findOneAndUpdate({clientId:req.params.clientId},{$set:updates},{returnDocument:'after'})
if(!result)return res.status(404).json({error:'Client not found'})
res.json(result)
}catch(err){res.status(500).json({error:err.message})}
})
app.delete('/admin/clients/:clientId',requireAdminKey,async(req,res)=>{
try{
const{clientId}=req.params
const database=await getDb()
const client=await database.collection('clients').findOne({clientId})
if(!client)return res.status(404).json({error:'Client not found'})
if(client.apiKey)evictCache(client.apiKey)
await database.collection('clients').deleteOne({clientId})
invalidateChunkCache(clientId)
const blobsDeleted=[],blobsFailed=[]
if(blobServiceClient){
try{
const containerClient=blobServiceClient.getContainerClient(AZURE_CONTAINER_NAME)
for(const prefix of[`raw/${clientId}/`,`faiss/${clientId}/`,`bm25/${clientId}/`,`meta/${clientId}/`]){
for await(const blob of containerClient.listBlobsFlat({prefix})){
try{await containerClient.deleteBlob(blob.name);blobsDeleted.push(blob.name)}
catch(e){blobsFailed.push({name:blob.name,error:e.message})}
}
}
}catch(azureErr){blobsFailed.push({name:'azure-connection',error:azureErr.message})}
}
res.json({ok:true,deleted:clientId,blobsDeleted:blobsDeleted.length,blobsFailed:blobsFailed.length>0?blobsFailed:undefined})
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/admin/clients/:clientId/invalidate-cache',requireAdminKey,(req,res)=>{
invalidateChunkCache(req.params.clientId)
const{RESPONSE_CACHE}=require('./src/config')
RESPONSE_CACHE.clear()
res.json({ok:true,clientId:req.params.clientId,message:'Chunk + response cache invalidated'})
})
app.post('/client/login',async(req,res)=>{
try{
const apiKey=extractApiKey(req)||req.body?.apiKey
if(!apiKey)return res.status(400).json({error:'apiKey is required'})
const client=await verifyApiKey(apiKey)
if(!client)return res.status(401).json({error:'Invalid API key'})
if(blobServiceClient)loadChunksForClient(client.clientId).catch(err=>console.warn(`[login warmup] ${client.clientId}: ${err.message}`))
res.json({ok:true,client})
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/login',async(req,res)=>{
try{
const apiKey=extractApiKey(req)||req.body?.apiKey
if(!apiKey)return res.status(400).json({error:'apiKey is required'})
const client=await verifyApiKey(apiKey)
if(!client)return res.status(401).json({error:'Invalid API key'})
if(blobServiceClient)loadChunksForClient(client.clientId).catch(err=>console.warn(`[chat/login warmup] ${client.clientId}: ${err.message}`))
res.json({ok:true,client})
}catch(err){res.status(500).json({error:err.message})}
})
app.get('/client/me',requireClientKey,async(req,res)=>{
try{
const database=await getDb()
const client=await database.collection('clients').findOne({clientId:req.client.clientId},{projection:{apiKey:0}})
if(!client)return res.status(404).json({error:'Client not found'})
res.json(client)
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/conversations',requireClientKey,async(req,res)=>{
try{
const{title}=req.body
const database=await getChatDb()
const now=new Date()
const conversation={clientId:req.client.clientId,title:title||'New Conversation',messages:[],createdAt:now,updatedAt:now}
const result=await database.collection('conversations').insertOne(conversation)
res.status(201).json({...conversation,_id:result.insertedId})
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/list',requireClientKey,async(req,res)=>{
try{
const database=await getChatDb()
const conversations=await database.collection('conversations').find({clientId:req.client.clientId},{projection:{messages:0}}).sort({updatedAt:-1}).toArray()
res.json({conversations})
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/get',requireClientKey,async(req,res)=>{
try{
const{conversationId}=req.body
if(!conversationId)return res.status(400).json({error:'conversationId is required'})
const database=await getChatDb()
const conversation=await database.collection('conversations').findOne({_id:new ObjectId(conversationId),clientId:req.client.clientId})
if(!conversation)return res.status(404).json({error:'Conversation not found'})
res.json(conversation)
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/rename',requireClientKey,async(req,res)=>{
try{
const{conversationId,title}=req.body
if(!conversationId||!title)return res.status(400).json({error:'conversationId and title are required'})
const database=await getChatDb()
const result=await database.collection('conversations').findOneAndUpdate(
{_id:new ObjectId(conversationId),clientId:req.client.clientId},
{$set:{title:title.trim(),updatedAt:new Date()}},
{returnDocument:'after',projection:{messages:0}}
)
if(!result)return res.status(404).json({error:'Conversation not found'})
res.json(result)
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/conversations/delete',requireClientKey,async(req,res)=>{
try{
const{conversationId}=req.body
if(!conversationId)return res.status(400).json({error:'conversationId is required'})
const database=await getChatDb()
const result=await database.collection('conversations').deleteOne({_id:new ObjectId(conversationId),clientId:req.client.clientId})
if(result.deletedCount===0)return res.status(404).json({error:'Conversation not found'})
res.json({ok:true,deleted:conversationId})
}catch(err){res.status(500).json({error:err.message})}
})
app.post('/chat/message',requireClientKey,withRequestTimeout(async(req,res)=>{
try{
const{query,topK=6,conversationId}=req.body
if(!query?.trim())return res.status(400).json({error:'query is required'})
const validation=validateQuery(query)
if(!validation.valid)return res.json({answer:validation.message,sources:[],conversationId:conversationId||null,client:req.client})
const{clientId,name}=req.client
const intent=detectQueryIntent(query.trim())
if(intent==='greeting'){
return res.json({
answer:"Hello! I'm your document assistant. Ask me anything about your data or documents.",
sources:[],conversationId:conversationId||null,client:{clientId,name},
})
}
const cacheKey=getCacheKey(clientId,query)
const cached=responseCacheGet(cacheKey)
if(cached){
const activeConversationId=await saveConversationMessage(clientId,conversationId||null,query.trim(),cached.answer,cached.sources||[])
return res.json({...cached,cached:true,conversationId:activeConversationId})
}
if(IN_FLIGHT.has(cacheKey)){
try{
const result=await IN_FLIGHT.get(cacheKey)
const activeConversationId=await saveConversationMessage(clientId,conversationId||null,query.trim(),result.answer,result.sources||[])
return res.json({...result,conversationId:activeConversationId})
}catch{}
}
const requestPromise=(async()=>{
const{chunks,invertedIndexes}=await loadChunksForClient(clientId)
if(chunks.length===0)return{answer:'No documents found for your account. Please ensure your documents have been ingested first.',sources:[],client:{clientId,name}}
let processedQuery=applyTypos(query.trim())
processedQuery=applySynonyms(processedQuery)
const hasSFChunks=chunks.some(c=>c.docType==='structured')
const hasUDChunks=chunks.some(c=>c.docType!=='data_dictionary'&&c.docType!=='structured')
if(hasSFChunks||hasUDChunks){
const rewritten=await preprocessQueryUD(processedQuery)
if(rewritten!==processedQuery)console.log(`[QueryPipeline] Rewritten: "${rewritten}"`)
processedQuery=rewritten
}else{
const corrected=fuzzyCorrectQuery(processedQuery,chunks.filter(c=>c.docType==='data_dictionary'))
if(corrected!==processedQuery)console.log(`[QueryPipeline] Fuzzy corrected: "${corrected}"`)
processedQuery=corrected
}
if(intent==='all_urls'){
const urlChunks=chunks.filter(c=>/https?:\/\/\S+/.test(c.text||''))
const urlEntries=extractAllUrlsFromChunks(urlChunks)
const answer=urlEntries.length>0?urlEntries.map(e=>`**${e.name}:** ${e.url}`).join('\n'):'I could not find any URLs in your documents.'
const sources=urlChunks.slice(0,6).map(h=>({source_file:h.source_file||'unknown',chunk_index:h.chunk_index??0,score:null,preview:(h.text||'').slice(0,200)}))
return{answer,sources,client:{clientId,name}}
}
const multiTopicCheck=detectMultiTopicQuery(processedQuery)
if(multiTopicCheck.isMulti){
console.log(`[chat/message] Multi-topic: ${JSON.stringify(multiTopicCheck.topics)} mode=${multiTopicCheck.mode}`)
const answer=await handleMultiTopicQuery(multiTopicCheck.topics,multiTopicCheck.mode,chunks,Math.min(topK,MAX_HITS_GLOBAL),invertedIndexes)
return{answer,sources:[],client:{clientId,name}}
}
const uniqueFiles=[...new Set(chunks.map(c=>c.source_file).filter(Boolean))]
let sourceFilter=null
if(uniqueFiles.length>1){
sourceFilter=pickBestSourceFile(processedQuery,chunks)
if(sourceFilter)console.log(`[docRouter] Using source filter: ${sourceFilter}`)
}
const{hits,docType:routedDocType}=await retrieveBestHitsAcrossAllTypes(processedQuery,chunks,Math.min(topK,MAX_HITS_GLOBAL),invertedIndexes,intent,sourceFilter)
let finalHits=hits
let finalDocType=routedDocType
if(finalHits.length===0&&sourceFilter){
console.log(`[docRouter] No hits with filter, falling back to all docs`)
const fallback=await retrieveBestHitsAcrossAllTypes(processedQuery,chunks,Math.min(topK,MAX_HITS_GLOBAL),invertedIndexes,intent,null)
finalHits=fallback.hits
finalDocType=fallback.docType
}
if(finalHits.length===0){
finalHits=relaxedKeywordSearchDD(processedQuery,chunks,64,invertedIndexes.all)
finalDocType=detectDocTypeFromChunks(finalHits.length>0?finalHits:chunks)
}
const focusedHits=selectFocusedHits(finalHits,Math.min(topK,MAX_HITS_GLOBAL))
finalHits=focusedHits.length>0?focusedHits:finalHits
console.log(`[chat/message] "${query.slice(0,60)}" -> intent=${intent}, docType=${finalDocType}, hits=${finalHits.length}, sourceFilter=${sourceFilter||'none'}`)
if(finalHits.length===0)return{answer:'I could not find relevant information about this in your documents. Try rephrasing your question.',sources:[],client:{clientId,name}}
let rawAnswer=''
if(intent!=='url_lookup'){
try{
rawAnswer=await Promise.race([
generateAnswer(processedQuery,finalHits,intent,finalDocType),
new Promise((_,reject)=>setTimeout(()=>reject(new Error('All engines timed out')),55000)),
])
}catch(err){
console.warn(`[chat/message] All engines failed: ${err.message}`)
}
}
const isBlank=!rawAnswer||rawAnswer.trim().length<15
const rawFinal=isBlank?buildFallbackAnswer(processedQuery,finalHits,intent,finalDocType):rawAnswer
const answer=cleanAnswer(rawFinal)
if(isBlank)console.warn(`[chat/message] Used fallback for: "${query.slice(0,60)}"`)
const sources=finalHits.map(h=>({
source_file:h.source_file||'unknown',
chunk_index:h.chunk_index??0,
score:typeof h._score==='number'?parseFloat(h._score.toFixed(4)):null,
preview:(h.text||'').slice(0,200),
}))
return{answer,sources,client:{clientId,name}}
})()
IN_FLIGHT.set(cacheKey,requestPromise)
let result
try{result=await requestPromise}finally{IN_FLIGHT.delete(cacheKey)}
if(result.answer&&result.answer.length>15)responseCacheSet(cacheKey,result)
const activeConversationId=await saveConversationMessage(clientId,conversationId||null,query.trim(),result.answer,result.sources||[])
res.json({...result,conversationId:activeConversationId})
}catch(err){
console.error('[chat/message] Error:',err.message)
if(!res.headersSent)res.status(500).json({error:err.message})
}
}))
app.use((err,req,res,next)=>{
console.error('[global error handler]',err)
if(!res.headersSent)res.status(500).json({error:'An unexpected error occurred. Please try again.'})
})
const PORT=process.env.PORT||4000
app.listen(PORT,()=>{
console.log(`Service running on port ${PORT}`)
console.log(`ASKDATA: ${ASKDATA_ENDPOINT?'configured':'MISSING'} | ASKDATA2: ${ASKDATA2_ENDPOINT?'configured':'missing'}`)
console.log(`Retrieval: adaptive dual-retrieval — structured BM25+alias+exact (no reranker) | semantic hybrid+rerank`)
startApiKeyHealthChecker()
warmupChunkCaches()
})
module.exports=app
