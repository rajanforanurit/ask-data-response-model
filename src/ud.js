const {
escapeRegex,capFirst,ensureSinglePeriod,trimToCompleteSentence,
normalizeQuery,extractSubject,normalizeTerms,
computeNegativePenalty,MAX_HITS_GLOBAL,
detectQueryIntent,selectFocusedHits,
ASKDATA2_ENDPOINT,ASKDATA2_KEY,ASKDATA2_MODEL,ASKDATA2_REWRITE_TIMEOUT_MS,
fetchWithTimeout,
} = require('./config')
function slidingWindowChunk(text, sourceFile) {
const CHUNK_SIZE = 600
const CHUNK_OVERLAP = 120
const MIN_CHUNK = 30
const chunks = []
let chunkIndex = 0
const sentEnd = /(?<=[.!?])\s+/
const paragraphs = text.replace(/\r\n/g,'\n').split(/\n{2,}/).map(p => p.trim()).filter(Boolean)
const sentences = []
for (const para of paragraphs) {
const parts = para.split(sentEnd)
for (const part of parts) {
const sub = part.split('\n').map(s => s.trim()).filter(Boolean)
sentences.push(...sub)
}
}
let buffer = []
let bufferLen = 0
function flush() {
const text = buffer.join(' ').trim()
if (text.length >= MIN_CHUNK) {
chunks.push({ text, source_file: sourceFile, chunk_index: chunkIndex++, embedding: [] })
}
const overlapSents = []
let acc = 0
for (let i = buffer.length - 1; i >= 0; i--) {
acc += buffer[i].length + 1
overlapSents.unshift(buffer[i])
if (acc >= CHUNK_OVERLAP) break
}
buffer = overlapSents
bufferLen = overlapSents.reduce((s,t) => s + t.length + 1, 0)
}
for (const sent of sentences) {
const sentLen = sent.length + 1
if (!buffer.length && sentLen > CHUNK_SIZE) {
buffer = [sent]
bufferLen = sentLen
flush()
continue
}
if (buffer.length && bufferLen + sentLen > CHUNK_SIZE) flush()
buffer.push(sent)
bufferLen += sentLen
}
if (buffer.length) flush()
return chunks
}
function cleanOcrNoise(text) {
if (!text) return ''
return text
.replace(/[|]{2,}/g,' ')
.replace(/[_]{3,}/g,' ')
.replace(/\s{3,}/g,' ')
.replace(/([a-z])([A-Z])/g,'$1 $2')
.replace(/(\d)([A-Za-z])/g,'$1 $2')
.replace(/([A-Za-z])(\d)/g,'$1 $2')
.trim()
}
function buildInvertedIndexUD(chunks) {
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
function retrieveChunksUD(query, chunks, topK, invertedIndex) {
const intent = detectQueryIntent(query)
const subject = extractSubject(query)
const queryWords = normalizeQuery(query).replace(/[^\w\s]/g,' ').split(/\s+/).filter(w => w.length > 1)
const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
const union = new Set()
if (invertedIndex) {
for (const w of [...queryWords,...subjectWords]) {
for (const idx of (invertedIndex.get(w) || new Set())) union.add(idx)
for (const variant of normalizeTerms(w)) {
for (const idx of (invertedIndex.get(variant) || new Set())) union.add(idx)
}
}
}
const source = union.size > 0 ? [...union].map(i => chunks[i]).filter(Boolean) : chunks.slice(0,300)
const subjectPhraseStr = subject.toLowerCase()
const scored = source.map(c => {
const text = (c.text || '').toLowerCase()
let score = 0
const wordCoverage = queryWords.filter(w => text.includes(w)).length
score += wordCoverage * 3
const subjectCoverage = subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(text)).length
score += subjectCoverage * 5
if (text.includes(subjectPhraseStr)) score += 8
if (intent === 'definition' || intent === 'calculation') {
if (new RegExp(`\\b${escapeRegex(subjectPhraseStr)}\\b[\\s\\S]{0,50}(is|are|means|refers|defined)`,'i').test(text)) score += 10
}
if (intent === 'lookup' || intent === 'general') {
score += Math.min(text.length / 300, 4)
}
const penalty = computeNegativePenalty(subject, c.text || '')
score -= penalty
return {...c, _score: score}
}).filter(c => c._score > 0).sort((a,b) => b._score - a._score)
let top = selectFocusedHits(scored, Math.min(topK, MAX_HITS_GLOBAL))
if (top.length === 0) {
top = chunks.filter(c => {
const t = (c.text || '').toLowerCase()
return queryWords.some(w => t.includes(w))
}).slice(0, Math.min(topK, 10)).map(c => ({...c, _score: 1}))
}
return top
}
async function preprocessQueryUD(query) {
if (!ASKDATA2_ENDPOINT || !ASKDATA2_KEY) return query
if (query.trim().split(/\s+/).length <= 3) return query
try {
const response = await fetchWithTimeout(
ASKDATA2_ENDPOINT,
{
method: 'POST',
headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${ASKDATA2_KEY}`, 'Accept': 'application/json' },
body: JSON.stringify({
model: ASKDATA2_MODEL,
messages: [
{ role: 'system', content: 'Rewrite the query as a concise keyword search phrase. Return only the rewritten query, nothing else.' },
{ role: 'user', content: query }
],
max_tokens: 60, temperature: 0.0, stream: false
}),
},
ASKDATA2_REWRITE_TIMEOUT_MS
)
if (!response.ok) return query
const data = await response.json()
const rewritten = (data.choices?.[0]?.message?.content || '').trim().replace(/^["']|["']$/g,'')
if (rewritten && rewritten.length > 3 && rewritten.length < query.length * 2) return rewritten
return query
} catch {
return query
}
}
function buildSystemPromptUD(intent) {
if (intent === 'definition') return `Document assistant. Answer ONLY from context. Bold key term. 1-2 sentences. If not found: "I could not find this in your documents."`
if (intent === 'calculation') return `Document assistant. Extract formula/method ONLY from context. If not found: "I could not find this in your documents."`
if (intent === 'comparison') return `Document assistant. Compare ONLY from context. Bold each item. If not found: "I could not find this in your documents."`
return `Document assistant. Answer ONLY from context. Be concise. If not found: "I could not find this in your documents."`
}
function buildContextUD(hits) {
const seen = new Set()
const deduped = []
for (const h of hits) {
const fp = (h.text || '').trim().slice(0,80).toLowerCase()
if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
if (deduped.length >= 5) break
}
return deduped.map((h,i) => {
const limit = i === 0 ? 1400 : 1000
return (h.text || '').trim().slice(0,limit)
}).join('\n\n---\n\n')
}
function buildUserMessageUD(query, hits, intent) {
const context = buildContextUD(hits)
const subject = extractSubject(query)
let instruction = ''
if (intent === 'definition') {
instruction = `\n\nFrom context: what is "${subject}"? Bold the term. 1-2 sentences.`
} else if (intent === 'calculation') {
instruction = `\n\nFrom context: how is "${subject}" calculated?`
} else if (intent === 'comparison') {
instruction = `\n\nFrom context compare: ${query}. Bold each item.`
} else {
instruction = `\n\nFrom context answer: ${query}`
}
return `DOCUMENT CONTEXT:\n${context}${instruction}`
}
function buildFallbackAnswerUD(query, hits) {
if (!hits || hits.length === 0) return 'I could not find relevant information about this in your documents.'
const subject = extractSubject(query)
const subjectLower = subject.toLowerCase()
const subjectWords = subjectLower.split(/\s+/).filter(w => w.length > 1)
const escapedSubject = escapeRegex(subjectLower)
const definitionPattern = new RegExp(
`\\b${escapedSubject}\\b[^.!?\\n]{0,80}(?:is|are|means|refers to|defined as|describes)[^.!?\\n]{5,}[.!?]`,
'i'
)
for (const h of hits.slice(0, 6)) {
const text = (h.text || '').trim()
const m = text.match(definitionPattern)
if (m) {
const sentence = m[0].trim()
return ensureSinglePeriod(`**${capFirst(subject)}** ${sentence.replace(new RegExp(`^\\b${escapedSubject}\\b\\s*`, 'i'), '')}.`)
}
}
const sentences = []
for (const h of hits.slice(0, 4)) {
const text = (h.text || '').trim()
for (const sent of text.split(/(?<=[.!?])\s+/)) {
const s = sent.trim()
if (s.length < 20) continue
const hasSubject = subjectWords.some(w => s.toLowerCase().includes(w))
if (hasSubject) sentences.push(s)
}
}
if (sentences.length > 0) {
const best = sentences.sort((a, b) => {
const aScore = subjectWords.filter(w => a.toLowerCase().includes(w)).length
const bScore = subjectWords.filter(w => b.toLowerCase().includes(w)).length
return bScore - aScore
})[0]
return ensureSinglePeriod(`**${capFirst(subject)}:** ${trimToCompleteSentence(best, 300)}.`)
}
return `I could not find specific information about "${capFirst(subject)}" in your documents.`
}
module.exports = {
slidingWindowChunk,buildInvertedIndexUD,retrieveChunksUD,preprocessQueryUD,
buildSystemPromptUD,buildUserMessageUD,buildFallbackAnswerUD,cleanOcrNoise,
}
