const {
escapeRegex,capFirst,ensureSinglePeriod,trimToCompleteSentence,
applySynonyms,applyTypos,normalizeQuery,extractSubject,normalizeTerms,
computeNegativePenalty,CHUNK_SIZE,MAX_HITS_GLOBAL,
detectQueryIntent,ASKDATA2_ENDPOINT,ASKDATA2_KEY,ASKDATA2_MODEL,
ASKDATA2_REWRITE_TIMEOUT_MS,fetchWithTimeout,selectFocusedHits,
} = require('./config')

const UD_CHUNK_SIZE = 800
const UD_OVERLAP_SENTENCES = 3
const UD_MIN_CHUNK_LEN = 60

const ACRONYM_MAP = {
'occ':'occupancy','app':'application','appl':'application',
'rev':'revenue','mgmt':'management','maint':'maintenance',
'prop':'property','res':'resident','avg':'average',
'num':'number','cnt':'count','pct':'percentage',
'sq ft':'square feet','sqft':'square feet','bd':'bedroom',
'ba':'bathroom','mo':'month','yr':'year','q1':'quarter 1',
'q2':'quarter 2','q3':'quarter 3','q4':'quarter 4',
'ytd':'year to date','mtd':'month to date','roi':'return on investment',
'noi':'net operating income','gpr':'gross potential rent',
'egi':'effective gross income','opex':'operating expenses',
}

const NOISE_PATTERNS = [
/[^\x00-\x7F]{3,}/g,
/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g,
/[ \t]{4,}/g,
/\.{4,}/g,
/_{4,}/g,
/-{4,}/g,
/={4,}/g,
/\*{4,}/g,
]

function cleanOcrNoise(text) {
if (!text) return ''
let cleaned = text
for (const pattern of NOISE_PATTERNS) cleaned = cleaned.replace(pattern,' ')
cleaned = cleaned.replace(/\n{3,}/g,'\n\n')
.replace(/[ \t]+\n/g,'\n')
.replace(/\n[ \t]+/g,'\n')
.replace(/[ \t]{2,}/g,' ')
.trim()
return cleaned
}

function expandAcronyms(text) {
let result = text
for (const [abbr, full] of Object.entries(ACRONYM_MAP)) {
const pattern = new RegExp(`\\b${escapeRegex(abbr)}\\b`,'gi')
result = result.replace(pattern, full)
}
return result
}

function splitIntoSentences(text) {
const raw = text.replace(/([.!?])\s+([A-Z])/g,'$1\n$2')
.replace(/([.!?])\n/g,'$1\n')
.split('\n')
.map(s => s.trim())
.filter(s => s.length > 10)
return raw
}

function isBulletHeading(line) {
const t = line.trim()
if (!t) return false
if (/^\*\*[^*]+\*\*\s*:?\s*$/.test(t)) return true
if (/^#{1,6}\s+\S/.test(t)) return true
if (/^[A-Z][A-Za-z\s]{3,50}:\s*$/.test(t)) return true
return false
}

function extractBulletHeadingText(line) {
return line.trim()
.replace(/^\*\*/,'').replace(/\*\*\s*:?\s*$/,'')
.replace(/^#{1,6}\s+/,'')
.replace(/:\s*$/,'')
.trim()
}

function slidingWindowChunk(text, sourceFile) {
const chunks = []
let chunkIndex = 0
const cleaned = cleanOcrNoise(text)
const expanded = expandAcronyms(cleaned)

const lines = expanded.replace(/\r\n/g,'\n').split('\n')
let currentHeading = ''
let buffer = []
let bufferLen = 0

function flush() {
const combined = buffer.join('\n')
if (combined.trim().length >= UD_MIN_CHUNK_LEN) {
chunks.push({
text: combined.trim(),
source_file: sourceFile,
chunk_index: chunkIndex++,
embedding: [],
metadata: { docType: 'unstructured', section: currentHeading || '' },
})
}
}

let isBulletDoc = lines.filter(l => /^\s*[-*•]\s+/.test(l) || isBulletHeading(l)).length > lines.filter(l => l.trim()).length * 0.3

if (isBulletDoc) {
for (const line of lines) {
const trimmed = line.trim()
if (!trimmed) continue

if (isBulletHeading(line)) {
if (buffer.length > 0) { flush(); buffer = []; bufferLen = 0 }
currentHeading = extractBulletHeadingText(line)
buffer.push(line)
bufferLen = line.length
continue
}

const projected = bufferLen + line.length + 1
if (projected > UD_CHUNK_SIZE && buffer.length > 0) {
flush()
const overlap = buffer.slice(-UD_OVERLAP_SENTENCES)
buffer = [...overlap]
bufferLen = buffer.join('\n').length
}
buffer.push(line)
bufferLen += line.length + 1
}
if (buffer.length > 0) flush()
return chunks
}

const paragraphs = expanded.replace(/\r\n/g,'\n').split(/\n{2,}/).map(p => p.trim()).filter(p => p.length > 20)
let overlapBuffer = []

function flushPara() {
const combined = buffer.join('\n\n')
if (combined.trim().length >= UD_MIN_CHUNK_LEN) {
chunks.push({
text: combined.trim(),
source_file: sourceFile,
chunk_index: chunkIndex++,
embedding: [],
metadata: { docType: 'unstructured', section: '' },
})
}
}

for (const para of paragraphs) {
if (para.length > UD_CHUNK_SIZE * 1.5) {
if (buffer.length > 0) { flushPara(); overlapBuffer = buffer.slice(-UD_OVERLAP_SENTENCES); buffer = [...overlapBuffer]; bufferLen = buffer.join('\n\n').length }
const sentences = splitIntoSentences(para)
let sentBuf = [...overlapBuffer]
let sentLen = sentBuf.join(' ').length
for (const sent of sentences) {
if (sentLen + sent.length > UD_CHUNK_SIZE && sentBuf.length > 0) {
const combined = sentBuf.join(' ')
if (combined.length >= UD_MIN_CHUNK_LEN) {
chunks.push({ text: combined.trim(), source_file: sourceFile, chunk_index: chunkIndex++, embedding: [], metadata: { docType: 'unstructured', section: '' } })
}
sentBuf = sentBuf.slice(-UD_OVERLAP_SENTENCES)
sentLen = sentBuf.join(' ').length
}
sentBuf.push(sent)
sentLen += sent.length + 1
}
if (sentBuf.length > 0) {
const combined = sentBuf.join(' ')
if (combined.length >= UD_MIN_CHUNK_LEN) {
chunks.push({ text: combined.trim(), source_file: sourceFile, chunk_index: chunkIndex++, embedding: [], metadata: { docType: 'unstructured', section: '' } })
}
}
overlapBuffer = sentBuf.slice(-UD_OVERLAP_SENTENCES)
buffer = []
bufferLen = 0
continue
}
const projected = bufferLen + (bufferLen > 0 ? 2 : 0) + para.length
if (buffer.length > 0 && projected > UD_CHUNK_SIZE) {
flushPara()
overlapBuffer = buffer.slice(-UD_OVERLAP_SENTENCES)
buffer = [...overlapBuffer]
bufferLen = buffer.join('\n\n').length
}
buffer.push(para)
bufferLen += (bufferLen > 0 ? 2 : 0) + para.length
}
if (buffer.length > 0) flushPara()
return chunks
}

function buildInvertedIndexUD(chunks) {
const index = new Map()
for (let i = 0; i < chunks.length; i++) {
const text = (chunks[i].text || '').toLowerCase()
const section = (chunks[i].metadata?.section || '').toLowerCase()
const combined = text + ' ' + section
const words = combined.replace(/[^\w\s]/g,' ').split(/\s+/)
for (const w of words) {
if (w.length < 2) continue
if (!index.has(w)) index.set(w, new Set())
index.get(w).add(i)
const expanded = expandAcronyms(w)
if (expanded !== w) {
if (!index.has(expanded)) index.set(expanded, new Set())
index.get(expanded).add(i)
}
}
}
return index
}

function needsRewrite(query) {
const trimmed = query.trim()
const words = trimmed.split(/\s+/).filter(Boolean)
if (words.length <= 2) return true
if (/[^\x00-\x7F]/.test(trimmed) && words.length < 5) return true
if (/(.)\\1{3,}/.test(trimmed)) return true
if (words.length < 4 && !/\b(what|how|define|explain|formula|calculate|list|show|find)\b/i.test(trimmed)) return true
return false
}

async function rewriteQuery(query) {
if (!ASKDATA2_ENDPOINT || !ASKDATA2_KEY) return query
if (!needsRewrite(query)) return query
try {
const response = await fetchWithTimeout(
ASKDATA2_ENDPOINT,
{
method: 'POST',
headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${ASKDATA2_KEY}`, 'Accept': 'application/json' },
body: JSON.stringify({
model: ASKDATA2_MODEL,
messages: [
{ role: 'system', content: 'You are a query rewriter for a RAG system on unstructured text documents. Fix spelling, grammar, and expand abbreviations. Normalize ambiguous phrasing into a clear question. Return ONLY the rewritten query as plain text. No explanation. If already correct, return unchanged.' },
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
const rewritten = (data.choices?.[0]?.message?.content || '').trim().replace(/^["']|["']$/g,'').trim()
if (!rewritten || rewritten.length < 3 || rewritten.length > query.length * 4) return query
if (rewritten.toLowerCase() !== query.toLowerCase()) console.log(`[UD:rewrite] "${query}" -> "${rewritten}"`)
return rewritten
} catch (err) {
console.warn(`[UD:rewrite] Failed: ${err.message}`)
return query
}
}

async function preprocessQueryUD(query) {
let q = applyTypos(query)
q = applySynonyms(q)
q = expandAcronyms(q)
q = await rewriteQuery(q)
return q
}

function retrieveChunksUD(query, chunks, topK, invertedIndex) {
const intent = detectQueryIntent(query)
const subject = extractSubject(query)
const queryWords = normalizeQuery(query).replace(/[^\w\s]/g,' ').split(/\s+/).filter(w => w.length > 1)
const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
const expandedWords = [...new Set([...queryWords,...subjectWords,...queryWords.map(w => expandAcronyms(w))])]
const union = new Set()
if (invertedIndex) {
for (const w of expandedWords) {
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
const section = (c.metadata?.section || '').toLowerCase()
let score = 0
const wordCoverage = expandedWords.filter(w => text.includes(w)).length
score += wordCoverage * 2
if (text.includes(subjectPhraseStr)) score += 10
const subjectWordCoverage = subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(text)).length
score += subjectWordCoverage * 3
if (section) {
const sectionMatchCount = subjectWords.filter(w => section.includes(w)).length
score += sectionMatchCount * 8
if (section.includes(subjectPhraseStr)) score += 12
}
const density = Math.min(text.length / 300, 6)
score += density
if (intent === 'calculation' && (/formula|calculated|computed|equation/i.test(text))) score += 8
if (intent === 'definition' && (/is defined as|means|refers to|describes/i.test(text))) score += 6
const penalty = computeNegativePenalty(subject, c.text || '')
score -= penalty
return {...c, _score: score}
}).filter(c => c._score > 0).sort((a,b) => b._score - a._score)

let top = selectFocusedHits(scored, Math.min(topK, MAX_HITS_GLOBAL))
if (top.length < 3) {
const fallback = chunks
.filter(c => {
const t = (c.text || '').toLowerCase()
return expandedWords.some(w => t.includes(w))
})
.map(c => ({...c, _score: 1}))
.slice(0, Math.min(topK, 10))
const existingIds = new Set(top.map(c => c.chunk_index))
for (const c of fallback) {
if (!existingIds.has(c.chunk_index)) top.push(c)
if (top.length >= Math.min(topK, MAX_HITS_GLOBAL)) break
}
}
return top
}

function buildSystemPromptUD(intent) {
if (intent === 'definition') {
return `You are a precise document assistant. Your job is to answer ONLY what was asked using the provided context.
Rules:
- Answer in 1-2 sentences maximum.
- Bold the key term being defined.
- Extract ONLY the specific fact or definition requested. If the context contains multiple sections or bullet points, extract only the one that directly answers the question.
- Do NOT list or summarize other sections that were not asked about.
- If the context lacks the answer, say: "I could not find this in your documents."`
}
if (intent === 'calculation') {
return `You are a precise document assistant. Extract only the formula or calculation method from the context.
Rules:
- Return only the formula or calculation. Nothing else.
- If the context lacks the answer, say: "I could not find this in your documents."`
}
if (intent === 'lookup') {
return `You are a precise document assistant. Answer ONLY the specific question asked using the provided context.
Rules:
- Be direct and concise. Answer in 1-3 sentences.
- If the context is a list of bullet points or policy sections, extract ONLY the bullet or section that directly answers the question. Do not reproduce the entire list.
- Do not include document titles, greetings, preambles, or unrelated clauses.
- If the context lacks the answer, say: "I could not find this in your documents."`
}
return `You are a precise document assistant. Answer ONLY the specific question asked using the provided context.
Rules:
- Be direct and concise. Answer in 1-3 sentences.
- If the context contains a structured list, policy document, or bullet points, identify and extract ONLY the section or bullet that answers the question. Do not summarize or repeat the entire document.
- Do not repeat or summarize the entire document section.
- Do not include document titles, greetings, preambles, or content from sections that were not asked about.
- If the context lacks the answer, say: "I could not find this in your documents."`
}

function buildContextUD(hits, query) {
const seen = new Set()
const deduped = []
for (const h of hits) {
const fp = (h.text || '').trim().slice(0,80).toLowerCase()
if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
if (deduped.length >= 4) break
}

const subject = extractSubject(query)
const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)

return deduped.map((h, i) => {
const rawText = (h.text || '').trim()
const section = h.metadata?.section || ''
const lines = rawText.split('\n')
const isBulletList = lines.filter(l => /^\s*[-*•]\s+/.test(l) || isBulletHeading(l)).length > lines.length * 0.3

if (isBulletList && subjectWords.length > 0) {
const relevantLines = []
let inRelevantSection = false
for (const line of lines) {
const lineLower = line.toLowerCase()
const isHeading = isBulletHeading(line)
if (isHeading) {
const headingText = extractBulletHeadingText(line).toLowerCase()
inRelevantSection = subjectWords.some(w => headingText.includes(w))
if (inRelevantSection) relevantLines.push(line)
continue
}
if (inRelevantSection) {
relevantLines.push(line)
} else if (subjectWords.some(w => lineLower.includes(w))) {
relevantLines.push(line)
}
}
if (relevantLines.length > 0) {
const sectionPrefix = section ? `[Section: ${section}]\n` : ''
return sectionPrefix + relevantLines.join('\n').slice(0, i < 2 ? 600 : 400)
}
}

const limit = i < 2 ? 600 : 400
const sectionPrefix = section ? `[Section: ${section}]\n` : ''
return sectionPrefix + rawText.slice(0, limit)
}).join('\n\n---\n\n')
}

function buildUserMessageUD(query, hits, intent) {
const context = buildContextUD(hits, query)
const subject = extractSubject(query)
let instruction = ''
if (intent === 'definition') {
instruction = `\n\nUsing only the context above, what is "${subject}"? Write 1-2 sentences. Bold the key term. Return only the definition, nothing else. Do not list other sections.`
} else if (intent === 'calculation') {
instruction = `\n\nUsing only the context above, how is "${subject}" calculated? Return only the formula or method.`
} else {
instruction = `\n\nUsing only the context above, answer this specific question in 1-3 sentences: ${query}\n\nReturn only the direct answer. Do not include document titles, names, or content from sections that were not asked about.`
}
return `CONTEXT:\n${context}${instruction}`
}

function buildFallbackAnswerUD(query, hits) {
if (!hits || hits.length === 0) return 'I could not find relevant information in your documents.'
const subject = extractSubject(query)
const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
const queryLower = query.toLowerCase()
const scoredSentences = []
for (const h of hits.slice(0,4)) {
const text = (h.text || '').trim()
if (text.length < 20) continue
const lines = text.split('\n')
const isBulletList = lines.filter(l => /^\s*[-*•]\s+/.test(l) || isBulletHeading(l)).length > lines.length * 0.3
if (isBulletList) {
let inSection = false
for (const line of lines) {
const lineLower = line.toLowerCase()
if (isBulletHeading(line)) {
const ht = extractBulletHeadingText(line).toLowerCase()
inSection = subjectWords.some(w => ht.includes(w))
continue
}
if (inSection && line.trim().length > 10) {
scoredSentences.push({ sent: line.trim(), score: 5 })
}
if (!inSection && subjectWords.some(w => lineLower.includes(w)) && line.trim().length > 10) {
scoredSentences.push({ sent: line.trim(), score: 2 })
}
}
continue
}
const sentences = splitIntoSentences(text)
for (const sent of sentences) {
const sentLower = sent.toLowerCase()
const matchCount = subjectWords.filter(w => sentLower.includes(w)).length
const isHeader = /^(dear|letter|appointment|designation|department|offer details|compensation|bonus terms|other terms|notice period|employee|sincerely|we are pleased)/i.test(sent.trim())
if (matchCount > 0 && !isHeader && sent.length > 15) {
scoredSentences.push({ sent, score: matchCount + (sentLower.includes(queryLower) ? 5 : 0) })
}
}
}
scoredSentences.sort((a, b) => b.score - a.score)
const best = [...new Set(scoredSentences.slice(0,2).map(s => s.sent))]
if (best.length > 0) return ensureSinglePeriod(best.join(' '))
return 'I could not find relevant information in your documents.'
}

module.exports = {
slidingWindowChunk, buildInvertedIndexUD, retrieveChunksUD, preprocessQueryUD,
buildSystemPromptUD, buildUserMessageUD, buildFallbackAnswerUD, cleanOcrNoise, expandAcronyms,
}
