const {
escapeRegex,capFirst,ensureSinglePeriod,trimToCompleteSentence,
applySynonyms,normalizeQuery,extractSubject,normalizeTerms,
computeNegativePenalty,CHUNK_SIZE,MAX_HITS_GLOBAL,
detectQueryIntent,
} = require('./config')
const SF_CHUNK_SIZE = 1600
const SF_HEADING_OVERLAP = 1
const BM25_K1 = 1.2
const BM25_B = 0.75
const SECTION_SYNONYMS = {
salary: ['compensation','pay','ctc','breakup','remuneration','package','wages','earnings'],
compensation: ['salary','pay','ctc','breakup','remuneration','package'],
pay: ['salary','compensation','ctc','wages'],
ctc: ['salary','compensation','breakup','package'],
bonus: ['bonus terms','performance','incentive'],
package: ['salary','compensation','ctc','breakup'],
designation: ['position','role','title','job'],
role: ['designation','position','title'],
joining: ['joining date','start date','commencement'],
notice: ['notice period','exit','resignation'],
probation: ['probation period','trial period'],
benefits: ['perks','allowances','insurance','pf'],
}
function isHeading(line) {
const trimmed = line.trim()
if (!trimmed) return false
if (/^#{1,6}\s+\S/.test(trimmed)) return true
if (/^[A-Z][A-Z\s\d:.-]{3,60}$/.test(trimmed) && trimmed.length < 80) return true
if (/^\d+(\.\d+)*\s+[A-Z]/.test(trimmed) && trimmed.length < 80) return true
if (/^(Abstract|Introduction|Background|Methodology|Methods|Results|Discussion|Conclusion|References|Appendix|Summary|Overview|Executive Summary|Key Findings|Recommendations|Table of Contents)\b/i.test(trimmed)) return true
return false
}
function isTableRow(line) {return /\|.*\|/.test(line.trim()) || /^\s*[-+]{3,}/.test(line.trim())}
function extractHeading(line) {
return line.trim()
.replace(/^#{1,6}\s+/,'')
.replace(/^\d+(\.\d+)*\s+/,'')
.trim()
}
function chunkStructuredDocument(text, sourceFile) {
const chunks = []
let chunkIndex = 0
const lines = text.replace(/\r\n/g,'\n').split('\n')
let currentSection = ''
let currentHeading = ''
let buffer = []
let bufferLen = 0
let tableBuffer = []
let inTable = false
function flushBuffer(heading) {
if (buffer.length === 0 && tableBuffer.length === 0) return
const tableText = tableBuffer.length > 0 ? '\n' + tableBuffer.join('\n') : ''
const bodyText = buffer.join('\n')
const combined = (heading ? `[Section: ${heading}]\n` : '') + bodyText + tableText
if (combined.trim().length >= 40) {
chunks.push({
text:combined.trim(),
source_file:sourceFile,
chunk_index:chunkIndex++,
embedding:[],
metadata:{section:heading||'',docType:'structured'},
})
}
buffer = []
bufferLen = 0
tableBuffer = []
inTable = false
}
for (let i = 0; i < lines.length; i++) {
const line = lines[i]
const trimmed = line.trim()
if (isTableRow(line)) {
if (!inTable) {
if (buffer.length > 0 && bufferLen > SF_CHUNK_SIZE * 0.5) flushBuffer(currentHeading)
inTable = true
}
tableBuffer.push(line)
continue
}
if (inTable && !isTableRow(line) && trimmed !== '') {
if (tableBuffer.length > 0) {
const tableChunk = (currentHeading ? `[Section: ${currentHeading}]\n` : '') + tableBuffer.join('\n')
if (tableChunk.trim().length >= 20) {
chunks.push({
text:tableChunk.trim(),
source_file:sourceFile,
chunk_index:chunkIndex++,
embedding:[],
metadata:{section:currentHeading,isTable:true,docType:'structured'},
})
}
tableBuffer = []
}
inTable = false
}
if (isHeading(line)) {
flushBuffer(currentHeading)
currentHeading = extractHeading(line)
currentSection = currentHeading
continue
}
if (!trimmed) {
if (bufferLen > SF_CHUNK_SIZE * 0.6 && buffer.length > 0) {
const lastLines = buffer.slice(-SF_HEADING_OVERLAP)
flushBuffer(currentHeading)
buffer = lastLines
bufferLen = lastLines.join('\n').length
}
continue
}
if (bufferLen + line.length > SF_CHUNK_SIZE && buffer.length > 0) {
const lastLines = buffer.slice(-SF_HEADING_OVERLAP)
flushBuffer(currentHeading)
buffer = lastLines
bufferLen = lastLines.join('\n').length
}
buffer.push(line)
bufferLen += line.length + 1
}
flushBuffer(currentHeading)
if (tableBuffer.length > 0) {
const tableChunk = (currentHeading ? `[Section: ${currentHeading}]\n` : '') + tableBuffer.join('\n')
if (tableChunk.trim().length >= 20) {
chunks.push({text:tableChunk.trim(),source_file:sourceFile,chunk_index:chunkIndex++,embedding:[],metadata:{section:currentHeading,isTable:true,docType:'structured'}})
}
}
return chunks
}
function buildInvertedIndexSF(chunks) {
const index = new Map()
const df = new Map()
const totalLen = chunks.reduce((s,c) => s + (c.text || '').length, 0)
const avgdl = chunks.length > 0 ? totalLen / chunks.length : 1
for (let i = 0; i < chunks.length; i++) {
const text = (chunks[i].text || '').toLowerCase()
const words = text.replace(/[^\w\s]/g,' ').split(/\s+/)
const seen = new Set()
for (const w of words) {
if (w.length < 2) continue
if (!index.has(w)) index.set(w, new Set())
index.get(w).add(i)
if (!seen.has(w)) {
df.set(w, (df.get(w) || 0) + 1)
seen.add(w)
}
}
if (chunks[i].metadata && chunks[i].metadata.section) {
const sWords = chunks[i].metadata.section.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
for (const w of sWords) {
if (w.length >= 2) {
if (!index.has(w)) index.set(w, new Set())
index.get(w).add(i)
if (!seen.has(w)) {
df.set(w, (df.get(w) || 0) + 1)
seen.add(w)
}
}
}
}
}
index._meta = { df, avgdl, N: chunks.length }
return index
}
function termFreqInChunk(term, text) {
let count = 0
let pos = 0
while ((pos = text.indexOf(term, pos)) !== -1) {
count++
pos += term.length
}
return count
}
function bm25Score(term, chunkText, dl, avgdl, N, df) {
const f = termFreqInChunk(term, chunkText)
if (f === 0) return 0
const dfVal = df.get(term) || 0
const idf = Math.log((N - dfVal + 0.5) / (dfVal + 0.5) + 1)
const tfNorm = (f * (BM25_K1 + 1)) / (f + BM25_K1 * (1 - BM25_B + BM25_B * dl / avgdl))
return idf * tfNorm
}
function expandSectionTerms(words) {
const expanded = new Set(words)
for (const w of words) {
const synonyms = SECTION_SYNONYMS[w.toLowerCase()] || []
for (const s of synonyms) expanded.add(s.toLowerCase())
}
return [...expanded]
}
function sectionRelevanceScore(sectionName, queryWords, subjectWords) {
if (!sectionName) return 0
const sectionLower = sectionName.toLowerCase()
const sectionTokens = sectionLower.replace(/[^\w\s]/g,' ').split(/\s+/)
const allQueryTerms = expandSectionTerms([...queryWords, ...subjectWords])
let score = 0
for (const term of allQueryTerms) {
if (sectionLower.includes(term)) score += 10
}
for (const token of sectionTokens) {
const synonyms = SECTION_SYNONYMS[token] || []
for (const syn of synonyms) {
if (allQueryTerms.some(t => t === syn || syn.includes(t) || t.includes(syn))) {
score += 6
}
}
}
return score
}
function retrieveChunksSF(query, chunks, topK, invertedIndex) {
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
for (const syn of (SECTION_SYNONYMS[w.toLowerCase()] || [])) {
for (const idx of (invertedIndex.get(syn) || new Set())) union.add(idx)
}
}
}
const source = union.size > 0 ? [...union].map(i => chunks[i]).filter(Boolean) : chunks.slice(0,200)
const meta = invertedIndex?._meta || {}
const df = meta.df || new Map()
const avgdl = meta.avgdl || 400
const N = meta.N || chunks.length
const scored = source.map(c => {
const text = (c.text || '').toLowerCase()
const dl = text.length
let score = 0
for (const term of [...queryWords, ...subjectWords]) {
score += bm25Score(term, text, dl, avgdl, N, df)
}
const secScore = sectionRelevanceScore(c.metadata && c.metadata.section ? c.metadata.section : '', queryWords, subjectWords)
score += secScore
if (c.metadata && c.metadata.isTable) score += 4
const penalty = computeNegativePenalty(subject, c.text || '')
score -= penalty
return {...c, _score:score}
}).filter(c => c._score > 0).sort((a,b) => b._score - a._score)
let top = scored.slice(0, Math.min(topK, MAX_HITS_GLOBAL))
if (top.length === 0) {
top = chunks.filter(c => {
const t = (c.text || '').toLowerCase()
return queryWords.some(w => t.includes(w))
}).slice(0, Math.min(topK, 10)).map(c => ({...c, _score:0.1}))
}
return top
}
function buildSystemPromptSF(intent) {
if (intent === 'definition') {
return `You are a research document assistant. Answer questions using ONLY the provided document context. For definitions, provide a clear, concise explanation drawn directly from the document. Bold the key term. Write in complete sentences. If the context lacks the answer, say: "I could not find this in your documents."`
}
if (intent === 'comparison') {
return `You are a research document assistant. Compare the requested items using ONLY the provided context. Structure your comparison clearly with each item bolded. Derive all comparisons strictly from the document. If the context lacks the answer, say: "I could not find this in your documents."`
}
if (intent === 'calculation') {
return `You are a research document assistant. Extract the formula or calculation method from the document context. Present it clearly. If the context lacks the answer, say: "I could not find this in your documents."`
}
return `You are a research document assistant. Answer questions accurately using ONLY the provided context from the document. Synthesize information across sections when needed. Write in clear, complete sentences. Cite section names when helpful (e.g., "According to the Methodology section..."). Do not invent facts not in the context. If context lacks the answer, say: "I could not find this in your documents."`
}
function buildContextSF(hits) {
const seen = new Set()
const deduped = []
for (const h of hits) {
const fp = (h.text || '').trim().slice(0,80).toLowerCase()
if (!seen.has(fp)) {seen.add(fp);deduped.push(h)}
if (deduped.length >= 10) break
}
return deduped.map((h,i) => {
const section = h.metadata && h.metadata.section ? `[Section: ${h.metadata.section}]` : ''
const limit = i === 0 ? 1600 : 1200
const text = (h.text || '').trim().slice(0,limit)
return section ? `${section}\n${text}` : text
}).join('\n\n---\n\n')
}
function buildUserMessageSF(query, hits, intent) {
const context = buildContextSF(hits)
const subject = extractSubject(query)
let instruction = ''
if (intent === 'definition') {
instruction = `\n\nUsing the document context above, explain what "${subject}" means. Bold the key term. Write 2-3 complete sentences. Stay strictly within what the document says.`
} else if (intent === 'calculation') {
instruction = `\n\nUsing the document context above, describe how "${subject}" is calculated or measured. Present the formula or method clearly. Stay strictly within what the document says.`
} else if (intent === 'comparison') {
instruction = `\n\nUsing the document context above, compare: ${query}. Bold each item being compared. Synthesize information from across sections as needed. Derive all comparisons strictly from the document.`
} else if (intent === 'lookup') {
instruction = `\n\nUsing the document context above, answer: ${query}. Be specific and direct. Reference section names if helpful. Stay within what the document says.`
} else {
instruction = `\n\nUsing the document context above, answer: ${query}. Synthesize information across sections as needed. Write in clear, complete sentences. Reference section names when helpful. Stay strictly within the document content.`
}
return `DOCUMENT CONTEXT:\n${context}${instruction}`
}
function buildFallbackAnswerSF(query, hits) {
if (!hits || hits.length === 0) return 'I could not find relevant information about this topic in your documents.'
const subject = extractSubject(query)
const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
const relevantSections = []
for (const h of hits.slice(0,5)) {
if (h.metadata && h.metadata.section) {
const sectionLower = h.metadata.section.toLowerCase()
const matched = subjectWords.filter(w => sectionLower.includes(w)).length
if (matched > 0) relevantSections.push(h.metadata.section)
}
}
const matchingText = []
for (const h of hits.slice(0,4)) {
const text = (h.text || '').replace(/\[Section:[^\]]+\]/,'').trim()
if (text.length > 50) matchingText.push(trimToCompleteSentence(text, 400))
}
if (matchingText.length > 0) {
const combined = matchingText.join(' ').slice(0,800)
const sectionRef = relevantSections.length > 0 ? ` (from: ${relevantSections.slice(0,2).join(', ')})` : ''
return ensureSinglePeriod(`**${capFirst(subject)}**${sectionRef}: ${combined}.`)
}
return `I could not find specific information about "${capFirst(subject)}" in your documents.`
}
module.exports = {
chunkStructuredDocument,buildInvertedIndexSF,retrieveChunksSF,
buildSystemPromptSF,buildUserMessageSF,buildFallbackAnswerSF,
}
