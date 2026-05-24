const {
escapeRegex,capFirst,ensureSinglePeriod,trimToCompleteSentence,
applySynonyms,normalizeQuery,extractSubject,normalizeTerms,
computeNegativePenalty,CHUNK_SIZE,MAX_HITS_GLOBAL,
detectQueryIntent,
} = require('./config')

const SF_CHUNK_SIZE = 1600
const SF_HEADING_OVERLAP = 1

function isHeading(line) {
const trimmed = line.trim()
if (!trimmed) return false
if (/^#{1,6}\s+\S/.test(trimmed)) return true
if (/^[A-Z][A-Z\s\d:.-]{3,60}$/.test(trimmed) && trimmed.length < 80) return true
if (/^\d+(\.\d+)*\s+[A-Z]/.test(trimmed) && trimmed.length < 80) return true
if (/^(Abstract|Introduction|Background|Methodology|Methods|Results|Discussion|Conclusion|References|Appendix|Summary|Overview|Executive Summary|Key Findings|Recommendations|Table of Contents)\b/i.test(trimmed)) return true
if (/^\*\*[^*]{3,60}\*\*\s*:?\s*$/.test(trimmed)) return true
if (/^[A-Z][A-Za-z\s]{3,60}:\s*$/.test(trimmed) && trimmed.length < 80) return true
if (/^[-*•]\s*\*\*[^*]+\*\*\s*:?\s*$/.test(trimmed)) return true
return false
}

function isTableRow(line) { return /\|.*\|/.test(line.trim()) || /^\s*[-+]{3,}/.test(line.trim()) }

function extractHeading(line) {
return line.trim()
.replace(/^#{1,6}\s+/,'')
.replace(/^\d+(\.\d+)*\s+/,'')
.replace(/^\*\*/,'').replace(/\*\*\s*:?\s*$/,'')
.replace(/^[-*•]\s*/,'')
.replace(/:\s*$/,'')
.trim()
}

function chunkStructuredDocument(text, sourceFile) {
const chunks = []
let chunkIndex = 0
const lines = text.replace(/\r\n/g,'\n').split('\n')
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
text: combined.trim(),
source_file: sourceFile,
chunk_index: chunkIndex++,
embedding: [],
metadata: { section: heading || '', docType: 'structured' },
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
text: tableChunk.trim(),
source_file: sourceFile,
chunk_index: chunkIndex++,
embedding: [],
metadata: { section: currentHeading, isTable: true, docType: 'structured' },
})
}
tableBuffer = []
}
inTable = false
}

if (isHeading(line)) {
flushBuffer(currentHeading)
currentHeading = extractHeading(line)
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
chunks.push({ text: tableChunk.trim(), source_file: sourceFile, chunk_index: chunkIndex++, embedding: [], metadata: { section: currentHeading, isTable: true, docType: 'structured' } })
}
}

return chunks
}

function buildInvertedIndexSF(chunks) {
const index = new Map()
for (let i = 0; i < chunks.length; i++) {
const words = (chunks[i].text || '').toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
for (const w of words) {
if (w.length < 2) continue
if (!index.has(w)) index.set(w, new Set())
index.get(w).add(i)
}
if (chunks[i].metadata && chunks[i].metadata.section) {
const sWords = chunks[i].metadata.section.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
for (const w of sWords) {
if (w.length >= 2) {
if (!index.has(w)) index.set(w, new Set())
index.get(w).add(i)
}
}
}
}
return index
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
}
}
const source = union.size > 0 ? [...union].map(i => chunks[i]).filter(Boolean) : chunks.slice(0,200)
const subjectPhraseStr = subject.toLowerCase()
const scored = source.map(c => {
const text = (c.text || '').toLowerCase()
let score = 0
const wordCoverage = queryWords.filter(w => text.includes(w)).length
score += wordCoverage * 2
const subjectCoverage = subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(text)).length
score += subjectCoverage * 4
if (c.metadata && c.metadata.section) {
const sectionLower = c.metadata.section.toLowerCase()
const sectionMatchCount = subjectWords.filter(w => sectionLower.includes(w)).length
score += sectionMatchCount * 8
if (sectionLower.includes(subjectPhraseStr)) score += 12
}
if (c.metadata && c.metadata.isTable) score += 3
if (intent === 'lookup' || intent === 'general') {
const densityBonus = Math.min(text.length / 200, 5)
score += densityBonus
}
const penalty = computeNegativePenalty(subject, c.text || '')
score -= penalty
return {...c, _score: score}
}).filter(c => c._score > 0).sort((a,b) => b._score - a._score)

let top = scored.slice(0, Math.min(topK, MAX_HITS_GLOBAL))
if (top.length === 0) {
top = chunks.filter(c => {
const t = (c.text || '').toLowerCase()
return queryWords.some(w => t.includes(w))
}).slice(0, Math.min(topK, 10)).map(c => ({...c, _score: 1}))
}
return top
}

function buildSystemPromptSF(intent) {
if (intent === 'definition') {
return `You are a research document assistant. Answer questions using ONLY the provided document context.
Rules:
- For definitions, provide a clear, concise explanation drawn directly from the document. Bold the key term.
- Write 1-2 complete sentences. Extract ONLY the definition for what was asked.
- Do not list or repeat other sections from the document that were not asked about.
- If the context lacks the answer, say: "I could not find this in your documents."`
}
if (intent === 'comparison') {
return `You are a research document assistant. Compare the requested items using ONLY the provided context. Structure your comparison clearly with each item bolded. Derive all comparisons strictly from the document. If the context lacks the answer, say: "I could not find this in your documents."`
}
if (intent === 'calculation') {
return `You are a research document assistant. Extract the formula or calculation method from the document context. Present it clearly. If the context lacks the answer, say: "I could not find this in your documents."`
}
return `You are a research document assistant. Answer questions accurately using ONLY the provided context from the document.
Rules:
- Be direct and concise. Answer in 1-3 sentences.
- Synthesize information across sections when needed but extract only what is relevant to the question.
- Reference section names when helpful (e.g., "According to the Data Retention section...").
- Do not reproduce entire document sections. Extract only the relevant fact.
- Do not invent facts not in the context.
- If context lacks the answer, say: "I could not find this in your documents."`
}

function buildContextSF(hits) {
const seen = new Set()
const deduped = []
for (const h of hits) {
const fp = (h.text || '').trim().slice(0,80).toLowerCase()
if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
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
instruction = `\n\nUsing the document context above, explain what "${subject}" means. Bold the key term. Write 1-2 complete sentences. Extract only the definition for "${subject}" and nothing else from the document.`
} else if (intent === 'calculation') {
instruction = `\n\nUsing the document context above, describe how "${subject}" is calculated or measured. Present the formula or method clearly. Stay strictly within what the document says.`
} else if (intent === 'comparison') {
instruction = `\n\nUsing the document context above, compare: ${query}. Bold each item being compared. Synthesize information from across sections as needed. Derive all comparisons strictly from the document.`
} else if (intent === 'lookup') {
instruction = `\n\nUsing the document context above, answer: ${query}. Be specific and direct. Reference section names if helpful. Extract only what answers this specific question.`
} else {
instruction = `\n\nUsing the document context above, answer: ${query}. Be concise (1-3 sentences). Reference section names when helpful. Extract only what directly answers the question. Stay strictly within the document content.`
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
chunkStructuredDocument, buildInvertedIndexSF, retrieveChunksSF,
buildSystemPromptSF, buildUserMessageSF, buildFallbackAnswerSF,
}
