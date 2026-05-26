const {
  escapeRegex, capFirst, ensureSinglePeriod, trimToCompleteSentence,
  normalizeQuery, extractSubject, normalizeTerms,
  computeNegativePenalty, MAX_HITS_GLOBAL,
  detectQueryIntent, selectFocusedHits,
} = require('./config')

const SF_CHUNK_SIZE = 1600
const SF_HEADING_OVERLAP = 1

// ─── PATCH: Added cleanChunkText helper ─────────────────────────────
function cleanChunkText(text) {
  if (!text) return ''
  return text
    .replace(/\[File:[^\]]*\]\s*/g, '')
    .replace(/\[Sheet:[^\]]*\]\s*/g, '')
    .replace(/^\s*(Table Name|Measure Name|Attribute Name|Description|Source File|Sheet)\s*:[^\n]*/gim, '')
    .replace(/^[^\n]*(Table Name|Measure Name|Attribute Name)\s*:[^\n]*/gim, '')
    .replace(/\s*\|\s*(Table Name|Measure Name|Attribute Name|Description|Connected Fact Table)[^|]*\|?[^\n]*/gi, '')
    .replace(/Connected Fact Table[^\n]*/gi, '')
    .replace(/^\s*›\s*\d+\s+document\s+section[^\n]*/gim, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

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
      chunks.push({ 
        text: tableChunk.trim(), 
        source_file: sourceFile, 
        chunk_index: chunkIndex++, 
        embedding: [], 
        metadata: { section: currentHeading, isTable: true, docType: 'structured' } 
      })
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

  let top = selectFocusedHits(scored, Math.min(topK, MAX_HITS_GLOBAL))

  if (top.length === 0) {
    top = chunks.filter(c => {
      const t = (c.text || '').toLowerCase()
      return queryWords.some(w => t.includes(w))
    }).slice(0, Math.min(topK, 10)).map(c => ({...c, _score: 1}))
  }
  return top
}

function buildSystemPromptSF(intent) {
  if (intent === 'definition') return `Document assistant. Answer ONLY from context. Bold key term. 1-2 sentences. If not found: "I could not find this in your documents."`
  if (intent === 'comparison') return `Document assistant. Compare items ONLY from context. Bold each item. If not found: "I could not find this in your documents."`
  if (intent === 'calculation') return `Document assistant. Extract formula/method ONLY from context. If not found: "I could not find this in your documents."`
  return `Document assistant. Answer ONLY from context. Be direct, 1-3 sentences. Reference section names when helpful. If not found: "I could not find this in your documents."`
}

// ─── REPLACEMENT for buildContextSF ─────────────────────────────────
function buildContextSF(hits) {
  const seen = new Set()
  const deduped = []
  for (const h of hits) {
    const fp = (h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
    if (deduped.length >= 4) break
  }
  return deduped.map((h, i) => {
    const section = h.metadata && h.metadata.section ? `[Section: ${h.metadata.section}]` : ''
    const limit = i === 0 ? 1600 : 1200
    const text = cleanChunkText((h.text || '').trim()).slice(0, limit)
    return section ? `${section}\n${text}` : text
  }).join('\n\n---\n\n')
}

function buildUserMessageSF(query, hits, intent) {
  const context = buildContextSF(hits)
  const subject = extractSubject(query)
  let instruction = ''
  if (intent === 'definition') {
    instruction = `\n\nFrom context: what is "${subject}"? Bold the term. 1-2 sentences.`
  } else if (intent === 'calculation') {
    instruction = `\n\nFrom context: how is "${subject}" calculated? Present formula clearly.`
  } else if (intent === 'comparison') {
    instruction = `\n\nFrom context compare: ${query}. Bold each item.`
  } else if (intent === 'lookup') {
    instruction = `\n\nFrom context answer specifically: ${query}`
  } else {
    instruction = `\n\nFrom context answer: ${query}`
  }
  return `DOCUMENT CONTEXT:\n${context}${instruction}`
}

// ─── REPLACEMENT for buildFallbackAnswerSF ──────────────────────────
function buildFallbackAnswerSF(query, hits) {
  const { ensureSinglePeriod, trimToCompleteSentence, capFirst, extractSubject } = require('./config')
  if (!hits || hits.length === 0) return 'I could not find relevant information about this topic in your documents.'

  const subject = extractSubject(query)
  const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
  const relevantSections = []

  for (const h of hits.slice(0, 5)) {
    if (h.metadata && h.metadata.section) {
      const sectionLower = h.metadata.section.toLowerCase()
      const matched = subjectWords.filter(w => sectionLower.includes(w)).length
      if (matched > 0) relevantSections.push(h.metadata.section)
    }
  }

  const matchingText = []
  for (const h of hits.slice(0, 4)) {
    const text = cleanChunkText((h.text || ''))
      .replace(/\[Section:[^\]]+\]/, '')
      .trim()
    if (text.length > 50) matchingText.push(trimToCompleteSentence(text, 400))
  }

  if (matchingText.length > 0) {
    const combined = matchingText.join(' ').slice(0, 800)
    const sectionRef = relevantSections.length > 0 ? ` (from: ${relevantSections.slice(0, 2).join(', ')})` : ''
    return ensureSinglePeriod(`**${capFirst(subject)}**${sectionRef}: ${combined}.`)
  }

  return `I could not find specific information about "${capFirst(subject)}" in your documents.`
}

module.exports = {
  chunkStructuredDocument, 
  buildInvertedIndexSF, 
  retrieveChunksSF,
  buildSystemPromptSF, 
  buildUserMessageSF, 
  buildFallbackAnswerSF,
}
