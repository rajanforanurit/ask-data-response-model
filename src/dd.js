const XLSX = require('xlsx')
const {
  escapeRegex, capFirst, ensureSinglePeriod, trimToCompleteSentence,
  applySynonyms, normalizeQuery, extractSubject, extractUrlKeywords,
  normalizeTerms, extractFormulaFromText, computeNegativePenalty,
  MAX_HITS_GLOBAL, DOMAIN_SHORT_SAFELIST,
} = require('./config')
const stringSimilarity = require('string-similarity')
const { levenshteinSimilarity } = require('./config')
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
    .replace(/^(\s*\|[^\n]*){1,}/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
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
    [/^name$/, 90], [/report.?name/i, 85], [/\bname\b/, 70], [/\btitle\b/, 50],
  ]
  const TABLE_PATTERNS = [
    [/\b(table|module|category|group|domain|section|workspace)\b/, 100], [/^table$/, 90],
  ]
  const DESC_PATTERNS = [
    [/\b(description|desc|definition|about|summary|detail)\b/, 100], [/^desc$/, 90],
  ]
  const FORMULA_PATTERNS = [
    [/\b(formula|calculation|calc|how\s+calculated|computed\s+as)\b/, 100], [/^formula$/, 90],
  ]
  const URL_PATTERNS = [[/\b(url|link|href|report\s+link|dashboard)\b/, 100]]
  const ADDITIONAL_PATTERNS = [[/\b(additional|extra|notes?|info|configuration|config|mdm)\b/, 100]]

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
      if (h !== '' && h !== 'nan') { lastNonBlank = h; headers.push(h) }
      else headers.push(lastNonBlank || `Col${headers.length + 1}`)
    }

    const colIdx = detectColumns(headers)
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
          /(.*?\/.*?)/i, /(=.*?)/i, /(calculated\s+as.*)/i, /(computed\s+as.*)/i,
          /(divided\s+by.*)/i, /(multiplied\s+by.*)/i, /(sum\s+of.*)/i
        ]
        for (const pattern of formulaPatterns) {
          const match = descVal.match(pattern)
          if (match && match[0].trim().length > 3) { formulaVal = match[0].trim(); break }
        }
      }

      if (nameVal) {
        let synthesis = `${nameVal}`
        if (tableVal && tableVal !== sheetName) synthesis += ` (${tableVal})`
        if (descVal) synthesis += ` is defined as: ${descVal}`
        if (formulaVal && !descVal.toLowerCase().includes(formulaVal.toLowerCase())) synthesis += `. Formula: ${formulaVal}`
        if (additionalVal) synthesis += `. Additional Info: ${additionalVal}`
        if (urlVal) synthesis += `. URL: ${urlVal}`

        rows.push({ text: synthesis, metadata: { measure: nameVal, table: tableVal || sheetName, formula: formulaVal || '', description: descVal || '', url: urlVal || '', sourceSheet: sheetName } })

        if (formulaVal) {
          rows.push({ text: `How to calculate ${nameVal}: ${formulaVal}`, metadata: { measure: nameVal, table: tableVal || sheetName, formula: formulaVal, description: descVal || '', url: '', sourceSheet: sheetName } })
          rows.push({ text: `Formula for ${nameVal}: ${formulaVal}`, metadata: { measure: nameVal, table: tableVal || sheetName, formula: formulaVal, description: descVal || '', url: '', sourceSheet: sheetName } })
        }
        if (urlVal) {
          rows.push({ text: `Report URL for ${nameVal}: ${urlVal}`, metadata: { measure: nameVal, table: tableVal || sheetName, formula: '', description: '', url: urlVal, sourceSheet: sheetName } })
          rows.push({ text: `Power BI link for ${nameVal}: ${urlVal}`, metadata: { measure: nameVal, table: tableVal || sheetName, formula: '', description: '', url: urlVal, sourceSheet: sheetName } })
          if (tableVal && tableVal !== sheetName) rows.push({ text: `Report URL for ${nameVal} (${tableVal}): ${urlVal}`, metadata: { measure: nameVal, table: tableVal, formula: '', description: '', url: urlVal, sourceSheet: sheetName } })
        }
        rowsEmitted++
      } else if (descVal) {
        rows.push({ text: descVal, metadata: { measure: '', table: tableVal || sheetName, formula: '', description: descVal, url: '', sourceSheet: sheetName } })
      }
    }

    if (rowsEmitted === 0) {
      for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
        const row = rawRows[i]
        const cells = row.map(c => String(c || '').trim()).filter(Boolean)
        if (cells.length) rows.push({ text: cells.join(' | '), metadata: { measure: '', table: sheetName, formula: '', description: '', url: '', sourceSheet: sheetName } })
      }
    }
  }
  return rows
}

function buildVocabulary(chunks) {
  const vocab = new Set()
  const stopWords = new Set(['is','the','a','an','of','in','for','to','at','by','as','on','or','and','be','it','its','with','that','this','from','are','was','were'])
  for (const chunk of chunks) {
    const words = (chunk.text || '').toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
    for (const w of words) {
      if (w.length >= 3 && !stopWords.has(w)) vocab.add(w)
    }
    if (chunk.metadata && chunk.metadata.measure) {
      const measureWords = chunk.metadata.measure.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/)
      for (const w of measureWords) {if (w.length >= 3 && !stopWords.has(w)) vocab.add(w)}
    }
  }
  return [...vocab]
}

function fuzzyCorrectQuery(query, chunks) {
  if (!chunks || chunks.length === 0) return query
  const vocabulary = buildVocabulary(chunks)
  if (vocabulary.length === 0) return query
  const stopWords = new Set(['what','is','are','how','the','a','an','of','in','for','to','at','by','as','on','or','and','define','explain','show','find','get','list','give'])
  const words = query.split(/\s+/)
  const corrected = words.map(word => {
    const wordLower = word.toLowerCase()
    if (stopWords.has(wordLower)) return word
    if (DOMAIN_SHORT_SAFELIST.has(wordLower)) return word
    if (wordLower.length < 6) return word
    if (vocabulary.includes(wordLower)) return word
    const {bestMatch} = stringSimilarity.findBestMatch(wordLower, vocabulary)
    const levSim = levenshteinSimilarity(wordLower, bestMatch.target)
    const combinedScore = bestMatch.rating * 0.6 + levSim * 0.4
    if (combinedScore >= 0.72 && bestMatch.target !== wordLower) {
      return bestMatch.target
    }
    return word
  })
  return corrected.join(' ')
}

function keywordSearchDD(query, chunks, topK, intent, invertedIndex) {
  const subject = extractSubject(query)
  const subjectWords = subject.toLowerCase().split(/\s+/).filter(w => w.length > 1)
  const queryLower = normalizeQuery(query)
  const isMultiWord = subjectWords.length > 1
  const subjectPhraseRegex = isMultiWord
    ? new RegExp(escapeRegex(subject.toLowerCase()),'i')
    : new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b`,'i')

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
      for (const w of ['url','link','https','http','powerbi','app']) {
        for (const idx of (invertedIndex.get(w) || new Set())) union.add(idx)
      }
    }
    candidateIndices = union
  }

  const source = candidateIndices ? [...candidateIndices].map(i => chunks[i]).filter(Boolean) : chunks.slice(0,200)

  return source.map(c => {
    const text = (c.text || '').toLowerCase()
    let score = 0

    if (intent === 'all_urls') {
      if (!text.includes('http')) return {...c,_score:0}
      return {...c,_score:10}
    }
    if (intent === 'url_lookup') {
      if (!text.includes('http')) return {...c,_score:0}
      const kws = extractUrlKeywords(query)
      const matched = kws.filter(w => text.includes(w)).length
      if (matched === 0) return {...c,_score:0}
      score += matched * 10
      if (text.includes(kws.join(' '))) score += 15
    } else {
      const phraseFound = subjectPhraseRegex.test(c.text || '')
      if (phraseFound) {
        score += subjectWords.length * 6
        if (new RegExp(`\\|\\s*${escapeRegex(subject.toLowerCase())}\\s*\\|`,'i').test(c.text || '')) score += subjectWords.length * 4
        if (new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b[\\s\\S]{0,30}(is defined as|is calculated as|formula:)`,'i').test(c.text || '')) score += subjectWords.length * 8
      }
      const wordCoverage = subjectWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(c.text || '')).length
      score += wordCoverage * 2
      if (new RegExp(`\\b${escapeRegex(queryLower)}\\b`,'i').test(c.text || '')) score += 3
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
        if (measureLower === subject.toLowerCase().trim()) score += 100
        else if (subjectWords.some(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(measureLower))) score += 10
      }

      const penalty = computeNegativePenalty(subject, c.text || '')
      score -= penalty
    }
    return {...c,_score:score}
  }).filter(c => c._score > 0).sort((a,b) => b._score - a._score).slice(0,topK)
}

function relaxedKeywordSearchDD(query, chunks, topK, invertedIndex) {
  const subject = extractSubject(query)
  const allWords = [
    ...subject.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/).filter(w => w.length > 1),
    ...query.toLowerCase().replace(/[^\w\s]/g,' ').split(/\s+/).filter(w => w.length > 2),
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
  const source = union.size > 0 ? [...union].map(i => chunks[i]).filter(Boolean) : chunks.slice(0,300)

  return source.map(c => {
    const text = (c.text || '').toLowerCase()
    const matched = uniqueWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(text)).length
    const subjectMatch = subject.length > 2 && new RegExp(`\\b${escapeRegex(subject.toLowerCase())}\\b`,'i').test(text) ? 5 : 0
    let metaBoost = 0
    if (c.metadata && c.metadata.measure) {
      const ml = c.metadata.measure.toLowerCase()
      if (ml === subject.toLowerCase().trim()) metaBoost += 50
      else {
        const subjectMatched = uniqueWords.filter(w => new RegExp(`\\b${escapeRegex(w)}\\b`,'i').test(ml)).length
        metaBoost = subjectMatched * 3
      }
    }
    const penalty = computeNegativePenalty(subject, c.text || '')
    return {...c,_score:Math.max(0, matched + subjectMatch + metaBoost - penalty)}
  }).filter(c => c._score > 0).sort((a,b) => b._score - a._score).slice(0,topK)
}

async function retrieveChunksDD(query, chunks, topK, invertedIndex, _isRetry = false) {
  const intent = require('./config').detectQueryIntent(query)
  const normalizedQuery = normalizeQuery(query).replace(/[^\w\s]/g,' ').replace(/\s+/g,' ')
  if (intent === 'all_urls') return chunks.filter(c => /https?:\/\/\S+/.test(c.text || '')).slice(0,100)

  const candidates = keywordSearchDD(normalizedQuery, chunks, Math.min(150, chunks.length), intent, invertedIndex)
  const pool = candidates.length > 0 ? candidates : chunks.slice(0,150)
  const topScore = pool[0]?._score || 0
  let topCandidates = []

  if (topScore >= 6) topCandidates = pool.slice(0, Math.min(MAX_HITS_GLOBAL, pool.length))
  else if ((intent === 'definition' || intent === 'calculation') && topScore >= 3) topCandidates = pool.slice(0, Math.min(MAX_HITS_GLOBAL, pool.length))
  else if (intent === 'url_lookup' && pool.length > 0) return pool.slice(0, Math.min(topK, 6))

  if (topCandidates.length === 0 && pool.length > 0) topCandidates = pool.slice(0, Math.min(MAX_HITS_GLOBAL, pool.length))

  if (topCandidates.length === 0 && !_isRetry) {
    const corrected = fuzzyCorrectQuery(query, chunks)
    if (corrected.toLowerCase() !== query.toLowerCase()) {
      return retrieveChunksDD(corrected, chunks, topK, invertedIndex, true)
    }
  }

  if (topCandidates.length === 0) topCandidates = relaxedKeywordSearchDD(normalizedQuery, chunks, Math.min(topK * 2, 64), invertedIndex).slice(0, Math.min(topK, MAX_HITS_GLOBAL))
  return topCandidates.slice(0, Math.min(topK, MAX_HITS_GLOBAL))
}

function buildSystemPromptDD(intent) {
  const intentRule = intent === 'definition'
    ? `Definition: Bold measure name, one sentence definition only.`
    : intent === 'calculation'
    ? `Calculation: Output ONLY "**Formula for [Name]:** [formula]."`
    : intent === 'comparison'
    ? `Comparison: Bold each name. One definition each. End with "**Key Difference:**" from context only.`
    : `General: Answer in 2-4 sentences directly.`

  return `Data dictionary assistant. Answer ONLY from context. Bold subject. Complete sentences. No pipe characters. No source references.\nIf answer not in context say: "I could not find this in your documents."\n${intentRule}\n\nCRITICAL: Never include file names, sheet names, table metadata, or source references in your answer. Never repeat lines like "[File: ...]", "[Sheet: ...]", "Table Name:", "Measure Name:", "Attribute Name:", or "Description:" labels. Output ONLY the clean answer text.`
}

function buildUserMessageDD(query, hits, intent) {
  const context = buildContextDD(hits)
  const subject = extractSubject(query)
  let instruction = ''

  if (intent === 'definition') {
    instruction = `\n\nFrom context: one-sentence definition of "${subject}". Bold the name.`
  } else if (intent === 'calculation') {
    instruction = `\n\nFrom context: return only "**Formula for ${capFirst(subject)}:** [formula]."`
  } else if (intent === 'url_lookup') {
    instruction = `\n\nFrom context: return only the full URL for "${extractUrlKeywords(query).join(' ')}".`
  } else if (intent === 'all_urls') {
    instruction = `\n\nFrom context: list ALL URLs. Format: name: URL. One per line.`
  } else if (intent === 'comparison') {
    instruction = `\n\nFrom context compare: ${query}. Bold each name. End with "**Key Difference:**" from context only.`
  } else {
    instruction = `\n\nFrom context answer: ${query}`
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
        const cleanUrl = url.replace(/[.,;)]+$/,'').trim()
        if (!cleanUrl.startsWith('http') || seen.has(cleanUrl)) continue
        seen.add(cleanUrl)
        let name = 'Report'
        const reportUrlMatch = line.match(/^(?:Report URL|Power BI link)\s+for\s+(.+?)(?:\s*\([^)]+\))?\s*:\s*https?:/i)
        if (reportUrlMatch) {
          name = reportUrlMatch[1].trim()
        } else {
          const beforeUrl = line.slice(0, line.indexOf('http')).trim()
          if (beforeUrl) {
            const cleaned = beforeUrl.replace(/\.\s*URL\s*:?\s*$/i,'').replace(/\s*:\s*$/,'').replace(/^(URL|Link|Dashboard|Report)\s*:?\s*/i,'').trim()
            if (cleaned.length > 1 && cleaned.length < 120) name = cleaned
          }
        }
        results.push({name, url:cleanUrl})
      }
    }
  }
  return results
}

function buildContextDD(hits) {
  const seen = new Set()
  const deduped = []
  for (const h of hits) {
    const fp = (h.text || '').trim().slice(0, 80).toLowerCase()
    if (!seen.has(fp)) { seen.add(fp); deduped.push(h) }
    if (deduped.length >= 8) break
  }
  return deduped.map((h, i) => {
    const limit = i === 0 ? 1200 : 900
    const text = cleanChunkText((h.text || '').trim()).slice(0, limit)
    return text
  }).filter(Boolean).join('\n\n---\n\n')
}

function buildFallbackAnswerDD(query, hits, intent) {
  const { ensureSinglePeriod, trimToCompleteSentence, capFirst, escapeRegex, extractSubject, extractUrlKeywords, extractFormulaFromText, detectQueryIntent } = require('./config')

  if (!hits || hits.length === 0) return 'I could not find relevant information about this in your documents.'

  const resolvedIntent = intent || detectQueryIntent(query)
  const subject = extractSubject(query)
  const subjectLower = subject.toLowerCase().trim()
  const subjectWords = subjectLower.split(/\s+/).filter(w => w.length > 1)
  const escapedSubject = escapeRegex(subjectLower)

  if (resolvedIntent === 'all_urls') {
    const urlEntries = extractAllUrlsFromChunks(hits)
    if (urlEntries.length === 0) return 'I could not find any URLs in your documents.'
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
    return 'I could not find a matching URL in your documents.'
  }

  if (resolvedIntent === 'calculation') {
    for (const h of hits) {
      if (h.metadata && h.metadata.formula && new RegExp(`\\b${escapedSubject}\\b`, 'i').test(h.metadata.measure || '')) {
        return ensureSinglePeriod(`**Formula for ${capFirst(h.metadata.measure)}:** ${h.metadata.formula}.`)
      }
    }
    const calcPattern = new RegExp(`how to calculate ${escapedSubject}:\\s*([^\\n]+)`, 'im')
    for (const h of hits) {
      const m = cleanChunkText(h.text || '').match(calcPattern)
      if (m) return ensureSinglePeriod(`**Formula for ${capFirst(subject)}:** ${trimToCompleteSentence(m[1].trim(), 500)}.`)
    }
    const formulaPattern = new RegExp(`formula for ${escapedSubject}:\\s*([^\\n]+)`, 'im')
    for (const h of hits) {
      const m = cleanChunkText(h.text || '').match(formulaPattern)
      if (m) return ensureSinglePeriod(`**Formula for ${capFirst(subject)}:** ${trimToCompleteSentence(m[1].trim(), 500)}.`)
    }
    for (const h of hits) {
      const text = cleanChunkText(h.text || '')
      if (!new RegExp(`\\b${escapedSubject}\\b`, 'i').test(text)) continue
      const extracted = extractFormulaFromText(text)
      if (extracted) return ensureSinglePeriod(`**Formula for ${capFirst(subject)}:** ${extracted}.`)
    }
    return `I could not find a formula for ${capFirst(subject)} in your documents.`
  }

  for (const h of hits) {
    if (h.metadata && h.metadata.measure) {
      const measureLower = (h.metadata.measure || '').toLowerCase().trim()
      const isExactMatch = measureLower === subjectLower
      const isPartialMatch = !isExactMatch && (
        new RegExp(`\\b${escapedSubject}\\b`, 'i').test(measureLower) ||
        new RegExp(`\\b${escapeRegex(measureLower)}\\b`, 'i').test(subjectLower) ||
        (subjectWords.length > 1 && subjectWords.every(w => measureLower.includes(w)))
      )
      if (isExactMatch || isPartialMatch) {
        const cap = capFirst(h.metadata.measure)
        const desc = h.metadata.description || ''
        if (desc) {
          return ensureSinglePeriod(`**${cap}** is defined as: ${trimToCompleteSentence(desc, 600)}.`)
        }
        const cleanText = cleanChunkText(h.text || '')
        const definedAsMatch = cleanText.match(/is defined as:\s*(.+?)(?:\.\s*Formula:|$)/is)
        if (definedAsMatch) {
          return ensureSinglePeriod(`**${cap}** is defined as: ${trimToCompleteSentence(definedAsMatch[1].trim(), 600)}.`)
        }
        return ensureSinglePeriod(`**${cap}:** ${trimToCompleteSentence(cleanText, 400)}.`)
      }
    }
  }

  const synthesisPattern = new RegExp(
    `${escapedSubject}[^\\n]{0,60}is defined as:\\s*([^.\\n]+(?:\\.[^.\\n]+)?)(?:\\.\\s*Formula:\\s*([^.\\n]+(?:\\.[^.\\n]+)?))?(?:\\.\\s*Additional Info:\\s*([^.\\n]+))?`,
    'im'
  )
  for (const h of hits) {
    const cleanText = cleanChunkText(h.text || '')
    const m = cleanText.match(synthesisPattern)
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
    const cleanText = cleanChunkText(h.text || '')
    for (const line of cleanText.split('\n')) {
      if (!new RegExp(`\\b${escapedSubject}\\b`, 'i').test(line)) continue
      const trimmedLine = line.trim()
      if (trimmedLine.length <= 20) continue
      if ((trimmedLine.match(/\|/g) || []).length > 2) continue
      if (/^===\s*Sheet:/.test(trimmedLine)) continue
      if (resolvedIntent === 'definition' && /formula|calculated as|computed as/i.test(trimmedLine)) continue
      const cleaned = trimmedLine.replace(/\(from\s+[A-Za-z\s]+\)/g, '').trim()
      if (cleaned.length > 15) matchingLines.push(cleaned)
    }
  }

  if (matchingLines.length > 0) {
    const joined = trimToCompleteSentence([...new Set(matchingLines)].slice(0, 3).join(' '), 600)
    return ensureSinglePeriod(`**${capFirst(subject)}:** ${joined}.`)
  }

  return `I could not find information about "${capFirst(subject)}" in your documents.`
}

module.exports = {
  extractSpreadsheet,
  retrieveChunksDD,
  buildSystemPromptDD,
  buildUserMessageDD,
  buildFallbackAnswerDD,
  fuzzyCorrectQuery,
  relaxedKeywordSearchDD,
  extractAllUrlsFromChunks,
}
