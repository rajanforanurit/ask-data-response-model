'use strict'

const INTENTS = [
  {
    id: 'greeting',
    patterns: [
      /^\s*(hi|hello|hey|hii|helo|howdy|greetings|good\s*(morning|afternoon|evening|day))\s*[!.,]*\s*$/i,
      /^(hi|hello|hey)\s+there\s*[!.,]*$/i,
      /^(sup|what'?s\s+up|yo)\s*[!.,]*$/i,
      /^(hiya|heya|helo+|h+i+)\s*[!.,]*$/i,
      /^good\s*(morning|afternoon|evening|day)\s*[!.,]*$/i,
      /^how\s+are\s+you\s*[!?,.]*/i,
      /^how\s+r\s+u\s*[!?,.]*/i,
      /^how\s+are\s+u\s*[!?,.]*/i,
      /^hru\s*[!?,.]*/i,
      /^are\s+you\s+there\s*[!?,.]*/i,
      /^you\s+there\s*[!?,.]*/i,
      /^anyone\s+there\s*[!?,.]*/i,
      /^is\s+anyone\s+there\s*[!?,.]*/i,
      /^hello\s*[!?,.]*/i,
      /^hi\s*[!?,.]*/i,
      /^hey\s*[!?,.]*/i,
    ],
    responses: [
      'Hello! How can I assist you with your data today?',
      'Hi there! Ask me anything related to your enterprise data or reports.',
      'Hey! I\'m Ask Data, your AI assistant. What insights can I help you uncover?',
      'Hello! Ready to help you explore your enterprise data. What would you like to know?',
    ],
  },
  {
    id: 'identity_builder',
    patterns: [
      /who\s+(built|created|developed|made|designed|coded|engineered)\s+(this|it|you|the\s+(tool|app|chatbot|assistant|platform|system|visual|dashboard))/i,
      /who\s+is\s+behind\s+(this|the\s+(tool|platform|system|visual|chatbot|assistant))/i,
      /who\s+owns\s+(this|the\s+(tool|platform|system|visual|chatbot|assistant))/i,
      /who\s+maintains\s+(this|the\s+(tool|platform|system|visual|chatbot|assistant|dashboard))/i,
      /which\s+company\s+(made|built|created|developed|designed|owns)\s+(this|it|the\s+(tool|app|platform))/i,
      /what\s+company\s+(made|built|created|developed|is\s+behind)\s+(this|it)/i,
      /who\s+(designed|created)\s+this\s+(dashboard|visual|tool|app|platform|system)/i,
      /who\s+developed\s+this/i,
      /who\s+built\s+this/i,
      /who\s+made\s+this/i,
      /who\s+is\s+the\s+(developer|creator|author|maker|owner|builder)\s+of\s+(this|it|ask\s*data)/i,
      /who\s+(wrote|programmed|launched)\s+(this|it|ask\s*data)/i,
      /which\s+(team|group|org|organization)\s+(built|made|created|developed)\s+(this|it)/i,
    ],
    responses: [
      'Ask Data is an adaptive AI assistant developed by Anurit Innovation, designed to deliver fast, contextual insights from enterprise data.\n\nLearn more at https://www.anuritinnovation.com/',
    ],
  },
  {
    id: 'identity_self',
    patterns: [
      /what\s+(is\s+)?your\s+name/i,
      /who\s+are\s+you/i,
      /what\s+are\s+you/i,
      /tell\s+me\s+about\s+yourself/i,
      /introduce\s+yourself/i,
      /what\s+(is|are)\s+(this|the\s+(tool|assistant|chatbot|app|platform))/i,
      /what('s| is) ask\s*data/i,
      /are\s+you\s+an\s+(ai|bot|assistant|chatbot)/i,
      /what\s+kind\s+of\s+(ai|bot|assistant|tool)\s+are\s+you/i,
      /describe\s+yourself/i,
      /what\s+do\s+you\s+do/i,
    ],
    responses: [
      'I\'m Ask Data — an adaptive AI assistant developed by Anurit Innovation, designed to deliver fast, contextual insights from your enterprise data and Power BI reports.\n\nLearn more at https://www.anuritinnovation.com/',
    ],
  },
  {
    id: 'company_anurit',
    patterns: [
      /who\s+(is|are)\s+anurit(\s+innovation)?/i,
      /what\s+is\s+anurit(\s+innovation)?/i,
      /tell\s+me\s+about\s+anurit(\s+innovation)?/i,
      /about\s+anurit(\s+innovation)?/i,
      /anurit\s+innovation\s+(company|team|services|products|website|contact)/i,
      /anurit\s+(company|team|services|products|website|contact)/i,
      /what\s+does\s+anurit(\s+innovation)?\s+do/i,
      /anurit\s+innovation\s+kya\s+hai/i,
    ],
    responses: [
      'Anurit Innovation is a technology company specializing in intelligent enterprise data solutions and AI-powered analytics. They build tools like Ask Data — adaptive AI assistants embedded in platforms like Power BI — to help organizations make faster, smarter decisions from their data.\n\nVisit: https://www.anuritinnovation.com/',
    ],
  },
  {
    id: 'capabilities_general',
    patterns: [
      /what\s+can\s+you\s+do/i,
      /what\s+are\s+your\s+(capabilities|features|functions|abilities)/i,
      /what\s+do\s+you\s+(support|offer|provide|handle)/i,
      /how\s+can\s+you\s+help(\s+me)?/i,
      /what\s+kind\s+of\s+questions\s+can\s+i\s+(ask|use)/i,
      /what\s+questions\s+can\s+i\s+ask/i,
      /what\s+type\s+of\s+questions/i,
      /what\s+topics\s+do\s+you\s+(cover|handle|support)/i,
      /what\s+can\s+i\s+(ask|use\s+you\s+for)/i,
      /what\s+(features|things)\s+do\s+you\s+(have|support|offer)/i,
      /how\s+do\s+you\s+work/i,
      /what\s+is\s+your\s+(purpose|goal|function|role)/i,
    ],
    responses: [
      'Here\'s what I can help you with:\n\n- Data Analysis — Ask questions about your business metrics, KPIs, and trends.\n- Report Summarization — Get concise summaries of complex Power BI reports.\n- Context-Aware Querying — Retrieve specific data points from enterprise datasets.\n- Enterprise Insights — Uncover patterns, outliers, and actionable intelligence.\n- Power BI Integration — Natively embedded in your Power BI dashboards.\n\nJust type a business question and I\'ll find the answer from your data!',
    ],
  },
  {
    id: 'capabilities_powerbi',
    patterns: [
      /do\s+you\s+support\s+power\s*bi/i,
      /are\s+you\s+(a\s+)?power\s*bi/i,
      /power\s*bi\s+(support|integration|compatible|plugin|visual|tool|add-?in)/i,
      /is\s+this\s+(a\s+)?power\s*bi\s+(visual|tool|plugin|add-?in|report|dashboard)/i,
      /works?\s+with\s+power\s*bi/i,
      /power\s*bi\s+custom\s+visual/i,
      /embedded\s+in\s+power\s*bi/i,
      /can\s+i\s+use\s+(you|this)\s+(in|with|inside)\s+power\s*bi/i,
    ],
    responses: [
      'Yes! Ask Data is built as a native Power BI Custom Visual. It embeds directly into your Power BI dashboards and reports, giving your team the ability to query enterprise data using natural language — no SQL or technical knowledge required.\n\nPowered by Anurit Innovation: https://www.anuritinnovation.com/',
    ],
  },
  {
    id: 'capabilities_analysis',
    patterns: [
      /can\s+you\s+(analyze|analyse)\s+data/i,
      /can\s+you\s+do\s+(data\s+)?(analysis|analytics)/i,
      /do\s+you\s+(analyze|analyse)\s+data/i,
      /are\s+you\s+able\s+to\s+analyze/i,
      /can\s+you\s+(find|detect|identify)\s+(trends?|patterns?|outliers?|insights?)/i,
      /can\s+you\s+(compare|show\s+me)\s+(data|numbers|metrics|kpis?)/i,
      /can\s+you\s+give\s+me\s+(insights?|analysis|analytics)/i,
    ],
    responses: [
      'Absolutely! I\'m designed to analyze enterprise data in real time. You can ask things like:\n\n- "What were the top 5 products by revenue last quarter?"\n- "Show me the trend in customer acquisition over 6 months"\n- "Which region had the highest churn rate?"\n\nI retrieve answers directly from your connected data sources — no manual filtering needed.',
    ],
  },
  {
    id: 'capabilities_summarize',
    patterns: [
      /can\s+you\s+(summarize|summarise)\s+(reports?|dashboards?|data)/i,
      /can\s+you\s+summarize\s+data/i,
      /do\s+you\s+summarize/i,
      /report\s+summarization/i,
      /data\s+summary/i,
      /can\s+you\s+give\s+me\s+a\s+summary/i,
      /can\s+you\s+(summarize|summarise|recap|overview)/i,
      /can\s+you\s+(condense|digest|distill)\s+(this|data|reports?)/i,
    ],
    responses: [
      'Yes! I can summarize reports and data for you. Just ask something like:\n\n- "Summarize this month\'s sales report"\n- "Give me an overview of operational KPIs"\n- "What are the key takeaways from the Q3 dashboard?"\n\nI\'ll distill the most important insights from your data into a clear, concise response.',
    ],
  },
  {
    id: 'capabilities_database',
    patterns: [
      /are\s+you\s+connected\s+to\s+(a\s+)?(database|db|data\s*source|backend)/i,
      /do\s+you\s+(connect|link)\s+to\s+a\s+(database|db)/i,
      /is\s+this\s+connected\s+to/i,
      /where\s+does\s+(your|the)\s+data\s+come\s+from/i,
      /what\s+(data|database|source)\s+are\s+you\s+(using|connected\s+to|pulling\s+from)/i,
      /do\s+you\s+(use|access)\s+(real|live|actual)\s+data/i,
      /how\s+do\s+you\s+(get|fetch|retrieve|access)\s+(the\s+)?data/i,
      /is\s+(the\s+)?data\s+(live|real.?time|updated|current)/i,
    ],
    responses: [
      'Yes! Ask Data connects to your enterprise data sources through a secure RAG (Retrieval-Augmented Generation) pipeline. Your data is indexed and made searchable via vector embeddings, queries retrieve only the most relevant context, and responses are grounded in your actual enterprise data — not generic AI guesses. Your data stays within your organization\'s security perimeter.',
    ],
  },
  {
    id: 'thanks',
    patterns: [
      /^(thank\s*you|thanks|thank\s+u|thx|ty|cheers|great|awesome|perfect|nice|cool|got\s+it|ok(ay)?|sounds\s+good)\s*[!.,]*$/i,
      /^(thanks\s+a\s+lot|thank\s+you\s+so\s+much|much\s+appreciated)\s*[!.,]*$/i,
      /^(brilliant|wonderful|excellent|superb|amazing|fantastic)\s*[!.,]*$/i,
      /^(sure|alright|noted|understood|makes\s+sense)\s*[!.,]*$/i,
    ],
    responses: [
      'You\'re welcome! Feel free to ask me anything about your data.',
      'Happy to help! Let me know if you have any other questions.',
      'Anytime! Ask me anything about your enterprise data or reports.',
      'Glad I could help! What else would you like to explore?',
    ],
  },
  {
    id: 'help',
    patterns: [
      /^\s*help\s*[!?,]*\s*$/i,
      /i\s+need\s+help/i,
      /how\s+does\s+this\s+(work|tool\s+work)/i,
      /how\s+do\s+i\s+(use|start|begin)/i,
      /getting\s+started/i,
      /how\s+to\s+(use|start|begin|get\s+started)/i,
      /can\s+you\s+guide\s+me/i,
      /show\s+me\s+how\s+to\s+(use|start)/i,
      /i('m|\s+am)\s+(new|lost|confused|not\s+sure)/i,
    ],
    responses: [
      'Welcome to Ask Data! Here\'s how to get started:\n\n1. Ask a business question — type any question about your data in plain English.\n2. Be specific — the more context you provide, the more accurate the answer.\n3. Explore insights — ask for trends, comparisons, summaries, or specific metrics.\n\nExample questions:\n- "What is the total revenue for Q2 2026?"\n- "Which products had declining sales last month?"\n- "Compare performance across regions"\n\nLearn more: https://www.anuritinnovation.com/',
    ],
  },
  {
    id: 'irrelevant_general',
    patterns: [
      /^(bored|boredom|nothing)\s*[!.,]*$/i,
      /tell\s+me\s+a\s+(joke|story|fun\s+fact)/i,
      /what('s|\s+is)\s+the\s+weather/i,
      /what\s+(time|day|date)\s+is\s+it/i,
      /play\s+(music|song|video)/i,
      /can\s+you\s+(sing|dance|draw|paint)/i,
      /do\s+you\s+(dream|sleep|eat|feel)/i,
      /are\s+you\s+(human|real|alive|sentient|conscious)/i,
      /do\s+you\s+have\s+(feelings|emotions|a\s+soul)/i,
      /what\s+is\s+(life|love|happiness|the\s+meaning)/i,
    ],
    responses: [
      'I\'m Ask Data — purpose-built for enterprise data queries. I\'m not able to help with that, but I can answer questions about your business data, KPIs, reports, and more. What would you like to explore?',
      'That\'s outside my area! I specialize in enterprise data analysis and Power BI insights. Try asking me a business question — like revenue trends, top products, or regional performance.',
    ],
  },
]

const HARD_OUT_OF_SCOPE_PATTERNS = [
  /^who\s+is\s+the\s+(first|second|third|current|former|last|next|prime\s+minister|president|king|queen|emperor|governor|mayor|ceo|founder|chancellor)\s+of\s+(india|china|usa|uk|russia|france|germany|japan|australia|canada|pakistan|bangladesh|nepal|brazil|mexico|egypt|nigeria|kenya|south\s+africa|argentina|iran|iraq|saudi|uae|israel|turkey|greece|italy|spain|portugal|poland|ukraine|sweden|norway|denmark|finland|switzerland|austria|belgium|netherlands|new\s+zealand|south\s+korea|north\s+korea|philippines|vietnam|indonesia|malaysia|singapore)\b/i,
  /^what\s+is\s+the\s+(capital|currency|official\s+language|population|national\s+anthem|national\s+flag|national\s+animal|independence\s+day)\s+of\s+(india|china|usa|uk|russia|france|germany|japan|australia|canada|pakistan|bangladesh|nepal|brazil|mexico|egypt|nigeria|south\s+africa|argentina|iran|iraq|saudi\s+arabia|turkey|israel|greece|italy|spain|sweden|norway|denmark|finland|switzerland|austria|belgium|netherlands|new\s+zealand|south\s+korea|north\s+korea|philippines|vietnam|indonesia|malaysia|singapore|bangladesh|nepal|sri\s+lanka|afghanistan|myanmar|thailand)\b/i,
  /^(what|who|where|when|why|how)\s+(is|are|was|were|did)\s+(photosynthesis|mitosis|meiosis|osmosis|evaporation|condensation|precipitation|erosion|tectonic|radioactive|electromagnetic)\b/i,
  /^(what\s+is|who\s+is|where\s+is|when\s+was|why\s+did|how\s+did)\s+(world\s+war|cold\s+war|vietnam\s+war|korean\s+war|gulf\s+war|civil\s+war|holocaust|slavery|colonialism|imperialism|feudalism|fascism|nazism|communism)\b/i,
  /^tell\s+me\s+a\s+(joke|riddle|story|poem|fun\s+fact)\b/i,
  /^(play|sing|draw|paint|dance)\b/i,
  /^what\s+is\s+(the\s+meaning\s+of\s+life|love|happiness|god|religion|spirituality)\b/i,
  /^(latest|current|today\'?s?)\s+(news|headlines|stock\s+price|cricket\s+score|sports\s+score|weather)\b/i,
  /^who\s+(won|is\s+winning|will\s+win)\s+(the\s+)?(ipl|icc|world\s+cup|olympics|fifa|nba|nfl|premier\s+league|champions\s+league|formula\s+1|f1)\b/i,
]

function isDefinitionOrDocQuery(query) {
  const definitionPrefixes = /^(define|what\s+is|what\s+are|explain|tell\s+me\s+about|describe|how\s+(is|are|do\s+you\s+calculate)|what\s+is\s+the\s+(formula|definition|meaning|calculation)\s+(for|of)|show\s+me|list|find|get)\s+/i
  const words = query.trim().split(/\s+/)
  if (definitionPrefixes.test(query) && words.length <= 8) return true
  if (words.length <= 4) return true
  return false
}

function isOutOfScope(query) {
  const q = query.trim()

  if (isDefinitionOrDocQuery(q)) return false

  for (const pattern of HARD_OUT_OF_SCOPE_PATTERNS) {
    if (pattern.test(q)) return true
  }

  return false
}

function pickRandom(arr) {
  return arr[Math.floor(Math.random() * arr.length)]
}

const OUT_OF_SCOPE_RESPONSE = 'I can only answer questions based on your uploaded documents. That topic isn\'t covered in them. Please ask something related to your enterprise data, KPIs, or reports.'

function resolveIntent(query) {
  if (!query || !query.trim()) return null
  const normalized = query.trim().replace(/\s+/g, ' ')
  for (const intent of INTENTS) {
    for (const pattern of intent.patterns) {
      if (pattern.test(normalized)) {
        return { intent: intent.id, response: pickRandom(intent.responses) }
      }
    }
  }
  if (isOutOfScope(normalized)) {
    return { intent: 'out_of_scope', response: OUT_OF_SCOPE_RESPONSE }
  }
  return null
}

module.exports = { resolveIntent, isOutOfScope }
