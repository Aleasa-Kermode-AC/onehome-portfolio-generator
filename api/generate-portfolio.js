const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

// Disable automatic body parsing so we can handle control characters
export const config = {
  api: {
    bodyParser: false,
  },
};

// OpenAI API configuration
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

// OneHome Education Logo URL (hosted on GitHub or Vercel)
const LOGO_URL = 'https://raw.githubusercontent.com/Aleasa-Kermode-AC/onehome-portfolio-generator/main/assets/OneHomeEd_Logo.jpg';

/**
 * Fetch the OneHome Education logo
 */
async function fetchLogo() {
  try {
    const response = await fetch(LOGO_URL);
    if (!response.ok) {
      console.log('Could not fetch logo:', response.status);
      return null;
    }
    const arrayBuffer = await response.arrayBuffer();
    return Buffer.from(arrayBuffer);
  } catch (error) {
    console.log('Error fetching logo:', error.message);
    return null;
  }
}

/**
 * Fetch image from URL and return as buffer
 */
async function fetchImageAsBuffer(url) {
  try {
    // Set a timeout for image fetching (10 seconds)
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10000);
    
    const response = await fetch(url, { signal: controller.signal });
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      console.error('Failed to fetch image:', response.status);
      return null;
    }
    const arrayBuffer = await response.arrayBuffer();
    return Buffer.from(arrayBuffer);
  } catch (error) {
    console.error('Error fetching image:', error.message);
    return null;
  }
}

/**
 * Process attachments and fetch image data
 */
async function processAttachments(attachments) {
  if (!attachments || !Array.isArray(attachments) || attachments.length === 0) {
    return [];
  }
  
  const processedImages = [];
  
  // Limit to first 3 images per evidence to avoid timeout
  const limitedAttachments = attachments.slice(0, 3);
  
  for (const att of limitedAttachments) {
    try {
      // Check if it's an image
      const mimeType = att['MIME type'] || att.type || '';
      if (!mimeType.startsWith('image/')) {
        continue;
      }
      
      // Use thumbnail for smaller file size in document (Large size is good balance)
      const imageUrl = att.Thumbnails?.Large?.URL || att.URL || att.url;
      const width = att.Thumbnails?.Large?.Width || att.Width || 400;
      const height = att.Thumbnails?.Large?.Height || att.Height || 300;
      
      if (!imageUrl) continue;
      
      console.log(`Fetching image: ${att['File name'] || 'unknown'}`);
      const imageBuffer = await fetchImageAsBuffer(imageUrl);
      
      if (imageBuffer) {
        processedImages.push({
          buffer: imageBuffer,
          filename: att['File name'] || att.filename || 'image.jpg',
          width: Math.min(width, 400), // Cap width at 400px for document
          height: Math.min(height, 500), // Cap height at 500px
          mimeType: mimeType
        });
      }
    } catch (error) {
      console.error('Error processing attachment:', error.message);
      // Continue with other attachments instead of crashing
    }
  }
  
  return processedImages;
}

/**
 * Call OpenAI API to generate text
 */
async function callOpenAI(prompt, maxTokens = 300) {
  if (!OPENAI_API_KEY) {
    console.log('OpenAI API key not configured, skipping AI enhancement');
    return null;
  }
  
  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          {
            role: 'system',
            content: 'You are an experienced Australian home education assessor who writes warm, strengths-based, professional summaries for neurodivergent homeschooled children. Use Australian English spelling. Be concise but thorough. Focus on growth, engagement, and demonstrated learning.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        max_tokens: maxTokens,
        temperature: 0.7
      })
    });
    
    if (!response.ok) {
      console.error('OpenAI API error:', response.status, await response.text());
      return null;
    }
    
    const data = await response.json();
    return data.choices[0]?.message?.content?.trim() || null;
  } catch (error) {
    console.error('OpenAI API call failed:', error.message);
    return null;
  }
}

/**
 * Generate AI progress summary for a learning area - strictly subject-specific
 */
async function generateProgressSummary(area, evidenceList, childName) {
  if (!evidenceList || evidenceList.length === 0) {
    return null;
  }
  
  const evidenceSummary = evidenceList.map(e => 
    `- "${e.title}": ${e.description} (Engagement: ${e.engagement})`
  ).join('\n');
  
  const prompt = `Write a 2-3 sentence progress summary for ${childName}'s learning ONLY in ${area} based on this evidence:

${evidenceSummary}

CRITICAL RULES:
- ONLY describe skills and activities directly related to ${area}
- Do NOT mention skills from other subject areas (e.g., if this is English, don't mention maths/geography/science skills)
- Focus on ${area}-specific learning demonstrated in the evidence
- Use warm, strengths-based language
- Do NOT use asterisks or markdown formatting around titles or words
- Be suitable for a formal NSW home education portfolio

For reference:
- English: reading, writing, speaking, listening, comprehension, vocabulary, grammar, punctuation
- Mathematics: numbers, calculations, measurement, geometry, data, patterns
- Science & Technology: scientific inquiry, experiments, technology, design, natural world
- HSIE: history, geography, civics, society, environment, culture
- PDHPE: physical activity, health, wellbeing, relationships, safety, emotions
- Creative Arts: visual arts, music, drama, dance, creativity, artistic expression

Write only the summary paragraph, no headings or labels.`;

  return await callOpenAI(prompt, 200);
}

/**
 * Enhance parent assessment with AI based on evidence - domain specific
 */
async function enhanceParentAssessment(domain, parentInput, childName, allEvidence) {
  if (!parentInput || parentInput.trim() === '') {
    return parentInput;
  }
  
  // Get evidence relevant to this domain
  const domainKeywords = {
    'Cognitive Development': ['thinking', 'learning', 'question', 'understand', 'problem', 'reading', 'writing', 'calculating', 'analysing'],
    'Social Development': ['friend', 'peer', 'social', 'people', 'interact', 'play', 'share', 'communicate', 'relationship'],
    'Emotional Development': ['feeling', 'emotion', 'confidence', 'anxiety', 'self', 'regulate', 'cope', 'express'],
    'Physical Development': ['swim', 'run', 'motor', 'fitness', 'physical', 'movement', 'sport', 'exercise', 'coordination']
  };
  
  const keywords = domainKeywords[domain] || [];
  
  // Find relevant evidence snippets for this domain
  const relevantSnippets = [];
  Object.values(allEvidence).forEach(evidenceList => {
    if (Array.isArray(evidenceList)) {
      evidenceList.forEach(e => {
        const desc = (e.description || '').toLowerCase();
        const isRelevant = keywords.some(kw => desc.includes(kw));
        if (isRelevant) {
          relevantSnippets.push(`${e.title}: ${e.description}`);
        }
      });
    }
  });
  
  // Limit to 3 most relevant snippets
  const evidenceContext = relevantSnippets.slice(0, 3).join('\n');
  
  const prompt = `A parent wrote this brief ${domain.toLowerCase()} assessment for their child ${childName}:

"${parentInput}"

${evidenceContext ? `Relevant learning evidence for ${domain.toLowerCase()}:\n${evidenceContext}` : ''}

Expand this into a professional 2-3 sentence assessment that:
- Stays focused ONLY on ${domain.toLowerCase()} (do not mention unrelated learning areas)
- Keeps the parent's voice and observations as the foundation
- Only adds evidence examples if they directly relate to ${domain.toLowerCase()}
- Uses warm, strengths-based educational language
- Does NOT use asterisks or markdown formatting around any words
- Is suitable for a formal NSW home education portfolio

Write only the enhanced assessment paragraph, no headings or labels.`;

  const enhanced = await callOpenAI(prompt, 200);
  return enhanced || parentInput;
}

/**
 * Enhance future plans overview with AI
 */
async function enhanceFuturePlans(overview, goals, childName) {
  if (!overview || overview.trim() === '') {
    return { enhancedOverview: null, enhancedGoals: null };
  }
  
  const overviewPrompt = `A parent wrote this brief future learning plans overview for their child ${childName}:

"${overview}"

${goals ? `Their learning goals are: "${goals}"` : ''}

Expand this into a professional 2-3 sentence overview that:
- Maintains the parent's positive sentiment
- Adds educational context about child-led learning approaches
- Uses warm, strengths-based language suitable for NSW home education
- Does NOT use asterisks or markdown formatting

Write only the enhanced overview paragraph.`;

  const enhancedOverview = await callOpenAI(overviewPrompt, 150);
  
  return { 
    enhancedOverview: enhancedOverview || overview,
    enhancedGoals: null // Goals stay as bullet points from parent
  };
}

module.exports = async (req, res) => {
  // Set CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Handle preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Health check
  if (req.method === 'GET') {
    return res.status(200).json({ 
      status: 'healthy', 
      service: 'OneHome Education Portfolio Generator',
      version: '1.2.0'
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  /**
   * Sanitize string to remove control characters that break JSON
   */
  function sanitizeString(str) {
    if (typeof str !== 'string') return str;
    // Remove control characters (except newline, carriage return, tab)
    // Then normalize whitespace
    return str
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '') // Remove control chars
      .replace(/\r\n/g, '\n')  // Normalize line endings
      .replace(/\r/g, '\n')    // Normalize line endings
      .replace(/\t/g, ' ')     // Replace tabs with spaces
      .trim();
  }

  /**
   * Recursively sanitize all strings in an object
   */
  function sanitizeObject(obj) {
    if (obj === null || obj === undefined) return obj;
    if (typeof obj === 'string') return sanitizeString(obj);
    if (Array.isArray(obj)) return obj.map(item => sanitizeObject(item));
    if (typeof obj === 'object') {
      const sanitized = {};
      for (const [key, value] of Object.entries(obj)) {
        sanitized[key] = sanitizeObject(value);
      }
      return sanitized;
    }
    return obj;
  }

  /**
   * Sanitize raw JSON string before parsing
   */
  function sanitizeJsonString(jsonStr) {
    if (typeof jsonStr !== 'string') return jsonStr;
    // Remove control characters that break JSON parsing
    // Keep only valid JSON whitespace: space, tab, newline, carriage return
    return jsonStr.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  }

  /**
   * Read raw body from request
   */
  function getRawBody(req) {
    return new Promise((resolve, reject) => {
      let data = '';
      req.on('data', chunk => {
        data += chunk;
      });
      req.on('end', () => {
        resolve(data);
      });
      req.on('error', err => {
        reject(err);
      });
    });
  }

  try {
    console.log('=== PORTFOLIO GENERATION REQUEST ===');
    console.log('Request received at:', new Date().toISOString());
    
    let portfolioData;
    
    // Read and parse raw body with sanitization
    try {
      const rawBody = await getRawBody(req);
      console.log('Raw body length:', rawBody.length);
      const sanitizedBody = sanitizeJsonString(rawBody);
      portfolioData = JSON.parse(sanitizedBody);
    } catch (parseError) {
      console.error('Failed to parse request body:', parseError.message);
      return res.status(400).json({ error: 'Invalid JSON in request body', details: parseError.message });
    }
    
    // Sanitize all input data to remove any remaining control characters
    portfolioData = sanitizeObject(portfolioData);
    
    // Log raw input for debugging
    console.log('Raw parentname:', portfolioData.parentname);
    console.log('Raw parentName:', portfolioData.parentName);
    console.log('Raw futurePlans type:', typeof portfolioData.futurePlans);
    console.log('Raw evidenceEntries type:', typeof portfolioData.evidenceEntries);
    console.log('Raw curriculumOutcomes type:', typeof portfolioData.curriculumOutcomes);
    
    // === FIX 1: Handle parentName case sensitivity ===
    // Accept both 'parentname' and 'parentName'
    if (!portfolioData.parentName && portfolioData.parentname) {
      portfolioData.parentName = portfolioData.parentname;
    }
    console.log('Final parentName:', portfolioData.parentName);
    
    // === FIX 2: Parse futurePlans if it's a string ===
    if (typeof portfolioData.futurePlans === 'string') {
      try {
        // Handle the specific format from Make.com with curly braces and newlines
        let cleanedString = sanitizeJsonString(portfolioData.futurePlans)
          .replace(/\{\s+"/g, '{"')  // Fix leading whitespace after {
          .replace(/"\s+\}/g, '"}')  // Fix trailing whitespace before }
          .replace(/\\n/g, ' ')      // Replace escaped newlines with spaces
          .replace(/\n/g, ' ')       // Replace actual newlines with spaces
          .trim();
        
        portfolioData.futurePlans = JSON.parse(cleanedString);
        console.log('Parsed futurePlans successfully:', Object.keys(portfolioData.futurePlans));
      } catch (e) {
        console.log('Could not parse futurePlans as JSON:', e.message);
        console.log('Raw futurePlans:', portfolioData.futurePlans.substring(0, 200));
        // Keep as object with raw string as overview
        portfolioData.futurePlans = { overview: sanitizeString(portfolioData.futurePlans) };
      }
    }
    
    // === FIX 3: Parse progressAssessment if it's a string ===
    console.log('progressAssessment type:', typeof portfolioData.progressAssessment);
    console.log('progressAssessment raw:', JSON.stringify(portfolioData.progressAssessment)?.substring(0, 300));
    
    if (typeof portfolioData.progressAssessment === 'string') {
      try {
        portfolioData.progressAssessment = JSON.parse(sanitizeJsonString(portfolioData.progressAssessment));
        console.log('Parsed progressAssessment:', JSON.stringify(portfolioData.progressAssessment)?.substring(0, 300));
      } catch (e) {
        console.log('Could not parse progressAssessment as JSON:', e.message);
        // Try to extract values if it looks like a simple format
        portfolioData.progressAssessment = {};
      }
    }
    
    // Log final progressAssessment structure
    console.log('Final progressAssessment keys:', Object.keys(portfolioData.progressAssessment || {}));
    
    // === FIX 4: Parse evidenceEntries if it's a string ===
    if (typeof portfolioData.evidenceEntries === 'string') {
      try {
        // Make.com sometimes sends as comma-separated JSON objects without array brackets
        let evidenceStr = sanitizeJsonString(portfolioData.evidenceEntries).trim();
        if (!evidenceStr.startsWith('[')) {
          evidenceStr = '[' + evidenceStr + ']';
        }
        portfolioData.evidenceEntries = JSON.parse(evidenceStr);
        console.log('Parsed evidenceEntries count:', portfolioData.evidenceEntries.length);
      } catch (e) {
        console.log('Could not parse evidenceEntries:', e.message);
        portfolioData.evidenceEntries = [];
      }
    }
    
    // Ensure evidenceEntries is an array
    if (!Array.isArray(portfolioData.evidenceEntries)) {
      portfolioData.evidenceEntries = [];
    }
    
    // === FIX 5: Parse curriculumOutcomes if it's a string ===
    if (typeof portfolioData.curriculumOutcomes === 'string') {
      try {
        let outcomesStr = sanitizeJsonString(portfolioData.curriculumOutcomes).trim();
        if (!outcomesStr.startsWith('[')) {
          outcomesStr = '[' + outcomesStr + ']';
        }
        portfolioData.curriculumOutcomes = JSON.parse(outcomesStr);
        console.log('Parsed curriculumOutcomes count:', portfolioData.curriculumOutcomes.length);
      } catch (e) {
        console.log('Could not parse curriculumOutcomes:', e.message);
        portfolioData.curriculumOutcomes = [];
      }
    }
    
    // Ensure curriculumOutcomes is an array
    if (!Array.isArray(portfolioData.curriculumOutcomes)) {
      portfolioData.curriculumOutcomes = [];
    }
    
    // === FIX 6: Build evidenceByArea from evidenceEntries ===
    // Group evidence by learning area for the document sections
    if (portfolioData.evidenceEntries.length > 0) {
      
      console.log('Building evidenceByArea from evidenceEntries...');
      portfolioData.evidenceByArea = {};
      
      // Process each evidence entry - using for...of to support await
      for (const entry of portfolioData.evidenceEntries) {
        // Get learning areas - could be array or string
        let areas = entry['Learning Areas'] || entry.learningAreas || [];
        if (typeof areas === 'string') {
          areas = areas.split(',').map(a => a.trim().replace(/"/g, ''));
        }
        if (!Array.isArray(areas)) {
          areas = [areas];
        }
        
        // Process attachments once per entry (not per area)
        let processedAttachments = [];
        try {
          const rawAttachments = entry.Attachments || entry.attachments || [];
          if (Array.isArray(rawAttachments) && rawAttachments.length > 0) {
            processedAttachments = await processAttachments(rawAttachments);
            console.log(`Processed ${processedAttachments.length} images for "${entry.Title || 'Untitled'}"`);
          }
        } catch (attachError) {
          console.error('Error processing attachments, skipping images:', attachError.message);
          processedAttachments = [];
        }
        
        // Get outcome codes - could be array (from Make.com) or string (from rollup)
        let outcomeCodesData = entry['Outcome Code Rollup (from Matched Outcomes 3)'] || 
                               entry['Outcome Code Rollup'] ||
                               entry['Outcome Codes Text'] ||
                               entry.outcomeCodesText ||
                               entry['Matched Outcomes 3'] || 
                               entry.matchedOutcomes || [];
        
        // If it's an array of strings, join them; if it's already a string, use as-is
        let outcomeCodesString = '';
        if (Array.isArray(outcomeCodesData)) {
          outcomeCodesString = outcomeCodesData.join(', ');
        } else if (typeof outcomeCodesData === 'string') {
          outcomeCodesString = outcomeCodesData;
        }
        
        // Add to each learning area
        for (const area of areas) {
          if (!area) continue;
          
          // Skip numeric or invalid area names
          if (area === '0' || area === '1' || /^\d+$/.test(String(area))) {
            continue;
          }
          
          // Normalise area name
          const normalizedArea = normalizeAreaName(area);
          
          if (!portfolioData.evidenceByArea[normalizedArea]) {
            portfolioData.evidenceByArea[normalizedArea] = [];
          }
          
          portfolioData.evidenceByArea[normalizedArea].push({
            title: entry.Title || entry.title || 'Untitled',
            date: formatDate(entry.Date || entry.date),
            description: entry['What Happened?'] || entry.whatHappened || entry.description || '',
            engagement: entry['Child Engagement'] || entry.childEngagement || entry.engagement || '',
            matchedOutcomes: outcomeCodesString,
            attachments: processedAttachments
          });
        }
      }
      
      console.log('Built evidenceByArea with areas:', Object.keys(portfolioData.evidenceByArea));
      console.log('Evidence counts per area:', Object.fromEntries(
        Object.entries(portfolioData.evidenceByArea).map(([k, v]) => [k, Array.isArray(v) ? v.length : 'not array'])
      ));
    } else {
      portfolioData.evidenceByArea = {};
    }
    
    // === FIX 7: Build learningAreaOverviews from curriculumOutcomes ===
    if (portfolioData.curriculumOutcomes.length > 0 &&
        (!portfolioData.learningAreaOverviews || Object.keys(portfolioData.learningAreaOverviews).length === 0)) {
      
      console.log('Building learningAreaOverviews from curriculumOutcomes...');
      portfolioData.learningAreaOverviews = {};
      
      // Group outcomes by learning area
      const outcomesByArea = {};
      portfolioData.curriculumOutcomes.forEach(outcome => {
        const area = normalizeAreaName(outcome['Learning Area'] || outcome.learningArea || 'Other');
        if (!outcomesByArea[area]) {
          outcomesByArea[area] = [];
        }
        outcomesByArea[area].push(outcome);
      });
      
      // Create overview for each area
      Object.entries(outcomesByArea).forEach(([area, outcomes]) => {
        const evidenceForArea = portfolioData.evidenceByArea[area] || [];
        const hasEvidence = evidenceForArea.length > 0;
        
        // Build stage statement from outcomes
        const outcomeDescriptions = outcomes
          .slice(0, 5) // Take first 5 for the overview
          .map(o => o['Outcome Description'] || o.outcomeDescription || '')
          .filter(d => d.length > 0)
          .join('; ');
        
        portfolioData.learningAreaOverviews[area] = {
          stageStatement: `In ${area}, Stage 2 students work towards outcomes including: ${outcomeDescriptions}`,
          progressSummary: hasEvidence 
            ? `${evidenceForArea.length} learning evidence entries documented for this area.`
            : '',
          outcomes: outcomes.map(o => ({
            code: o['Outcome Title'] || o.outcomeTitle || o.code || '',
            description: o['Outcome Description'] || o.outcomeDescription || o.description || ''
          }))
        };
      });
      
      console.log('Built learningAreaOverviews for areas:', Object.keys(portfolioData.learningAreaOverviews));
    }
    
    // Validate required fields
    if (!portfolioData.childName || !portfolioData.yearLevel) {
      return res.status(400).json({ 
        success: false, 
        error: 'Missing required fields: childName and yearLevel are required' 
      });
    }
    
    console.log('Generating portfolio for:', portfolioData.childName);
    console.log('Evidence entries:', portfolioData.evidenceEntries?.length || 0);
    console.log('Curriculum outcomes:', portfolioData.curriculumOutcomes?.length || 0);
    console.log('Evidence by area:', Object.keys(portfolioData.evidenceByArea || {}));
    console.log('Learning area overviews:', Object.keys(portfolioData.learningAreaOverviews || {}));
    
    // === AI ENHANCEMENTS ===
    if (OPENAI_API_KEY) {
      console.log('Starting AI enhancements...');
      
      // 1. Generate AI progress summaries for each learning area
      const aiProgressSummaries = {};
      for (const [area, evidenceList] of Object.entries(portfolioData.evidenceByArea || {})) {
        if (Array.isArray(evidenceList) && evidenceList.length > 0) {
          console.log(`Generating AI summary for ${area}...`);
          const summary = await generateProgressSummary(area, evidenceList, portfolioData.childName);
          if (summary) {
            aiProgressSummaries[area] = summary;
            console.log(`AI summary generated for ${area}`);
          }
        }
      }
      portfolioData.aiProgressSummaries = aiProgressSummaries;
      
      // 2. Enhance parent assessments
      const parsedAssessment = typeof portfolioData.progressAssessment === 'string' 
        ? (() => { try { return JSON.parse(sanitizeJsonString(portfolioData.progressAssessment)); } catch(e) { return {}; } })()
        : (portfolioData.progressAssessment || {});
      
      const enhancedAssessment = {};
      
      if (parsedAssessment.cognitive) {
        console.log('Enhancing cognitive assessment...');
        enhancedAssessment.cognitive = await enhanceParentAssessment(
          'Cognitive Development', 
          parsedAssessment.cognitive, 
          portfolioData.childName,
          portfolioData.evidenceByArea
        );
      }
      
      if (parsedAssessment.social) {
        console.log('Enhancing social assessment...');
        enhancedAssessment.social = await enhanceParentAssessment(
          'Social Development', 
          parsedAssessment.social, 
          portfolioData.childName,
          portfolioData.evidenceByArea
        );
      }
      
      if (parsedAssessment.emotional) {
        console.log('Enhancing emotional assessment...');
        enhancedAssessment.emotional = await enhanceParentAssessment(
          'Emotional Development', 
          parsedAssessment.emotional, 
          portfolioData.childName,
          portfolioData.evidenceByArea
        );
      }
      
      if (parsedAssessment.physical) {
        console.log('Enhancing physical assessment...');
        enhancedAssessment.physical = await enhanceParentAssessment(
          'Physical Development', 
          parsedAssessment.physical, 
          portfolioData.childName,
          portfolioData.evidenceByArea
        );
      }
      
      portfolioData.enhancedProgressAssessment = enhancedAssessment;
      
      // 3. Enhance future plans overview
      const parsedFuturePlans = typeof portfolioData.futurePlans === 'string' 
        ? (() => { try { return JSON.parse(sanitizeJsonString(portfolioData.futurePlans)); } catch(e) { return { overview: sanitizeString(portfolioData.futurePlans) }; } })()
        : (portfolioData.futurePlans || {});
      
      if (parsedFuturePlans.overview) {
        console.log('Enhancing future plans overview...');
        const { enhancedOverview } = await enhanceFuturePlans(
          parsedFuturePlans.overview,
          parsedFuturePlans.goals,
          portfolioData.childName
        );
        if (enhancedOverview) {
          portfolioData.enhancedFuturePlansOverview = enhancedOverview;
        }
      }
      
      console.log('AI enhancements complete');
    } else {
      console.log('OpenAI API key not configured, skipping AI enhancements');
    }
    
    // Fetch the logo
    console.log('Fetching logo...');
    const logoBuffer = await fetchLogo();
    if (logoBuffer) {
      portfolioData.logoBuffer = logoBuffer;
      console.log('Logo fetched successfully');
    } else {
      console.log('Logo not available, continuing without logo');
    }
    
    // Generate the portfolio document
    const doc = generatePortfolio(portfolioData);
    
    // Convert to buffer
    const buffer = await Packer.toBuffer(doc);
    
    console.log('Portfolio generated successfully, size:', buffer.length, 'bytes');
    
    // Create safe filename
    const safeName = (portfolioData.childName || 'Child').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 50);
    const safePeriod = (portfolioData.reportingPeriod || 'Portfolio').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 30);
    const filename = `${safeName}-Portfolio-${safePeriod}.docx`;
    
    // Upload to Vercel Blob
    const blob = await put(`portfolios/${filename}`, buffer, {
      access: 'public',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    
    console.log('File uploaded to Vercel Blob:', blob.url);
    
    // Return success with URL
    return res.status(200).json({
      success: true,
      filename: filename,
      url: blob.url,
      fileSize: buffer.length
    });

  } catch (error) {
    console.error('Error generating portfolio:', error);
    return res.status(500).json({ 
      success: false, 
      error: error.message,
      stack: error.stack
    });
  }
};

// Helper function to normalise learning area names
function normalizeAreaName(area) {
  if (!area) return 'Other';
  
  const areaStr = String(area).trim();
  
  // Map various names to standard names
  const mappings = {
    'english': 'English',
    'mathematics': 'Mathematics',
    'maths': 'Mathematics',
    'math': 'Mathematics',
    'science & technology': 'Science & Technology',
    'science and technology': 'Science & Technology',
    'science': 'Science & Technology',
    'hsie': 'HSIE',
    'hsie (history, geography etc)': 'HSIE',
    'history': 'HSIE',
    'geography': 'HSIE',
    'pdhpe': 'PDHPE',
    'pdhpe (health, physical education)': 'PDHPE',
    'health': 'PDHPE',
    'pe': 'PDHPE',
    'creative arts': 'Creative Arts',
    'art': 'Creative Arts',
    'music': 'Creative Arts',
    'drama': 'Creative Arts',
    'dance': 'Creative Arts'
  };
  
  const lowerArea = areaStr.toLowerCase();
  return mappings[lowerArea] || areaStr;
}

// Helper function to format dates
function formatDate(dateInput) {
  if (!dateInput) return 'Date not specified';
  
  try {
    const date = new Date(dateInput);
    if (isNaN(date.getTime())) return String(dateInput);
    
    return date.toLocaleDateString('en-AU', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  } catch (e) {
    return String(dateInput);
  }
}
