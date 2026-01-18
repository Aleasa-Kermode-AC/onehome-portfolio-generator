const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

// OpenAI API configuration
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

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
 * Generate AI progress summary for a learning area
 */
async function generateProgressSummary(area, evidenceList, childName) {
  if (!evidenceList || evidenceList.length === 0) {
    return null;
  }
  
  const evidenceSummary = evidenceList.map(e => 
    `- "${e.title}": ${e.description} (Engagement: ${e.engagement})`
  ).join('\n');
  
  const prompt = `Write a 2-3 sentence progress summary for ${childName}'s learning in ${area} based on this evidence:

${evidenceSummary}

The summary should:
- Highlight key skills demonstrated
- Note engagement levels and interests
- Use warm, strengths-based language
- Be suitable for a formal home education portfolio

Write only the summary paragraph, no headings or labels.`;

  return await callOpenAI(prompt, 200);
}

/**
 * Enhance parent assessment with AI based on evidence
 */
async function enhanceParentAssessment(domain, parentInput, childName, allEvidence) {
  if (!parentInput || parentInput.trim() === '') {
    return parentInput;
  }
  
  // Get relevant evidence snippets
  const evidenceSnippets = [];
  Object.values(allEvidence).forEach(evidenceList => {
    if (Array.isArray(evidenceList)) {
      evidenceList.forEach(e => {
        evidenceSnippets.push(`${e.title}: ${e.description}`);
      });
    }
  });
  
  const evidenceContext = evidenceSnippets.slice(0, 5).join('\n');
  
  const prompt = `A parent wrote this brief ${domain.toLowerCase()} assessment for their child ${childName}:

"${parentInput}"

Based on this parent input and the following learning evidence:
${evidenceContext}

Expand this into a professional 3-4 sentence assessment that:
- Keeps the parent's voice and observations
- Adds specific examples from the evidence where relevant
- Uses warm, strengths-based educational language
- Is suitable for a formal home education portfolio

Write only the enhanced assessment paragraph, no headings or labels.`;

  const enhanced = await callOpenAI(prompt, 250);
  return enhanced || parentInput; // Fall back to original if AI fails
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

  try {
    console.log('=== PORTFOLIO GENERATION REQUEST ===');
    
    let portfolioData = req.body;
    
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
        let cleanedString = portfolioData.futurePlans
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
        portfolioData.futurePlans = { overview: portfolioData.futurePlans };
      }
    }
    
    // === FIX 3: Parse progressAssessment if it's a string ===
    if (typeof portfolioData.progressAssessment === 'string') {
      try {
        portfolioData.progressAssessment = JSON.parse(portfolioData.progressAssessment);
        console.log('Parsed progressAssessment:', portfolioData.progressAssessment);
      } catch (e) {
        console.log('Could not parse progressAssessment as JSON, using as-is');
      }
    }
    
    // === FIX 4: Parse evidenceEntries if it's a string ===
    if (typeof portfolioData.evidenceEntries === 'string') {
      try {
        // Make.com sometimes sends as comma-separated JSON objects without array brackets
        let evidenceStr = portfolioData.evidenceEntries.trim();
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
        let outcomesStr = portfolioData.curriculumOutcomes.trim();
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
      
      portfolioData.evidenceEntries.forEach(entry => {
        // Get learning areas - could be array or string
        let areas = entry['Learning Areas'] || entry.learningAreas || [];
        if (typeof areas === 'string') {
          areas = areas.split(',').map(a => a.trim().replace(/"/g, ''));
        }
        if (!Array.isArray(areas)) {
          areas = [areas];
        }
        
        areas.forEach(area => {
          if (!area) return;
          
          // Skip numeric or invalid area names
          if (area === '0' || area === '1' || /^\d+$/.test(String(area))) {
            return;
          }
          
          // Normalise area name
          const normalizedArea = normalizeAreaName(area);
          
          if (!portfolioData.evidenceByArea[normalizedArea]) {
            portfolioData.evidenceByArea[normalizedArea] = [];
          }
          
          // Format the evidence entry
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
          
          portfolioData.evidenceByArea[normalizedArea].push({
            title: entry.Title || entry.title || 'Untitled',
            date: formatDate(entry.Date || entry.date),
            description: entry['What Happened?'] || entry.whatHappened || entry.description || '',
            engagement: entry['Child Engagement'] || entry.childEngagement || entry.engagement || '',
            matchedOutcomes: outcomeCodesString
          });
        });
      });
      
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
        ? JSON.parse(portfolioData.progressAssessment) 
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
      console.log('AI enhancements complete');
    } else {
      console.log('OpenAI API key not configured, skipping AI enhancements');
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
