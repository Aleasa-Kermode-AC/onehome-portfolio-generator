const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

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
        portfolioData.futurePlans = JSON.parse(portfolioData.futurePlans);
        console.log('Parsed futurePlans:', portfolioData.futurePlans);
      } catch (e) {
        console.log('Could not parse futurePlans as JSON, using as-is');
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
          
          // Normalise area name
          const normalizedArea = normalizeAreaName(area);
          
          if (!portfolioData.evidenceByArea[normalizedArea]) {
            portfolioData.evidenceByArea[normalizedArea] = [];
          }
          
          // Format the evidence entry
          portfolioData.evidenceByArea[normalizedArea].push({
            title: entry.Title || entry.title || 'Untitled',
            date: formatDate(entry.Date || entry.date),
            description: entry['What Happened?'] || entry.whatHappened || entry.description || '',
            engagement: entry['Child Engagement'] || entry.childEngagement || entry.engagement || '',
            matchedOutcomes: entry['Matched Outcomes 3'] || entry.matchedOutcomes || []
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
