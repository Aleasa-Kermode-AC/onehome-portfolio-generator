const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

// Helper to normalize arrays from Make.com
function toArray(val) {
  if (Array.isArray(val)) return val;
  if (!val) return [];
  if (typeof val === 'object') return Object.values(val).filter(v => v !== null && v !== undefined);
  return [];
}

// Normalize learning area names to standard format
function normalizeAreaName(area) {
  if (!area) return 'Other';
  const areaStr = String(area).trim();
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
  return mappings[areaStr.toLowerCase()] || areaStr;
}

// Build evidenceByArea object from flat evidenceEntries array
function buildEvidenceByArea(evidenceEntries) {
  const byArea = {};

  evidenceEntries.forEach(entry => {
    // Learning areas can be an array, a string, or nested
    let areas = [];

    if (entry['Learning Areas'] || entry.learningAreas) {
      const raw = entry['Learning Areas'] || entry.learningAreas;
      if (Array.isArray(raw)) {
        areas = raw;
      } else if (typeof raw === 'string') {
        areas = raw.split(',').map(a => a.trim());
      }
    }

    // Also check Areas field
    if (entry.Areas) {
      const raw = entry.Areas;
      if (Array.isArray(raw)) {
        areas = [...areas, ...raw];
      } else if (typeof raw === 'string') {
        areas = [...areas, ...raw.split(',').map(a => a.trim())];
      }
    }

    // Normalize and deduplicate areas
    areas = [...new Set(areas.map(normalizeAreaName).filter(a => a && a !== 'Other'))];

    // If no areas found, put in Other
    if (areas.length === 0) areas = ['Other'];

    // Build the evidence object for this entry
    const evidenceObj = {
      title: entry.Title || entry.title || 'Untitled',
      date: entry.Date || entry.date || '',
      description: entry['What Happened?'] || entry.whatHappened || entry.description || '',
      engagement: entry['Child Engagement'] || entry.childEngagement || entry.engagement || '',
      matchedOutcomes: entry['Outcome Code Rollup (from Matched Outcomes 3)'] || 
                       entry['Matched Outcomes 3'] ||
                       entry.matchedOutcomes || [],
      attachments: entry.Attachments || entry.attachments || []
    };

    // Add to each learning area
    areas.forEach(area => {
      if (!byArea[area]) byArea[area] = [];
      byArea[area].push(evidenceObj);
    });
  });

  return byArea;
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method === 'GET') {
    return res.status(200).json({
      status: 'healthy',
      service: 'OneHome Education Portfolio Generator',
      version: '3.1.0',
      mode: 'synchronous',
    });
  }

  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const portfolioData = req.body;

  if (!portfolioData || !portfolioData.childName) {
    return res.status(400).json({ success: false, error: 'Missing required field: childName' });
  }

  try {
    console.log('Portfolio generation started for:', portfolioData.childName);

    // Normalize arrays
    portfolioData.curriculumOutcomes = toArray(portfolioData.curriculumOutcomes);
    portfolioData.evidenceEntries = toArray(portfolioData.evidenceEntries);
    portfolioData.learningAreaOverviews = toArray(portfolioData.learningAreaOverviews);

    // Build evidenceByArea from evidenceEntries if not already populated
    const existingByArea = portfolioData.evidenceByArea;
    const hasValidByArea = existingByArea && 
                           typeof existingByArea === 'object' && 
                           !Array.isArray(existingByArea) && 
                           Object.keys(existingByArea).length > 0;

    if (!hasValidByArea && portfolioData.evidenceEntries.length > 0) {
      console.log('Building evidenceByArea from', portfolioData.evidenceEntries.length, 'evidence entries');
      portfolioData.evidenceByArea = buildEvidenceByArea(portfolioData.evidenceEntries);
      console.log('Built evidenceByArea areas:', Object.keys(portfolioData.evidenceByArea));
    } else if (Array.isArray(existingByArea)) {
      // It was sent as a flat array — build from it
      portfolioData.evidenceByArea = buildEvidenceByArea(existingByArea);
    }

    // learningAreaOverviews should be an object, not array
    if (Array.isArray(portfolioData.learningAreaOverviews)) {
      portfolioData.learningAreaOverviews = {};
    }

    // Generate the DOCX
    const doc = generatePortfolio(portfolioData);
    const buffer = await Packer.toBuffer(doc);
    console.log('DOCX generated, size:', buffer.length, 'bytes');

    // Build filename
    const safeName = (portfolioData.childName || 'Child').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 50);
    const safePeriod = (portfolioData.reportingPeriod || 'Portfolio').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 30);
    const filename = `${safeName}-Portfolio-${safePeriod}.docx`;

    // Upload to Vercel Blob
    const blob = await put(`portfolios/${filename}`, buffer, {
      access: 'public',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    console.log('Uploaded to Blob:', blob.url);

    return res.status(200).json({
      success: true,
      filename: filename,
      url: blob.url,
      fileSize: buffer.length,
    });

  } catch (error) {
    console.error('Error generating portfolio:', error.message);
    console.error(error.stack);
    return res.status(500).json({
      success: false,
      error: error.message,
    });
  }
};
