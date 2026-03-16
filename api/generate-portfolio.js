const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

// Helper to normalize arrays from Make.com (sometimes sends objects instead of arrays)
function toArray(val) {
  if (Array.isArray(val)) return val;
  if (!val) return [];
  if (typeof val === 'object') return Object.values(val).filter(v => v !== null && v !== undefined);
  return [];
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
      version: '3.0.0',
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

    if (portfolioData.evidenceByArea && typeof portfolioData.evidenceByArea === 'object') {
      Object.keys(portfolioData.evidenceByArea).forEach(key => {
        portfolioData.evidenceByArea[key] = toArray(portfolioData.evidenceByArea[key]);
      });
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

    // Return URL to Make.com - Make.com will update Airtable
    return res.status(200).json({
      success: true,
      filename: filename,
      url: blob.url,
      fileSize: buffer.length,
    });

  } catch (error) {
    console.error('Error generating portfolio:', error.message);
    return res.status(500).json({
      success: false,
      error: error.message,
    });
  }
};
