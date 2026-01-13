const express = require('express');
const { generatePortfolio } = require('./generate-portfolio');
const { Packer } = require('docx');

const app = express();
const PORT = process.env.PORT || 10000;

// Increase payload limit
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Health check
app.get('/', (req, res) => {
  res.json({ 
    status: 'healthy', 
    service: 'OneHome Education Portfolio Generator',
    version: '1.0.0'
  });
});

// Portfolio generation with detailed logging
app.post('/generate-portfolio', async (req, res) => {
  try {
    console.log('=== INCOMING REQUEST ===');
    console.log('Request body type:', typeof req.body);
    console.log('Request body:', JSON.stringify(req.body, null, 2));
    
    const portfolioData = req.body;
    
    // Log specific fields
    console.log('evidenceEntries type:', typeof portfolioData.evidenceEntries);
    console.log('evidenceEntries value:', portfolioData.evidenceEntries);
    console.log('evidenceEntries is Array?:', Array.isArray(portfolioData.evidenceEntries));
    
    console.log('curriculumOutcomes type:', typeof portfolioData.curriculumOutcomes);
    console.log('curriculumOutcomes is Array?:', Array.isArray(portfolioData.curriculumOutcomes));
    
    // Validate required fields
    if (!portfolioData.childName || !portfolioData.yearLevel) {
      console.error('Missing required fields');
      return res.status(400).json({ 
        success: false, 
        error: 'Missing required fields: childName and yearLevel are required' 
      });
    }
    
    // Convert evidence entries if needed
    let evidenceList = portfolioData.evidenceEntries || [];
    console.log('Initial evidenceList type:', typeof evidenceList);
    
    // If it's a string, parse it as JSON
    if (typeof evidenceList === 'string') {
      console.log('Parsing evidenceEntries from JSON string');
      try {
        evidenceList = JSON.parse(evidenceList);
        console.log('Parsed evidenceList, is array?:', Array.isArray(evidenceList));
      } catch (e) {
        console.error('Failed to parse evidenceEntries:', e.message);
        evidenceList = [];
      }
    }
    
    // If it's an object, try to convert it to an array
    if (!Array.isArray(evidenceList) && typeof evidenceList === 'object') {
      console.log('Converting evidenceEntries object to array');
      // Try to extract array values
      if (evidenceList.array) {
        evidenceList = evidenceList.array;
      } else {
        evidenceList = Object.values(evidenceList).filter(item => item && typeof item === 'object');
      }
      console.log('Converted evidenceList length:', evidenceList.length);
    }
    
    // Same for curriculum outcomes
    let outcomesList = portfolioData.curriculumOutcomes || [];
    console.log('Initial outcomesList type:', typeof outcomesList);
    
    // If it's a string, parse it as JSON
    if (typeof outcomesList === 'string') {
      console.log('Parsing curriculumOutcomes from JSON string');
      try {
        outcomesList = JSON.parse(outcomesList);
        console.log('Parsed outcomesList, is array?:', Array.isArray(outcomesList));
      } catch (e) {
        console.error('Failed to parse curriculumOutcomes:', e.message);
        outcomesList = [];
      }
    }
    
    if (!Array.isArray(outcomesList) && typeof outcomesList === 'object') {
      console.log('Converting curriculumOutcomes object to array');
      if (outcomesList.array) {
        outcomesList = outcomesList.array;
      } else {
        outcomesList = Object.values(outcomesList).filter(item => item && typeof item === 'object');
      }
      console.log('Converted outcomesList length:', outcomesList.length);
    }
    
    // Update the portfolioData with converted arrays
    portfolioData.evidenceEntries = evidenceList;
    portfolioData.curriculumOutcomes = outcomesList;
    
    console.log('Generating portfolio for:', portfolioData.childName);
    console.log('Evidence count:', evidenceList.length);
    console.log('Outcomes count:', outcomesList.length);
    
    // Generate the portfolio document
    const doc = generatePortfolio(portfolioData);
    
    // Convert to buffer
    const buffer = await Packer.toBuffer(doc);
    
    console.log('Portfolio generated successfully, size:', buffer.length, 'bytes');
    
    // Return as downloadable file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${portfolioData.childName}-Portfolio-${portfolioData.reportingPeriod.replace(/[^a-zA-Z0-9]/g, '-')}.docx"`);
    res.setHeader('Content-Length', buffer.length);
    
    res.send(buffer);
    
  } catch (error) {
    console.error('=== ERROR ===');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    res.status(500).json({ 
      success: false, 
      error: error.message
    });
  }
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ 
    success: false, 
    error: 'Internal server error',
    message: err.message 
  });
});

app.listen(PORT, () => {
  console.log(`Portfolio Generator webhook running on port ${PORT}`);
  console.log('Debug logging enabled');
});
