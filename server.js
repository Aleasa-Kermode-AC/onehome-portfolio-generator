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
    
    // Helper function to robustly parse Make.com data
    function parseMakeComData(data, fieldName) {
      console.log(`Parsing ${fieldName}, type:`, typeof data);
      
      // Already an array - return it
      if (Array.isArray(data)) {
        console.log(`${fieldName} is already array, length:`, data.length);
        return data;
      }
      
      // Null/undefined - return empty array
      if (!data) {
        console.log(`${fieldName} is null/undefined`);
        return [];
      }
      
      // String - try to parse as JSON (handles double-encoding)
      if (typeof data === 'string') {
        console.log(`${fieldName} is string, attempting to parse...`);
        let parsed = data;
        let attempts = 0;
        
        // Sometimes Make.com double or triple encodes - keep parsing until we get an object/array
        while (typeof parsed === 'string' && attempts < 5) {
          try {
            parsed = JSON.parse(parsed);
            attempts++;
            console.log(`Parse attempt ${attempts} succeeded, new type:`, typeof parsed);
          } catch (e) {
            console.error(`Parse attempt ${attempts} failed:`, e.message);
            break;
          }
        }
        
        // If we got an array, return it
        if (Array.isArray(parsed)) {
          console.log(`${fieldName} parsed to array, length:`, parsed.length);
          return parsed;
        }
        
        // If we got an object, try to extract array
        if (parsed && typeof parsed === 'object') {
          if (parsed.array && Array.isArray(parsed.array)) {
            console.log(`${fieldName} has array property, length:`, parsed.array.length);
            return parsed.array;
          }
          // Try to convert object values to array
          const values = Object.values(parsed).filter(item => item && typeof item === 'object');
          console.log(`${fieldName} converted object to array, length:`, values.length);
          return values;
        }
        
        console.log(`${fieldName} could not be parsed, returning empty array`);
        return [];
      }
      
      // Object - try to extract array
      if (typeof data === 'object') {
        if (data.array && Array.isArray(data.array)) {
          console.log(`${fieldName} has array property, length:`, data.array.length);
          return data.array;
        }
        const values = Object.values(data).filter(item => item && typeof item === 'object');
        console.log(`${fieldName} converted object to array, length:`, values.length);
        return values;
      }
      
      console.log(`${fieldName} unknown format, returning empty array`);
      return [];
    }
    
    // Parse both arrays
    const evidenceList = parseMakeComData(portfolioData.evidenceEntries, 'evidenceEntries');
    const outcomesList = parseMakeComData(portfolioData.curriculumOutcomes, 'curriculumOutcomes');
    
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
