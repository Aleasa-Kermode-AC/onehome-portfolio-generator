const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

module.exports = async (req, res) => {
  // Handle CORS
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
      version: '1.0.0'
    });
  }
  
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }
  
  try {
    console.log('=== INCOMING REQUEST ===');
    console.log('Request body type:', typeof req.body);
    
    const portfolioData = req.body;
    
    // Helper function to ensure data is in array format
    function ensureArray(data, fieldName) {
      console.log(`Parsing ${fieldName}, type:`, typeof data);
      
      if (Array.isArray(data)) {
        console.log(`${fieldName} is already array, length:`, data.length);
        return data;
      }
      
      if (!data) {
        console.log(`${fieldName} is null/undefined`);
        return [];
      }
      
      if (typeof data === 'string') {
        console.log(`${fieldName} is string, attempting to parse...`);
        let parsed = data;
        let attempts = 0;
        
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
        
        if (Array.isArray(parsed)) {
          console.log(`${fieldName} parsed to array, length:`, parsed.length);
          return parsed;
        }
        
        if (parsed && typeof parsed === 'object') {
          if (parsed.array && Array.isArray(parsed.array)) {
            console.log(`${fieldName} has array property, length:`, parsed.array.length);
            return parsed.array;
          }
          const values = Object.values(parsed).filter(item => item && typeof item === 'object');
          console.log(`${fieldName} converted object to array, length:`, values.length);
          return values;
        }
        
        console.log(`${fieldName} could not be parsed, returning empty array`);
        return [];
      }
      
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
    
    // Parse arrays
    const evidenceList = ensureArray(portfolioData.evidenceEntries, 'evidenceEntries');
    const outcomesList = ensureArray(portfolioData.curriculumOutcomes, 'curriculumOutcomes');
    
    console.log('Evidence count:', evidenceList.length);
    console.log('Outcomes count:', outcomesList.length);
    
    console.log('Generating portfolio for:', {
      childName: portfolioData.childName,
      yearLevel: portfolioData.yearLevel,
      reportingPeriod: portfolioData.reportingPeriod,
      parentName: portfolioData.parentName,
      state: portfolioData.state,
      curriculum: portfolioData.curriculum
    });
    
    // Generate the portfolio document
    const doc = generatePortfolio(portfolioData);
    
    // Convert to buffer
    const buffer = await Packer.toBuffer(doc);
    
    console.log('Portfolio generated successfully, size:', buffer.length, 'bytes');
    
    // Create safe filename
    const safeName = (portfolioData.childName || 'Child').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 50);
    const safePeriod = (portfolioData.reportingPeriod || 'Portfolio').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 30);
    const filename = `${safeName}-Portfolio-${safePeriod}.docx`;
    
    console.log('Uploading file to Vercel Blob:', filename);
    
    // Upload to Vercel Blob storage
    const blob = await put(`portfolios/${filename}`, buffer, {
      access: 'public',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    
    console.log('File uploaded successfully, URL:', blob.url);
    
    // Return the public URL
    return res.status(200).json({
      success: true,
      filename: filename,
      url: blob.url,
      fileSize: buffer.length
    });
    
  } catch (error) {
    console.error('=== ERROR ===');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    return res.status(500).json({ 
      success: false, 
      error: error.message
    });
  }
};
