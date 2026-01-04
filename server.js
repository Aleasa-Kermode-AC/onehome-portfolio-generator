const express = require('express');
const { generatePortfolio } = require('./generate-portfolio');
const { Packer } = require('docx');

const app = express();
const PORT = process.env.PORT || 3000;

// Increase payload limit for large requests with images
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Health check endpoint
app.get('/', (req, res) => {
  res.json({ 
    status: 'healthy', 
    service: 'OneHome Education Portfolio Generator',
    version: '1.0.0'
  });
});

// Portfolio generation endpoint
app.post('/generate-portfolio', async (req, res) => {
  try {
    console.log('Received portfolio generation request');
    
    const portfolioData = req.body;
    
    // Validate required fields
    if (!portfolioData.childName || !portfolioData.yearLevel) {
      return res.status(400).json({ 
        success: false, 
        error: 'Missing required fields: childName and yearLevel are required' 
      });
    }
    
    // Generate the portfolio document
    console.log(`Generating portfolio for ${portfolioData.childName}`);
    const doc = generatePortfolio(portfolioData);
    
    // Convert to buffer
    const buffer = await Packer.toBuffer(doc);
    
    console.log('Portfolio generated successfully');
    
    // Return as downloadable file
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${portfolioData.childName}-Portfolio-${portfolioData.reportingPeriod}.docx"`);
    res.setHeader('Content-Length', buffer.length);
    
    res.send(buffer);
    
  } catch (error) {
    console.error('Error generating portfolio:', error);
    res.status(500).json({ 
      success: false, 
      error: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
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
});
