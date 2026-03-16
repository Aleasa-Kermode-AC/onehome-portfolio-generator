const { generatePortfolio } = require('../generate-portfolio');
const { Packer } = require('docx');
const { put } = require('@vercel/blob');

// This function runs AFTER we've already responded 202 to Make.com
async function processPortfolio(portfolioData) {
  const recordId = portfolioData.recordId;

  try {
    console.log('Background processing started for record:', recordId);

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

    // Write URL + status back to Airtable
    const airtableBaseId = process.env.AIRTABLE_BASE_ID;
    const airtableApiKey = process.env.AIRTABLE_API_KEY;

    if (!airtableBaseId || !airtableApiKey) {
      throw new Error('Missing AIRTABLE_BASE_ID or AIRTABLE_API_KEY environment variables');
    }

    const airtableResponse = await fetch(
      `https://api.airtable.com/v0/${airtableBaseId}/Portfolio%20Requests/${recordId}`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${airtableApiKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          fields: {
            'Generated Document URL': blob.url,
            Status: 'Complete',
          },
        }),
      }
    );

    if (!airtableResponse.ok) {
      const errText = await airtableResponse.text();
      throw new Error(`Airtable update failed: ${errText}`);
    }

    console.log('Airtable updated successfully for record:', recordId);

  } catch (error) {
    console.error('Background processing error for record:', recordId, error.message);

    // Try to write the error status back to Airtable so the user knows it failed
    try {
      const airtableBaseId = process.env.AIRTABLE_BASE_ID;
      const airtableApiKey = process.env.AIRTABLE_API_KEY;
      if (airtableBaseId && airtableApiKey && recordId) {
        await fetch(
          `https://api.airtable.com/v0/${airtableBaseId}/Portfolio%20Requests/${recordId}`,
          {
            method: 'PATCH',
            headers: {
              Authorization: `Bearer ${airtableApiKey}`,
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              fields: {
                Status: 'Error',
              },
            }),
          }
        );
      }
    } catch (airtableErr) {
      console.error('Could not write error status to Airtable:', airtableErr.message);
    }
  }
}

module.exports = async (req, res) => {
  // Handle CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Health check
  if (req.method === 'GET') {
    return res.status(200).json({
      status: 'healthy',
      service: 'OneHome Education Portfolio Generator',
      version: '2.0.0',
      mode: 'async',
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const portfolioData = req.body;

  // Basic validation
  if (!portfolioData || !portfolioData.childName) {
    return res.status(400).json({ success: false, error: 'Missing required field: childName' });
  }

  if (!portfolioData.recordId) {
    return res.status(400).json({ success: false, error: 'Missing required field: recordId — needed to write result back to Airtable' });
  }

  // Respond immediately so Make.com doesn't time out
  res.status(202).json({
    success: true,
    message: 'Portfolio generation started. Result will be written directly to Airtable.',
    recordId: portfolioData.recordId,
  });

  // Process in background (Vercel Fluid Compute keeps the function alive after response)
  processPortfolio(portfolioData).catch((err) => {
    console.error('Unhandled error in processPortfolio:', err);
  });
};
