const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, 
        LevelFormat, PageBreak, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');
const fs = require('fs');
const https = require('https');

// This script receives data from Make.com and generates a portfolio DOCX

// Helper function to ensure data is in array format
function ensureArray(data) {
  // If it's already an array, return it
  if (Array.isArray(data)) {
    return data;
  }
  
  // If it's null or undefined, return empty array
  if (!data) {
    return [];
  }
  
  // If it's an object with an 'array' property
  if (typeof data === 'object' && data.array) {
    return Array.isArray(data.array) ? data.array : [];
  }
  
  // If it's an object, try to convert it to an array
  if (typeof data === 'object') {
    // Filter out empty values and metadata
    const values = Object.values(data).filter(item => 
      item && 
      typeof item === 'object' && 
      !Array.isArray(item)
    );
    return values.length > 0 ? values : [];
  }
  
  // Default: return empty array
  return [];
}

async function generateAIProgressSummary(learningArea, evidenceList, apiKey) {
  // If no evidence, return empty (will trigger "no evidence" statement)
  if (!evidenceList || evidenceList.length === 0) {
    return '';
  }

  // Prepare evidence summary for AI
  const evidenceSummary = evidenceList.map(e => 
    `- ${e.title} (${e.date}): ${e.description}. Engagement: ${e.engagement}. Outcomes: ${e.outcomes}`
  ).join('\n');

  const prompt = `You are an educational assessor reviewing a homeschooled child's learning progress.

Learning Area: ${learningArea}

Evidence entries for this term:
${evidenceSummary}

Write a concise 2-3 sentence progress summary that:
- Highlights key learning demonstrated
- Notes engagement and development trends
- Uses professional but warm tone
- Focuses on growth and strengths

Progress Summary:`;

  return new Promise((resolve, reject) => {
    const data = JSON.stringify({
      model: 'gpt-4-turbo-preview',
      messages: [
        { role: 'system', content: 'You are an experienced educational assessor who writes warm, strengths-based progress summaries for homeschooled children.' },
        { role: 'user', content: prompt }
      ],
      max_tokens: 200,
      temperature: 0.7
    });

    const options = {
      hostname: 'api.openai.com',
      port: 443,
      path: '/v1/chat/completions',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': data.length,
        'Authorization': `Bearer ${apiKey}`
      }
    };

    const req = https.request(options, (res) => {
      let responseData = '';
      res.on('data', (chunk) => { responseData += chunk; });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(responseData);
          const summary = parsed.choices[0].message.content.trim();
          resolve(summary);
        } catch (e) {
          reject(e);
        }
      });
    });

    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

function getCurriculumTerm(state) {
  return state === 'NSW' ? 'syllabus' : 'curriculum';
}

function getSectionNumber(area) {
  const areaOrder = {
    'English': 1,
    'Mathematics': 2,
    'Science': 3,
    'HASS': 4,
    'HPE': 5,
    'The Arts': 6,
    'Technologies': 7
  };
  return areaOrder[area] || 8;
}

function generateLearningAreaOverviews(learningAreaOverviews, yearLevel, curriculumTermCap) {
  const sections = [];
  
  Object.entries(learningAreaOverviews).forEach(([area, overview]) => {
    sections.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun(`2.${getSectionNumber(area)} ${area}`)]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: `${curriculumTermCap} Expectations:`, bold: true })]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: overview.stageStatement, italics: true })]
      })
    );

    if (!overview.progressSummary || overview.progressSummary.trim() === '') {
      sections.push(
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Evidence Status: ", bold: true }),
            new TextRun("No formal evidence was documented for this learning area during the current reporting period. Learning in this area has occurred informally through daily activities, conversations, and integrated experiences.")
          ]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({ text: "Future Planning: ", bold: true }),
            new TextRun(`Specific evidence of learning in ${area} will be documented in upcoming reporting periods through planned activities and observations aligned with the relevant outcomes.`)
          ]
        })
      );
    } else {
      sections.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({ text: "Progress Summary: ", bold: true }),
            new TextRun(overview.progressSummary)
          ]
        })
      );
    }
  });

  return sections;
}

function selectDiverseEvidence(evidenceByArea, maxTotal = 30) {
  const allEvidence = Object.values(evidenceByArea).flat();
  
  if (allEvidence.length <= maxTotal) {
    return evidenceByArea; // Return all if under limit
  }

  // Select diverse evidence
  const selected = {};
  const maxPerArea = 5;
  
  Object.entries(evidenceByArea).forEach(([area, evidenceList]) => {
    if (!evidenceList || evidenceList.length === 0) return;
    
    // Sort by engagement level and take most engaged + spread across dates
    const sorted = evidenceList.sort((a, b) => {
      const engagementOrder = { 'Very High': 3, 'High': 2, 'Medium': 1, 'Low': 0 };
      const aScore = engagementOrder[a.engagement] || 0;
      const bScore = engagementOrder[b.engagement] || 0;
      return bScore - aScore;
    });
    
    selected[area] = sorted.slice(0, maxPerArea);
  });
  
  return selected;
}

function generateEvidenceStatistics(evidenceByArea, selectedEvidence) {
  const stats = {};
  
  Object.entries(evidenceByArea).forEach(([area, fullList]) => {
    const selectedList = selectedEvidence[area] || [];
    const total = fullList.length;
    const shown = selectedList.length;
    
    if (total > shown) {
      stats[area] = {
        total: total,
        shown: shown,
        additional: total - shown
      };
    }
  });
  
  return stats;
}

function identifyUnaddressedOutcomes(allOutcomes, evidenceByArea) {
  // Collect all outcome codes that WERE addressed
  const addressedCodes = new Set();
  
  Object.values(evidenceByArea).forEach(evidenceList => {
    evidenceList.forEach(evidence => {
      if (evidence.matchedOutcomes) {
        evidence.matchedOutcomes.forEach(outcome => {
          addressedCodes.add(outcome.code);
        });
      }
    });
  });
  
  // Find outcomes that were NOT addressed
  const unaddressed = allOutcomes.filter(outcome => 
    !addressedCodes.has(outcome.code)
  );
  
  return unaddressed;
}

function generateUnaddressedOutcomesSection(unaddressedOutcomes) {
  if (!unaddressedOutcomes || unaddressedOutcomes.length === 0) {
    return []; // No unaddressed outcomes
  }

  const sections = [];
  const tableBorder = { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC" };
  
  sections.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("3.9 Curriculum Outcomes Not Yet Addressed")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({
        text: "The following curriculum outcomes have not yet been addressed through documented evidence during this reporting period. These will be incorporated into future learning activities as outlined in Section 5: Future Learning Plans."
      })]
    })
  );

  // Create table rows
  const tableRows = [
    // Header row
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          width: { size: 2500, type: WidthType.DXA },
          borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Outcome Code", bold: true })]
          })]
        }),
        new TableCell({
          width: { size: 6860, type: WidthType.DXA },
          borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Outcome Description", bold: true })]
          })]
        })
      ]
    })
  ];

  // Data rows
  unaddressedOutcomes.forEach(outcome => {
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({
            width: { size: 2500, type: WidthType.DXA },
            borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
            children: [new Paragraph({ children: [new TextRun(outcome.code)] })]
          }),
          new TableCell({
            width: { size: 6860, type: WidthType.DXA },
            borders: { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder },
            children: [new Paragraph({ children: [new TextRun(outcome.description)] })]
          })
        ]
      })
    );
  });

  sections.push(
    new Table({
      columnWidths: [2500, 6860],
      rows: tableRows
    })
  );

  return sections;
}

function generateEvidenceSections(evidenceByArea, allOutcomes) {
  const sections = [];
  
  // Select diverse evidence if > 30 total  
  const selectedEvidence = selectDiverseEvidence(evidenceByArea, 30);
  const statistics = generateEvidenceStatistics(evidenceByArea, selectedEvidence);
  
  // Group curriculum outcomes by learning area
  const outcomesByArea = {};
  if (allOutcomes && allOutcomes.length > 0) {
    allOutcomes.forEach(outcome => {
      // Extract learning area from code prefix
      const codePrefix = outcome.code.substring(0, 2).toUpperCase();
      const areaMap = {
        'EN': 'English',
        'MA': 'Mathematics', 
        'ST': 'Science and Technology',
        'HS': 'Human Society and Its Environment',
        'CA': 'Creative Arts',
        'PH': 'Personal Development, Health and Physical Education',
        'PD': 'Personal Development, Health and Physical Education'
      };
      const area = areaMap[codePrefix];
      
      if (area && !outcomesByArea[area]) {
        outcomesByArea[area] = [];
      }
      if (area) {
        outcomesByArea[area].push(outcome);
      }
    });
  }
  
  Object.entries(selectedEvidence).forEach(([area, evidenceList]) => {
    if (!evidenceList || evidenceList.length === 0) return;
    
    sections.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun(`3.${getSectionNumber(area)} ${area}`)]
      })
    );

    evidenceList.forEach((evidence) => {
      sections.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun(`${evidence.title}`)]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Date: ", bold: true }),
            new TextRun(evidence.date)
          ]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Description: ", bold: true }),
            new TextRun(evidence.description)
          ]
        })
      );

      if (evidence.matchedOutcomes && evidence.matchedOutcomes.length > 0) {
        sections.push(
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: "Curriculum Outcomes Addressed:", bold: true })]
          })
        );
        evidence.matchedOutcomes.forEach(outcome => {
          sections.push(
            new Paragraph({
              numbering: { reference: "bullet-list", level: 0 },
              children: [new TextRun(`${outcome.code}: ${outcome.description}`)]
            })
          );
        });
      }

      if (evidence.engagement) {
        sections.push(
          new Paragraph({
            spacing: { before: 60, after: 60 },
            children: [
              new TextRun({ text: "Child Engagement: ", bold: true }),
              new TextRun(evidence.engagement)
            ]
          })
        );
      }

      // Add images if provided
      if (evidence.images && evidence.images.length > 0) {
        evidence.images.forEach(imagePath => {
          try {
            if (fs.existsSync(imagePath)) {
              const imageData = fs.readFileSync(imagePath);
              const imageExtension = imagePath.split('.').pop().toLowerCase();
              
              sections.push(
                new Paragraph({
                  spacing: { before: 60, after: 120 },
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: imageData,
                      transformation: {
                        width: 400,
                        height: 300
                      },
                      type: imageExtension === 'jpg' || imageExtension === 'jpeg' ? 'jpg' : imageExtension
                    })
                  ]
                })
              );
            }
          } catch (err) {
            console.error(`Error loading image ${imagePath}:`, err.message);
          }
        });
      } else {
        // Add spacing after last text element if no images
        sections.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      }
    });
    
    // Add statistics summary if evidence was limited
    if (statistics[area]) {
      sections.push(
        new Paragraph({
          spacing: { before: 60, after: 120 },
          children: [
            new TextRun({ text: `Additional Evidence: `, bold: true }),
            new TextRun({ 
              text: `Plus ${statistics[area].additional} additional ${area} ${statistics[area].additional === 1 ? 'activity' : 'activities'} documented during this reporting period.`
            })
          ]
        })
      );
    }
    
    // Check for unaddressed outcomes in THIS learning area
    if (outcomesByArea[area]) {
      const addressedOutcomeCodes = new Set();
      
      // Get all evidence for this area (not just selected)
      const fullEvidenceList = evidenceByArea[area] || [];
      fullEvidenceList.forEach(evidence => {
        if (evidence.matchedOutcomes) {
          evidence.matchedOutcomes.forEach(outcome => {
            addressedOutcomeCodes.add(outcome.code);
          });
        }
      });
      
      const unaddressedOutcomes = outcomesByArea[area].filter(
        outcome => !addressedOutcomeCodes.has(outcome.code)
      );
      
      if (unaddressedOutcomes.length > 0) {
        sections.push(
          new Paragraph({
            spacing: { before: 180, after: 60 },
            children: [
              new TextRun({ text: `${area} Outcomes Not Yet Addressed:`, bold: true, italics: true })
            ]
          }),
          new Paragraph({
            spacing: { after: 60 },
            children: [
              new TextRun({ 
                text: `The following outcomes have not yet been formally documented in this reporting period and will be addressed in future learning activities:`,
                italics: true 
              })
            ]
          })
        );
        
        unaddressedOutcomes.forEach(outcome => {
          sections.push(
            new Paragraph({
              numbering: { reference: "bullet-list", level: 0 },
              children: [new TextRun({ text: `${outcome.code}: ${outcome.description}`, italics: true })]
            })
          );
        });
        
        sections.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      }
    }
  });

  return sections;
}

function generatePortfolio(portfolioData) {
  // Ensure arrays are properly formatted (Make.com sometimes sends objects instead of arrays)
  if (portfolioData.evidenceEntries) {
    portfolioData.evidenceEntries = ensureArray(portfolioData.evidenceEntries);
  }
  if (portfolioData.curriculumOutcomes) {
    portfolioData.curriculumOutcomes = ensureArray(portfolioData.curriculumOutcomes);
  }
  if (portfolioData.allOutcomes) {
    portfolioData.allOutcomes = ensureArray(portfolioData.allOutcomes);
  }
  if (portfolioData.evidenceByArea) {
    Object.keys(portfolioData.evidenceByArea).forEach(key => {
      portfolioData.evidenceByArea[key] = ensureArray(portfolioData.evidenceByArea[key]);
    });
  }
  
  const {
    childName,
    yearLevel,
    reportingPeriod,
    parentName,
    state,
    curriculum,
    learningAreaOverviews,
    evidenceByArea,
    progressAssessment,
    futurePlans,
    allOutcomes // Array of all curriculum outcomes for this stage
  } = portfolioData;

  const curriculumTerm = getCurriculumTerm(state);
  const curriculumTermCap = curriculumTerm.charAt(0).toUpperCase() + curriculumTerm.slice(1);

  const currentDate = new Date().toLocaleDateString('en-AU', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });

  const numberingConfig = {
    config: [
      {
        reference: "bullet-list",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  };

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 24 }
        }
      },
      paragraphStyles: [
        {
          id: "Title",
          name: "Title",
          basedOn: "Normal",
          run: { size: 56, bold: true, color: "000000", font: "Arial" },
          paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.CENTER }
        },
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 32, bold: true, color: "000000", font: "Arial" },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 28, bold: true, color: "000000", font: "Arial" },
          paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 1 }
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 26, bold: true, color: "000000", font: "Arial" },
          paragraph: { spacing: { before: 120, after: 60 }, outlineLevel: 2 }
        }
      ]
    },
    numbering: numberingConfig,
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        new Paragraph({
          heading: HeadingLevel.TITLE,
          children: [new TextRun("Home Education Learning Portfolio")]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
          children: [new TextRun({ text: childName, size: 32, bold: true })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: yearLevel, size: 28 })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: reportingPeriod, size: 28 })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 240 },
          children: [new TextRun({ text: `Prepared by: ${parentName}`, size: 24 })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 240 },
          children: [new TextRun({ text: `Date: ${currentDate}`, size: 24 })]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("1. Learning Program Overview")]
        }),
        
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun(`1.1 ${curriculumTermCap} Framework`)]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({
            text: `This learning portfolio demonstrates ${childName}'s educational progress during ${reportingPeriod}. Our home education program aligns with the ${curriculum} and covers all key learning areas required under ${state} homeschooling regulations.`
          })]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("1.2 Compliance with Disability Standards for Education 2005")]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({
            text: `This educational program has been developed in accordance with the Disability Standards for Education 2005 (Cth), which ensure that students with disability are able to access and participate in education on the same basis as students without disability.`
          })]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({
            text: `${childName} has a neurodivergent learning profile, and our program incorporates reasonable adjustments as outlined in the Standards. These adjustments are not considered "special" or "extra" support, but rather the necessary adaptations that enable ${childName} to access the ${curriculumTerm} effectively.`
          })]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({ text: "Key adjustments implemented:", bold: true })]
        }),
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun("Flexible pacing that honours the child's autonomy and reduces demand-related anxiety")]
        }),
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun("Collaborative approach to learning activities, allowing the child to maintain a sense of control and choice")]
        }),
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun("Integration of special interests and preferred learning modalities to enhance engagement")]
        }),
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun("Low-demand presentation of learning opportunities that reduces pressure while maintaining educational rigour")]
        }),
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          spacing: { after: 120 },
          children: [new TextRun("Recognition that anxiety and overwhelm are communication, not misbehaviour, requiring adaptive responses")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({
            text: `These adjustments are fundamental to our educational approach and enable ${childName} to demonstrate learning and progress toward ${curriculumTerm} outcomes in ways that respect neurodivergent learning patterns.`
          })]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("1.3 Educational Philosophy and Approach")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({
            text: `Our home education program recognises that meaningful learning occurs when children feel safe, autonomous, and connected. We provide a rich learning environment that allows natural curiosity to drive engagement with ${curriculumTerm} content, while maintaining clear alignment with ${curriculum} outcomes.`
          })]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("2. Learning Areas Overview")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({
            text: `The following provides an overview of ${curriculum} expectations for ${yearLevel} students in each learning area, along with a brief summary of ${childName}'s progress toward these standards.`
          })]
        }),

        ...generateLearningAreaOverviews(learningAreaOverviews, yearLevel, curriculumTermCap),

        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("3. Detailed Learning Evidence by Subject Area")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun("The following sections present specific evidence of learning across all curriculum areas, with each entry linked to curriculum outcomes.")]
        }),

        ...generateEvidenceSections(evidenceByArea, allOutcomes || []),
        
        // Add unaddressed outcomes section if we have curriculum data
        ...(allOutcomes && allOutcomes.length > 0 
          ? generateUnaddressedOutcomesSection(
              identifyUnaddressedOutcomes(allOutcomes, evidenceByArea)
            )
          : []),

        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("4. Parent Assessment of Progress")]
        }),
        
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("4.1 Cognitive Development")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(progressAssessment.cognitive)]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("4.2 Social Development")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(progressAssessment.social)]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("4.3 Emotional Development")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(progressAssessment.emotional)]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("4.4 Physical Development")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(progressAssessment.physical)]
        }),

        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("5. Future Learning Plans")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(futurePlans.overview)]
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("5.1 Learning Goals")]
        }),
        ...futurePlans.goals.map(goal => 
          new Paragraph({
            numbering: { reference: "bullet-list", level: 0 },
            children: [new TextRun(goal)]
          })
        ),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("5.2 Planned Strategies")]
        }),
        ...futurePlans.strategies.map(strategy =>
          new Paragraph({
            numbering: { reference: "bullet-list", level: 0 },
            children: [new TextRun(strategy)]
          })
        ),
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun("6. Resources for Learning")]
        }),
        
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("6.1 Resources Used During This Period")]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun("The following resources supported learning across curriculum areas during this reporting period:")]
        }),
        
        ...generateResourcesList(evidenceByArea, 'used'),
        
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("6.2 Planned Resources for Next Learning Period")]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun(futurePlans.plannedResources || "We will continue using many of the resources that have proven effective, supplemented with additional materials as learning needs develop.")]
        }),
        
        ...generateResourcesList(evidenceByArea, 'planned', futurePlans.additionalResources)
      ]
    }]
  });

  return doc;
}

function generateResourcesList(evidenceByArea, type, additionalResources = []) {
  const sections = [];
  const resourceCategories = {
    'People': new Set(),
    'Places': new Set(),
    'Physical Resources': new Set(),
    'Digital Resources': new Set()
  };
  
  // Collect resources from evidence
  if (type === 'used') {
    Object.values(evidenceByArea).forEach(evidenceList => {
      if (!evidenceList) return;
      evidenceList.forEach(evidence => {
        if (evidence.resources) {
          evidence.resources.forEach(resource => {
            // Categorize resource
            const category = categorizeResource(resource);
            resourceCategories[category].add(resource);
          });
        }
      });
    });
  }
  
  // Add additional planned resources
  if (type === 'planned' && additionalResources && additionalResources.length > 0) {
    additionalResources.forEach(resource => {
      const category = categorizeResource(resource);
      resourceCategories[category].add(resource);
    });
  }
  
  // Generate sections for each category
  Object.entries(resourceCategories).forEach(([category, resourceSet]) => {
    if (resourceSet.size === 0) return;
    
    sections.push(
      new Paragraph({
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: category + ":", bold: true })]
      })
    );
    
    Array.from(resourceSet).sort().forEach(resource => {
      sections.push(
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun(resource)]
        })
      );
    });
  });
  
  if (sections.length === 0) {
    sections.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "Resources will be documented as the learning program develops.", italics: true })]
      })
    );
  }
  
  return sections;
}

function categorizeResource(resource) {
  const lower = resource.toLowerCase();
  
  // People indicators
  if (lower.includes('tutor') || lower.includes('teacher') || lower.includes('instructor') || 
      lower.includes('coach') || lower.includes('mentor') || lower.includes('parent') ||
      lower.includes('specialist') || lower.includes('therapist')) {
    return 'People';
  }
  
  // Places indicators
  if (lower.includes('library') || lower.includes('museum') || lower.includes('park') || 
      lower.includes('centre') || lower.includes('center') || lower.includes('club') ||
      lower.includes('venue') || lower.includes('location') || lower.includes('facility')) {
    return 'Places';
  }
  
  // Digital indicators
  if (lower.includes('app') || lower.includes('website') || lower.includes('online') || 
      lower.includes('software') || lower.includes('digital') || lower.includes('.com') ||
      lower.includes('video') || lower.includes('youtube') || lower.includes('platform') ||
      lower.includes('program') || lower.includes('subscription')) {
    return 'Digital Resources';
  }
  
  // Default to Physical Resources
  return 'Physical Resources';
}

// Main execution
async function main() {
  const inputData = JSON.parse(process.argv[2]);
  const outputPath = process.argv[3] || '/mnt/user-data/outputs/portfolio.docx';
  
  const doc = generatePortfolio(inputData);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  
  console.log(JSON.stringify({ 
    success: true, 
    filepath: outputPath,
    message: 'Portfolio generated successfully'
  }));
}

if (require.main === module) {
  main().catch(err => {
    console.error(JSON.stringify({ 
      success: false, 
      error: err.message 
    }));
    process.exit(1);
  });
}

module.exports = { generatePortfolio, generateAIProgressSummary };
