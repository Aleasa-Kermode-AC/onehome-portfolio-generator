const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, 
        LevelFormat, PageBreak, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');

/**
 * OneHome Education Portfolio Generator
 * Generates NSW-compliant homeschool learning portfolios as DOCX documents
 * Version 1.2.0 - Fixed data parsing and section generation
 */

function generatePortfolio(portfolioData) {
  // Extract data with defaults
  const {
    childName = 'Child',
    yearLevel = 'Stage 2',
    reportingPeriod = 'Current Period',
    parentName,
    parentname, // Handle lowercase version from Make.com
    state = 'NSW',
    curriculum = 'NSW Syllabus',
    learningAreaOverviews = {},
    evidenceByArea = {},
    progressAssessment = {},
    futurePlans = {},
    curriculumOutcomes = [],
    evidenceEntries = []
  } = portfolioData;
  
  // Use parentName or parentname (handle case sensitivity)
  const finalParentName = parentName || parentname || 'Parent/Carer';
  
  // Parse futurePlans if it's a string
  let parsedFuturePlans = futurePlans;
  if (typeof futurePlans === 'string') {
    try {
      // Remove any escaped characters and parse
      const cleanedString = futurePlans.replace(/\\n/g, '\n').replace(/\\"/g, '"');
      parsedFuturePlans = JSON.parse(cleanedString);
    } catch (e) {
      // If it looks like JSON but failed to parse, try to extract values manually
      if (futurePlans.includes('"overview"') || futurePlans.includes('"goals"')) {
        try {
          // Try wrapping in proper JSON if it's malformed
          const fixed = futurePlans.replace(/^\{?\s*/, '{').replace(/\s*\}?$/, '}');
          parsedFuturePlans = JSON.parse(fixed);
        } catch (e2) {
          parsedFuturePlans = { overview: futurePlans };
        }
      } else {
        parsedFuturePlans = { overview: futurePlans };
      }
    }
  }
  
  // Log for debugging
  console.log('Parsed futurePlans:', parsedFuturePlans);
  
  // Parse progressAssessment if it's a string
  let parsedProgressAssessment = progressAssessment;
  if (typeof progressAssessment === 'string') {
    try {
      parsedProgressAssessment = JSON.parse(progressAssessment);
    } catch (e) {
      parsedProgressAssessment = {};
    }
  }
  
  const curriculumTerm = state === 'NSW' ? 'syllabus' : 'curriculum';
  const curriculumTermCap = state === 'NSW' ? 'Syllabus' : 'Curriculum';
  const currentDate = new Date().toLocaleDateString('en-AU', { 
    day: 'numeric', 
    month: 'long', 
    year: 'numeric' 
  });

  // Build document sections
  const children = [];
  
  // === TITLE PAGE ===
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 200 },
      children: [new TextRun({ text: "Home Education Learning Portfolio", bold: true, size: 56 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: childName, size: 48 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: yearLevel, size: 36 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: reportingPeriod, size: 36 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: `Prepared by: ${finalParentName}`, size: 28 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: [new TextRun({ text: `Date: ${currentDate}`, size: 28 })]
    }),
    new Paragraph({ children: [new PageBreak()] })
  );

  // === SECTION 1: LEARNING PROGRAM OVERVIEW ===
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("1. Learning Program Overview")]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("1.1 Syllabus Framework")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(`This learning portfolio demonstrates ${childName}'s educational progress during ${reportingPeriod}. Our home education program aligns with the ${curriculum} and covers all key learning areas required under ${state} homeschooling regulations.`)]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("1.2 Compliance with Disability Standards for Education 2005")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun("This educational program has been developed in accordance with the Disability Standards for Education 2005 (Cth), which ensure that students with disability are able to access and participate in education on the same basis as students without disability.")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(`${childName} has a neurodivergent learning profile, and our program incorporates reasonable adjustments as outlined in the Standards. These adjustments are not considered "special" or "extra" support, but rather the necessary adaptations that enable ${childName} to access the ${curriculumTerm} effectively.`)]
    }),
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "Key adjustments implemented:", bold: true })]
    })
  );
  
  // Add adjustment bullet points
  const adjustments = [
    "Flexible pacing that honours the child's autonomy and reduces demand-related anxiety",
    "Collaborative approach to learning activities, allowing the child to maintain a sense of control and choice",
    "Integration of special interests and preferred learning modalities to enhance engagement",
    "Low-demand presentation of learning opportunities that reduces pressure while maintaining educational rigour",
    "Recognition that anxiety and overwhelm are communication, not misbehaviour, requiring adaptive responses"
  ];
  
  adjustments.forEach(adj => {
    children.push(
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        children: [new TextRun(adj)]
      })
    );
  });
  
  children.push(
    new Paragraph({
      spacing: { before: 120, after: 120 },
      children: [new TextRun(`These adjustments are fundamental to our educational approach and enable ${childName} to demonstrate learning and progress toward ${curriculumTerm} outcomes in ways that respect neurodivergent learning patterns.`)]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("1.3 Educational Philosophy and Approach")]
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun(`Our home education program recognises that meaningful learning occurs when children feel safe, autonomous, and connected. We provide a rich learning environment that allows natural curiosity to drive engagement with ${curriculumTerm} content, while maintaining clear alignment with ${curriculum} outcomes.`)]
    })
  );

  // === SECTION 2: LEARNING AREAS OVERVIEW ===
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("2. Learning Areas Overview")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(`The following provides an overview of ${curriculum} expectations for ${yearLevel} students in each learning area, along with a brief summary of ${childName}'s progress toward these standards.`)]
    })
  );
  
  // Generate learning area sections
  const learningAreaSections = generateLearningAreaOverviews(
    learningAreaOverviews, 
    evidenceByArea,
    curriculumOutcomes,
    yearLevel, 
    curriculumTermCap,
    childName
  );
  children.push(...learningAreaSections);

  // === SECTION 3: DETAILED LEARNING EVIDENCE ===
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("3. Detailed Learning Evidence by Subject Area")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun("The following sections present specific evidence of learning across all curriculum areas, with each entry linked to curriculum outcomes.")]
    })
  );
  
  // Generate evidence sections
  const evidenceSections = generateEvidenceSections(evidenceByArea, curriculumOutcomes);
  children.push(...evidenceSections);

  // === SECTION 4: PARENT ASSESSMENT OF PROGRESS ===
  children.push(
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
      children: [new TextRun(parsedProgressAssessment.cognitive || "No assessment provided.")]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("4.2 Social Development")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(parsedProgressAssessment.social || "No assessment provided.")]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("4.3 Emotional Development")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(parsedProgressAssessment.emotional || "No assessment provided.")]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("4.4 Physical Development")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(parsedProgressAssessment.physical || "No assessment provided.")]
    })
  );

  // === SECTION 5: FUTURE LEARNING PLANS ===
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("5. Future Learning Plans")]
    })
  );
  
  // Overview
  const futureOverview = parsedFuturePlans.overview || '';
  if (futureOverview && futureOverview.trim() !== '') {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun(futureOverview)]
      })
    );
  } else {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "No future plans overview provided.", italics: true })]
      })
    );
  }
  
  // Learning Goals
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("5.1 Learning Goals")]
    })
  );
  
  const futureGoals = parsedFuturePlans.goals || '';
  if (futureGoals && futureGoals.trim() !== '') {
    // Split goals by newlines or periods and create bullet points
    const goalsList = futureGoals.split(/[\n\r]+/).filter(g => g.trim());
    if (goalsList.length > 1) {
      goalsList.forEach(goal => {
        children.push(
          new Paragraph({
            numbering: { reference: "bullet-list", level: 0 },
            children: [new TextRun(goal.trim())]
          })
        );
      });
    } else {
      children.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(futureGoals)]
        })
      );
    }
  } else {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "No learning goals specified.", italics: true })]
      })
    );
  }
  
  // Planned Strategies
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("5.2 Planned Strategies")]
    })
  );
  
  const futureStrategies = parsedFuturePlans.strategies || '';
  if (futureStrategies && futureStrategies.trim() !== '') {
    // Split only by newlines or by ", " followed by a capital letter (indicating a new strategy)
    // This preserves phrases like "Allow for breaks, movement, and regulation needs"
    let strategiesList = [];
    
    // First try splitting by newlines
    if (futureStrategies.includes('\n')) {
      strategiesList = futureStrategies.split(/[\n\r]+/).filter(s => s.trim());
    } else {
      // Split by comma followed by space and capital letter (new sentence/strategy)
      strategiesList = futureStrategies.split(/,\s+(?=[A-Z])/).filter(s => s.trim());
    }
    
    if (strategiesList.length > 1) {
      strategiesList.forEach(strategy => {
        children.push(
          new Paragraph({
            numbering: { reference: "bullet-list", level: 0 },
            children: [new TextRun(strategy.trim())]
          })
        );
      });
    } else {
      children.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(futureStrategies)]
        })
      );
    }
  } else {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "No strategies specified.", italics: true })]
      })
    );
  }
  
  // Planned Resources
  const plannedResources = parsedFuturePlans.plannedResources || '';
  if (plannedResources && plannedResources.trim() !== '') {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("5.3 Planned Resources")]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun(plannedResources)]
      })
    );
  }

  // === SECTION 6: RESOURCES ===
  // Extract resources from evidence descriptions
  const extractedResources = extractResourcesFromEvidence(evidenceByArea);
  
  children.push(
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
    })
  );
  
  if (extractedResources.length > 0) {
    extractedResources.forEach(resource => {
      children.push(
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun(resource)]
        })
      );
    });
  } else {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "Resources will be documented as the learning program develops.", italics: true })]
      })
    );
  }
  
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("6.2 Planned Resources for Next Learning Period")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun("We will continue using many of the resources that have proven effective, supplemented with additional materials as learning needs develop.")]
    })
  );
  
  // Add planned resources if provided
  const plannedResourcesText = parsedFuturePlans.plannedResources || '';
  if (plannedResourcesText && plannedResourcesText.trim() !== '') {
    const plannedList = plannedResourcesText.split(/[,\n]+/).filter(r => r.trim());
    plannedList.forEach(resource => {
      children.push(
        new Paragraph({
          numbering: { reference: "bullet-list", level: 0 },
          children: [new TextRun(resource.trim())]
        })
      );
    });
  } else {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: "Planned resources will be identified based on emerging learning interests and needs.", italics: true })]
      })
    );
  }

  // Create the document
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "bullet-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 360 }
                }
              }
            }
          ]
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440
          }
        }
      },
      children: children
    }]
  });

  return doc;
}

/**
 * Generate learning area overview sections
 */
function generateLearningAreaOverviews(learningAreaOverviews, evidenceByArea, curriculumOutcomes, yearLevel, curriculumTermCap, childName) {
  const sections = [];
  
  // Ensure evidenceByArea is an object
  if (!evidenceByArea || typeof evidenceByArea !== 'object') {
    evidenceByArea = {};
  }
  
  // Define standard learning areas for NSW
  const standardAreas = [
    'English',
    'Mathematics', 
    'Science & Technology',
    'HSIE',
    'PDHPE',
    'Creative Arts'
  ];
  
  // Get all areas that have either overviews or evidence
  const allAreas = new Set([
    ...standardAreas,
    ...Object.keys(learningAreaOverviews || {}),
    ...Object.keys(evidenceByArea || {})
  ]);
  
  let sectionNum = 1;
  
  allAreas.forEach(area => {
    // Skip numeric or empty area names (these are artifacts from bad data)
    if (!area || area === '0' || area === '1' || /^\d+$/.test(area)) {
      return;
    }
    
    // Skip non-standard areas that have no content
    const areaEvidence = evidenceByArea[area];
    const evidenceArray = Array.isArray(areaEvidence) ? areaEvidence : 
                          (areaEvidence && typeof areaEvidence === 'object') ? Object.values(areaEvidence) : [];
    
    if (!standardAreas.includes(area) && 
        !learningAreaOverviews[area] && 
        evidenceArray.length === 0) {
      return;
    }
    
    const overview = learningAreaOverviews[area] || {};
    const evidence = evidenceArray;
    const areaOutcomes = (curriculumOutcomes || []).filter(o => 
      normalizeAreaName(o['Learning Area'] || o.learningArea) === area
    );
    
    sections.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun(`2.${sectionNum} ${area}`)]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: `${curriculumTermCap} Expectations:`, bold: true })]
      })
    );
    
    // Add stage statement or generate from outcomes
    if (overview.stageStatement) {
      sections.push(
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: overview.stageStatement, italics: true })]
        })
      );
    } else if (areaOutcomes.length > 0) {
      // Show up to 6 outcomes with full descriptions
      const outcomeDescriptions = areaOutcomes
        .slice(0, 6)
        .map(o => {
          const desc = o['Outcome Description'] || o.outcomeDescription || '';
          const code = o['Outcome Title'] || o.outcomeTitle || '';
          // Return full description, not truncated
          return desc.trim();
        })
        .filter(d => d.length > 0)
        .join('; ');
      
      sections.push(
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ 
            text: `In ${area}, ${yearLevel} students work towards outcomes including: ${outcomeDescriptions}.`,
            italics: true 
          })]
        })
      );
    } else {
      sections.push(
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ 
            text: `${yearLevel} students develop skills and knowledge in ${area} through engaging activities and experiences.`,
            italics: true 
          })]
        })
      );
    }
    
    // Add progress summary
    if (evidence.length > 0) {
      sections.push(
        new Paragraph({
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: "Progress Summary:", bold: true })]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(`${evidence.length} learning evidence ${evidence.length === 1 ? 'entry' : 'entries'} documented for this area during the reporting period.`)]
        })
      );
    } else {
      sections.push(
        new Paragraph({
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: "Evidence Status:", bold: true })]
        }),
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun({ 
            text: `No formal evidence was documented for ${area} during this reporting period. Learning in this area has occurred informally through daily activities, conversations, and integrated experiences.`,
            italics: true 
          })]
        })
      );
    }
    
    sectionNum++;
  });
  
  return sections;
}

/**
 * Generate detailed evidence sections
 */
function generateEvidenceSections(evidenceByArea, curriculumOutcomes) {
  const sections = [];
  
  if (!evidenceByArea || typeof evidenceByArea !== 'object' || Object.keys(evidenceByArea).length === 0) {
    sections.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ 
          text: "Detailed evidence will be documented as learning activities are recorded.",
          italics: true 
        })]
      })
    );
    return sections;
  }
  
  let sectionNum = 1;
  
  Object.entries(evidenceByArea).forEach(([area, evidenceList]) => {
    // Ensure evidenceList is an array
    if (!evidenceList) return;
    if (!Array.isArray(evidenceList)) {
      // Try to convert to array if it's an object
      if (typeof evidenceList === 'object') {
        evidenceList = Object.values(evidenceList);
      } else {
        return;
      }
    }
    if (evidenceList.length === 0) return;
    
    sections.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun(`3.${sectionNum} ${area}`)]
      })
    );
    
    evidenceList.forEach((evidence, idx) => {
      sections.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_3,
          children: [new TextRun(evidence.title || `Evidence ${idx + 1}`)]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Date: ", bold: true }),
            new TextRun(evidence.date || 'Not specified')
          ]
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Description: ", bold: true }),
            new TextRun(evidence.description || 'No description provided.')
          ]
        })
      );
      
      // Add matched outcomes if available
      if (evidence.matchedOutcomes && evidence.matchedOutcomes.length > 0) {
        sections.push(
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: "Curriculum Outcomes Addressed:", bold: true })]
          })
        );
        
        // Handle matched outcomes - could be array of IDs or objects
        evidence.matchedOutcomes.forEach(outcome => {
          let outcomeText = '';
          if (typeof outcome === 'string') {
            // It's an ID - try to find the outcome details
            const foundOutcome = (curriculumOutcomes || []).find(o => o.ID === outcome || o.id === outcome);
            if (foundOutcome) {
              outcomeText = `${foundOutcome['Outcome Title'] || foundOutcome.outcomeTitle}: ${foundOutcome['Outcome Description'] || foundOutcome.outcomeDescription}`;
            } else {
              outcomeText = outcome; // Just show the ID
            }
          } else if (typeof outcome === 'object') {
            outcomeText = `${outcome.code || outcome['Outcome Title'] || ''}: ${outcome.description || outcome['Outcome Description'] || ''}`;
          }
          
          if (outcomeText) {
            sections.push(
              new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun(outcomeText)]
              })
            );
          }
        });
      }
      
      // Add engagement level if available
      if (evidence.engagement) {
        sections.push(
          new Paragraph({
            spacing: { before: 60, after: 120 },
            children: [
              new TextRun({ text: "Child Engagement: ", bold: true }),
              new TextRun(evidence.engagement)
            ]
          })
        );
      } else {
        sections.push(
          new Paragraph({ spacing: { after: 120 }, children: [] })
        );
      }
    });
    
    sectionNum++;
  });
  
  return sections;
}

/**
 * Normalise learning area names to standard format
 */
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
  
  const lowerArea = areaStr.toLowerCase();
  return mappings[lowerArea] || areaStr;
}

/**
 * Extract resources mentioned in evidence descriptions
 */
function extractResourcesFromEvidence(evidenceByArea) {
  const resources = new Set();
  
  // Common resource patterns to look for
  const resourcePatterns = [
    // Apps and software
    /\bapp[s]?\s+['"]?([^'".,]+)['"]?/gi,
    /\busing\s+(?:the\s+)?(?:app\s+)?['"]?([A-Z][a-zA-Z\s]+)['"]?/gi,
    // Games
    /\b(?:played|playing)\s+(?:a\s+)?(?:game\s+)?(?:called\s+)?['"]?([A-Z][a-zA-Z\s]+)['"]?/gi,
    /\bMinecraft\b/gi,
    /\bRoblox\b/gi,
    /\bLEGO\b/gi,
    /\bLego\b/gi,
    // Media
    /\b(?:watched|watching)\s+(?:the\s+)?(?:movie\s+|film\s+|show\s+)?['"]?([A-Z][a-zA-Z\s]+)['"]?/gi,
    /\bYouTube\b/gi,
    // Platforms and tools
    /\bFlight\s+Simulator\b/gi,
    /\bGoogle\b/gi,
    // Physical resources
    /\bgeodes?\b/gi,
    /\bpool\b/gi,
    /\bswimming\s+pool\b/gi,
    /\blibrary\b/gi,
    // Educational
    /\bpsychologist\b/gi,
  ];
  
  // Known resources to extract (case-insensitive matching, proper case output)
  const knownResources = {
    'minecraft': 'Minecraft (digital game)',
    'roblox': 'Roblox (digital game)',
    'lego': 'LEGO (construction toys)',
    'youtube': 'YouTube (video platform)',
    'flight simulator': 'Flight Simulator app',
    'flightradar': 'Flight tracking app',
    'pool': 'Local swimming pool',
    'swimming pool': 'Local swimming pool',
    'geode': 'Geodes (geological specimens)',
    'geodes': 'Geodes (geological specimens)',
    'psychologist': 'Child psychologist sessions',
    'the lorax': 'The Lorax (film/book)',
  };
  
  if (!evidenceByArea || typeof evidenceByArea !== 'object') {
    return [];
  }
  
  // Go through all evidence
  Object.values(evidenceByArea).forEach(evidenceList => {
    if (!Array.isArray(evidenceList)) return;
    
    evidenceList.forEach(evidence => {
      const description = (evidence.description || '').toLowerCase();
      const title = (evidence.title || '').toLowerCase();
      const combined = description + ' ' + title;
      
      // Check for known resources
      Object.entries(knownResources).forEach(([pattern, resourceName]) => {
        if (combined.includes(pattern.toLowerCase())) {
          resources.add(resourceName);
        }
      });
      
      // Extract apps mentioned with quotes
      const appMatches = combined.match(/app\s+['"]([^'"]+)['"]/gi);
      if (appMatches) {
        appMatches.forEach(match => {
          const appName = match.replace(/app\s+['"]?/i, '').replace(/['"]?$/, '');
          if (appName.length > 2) {
            resources.add(`${appName} (app)`);
          }
        });
      }
    });
  });
  
  return Array.from(resources).sort();
}

module.exports = { generatePortfolio };
