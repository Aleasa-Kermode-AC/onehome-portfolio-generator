const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, 
        LevelFormat, PageBreak, Table, TableRow, TableCell, WidthType, BorderStyle,
        ImageRun } = require('docx');

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
    evidenceEntries = [],
    aiProgressSummaries = {},           // AI-generated progress summaries per area
    enhancedProgressAssessment = {},    // AI-enhanced parent assessments
    enhancedFuturePlansOverview = null  // AI-enhanced future plans overview
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
    new Paragraph({ children: [new PageBreak()] }),
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
    childName,
    aiProgressSummaries  // Pass AI summaries
  );
  children.push(...learningAreaSections);

  // === SECTION 3: DETAILED LEARNING EVIDENCE ===
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("3. Detailed Learning Evidence")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun("Our learning activities integrate learning into everyday life experiences, encouraging self-directed inquiry, real-world problem-solving, and project-based activities that emerge from our child's passions.")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun("In our homeschool program, we use flexible learning models that ensure all key syllabus areas are covered, but not always through a traditional, linear sequence. Instead of progressing subject-by-subject in a fixed order, learning happens naturally across multiple areas at once, often integrated into real-world projects, play, or child-led inquiry. This is how we create an interdisciplinary approach to education.")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(`This approach allows ${childName} to engage deeply and meaningfully with the syllabus in ways that suit their developmental stage, interests, and learning style. It also aligns with the principles of our flexible learning model and inclusive education practices, ensuring that learning remains accessible, engaging, and personalised.`)]
    })
  );
  
  // Generate evidence sections - NO LONGER grouped by subject area
  const evidenceSections = generateEvidenceSectionsFlat(evidenceByArea, curriculumOutcomes);
  children.push(...evidenceSections);

  // === SECTION 4: PARENT ASSESSMENT OF PROGRESS ===
  // Use enhanced assessments if available, fall back to original
  const finalCognitive = enhancedProgressAssessment.cognitive || parsedProgressAssessment.cognitive || "No assessment provided.";
  const finalSocial = enhancedProgressAssessment.social || parsedProgressAssessment.social || "No assessment provided.";
  const finalEmotional = enhancedProgressAssessment.emotional || parsedProgressAssessment.emotional || "No assessment provided.";
  const finalPhysical = enhancedProgressAssessment.physical || parsedProgressAssessment.physical || "No assessment provided.";
  
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
      children: [new TextRun(finalCognitive)]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("4.2 Social Development")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(finalSocial)]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("4.3 Emotional Development")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(finalEmotional)]
    }),
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("4.4 Physical Development")]
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun(finalPhysical)]
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
  
  // Overview - use AI enhanced version if available
  const futureOverviewText = enhancedFuturePlansOverview || parsedFuturePlans.overview || '';
  if (futureOverviewText && futureOverviewText.trim() !== '') {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun(futureOverviewText)]
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
    // Split goals by newlines, or by "To " at start of sentences (common goal pattern)
    let goalsList = [];
    
    if (futureGoals.includes('\n')) {
      goalsList = futureGoals.split(/[\n\r]+/).filter(g => g.trim());
    } else {
      // Split by ". To" or just "To" at the start of a new goal
      // Pattern: "To find fiction. To build skills" -> ["To find fiction", "To build skills"]
      goalsList = futureGoals.split(/(?:\.?\s+)(?=To\s)/i).filter(g => g.trim());
    }
    
    // Always use bullet points for goals
    if (goalsList.length > 0) {
      goalsList.forEach(goal => {
        children.push(
          new Paragraph({
            numbering: { reference: "bullet-list", level: 0 },
            children: [new TextRun(goal.trim())]
          })
        );
      });
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
      children: [new TextRun("The following resources supported learning across syllabus areas during this reporting period:")]
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
    styles: {
      default: {
        document: {
          run: {
            font: "Aptos",
            size: 24 // 12pt
          }
        },
        heading1: {
          run: {
            font: "Aptos",
            size: 32,
            bold: true,
            color: "2E74B5"
          },
          paragraph: {
            spacing: { before: 240, after: 120 }
          }
        },
        heading2: {
          run: {
            font: "Aptos",
            size: 26,
            bold: true,
            color: "2E74B5"
          },
          paragraph: {
            spacing: { before: 200, after: 80 }
          }
        },
        heading3: {
          run: {
            font: "Aptos",
            size: 24,
            bold: true
          },
          paragraph: {
            spacing: { before: 160, after: 60 }
          }
        }
      }
    },
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
                },
                run: {
                  font: "Aptos"
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
function generateLearningAreaOverviews(learningAreaOverviews, evidenceByArea, curriculumOutcomes, yearLevel, curriculumTermCap, childName, aiProgressSummaries = {}) {
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
    
    // Get AI summary if available
    const aiSummary = aiProgressSummaries[area];
    
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
    
    // Add progress summary - use AI summary if available, otherwise show count
    sections.push(
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Progress Summary:", bold: true })]
      })
    );
    
    if (aiSummary) {
      // Use AI-generated summary
      sections.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(aiSummary)]
        })
      );
    } else if (evidence.length > 0) {
      // Fall back to count-based summary
      sections.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [new TextRun(`${evidence.length} learning evidence ${evidence.length === 1 ? 'entry' : 'entries'} documented for this area during the reporting period.`)]
        })
      );
    } else {
      sections.push(
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
 * Generate flat evidence sections without duplicating entries across subject areas
 * Each evidence entry appears once with all its associated outcomes
 */
function generateEvidenceSectionsFlat(evidenceByArea, curriculumOutcomes) {
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
  
  // Collect all unique evidence entries (avoid duplicates)
  const seenTitles = new Set();
  const uniqueEvidence = [];
  
  Object.entries(evidenceByArea).forEach(([area, evidenceList]) => {
    if (!Array.isArray(evidenceList)) {
      if (typeof evidenceList === 'object') {
        evidenceList = Object.values(evidenceList);
      } else {
        return;
      }
    }
    
    evidenceList.forEach(evidence => {
      const title = evidence.title || 'Untitled';
      if (!seenTitles.has(title)) {
        seenTitles.add(title);
        uniqueEvidence.push({
          ...evidence,
          primaryArea: area
        });
      }
    });
  });
  
  // Sort by date (most recent first)
  uniqueEvidence.sort((a, b) => {
    const dateA = new Date(a.date || 0);
    const dateB = new Date(b.date || 0);
    return dateB - dateA;
  });
  
  // Generate sections for each unique evidence entry
  uniqueEvidence.forEach((evidence, idx) => {
    // Add a horizontal line separator before each entry (except the first)
    if (idx > 0) {
      sections.push(
        new Paragraph({
          spacing: { before: 240, after: 240 },
          border: {
            bottom: {
              color: "CCCCCC",
              space: 1,
              style: BorderStyle.SINGLE,
              size: 6
            }
          },
          children: []
        })
      );
    }
    
    // Evidence title with background styling
    sections.push(
      new Paragraph({
        spacing: { before: 120, after: 80 },
        shading: {
          fill: "E8F4FC"
        },
        children: [
          new TextRun({ 
            text: `${idx + 1}. ${evidence.title || `Evidence ${idx + 1}`}`,
            bold: true,
            size: 26,
            font: "Aptos"
          })
        ]
      })
    );
    
    // Date
    sections.push(
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Date: ", bold: true, font: "Aptos" }),
          new TextRun({ text: evidence.date || 'Not specified', font: "Aptos" })
        ]
      })
    );
    
    // Description
    sections.push(
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({ text: "Description: ", bold: true, font: "Aptos" }),
          new TextRun({ text: evidence.description || 'No description provided.', font: "Aptos" })
        ]
      })
    );
    
    // Add matched outcomes if available (filtered to NSW only)
    const matchedOutcomes = evidence.matchedOutcomes;
    if (matchedOutcomes && (Array.isArray(matchedOutcomes) ? matchedOutcomes.length > 0 : matchedOutcomes.toString().trim() !== '')) {
      
      let outcomesList = [];
      if (typeof matchedOutcomes === 'string') {
        outcomesList = matchedOutcomes.split(',').map(o => o.trim()).filter(o => o.length > 0);
      } else if (Array.isArray(matchedOutcomes)) {
        matchedOutcomes.forEach(outcome => {
          if (typeof outcome === 'string') {
            if (outcome.startsWith('rec') && outcome.length === 17) return;
            outcomesList.push(outcome);
          } else if (typeof outcome === 'object') {
            const text = `${outcome.code || outcome['Outcome Title'] || ''}: ${outcome.description || outcome['Outcome Description'] || ''}`;
            if (text.trim() !== ':') outcomesList.push(text);
          }
        });
      }
      
      // Filter to NSW syllabus codes only
      const nswPattern = /^(EN|MA|ST|HS|PH|CA)\d-[A-Z]{2,6}-\d{2}$/;
      const filteredOutcomes = outcomesList
        .map(outcomeText => {
          let displayText = outcomeText.trim();
          const codeMatch = displayText.match(/([A-Z]{2,3}\d?-[A-Z]{2,6}-\d{2})/);
          if (codeMatch) displayText = codeMatch[1];
          return displayText;
        })
        .filter(code => nswPattern.test(code));
      
      if (filteredOutcomes.length > 0) {
        sections.push(
          new Paragraph({
            spacing: { before: 60, after: 60 },
            children: [new TextRun({ text: "Syllabus Outcomes Addressed:", bold: true, font: "Aptos" })]
          })
        );
        
        filteredOutcomes.forEach(displayText => {
          sections.push(
            new Paragraph({
              numbering: { reference: "bullet-list", level: 0 },
              children: [new TextRun({ text: displayText, font: "Aptos" })]
            })
          );
        });
      }
    }
    
    // Add engagement level if available
    if (evidence.engagement) {
      sections.push(
        new Paragraph({
          spacing: { before: 60, after: 80 },
          children: [
            new TextRun({ text: "Child Engagement: ", bold: true, font: "Aptos" }),
            new TextRun({ text: evidence.engagement, font: "Aptos" })
          ]
        })
      );
    }
    
    // Add attachments/photos if available
    const attachments = evidence.attachments;
    if (attachments && Array.isArray(attachments) && attachments.length > 0) {
      sections.push(
        new Paragraph({
          spacing: { before: 80, after: 60 },
          children: [
            new TextRun({ text: "Evidence Photos:", bold: true, font: "Aptos" })
          ]
        })
      );
      
      // Embed each image
      attachments.forEach((att) => {
        if (att.buffer) {
          try {
            // Determine image type from mime type
            let imageType = 'jpeg';
            if (att.mimeType) {
              if (att.mimeType.includes('png')) imageType = 'png';
              else if (att.mimeType.includes('gif')) imageType = 'gif';
              else if (att.mimeType.includes('webp')) imageType = 'png';
            }
            
            // Calculate dimensions to fit nicely in document
            let width = att.width || 400;
            let height = att.height || 300;
            const maxWidth = 400;
            const maxHeight = 400;
            
            if (width > maxWidth) {
              const ratio = maxWidth / width;
              width = maxWidth;
              height = Math.round(height * ratio);
            }
            if (height > maxHeight) {
              const ratio = maxHeight / height;
              height = maxHeight;
              width = Math.round(width * ratio);
            }
            
            sections.push(
              new Paragraph({
                spacing: { after: 80 },
                children: [
                  new ImageRun({
                    data: att.buffer,
                    transformation: {
                      width: width,
                      height: height
                    },
                    type: imageType
                  })
                ]
              })
            );
            // No filename caption - removed as requested
          } catch (imgError) {
            console.error('Error embedding image:', imgError.message);
          }
        }
      });
    }
    
    // Add bottom spacing for the entry
    sections.push(
      new Paragraph({ spacing: { after: 120 }, children: [] })
    );
  });
  
  return sections;
}

/**
 * Generate detailed evidence sections (legacy - grouped by subject)
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
      const matchedOutcomes = evidence.matchedOutcomes;
      if (matchedOutcomes && (Array.isArray(matchedOutcomes) ? matchedOutcomes.length > 0 : matchedOutcomes.toString().trim() !== '')) {
        
        // Handle outcomes - could be string (from rollup) or array
        let outcomesList = [];
        if (typeof matchedOutcomes === 'string') {
          // It's a comma-separated string from the rollup field
          // Format: "English - Stage 2 - EN2-OLC-01, English - Stage 2 - EN2-HANDW-01"
          outcomesList = matchedOutcomes.split(',').map(o => o.trim()).filter(o => o.length > 0);
        } else if (Array.isArray(matchedOutcomes)) {
          // Check if it's an array of record IDs or actual outcome objects
          matchedOutcomes.forEach(outcome => {
            if (typeof outcome === 'string') {
              // Check if it's a record ID (starts with 'rec') or actual text
              if (outcome.startsWith('rec') && outcome.length === 17) {
                // Skip record IDs - we don't have the lookup data
                return;
              }
              outcomesList.push(outcome);
            } else if (typeof outcome === 'object') {
              const text = `${outcome.code || outcome['Outcome Title'] || ''}: ${outcome.description || outcome['Outcome Description'] || ''}`;
              if (text.trim() !== ':') {
                outcomesList.push(text);
              }
            }
          });
        }
        
        // Filter to NSW syllabus codes only (EN2, MA2, ST2, HS2, PH2, CA2 patterns)
        // Skip Australian Curriculum codes (AC9...) and Victorian codes
        const nswPattern = /^(EN|MA|ST|HS|PH|CA)\d-[A-Z]{2,6}-\d{2}$/;
        
        const filteredOutcomes = outcomesList
          .map(outcomeText => {
            let displayText = outcomeText.trim();
            // Extract just the code
            const codeMatch = displayText.match(/([A-Z]{2,3}\d?-[A-Z]{2,6}-\d{2})/);
            if (codeMatch) {
              displayText = codeMatch[1];
            }
            return displayText;
          })
          .filter(code => nswPattern.test(code)); // Only NSW codes
        
        // Only show section if there are filtered outcomes
        if (filteredOutcomes.length > 0) {
          sections.push(
            new Paragraph({
              spacing: { after: 60 },
              children: [new TextRun({ text: "Syllabus Outcomes Addressed:", bold: true })]
            })
          );
          
          filteredOutcomes.forEach(displayText => {
            sections.push(
              new Paragraph({
                numbering: { reference: "bullet-list", level: 0 },
                children: [new TextRun(displayText)]
              })
            );
          });
        }
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
  const addedLowerCase = new Set(); // Track lowercase versions to avoid duplicates
  
  // Known resources to extract (case-insensitive matching, proper case output)
  const knownResources = {
    'minecraft': 'Minecraft (digital game)',
    'roblox': 'Roblox (digital game)',
    'lego': 'LEGO (construction toys)',
    'youtube': 'YouTube (video platform)',
    'flight simulator': 'Flight Simulator (app)',
    'flightradar': 'Flight tracking app',
    'pool': 'Local swimming pool',
    'swimming pool': 'Local swimming pool',
    'geode': 'Geodes (geological specimens)',
    'geodes': 'Geodes (geological specimens)',
    'psychologist': 'Child psychologist sessions',
    'the lorax': 'The Lorax (film/book)',
    // Art supplies
    'watercolour': 'Watercolour paints',
    'watercolor': 'Watercolour paints',
    'acrylic paint': 'Acrylic paints',
    'texta': 'Textas/markers',
    'textas': 'Textas/markers',
    'marker': 'Textas/markers',
    'pencil': 'Pencils',
    'pencils': 'Pencils',
    'crayon': 'Crayons',
    'crayons': 'Crayons',
    'hot glue': 'Hot glue gun',
    'bead': 'Beads',
    'beads': 'Beads',
    'paint': 'Paints',
    // Other common resources
    'book': 'Books',
    'library': 'Library',
    'ipad': 'iPad/tablet',
    'tablet': 'iPad/tablet',
    'computer': 'Computer',
  };
  
  if (!evidenceByArea || typeof evidenceByArea !== 'object') {
    return [];
  }
  
  // Helper function to add resource without duplicates
  function addResource(resourceName) {
    const lowerName = resourceName.toLowerCase();
    if (!addedLowerCase.has(lowerName)) {
      addedLowerCase.add(lowerName);
      resources.add(resourceName);
    }
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
          addResource(resourceName);
        }
      });
    });
  });
  
  return Array.from(resources).sort();
}

module.exports = { generatePortfolio };
