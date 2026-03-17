const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, 
        LevelFormat, PageBreak, Table, TableRow, TableCell, WidthType, BorderStyle,
        ImageRun, Header, Footer, PageNumber } = require('docx');
const { put } = require('@vercel/blob');


// ============================================================
// HELPER FUNCTIONS
// ============================================================

function toArray(val) {
  if (Array.isArray(val)) return val;
  if (!val) return [];
  if (typeof val === 'object') return Object.values(val).filter(v => v !== null && v !== undefined);
  return [];
}

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

function buildEvidenceByArea(evidenceEntries) {
  const byArea = {};
  evidenceEntries.forEach(entry => {
    let areas = [];
    const rawAreas = entry['Learning Areas'] || entry.learningAreas || entry.Areas || entry.areas;
    if (rawAreas) {
      if (Array.isArray(rawAreas)) areas = rawAreas;
      else if (typeof rawAreas === 'string') areas = rawAreas.split(',').map(a => a.trim());
    }
    areas = [...new Set(areas.map(normalizeAreaName).filter(a => a && a !== 'Other'))];
    if (areas.length === 0) areas = ['Other'];

    const evidenceObj = {
      title: entry.Title || entry.title || 'Untitled',
      date: entry.Date || entry.date || '',
      description: entry['What Happened?'] || entry.whatHappened || entry.description || '',
      engagement: entry['Child Engagement'] || entry.childEngagement || entry.engagement || '',
      matchedOutcomes: entry['Outcome Code Rollup (from Matched Outcomes 3)'] ||
                       entry['Matched Outcomes 3'] || entry.matchedOutcomes || [],
      attachments: entry.Attachments || entry.attachments || []
    };

    areas.forEach(area => {
      if (!byArea[area]) byArea[area] = [];
      byArea[area].push(evidenceObj);
    });
  });
  return byArea;
}

// ============================================================
// PORTFOLIO GENERATOR
// ============================================================

function generatePortfolio(portfolioData) {
  const {
    childName = 'Child',
    yearLevel = 'Stage 2',
    reportingPeriod = 'Current Period',
    parentName,
    parentname,
    state = 'NSW',
    curriculum = 'NSW Syllabus',
    learningAreaOverviews = {},
    evidenceByArea = {},
    progressAssessment = {},
    futurePlans = {},
    curriculumOutcomes = [],
    evidenceEntries = [],
    aiProgressSummaries = {},
    enhancedProgressAssessment = {},
    enhancedFuturePlansOverview = null,
    logoBuffer = null
  } = portfolioData;

  const finalParentName = parentName || parentname || 'Parent/Carer';

  let parsedFuturePlans = futurePlans;
  if (typeof futurePlans === 'string') {
    try {
      parsedFuturePlans = JSON.parse(futurePlans.replace(/\\n/g, '\n').replace(/\\"/g, '"'));
    } catch (e) {
      parsedFuturePlans = { overview: futurePlans };
    }
  }
  console.log('Parsed futurePlans:', parsedFuturePlans);

  let parsedProgressAssessment = progressAssessment;
  if (typeof progressAssessment === 'string') {
    try { parsedProgressAssessment = JSON.parse(progressAssessment); }
    catch (e) { parsedProgressAssessment = {}; }
  }

  const curriculumTermCap = state === 'NSW' ? 'Syllabus' : 'Curriculum';
  const curriculumTerm = state === 'NSW' ? 'syllabus' : 'curriculum';
  const currentDate = new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' });

  const children = [];

  // TITLE PAGE
  if (logoBuffer) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 200 },
      children: [new ImageRun({ data: logoBuffer, transformation: { width: 150, height: 150 }, type: 'jpeg' })]
    }));
  }

  children.push(
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: logoBuffer ? 100 : 400, after: 200 }, children: [new TextRun({ text: "Home Education Learning Portfolio", bold: true, size: 56, font: "Aptos" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: childName, size: 48, font: "Aptos" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: yearLevel, size: 36, font: "Aptos" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: reportingPeriod, size: 36, font: "Aptos" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: `Prepared by: ${finalParentName}`, size: 28, font: "Aptos" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: `Date: ${currentDate}`, size: 28, font: "Aptos" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 }, shading: { fill: "F5F5F5" }, children: [new TextRun({ text: "This portfolio was generated using OneHome Education's automated portfolio system. Learning evidence and assessments were provided by the parent/carer and enhanced using AI assistance.", size: 18, italics: true, font: "Aptos" })] }),
    new Paragraph({ children: [new PageBreak()] })
  );

  // SECTION 1
  children.push(
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. Learning Program Overview")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.1 Syllabus Framework")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(`This learning portfolio demonstrates ${childName}'s educational progress during ${reportingPeriod}. Our home education program aligns with the ${curriculum} and covers all key learning areas required under ${state} homeschooling regulations.`)] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.2 Compliance with Disability Standards for Education 2005")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun("This educational program has been developed in accordance with the Disability Standards for Education 2005 (Cth), which ensure that students with disability are able to access and participate in education on the same basis as students without disability.")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(`${childName} has a neurodivergent learning profile, and our program incorporates reasonable adjustments as outlined in the Standards.`)] }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Key adjustments implemented:", bold: true })] })
  );

  ["Flexible pacing that honours the child's autonomy and reduces demand-related anxiety",
   "Collaborative approach to learning activities, allowing the child to maintain a sense of control and choice",
   "Integration of special interests and preferred learning modalities to enhance engagement",
   "Low-demand presentation of learning opportunities that reduces pressure while maintaining educational rigour",
   "Recognition that anxiety and overwhelm are communication, not misbehaviour, requiring adaptive responses"
  ].forEach(adj => children.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun(adj)] })));

  children.push(
    new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun(`These adjustments are fundamental to our educational approach and enable ${childName} to demonstrate learning in ways that respect neurodivergent learning patterns.`)] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.3 Educational Philosophy and Approach")] }),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun(`Our home education program recognises that meaningful learning occurs when children feel safe, autonomous, and connected. We provide a rich learning environment that allows natural curiosity to drive engagement with ${curriculumTerm} content.`)] })
  );

  // SECTION 2
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. Learning Areas Overview")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(`The following provides an overview of ${curriculum} expectations for ${yearLevel} students in each learning area, along with a summary of ${childName}'s progress.`)] })
  );
  children.push(...generateLearningAreaOverviews(learningAreaOverviews, evidenceByArea, curriculumOutcomes, yearLevel, curriculumTermCap, childName, aiProgressSummaries));

  // SECTION 3
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. Detailed Learning Evidence")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun("Our learning activities integrate learning into everyday life experiences, encouraging self-directed inquiry, real-world problem-solving, and project-based activities.")] })
  );
  children.push(...generateEvidenceSectionsFlat(evidenceByArea, curriculumOutcomes, state));

  // SECTION 4
  const finalCognitive = enhancedProgressAssessment.cognitive || parsedProgressAssessment.cognitive || "No assessment provided.";
  const finalSocial = enhancedProgressAssessment.social || parsedProgressAssessment.social || "No assessment provided.";
  const finalEmotional = enhancedProgressAssessment.emotional || parsedProgressAssessment.emotional || "No assessment provided.";
  const finalPhysical = enhancedProgressAssessment.physical || parsedProgressAssessment.physical || "No assessment provided.";

  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. Parent Assessment of Progress")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.1 Cognitive Development")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(finalCognitive)] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.2 Social Development")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(finalSocial)] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.3 Emotional Development")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(finalEmotional)] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.4 Physical Development")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun(finalPhysical)] })
  );

  // SECTION 5
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. Future Learning Plans")] })
  );

  const futureOverviewText = enhancedFuturePlansOverview || parsedFuturePlans.overview || '';
  children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun(futureOverviewText || 'No future plans overview provided.')] }));

  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("5.1 Learning Goals")] }));
  const futureGoals = parsedFuturePlans.goals || '';
  if (futureGoals.trim()) {
    const goalsList = futureGoals.includes('\n') ? futureGoals.split(/[\n\r]+/).filter(g => g.trim()) : futureGoals.split(/(?:\.?\s+)(?=To\s)/i).filter(g => g.trim());
    goalsList.forEach(goal => children.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun(goal.trim())] })));
  } else {
    children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "No learning goals specified.", italics: true })] }));
  }

  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("5.2 Planned Strategies")] }));
  const futureStrategies = parsedFuturePlans.strategies || '';
  if (futureStrategies.trim()) {
    const strategiesList = futureStrategies.includes('\n') ? futureStrategies.split(/[\n\r]+/).filter(s => s.trim()) : futureStrategies.split(/,\s+(?=[A-Z])/).filter(s => s.trim());
    if (strategiesList.length > 1) {
      strategiesList.forEach(s => children.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun(s.trim())] })));
    } else {
      children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun(futureStrategies)] }));
    }
  } else {
    children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "No strategies specified.", italics: true })] }));
  }

  const plannedResources = parsedFuturePlans.plannedResources || '';
  if (plannedResources.trim()) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("5.3 Planned Resources")] }));
    const resList = plannedResources.includes(',') || plannedResources.includes('\n')
      ? plannedResources.split(/[,\n]+/).map(r => r.trim()).filter(r => r.length > 0)
      : [plannedResources.trim()];
    resList.forEach(r => children.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun(r)] })));
  }

  // SECTION 6
  const extractedResources = extractResourcesFromEvidence(evidenceByArea);
  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("6. Resources for Learning")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("6.1 Resources Used During This Period")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun("The following resources supported learning across syllabus areas during this reporting period:")] })
  );
  if (extractedResources.length > 0) {
    extractedResources.forEach(r => children.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun(r)] })));
  } else {
    children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Resources will be documented as the learning program develops.", italics: true })] }));
  }

  children.push(
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("6.2 Planned Resources for Next Learning Period")] }),
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun("We will continue using many of the resources that have proven effective, supplemented with additional materials as learning needs develop.")] })
  );

  if (plannedResources.trim()) {
    const plannedList = plannedResources.includes(',') || plannedResources.includes('\n')
      ? plannedResources.split(/[,\n]+/).map(r => r.trim()).filter(r => r.length > 0)
      : [plannedResources.trim()];
    plannedList.forEach(r => children.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun(r)] })));
  }

  return new Document({
    styles: {
      default: {
        document: { run: { font: "Aptos", size: 24 } },
        heading1: { run: { font: "Aptos", size: 32, bold: true, color: "2E74B5" }, paragraph: { spacing: { before: 240, after: 120 } } },
        heading2: { run: { font: "Aptos", size: 26, bold: true, color: "2E74B5" }, paragraph: { spacing: { before: 200, after: 80 } } },
        heading3: { run: { font: "Aptos", size: 24, bold: true }, paragraph: { spacing: { before: 160, after: 60 } } }
      }
    },
    numbering: {
      config: [{
        reference: "bullet-list",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { font: "Aptos" } } }]
      }]
    },
    sections: [{
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "OneHome Education | Affirming Connections | ABN: 57 886 895 482", size: 16, font: "Aptos", color: "666666" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "www.affirmingconnections.com.au", size: 16, font: "Aptos", color: "666666" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: "Page ", size: 16, font: "Aptos", color: "666666" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Aptos", color: "666666" }), new TextRun({ text: " of ", size: 16, font: "Aptos", color: "666666" }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: "Aptos", color: "666666" })] })
          ]
        })
      },
      children: children
    }]
  });
}

// ============================================================
// LEARNING AREA OVERVIEWS
// ============================================================

function generateLearningAreaOverviews(learningAreaOverviews, evidenceByArea, curriculumOutcomes, yearLevel, curriculumTermCap, childName, aiProgressSummaries = {}) {
  const sections = [];
  if (!evidenceByArea || typeof evidenceByArea !== 'object') evidenceByArea = {};

  const standardAreas = ['English', 'Mathematics', 'Science & Technology', 'HSIE', 'PDHPE', 'Creative Arts'];
  const allAreas = new Set([...standardAreas, ...Object.keys(learningAreaOverviews || {}), ...Object.keys(evidenceByArea || {})]);

  let sectionNum = 1;
  allAreas.forEach(area => {
    if (!area || /^\d+$/.test(area)) return;
    const areaEvidence = evidenceByArea[area];
    const evidenceArray = Array.isArray(areaEvidence) ? areaEvidence : (areaEvidence && typeof areaEvidence === 'object') ? Object.values(areaEvidence) : [];
    if (!standardAreas.includes(area) && !(learningAreaOverviews || {})[area] && evidenceArray.length === 0) return;

    const overview = (learningAreaOverviews || {})[area] || {};
    const areaOutcomes = (curriculumOutcomes || []).filter(o => normalizeAreaName(o['Learning Area'] || o.learningArea) === area);
    const aiSummary = aiProgressSummaries[area];

    sections.push(
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`2.${sectionNum} ${area}`)] }),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `${curriculumTermCap} Expectations:`, bold: true })] })
    );

    if (overview.stageStatement) {
      sections.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: overview.stageStatement, italics: true })] }));
    } else if (areaOutcomes.length > 0) {
      const outcomeDescriptions = areaOutcomes.slice(0, 6).map(o => (o['Outcome Description'] || o.outcomeDescription || '').trim()).filter(d => d.length > 0).join('; ');
      sections.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `In ${area}, ${yearLevel} students work towards outcomes including: ${outcomeDescriptions}.`, italics: true })] }));
    } else {
      sections.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `${yearLevel} students develop skills and knowledge in ${area} through engaging activities and experiences.`, italics: true })] }));
    }

    sections.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Progress Summary:", bold: true })] }));

    if (aiSummary) {
      sections.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun(aiSummary)] }));
    } else if (evidenceArray.length > 0) {
      sections.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun(`${evidenceArray.length} learning evidence ${evidenceArray.length === 1 ? 'entry' : 'entries'} documented for this area during the reporting period.`)] }));
    } else {
      sections.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: `No formal evidence was documented for ${area} during this reporting period. Learning in this area has occurred informally through daily activities, conversations, and integrated experiences.`, italics: true })] }));
    }

    sectionNum++;
  });

  return sections;
}

// ============================================================
// EVIDENCE SECTIONS
// ============================================================

function generateEvidenceSectionsFlat(evidenceByArea, curriculumOutcomes, state = 'NSW') {
  const sections = [];
  const outcomesLabel = state === 'NSW' ? 'Syllabus Outcomes Addressed:' : 'Curriculum Outcomes Addressed:';

  if (!evidenceByArea || typeof evidenceByArea !== 'object' || Object.keys(evidenceByArea).length === 0) {
    sections.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Detailed evidence will be documented as learning activities are recorded.", italics: true })] }));
    return sections;
  }

  const seenTitles = new Set();
  const uniqueEvidence = [];

  Object.entries(evidenceByArea).forEach(([area, evidenceList]) => {
    if (!Array.isArray(evidenceList)) {
      if (typeof evidenceList === 'object') evidenceList = Object.values(evidenceList);
      else return;
    }
    evidenceList.forEach(evidence => {
      const title = evidence.title || 'Untitled';
      if (!seenTitles.has(title)) {
        seenTitles.add(title);
        uniqueEvidence.push({ ...evidence, primaryArea: area });
      }
    });
  });

  uniqueEvidence.sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0));

  uniqueEvidence.forEach((evidence, idx) => {
    if (idx > 0) {
      sections.push(new Paragraph({ spacing: { before: 240, after: 240 }, border: { bottom: { color: "CCCCCC", space: 1, style: BorderStyle.SINGLE, size: 6 } }, children: [] }));
    }

    sections.push(
      new Paragraph({ spacing: { before: 120, after: 80 }, shading: { fill: "E8F4FC" }, children: [new TextRun({ text: `${idx + 1}. ${evidence.title || `Evidence ${idx + 1}`}`, bold: true, size: 26, font: "Aptos" })] }),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Date: ", bold: true, font: "Aptos" }), new TextRun({ text: evidence.date || 'Not specified', font: "Aptos" })] }),
      new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "Description: ", bold: true, font: "Aptos" }), new TextRun({ text: evidence.description || 'No description provided.', font: "Aptos" })] })
    );

    const matchedOutcomes = evidence.matchedOutcomes;
    if (matchedOutcomes && (Array.isArray(matchedOutcomes) ? matchedOutcomes.length > 0 : matchedOutcomes.toString().trim() !== '')) {
      let outcomesList = [];
      if (typeof matchedOutcomes === 'string') {
        outcomesList = matchedOutcomes.split(',').map(o => o.trim()).filter(o => o.length > 0);
      } else if (Array.isArray(matchedOutcomes)) {
        matchedOutcomes.forEach(outcome => {
          if (typeof outcome === 'string' && !(outcome.startsWith('rec') && outcome.length === 17)) outcomesList.push(outcome);
          else if (typeof outcome === 'object') {
            const text = `${outcome.code || outcome['Outcome Title'] || ''}: ${outcome.description || outcome['Outcome Description'] || ''}`;
            if (text.trim() !== ':') outcomesList.push(text);
          }
        });
      }

      const nswPattern = /^(EN|MA|ST|HS|PH|CA)\d-[A-Z]{2,6}-\d{2}$/;
      const acPattern = /^AC9[A-Z]\d[A-Z]{1,2}\d{2}$/;
      const filteredOutcomes = outcomesList.map(t => {
        const nswMatch = t.match(/([A-Z]{2,3}\d?-[A-Z]{2,6}-\d{2})/);
        if (nswMatch) return nswMatch[1];
        const acMatch = t.match(/(AC9[A-Z]\d[A-Z]{1,2}\d{2})/);
        if (acMatch) return acMatch[1];
        return t.trim();
      }).filter(code => nswPattern.test(code) || acPattern.test(code));

      if (filteredOutcomes.length > 0) {
        sections.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: outcomesLabel, bold: true, font: "Aptos" })] }));
        filteredOutcomes.forEach(t => sections.push(new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: t, font: "Aptos" })] })));
      }
    }

    if (evidence.engagement) {
      sections.push(new Paragraph({ spacing: { before: 60, after: 80 }, children: [new TextRun({ text: "Child Engagement: ", bold: true, font: "Aptos" }), new TextRun({ text: evidence.engagement, font: "Aptos" })] }));
    }

    const attachments = evidence.attachments;
    if (attachments && Array.isArray(attachments) && attachments.length > 0) {
      sections.push(new Paragraph({ spacing: { before: 80, after: 60 }, children: [new TextRun({ text: "Evidence Photos:", bold: true, font: "Aptos" })] }));
      attachments.forEach(att => {
        if (att.buffer) {
          try {
            let imageType = 'jpeg';
            if (att.mimeType) {
              if (att.mimeType.includes('png')) imageType = 'png';
              else if (att.mimeType.includes('gif')) imageType = 'gif';
            }
            let width = Math.min(att.width || 400, 400);
            let height = att.height || 300;
            if (att.width > 400) height = Math.round(height * (400 / att.width));
            if (height > 400) { width = Math.round(width * (400 / height)); height = 400; }
            sections.push(new Paragraph({ spacing: { after: 80 }, children: [new ImageRun({ data: att.buffer, transformation: { width, height }, type: imageType })] }));
          } catch (e) { console.error('Image embed error:', e.message); }
        }
      });
    }

    sections.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  });

  return sections;
}

// ============================================================
// RESOURCES EXTRACTOR
// ============================================================

function extractResourcesFromEvidence(evidenceByArea) {
  const resources = new Set();
  const addedLower = new Set();
  const knownResources = {
    'minecraft': 'Minecraft (digital game)', 'roblox': 'Roblox (digital game)',
    'lego': 'LEGO (construction toys)', 'youtube': 'YouTube (video platform)',
    'flight simulator': 'Flight Simulator (app)', 'flightradar': 'Flight tracking app',
    'pool': 'Local swimming pool', 'swimming pool': 'Local swimming pool',
    'geode': 'Geodes (geological specimens)', 'geodes': 'Geodes (geological specimens)',
    'psychologist': 'Child psychologist sessions', 'the lorax': 'The Lorax (film/book)',
    'watercolour': 'Watercolour paints', 'watercolor': 'Watercolour paints',
    'acrylic paint': 'Acrylic paints', 'texta': 'Textas/markers', 'textas': 'Textas/markers',
    'marker': 'Textas/markers', 'pencil': 'Pencils', 'pencils': 'Pencils',
    'crayon': 'Crayons', 'crayons': 'Crayons', 'hot glue': 'Hot glue gun',
    'bead': 'Beads', 'beads': 'Beads', 'paint': 'Paints', 'book': 'Books',
    'library': 'Library', 'ipad': 'iPad/tablet', 'tablet': 'iPad/tablet', 'computer': 'Computer'
  };

  if (!evidenceByArea || typeof evidenceByArea !== 'object') return [];

  Object.values(evidenceByArea).forEach(evidenceList => {
    if (!Array.isArray(evidenceList)) return;
    evidenceList.forEach(evidence => {
      const combined = ((evidence.description || '') + ' ' + (evidence.title || '')).toLowerCase();
      Object.entries(knownResources).forEach(([pattern, name]) => {
        if (combined.includes(pattern.toLowerCase())) {
          const lower = name.toLowerCase();
          if (!addedLower.has(lower)) { addedLower.add(lower); resources.add(name); }
        }
      });
    });
  });

  return Array.from(resources).sort();
}

// ============================================================
// VERCEL HANDLER
// ============================================================

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method === 'GET') {
    return res.status(200).json({ status: 'healthy', service: 'OneHome Education Portfolio Generator', version: '4.0.1', mode: 'self-contained' });
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

    // Build evidenceByArea from evidenceEntries
    const existingByArea = portfolioData.evidenceByArea;
    const hasValidByArea = existingByArea && typeof existingByArea === 'object' && !Array.isArray(existingByArea) && Object.keys(existingByArea).length > 0;

    if (!hasValidByArea && portfolioData.evidenceEntries.length > 0) {
      console.log('Building evidenceByArea from', portfolioData.evidenceEntries.length, 'evidence entries');
      portfolioData.evidenceByArea = buildEvidenceByArea(portfolioData.evidenceEntries);
      console.log('Built areas:', Object.keys(portfolioData.evidenceByArea));
    } else if (Array.isArray(existingByArea)) {
      portfolioData.evidenceByArea = buildEvidenceByArea(existingByArea);
    } else if (hasValidByArea) {
      Object.keys(portfolioData.evidenceByArea).forEach(key => {
        portfolioData.evidenceByArea[key] = toArray(portfolioData.evidenceByArea[key]);
      });
    }

    // learningAreaOverviews should be an object not array
    if (Array.isArray(portfolioData.learningAreaOverviews) || !portfolioData.learningAreaOverviews) {
      portfolioData.learningAreaOverviews = {};
    }

    const doc = generatePortfolio(portfolioData);
    const buffer = await Packer.toBuffer(doc);
    console.log('DOCX generated, size:', buffer.length, 'bytes');

    const safeName = (portfolioData.childName || 'Child').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 50);
    const safePeriod = (portfolioData.reportingPeriod || 'Portfolio').replace(/[^a-zA-Z0-9]/g, '-').substring(0, 30);
    const filename = `${safeName}-Portfolio-${safePeriod}.docx`;

    const blob = await put(`portfolios/${filename}`, buffer, {
      access: 'public',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    console.log('Uploaded to Blob:', blob.url);

    return res.status(200).json({ success: true, filename, url: blob.url, fileSize: buffer.length });

  } catch (error) {
    console.error('Error generating portfolio:', error.message);
    console.error(error.stack);
    return res.status(500).json({ success: false, error: error.message });
  }
};
