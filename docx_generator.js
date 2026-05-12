const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType,
  PageBreak, TabStopPosition, TabStopType, LevelFormat
} = require('docx');
const fs = require('fs');

// ─── helpers ────────────────────────────────────────────────────────────────

const BLUE_DARK  = "1F4E79";
const BLUE_MED   = "2E75B6";
const BLUE_LIGHT = "DEEAF1";
const GREEN_DARK = "375623";
const GREEN_LIGHT= "E2EFDA";
const RED_LIGHT  = "FCE4D6";
const GREY_LIGHT = "F2F2F2";
const YELLOW_LIGHT = "FFF2CC";

// docx.js generates tblGrid with w:w="100" per column regardless of actual
// widths — this is a known bug. Fix tblGrid by extracting tcW values and
// regenerating the grid columns to match. Also fix tblW type to dxa.
function fixTblGrid(xml) {
  // Split on tables
  return xml.replace(/<w:tbl>([\s\S]*?)<\/w:tbl>/g, (tblBlock) => {
    if (!tblBlock.includes('<w:tblGrid>')) return tblBlock;

    // Extract all tcW values (w:w attribute)
    const tcWidths = [];
    const tcWRegex = /<w:tcW[^>]+w:w="(\d+)"[^>]*\/>/g;
    let m;
    while ((m = tcWRegex.exec(tblBlock)) !== null) {
      tcWidths.push(m[1]);
    }
    if (tcWidths.length === 0) return tblBlock;

    // Build corrected tblGrid
    const gridCols = tcWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('');
    const fixedGrid = `<w:tblGrid>${gridCols}</w:tblGrid>`;

    // Fix tblW: replace type="pct" with type="dxa" and sum widths
    const totalW = tcWidths.reduce((s, v) => s + parseInt(v), 0);
    const fixedTblW = `<w:tblW w:type="dxa" w:w="${totalW}"/>`;

    let result = tblBlock
      .replace(/<w:tblW[^>]+>/, fixedTblW)
      .replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, fixedGrid);

    return result;
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 120 },
    shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK },
    children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 28 })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    shading: { type: ShadingType.SOLID, color: BLUE_MED, fill: BLUE_MED },
    children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 22 })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED } },
    children: [new TextRun({ text, color: BLUE_DARK, size: 22 })],
  });
}

function para(runs, spacing = {}) {
  const runArray = typeof runs === 'string'
    ? [new TextRun(runs)]
    : runs;
  return new Paragraph({
    children: runArray,
    spacing: { before: 60, after: 60, ...spacing },
  });
}

function bold(text, color = null, size = null) {
  const t = new TextRun({ text, bold: true, ...(color ? { color } : {}), ...(size ? { size } : {}) });
  return t;
}

function normal(text, color = null, italics = false) {
  return new TextRun({ text, ...(color ? { color } : {}), italics });
}

function bullet(text, level = 0) {
  return new Paragraph({
    children: [new TextRun({ text, size: 20 })],
    bullet: { level },
    spacing: { before: 40, after: 40 },
  });
}

function checkBullet(text, color = "375623") {
  return new Paragraph({
    children: [new TextRun({ text: "✅ " + text, color, size: 20 })],
    spacing: { before: 40, after: 40 },
  });
}

function crossBullet(text, color = "C00000") {
  return new Paragraph({
    children: [new TextRun({ text: "🚩 " + text, color, size: 20 })],
    spacing: { before: 40, after: 40 },
  });
}

function spacer() {
  return new Paragraph({ text: "", spacing: { before: 60, after: 60 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// 2-col info table — uses DXA twips to ensure tblGrid renders correctly
// 8640 twips = 6 inches (page content width at 1440 twips/inch)
function infoTable(rows) {
  return new Table({
    width: { size: 8640, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      left: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      right: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
    },
    rows: rows.map(([label, value], i) =>
      new TableRow({
        children: [
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: label, bold: true, color: "FFFFFF", size: 20 })],
              shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK },
              spacing: { before: 60, after: 60 },
            })],
            width: { size: 2592, type: WidthType.DXA },  // 30% of 8640
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: value, size: 20 })],
              shading: { type: ShadingType.SOLID, color: i % 2 === 0 ? GREY_LIGHT : "FFFFFF", fill: i % 2 === 0 ? GREY_LIGHT : "FFFFFF" },
              spacing: { before: 60, after: 60 },
            })],
            width: { size: 6048, type: WidthType.DXA },  // 70% of 8640
          }),
        ],
      })
    ),
  });
}

// ─── Q&A block ───────────────────────────────────────────────────────────────

function qBlock(num, question, whatToListen, probes, flagLabel) {
  const elements = [];

  // Question
  elements.push(new Paragraph({
    children: [
      new TextRun({ text: `Q${num}: `, bold: true, color: BLUE_DARK, size: 22 }),
      new TextRun({ text: question, bold: true, size: 22 }),
    ],
    spacing: { before: 200, after: 100 },
    shading: { type: ShadingType.SOLID, color: BLUE_LIGHT, fill: BLUE_LIGHT },
    border: { left: { style: BorderStyle.THICK, size: 6, color: BLUE_DARK } },
  }));

  // What to listen for
  if (whatToListen && whatToListen.length > 0) {
    elements.push(new Paragraph({
      children: [new TextRun({ text: "What to listen for:", bold: true, color: "404040", size: 20 })],
      spacing: { before: 80, after: 40 },
    }));
    whatToListen.forEach(item => {
      if (item.check) elements.push(checkBullet(item.text, "375623"));
      else if (item.cross) elements.push(crossBullet(item.text, "C00000"));
      else elements.push(bullet(item.text));
    });
  }

  // Probes
  if (probes && probes.length > 0) {
    elements.push(new Paragraph({
      children: [new TextRun({ text: "Probes:", bold: true, color: "404040", italics: true, size: 20 })],
      spacing: { before: 80, after: 40 },
    }));
    probes.forEach(p => elements.push(bullet(p)));
  }

  elements.push(spacer());
  return elements;
}

// ─── scoring table ────────────────────────────────────────────────────────────

function scoringTable(headers, rows) {
  return new Table({
    width: { size: 8640, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      left: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      right: { style: BorderStyle.SINGLE, size: 1, color: BLUE_MED },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" },
    },
    columns: headers.map(() => ({ width: { size: 2160, type: WidthType.DXA } })),
    rows: [
      new TableRow({
        children: headers.map(h => new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 20 })],
            shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK },
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 60 },
          })],
          width: { size: 2160, type: WidthType.DXA },
        })),
      }),
      ...rows.map((row, ri) =>
        new TableRow({
          children: row.map(cell => new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: cell, size: 20 })],
              shading: { type: ShadingType.SOLID, color: ri % 2 === 0 ? GREY_LIGHT : "FFFFFF", fill: ri % 2 === 0 ? GREY_LIGHT : "FFFFFF" },
              spacing: { before: 60, after: 60 },
            })],
            width: { size: 2160, type: WidthType.DXA },
          })),
        })
      ),
    ],
  });
}

// ─── decision guide ──────────────────────────────────────────────────────────

function decisionBox(text, fill, textColor) {
  return new Table({
    width: { size: 8640, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.THICK, size: 4, color: fill },
      bottom: { style: BorderStyle.THICK, size: 4, color: fill },
      left: { style: BorderStyle.THICK, size: 4, color: fill },
      right: { style: BorderStyle.THICK, size: 4, color: fill },
    },
    rows: [new TableRow({
      children: [new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text, bold: true, color: textColor, size: 22 })],
          shading: { type: ShadingType.SOLID, color: fill, fill },
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
        })],
        width: { size: 8640, type: WidthType.DXA },
      })],
    })],
  });
}

// ─── document ─────────────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 22 },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 32, bold: true, color: "FFFFFF" },
        paragraph: { spacing: { before: 300, after: 120 } },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 26, bold: true, color: "FFFFFF" },
        paragraph: { spacing: { before: 240, after: 80 } },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 24, bold: true, color: BLUE_DARK },
        paragraph: { spacing: { before: 180, after: 60 } },
      },
    ],
  },
  sections: [{
    properties: {
      page: { margin: { top: 720, bottom: 720, left: 900, right: 900 } },
    },
    children: [

      // ══════════════════════════════════════════════════════
      // TITLE
      // ══════════════════════════════════════════════════════
      new Paragraph({
        children: [new TextRun({ text: "INTERVIEW PREPARATION", bold: true, size: 52, color: "FFFFFF" })],
        alignment: AlignmentType.CENTER,
        shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK },
        spacing: { before: 0, after: 0 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "Beverley Tan", bold: true, size: 40, color: "FFFFFF" })],
        alignment: AlignmentType.CENTER,
        shading: { type: ShadingType.SOLID, color: BLUE_MED, fill: BLUE_MED },
        spacing: { before: 0, after: 0 },
      }),
      spacer(),

      // ══════════════════════════════════════════════════════
      // CANDIDATE INFO
      // ══════════════════════════════════════════════════════
      h1("Candidate Summary"),
      infoTable([
        ["Source",          "ResumeScanner Batch 2026-04-24_141024"],
        ["ResumeScanner Score", "90.0 — Invite Interview"],
        ["Epic Sphinx Test", "Passed (externally confirmed)"],
        ["Education",        "BSocSci (Hons), NUS 2017-2021 — Sociology + SE Asian Studies"],
        ["Total Experience", "~13 years (mixed background)"],
        ["NUHS Tenure",      "Sep 2021 – Dec 2022 (~15 months)"],
        ["Current Role",     "Consultant, Deloitte (Nov 2024 – Present)"],
        ["EPIC Modules",     "Cadence, Prelude, Bridges, Training Environment Build"],
        ["EPIC Experience",  "Training + Configuration/Build/Testing + Go-live Support"],
        ["Clusters Covered", "NUHS, NHG, YH (Yishun Hospital), IMH, WH (Khoo Teck Puat)"],
        ["Score Breakdown",  "Healthcare: 5/5 | EMR/EPIC: 5/5 | Comm: 5/5 | Exp: 4/5 | Analytical: 3/5"],
      ]),
      spacer(),

      // ══════════════════════════════════════════════════════
      // FLAG SIGNALS
      // ══════════════════════════════════════════════════════
      h1("Flag Signals — What to Watch For"),
      h2("🚩 NO-HIRE Red Flags"),
      crossBullet("Blames NUHS / Accenture / Deloitte for project failures"),
      crossBullet("Can't explain what 'configuration, build and testing' means — only did training, not technical"),
      crossBullet("Gap between Dec 2022 and Nov 2024 (~11 months) — unexplained"),
      crossBullet("Vague on Accenture role at YH / IMH / WH go-lives — was she technical or just floating?"),
      crossBullet("Only 2 months at Flying Chalks — pattern of short stints?"),
      crossBullet("Can't speak to Deloitte's current work — potential conflict with NUHS"),
      spacer(),
      h2("✅ HIRE Green Flags"),
      checkBullet("Owns specific EPIC module work — 'I configured Cadence for...'"),
      checkBullet("Describes cross-cluster coordination with named stakeholders"),
      checkBullet("Change management webinar of 350 — shows initiative and communication skill"),
      checkBullet("Explains employment gap clearly with constructive activity"),
      checkBullet("Says 'we' at Accenture but can clarify her specific role"),
      checkBullet("'Dispatched configuration, build and testing' — strong ownership phrasing"),
      spacer(),

      pageBreak(),

      // ══════════════════════════════════════════════════════
      // SECTION 1 — ACCOUNTABILITY
      // ══════════════════════════════════════════════════════
      h1("Section 1: Accountability & Responsibility (STAR Format)"),

      ...qBlock(1,
        "Tell me about a time when a task you were responsible for didn't go as planned. What did you do?",
        [
          { check: true,  text: "Owns it fully — 'I realised I missed X, so I did Y to fix it'" },
          { check: true,  text: "Traces root cause without blaming others" },
          { check: true,  text: "Shows learning — 'Next time I would...'" },
          { cross: true, text: "Deflects — 'The system was slow so I couldn't finish'" },
          { cross: true, text: "'We' throughout with no clarification of their role" },
        ],
        ["When you were training NUHS staff on EPIC, were there sessions where the training wasn't landing? What happened?"]
      ),

      ...qBlock(2,
        "Describe a situation where you had multiple deadlines or competing priorities. How did you manage?",
        [
          { check: true,  text: "Structured approach — triage, communicate early if slipping" },
          { check: true,  text: "Finishes what they start even when hard" },
          { cross: true, text: "Drops things when they get difficult" },
          { cross: true, text: "Says yes to everything then can't deliver" },
        ],
        ["At Accenture you supported YH, IMH and WH go-live command centres simultaneously — how did you prioritise?"]
      ),

      ...qBlock(3,
        "Tell me about a time you had to hold someone (or yourself) accountable for a standard or deadline.",
        [
          { check: true,  text: "Doesn't let small issues become big ones" },
          { check: true,  text: "Has honest conversations early" },
          { check: true,  text: "Follows through even when uncomfortable" },
          { cross: true, text: "Avoids confrontation, lets problems fester" },
          { cross: true, text: "Passes the buck upward without trying first" },
        ],
        []
      ),

      ...qBlock(4,
        "Describe a situation where you realised a project or deliverable was at risk of missing its target. What did you do?",
        [
          { check: true,  text: "Surfaces problems early before they become crises" },
          { check: true,  text: "Proposes solutions, not just reports problems" },
          { check: true,  text: "Keeps stakeholders informed so they can make decisions" },
          { cross: true, text: "Hides problems until they blow up" },
          { cross: true, text: "Only escalates after missing the deadline" },
        ],
        ["During your NUHS tenure, when you were dispatching configuration/build/testing of EPIC — were there moments you had to flag something to management?"]
      ),

      ...qBlock(5,
        "Tell me about the biggest mistake you made in your career and what you did about it.",
        [
          { check: true,  text: "Takes full ownership" },
          { check: true,  text: "No excuses, no blame on circumstances" },
          { check: true,  text: "Concrete action taken to fix it" },
          { check: true,  text: "Clear learning applied since" },
          { cross: true, text: "'It wasn't really a mistake, more like...'" },
          { cross: true, text: "Blaming external factors" },
        ],
        []
      ),

      pageBreak(),

      // ══════════════════════════════════════════════════════
      // SECTION 2 — TECHNICAL
      // ══════════════════════════════════════════════════════
      h1("Section 2: Technical — Epic Healthcare, NGEMR & IT Project Stages"),

      h2("Background: What Candidates Should Know"),
      para([
        normal("NGEMR = Next Generation Electronic Medical Record — Singapore's national healthtech initiative connecting Regional Health Systems (NUHS, SingHealth, NHG), specialty centres, and polyclinics into a unified national EHR. EPIC is the primary EMR platform at the RHS level. Candidates should show awareness that EPIC fits into this broader national context, not just their institution."),
      ]),
      spacer(),

      ...qBlock(6,
        "The NGEMR project is Singapore's national healthtech initiative. Based on your NUHS experience with EPIC, what do you understand about how EPIC fits into the broader NGEMR landscape?",
        [
          { check: true,  text: "Understands NGEMR as national-level, multi-cluster integration" },
          { check: true,  text: "Knows EPIC is the EMR system within NGEMR at RHS level" },
          { check: true,  text: "Mentions interoperability, data exchange between clusters" },
          { check: true,  text: "Understands clinical workflows spanning multiple institutions" },
          { cross: true, text: "'I'm not sure what NGEMR is' — red flag for Medical Informatics role" },
          { cross: true, text: "Only knows EPIC as a local system, no national context" },
        ],
        ["How does patient data flow between NUHS and other clusters under NGEMR?"]
      ),

      ...qBlock(7,
        "You listed EPIC Cadence, Prelude, and Bridges. Walk me through what you specifically did with each module — not just training, but build, configuration, or testing.",
        [
          { check: true,  text: "Can describe specific configuration or build tasks per module" },
          { check: true,  text: "Cadence = scheduling — can describe what she configured (resource types, rules?)" },
          { check: true,  text: "Training Environment Build — built from scratch, not just used" },
          { cross: true, text: "Can only speak about training, not configuration or build" },
          { cross: true, text: "Says she 'used' EPIC but can't describe what she built or configured" },
        ],
        [
          "What resource types and appointment types did you configure in Cadence?",
          "Who did you work with when building the Training Environment — EPIC team or internal?",
          "Besides training, did you participate in any build or configuration sessions with the EPIC team?",
        ]
      ),

      ...qBlock(8,
        "Healthcare IT implementations follow stages: Requirements → Design/Build → Testing (UAT) → Deployment → Support/Optimization. Walk me through your Accenture tenure and map your work to these stages.",
        [
          { check: true,  text: "Maps to 4+ stages clearly with specific examples" },
          { check: true,  text: "Accenture go-live support = Deployment + Support stages" },
          { check: true,  text: "Configuration/build at NUHS = Design/Build stage" },
          { cross: true, text: "Vague — 'I did a bit of everything' without structure" },
          { cross: true, text: "Can only describe one stage" },
          { cross: true, text: "'They put me where they needed me' without showing ownership" },
        ],
        ["What was your specific role during YH's go-live? Were you in the command centre or on the floor?", "Did you do any UAT testing work, or was that handled by the EPIC team?"]
      ),

      ...qBlock(9,
        "What do you think is the biggest challenge in implementing an EMR system like EPIC in a hospital environment? Give an example from your experience.",
        [
          { check: true,  text: "User adoption challenges — clinicians resistant to change" },
          { check: true,  text: "Data migration risks and clinical workflow complexity" },
          { check: true,  text: "Stakeholder management — doctors, nurses, admins all have different needs" },
          { cross: true, text: "Only technical answers, no mention of human/clinical factors" },
          { cross: true, text: "'It's just software' — misses the healthcare context" },
        ],
        ["You ran a 350-person change management webinar. What were the biggest pushbacks you faced from clinicians?", "What was harder — getting doctors on board or nurses?"]
      ),

      ...qBlock(10,
        "You worked across NUHS and NHG clusters, and then Accenture supporting YH, IMH, WH. What was the hardest part about coordinating across multiple institutions?",
        [
          { check: true,  text: "Different hospital cultures, different EHR maturity levels" },
          { check: true,  text: "Different stakeholder priorities (doctors vs nurses vs admin)" },
          { check: true,  text: "Had to tailor approach per institution" },
          { check: true,  text: "Mentions specific teams or people she worked with" },
          { cross: true, text: "'It was fine' — doesn't acknowledge complexity" },
          { cross: true, text: "No specific examples" },
        ],
        ["Which hospital had the most resistance to EPIC? How did you handle it?", "What was different between IMH and WH implementations?"]
      ),

      pageBreak(),

      // ══════════════════════════════════════════════════════
      // SECTION 3 — SCORING
      // ══════════════════════════════════════════════════════
      h1("Section 3: Decision Summary — Scoring Rubric"),
      scoringTable(
        ["Dimension", "Strong (✅ HIRE)", "Moderate (⚠️ CONDITIONAL)", "Weak (🚩 NO-HIRE)"],
        [
          ["Accountability",          "Owns all, no deflection,\nclear learning",         "Owns most, some vagueness",             "Blames, deflects, vague" ],
          ["EPIC Technical Depth",    "Can describe config/build\nper module",             "Training + some config",                "Training only" ],
          ["IT Project Stage Awareness", "Maps to 4+ stages clearly",                    "Maps to 2-3 stages",                    "Can't explain stages" ],
          ["NGEMR Understanding",    "Clear national context",                           "Basic awareness",                       "No understanding" ],
          ["Cross-Cluster Collab",   "Specific examples,\nnamed teams",                "General 'worked with others'",          "No examples" ],
        ]
      ),
      spacer(),

      h2("Decision Guide"),
      decisionBox("HIRE  —  4-5 Strong ✅,  no 🚩 flags", GREEN_LIGHT, GREEN_DARK),
      spacer(),
      decisionBox("CONDITIONAL  —  3 Strong, 1-2 Moderate,  no 🚩 flags in Accountability", YELLOW_LIGHT, "7F6000"),
      spacer(),
      decisionBox("NO-HIRE  —  Any 🚩 flag in Accountability  OR  2+ 🚩 flags total  OR  2+ Weaks", RED_LIGHT, "C00000"),
      spacer(),

      pageBreak(),

      // ══════════════════════════════════════════════════════
      // CANDIDATE-SPECIFIC NOTES
      // ══════════════════════════════════════════════════════
      h1("Notes for Beverley Specifically"),
      h2("Areas to Probe Deeper"),
      bullet("~11 month gap between Dec 2022 and Nov 2024 — what's the story here?"),
      bullet("Flying Chalks only 2 months — is this a pattern of short stints?"),
      bullet("Accenture's exact role at YH/IMH/WH — was she technical or functional?"),
      bullet("EPIC 'Training Environment Build' — did she build from scratch or just maintain?"),
      bullet("Deloitte current role — what is she actually doing? Any NUHS conflict?"),
      spacer(),
      h2("Strengths to Validate"),
      checkBullet("350-person change management webinar — impressive scope, ask for specifics"),
      checkBullet("Cross-cluster NUHS + NHG training — ask for numbers and timeline"),
      checkBullet("'Dispatched configuration, build and testing' — strong phrasing, confirm she did it vs coordinated it"),
      spacer(),

      // footer
      new Paragraph({
        children: [new TextRun({ text: "Generated by Hermes InterviewGenerator | ResumeScanner + Epic Sphinx pipeline", italics: true, color: "808080", size: 16 })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 300 },
      }),

    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = '/mnt/d/Hermes/Project/InterviewGenerator/output/Beverley_Tan_Interview_Prep.docx';
  fs.writeFileSync(outPath, buffer);
  console.log('SUCCESS: ' + outPath);
}).catch(err => {
  console.error('ERROR:', err.message);
  process.exit(1);
});
