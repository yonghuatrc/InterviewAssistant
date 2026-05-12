const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType, PageBreak
} = require('docx');
const fs = require('fs');

const candidate = CANDIDATE_JSON;
const section1Questions = SECTION1_QUESTIONS;
const section2Questions = SECTION2_QUESTIONS;
const section3Questions = SECTION3_QUESTIONS;

// ─── helpers ────────────────────────────────────────────────────────────────
const BLUE_DARK  = "1F4E79";
const BLUE_MED   = "2E75B6";
const BLUE_LIGHT = "DEEAF1";
const GREEN_DARK = "375623";
const GREEN_LIGHT= "E2EFDA";
const RED_LIGHT  = "FCE4D6";
const GREY_LIGHT = "F2F2F2";
const PURPLE_DARK = "7030A0";
const PURPLE_LIGHT= "EAD1DC";

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
function para(runs, spacing) {
  if (spacing === undefined) spacing = {};
  const runArray = typeof runs === 'string' ? [new TextRun(runs)] : runs;
  return new Paragraph({ children: runArray, spacing: { before: 60, after: 60, ...spacing } });
}
function bold(text, color) { return new TextRun({ text, bold: true, ...(color ? { color } : {}) }); }
function normal(text, color, italics) {
  return new TextRun({ text, ...(color ? { color } : {}), ...(italics ? { italics: true } : {}) });
}
function bullet(text, level) {
  if (level === undefined) level = 0;
  return new Paragraph({
    children: [new TextRun({ text, size: 20 })],
    bullet: { level },
    spacing: { before: 40, after: 40 },
  });
}
function spacer() { return new Paragraph({ text: "", spacing: { before: 60, after: 60 } }); }
function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

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
            width: { size: 2592, type: WidthType.DXA },
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: value, size: 20 })],
              shading: { type: ShadingType.SOLID, color: i % 2 === 0 ? GREY_LIGHT : "FFFFFF", fill: i % 2 === 0 ? GREY_LIGHT : "FFFFFF" },
              spacing: { before: 60, after: 60 },
            })],
            width: { size: 6048, type: WidthType.DXA },
          }),
        ],
      })
    ),
  });
}

function qBlock(num, question, whatToListen, probes) {
  const elements = [];
  elements.push(new Paragraph({
    children: [
      new TextRun({ text: "Q" + num + ": ", bold: true, color: BLUE_DARK, size: 22 }),
      new TextRun({ text: question, bold: true, size: 22 }),
    ],
    spacing: { before: 200, after: 100 },
    shading: { type: ShadingType.SOLID, color: BLUE_LIGHT, fill: BLUE_LIGHT },
    border: { left: { style: BorderStyle.THICK, size: 6, color: BLUE_DARK } },
  }));
    if (whatToListen && whatToListen.length > 0) {
    var checks = whatToListen.filter(function(item) { return item && (item.check === true || typeof item.check === 'string'); });
    var crosses = whatToListen.filter(function(item) { return item && (item.cross === true || typeof item.cross === 'string'); });
    if (checks.length > 0) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: "Good signals:", bold: true, color: "375623", size: 20 })],
        spacing: { before: 80, after: 40 },
      }));
      checks.forEach(function(item) {
        var raw = typeof item.check === 'string' ? item.check : (item.text || "");
        var label = raw.replace(/^[^\w\u0080-\uFFFF]+/, '').trim();
        elements.push(new Paragraph({
          children: [new TextRun({ text: "\u2705 " + (label || raw), color: "375623", size: 20 })],
          spacing: { before: 40, after: 40 },
        }));
      });
    }
    if (crosses.length > 0) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: "Red flags:", bold: true, color: "C00000", size: 20 })],
        spacing: { before: 60, after: 40 },
      }));
      crosses.forEach(function(item) {
        var raw = typeof item.cross === 'string' ? item.cross : (item.text || "");
        var label = raw.replace(/^[^\w\u0080-\uFFFF]+/, '').trim();
        elements.push(new Paragraph({
          children: [new TextRun({ text: "\uD83D\uDEA7 " + (label || raw), color: "C00000", size: 20 })],
          spacing: { before: 40, after: 40 },
        }));
      });
    }
  }
  if (probes && probes.length > 0) {
    elements.push(new Paragraph({
      children: [new TextRun({ text: "Probes:", bold: true, color: "404040", italics: true, size: 20 })],
      spacing: { before: 80, after: 40 },
    }));
    probes.forEach(function(p) { elements.push(bullet(p)); });
  }
  elements.push(spacer());
  return elements;
}

function qBlockPersonalized(num, question, whatToListen, probes) {
  const elements = [];
  elements.push(new Paragraph({
    children: [
      new TextRun({ text: "\uD83C\uDFAF Q" + num + ": ", bold: true, color: PURPLE_DARK, size: 22 }),
      new TextRun({ text: question, bold: true, size: 22 }),
    ],
    spacing: { before: 200, after: 100 },
    shading: { type: ShadingType.SOLID, color: PURPLE_LIGHT, fill: PURPLE_LIGHT },
    border: { left: { style: BorderStyle.THICK, size: 6, color: PURPLE_DARK } },
  }));
    if (whatToListen && whatToListen.length > 0) {
    var checks = whatToListen.filter(function(item) { return item && (item.check === true || typeof item.check === 'string'); });
    var crosses = whatToListen.filter(function(item) { return item && (item.cross === true || typeof item.cross === 'string'); });
    if (checks.length > 0) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: "Good signals:", bold: true, color: "375623", size: 20 })],
        spacing: { before: 80, after: 40 },
      }));
      checks.forEach(function(item) {
        var raw = typeof item.check === 'string' ? item.check : (item.text || "");
        var label = raw.replace(/^[^\w\u0080-\uFFFF]+/, '').trim();
        elements.push(new Paragraph({
          children: [new TextRun({ text: "\u2705 " + (label || raw), color: "375623", size: 20 })],
          spacing: { before: 40, after: 40 },
        }));
      });
    }
    if (crosses.length > 0) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: "Red flags:", bold: true, color: "C00000", size: 20 })],
        spacing: { before: 60, after: 40 },
      }));
      crosses.forEach(function(item) {
        var raw = typeof item.cross === 'string' ? item.cross : (item.text || "");
        var label = raw.replace(/^[^\w\u0080-\uFFFF]+/, '').trim();
        elements.push(new Paragraph({
          children: [new TextRun({ text: "\uD83D\uDEA7 " + (label || raw), color: "C00000", size: 20 })],
          spacing: { before: 40, after: 40 },
        }));
      });
    }
  }
  if (probes && probes.length > 0) {
    elements.push(new Paragraph({
      children: [new TextRun({ text: "Probes:", bold: true, color: "404040", italics: true, size: 20 })],
      spacing: { before: 80, after: 40 },
    }));
    probes.forEach(function(p) { elements.push(bullet(p)); });
  }
  elements.push(spacer());
  return elements;
}

// qBlock for 🎯 personalized questions that ALSO carry contextual red/green flags
// (used in Sections 1 & 2 when personalized questions have whatToListen attached)
function qBlockPersonalizedWithFlags(num, question, whatToListen, probes) {
  const elements = [];
  elements.push(new Paragraph({
    children: [
      new TextRun({ text: "\uD83C\uDFAF Q" + num + ": ", bold: true, color: PURPLE_DARK, size: 22 }),
      new TextRun({ text: question, bold: true, size: 22 }),
    ],
    spacing: { before: 200, after: 100 },
    shading: { type: ShadingType.SOLID, color: PURPLE_LIGHT, fill: PURPLE_LIGHT },
    border: { left: { style: BorderStyle.THICK, size: 6, color: PURPLE_DARK } },
  }));
    if (whatToListen && whatToListen.length > 0) {
    var checks = whatToListen.filter(function(item) { return item && (item.check === true || typeof item.check === 'string'); });
    var crosses = whatToListen.filter(function(item) { return item && (item.cross === true || typeof item.cross === 'string'); });
    if (checks.length > 0) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: "Good signals:", bold: true, color: "375623", size: 20 })],
        spacing: { before: 80, after: 40 },
      }));
      checks.forEach(function(item) {
        var raw = typeof item.check === 'string' ? item.check : (item.text || "");
        var label = raw.replace(/^[^\w\u0080-\uFFFF]+/, '').trim();
        elements.push(new Paragraph({
          children: [new TextRun({ text: "\u2705 " + (label || raw), color: "375623", size: 20 })],
          spacing: { before: 40, after: 40 },
        }));
      });
    }
    if (crosses.length > 0) {
      elements.push(new Paragraph({
        children: [new TextRun({ text: "Red flags:", bold: true, color: "C00000", size: 20 })],
        spacing: { before: 60, after: 40 },
      }));
      crosses.forEach(function(item) {
        var raw = typeof item.cross === 'string' ? item.cross : (item.text || "");
        var label = raw.replace(/^[^\w\u0080-\uFFFF]+/, '').trim();
        elements.push(new Paragraph({
          children: [new TextRun({ text: "\uD83D\uDEA7 " + (label || raw), color: "C00000", size: 20 })],
          spacing: { before: 40, after: 40 },
        }));
      });
    }
  }
  if (probes && probes.length > 0) {
    elements.push(new Paragraph({
      children: [new TextRun({ text: "Probes:", bold: true, color: "404040", italics: true, size: 20 })],
      spacing: { before: 80, after: 40 },
    }));
    probes.forEach(function(p) { elements.push(bullet(p)); });
  }
  elements.push(spacer());
  return elements;
}

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
    columns: headers.map(function() { return { width: { size: 2160, type: WidthType.DXA } }; }),
    rows: [
      new TableRow({
        children: headers.map(function(h) {
          return new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 20 })],
              shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK },
              alignment: AlignmentType.CENTER,
              spacing: { before: 60, after: 60 },
            })],
            width: { size: 2160, type: WidthType.DXA },
          });
        }),
      }),
      ...rows.map(function(row, ri) {
        return new TableRow({
          children: row.map(function(cell) {
            return new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: cell, size: 20 })],
                shading: { type: ShadingType.SOLID, color: ri % 2 === 0 ? GREY_LIGHT : "FFFFFF", fill: ri % 2 === 0 ? GREY_LIGHT : "FFFFFF" },
                spacing: { before: 60, after: 60 },
              })],
              width: { size: 2160, type: WidthType.DXA },
            });
          }),
        });
      }),
    ],
  });
}

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
          children: [new TextRun({ text: text, bold: true, color: textColor, size: 22 })],
          shading: { type: ShadingType.SOLID, color: fill, fill: fill },
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
        })],
        width: { size: 8640, type: WidthType.DXA },
      })],
    })],
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { size: 32, bold: true, color: "FFFFFF" },
        paragraph: { spacing: { before: 300, after: 120 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        run: { size: 26, bold: true, color: "FFFFFF" },
        paragraph: { spacing: { before: 240, after: 80 } } },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 720, bottom: 720, left: 900, right: 900 } } },
    children: buildDoc(candidate),
  }],
});

Packer.toBuffer(doc).then(function(buf) {
  fs.writeFileSync(OUTPUT_PATH, buf);
  console.log('DONE:' + OUTPUT_PATH);
}).catch(function(e) { console.error('ERROR:' + e.message); process.exit(1); });

function buildDoc(c) {
  var name = c.name || "Unknown";
  var score = "SCORE_VAL" || "N/A";
  var criteriaData = CRITERIA_DATA || {};
  var batch = BATCH_ID_VAL || "Unknown";
  var ep = c;

  var sysText = ((ep.systems_summary || "") + " " + (ep.healthcare_summary || "")).toLowerCase();
  var hasEpic = sysText.indexOf("epic") !== -1;
  var hasCadence = sysText.indexOf("cadence") !== -1;
  var hasPrelude = sysText.indexOf("prelude") !== -1;
  var hasBridges = sysText.indexOf("bridges") !== -1;
  var hasNGEMR = sysText.indexOf("ngemr") !== -1 || sysText.indexOf("national") !== -1;
  var hasEMR = sysText.indexOf("emr") !== -1 || sysText.indexOf("electronic medical") !== -1;

  var epicModules = [];
  if (hasEpic)    epicModules.push("Epic");
  if (hasCadence) epicModules.push("Cadence");
  if (hasPrelude) epicModules.push("Prelude");
  if (hasBridges) epicModules.push("Bridges");
  var epicStr = epicModules.length ? epicModules.join(", ") : "Not specified from resume";

  var eduText = Array.isArray(ep.education) ? ep.education.join(" | ") : (ep.education || "");
  var eduSummary = eduText.length > 120 ? eduText.substring(0, 120) + "..." : eduText;

  // Build personalized question blocks (Phase 2)
  // Build Section 1 question blocks dynamically
  var section1Blocks = [];
  section1Questions.forEach(function(q, i) {
    var qNum = i + 1;
    if (q.isPersonalized && q.whatToListen && q.whatToListen.length > 0) {
      section1Blocks.push.apply(section1Blocks, qBlockPersonalizedWithFlags(qNum, q.question, q.whatToListen, q.probes));
    } else if (q.isPersonalized) {
      section1Blocks.push.apply(section1Blocks, qBlockPersonalized(qNum, q.question, q.whatToListen, q.probes));
    } else {
      section1Blocks.push.apply(section1Blocks, qBlock(qNum, q.question, q.whatToListen, q.probes));
    }
  });

  // Build Section 2 question blocks dynamically
  var section2Blocks = [];
  section2Questions.forEach(function(q, i) {
    var qNum = i + 1;
    if (q.isPersonalized && q.whatToListen && q.whatToListen.length > 0) {
      section2Blocks.push.apply(section2Blocks, qBlockPersonalizedWithFlags(qNum, q.question, q.whatToListen, q.probes));
    } else if (q.isPersonalized) {
      section2Blocks.push.apply(section2Blocks, qBlockPersonalized(qNum, q.question, q.whatToListen, q.probes));
    } else {
      section2Blocks.push.apply(section2Blocks, qBlock(qNum, q.question, q.whatToListen, q.probes));
    }
  });

  // Build Section 3 question blocks (Motivation & Retention)
  var section3Blocks = [];
  section3Questions.forEach(function(q, i) {
    var qNum = i + 1;
    if (q.isPersonalized && q.whatToListen && q.whatToListen.length > 0) {
      section3Blocks.push.apply(section3Blocks, qBlockPersonalizedWithFlags(qNum, q.question, q.whatToListen, q.probes));
    } else if (q.isPersonalized) {
      section3Blocks.push.apply(section3Blocks, qBlockPersonalized(qNum, q.question, q.whatToListen, q.probes));
    } else {
      section3Blocks.push.apply(section3Blocks, qBlock(qNum, q.question, q.whatToListen, q.probes));
    }
  });

  var result = [
    // TITLE
    new Paragraph({
      children: [new TextRun({ text: "INTERVIEW PREPARATION", bold: true, size: 52, color: "FFFFFF" })],
      alignment: AlignmentType.CENTER,
      shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK },
      spacing: { before: 0, after: 0 },
    }),
    new Paragraph({
      children: [new TextRun({ text: name, bold: true, size: 40, color: "FFFFFF" })],
      alignment: AlignmentType.CENTER,
      shading: { type: ShadingType.SOLID, color: BLUE_MED, fill: BLUE_MED },
      spacing: { before: 0, after: 0 },
    }),
    spacer(),

    // CANDIDATE INFO
    h1("Candidate Summary"),
    infoTable([
      ["Source",              "ResumeScanner Batch " + batch],
      ["Total Score",         criteriaData.total_score ? String(criteriaData.total_score) + " / 100" : String(score) + " — N/A"],
      ["Recommendation",      criteriaData.recommendation || "N/A"],
      ["Epic Sphinx Test",    "Passed (externally confirmed)"],
      ["Education",           eduSummary || "See resume"],
      ["Total Experience",    TOTAL_EXP_YEARS + " years"],
      ["EPIC Modules",         epicStr],
      ["EPIC Exposure",        hasEpic ? "Epic mentioned in resume" : "No Epic keywords found"],
      ["NGEMR Awareness",      hasNGEMR ? "NGEMR/national initiative mentioned" : "Not explicitly mentioned"],
      ["EMR Experience",      hasEMR ? "EMR/EHR directly mentioned" : "General healthcare IT exposure"],
    ]),
    spacer(),

    // ResumeScanner Criteria Breakdown
    criteriaData.criteria && Object.keys(criteriaData.criteria).length > 0 ? (function() {
      var crit = criteriaData.criteria;
      var rows = [];
      // Header row
      rows.push(new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Criteria", bold: true, size: 18, color: "FFFFFF" })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Score", bold: true, size: 18, color: "FFFFFF" })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Weight", bold: true, size: 18, color: "FFFFFF" })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Weighted", bold: true, size: 18, color: "FFFFFF" })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: BLUE_DARK, fill: BLUE_DARK } }),
      ]}));
      var labelMap = {
        working_experience: "Working Experience",
        healthcare_experience: "Healthcare Experience",
        systems_emr_exposure: "Systems / EMR Exposure",
        stakeholder_communication: "Stakeholder Communication",
        analytical_thinking: "Analytical Thinking",
      };
      var alt = false;
      Object.keys(crit).forEach(function(key) {
        var item = crit[key];
        var bg = alt ? "F0F4FA" : "FFFFFF";
        alt = !alt;
        rows.push(new TableRow({ children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: labelMap[key] || key, size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: bg, fill: bg } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(item.score) + " / 5", size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: bg, fill: bg } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(item.weight) + "%", size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: bg, fill: bg } }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(item.weighted.toFixed(1)), size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: bg, fill: bg } }),
        ]}));
      });
      // Total row
      var totalWeighted = Object.keys(crit).reduce(function(s, k) { return s + (crit[k].weighted || 0); }, 0);
      rows.push(new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "TOTAL", bold: true, size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: "E8F0FE", fill: "E8F0FE" } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", bold: true, size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: "E8F0FE", fill: "E8F0FE" } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "100%", bold: true, size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: "E8F0FE", fill: "E8F0FE" } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(totalWeighted.toFixed(1)) + " / 100", bold: true, size: 18 })], spacing: { before: 60, after: 60 } })], shading: { type: ShadingType.SOLID, color: "E8F0FE", fill: "E8F0FE" } }),
      ]}));
      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: rows,
      });
    })() : [],  // Use [] instead of null to avoid null entries in array
    spacer(),

    // FLAG SIGNALS
    h1("Flag Signals — What to Watch For"),
    h2("🚩 NO-HIRE Red Flags"),
    new Paragraph({ children: [new TextRun({ text: "🚩 Blames employer or teammates for project failures", color: "C00000", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "🚩 Cannot describe specific EPIC modules or configuration work", color: "C00000", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "🚩 Employment gaps without clear explanation", color: "C00000", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "🚩 Short job tenures without justification", color: "C00000", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "🚩 Vague answers with no concrete examples — uses 'we' without clarifying personal role", color: "C00000", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "🚩 No awareness of NGEMR or national healthcare IT context", color: "C00000", size: 20 })], spacing: { before: 40, after: 40 } }),
    spacer(),
    h2("✅ HIRE Green Flags"),
    new Paragraph({ children: [new TextRun({ text: "✅ Owns specific work — uses 'I owned...', 'I was responsible for...'", color: "375623", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "✅ Describes specific EPIC modules with configuration or build details", color: "375623", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "✅ Mentions cross-team or cross-institution collaboration with named examples", color: "375623", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "✅ Shows learning from mistakes — what they would do differently", color: "375623", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "✅ Asks clarifying questions — shows they care about getting it right", color: "375623", size: 20 })], spacing: { before: 40, after: 40 } }),
    new Paragraph({ children: [new TextRun({ text: "✅ Clear NGEMR/EPIC national context awareness", color: "375623", size: 20 })], spacing: { before: 40, after: 40 } }),
    spacer(),

    pageBreak(),

    // SECTION 1 — ACCOUNTABILITY
    h1("Section 1: Accountability & Responsibility (STAR Format)"),
    ...section1Blocks,

    // SECTION 2 — TECHNICAL
    pageBreak(),
    h1("Section 2: Technical — Epic Healthcare, NGEMR & IT Project Stages"),

    h2("Background: What Candidates Should Know"),
    para([normal("NGEMR = Next Generation Electronic Medical Record — Singapore's national healthtech initiative connecting Regional Health Systems (NUHS, SingHealth, NHG), specialty centres, and polyclinics into a unified national EHR. EPIC is the primary EMR platform at the RHS level. Candidates should show awareness that EPIC fits into this broader national context, not just their institution.")]),
    spacer(),

    ...section2Blocks,

    // SECTION 3 — MOTIVATION, SELF-AWARENESS & RETENTION
    pageBreak(),
    h1("Section 3: Motivation, Self-Awareness & Retention"),
    h2("Why this role? Why us? Will they stay?"),
    para([normal("These questions probe genuine motivation and self-awareness. Fresh graduates and career changers especially need to show they understand what this role involves and have thought carefully about whether it's the right fit. Watch for generic answers, lack of research, and unrealistic expectations — all strong no-hire signals.")]),
    spacer(),

    ...section3Blocks,

    // SECTION 4 — SCORING
    pageBreak(),
    h1("Section 4: Decision Summary — Scoring Rubric"),
    scoringTable(
      ["Dimension", "Strong (✅ HIRE)", "Moderate (⚠️ CONDITIONAL)", "Weak (🚩 NO-HIRE)"],
      [
        ["Accountability",           "Owns all, no deflection,\nclear learning",               "Owns most, some vagueness",  "Blames, deflects, vague"],
        ["EPIC Technical Depth",     "Can describe config/build\nper module",              "Training + some config",      "Training only / vague"],
        ["IT Project Stage Awareness","Maps to 4+ stages clearly",                       "Maps to 2-3 stages",           "Can't explain stages"],
        ["NGEMR Understanding",      "Clear national context",                            "Basic awareness",             "No understanding"],
        ["Cross-Team Collaboration", "Specific examples,\nnamed teams",                "General 'worked with others'", "No examples"],
        ["Motivation & Retention",  "Clear why this role,\nrealistic expectations,\nshows research", "Some specificity,\nmoderate understanding", "Generic, vague,\nno research shown"],
      ]
    ),
    spacer(),
    h2("Decision Guide"),
    decisionBox("HIRE  —  4-5 Strong ✅,  no 🚩 flags", "E2EFDA", "375623"),
    spacer(),
    decisionBox("CONDITIONAL  —  3 Strong, 1-2 Moderate,  no 🚩 flags in Accountability", "FFF2CC", "7F6000"),
    spacer(),
    decisionBox("NO-HIRE  —  Any 🚩 flag in Accountability  OR  2+ 🚩 flags total  OR  2+ Weaks", "FCE4D6", "C00000"),
    spacer(),

    // FOOTER
    new Paragraph({
      children: [new TextRun({ text: "Generated by Hermes InterviewGenerator | ResumeScanner + Epic Sphinx pipeline", italics: true, color: "808080", size: 16 })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 300 },
    }),
  ];

  return result;
}
