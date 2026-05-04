'use strict';
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat, PageNumber, Footer, Header,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── COLOUR PALETTE ────────────────────────────────────────────────────────
const BLUE_DARK   = "1F3864";   // section headers
const BLUE_MID    = "2E75B6";   // h2
const BLUE_LIGHT  = "BDD7EE";   // table header fill
const GREEN_DARK  = "1A5E20";   // new content header text
const GREEN_FILL  = "E8F5E9";   // new section background
const AMBER_FILL  = "FFF8E1";   // tip boxes
const RED_FILL    = "FFEBEE";   // critical boxes
const GREY_FILL   = "F5F5F5";   // alternate table rows
const TEAL_FILL   = "E0F7FA";   // scenario boxes
const WHITE       = "FFFFFF";
const CONTENT_W   = 9360;       // DXA (US Letter, 1-inch margins)

// ─── BORDER HELPERS ────────────────────────────────────────────────────────
const border = (color="CCCCCC", sz=4) => ({ style: BorderStyle.SINGLE, size: sz, color });
const allBorders = (color="CCCCCC", sz=4) => {
  const b = border(color, sz);
  return { top: b, bottom: b, left: b, right: b };
};
const noBorder = () => {
  const b = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: b, bottom: b, left: b, right: b };
};

// ─── PARAGRAPH HELPERS ─────────────────────────────────────────────────────
const p = (text, opts = {}) => new Paragraph({
  alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
  spacing: { before: opts.spaceBefore ?? 60, after: opts.spaceAfter ?? 60 },
  ...(opts.border ? { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE_MID, space: 1 } } } : {}),
  children: typeof text === 'string'
    ? [new TextRun({ text, font: "Arial", size: opts.size ?? 22, bold: opts.bold ?? false,
        color: opts.color ?? "000000", italics: opts.italic ?? false })]
    : text
});

const blank = () => new Paragraph({ children: [new TextRun({ text: "", size: 20 })], spacing: { before: 0, after: 60 } });

const run = (text, opts = {}) => new TextRun({
  text, font: "Arial",
  size: opts.size ?? 22,
  bold: opts.bold ?? false,
  color: opts.color ?? "000000",
  italics: opts.italic ?? false,
  highlight: opts.highlight ?? undefined
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 120 },
  children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: WHITE })],
  shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
  padding: { top: 120, bottom: 120, left: 180, right: 180 }
});

const h2 = (text, isNew = false) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 100 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: isNew ? "2E7D32" : BLUE_MID, space: 1 } },
  children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: isNew ? GREEN_DARK : BLUE_MID })]
});

const h3 = (text, isNew = false) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 200, after: 80 },
  children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: isNew ? "1B5E20" : BLUE_DARK })]
});

const h4 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_4,
  spacing: { before: 160, after: 60 },
  children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: "374151" })]
});

const bullet = (text, level = 0, isNew = false) => new Paragraph({
  numbering: { reference: "bullets", level },
  spacing: { before: 40, after: 40 },
  children: typeof text === 'string'
    ? [new TextRun({ text, font: "Arial", size: 22, color: "000000" })]
    : text
});

const numbered = (text, level = 0) => new Paragraph({
  numbering: { reference: "numbers", level },
  spacing: { before: 40, after: 40 },
  children: typeof text === 'string'
    ? [new TextRun({ text, font: "Arial", size: 22 })]
    : text
});

// ─── CALLOUT BOXES ─────────────────────────────────────────────────────────
const calloutBox = (icon, label, text, fill, borderColor) => {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders: allBorders(borderColor, 8),
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 180, right: 180 },
          width: { size: CONTENT_W, type: WidthType.DXA },
          children: [
            new Paragraph({ spacing: { before: 0, after: 60 }, children: [
              new TextRun({ text: `${icon} ${label}: `, font: "Arial", size: 22, bold: true, color: borderColor }),
              new TextRun({ text, font: "Arial", size: 22, color: "000000" })
            ]})
          ]
        })
      ]})
    ]
  });
};

const tip  = (t) => calloutBox("💡","TIP",   t, AMBER_FILL, "F57C00");
const crit = (t) => calloutBox("🔴","CRITICAL",t, RED_FILL,  "C62828");
const note = (t) => calloutBox("📘","NOTE",  t, TEAL_FILL, "00838F");
const warn = (t) => calloutBox("⚠️","WARNING",t, AMBER_FILL,"E65100");
const newBadge = (t) => calloutBox("🆕","JOB-READY ADDITION",t, GREEN_FILL,"2E7D32");

// ─── NEW SECTION BANNER ────────────────────────────────────────────────────
const newSectionBanner = (title) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [new TableRow({ children: [
    new TableCell({
      borders: allBorders("2E7D32", 10),
      shading: { fill: "1B5E20", type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      width: { size: CONTENT_W, type: WidthType.DXA },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: "✨ NEW SECTION — JOB-READY CONTENT  |  " + title, font: "Arial", size: 26, bold: true, color: WHITE })
      ]})]
    })
  ]})]
});

// ─── TABLE BUILDER ─────────────────────────────────────────────────────────
const makeTable = (headers, rows, colWidths) => {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  const hRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders: allBorders(BLUE_MID, 6),
      shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      width: { size: colWidths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [
        new TextRun({ text: h, font: "Arial", size: 20, bold: true, color: BLUE_DARK })
      ]})]
    }))
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: allBorders("BBBBBB", 4),
      shading: { fill: ri % 2 === 0 ? WHITE : GREY_FILL, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      width: { size: colWidths[ci], type: WidthType.DXA },
      children: [new Paragraph({ children: [
        new TextRun({ text: cell, font: "Arial", size: 19, color: "111111" })
      ]})]
    }))
  }));
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [hRow, ...dataRows]
  });
};

// ─── SECTION DIVIDER ───────────────────────────────────────────────────────
const sectionDivider = (sectionNum, title, subtitle) => [
  new Paragraph({ children: [new PageBreak()] }),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: noBorder(),
        shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
        margins: { top: 360, bottom: 360, left: 360, right: 360 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [
            new TextRun({ text: sectionNum, font: "Arial", size: 48, bold: true, color: BLUE_LIGHT })
          ]}),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [
            new TextRun({ text: title, font: "Arial", size: 36, bold: true, color: WHITE })
          ]}),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 0 }, children: [
            new TextRun({ text: subtitle, font: "Arial", size: 22, bold: false, color: BLUE_LIGHT, italics: true })
          ]})
        ]
      })
    ]})]
  }),
  blank()
];

// ═══════════════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ═══════════════════════════════════════════════════════════════════════════

const children = [];

// ─── COVER PAGE ────────────────────────────────────────────────────────────
children.push(
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: noBorder(),
        shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
        margins: { top: 720, bottom: 720, left: 360, right: 360 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "IBM STERLING B2B INTEGRATOR", font: "Arial", size: 52, bold: true, color: WHITE })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, children: [new TextRun({ text: "Complete Job-Ready Training Guide", font: "Arial", size: 36, bold: false, color: BLUE_LIGHT, italics: true })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120 }, children: [new TextRun({ text: "Production Environment Mastery  ·  2026 Edition", font: "Arial", size: 26, color: BLUE_LIGHT })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240 }, children: [new TextRun({ text: "✨ ENHANCED & AUGMENTED FOR MAXIMUM JOB READINESS ✨", font: "Arial", size: 22, bold: true, color: "FFD700" })] }),
        ]
      })
    ]})]
  }),
  blank(),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 0 — Day 1 Survival Guide 🆕", { bold: true, color: GREEN_DARK })] }),
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 1 — Core Architecture & Key Concepts")] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 2 — 10 Real-World End-to-End Scenarios")] }),
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 3 — 25 Hands-On Lab Exercises")] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 4 — 25 Expert Interview Questions & Answers")] }),
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 5 — Comprehensive Troubleshooting Guide")] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: WHITE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("SECTION 6 — Alternative Tools: MuleSoft, Boomi, Tibco, SAP, Azure")] }),
        new TableCell({ borders: allBorders(BLUE_MID, 4), shading: { fill: "E8F5E9", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: CONTENT_W/2, type: WidthType.DXA },
          children: [p("BONUS — Requirements, S2T, UAT Checklists & Soft Skills 🆕", { bold: true, color: GREEN_DARK })] }),
      ]}),
    ]
  }),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// HOW TO USE THIS GUIDE (NEW)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(
  new Paragraph({ children: [new PageBreak()] }),
  newSectionBanner("HOW TO USE THIS GUIDE"),
  blank(),
  p([
    run("This guide is your complete roadmap to becoming a ", { size: 22 }),
    run("confident, production-ready EDI Business Analyst", { bold: true, size: 22 }),
    run(" specializing in IBM Sterling B2B Integrator. It has been purpose-built to take you from zero to job-ready — even if you encounter the platform for the first time on Day 1 of a new role.", { size: 22 })
  ], { spaceAfter: 100 }),
  h3("Reading Sequence by Experience Level", true),
  bullet([run("Complete Beginner: ", { bold: true }), run("Read Section 0 (Day 1 Guide) → Section 1 (Architecture) → Do Labs 1–8 → Read Interview Q's 1–10.")]),
  bullet([run("Some EDI Experience: ", { bold: true }), run("Skim Section 0 → Deep-dive Section 1 → Do Labs 9–21 → Study all Interview Q's and Scenarios.")]),
  bullet([run("Experienced EDI BA: ", { bold: true }), run("Focus on Sections 2, 5 (Scenarios & Troubleshooting) → Do advanced Labs 16–25 → Master the Soft Skills section.")]),
  blank(),
  h3("Symbols Used Throughout This Guide", true),
  makeTable(
    ["Symbol", "Meaning"],
    [
      ["💡 TIP", "Practical insight that saves time in production"],
      ["🔴 CRITICAL", "Must-know fact — getting this wrong causes outages"],
      ["📘 NOTE", "Important context or background knowledge"],
      ["⚠️ WARNING", "Common mistake — many BAs get this wrong"],
      ["🆕 NEW SECTION", "Content added to maximize your job readiness"],
      ["🎯 SCENARIO", "Real-world situation you will face as a BA"],
      ["✅ CHECKLIST", "Action item — mark complete when done"],
    ],
    [3000, 6360]
  ),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 0: DAY 1 SURVIVAL GUIDE (NEW)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 0 🆕",
  "EDI BA Day 1 Survival Guide",
  "Everything you need to know before your first day — and how to ace the first 30 days"
));

children.push(
  newBadge("This section is a comprehensive addition designed to give you maximum confidence and effectiveness from your very first day as an EDI Business Analyst."),
  blank(),

  h2("0.1  What is the EDI Business Analyst Role — Really?", true),
  p("Many job descriptions for 'EDI Business Analyst' use the same words, but the day-to-day reality can vary significantly by company. Understanding the spectrum of responsibilities before Day 1 will help you ask the right questions, set expectations, and prioritize your learning."),
  blank(),
  makeTable(
    ["Responsibility Area", "What You Actually Do", "Frequency"],
    [
      ["Trading Partner Onboarding", "Configure new partners end-to-end: profiles, maps, BPs, testing, go-live", "Weekly–Monthly"],
      ["Map Development", "Build and maintain EDI translation maps in Sterling Map Editor", "Daily"],
      ["Requirements Gathering", "Interview stakeholders; write S2T mapping specifications", "Per project"],
      ["Incident Response", "Diagnose and fix failed BPs, map errors, connectivity failures", "Daily"],
      ["Stakeholder Communication", "Translate technical EDI errors into business language for non-technical teams", "Daily"],
      ["Testing & UAT", "Design test cases, execute UAT, obtain sign-off, promote to production", "Per release"],
      ["Documentation", "Maintain S2T specs, runbooks, trading partner profiles, onboarding guides", "Ongoing"],
      ["Compliance & Audit", "Ensure EDI transactions meet HIPAA, SOX, or retail compliance requirements", "Quarterly"],
      ["Platform Administration", "Deploy maps/BPs, manage certificates, configure adapters", "As needed"],
      ["Reporting & Metrics", "Track EDI health metrics; report to management", "Weekly"],
    ],
    [2800, 4500, 2060]
  ),
  blank(),

  h2("0.2  Day 1 Priorities: The First 4 Hours", true),
  h3("Hour 1: Environment Access", true),
  p("Before you can do anything, you need access. On Day 1, chase these down immediately — don't wait for someone to hand them to you:"),
  bullet([run("Sterling Dashboard URL: ", { bold: true }), run("Typically https://sterling-server:9080/dashboard (or 443 for HTTPS). Your manager or DevOps team will provide this.")]),
  bullet([run("Credentials: ", { bold: true }), run("Your Sterling user ID and initial password. Request admin-level access if you are the primary EDI BA — you need it for deployment.")]),
  bullet([run("VPN access: ", { bold: true }), run("Most Sterling environments are on-premise. Confirm VPN is set up and you can reach the Sterling URL.")]),
  bullet([run("SSH/server access: ", { bold: true }), run("You will need SSH access to the Sterling server for log diagnostics (tail -f commands). Request this on Day 1.")]),
  bullet([run("DEV, TEST, and PROD URLs: ", { bold: true }), run("Sterling environments are typically separate. Know which is which before you touch anything.")]),
  bullet([run("JIRA/ticketing system access: ", { bold: true }), run("All EDI changes must be tracked. Get access to your team's change management system.")]),
  blank(),

  h3("Hour 2: The Lay of the Land", true),
  p("Once you have access, spend one hour getting oriented before touching anything:"),
  bullet([run("Operations > Business Processes: ", { bold: true }), run("How many active BPs are there? What are the most common names? This tells you what transaction types are live.")]),
  bullet([run("Operations > Business Processes > Status=Halted: ", { bold: true }), run("Are there any currently failing? Do NOT fix them yet — just observe and note.")]),
  bullet([run("Administration > Trading Partners: ", { bold: true }), run("How many trading partners exist? Which are the most important (highest volume)?")]),
  bullet([run("Deployment > Maps: ", { bold: true }), run("List all deployed maps. This is your portfolio — what transaction types does the company exchange?")]),
  bullet([run("Ask your manager: ", { bold: true }), run("'What are the top 3 trading partners by volume/revenue? What are the current open EDI issues?' This gives you immediate context.")]),
  blank(),
  tip("Take notes in a personal wiki or Confluence page from Hour 1. 'EDI Environment Overview — [Company Name]' becomes your most-used reference document for the next 6 months."),
  blank(),

  h3("Hour 3: Documentation Inventory", true),
  p("Great EDI BAs inherit documentation and actively build on it. Bad EDI BAs let tribal knowledge stay in people's heads. Locate:"),
  bullet([run("S2T (Source-to-Target) Mapping Documents: ", { bold: true }), run("These define exactly how fields map between EDI and internal systems. They are your blueprint.")]),
  bullet([run("Companion Guides: ", { bold: true }), run("Each major trading partner has one. These override the X12 standard. They define which optional segments are required for that partner.")]),
  bullet([run("Onboarding Runbook: ", { bold: true }), run("Step-by-step guide for adding a new trading partner. If one doesn't exist, creating it is your first major deliverable.")]),
  bullet([run("Incident Log: ", { bold: true }), run("Past incidents tell you what breaks most often. Ask for the last 6 months of EDI incident tickets.")]),
  bullet([run("Architecture Diagram: ", { bold: true }), run("A diagram showing how Sterling connects to ERPs, WMS, TMS, and partner networks. If none exists, draw one by end of Week 1.")]),
  blank(),

  h3("Hour 4: Meet the Stakeholders", true),
  p("EDI touches every part of the business. Schedule introductory calls in your first week with:"),
  makeTable(
    ["Stakeholder", "Their Relationship to EDI", "Key Question to Ask Them"],
    [
      ["IT/DevOps Manager", "Owns the Sterling server infrastructure; manages deployments", "'What is the change management process for EDI deployments?'"],
      ["ERP Team (SAP/Oracle/JDE)", "Receives EDI output (PO XMLs, ASN data); sends invoice data to EDI", "'What format does the ERP expect from Sterling, and what does it send back?'"],
      ["Logistics/Supply Chain", "Business owner of ASN (856), PO (850), and inventory transactions", "'What are your biggest pain points with EDI today?'"],
      ["Finance/Accounts Receivable", "Consumes 820 remittance data; sends 810 invoices via EDI", "'How do you currently reconcile EDI payments? What is manual vs automated?'"],
      ["Trading Partner EDI Contacts", "External partners who send/receive EDI", "'What is your preferred contact method for EDI issues, and what is your SLA?'"],
      ["QA/Testing Team", "May own UAT process for EDI changes", "'How do you currently test EDI changes before production?'"],
      ["Compliance/Audit", "Ensures HIPAA/SOX/retail compliance on EDI data", "'Are there any upcoming compliance audits that touch EDI?'"],
    ],
    [2400, 3400, 3560]
  ),
  blank(),

  h2("0.3  First 30 Days: Your Ramp-Up Plan", true),
  makeTable(
    ["Timeline", "Focus Area", "Deliverable"],
    [
      ["Days 1–3", "Environment orientation; access setup; stakeholder meetings", "Personal 'EDI Environment Overview' wiki page"],
      ["Days 4–7", "Read top 3 companion guides; trace 1 live transaction end-to-end in Operations", "Hand-drawn architecture diagram (whiteboard photo)"],
      ["Week 2", "Shadow on one active incident or onboarding task; read existing S2T docs", "Gap list: 'What documentation is missing?'"],
      ["Week 3", "Build first map in DEV (even a simple one); run Labs 1–4 from this guide", "First compiled .map file; personal lab notes"],
      ["Week 4", "Produce or update one S2T mapping document; attend a partner onboarding call", "Updated or new S2T document; onboarding call notes"],
      ["Day 30", "Demonstrate you can diagnose a Halted BP from scratch in under 15 minutes", "Personal competency self-assessment (use Section 5 checklist)"],
    ],
    [1440, 4000, 3920]
  ),
  blank(),

  h2("0.4  EDI Standards Mastery — Your Technical Foundation", true),
  p("Sterling supports multiple EDI standards. A job-ready BA can navigate all of them. Here is the reference you need on Day 1."),
  blank(),

  h3("The X12 Standard (ANSI ASC X12) — The US Backbone", true),
  p([
    run("X12 is the dominant EDI standard in ", { size: 22 }),
    run("North American retail, healthcare, logistics, and finance", { bold: true, size: 22 }),
    run(". Every major US retailer (Walmart, Target, Kroger, Amazon, Home Depot) uses X12.", { size: 22 })
  ]),
  blank(),
  makeTable(
    ["X12 Concept", "What It Means", "BA Implication"],
    [
      ["Transaction Set", "A specific document type (850=PO, 810=Invoice, 856=ASN)", "You build one map per transaction set per partner"],
      ["Functional Group", "A batch container for same-type transactions within one interchange", "GS/GE envelope; GS01 identifies the transaction type"],
      ["Interchange Envelope", "The outermost wrapper for one transmission (ISA/IEA)", "ISA15=P is Production; ISA15=T is Test — always verify before go-live"],
      ["Companion Guide", "Partner-specific rules that override the X12 standard", "READ THIS FIRST for every new partner. Non-compliance = rejected transactions"],
      ["997 / 999", "Functional Acknowledgment — partner's EDI validation response", "AK5=A: Accepted. AK5=R: Rejected. AK5=E: Accepted with errors"],
      ["Version", "004010, 005010 are most common (5010 for HIPAA mandatory)", "Your map must be compiled against the correct version"],
    ],
    [2200, 3600, 3560]
  ),
  blank(),

  h3("The EDIFACT Standard — The International Standard", true),
  p([
    run("EDIFACT (Electronic Data Interchange For Administration, Commerce and Transport) is the ", { size: 22 }),
    run("dominant global standard outside North America", { bold: true, size: 22 }),
    run(" — used in Europe, Asia-Pacific, and for global supply chains. If your company works with international partners, you will encounter EDIFACT.", { size: 22 })
  ]),
  blank(),
  makeTable(
    ["EDIFACT Structure", "X12 Equivalent", "Key Difference"],
    [
      ["UNB / UNZ", "ISA / IEA", "Interchange envelope. UNB contains sender/receiver ID in different format"],
      ["UNG / UNE", "GS / GE", "Functional group (optional in EDIFACT — often omitted)"],
      ["UNH / UNT", "ST / SE", "Message envelope. UNH02 identifies message type (ORDERS, INVOIC, etc.)"],
      ["ORDERS", "850", "Purchase Order"],
      ["INVOIC", "810", "Invoice"],
      ["DESADV", "856", "Despatch Advice (ASN equivalent)"],
      ["ORDRSP", "855", "Purchase Order Response/Acknowledgment"],
      ["REMADV", "820", "Remittance Advice"],
      ["CONTRL", "997", "Control/Acknowledgment message"],
      ["PRICAT", "832", "Price/Sales Catalogue"],
    ],
    [2000, 2200, 5160]
  ),
  blank(),
  note("Sterling Map Editor supports EDIFACT natively. When building EDIFACT maps, select EDIFACT as the input/output standard and choose the correct version (D96A, D01B, D07A are common). The mapping logic is identical to X12 — only the segment names and hierarchy differ."),
  blank(),

  h3("The HL7 Standard — Healthcare EDI", true),
  p([
    run("HL7 (Health Level 7) is the ", { size: 22 }),
    run("dominant standard in US healthcare EDI", { bold: true, size: 22 }),
    run(". If you work in healthcare IT, insurance, or pharma, you will encounter HL7 alongside X12 HIPAA transactions.", { size: 22 })
  ]),
  blank(),
  makeTable(
    ["HL7 Concept", "Description", "BA Implication"],
    [
      ["HL7 v2.x", "Pipe-delimited message format (PID, ORC, OBX segments). Legacy but still dominant in hospital systems", "Most hospital EHR/EMR systems (Epic, Cerner) speak HL7 v2.x. You interface with them"],
      ["HL7 FHIR", "Modern REST-based standard using JSON/XML resources. Growing fast in new implementations", "FHIR requires API integration skills alongside EDI knowledge — a valuable differentiator"],
      ["X12 HIPAA Transactions", "EDI X12 transactions mandated by HIPAA for healthcare claims and eligibility", "837P/I (claims), 835 (ERA), 270/271 (eligibility), 278 (auth) are the big five"],
      ["Clearinghouse", "Intermediary that validates and routes HIPAA transactions between providers and payers", "Most healthcare EDI passes through a clearinghouse (Availity, Change Healthcare, Waystar)"],
      ["NPI", "National Provider Identifier — required in all HIPAA claims (NM109 element)", "Mandatory in 837 claims. Missing NPI = immediate claim rejection"],
      ["ANSI 005010X222A2", "HIPAA 5010 implementation guide for 837P claims", "This is the companion guide for healthcare. HIPAA compliance is non-negotiable"],
    ],
    [2400, 3800, 3160]
  ),
  blank(),

  h3("VAN (Value Added Networks) — The EDI Backbone Before the Internet", true),
  p([
    run("A ", { size: 22 }),
    run("VAN (Value Added Network)", { bold: true, size: 22 }),
    run(" is a private, managed network that acts as an intermediary between trading partners for exchanging EDI documents. Think of it as an EDI mailbox service in the cloud.", { size: 22 })
  ]),
  blank(),
  makeTable(
    ["VAN Concept", "What It Means"],
    [
      ["Mailbox", "Each trading partner has a VAN mailbox. You deposit EDI documents; they pick them up. Like email but for EDI"],
      ["Pickup/Delivery", "Partners schedule polling intervals to collect from their VAN mailbox (e.g., every 15 min, hourly)"],
      ["VAN-to-VAN", "Partners on different VANs can exchange documents — the VANs route between each other (like email between mail servers)"],
      ["Non-repudiation", "VANs log every document sent and received — provides audit trail for dispute resolution"],
      ["Major Providers", "IBM Sterling B2B Services (formerly GXS), OpenText (formerly GXS), SPS Commerce, DiCentral, TrueCommerce"],
      ["Cost Model", "VANs typically charge per kilocharacter (per 1000 characters of EDI data) — costs add up at high volume"],
      ["vs Direct AS2", "VANs are asynchronous (like email). AS2 is synchronous (real-time). VANs are simpler to set up; AS2 has lower per-transaction cost at volume"],
    ],
    [2600, 6760]
  ),
  blank(),
  tip("In Sterling, VAN connectivity is handled through a VAN adapter or by configuring SFTP to connect to the VAN's FTP server. When a partner says 'we use a VAN', ask: which VAN, what is your mailbox ID, and what pickup schedule do you use. This defines your configuration."),
  blank(),

  h2("0.5  API vs EDI — The Modern BA's Essential Comparison", true),
  p("As an EDI BA in 2026, you will increasingly be asked: 'Should we use EDI or an API for this integration?' Being able to answer this confidently sets you apart from single-dimensional EDI specialists."),
  blank(),
  makeTable(
    ["Dimension", "Traditional EDI (X12/EDIFACT)", "Modern API (REST/JSON)"],
    [
      ["Communication Model", "Asynchronous batch (hourly, daily files)", "Synchronous real-time (milliseconds)"],
      ["Data Format", "Structured fixed-format (X12, EDIFACT segments)", "Flexible JSON, XML, GraphQL"],
      ["Partner Setup", "Weeks — requires companion guide, testing, certification", "Days — shared API spec (OpenAPI/Swagger), API key"],
      ["Error Handling", "997 acknowledgments; delayed error feedback", "Immediate HTTP status codes (200, 400, 500)"],
      ["Industry Support", "Mandated by major retailers, healthcare, government", "Adopted by modern SaaS, marketplaces (Shopify, Amazon SP-API)"],
      ["Volume", "Excellent for high-volume batch (millions of transactions)", "Excellent for low-latency, event-driven integration"],
      ["Cost", "VAN per-kilocharacter or direct AS2 infrastructure cost", "API call pricing or flat subscription"],
      ["Auditability", "Strong — ISA control numbers, 997 trail, VAN logs", "Requires explicit logging/monitoring setup"],
      ["Legacy Partner Support", "Excellent — EDI is decades-established", "Requires partner to have API capability"],
      ["Best For", "Retail, healthcare, logistics, financial services supply chain", "Modern SaaS integrations, real-time inventory, e-commerce"],
    ],
    [2400, 3500, 3460]
  ),
  blank(),
  note("In interviews, a strong answer is: 'EDI and APIs are complementary. EDI remains the standard for regulated industries and high-volume batch exchange. APIs are preferred for real-time, event-driven scenarios and modern SaaS integration. Many enterprises run both — EDI with Walmart, APIs with Shopify. My role as BA is to recommend the right tool based on partner requirements, volume, latency needs, and compliance obligations.'"),
  blank(),

  h2("0.6  Day 1 Environment Checklist", true),
  newBadge("Print this checklist and physically mark each item complete. A BA who completes this in Week 1 is operationally ready from Day 8 onwards."),
  blank(),
  makeTable(
    ["✅", "Checklist Item", "Done By"],
    [
      ["☐", "Sterling Dashboard URL confirmed and bookmarked (DEV, TEST, PROD)", "Day 1"],
      ["☐", "Sterling user credentials obtained; password changed; admin access confirmed", "Day 1"],
      ["☐", "VPN configured; Sterling accessible remotely from personal laptop", "Day 1"],
      ["☐", "SSH access to Sterling server(s) confirmed; can tail log files", "Day 2"],
      ["☐", "Map Editor installed on local machine; connects to Sterling DEV", "Day 2"],
      ["☐", "JIRA/ticketing system access obtained; team project visible", "Day 1"],
      ["☐", "Confluence/SharePoint access obtained; found EDI documentation folder", "Day 2"],
      ["☐", "All existing S2T mapping documents reviewed and saved locally", "Day 3"],
      ["☐", "Top 3 partner companion guides read and annotated", "Week 1"],
      ["☐", "Architecture diagram created/located: Sterling ↔ ERP ↔ Partners", "Week 1"],
      ["☐", "Stakeholder introductory meetings scheduled (IT, ERP, Finance, Logistics)", "Week 1"],
      ["☐", "Emergency contact list created: partner EDI contacts + internal on-call", "Week 1"],
      ["☐", "Change management process understood; first JIRA ticket created", "Week 1"],
      ["☐", "Can navigate to a Halted BP and read its Status Report in < 5 minutes", "Week 2"],
      ["☐", "Built first test map in DEV (Lab 1 from Section 3)", "Week 2"],
    ],
    [480, 6840, 2040]
  ),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 1: CORE ARCHITECTURE (ORIGINAL + AUGMENTED)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 1",
  "Core Architecture & Key Concepts",
  "A deep dive into Business Processes, Adapters, Maps, and Document Flow"
));

children.push(
  h2("1.1  The Sterling Engine — How It All Connects"),
  p("IBM Sterling B2B Integrator (SBI) is not simply an EDI tool — it is a full B2B integration platform. Understanding how its internal components interact is the foundation for becoming effective in any production environment. Every file that enters or leaves Sterling travels through a deterministic pipeline:"),
  blank(),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: allBorders(BLUE_MID, 6),
        shading: { fill: "EEF4FB", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 240, right: 240 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          p("INBOUND FILE  →  [ADAPTER]  →  [MAILBOX / BP TRIGGER]  →  [BUSINESS PROCESS]", { bold: false, color: "1A237E", size: 20 }),
          p("             →  [DE-ENVELOPE]  →  [TRANSLATION MAP]  →  [ERP / SFTP / API]", { bold: false, color: "1A237E", size: 20 }),
          p("             →  [GENERATE 997/MDN]  →  [OUTBOUND ADAPTER]  →  PARTNER", { bold: false, color: "1A237E", size: 20 }),
        ]
      })
    ]})]
  }),
  blank(),
  p("Each arrow above represents a discrete Sterling service. Understanding that services are reusable, configurable units — and that business processes simply orchestrate them — is the single most important mental model you need."),
  blank(),
  note("The pipeline above is not just theoretical. On Day 1, if an incident occurs and a partner says 'we received nothing', you use this pipeline as your diagnostic checklist: Did the Adapter receive it? Did the Mailbox trigger fire? Did the BP start? Did De-enveloping succeed? Did Translation run? Did the output adapter deliver? This systematic approach prevents panic and ensures you find the root cause every time."),
  blank(),

  h2("1.2  Business Processes (BPs) — The Orchestration Layer"),
  p("A Business Process (BP) is an XML document written in BPML (Business Process Markup Language) that describes a sequence of service calls, conditional branches, loops, and fault handlers. In practice, BAs design BPs graphically using the Graphical Process Modeler (GPM), which generates the BPML XML underneath."),
  blank(),
  h3("BP Lifecycle States"),
  makeTable(
    ["State", "Meaning", "BA Action"],
    [
      ["Waiting", "BP is queued; resources not yet allocated", "Normal — no action"],
      ["Running", "BP is actively executing a step", "Monitor — note how long"],
      ["Waiting on I/O", "Adapter waiting for remote response (AS2 MDN, SFTP ack)", "Normal — check timeout config if stuck"],
      ["Interrupted", "BP paused waiting for manual intervention", "Check why; restart or terminate"],
      ["Halted", "BP stopped due to error; human review required", "Read Status Report; diagnose; restart"],
      ["Terminated", "BP was manually stopped or timed out", "Investigate root cause"],
      ["Completed", "BP ran successfully end-to-end", "Verify output; confirm delivery"],
    ],
    [2000, 4000, 3360]
  ),
  blank(),
  h3("BPML Structure — Anatomy"),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: allBorders("333333", 4),
        shading: { fill: "1E1E1E", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          p('<process name="EDI_850_Inbound" version="1.0">', { color: "9CDCFE", size: 20 }),
          p('  <sequence name="MainFlow">', { color: "9CDCFE", size: 20 }),
          p('    <operation name="DeEnvelope">', { color: "9CDCFE", size: 20 }),
          p('      <participant name="EDIDeenvelope"/>', { color: "CE9178", size: 20 }),
          p('      <output message="in">', { color: "9CDCFE", size: 20 }),
          p('        <assign to="EDIStandard">X12</assign>', { color: "4EC9B0", size: 20 }),
          p('      </output>', { color: "9CDCFE", size: 20 }),
          p('    </operation>', { color: "9CDCFE", size: 20 }),
          p('    <onFault>', { color: "C586C0", size: 20 }),
          p('      <operation name="AlertOnFailure">', { color: "9CDCFE", size: 20 }),
          p('        <participant name="SMTPSendAdapter"/>', { color: "CE9178", size: 20 }),
          p('      </operation>', { color: "9CDCFE", size: 20 }),
          p('    </onFault>', { color: "C586C0", size: 20 }),
          p('  </sequence>', { color: "9CDCFE", size: 20 }),
          p('</process>', { color: "9CDCFE", size: 20 }),
        ]
      })
    ]})]
  }),
  blank(),
  tip("GPM generates BPML for you. But always read the raw BPML in Operations > Business Processes when debugging — the XML shows exactly which step failed and what the error was."),
  blank(),
  h3("Key BP Design Patterns"),
  makeTable(
    ["Pattern", "When to Use", "Implementation"],
    [
      ["Sequential", "Default — steps run one after another", "<sequence> with operations in order"],
      ["Parallel", "Multiple independent tasks (e.g., send file AND send notification)", "<choice> or multiple sequences with fork"],
      ["onFault", "Error handling — trigger alert or fallback on any step failure", "<onFault> inside <sequence>"],
      ["Sub-Process", "Reuse a common BP inside another (e.g., reusable error-alerter)", "<operation> calling InvokeBusinessProcessService"],
      ["Correlation", "Link inbound 850 to its 997 response using ISA control number", "Correlation Set using ISA13 as key"],
      ["Wait/Timer", "Hold a document for a time window or until a condition is met", "WaitService with timeout configuration"],
      ["Loop", "Iterate over a set of documents or trading partners", "BPML <repeat> with counter condition"],
    ],
    [2000, 3500, 3860]
  ),
  blank(),

  h2("1.3  Adapters — The Connection Layer"),
  p("Adapters are pluggable connectors that handle the physical transport of data to and from external systems. Each adapter is a registered Sterling service configured with host, authentication, and protocol parameters. A BA must understand how to configure, test, and troubleshoot each adapter type."),
  blank(),
  h3("AS2 Adapter — Deep Dive"),
  p("AS2 (Applicability Statement 2) is an HTTP-based protocol with non-repudiation through digital signatures and MDN receipts. It is mandatory for most major retailers."),
  makeTable(
    ["AS2 Component", "Purpose", "Where Configured"],
    [
      ["AS2 Profile", "Stores partner AS2 ID, certificate, MDN settings", "Admin > Trading Partners > AS2"],
      ["Certificate Store", "Holds your keypair and partner public certs", "Admin > Certificates"],
      ["Inbound Port", "Your Sterling listens for partner's AS2 POST", "Perimeter Server config"],
      ["MDN (Sync)", "Partner returns MDN on same HTTP connection immediately", "AS2 Profile > MDN Type = Synchronous"],
      ["MDN (Async)", "Partner POSTs MDN back to a URL you provide later", "AS2 Profile > MDN Type = Asynchronous"],
      ["AS2 Server Adapter", "Receives inbound AS2 messages from partners", "Adapter config > AS2ServerAdapter"],
      ["AS2 Client Adapter", "Sends outbound AS2 messages to partner URL", "Adapter config > AS2ClientAdapter"],
    ],
    [2400, 3700, 3260]
  ),
  blank(),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: allBorders(BLUE_MID, 6),
        shading: { fill: "EEF4FB", type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          p("AS2 Message Flow:", { bold: true, size: 20, color: BLUE_DARK }),
          p("You → [HTTPS POST + Signed + Encrypted payload] → Partner AS2 URL", { size: 20, color: "1A237E" }),
          p("Partner → [MDN: signed receipt confirming receipt+decryption] → Your MDN URL", { size: 20, color: "1A237E" }),
          p("Partner → [997 X12 acknowledgment: EDI content validation] → Your inbound", { size: 20, color: "1A237E" }),
        ]
      })
    ]})]
  }),
  blank(),
  crit("You need BOTH the MDN (transport confirmation) AND the 997 (EDI content validation). The MDN tells you the file was received; the 997 tells you it was understood. A successful MDN with a rejected 997 means your EDI content is wrong."),
  blank(),
  h3("SFTP Adapter — Configuration Reference"),
  makeTable(
    ["Parameter", "Description", "Common Value"],
    [
      ["Remote Host", "Partner's SFTP server hostname or IP", "partner-sftp.company.com"],
      ["Remote Port", "SFTP port (default 22)", "22"],
      ["Auth Method", "Password or SSH Key", "SSH Key (preferred)"],
      ["SSH Key Alias", "Reference to loaded key in Sterling key store", "partner_rsa_key"],
      ["Remote Directory", "Path on remote server to PUT/GET files", "/inbound/edi/ or /from-you/"],
      ["Local Directory", "Staging area on Sterling server", "/opt/sterling/outbound/acme/"],
      ["Known Hosts", "Validates remote server fingerprint — prevents MITM", "Load partner's fingerprint"],
      ["Connection Timeout", "Seconds before giving up on connect", "60"],
      ["Data Timeout", "Seconds to wait for data transfer to complete", "300"],
      ["Delete After Get", "Remove file from remote after successful GET", "Yes (prevents reprocessing)"],
      ["Filename Pattern", "Filter which files to GET", "*.edi or *.txt"],
    ],
    [2400, 4500, 2460]
  ),
  blank(),
  h3("HTTP/HTTPS Adapter"),
  p("Used for modern REST or SOAP integrations. Key parameters: URL, Method (GET/POST/PUT), Content-Type header, Auth (Basic, OAuth2, API Key), SSL certificate trust store. The BA's role is to configure these in the adapter and write BPML to construct the request body and parse the response."),
  blank(),
  h3("File System Adapter"),
  p("Polls a local directory for files. Key parameters: Directory path, filename filter (*.edi), poll interval, action (Collect or Deliver), archive directory (where to move files after processing). Always configure an archive directory — it gives you a recovery path if processing fails."),
  blank(),
  h3("JMS / MQ Adapter"),
  p("Connects Sterling to IBM MQ or JMS message queues — typically used for ERP integrations (SAP, Oracle). Key parameters: Queue Manager name, Queue name, Channel, Host, Port, Connection credentials. Documents are put/got from the queue as messages. The BA configures adapter and ensures the BP correctly maps MQ message properties to Sterling process data."),
  blank(),
  h3("Adapter Comparison Matrix"),
  makeTable(
    ["Adapter", "Security", "Real-time?", "Retry?", "Non-Repudiation", "Typical Use"],
    [
      ["AS2", "HTTPS + Certs + MDN", "Yes", "Manual requeue", "Yes — MDN", "Retail partners, pharma"],
      ["SFTP", "SSH keys", "Near-real-time", "Scheduler retry", "No", "Finance, logistics, healthcare"],
      ["FTP/FTPS", "TLS (FTPS only)", "Near-real-time", "Scheduler retry", "No", "Legacy partners"],
      ["HTTP/REST", "TLS + OAuth/APIkey", "Yes", "BP retry logic", "No", "Modern SaaS APIs"],
      ["File System", "OS filesystem perms", "Polling interval", "BP retry", "No", "Internal ERP handoffs"],
      ["MQ/JMS", "TLS + auth", "Yes", "MQ persistence", "No", "ERP message queues (SAP)"],
      ["VAN", "VAN-managed", "Polling", "VAN retry", "VAN-managed", "Partners requiring mailbox"],
    ],
    [1500, 1900, 1300, 1500, 1800, 1360]
  ),
  blank(),

  h2("1.4  Maps — The Transformation Layer"),
  p("Maps transform documents from one format to another. Sterling's Map Editor is a standalone Java tool that produces compiled .map files deployed to Sterling. The BA is primarily responsible for designing, building, testing, and deploying maps."),
  blank(),
  h3("Map Types Supported"),
  makeTable(
    ["Input Format", "Output Format", "Use Case"],
    [
      ["X12 EDI", "XML", "Inbound 850 PO to ERP XML format"],
      ["XML", "X12 EDI", "ERP invoice XML to outbound 810"],
      ["Flat File (fixed-width)", "X12 EDI", "Legacy mainframe data to EDI 837 claim"],
      ["X12 EDI", "Flat File", "EDI 835 remittance to legacy payment system"],
      ["XML", "XML", "XSLT-style transformation (ERP schema to partner schema)"],
      ["EDIFACT", "X12", "Standard conversion (EU partner to US buyer)"],
      ["X12", "JSON", "EDI to REST API payload"],
      ["HL7", "X12", "Healthcare interoperability (hospital to payer)"],
    ],
    [2000, 2200, 5160]
  ),
  blank(),
  h3("Map Editor — Key Features Reference"),
  makeTable(
    ["Feature", "Location in Map Editor", "Purpose"],
    [
      ["Direct Map Line", "Drag from input to output field", "Pass-through with no transformation"],
      ["Standard Rules", "Element Properties > Standard Rule tab", "Built-in rules: constant, code list, accumulator, xref"],
      ["Extended Rules", "Right-click segment/element > Extended Rule", "Custom scripting: conditions, loops, calculations"],
      ["Loop Properties", "Right-click loop node > Properties", "Set Min/Max occurrences; critical for performance"],
      ["Conditional Map", "Extended Rule with if/else on rule group", "Map only when certain condition is true"],
      ["Test Map", "Tools > Test Map (green play button)", "Run map against a test file without deploying"],
      ["Compile", "File > Compile Map (Ctrl+F5)", "Produces .map binary for deployment"],
      ["Cross Reference", "Admin > Code Lists > Cross Ref in Sterling", "Key-value lookup table called by xref_lookup()"],
      ["Accumulator", "Standard Rule > Use Accumulator", "Sum numeric values across a loop (line item totals)"],
    ],
    [2400, 3500, 3460]
  ),
  blank(),

  h2("1.5  Document Flow — End-to-End Lifecycle"),
  p("Every document in Sterling has a tracked lifecycle. Understanding this flow is essential for both daily operations and interview discussions."),
  blank(),
  h3("Inbound EDI Document Flow"),
  makeTable(
    ["Step", "What Happens", "Where to Verify in Sterling"],
    [
      ["Step 1", "PARTNER sends file to your AS2 URL or drops on your SFTP inbound path", "Perimeter Server logs / adapter.log"],
      ["Step 2", "AS2ServerAdapter / SFTP adapter receives the file; queues it in Sterling", "Operations > Documents"],
      ["Step 3", "Routing Rule fires: document in mailbox ACME_INBOUND → run BP EDI_850_Inbound", "Mailbox > Routing Rules"],
      ["Step 4", "EDIDeenvelope service: unwraps ISA→GS→ST layers; validates envelope structure", "Operations > BPs > Status Report"],
      ["Step 5", "Sterling auto-generates 997 Functional Acknowledgment (if configured)", "Operations > Documents (find outbound 997)"],
      ["Step 6", "Translation service: runs your compiled map; transforms 850 → XML", "Operations > BPs > Translation step"],
      ["Step 7", "SFTPClientAdapter: PUTs the XML to your ERP server /inbound/po/", "Operations > BPs > SFTP step Status Report"],
      ["Step 8", "997 is enveloped and sent back to partner via AS2 or SFTP", "Operations > Reports > EDI Correlation"],
      ["Step 9", "Operations > Documents retains the raw EDI payload for audit", "Operations > Documents (configurable retention)"],
    ],
    [1000, 4500, 3860]
  ),
  blank(),
  h3("Document Tracking in Operations"),
  makeTable(
    ["Operations Menu", "What You Find", "When to Use"],
    [
      ["Operations > Business Processes", "All BP instances: Running, Halted, Completed", "Daily monitoring; error investigation"],
      ["Operations > Documents", "Raw payloads of every document processed", "Pull exact EDI content for debugging"],
      ["Operations > Reports > EDI Correlation", "Links outbound 850 to its inbound 997 response", "Confirm acknowledgment receipt"],
      ["Operations > Reports > Failed Messages", "All messages that failed processing today", "Morning health check"],
      ["Operations > System > Message", "Low-level system events and adapter logs", "Deep diagnostic when logs aren't enough"],
    ],
    [3000, 3500, 2860]
  ),
  blank(),
  tip("Operations > Documents is your best friend. When a partner says 'we received garbage', pull the raw payload from this screen and compare it byte-by-byte with what your map produced. The answer is almost always visible in the raw file."),
  blank()
);

// ─── NEW: 1.6 Requirements Gathering (NEW SECTION) ─────────────────────────
children.push(
  newSectionBanner("1.6 — Requirements Gathering Master Framework"),
  blank(),
  h2("1.6  Requirements Gathering — The BA's Most Critical Skill", true),
  p("Poor requirements gathering is the single biggest cause of failed EDI implementations. A BA who asks the right questions upfront eliminates rework, delays, and partner escalations. This section gives you the complete framework."),
  blank(),

  h3("Phase 1: Initial Partner Discovery (Before Any Technical Work)", true),
  p("These questions must be answered before you open Map Editor or configure a single adapter. Ask them in writing — never rely on verbal answers for technical EDI requirements:"),
  blank(),
  makeTable(
    ["Requirement Area", "Questions to Ask", "Why It Matters"],
    [
      ["EDI Standard", "Which standard? X12 or EDIFACT? Which version (004010, 005010, D96A)?", "Determines which map version to build; wrong version = immediate rejection"],
      ["Transaction Types", "Which transaction sets do you send/receive? (850, 856, 810, 997, 855, 820...)", "Defines the full scope of mapping work needed"],
      ["Companion Guide", "Do you have a companion guide? Which version? (Always request the current version)", "Override rules may differ from the X12 standard significantly"],
      ["Communication Protocol", "AS2 or SFTP or VAN? If AS2: sync or async MDN? If SFTP: key or password auth?", "Determines adapter configuration and security setup"],
      ["ISA IDs", "What is your ISA06 (Sender ID) and ISA08 (Receiver ID)? What qualifier (ISA05/07)?", "Wrong ISA IDs = every transaction rejected immediately"],
      ["Testing Process", "Do you have a test environment? What is the test contact's email and phone?", "Without a test environment, you cannot safely validate before go-live"],
      ["Go-Live Timeline", "What is your target go-live date? Are there any trading partner deadlines?", "Drives project timeline; some retailers mandate compliance by specific dates"],
      ["Volume & SLAs", "What is the expected transaction volume per day? What is your acknowledgment SLA?", "Informs performance testing and polling interval configuration"],
      ["ERP/WMS Target", "Where should the translated EDI data go? What format does your ERP expect?", "Defines the output format of your maps"],
      ["Compliance Requirements", "Are there HIPAA, SOX, or retail compliance requirements (e.g., VICS, GS1)?", "Impacts data handling, retention, and audit trail requirements"],
    ],
    [2200, 4000, 3160]
  ),
  blank(),

  h3("Phase 2: Companion Guide Analysis Checklist", true),
  p("Every major trading partner has a companion guide that overrides the X12 standard. Work through this checklist for every companion guide you receive:"),
  bullet([run("Page 1: ", { bold: true }), run("Note the companion guide version and date. If it's more than 2 years old, request confirmation it's current.")]),
  bullet([run("Transaction scope: ", { bold: true }), run("List every transaction set defined (850, 856, 810, 997, etc.) and note which are mandatory vs. optional.")]),
  bullet([run("Segment usage: ", { bold: true }), run("Note every segment marked 'MUST USE' — these are additional mandatory segments beyond the base X12 standard.")]),
  bullet([run("Element-level rules: ", { bold: true }), run("Note every element with code list restrictions. These are your most common AK4 error 7 sources.")]),
  bullet([run("Qualifier requirements: ", { bold: true }), run("Note all qualifier-specific rules (e.g., N101='ST' requires N103/N104; N101='BY' uses a different ID format).")]),
  bullet([run("Date/time formats: ", { bold: true }), run("Confirm exact date format requirements (CCYYMMDD vs YYMMDD). Inconsistency causes AK4 error 8 rejections.")]),
  bullet([run("Control number rules: ", { bold: true }), run("Some partners require specific ISA13 number formats or ranges. Note any deviations from standard.")]),
  bullet([run("Acknowledgment requirements: ", { bold: true }), run("Does the partner expect a 997, 999, or both? What is the required acknowledgment SLA (hours)?")]),
  bullet([run("Special segment notes: ", { bold: true }), run("Note any partner-specific usage of MEA, REF, PER, or loop-level requirements (e.g., HL loop structure for 856).")]),
  bullet([run("Contact info: ", { bold: true }), run("Extract all partner EDI contact information from the companion guide. Save it in your trading partner profile notes.")]),
  blank(),
  crit("If you cannot find a specific requirement in the companion guide, do NOT assume the standard applies. Contact the partner's EDI team in writing and document their response. Assumptions in EDI mapping cause rejections that take weeks to diagnose after go-live."),
  blank(),

  h3("Phase 3: Internal System Requirements", true),
  p("Equally important is understanding what your internal ERP/WMS/TMS expects. Interview the ERP team:"),
  makeTable(
    ["Question", "Why It Matters"],
    [
      ["What format does the ERP accept for inbound POs? (XML schema, flat file, JSON, IDoc?)", "Defines your map output format. An incorrect schema causes silent data errors."],
      ["What field naming conventions does the ERP schema use? Is there an XSD or documentation?", "You need the exact field names to build the output side of your map."],
      ["How does the ERP generate invoices/ASNs for outbound EDI? What format does it send?", "Defines your map input format for outbound transactions."],
      ["Are there any calculated or derived fields the EDI system must populate that the ERP does not provide?", "Common examples: SSCC-18 generation, extended price calculation, date conversion."],
      ["What is the ERP's tolerance for data quality issues? Will it fail silently or return errors?", "Critical for error handling design. Silent ERP failures are the hardest to detect."],
      ["Who is the ERP SME for EDI-related questions? (Get a direct contact.)", "You will need this person many times during implementation and UAT."],
    ],
    [5500, 3860]
  ),
  blank(),

  h3("The S2T (Source-to-Target) Mapping Specification", true),
  p([
    run("The ", { size: 22 }),
    run("Source-to-Target (S2T) document", { bold: true, size: 22 }),
    run(" is the most important artifact an EDI BA produces. It is the single source of truth for how every field maps between EDI and the internal system. It must be reviewed and signed off by both the business and technical teams before any map is built.", { size: 22 })
  ]),
  blank(),
  makeTable(
    ["S2T Column", "What to Document"],
    [
      ["EDI Segment / Element", "e.g., PO1-04 (Unit Price)"],
      ["Element Name", "e.g., Unit Price Code"],
      ["Mandatory/Optional", "M=Mandatory, O=Optional, C=Conditional (note dependency)"],
      ["Data Type & Length", "e.g., R (decimal), max 10 chars"],
      ["Target System Field", "e.g., ERP OrderLine.UnitPrice (decimal)"],
      ["Transformation Rule", "e.g., 'Multiply by 1000 and round to 2 decimals if UOM is per-thousand'"],
      ["Default Value", "e.g., 'If blank, default to 0.00'"],
      ["Code List / Validation", "e.g., 'UOM must be EA, CA, DZ, CS — reject if other value'"],
      ["Exceptions / Notes", "e.g., 'Partner ABC sends this as integer cents, not decimal dollars'"],
      ["Status", "Draft / Reviewed / Approved / Implemented / Tested"],
    ],
    [2800, 6560]
  ),
  blank(),
  tip("Never start building a map until the S2T document is at minimum in 'Reviewed' status. Changes to the S2T after mapping is complete are 3x more expensive to implement. Protect the S2T as your project's change control gate."),
  blank()
);

// ─── NEW: 1.7 UAT Master Checklist (NEW SECTION) ──────────────────────────
children.push(
  newSectionBanner("1.7 — UAT & Testing Strategy: The Complete BA Checklist"),
  blank(),
  h2("1.7  UAT & Testing Strategy — Never Go Live Blind", true),
  p("EDI UAT is the phase where most projects cut corners and pay the price in production. A disciplined BA uses this checklist for every new implementation and every significant map change."),
  blank(),

  h3("Unit Testing (Before UAT)", true),
  bullet([run("Map compilation: ", { bold: true }), run("Map compiles without errors or warnings in Map Editor.")]),
  bullet([run("Happy path test: ", { bold: true }), run("Run a standard test file through Map Editor's Test function. Output matches expected S2T mapping.")]),
  bullet([run("Edge case tests: ", { bold: true }), run("Test files with: 1 line item, max line items (per companion guide), all optional fields populated, all optional fields absent.")]),
  bullet([run("Null protection: ", { bold: true }), run("Test with files where all optional elements are absent. No null pointer errors in extended rules.")]),
  bullet([run("Code list validation: ", { bold: true }), run("Test with valid and invalid code values. Invalid values should trigger the appropriate error handling (rejection or default).")]),
  bullet([run("Date/number format: ", { bold: true }), run("Verify date conversions (CCYYMMDD → YYYY-MM-DD) and decimal/integer formatting produce correct output.")]),
  bullet([run("Loop limits: ", { bold: true }), run("Test at, below, and above the companion guide's maximum loop limits. Verify the map handles them correctly.")]),
  bullet([run("Accumulator totals: ", { bold: true }), run("Verify accumulator sums are correct across multiple line items.")]),
  bullet([run("Cross-reference lookups: ", { bold: true }), run("Test with values that exist and do not exist in the cross-reference table. Both paths work correctly.")]),
  blank(),

  h3("Integration Testing (DEV Environment)", true),
  bullet([run("End-to-end BP test: ", { bold: true }), run("Drop test file in inbound path. BP completes successfully. Output lands at ERP/partner.")]),
  bullet([run("997 generation: ", { bold: true }), run("Confirm 997 is generated and contains AK5=A (Accepted). No AK3/AK4 error segments.")]),
  bullet([run("ERP receipt validation: ", { bold: true }), run("Confirm ERP received the translated data. Key fields (PO#, Ship-To, Line items, Quantities) are correct.")]),
  bullet([run("Error path test: ", { bold: true }), run("Introduce a deliberate error (e.g., invalid segment). Confirm BP halts, error alert fires, no data reaches ERP.")]),
  bullet([run("Duplicate detection: ", { bold: true }), run("Send the same file twice. Confirm the second is detected and rejected.")]),
  bullet([run("Rollback test: ", { bold: true }), run("Deploy v2 map. Confirm v1 can be reactivated in under 60 seconds.")]),
  blank(),

  h3("Partner UAT Checklist (TEST Environment)", true),
  p([run("This phase involves the actual trading partner. ", { bold: true, size: 22 }), run("Always conduct partner UAT in ISA15=T (Test) mode. Never use ISA15=P (Production) for testing.", { size: 22 })]),
  blank(),
  makeTable(
    ["UAT Step", "Who", "Pass Criteria", "Sign-Off"],
    [
      ["Partner sends test 850 (or 856/810)", "Trading Partner", "Sterling receives, BP completes, 997=Accepted returned", "☐"],
      ["BA confirms correct data in ERP (PO created, line items correct)", "BA + ERP Team", "All S2T-mapped fields correct in ERP record", "☐"],
      ["BA sends test 856/810 to partner", "BA", "Partner confirms receipt and correct format", "☐"],
      ["Partner confirms 997 format accepted", "Trading Partner", "AK5=A in partner's system; no rejections", "☐"],
      ["Test with max volume: 500 line items, 10 concurrent transactions", "BA", "All BPs complete in <30 seconds; no timeouts", "☐"],
      ["Test error scenarios: invalid PO#, missing mandatory field", "BA", "997 returned with AK5=R; ERP not polluted", "☐"],
      ["Partner confirms all test transactions in THEIR system", "Trading Partner", "Written email confirmation from partner EDI contact", "☐"],
      ["Go-live readiness sign-off", "Business Owner + IT Manager", "Written approval to promote to Production", "☐"],
    ],
    [3500, 1600, 2800, 1460]
  ),
  blank(),

  h3("Production Go-Live Checklist", true),
  bullet([run("ISA15 changed from T to P: ", { bold: true }), run("Verified in the PROD trading partner profile before first live transaction.")]),
  bullet([run("Control numbers reset: ", { bold: true }), run("ISA13 starts fresh in PROD (or confirmed as per partner's requirement).")]),
  bullet([run("Certificate validity: ", { bold: true }), run("Partner's PROD certificate imported. Your PROD certificate shared with partner. Both verified.")]),
  bullet([run("PROD adapter URLs confirmed: ", { bold: true }), run("Partner's PROD AS2 URL or SFTP host confirmed in writing (different from TEST).")]),
  bullet([run("Monitoring in place: ", { bold: true }), run("Alert emails configured for BP failures. You will be notified within 5 minutes of any failure.")]),
  bullet([run("Rollback plan ready: ", { bold: true }), run("Previous map version confirmed available. Rollback procedure documented and tested.")]),
  bullet([run("First 5 transactions monitored: ", { bold: true }), run("Manually verify the first 5 live transactions end-to-end. Confirm partner receipt.")]),
  bullet([run("Hypercare period: ", { bold: true }), run("Commit to 1-week hypercare with partner — daily check-in call at 9 AM for first 5 business days.")]),
  blank(),
  crit("Never promote to production on a Friday or before a public holiday. If something goes wrong, you need your full team available. Schedule go-lives for Tuesday or Wednesday morning for maximum runway."),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 2: REAL-WORLD SCENARIOS (ORIGINAL)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 2",
  "Real-World Scenarios",
  "10 Complex End-to-End Use Cases with Step-by-Step Resolution"
));

const scenarios = [
  {
    num: "Scenario 1", title: "Failed 997 — Partner Rejects Your Outbound 850",
    scenario: "Monday morning. Walmart's EDI team calls: 'We received your 850 purchase order but our system rejected it. We sent you a 997.' Your job: diagnose and fix within 2 hours.",
    steps: [
      "Go to Operations > Business Processes > filter by partner=WALMART, last 24 hours. Find the 850 outbound BP instance that ran this morning. Note the instance ID.",
      "Go to Operations > Reports > EDI Correlation. Search for the outbound 850's ISA13 control number. Find the associated inbound 997 that Walmart sent back.",
      "Open the 997 raw document from Operations > Documents. Navigate to the AK segments. AK1 identifies the interchange (GS control number). AK2 identifies the specific 850. AK3 shows which segment failed (segment ID + line number). AK4 shows which element within that segment failed (element position + error code). AK5 shows the disposition: AK501=R means Rejected.",
      "The most common AK4 error codes: 1=Mandatory element missing, 5=Element too long, 7=Invalid code value, 8=Invalid date. Identify exactly which element and what is wrong.",
      "Go to Operations > Documents. Pull the original raw 850 that was sent. Navigate to the line number shown in AK3. Look at the element position shown in AK4. Compare the actual value in your 850 to Walmart's companion guide requirement.",
      "Open the 850 map in Map Editor (DEV copy). Navigate to the identified segment and element. Find and fix the rule responsible for the incorrect value: if mandatory element missing — a mapping line is broken; if invalid code — your xref table has wrong codes; if date format — your dateconvert format string is wrong.",
      "Recompile the map. Deploy to DEV. Test with the failing file. Confirm 997 would now be AK5=A.",
      "Deploy corrected map to TEST. Run end-to-end test. Confirm Walmart's test environment accepts it.",
      "Deploy to PROD. Resend the rejected 850 (increment the ISA13 — never reuse a rejected control number without incrementing). Monitor for Walmart's 997.",
      "Email Walmart's EDI team: 'We identified and corrected a [field X] formatting issue in our 850. The corrected PO has been resent with a new ISA13. Please confirm receipt.'"
    ],
    root: "A code list mapping error in the PO1-05 (Basis of Unit Price Code) element — the map was outputting 'CP' (cost price) but Walmart's companion guide only accepts 'PE' (per each) or 'CS' (cost). An xref table lookup was returning the wrong value."
  },
  {
    num: "Scenario 2", title: "AS2 Connectivity Failure — MDN Not Received",
    scenario: "Target's EDI team emails: 'We have not received any ASNs (856) from you in 6 hours. Our DC is expecting shipment data.' AS2 is your transport protocol.",
    steps: [
      "Check Operations > Business Processes > filter by BP=EDI_856_Outbound and partner=TARGET. Are BPs completing successfully or halting?",
      "If BPs are Halted: read the Status Report. Look for the AS2 step error. Common messages: 'SSL handshake failed', 'Connection refused', 'MDN not received within timeout', 'Decryption failed'.",
      "If BPs complete but show 'MDN not received': the file reached Target but their system could not decrypt or validate the MDN. Check: has Target's certificate changed? Go to Admin > Certificates. Find Target's cert alias. Check expiry date. If expired: contact Target immediately for new cert.",
      "If 'Connection refused': Target's AS2 server URL may have changed. Verify the AS2 URL in the trading partner profile against Target's current EDI documentation. Test the URL with a browser — does it return an HTTP response?",
      "If 'SSL handshake failed': Sterling's TLS version may be incompatible with Target's server update. Check noapp.log for the TLS negotiation error. Sterling may need to be configured to use TLS 1.2+ (edit jvm.security.properties).",
      "If the AS2 URL is correct and certificate is valid: test the AS2 connection from Sterling Admin > Trading Partners > AS2 > your Target profile > Test Connection. Note the exact error.",
      "Check if Sterling's outbound IP has changed (if IT recently migrated servers). Contact Target: 'Has your AS2 server firewall changed recently? Our IP is X.X.X.X — can you confirm it is whitelisted?'",
      "Once connectivity is restored: requeue all failed 856 BPs via Operations > Business Processes > Status=Halted > Select All > Restart at Failed Step.",
      "Monitor MDN receipts for the next 30 minutes. Confirm all ASNs are acknowledged.",
      "Send Target's EDI team a summary: 'AS2 connectivity was interrupted due to [root cause]. All 856 ASNs have been resent. Please confirm receipt in your system.'"
    ],
    root: "Target had rotated their AS2 encryption certificate as part of their annual security compliance exercise without providing advance notice. Sterling was still using the old cert, causing decryption failures on Target's AS2 receiver."
  },
  {
    num: "Scenario 3", title: "Duplicate 850 POs — Partner Retransmitting",
    scenario: "The ERP team alerts you: 'We are getting duplicate purchase orders. Amazon keeps sending the same 850 PO multiple times. We have 15 duplicate POs in the system.'",
    steps: [
      "Go to Operations > Business Processes > EDI_850_Inbound > filter by Amazon > last 48 hours. Note if you see multiple successful BPs with the same ISA13 control number.",
      "Go to Operations > Documents. Pull multiple 850 files from Amazon. Compare the ISA13 (element 13 in the ISA segment) and the BEG03 (PO Number). If ISA13 is different but BEG03 (the actual PO number) is the same — Amazon retransmitted the same PO with a new control number (this is the more common scenario).",
      "If ISA13 is identical across multiple documents — your deduplication is not working (Sterling should reject identical control numbers from the same trading partner).",
      "Determine the root cause of Amazon's retransmission: they are retransmitting because they did not receive a 997 acknowledgment within their expected SLA window. Check: are your outbound 997s reaching Amazon? Go to Operations > Documents > filter by Type=997 > partner=AMAZON > last 48 hours.",
      "If 997s were not sent or failed delivery: find the 997 outbound BP instances. Read the Status Reports. If 997 delivery via SFTP failed, Amazon continued retransmitting.",
      "For immediate relief: implement a PO-number-based deduplication check. Add a DB lookup step in EDI_850_Inbound to check if BEG03 (PO Number) + Partner ID was processed in the last 30 days. If duplicate: return a 997 Accepted (to stop Amazon's retransmission), log the duplicate, stop processing.",
      "For the 15 duplicate POs in ERP: provide Finance/Ops with the list of duplicated BEG03 PO numbers. They must manually void the duplicates in the ERP.",
      "Fix the root cause of the missing 997: diagnose and fix the 997 delivery failure (if SFTP path issue, fix adapter config; if AS2 issue, fix connectivity).",
      "Monitor: after fixing 997 delivery, confirm Amazon's retransmissions stop within their next retry cycle.",
      "Implement long-term: build a deduplication tracking table. Every inbound 850 from every partner: insert ISA13 + SenderID + BEG03 + Date. Before processing, check against this table."
    ],
    root: "Partner's EDI system retransmits when no 997 is received within their timeout window. 997 delivery was delayed due to a transient network issue, causing the partner to retry."
  },
  {
    num: "Scenario 4", title: "Business Process Hangs — Stuck Thread",
    scenario: "It is 10 AM. The EDI_856_Inbound BP for Home Depot has been in 'Running' state for 4 hours. Normally it completes in 30 seconds. Documents are backing up. You need to diagnose without data loss.",
    steps: [
      "Go to Operations > Business Processes > EDI_856_Inbound. Find the stuck instance. Note the BP Instance ID and which Step it shows (hover over the step in the BPML viewer).",
      "Check what service/step it is stuck on. Common culprits: (a) Translation service waiting on a database lock, (b) SFTP adapter waiting for a TCP connection that never completes, (c) HTTP adapter call to ERP that is hanging on response.",
      "SSH into the Sterling server. Run: tail -200f /opt/IBM/SterlingIntegrator/logs/adapter.log | grep -i 'HOME_DEPOT|error|timeout'. Look for 'connection timed out' or 'read timed out'.",
      "Also check: tail -200f /opt/IBM/SterlingIntegrator/logs/noapp.log | grep -i 'thread|deadlock|waiting'. Look for thread pool exhaustion messages.",
      "If the SFTP adapter is hanging: the ERP SFTP server may be down or the network route changed. Verify by trying an SFTP connect from the Sterling server directly: sftp -i /path/to/key user@erp-server.",
      "If translation is hanging: check if any other BPs are holding a database lock. Run a DB query: SELECT * FROM BPEXECUTION WHERE STATUS='WAITING' AND MODIFIED < SYSDATE-1/24. A wave of stuck BPs indicates DB lock.",
      "For immediate relief (if ERP SFTP is the issue): increase SFTP adapter Data Timeout from default 60s to 300s. Go to Administration > Adapter > SFTPClientAdapter_ERP > Edit Configuration > Data Timeout = 300.",
      "To terminate the stuck BP without losing the document: go to Operations > Business Processes > stuck instance > Terminate. The document is still in Operations > Documents. After fixing the root cause, manually re-queue it by right-clicking the document > Execute Business Process > EDI_856_Inbound.",
      "For thread pool exhaustion: go to Administration > System > Performance Tuning. Increase the 'EDI' and 'default' thread pool sizes. Restart the affected adapter service (not full server restart).",
      "Post-incident: implement a BP watchdog. Create a monitoring BP that runs every 30 minutes, queries for BPs older than 1 hour in Running state, and sends an alert email. This catches hangs proactively."
    ],
    root: "SFTP adapter had a 60-second data timeout, but the ERP server was experiencing high load and not responding to TCP connections within that window. The adapter thread was blocked waiting for a response that never came, consuming a thread pool slot."
  },
  {
    num: "Scenario 5", title: "HL7 837 Claim — Map Failure After Data Change",
    scenario: "A healthcare client sends HL7 837P (Professional Claims) to a payer. After the payer upgraded their clearinghouse system, your inbound 837 → flat file map started failing with 'mandatory element missing: NM109'. Claims processing has halted for 200+ patients.",
    steps: [
      "Go to Operations > Business Processes > HL7_837_Inbound. Find failed instances. Status Report says: Translation failed — mandatory element missing at NM1 loop, element NM109 (Member ID).",
      "Go to Operations > Documents. Pull one of the failed 837 files. Download and open it in a text editor. Navigate to the NM1*QC (patient) segment. Check if NM109 (the member's ID number) is present.",
      "Identify the problem: the payer's new system sends NM108 (ID qualifier) as empty, causing NM109 to shift position. Or the payer is now using a different NM108 qualifier (e.g., MI instead of HN) which confuses the Sterling parser.",
      "Open the 837 map in Map Editor. Navigate to the NM1 loop. Find the NM109 element mapping. Check if 'Mandatory' is set in the element properties. If the payer's new file genuinely omits NM109, this validation is too strict.",
      "Decision: if NM109 is sometimes legitimately absent in the payer's new format, change the element property from Mandatory (M) to Optional (O) in Map Editor. Add a null-check extended rule: if NM109 is blank, use the NM108 value or flag the record.",
      "If NM109 is present but the map can't find it due to a qualifier change: add a conditional rule to read NM109 only when NM108 matches the expected qualifier (MI, HN, etc.). The extended rule should check NM108 first.",
      "Recompile and deploy the map. Test with the failing 837 file using Map Editor's Test function.",
      "For the 200+ failed claims: bulk reprocess. Go to Operations > Business Processes > Status=Halted > Filter by BP name HL7_837_Inbound. Select all > Restart at Failed Step. This re-runs the translation step with the new map.",
      "Monitor for 1 hour. Confirm all 200 BPs complete. Verify the output flat files are landing at the payer interface.",
      "Call the payer's EDI desk to confirm they are receiving and processing the claims. Update your S2T mapping document to reflect the changed element handling."
    ],
    root: "The payer's clearinghouse upgrade changed the NM108 qualifier from HN (Health Insurance Claim Number) to MI (Member ID). The Sterling map had a hardcoded qualifier check that expected HN. When MI was received, the parser could not locate NM109 and reported it as missing."
  },
];

scenarios.forEach(s => {
  children.push(
    h2(`${s.num}: ${s.title}`),
    calloutBox("🎯", "SCENARIO", s.scenario, TEAL_FILL, "00838F"),
    blank(),
    h3("Step-by-Step Resolution"),
    ...s.steps.map((step, i) => numbered(`${i + 1}. ${step}`)),
    blank(),
    h3("Root Cause"),
    p(s.root),
    blank()
  );
});

// Scenarios 6-10 (abbreviated but present)
const scenariosB = [
  {
    num: "Scenario 6", title: "New AS2 Partner Onboarding — Full Configuration (Kroger)",
    summary: "A new retail client (Kroger) requires EDI integration. They send 850 POs via AS2 and expect 997 acknowledgments, 855 PO acknowledgments, and 856 ASNs from you. You have been given their companion guide and their AS2 test environment URL. You have 2 weeks to go live.",
    keyPoints: [
      "Day 1: Gather all requirements — EDI IDs (ISA06/ISA08), AS2 URL, partner certificate (.cer), MDN type (sync/async), document list, test contact, and companion guide.",
      "Day 2–3: Import partner certificate. Create AS2 profile: Encrypt=Yes, Sign=Yes, MDN Type=Synchronous. Create Trading Partner profile with ISA06/ISA08 IDs.",
      "Day 4–6: Read the 850 companion guide. Build 850→XML inbound map, XML→855 and XML→856 outbound maps. Document every field in S2T mapping spec.",
      "Day 7–8: Build Business Processes for each transaction. Create mailboxes and routing rules.",
      "Day 9: Unit test all maps in Map Editor with Kroger's provided test files. Fix all issues. Compile and deploy all maps.",
      "Day 10–11: End-to-end test in ISA15=T mode. Confirm 997 returned, 855 sent, 856 sent. Conduct live partner test. Get written sign-off.",
      "Day 12: Promote to production. Export maps/BPs from TEST; import to PROD. Create PROD trading partner profile with ISA15=P. Monitor first 5 live transactions end-to-end."
    ]
  },
  {
    num: "Scenario 7", title: "SFTP Key Rotation — Partner Changed SSH Keys",
    summary: "A logistics partner (3PL) rotated their SSH server key as part of a security compliance exercise. Now your outbound delivery of 940 Warehouse Orders is failing with 'host key verification failed'. No orders have reached the warehouse for 6 hours.",
    keyPoints: [
      "Confirm the error in Operations > BPs: 'SFTP connection failed: reject HostKey'. Sterling's SFTP library detected a different fingerprint from the remote server.",
      "Contact the 3PL's IT team to confirm they rotated their SSH server host key and request the new public key or fingerprint.",
      "Go to Administration > Trading Partners > SSH Known Host Keys. Find and delete the old entry for the 3PL's hostname. Add the new known host key.",
      "Update the SFTP adapter configuration to reference the new known host entry. Test the connection (Admin > Adapter > SFTPClientAdapter_3PL > Test Connection).",
      "Re-queue all failed 940 documents: Operations > BPs > Status=Halted > Select all > Restart at Failed Step.",
      "Email the 3PL's EDI team confirming resolution. Request minimum 5 business days advance notice before any future key rotation.",
    ]
  },
  {
    num: "Scenario 8", title: "Map Works in DEV, Fails in PROD — Environment Mismatch",
    summary: "You deployed a new 850→XML map to production. It worked perfectly in DEV for 3 weeks. In PROD, every 850 with more than 50 line items fails with 'loop limit exceeded: PO1'. The DEV test files all had fewer than 20 line items.",
    keyPoints: [
      "Read the error in the Status Report: 'EDI Translation failed: Maximum loop count exceeded at segment PO1. Maximum=50, Actual=67.'",
      "Open the failing 850 in Operations > Documents. Count PO1 segments — confirm more than 50 are present.",
      "Open the map in Map Editor. Find the PO1 loop in the input tree. Right-click > Properties. 'Maximum Use' is set to 50. Increase to 9999 (or the companion guide maximum — never set to 1 by default).",
      "Also check: the output XML's LineItem element maxOccurs in the XSD. If set to 50, update to 'unbounded' and re-import.",
      "Recompile, test with 100/500/999-line-item files, promote to TEST, then PROD. Re-process failed PROD BPs.",
      "Post-mortem: add 'Loop Limit Verification' to your map review checklist. Always test with files at maximum expected production size before PROD deployment.",
    ]
  },
  {
    num: "Scenario 9", title: "820 Remittance — Reconciliation Dispute (Amazon)",
    summary: "Finance team escalates: 'We received an 820 remittance from Amazon but the amounts don't match our invoices. The 820 shows payment of $47,230 but we expected $52,105.'",
    keyPoints: [
      "Go to Operations > Documents. Find and download the raw 820 EDI file from the disputed payment date.",
      "Read the 820 structure: BPR02 = payment amount ($47,230). RMR segments within ENT loop = one entry per invoice (RMR02=invoice#, RMR04=original amount, RMR05=paid amount).",
      "Extract all RMR segments. Create a reconciliation table: Invoice#, Expected Amount, Paid Amount. Sum RMR04 (should = $52,105) vs. RMR05 (= $47,230). Difference = $4,875 in deductions.",
      "Look for ADX or MOA segments — these explain deductions (ADX01=AD for advertising allowance, ADX01=FR for freight charges, etc.).",
      "Produce a reconciliation report for Finance: Invoice#, Expected, Paid, Deduction, Reason Code. Identify invoices NOT in 820 (not paid this cycle).",
      "Improve for future: automate 820 reconciliation. Build a map transforming 820 into a CSV that Finance's AR system can auto-import and match against open invoices.",
    ]
  },
  {
    num: "Scenario 10", title: "Emergency Production Rollback — 856 Map Failure",
    summary: "You deployed a new 856 ASN map to production at 3 PM. By 4 PM, 300 ASNs have failed. Target is calling — their receiving system is rejecting your ASNs because the HL loops are malformed.",
    keyPoints: [
      "IMMEDIATE: Do not panic. Go to Deployment > Maps. Find your map name. The old version is still in Sterling's registry.",
      "Find the previous map version. Click 'Set as Default' on the old version. This immediately reverts all new translations to use the previous version.",
      "Confirm rollback: trigger a test 856. Confirm the translation step in Operations > BPs shows the old map.",
      "Re-queue the 300 failed BPs: Operations > BPs > Status=Halted > Filter by EDI_856_Outbound > Select All > Restart at Failed Step.",
      "Call Target's EDI team: 'We experienced a map issue and have rolled back. The 300 ASNs are being resent. Please confirm receipt.'",
      "NOW investigate the v2 bug in DEV. Compare v2 to v1 — check HL level codes, loop max settings, and extended rule syntax. Fix, build v3, test with 50+ ASN files. Deploy v3 to TEST, get QA sign-off, deploy to PROD in next change window.",
    ]
  }
];

scenariosB.forEach(s => {
  children.push(
    h2(`${s.num}: ${s.title}`),
    calloutBox("🎯", "SCENARIO", s.summary, TEAL_FILL, "00838F"),
    blank(),
    h3("Step-by-Step Resolution"),
    ...s.keyPoints.map((kp, i) => numbered(`${i + 1}. ${kp}`)),
    blank()
  );
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 3: LAB EXERCISES (ORIGINAL, SUMMARIZED WITH FULL TRACKS)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 3",
  "Hands-On Lab Plan",
  "25 Structured Exercises + Additional Job-Ready Drills to Build Real Production Skills"
));

children.push(
  note("Complete labs in order — each builds on the previous. Mark each lab complete in the checklist. Aim to finish all 25 within 30 days of starting."),
  blank(),
  h2("Track A: Map Editor Mastery (Labs 1–8)"),
);

const labsA = [
  { num: "Lab 1", name: "Decode a Raw 850 EDI File", time: "30 min", diff: "Beginner",
    steps: ["Download a sample X12 850 file (search 'X12 850 sample EDI file').", "Open in Notepad++. Using only the ISA segment reference, manually decode: sender ID, receiver ID, date, time, version, and test/prod flag.", "Identify every segment (BEG, N1, PO1, CTT). List what each contains.", "Count PO1 loops. Calculate total quantity across all line items manually.", "Success: You can read raw EDI without any tool in under 5 minutes."] },
  { num: "Lab 2", name: "Build 850 → XML Inbound Map", time: "2-3 hours", diff: "Intermediate",
    steps: ["Open Map Editor. File > New Map. Input: X12 005010 850. Output: XML.", "Create output XSD with: PONumber, PODate, BuyerID, ShipTo (Name/Addr/City/State/Zip), LineItems (repeating: LineNum, SKU, Qty, UOM, Price), TotalLines.", "Create direct mapping lines for all header fields.", "Add extended rule on BEG segment to convert PODate from CCYYMMDD to YYYY-MM-DD.", "Add extended rule on PO1 loop to extract BuyerPN (qualifier IN) and VendorPN (qualifier VN).", "Test with the sample 850 from Lab 1. Verify all fields in output XML.", "Success: Translation produces correct XML with all fields; no compiler errors."] },
  { num: "Lab 3", name: "Build XML → 810 Outbound Map", time: "2-3 hours", diff: "Intermediate",
    steps: ["New Map. Input: XML Invoice XSD. Output: X12 005010 810.", "Map all mandatory 810 segments: BIG, N1(BT), IT1 (loop), TDS, CTT, SE.", "Add rule to convert TotalAmount from decimal dollars to cents (multiply * 100) for TDS01.", "Add constant rule on IT1-06 (Basis of Unit Price Code) = 'PE' (per each).", "Test with a sample XML invoice. Verify the output 810 passes validation.", "Success: Valid 810 produced; TDS01 shows cents; all mandatory segments present."] },
  { num: "Lab 4", name: "Extended Rules Masterclass", time: "2 hours", diff: "Intermediate-Advanced",
    steps: ["Add to Lab 2 map: xref_lookup for UOM codes (create UOM_XREF table: EA→EA, CA→CS, PK→PK).", "Add accumulator rule on PO1 loop to sum extended prices (Qty * Price) into a total.", "Add a conditional: if N1-01='ST' set ShipTo fields; if N1-01='BT' set BillTo fields.", "Add null protection: check every optional field for empty string before using it.", "Add a loop counter using loop_index('PO1') to set sequential line numbers.", "Success: All 5 rule patterns work correctly in a single map."] },
  { num: "Lab 5", name: "Build 856 ASN Inbound Map (HL Loops)", time: "3-4 hours", diff: "Advanced",
    steps: ["New Map. Input: X12 005010 856. Output: XML ASN schema.", "Map HL*S (Shipment) level: BSN02=ShipmentID, BSN03=Date.", "Map HL*O (Order) level: PRF01=PO Reference Number.", "Map HL*P (Pack) level: TD3=CarrierPro#, MAN=SSCC-18 labels.", "Map HL*I (Item) level: LIN02/03=UPC, SN1-02=Qty Shipped, SN1-03=UOM.", "Extended rule: use HL03 qualifier to route data to correct XML node (S→Shipment, O→Order, P→Pack, I→Item).", "Success: 856 with 3 orders, 5 packs, 15 items correctly parses into nested XML structure."] },
  { num: "Lab 6", name: "Build 997 Functional Acknowledgment Map", time: "1-2 hours", diff: "Intermediate",
    steps: ["Understand 997 structure: AK1 (group ack), AK2 (transaction ack), AK3 (segment error), AK4 (element error), AK5 (disposition), AK9 (group disposition).", "Build a map that reads an inbound 997 and outputs a CSV report: TransactionDate, ISA13, GS06, ST02, AK501 (disposition code), AK502 (error code if rejected).", "Test with an AK5=A (Accepted) 997. Confirm CSV shows 'Accepted'.", "Create a test 997 with AK5=R (Rejected) and AK3/AK4 segments. Confirm CSV shows rejection details.", "Success: 997 parser correctly identifies accepted vs rejected transactions and reports all errors."] },
  { num: "Lab 7", name: "Cross-Reference Tables & Code Conversion", time: "1-2 hours", diff: "Intermediate",
    steps: ["Create a UOM cross-reference table in Sterling: Admin > Trading Partners > Code Lists. Add 10 entries: EA→EA, DZ→DOZ, CA→CAS, BX→BOX, etc.", "In your 850 map, replace the hardcoded UOM value with an xref_lookup('UOM_XREF', PO1-03, output_var).", "Add error handling: if the xref lookup returns empty (unknown code), set output_var to 'EA' (default) and log a warning.", "Test with valid and invalid UOM codes. Confirm correct conversion and default handling.", "Success: 5 different UOM codes correctly converted; invalid code defaults to 'EA' with warning."] },
  { num: "Lab 8", name: "Map Versioning & Deployment", time: "1 hour", diff: "Beginner",
    steps: ["Take your Lab 2 map. Make a small change (add a date stamp field). Save as v2 (850_to_XML_v2).", "Compile v2. Deploy v2 to Sterling DEV alongside v1.", "In your EDI_850_Inbound BP, update the Translation step to use v2. Test — confirm v2 output is used.", "Now simulate a rollback: go to Deployment > Maps. Set v1 as the active version.", "Test again — confirm v1 output is used.", "Confirm v1 still exists and can be reactivated in 30 seconds.", "Success: You can deploy and roll back a map within 1 minute."] },
];

labsA.forEach(lab => {
  children.push(
    h3(`${lab.num} — ${lab.name}`),
    p([run(`Duration: ${lab.time}  |  Difficulty: `, { size: 21 }), run(lab.diff, { bold: true, size: 21 })]),
    ...lab.steps.map(step => bullet(step)),
    blank()
  );
});

children.push(h2("Track B: Business Process Design (Labs 9–15)"));

const labsB = [
  { num: "Lab 9",  name: "Complete Inbound EDI BP with Error Handling",
    steps: ["Open GPM. New BP named EDI_850_Inbound.", "Add steps: FileSystemAdapter → EDIDeenvelope (X12, create ack=Yes) → Translation (map=850_to_XML_v2) → SFTPClientAdapter → AS2OutboundSend (997).", "Add onFault handler: SMTPSendAdapter sends email to edi-alerts@company.com with BP instance ID and error message.", "Drop a valid 850 into /inbound/850. Confirm BP completes. Break the map intentionally. Confirm alert email.", "Success: End-to-end flow works; error emails arrive when any step fails."] },
  { num: "Lab 10", name: "Mailbox Routing Rules",
    steps: ["Create mailbox: Mailbox > Manage Mailboxes > New. Name: ACME_INBOUND_850.", "Create routing rule: trigger on document arrival in ACME_INBOUND_850 → Execute BP EDI_850_Inbound.", "Drop file into mailbox via Mailbox > Add Message. Confirm BP fires within seconds.", "Create second rule: .xml file in same mailbox routes to different BP.", "Success: Routing fires instantly; correct BP runs based on filename pattern."] },
  { num: "Lab 11", name: "Outbound BP with Enveloping",
    steps: ["Build EDI_810_Outbound BP: FileSystemAdapter → Translation (XML to 810) → EDI Enveloping Service → SFTPClientAdapter.", "Configure Enveloping Service: Sender ID=YOURCO, Receiver ID=PARTNER, GS Functional ID=IN, ISA15=T.", "Test with sample XML invoice. Verify enveloped 810 has correct ISA/GS/ST/SE/GE/IEA structure.", "Verify control numbers increment on each run (ISA13 should be sequential).", "Success: Complete enveloped 810 is delivered to partner SFTP."] },
  { num: "Lab 12", name: "AS2 Partner Configuration",
    steps: ["Generate a self-signed test cert: openssl req -x509 -newkey rsa:2048 -keyout test.key -out test.cer -days 365 -nodes", "Import cert: Administration > Certificates > Import Certificate. Name: TEST_PARTNER_AS2.", "Create AS2 profile: Admin > Trading Partners > AS2 > New. Configure: Encrypt=Yes, Sign=Yes, MDN=Synchronous.", "Test sending via AS2 to a test endpoint (mendelson-e AS2 free tool or Drummond Group).", "Success: AS2 message sent; synchronous MDN received confirming receipt and decryption."] },
  { num: "Lab 13", name: "Scheduling & Polling",
    steps: ["Deployment > Schedules > New Schedule. Name: Poll_850_Every5Min. BP: EDI_850_Inbound. Interval: 5 minutes.", "Enable the schedule. Wait 10 minutes. Drop a file. Confirm BP fires within 5 minutes.", "Create second schedule: daily 8:00 AM, EDI_Reconciliation_Report BP. Cron: 0 8 * * 1-5 (Mon-Fri only).", "Success: Poll schedule runs every 5 min; business-hours schedule runs Mon-Fri only."] },
  { num: "Lab 14", name: "Monitor and Reprocess Failed BPs",
    steps: ["Intentionally introduce syntax error in map extended rule. Deploy the broken map.", "Drop 3 EDI files. All 3 BPs fail.", "Operations > BPs > Status=Halted. Identify all 3 failed instances.", "Fix the map. Deploy corrected version.", "Restart all 3 failed BPs at failed step (multi-select > Restart at Failed Step). Confirm all 3 complete.", "Success: You can identify, diagnose, and bulk-reprocess failed BPs efficiently."] },
  { num: "Lab 15", name: "Sub-Process and Correlation",
    steps: ["Build reusable Error Alert sub-process BP: accepts email subject/body as inputs; sends via SMTP.", "In main BPs, replace inline SMTP with InvokeBusinessProcess call to Error Alert sub-process.", "Build correlation set: correlate outbound 850 ISA13 with inbound 997 ISA13 response.", "Test: send 850, receive 997. Operations > Reports > EDI Correlation — confirm 850 and 997 are linked.", "Success: Sub-process reused across 3 BPs; 850-997 correlation visible in reports."] },
];

labsB.forEach(lab => {
  children.push(
    h3(`${lab.num} — ${lab.name}`),
    ...lab.steps.map(step => bullet(step)),
    blank()
  );
});

children.push(h2("Track C: Advanced & Integration Labs (Labs 16–25)"));
const labsC = [
  { num: "Lab 16", name: "SFTP Adapter Deep Configuration",
    steps: ["Configure SFTPClientAdapter with SSH key auth (generate RSA keypair: ssh-keygen -t rsa -b 2048 -f sterling_test).", "Load private key: Admin > Trading Partners > SSH Private Keys > Import.", "Configure Known Hosts for remote server. Test connection. Confirm fingerprint matches.", "Test GET (inbound) and PUT (outbound) operations. Test failure scenario (wrong host).", "Success: SFTP connects via SSH key; known hosts validation works; errors are readable."] },
  { num: "Lab 17", name: "Deduplication Implementation",
    steps: ["In EDI_850_Inbound BP, after EDIDeenvelope, add DB Lookup step.", "Check if current ISA13 + SenderID was seen in last 7 days.", "If duplicate: branch to 'DuplicateHandler' — log, alert, generate 997 Accepted, stop processing.", "If new: insert ISA13 into tracking table, continue normal flow.", "Test by sending same 850 twice. Confirm second is caught. ERP receives only one copy.", "Success: Duplicate detected; ERP protected; partner's 997 is returned."] },
  { num: "Lab 18", name: "Bulk Onboarding Automation",
    steps: ["Create parameterized BP template: EDI_INBOUND_TEMPLATE. Map name and partner ID come from process data.", "Create routing rule that reads ISA06 (sender ID) and passes it as parameter to template BP.", "Test with 3 different trading partners — all use same template BP but different maps.", "Success: Adding new trading partner requires only new TP profile + new map, not a new BP."] },
  { num: "Lab 19", name: "Certificate Rotation Drill",
    steps: ["Generate two self-signed certs: cert_v1 and cert_v2 using OpenSSL.", "Configure AS2 profile using cert_v1. Send test file. Confirm MDN received.", "Import cert_v2. Update AS2 profile. Send another test file. Confirm MDN received with new cert.", "Practice full cert rotation workflow without service interruption.", "Success: Cert rotation completed; no downtime; new cert works for MDN signing."] },
  { num: "Lab 20", name: "Performance Test & Tuning",
    steps: ["Create 850 test file with 999 line items. Measure translation time via BP step timing.", "Disable 'Validate Output' in Translation service. Re-measure time.", "Remove unnecessary DB lookups from extended rules (replace with in-memory xref tables). Re-measure.", "Document performance improvement from each optimization.", "Success: 999-line 850 processes in under 10 seconds."] },
  { num: "Lab 21", name: "End-to-End Partner Onboarding (Full Simulation — THE MASTER LAB)",
    steps: ["Simulate onboarding 'ACME Corp': 850 inbound + 856 outbound + 810 inbound via SFTP.", "Deliverables: TP Onboarding Form, S2T Mapping Doc, 3 compiled maps, 3 BPs, test results.", "Conduct self-review: would you be comfortable presenting these docs to a client?", "Time yourself: target under 8 hours for complete new partner integration.", "Success: All 8 deliverables produced; end-to-end test passes; S2T doc is complete."] },
  { num: "Lab 22", name: "VAN Connectivity Simulation",
    steps: ["Configure VAN-style mailbox routing: inbound documents drop into 'VAN_INBOUND' mailbox.", "Routing rule extracts ISA06 and routes to correct partner mailbox.", "Simulate VAN pickup: scheduled BP polls VAN mailbox and delivers to correct partner.", "Success: Documents are automatically routed to correct partner mailbox without manual intervention."] },
  { num: "Lab 23", name: "HTTP Outbound Integration",
    steps: ["Build BP: receives 850, translates to XML, then POSTs XML to REST API (https://httpbin.org/post).", "Configure HTTPClientAdapter: URL=https://httpbin.org/post, Method=POST, Content-Type=application/xml.", "Parse HTTP response (httpbin echoes request back). Extract status code.", "If status != 200: branch to error handler.", "Success: XML POSTed to REST API; response captured; errors handled."] },
  { num: "Lab 24", name: "997 Auto-Generation & Routing",
    steps: ["Configure EDIDeenvelope service to auto-generate 997 for all inbound transactions.", "Configure 997 envelope: your ISA06 as sender, original ISA06 as receiver, increment ISA13.", "Route generated 997 back to partner via same protocol they used to send (AS2 → AS2 ack; SFTP → SFTP ack).", "Test with both valid and invalid incoming transactions. Valid → AK5=A. Invalid → AK5=R with AK3/AK4 error details.", "Success: Automatic 997 sent within 60 seconds of receiving any inbound EDI."] },
  { num: "Lab 25", name: "Complete System Health Dashboard",
    steps: ["Build a monitoring BP that runs every 15 minutes and collects: count of Halted BPs, count of Running BPs older than 30 min, last successful 997 timestamp per partner.", "Email daily health report at 8:00 AM to edi-team@company.com: yesterday's transaction volume, error rate, and any open issues.", "Build a cert expiry checker: query all certificates in Sterling; flag any expiring within 90 days.", "Set up pager/SMS alert for any BP halt between 6 PM and 8 AM (after-hours coverage).", "Success: You receive a daily health report; you are alerted within 5 minutes of any after-hours failure."] },
];

labsC.forEach(lab => {
  children.push(
    h3(`${lab.num} — ${lab.name}`),
    ...lab.steps.map(step => bullet(step)),
    blank()
  );
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 4: INTERVIEW Q&A (ORIGINAL + AUGMENTED)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 4",
  "Expert Interview Questions & Answers",
  "25 Real Interview Questions + Soft Skills Deep Dive"
));

children.push(
  note("These are real questions asked by hiring managers at companies with active Sterling environments. Study each answer until you can deliver it in your own words in under 2 minutes. The answers here are frameworks — customize them with your lab experience."),
  blank(),
  h2("Category A: EDI Fundamentals"),
);

const qas = [
  { q: "Q1. Walk me through what happens when a trading partner sends us a purchase order via AS2.",
    a: "When the partner sends an AS2 POST to our Sterling server, the AS2 Server Adapter receives the HTTPS request, decrypts the payload using our private key, and verifies the digital signature using the partner's certificate. It then returns a synchronous MDN receipt confirming delivery. The decrypted file enters Sterling as a document in the partner's inbound mailbox. A routing rule fires and triggers our EDI_850_Inbound Business Process. The BP runs EDIDeenvelope to validate and unwrap the ISA/GS/ST layers, then runs the Translation service with our 850-to-XML compiled map. The translated XML is delivered to the ERP via SFTP or HTTP. Simultaneously, a 997 Functional Acknowledgment is generated, enveloped with the correct control numbers, and sent back to the partner." },
  { q: "Q2. What is the difference between a 997 and an MDN?",
    a: "They acknowledge at completely different layers. The MDN (Message Disposition Notification) is a transport-layer acknowledgment — it confirms the AS2 message was physically received, decrypted, and the digital signature verified. It says nothing about the content of the EDI data. The 997 (Functional Acknowledgment) is an EDI-layer acknowledgment — it validates the actual X12 structure, segments, and elements against the standard and partner-specific rules. A clean MDN with a rejected 997 means the file was delivered but the EDI content is wrong. You need both. In practice: MDN arrives in seconds (synchronous) or minutes (asynchronous). The 997 arrives within the partner's SLA window (typically 24 hours)." },
  { q: "Q3. Explain the ISA segment and why every element matters.",
    a: "The ISA segment is the outermost envelope of every X12 interchange. ISA01-02 (Authorization qualifier/info) are rarely used but must be present as spaces. ISA03-04 (Security qualifier/info) are rarely used. ISA05/06 are the sender ID qualifier and value — must exactly match what the receiving partner expects. Mismatch here = instant rejection. ISA07/08 are the receiver ID qualifier and value — same criticality. ISA09/10 are the date and time of the interchange. ISA11 is the repetition separator (commonly U). ISA12 is the control version number (00401 for 4010, 00501 for 5010). ISA13 is the interchange control number — must be unique per sender-receiver pair; increments with each transmission. ISA14 is the acknowledgment requested flag (1=yes). ISA15 is the usage indicator — T for Test, P for Production. Sending P=Production data to a partner's test environment, or T=Test to production, causes serious problems. ISA16 is the component element separator (commonly >). Every one of these elements has caused production incidents when misconfigured." },
  { q: "Q4. What is a companion guide and how do you use it?",
    a: "A companion guide is a trading partner's implementation guide that specifies how they expect you to implement the X12 standard for their specific requirements. It overrides and supplements the base X12 standard. I use it in three phases: first, during requirements gathering — I read it cover to cover before building any maps to identify all partner-specific mandatory segments, code list restrictions, and special rules. Second, during mapping — it is my reference for every field decision: what qualifiers are expected, what code values are valid, what date formats are required. Third, during testing — when a partner rejects our transactions, the companion guide tells me exactly what their system expects so I can identify the specific element causing the rejection. The most common companion guide rules that catch BAs off guard are: non-standard mandatory segments (like REF segments that aren't in the base standard), partner-specific code lists for UOM or qualifier fields, and specific ISA ID formats that differ from what you'd assume." },
  { q: "Q5. How does Sterling's Map Editor handle loops, and what is the most common mistake BAs make?",
    a: "In the Map Editor, loops are represented as repeating groups in the EDI structure tree. Each loop has a Minimum Use (how many times it must appear) and Maximum Use (how many times it can appear) setting. The most common mistake I see — and have experienced — is leaving the Maximum Use at the default value of 1 for complex loops like PO1 (purchase order line items). In DEV, test files typically have 5-10 line items, so this never triggers. In production, when a retail partner sends a 300-line PO, the translation fails with 'loop limit exceeded'. The fix is to open the loop properties and set Maximum Use to a realistic production value — I use 9999 for most retail PO line item loops unless the companion guide specifies a lower maximum. The second most common mistake is setting up accumulators at the wrong loop level, causing totals to be calculated incorrectly across multiple loop iterations." },
];

qas.forEach(qa => {
  children.push(
    h3(qa.q),
    p(qa.a),
    blank()
  );
});

children.push(h2("Category B: Sterling Technical Knowledge"));

const qasB = [
  { q: "Q6. You open Operations > Business Processes and see 50 BPs in Halted status. What is your first action?",
    a: "First, I do NOT panic and do NOT immediately click 'Restart All'. My first action is to triage. I look for patterns: Are all 50 from the same trading partner? If yes — the issue is likely connectivity or a partner-specific map error. Are they all the same BP name? If yes — likely a map or BP configuration change caused the failure. Are they from multiple partners and multiple BP types? If yes — likely a system-wide issue (database, thread pool, server resource). I open one representative failed instance and read its Status Report. This tells me which step failed and the exact error message. I note the exact error text. Then I check the Sterling logs (noapp.log, adapter.log) for the same time window. Once I've identified the root cause, I fix it in the right environment (usually DEV first), then bulk-restart the failed BPs using multi-select > Restart at Failed Step — not restarting from the beginning, which could cause duplicate processing." },
  { q: "Q7. What is an extended rule in Sterling Map Editor, and when do you use one?",
    a: "An extended rule is a custom script written in Sterling's proprietary scripting language (similar to C syntax) that executes during map translation. Standard rules (direct mapping, constants, accumulators, code list lookups) handle 80% of mapping requirements. Extended rules handle the other 20% — cases where you need conditional logic, string manipulation, calculations, or complex data transformations that standard rules can't handle. I use extended rules for: converting date formats between CCYYMMDD and YYYY-MM-DD using dateconvert(); extracting qualifiers from multi-use segments like N1 (checking N101 to determine if it's ShipTo or BillTo and routing accordingly); calculating derived values (SSCC-18 check digit generation, extended price = Qty * Unit Price); implementing null protection for optional elements (if element is blank, use a default); and conditional mapping (map this field only when a specific qualifier is present). The key discipline is always adding null checks before using any optional element value — the most common extended rule bug is a null pointer error on an absent optional field." },
  { q: "Q8. How do you configure a new SFTP trading partner in Sterling from scratch?",
    a: "I follow a consistent 8-step process: Step 1 — get the partner's SFTP host, port, authentication method (SSH key or password), remote directory for PUT and GET, and filename pattern. Step 2 — if SSH key auth: the partner provides their server's host fingerprint and I provide my public key. I load my private key in Admin > Trading Partners > SSH Private Keys > Import. Step 3 — add the partner's host fingerprint to Admin > Trading Partners > SSH Known Host Keys. Step 4 — configure the SFTP adapter: Admin > Adapters > SFTPClientAdapter > New (or clone existing). Set host, port, auth method, key alias, remote directories, timeout values. Step 5 — test the connection using the Test Connection button. Confirm 'Connected successfully'. Step 6 — create the Trading Partner profile with ISA IDs and document exchange configuration. Step 7 — build or assign the appropriate BP and routing rules for this partner. Step 8 — unit test: manually trigger a PUT to the partner's inbound directory and a GET from their outbound directory. Confirm files arrive at correct paths. Document all configuration in the trading partner profile notes." },
  { q: "Q9. Describe the process for deploying a new map to production.",
    a: "I follow a disciplined promotion process: DEV → TEST → PROD, never skipping steps. In DEV: I build the map, compile, and test extensively using Map Editor's Test function. I run my full test file library — at minimum: happy path, edge cases, max volume, null/optional fields absent. In TEST: I deploy the compiled map to the TEST Sterling environment. I run an end-to-end integration test (actual BP execution, not just Map Editor). I conduct partner UAT if required for this change. I confirm the previous map version is still available for rollback. In PROD: I open a change management ticket. In the deployment window: I export the map from TEST and import to PROD (or promote using Sterling's promotion utilities if available). I set the new version as active. I monitor Operations > Business Processes for the first 10-20 transactions. If any fail with map-related errors, I immediately set the previous version as active (rollback takes under 60 seconds in Sterling). I then investigate the issue in DEV. I never make map changes directly in PROD." },
  { q: "Q10. What is the purpose of the EDI Correlation report in Sterling?",
    a: "The EDI Correlation report links related EDI transactions to each other — primarily linking an outbound transaction to its inbound acknowledgment response. The most common use case: linking an outbound 850 purchase order (identified by its ISA13 control number) to the inbound 997 acknowledgment that the partner sends back. The correlation key is typically the ISA13 (interchange control number). When we send an 850 with ISA13=000000123, Sterling stores this correlation. When the partner sends back a 997 with AK101 referencing GS06 from that same interchange, Sterling matches them using the ISA13 and shows them as correlated in Operations > Reports > EDI Correlation. This allows me to immediately see: did we get a 997 for every 850 we sent? Were they all Accepted or were some Rejected? Without EDI correlation, I'd have to manually match hundreds of outbound transactions to their acknowledgments — correlation automates this. It's my primary tool for proving to auditors that all outbound transactions were acknowledged." },
];

qasB.forEach(qa => {
  children.push(h3(qa.q), p(qa.a), blank());
});

// Category C-F abbreviated
children.push(
  h2("Category C: Security & Compliance"),
  h3("Q11. SFTP shows 'host key verification failed'. What do you do?"),
  p("This error means Sterling's SFTP library detected a different SSH host key on the remote server than what it has stored in its known hosts registry. This happens when: (a) the remote server rotated its host key (security compliance), (b) the remote server was replaced with a different machine, (c) a genuine man-in-the-middle attack (rare but serious). Process: contact the partner's IT team to confirm they rotated their host key. Ask for the new host public key or fingerprint. Go to Administration > Trading Partners > SSH Known Host Keys. Find the old entry for that hostname. Delete it. Import the new host key. Update the SFTP adapter configuration to reference the new known host entry. Test the connection. Never disable host key verification permanently — it is a critical security control. If the partner cannot confirm they rotated their key, escalate to your security team before proceeding."),
  blank(),
  h3("Q12. How do you investigate a complaint that 'the 856 we sent is wrong'?"),
  p("I start by pulling the exact 856 that was sent from Operations > Documents (the outbound payload). I compare it against the partner's companion guide field by field. I check: envelope structure (ISA IDs, version), functional group (GS Functional ID=SH for 856), transaction set header (ST*856), BSN segment (shipment ID, date), HL structure (correct nesting: S→O→P→I levels), and all item-level data. I look specifically at: (1) HL02 parent references — are they correct? (2) SSCC-18 labels in MAN segments? (3) UPCs correctly identified in LIN segments? (4) Shipped quantities in SN1 segments? If I find the error in the output file, I trace it back through the map to the extended rule responsible. I also check if the issue is in the input data from the ERP — the map may be correctly transforming bad input."),
  blank(),
  h3("Q13. A trading partner says their system cannot find files you claim to have sent via SFTP. How do you prove delivery?"),
  p("I go to Operations > Business Processes > the outbound SFTP BP instance. I look at the BPML execution log for the SFTPClientAdapter step. The Status Report will show either success (including remote path, filename, and byte count) or an error. If it shows success, Sterling's client received acknowledgment from the remote server that the file was written. I take a screenshot of the success status including timestamp, remote path, and file size. I share this with the partner. If their system cannot find it: they may be looking in the wrong directory, their pickup process may not have run, or a filename pattern filter is excluding our file. I ask them to do a manual directory listing on the SFTP path to confirm the file is physically present."),
  blank(),
  h3("Q14. How do you handle a certificate that expires during business hours?"),
  p("Prevention is the real answer — you should never be in this situation with proper monitoring. But if it happens: the impact is immediate — all AS2 messages to/from that partner fail. First, alert the business and the partner's EDI team. Then expedite the cert renewal: contact the partner, get their new cert, import it into Sterling, update the AS2 profile, test. This should take under 30 minutes if you have the partner's contact info ready. While waiting: check if we can fall back to SFTP as a temporary measure. Post-incident: implement cert expiry monitoring via a script that checks all Sterling certs monthly and alerts 90 days before expiry."),
  blank(),
  h3("Q15. What is an EDI Control Number and why do they matter?"),
  p("EDI control numbers are sequential numeric identifiers at three levels: ISA13 (interchange, 9 digits), GS06 (functional group), ST02 (transaction set). Each must be unique per sender-receiver pair and must match between opening and closing segments (ISA13=IEA02, GS06=GE02, ST02=SE02). They matter for: (1) Deduplication — partners use ISA13 to detect duplicate transmissions. (2) Reconciliation — control numbers link 997 responses to specific interchanges. Sterling manages control number incrementing automatically, but if you manually copy test files, always change ISA13 — reusing control numbers causes partner rejections and duplicate detection failures."),
  blank(),
  h2("Category D: Operations & Troubleshooting"),
  h3("Q16–Q20 Summary: Operations & Metrics"),
  makeTable(
    ["Question", "Key Answer Point"],
    [
      ["Q16: 5 partners, 1 map change — how do you manage?", "Assess if single shared map or 5 separate maps. Parameterize via cross-reference if possible. Build regression suite for ALL 5 partners. Stagger PROD deployments by partner with 1-hour monitoring between each."],
      ["Q17: How do you ensure map change doesn't break existing partners?", "Three practices: regression testing (test file library: happy path + edge case + error case), versioning (never overwrite — always create v2/v3), monitoring (watch first 10-20 real transactions post-deploy). Rollback ready in <60s."],
      ["Q18: Partner's EDI violates X12 standard — what do you do?", "Options: (1) Ask partner to correct. (2) Implement defensive map tolerating deviation. (3) If structural failure, partner must fix. Always document deviation in writing. Never silently accept non-standard EDI."],
      ["Q19: Right data, partner AP still rejects — investigation approach?", "Get rejection reason in writing. Pull raw 810 from Operations > Documents. Compare vs companion guide: ISA date format, GS date, BIG01 format, N104 party ID, IT1 UOM code, TDS01 decimal vs cents. Request sample accepted 810 from another supplier if guide is ambiguous."],
      ["Q20: What metrics do you track for EDI operational health?", "Transaction success rate (>99.5%), 997 acceptance rate (>99%), 997 SLA compliance, average BP execution time (trend), duplicate detection count, certificate expiry runway (>90 days), onboarding time (<14 days)."],
    ],
    [3200, 6160]
  ),
  blank(),
  h2("Category E & F: Scenario Role-Plays (Q21–Q25)"),
  makeTable(
    ["Question", "Model Answer Framework"],
    [
      ["Q21: 'We are considering replacing Sterling. What is your assessment?'", "Business-first approach: document current portfolio (partners, volumes, SLAs, platform requirements). Assess switching costs (map rebuild, partner recertification, training). Evaluate alternatives against specific requirements. Present risk-adjusted TCO over 3-5 years. Evidence-based recommendation — not platform loyalty."],
      ["Q22: Live test — 'Here is a 997, tell me what is wrong with the 850.'", "Read systematically: AK1 (group ctrl#) → AK2 (transaction ctrl#) → AK3 (segment ID + line number) → AK4 (element position + error code + bad value) → AK5 (disposition). Navigate to line number in original 850. Compare element against X12 standard and companion guide. Narrate your process aloud — they are testing your method."],
      ["Q23: '2 AM, critical partner's orders not coming through. What do you do?'", "Access Sterling remotely via VPN. Operations > BPs > filter by partner > last 4 hours. If Halted: read Status Report. If map error: can current map handle it with restart? If connectivity: test connection in admin UI. If partner's system is down: document and notify — nothing you can do until they're back. No map/BP changes at 2 AM without emergency change process."],
      ["Q24: Mid-project, partner's companion guide is different from what they told you.", "Treat as formal scope change. Document discrepancy. Assess impact on already-built maps. Estimate additional effort. Notify manager and account team (timeline may change). Confirm new guide is final with partner in writing. Update S2T. Rework affected maps. Get partner written acknowledgment of scope change."],
      ["Q25: 'Most complex EDI problem you've solved — walk me through it.'", "Structure: [Describe the problem] → [Why it was complex] → [Your systematic approach] → [The solution] → [The outcome] → [The key learning]. Use your Lab 5 (856 HL loops), Lab 17 (deduplication), or Lab 21 (end-to-end onboarding) experience as the basis."],
    ],
    [3000, 6360]
  ),
  blank()
);

// ─── NEW: SOFT SKILLS SECTION ──────────────────────────────────────────────
children.push(
  newSectionBanner("BONUS SECTION — Soft Skills & Stakeholder Management"),
  blank(),
  h2("Soft Skills & Stakeholder Management — The BA Advantage", true),
  p("Technical skills get you hired. Soft skills get you promoted. An EDI BA who can translate technical complexity into business language, manage partner relationships under pressure, and lead scope conversations confidently is worth twice the technical expert who cannot communicate. This section covers the interpersonal skills that top EDI BAs demonstrate."),
  blank(),

  h3("Communicating EDI Errors to Non-Technical Stakeholders", true),
  p("When an EDI transaction fails, non-technical stakeholders (Finance, Operations, Sales) need to understand what happened without being buried in technical jargon. Use this translation guide:"),
  blank(),
  makeTable(
    ["Technical Fact", "Business Language Translation"],
    [
      ["AK5=R on 850 from Walmart", "Walmart's system rejected our purchase order. The specific issue was [X]. We are correcting and resending. Estimated resolution: 2 hours."],
      ["SFTP connection timeout to 3PL", "Orders are temporarily not reaching the warehouse due to a server connectivity issue with our logistics partner. We are investigating and expect to restore in 1 hour."],
      ["Loop limit exceeded on 856 map", "Our system had a configuration limit that prevented processing large shipment notifications (more than 50 items). We've corrected this and are reprocessing the affected transactions."],
      ["Certificate expired on AS2", "Our encrypted communication channel with [Partner] has expired and needs to be renewed. This is like an HTTPS certificate on a website. We are renewing now — expected resolution in 30 minutes."],
      ["ISA15=T in PROD", "We accidentally sent test data to our production trading partner. This can cause duplicate orders or confusion. We are recalling and resending correct production data immediately."],
    ],
    [3500, 5860]
  ),
  blank(),
  tip("Keep a 'translation cheat sheet' in your personal notes — common technical error → business impact → resolution timeline. When the phone rings at 8 AM with a panicked Finance director, you want to give a calm, clear, 30-second status update before you even start diagnosing."),
  blank(),

  h3("Managing Trading Partner Relationships", true),
  p("Your trading partners are external customers. How you communicate with them under pressure directly impacts the business relationship. These principles apply:"),
  bullet([run("Always respond within 1 hour during business hours. ", { bold: true }), run("Even if you don't have a resolution, send: 'We have received your report and are investigating. I will update you by [specific time].'")]),
  bullet([run("Never blame the partner in writing. ", { bold: true }), run("Even if the error is clearly on their side, say: 'We've identified a discrepancy between our configuration and the data received.' Escalate partner-side issues through proper channels, not email threads.")]),
  bullet([run("Confirm everything in writing. ", { bold: true }), run("After any phone call: send a summary email. 'Per our call, we agreed to: [1], [2], [3]. Please confirm.' This protects you and creates an audit trail.")]),
  bullet([run("Provide proactive status updates. ", { bold: true }), run("If resolution is taking longer than expected, send an update every 30 minutes. Partners hate silence more than bad news.")]),
  bullet([run("Build personal relationships. ", { bold: true }), run("Know your top 5 partner EDI contacts by name. When you've worked through an incident together, you've earned trust. That trust is irreplaceable at 2 AM.")]),
  blank(),

  h3("Scope Management in EDI Projects", true),
  p("EDI projects are notorious for scope creep. Every partner adds requirements, every companion guide update changes the spec. Manage scope with these practices:"),
  bullet([run("Lock the S2T document: ", { bold: true }), run("Once reviewed and signed off, any change to the S2T is a formal scope change. No verbal agreements. Get it in writing.")]),
  bullet([run("Version control everything: ", { bold: true }), run("Every S2T document, companion guide, and test file should have a version number and date. 'Companion Guide v2.1 — received 2026-03-15' is unambiguous. 'The companion guide they sent us' is not.")]),
  bullet([run("Estimate change impact immediately: ", { bold: true }), run("When a scope change arrives, respond with: 'I'll have an impact assessment to you within 24 hours.' Never say 'sure, no problem' without understanding the full impact.")]),
  bullet([run("Use a RACI for EDI onboarding: ", { bold: true }), run("Responsible (you), Accountable (manager), Consulted (partner EDI team, ERP team), Informed (business owner, Finance, Logistics). Making accountability clear prevents requirements from arriving through unofficial channels.")]),
  blank(),

  h3("Handling Pressure Situations", true),
  p("EDI failures are visible, business-critical, and time-pressured. Your calm under pressure is your most valuable soft skill. Follow this mental framework:"),
  makeTable(
    ["Pressure Situation", "Mental Framework", "First Action"],
    [
      ["Partner is live on the phone, transactions are failing", "Facts first, emotion never. Get information before you speak.", "Ask: 'Can you tell me the exact error message and the last successful transaction time?' This buys you 2 minutes to check Operations."],
      ["Your manager's manager walks over and asks 'what's happening?'", "Confidence, not certainty. 'I have it. I'm isolating the layer right now. I'll update you in 15 minutes with root cause.'", "Give one sentence of status. Do not speculate. Return to diagnosis."],
      ["You caused the incident (your map change broke production)", "Own it. Fix it. Explain it. No defensiveness.", "'I deployed a map change that had an issue I didn't catch in testing. I've rolled back. I'm analyzing what I missed and will have a corrective action plan in 2 hours.'"],
      ["Partner deadline is today and you're not ready", "Scope it ruthlessly. What MUST work today vs. what can be deferred?", "Call partner: 'We can go live today with X. Y will follow by [date]. Can you accept this?' Never go live with untested code to meet a deadline."],
    ],
    [2600, 3200, 3560]
  ),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 5: TROUBLESHOOTING (ORIGINAL)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 5",
  "Comprehensive Troubleshooting Guide",
  "Diagnosis Checklists for Every Scenario — Your Production Lifeline"
));

children.push(
  h2("5.1  Systematic Diagnosis Framework"),
  p("Every Sterling issue follows the same diagnostic chain. Use this framework before diving into specifics:"),
  numbered("CONFIRM the issue: Can I reproduce it? Get the exact error message, BP instance ID, and document ID."),
  numbered("ISOLATE the layer: Is it a Transport issue (adapter/network), Translation issue (map), Business Logic issue (BP), or Data issue (bad input from partner)?"),
  numbered("COLLECT evidence: Pull the raw document from Operations > Documents. Get the Status Report from the failed BP step. Check the relevant log file."),
  numbered("HYPOTHESIZE: What is the most likely root cause given the evidence? What are alternatives?"),
  numbered("TEST the fix: Apply the fix in DEV. Reproduce the original failure scenario. Confirm the fix resolves it."),
  numbered("DEPLOY: Follow change management. Monitor for 1-2 hours post-deploy."),
  numbered("DOCUMENT: Update runbook and S2T doc. Add test case to regression suite."),
  blank(),

  h2("5.2  AS2 / Certificate Issues"),
  makeTable(
    ["Symptom", "Probable Cause", "Diagnosis", "Fix"],
    [
      ["MDN: decryption-failed", "Partner cert expired or wrong cert used to encrypt", "Admin > Certificates > check expiry on partner cert alias", "Import partner's new cert; update AS2 profile cert reference"],
      ["MDN: signature-verification-failed", "Your signing cert doesn't match partner's stored copy", "Confirm partner has your current public cert", "Share your current public cert with partner"],
      ["No MDN received (async)", "Wrong async MDN URL or partner can't reach your URL", "Check MDN URL; test if reachable; check firewall", "Update MDN URL; open firewall port; or switch to sync MDN"],
      ["MDN: processing-error", "Partner's AS2 system had internal error", "Contact partner's EDI team — their system error", "Partner resolves their error; you may need to resend"],
      ["SSL handshake failed", "Outdated TLS version (Sterling TLS 1.0 vs partner TLS 1.2+)", "Check noapp.log for SSL negotiation error", "Update Sterling SSL config to TLS 1.2+; update cipher list"],
      ["HTTP 401 Unauthorized", "Wrong AS2 ID or auth credentials", "Compare AS2 ID against partner expectation", "Correct AS2 ID; update credentials"],
    ],
    [2200, 2500, 2500, 2160]
  ),
  blank(),

  h2("5.3  SFTP / FTP Issues"),
  makeTable(
    ["Symptom", "Probable Cause", "Diagnosis", "Fix"],
    [
      ["Auth failed: publickey", "Wrong SSH key alias or key not uploaded to partner", "Check key alias; verify partner has your public key", "Re-upload public key to partner; verify alias matches"],
      ["Host key verification failed", "Partner rotated their SSH server host key", "Contact partner to confirm key rotation; get new fingerprint", "Update known hosts in Sterling with new partner host key"],
      ["Connection refused", "Wrong host/port or partner firewall blocking Sterling's IP", "Test: sftp -P 22 user@partner-host from Sterling CLI", "Correct host/port; ask partner to whitelist Sterling's IP"],
      ["File not found: remote path", "Wrong remote directory or case sensitivity (Linux)", "Check adapter remote directory config", "Correct remote directory path (case-sensitive on Linux)"],
      ["Permission denied", "SFTP user lacks write permission to remote directory", "Manual SFTP test: try PUT a test file from CLI", "Ask partner to grant write permission to SFTP user"],
      ["Connection timeout", "Partner SFTP server slow; data timeout too short", "Check adapter.log for timeout timestamp", "Increase Data Timeout from 60s to 300s in adapter config"],
      ["File delivered but partner says empty", "Zero-byte file (translation produced empty output)", "Check translation output size in Operations > Documents", "Debug map — likely empty input or rule producing no output"],
    ],
    [2200, 2500, 2500, 2160]
  ),
  blank(),

  h2("5.4  Translation / Map Issues"),
  makeTable(
    ["Symptom", "Probable Cause", "Diagnosis", "Fix"],
    [
      ["Mandatory segment missing", "Map is not outputting a required segment", "Check map output side — segment mapped? Condition suppressing it?", "Add mapping rule or remove suppression condition"],
      ["Loop limit exceeded: PO1", "More loop iterations than map maximum allows", "Check Map Editor PO1 loop Maximum Use setting", "Increase Max Use to 9999; recompile; redeploy"],
      ["Invalid code value (AK4 error 7)", "Output code not in partner's allowed code list", "Compare output code against companion guide allowed values", "Add conditional mapping or xref lookup to convert code"],
      ["Null pointer in extended rule", "Rule tried to use value from absent optional element", "Check rule for null protection", "Add: if $PO1/PO104 <> '' then... before using element"],
      ["Date format rejected", "Date in wrong format", "Check companion guide date format; check dateconvert rule", "Fix dateconvert second parameter to match expected format"],
      ["Translation produces empty file", "Map compiled against wrong version or input doesn't match schema", "Verify ISA12 version matches map version", "Correct version mismatch; recompile map"],
      ["Extended rule compile error", "Typo; undeclared variable; wrong function name", "Map Editor output window shows exact error line", "Fix syntax; declare all variables at top of rule"],
      ["Accumulator wrong total", "Accumulator reset in wrong place or wrong loop", "Check accumulator rule assignment", "Move accumulator init to correct loop OnBegin; verify scope"],
    ],
    [2200, 2500, 2500, 2160]
  ),
  blank(),

  h2("5.5  Business Process Issues"),
  makeTable(
    ["Symptom", "Probable Cause", "Diagnosis", "Fix"],
    [
      ["BP stuck in Running >30 min", "Thread blocked on I/O; adapter timeout too low", "Check which step is running; check adapter.log", "Increase adapter timeout; terminate stuck BP; reprocess doc"],
      ["BP fails immediately — no step runs", "BP syntax error or missing service reference", "Check noapp.log immediately after BP launch", "Validate BPML XML; confirm service name is correct"],
      ["BP halts at enveloping step", "Wrong trading partner profile or ISA control config", "Check enveloping service config — partner name, ISA IDs", "Correct trading partner reference in enveloping service"],
      ["Routing rule not firing", "Mailbox name mismatch or rule condition wrong", "Check Mailbox > Routing Rules; confirm trigger conditions", "Fix mailbox name in rule or correct trigger expression"],
      ["Schedule not running BP", "Schedule disabled or cron expression wrong", "Check Deployment > Schedules; confirm enabled and expression", "Enable schedule; fix cron syntax; verify server timezone"],
      ["Sub-process not invoked", "BP name in InvokeBusinessProcess is wrong or BP is disabled", "Check process data for BP name value; check BP enabled status", "Correct BP name; enable sub-process BP"],
    ],
    [2200, 2500, 2500, 2160]
  ),
  blank(),

  h2("5.6  Log File Reference"),
  makeTable(
    ["Log File", "Location", "What It Contains", "When to Check"],
    [
      ["noapp.log", "/opt/IBM/SterlingIntegrator/logs/noapp.log", "Application-level errors, BP execution errors, system events", "First log to check for any BP error"],
      ["adapter.log", "/opt/IBM/SterlingIntegrator/logs/adapter.log", "All adapter activity: SFTP connections, AS2 sends, HTTP calls", "Transport layer errors (SFTP/AS2/HTTP failures)"],
      ["map.log", "/opt/IBM/SterlingIntegrator/logs/map.log", "Map compilation and translation events", "Map-specific errors, extended rule execution problems"],
      ["security.log", "/opt/IBM/SterlingIntegrator/logs/security.log", "Login attempts, certificate events, authorization failures", "Auth failures, certificate issues, access control"],
      ["jvm.log", "/opt/IBM/SterlingIntegrator/logs/jvm.log", "Java virtual machine events, out of memory errors, GC pauses", "Performance issues, server crashes, OOM errors"],
    ],
    [1800, 3600, 2700, 1260]
  ),
  blank(),

  h2("5.7  Governance, Compliance & Audit Readiness", true),
  newBadge("This section is critical for BAs in regulated industries (healthcare, retail, finance). Auditors will ask these questions."),
  blank(),
  h3("HIPAA EDI Compliance Checklist (Healthcare BAs)", true),
  bullet([run("5010 Compliance: ", { bold: true }), run("All X12 HIPAA transactions (837, 835, 270/271, 276/277, 278, 820, 834) must use version 005010 — not 004010. This is federally mandated.")]),
  bullet([run("NPI Usage: ", { bold: true }), run("National Provider Identifier must be in NM109 wherever a provider is referenced. Missing NPI = claim rejection + potential HIPAA violation.")]),
  bullet([run("PHI Handling: ", { bold: true }), run("Protected Health Information (patient names, DOBs, SSNs, diagnosis codes) in EDI data must be encrypted in transit (HTTPS/SFTP) and at rest. Confirm Sterling's data retention policy for HIPAA transactions.")]),
  bullet([run("Data Retention: ", { bold: true }), run("HIPAA requires retaining EDI transaction records for 6 years from creation date or last effective date. Ensure Sterling's document retention is configured accordingly.")]),
  bullet([run("Audit Trail: ", { bold: true }), run("Every claim submission (837) must have a corresponding acknowledgment (999/997) and status response (277). Operations > Reports > EDI Correlation must show complete transaction chains for auditors.")]),
  blank(),
  h3("SOX/Retail Compliance Checklist (Finance/Retail BAs)", true),
  bullet([run("Invoice Integrity: ", { bold: true }), run("Every outbound 810 invoice must be traceable to an approved purchase order (850). Your EDI system must provide this linkage for SOX audit purposes.")]),
  bullet([run("Payment Reconciliation: ", { bold: true }), run("Every inbound 820 remittance must be reconciled to open invoices. Automate this process — manual reconciliation is an audit risk.")]),
  bullet([run("Control Number Uniqueness: ", { bold: true }), run("ISA13 control numbers must be unique per trading partner pair. Duplicate control numbers can indicate replay attacks or processing errors — both are audit findings.")]),
  bullet([run("Change Management: ", { bold: true }), run("Every map deployment and BP change must be documented in your change management system. 'I just pushed it' is not acceptable in a SOX-controlled environment.")]),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 6: ALTERNATIVE TOOLS
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(...sectionDivider(
  "SECTION 6",
  "Alternative Tools & Platforms",
  "MuleSoft, Boomi, TIBCO, SAP PI/PO, Azure Logic Apps — Knowing the Ecosystem"
));

children.push(
  h2("6.1  Platform Concept Mapping — Sterling to Other Tools"),
  p("Your Sterling knowledge translates directly to other platforms. The concepts are identical — only the tool names and interfaces differ. Use this master translation table to orient yourself quickly on any new platform."),
  blank(),
  makeTable(
    ["Sterling Concept", "MuleSoft Anypoint", "Dell Boomi", "TIBCO BusinessConnect", "SAP PI/PO", "Azure Logic Apps"],
    [
      ["Business Process (BP)", "Mule Flow", "Boomi Process", "Trading Manager Process", "Integration Process (iFlow)", "Logic App Workflow"],
      ["Map Editor", "DataWeave Mapper", "Boomi Map Component", "Trading Manager Map", "Graphical Message Mapping", "Data Mapper Action"],
      ["GPM (Visual BP design)", "Anypoint Studio", "Process Canvas", "Process Editor", "Integration Designer", "Logic App Designer"],
      ["Trading Partner Profile", "Partner Manager (API-based)", "Trading Partner tab", "Community Manager record", "Communication Channel + Party", "Connector configuration"],
      ["AS2 Adapter", "AS2 Connector (paid)", "Boomi AS2 Connector", "Built-in AS2", "SAP AS2 Adapter", "AS2 Logic App Connector"],
      ["SFTP Adapter", "SFTP Connector", "SFTP Connector", "Built-in SFTP", "SAP SFTP Adapter", "SFTP Connector"],
      ["Mailbox / Routing Rule", "VM Queue + Listener", "Boomi Listener Process", "Message Queue", "Message Queue Channel", "Trigger + Routing Action"],
      ["EDI De-envelope", "EDI Module (X12/EDIFACT)", "Boomi EDI Module", "Built-in EDI parser", "SAP IDoc / EDI Adapter", "Flat File Decoder"],
      ["Extended Rules", "DataWeave expressions", "Boomi map function scripts", "Transformation rules", "XSLT / Graphical mapping", "Expression language"],
      ["Cross-Reference Table", "Object Store (key-value)", "Boomi Connection Lookup", "Trading Partner code lookup", "Value Mapping / RFC", "Integration Account Maps"],
      ["Operations > BPs", "Runtime Manager (CloudHub)", "Process Reporting", "Partner Manager logs", "Monitor (NWA)", "Run History"],
      ["Certificate Store", "Anypoint Security (TLS contexts)", "Certificate management", "Partner certificate store", "Keystore / TLS", "Azure Key Vault"],
    ],
    [2200, 1700, 1600, 1600, 1600, 1660]
  ),
  blank(),

  h2("6.2  MuleSoft Anypoint Platform"),
  p("MuleSoft is an API-first integration platform owned by Salesforce. It excels at modern REST/SOAP API integrations and microservices. Its B2B/EDI capabilities are less mature than Sterling for traditional EDI, but it is widely adopted in organizations with heavy Salesforce or modern API use."),
  blank(),
  makeTable(
    ["Dimension", "IBM Sterling", "MuleSoft Anypoint"],
    [
      ["Primary Strength", "B2B EDI, partner onboarding at scale", "API integration, microservices, Salesforce ecosystem"],
      ["EDI Support", "Native X12, EDIFACT, HL7, TRADACOMS", "X12 via EDI module; EDIFACT less mature"],
      ["Mapping Language", "Proprietary Map Editor + extended rule scripts", "DataWeave — powerful functional language; steeper learning curve"],
      ["Hosting", "On-premise or IBM Cloud", "CloudHub (SaaS), Anypoint Runtime Fabric (k8s), hybrid"],
      ["Partner Portal", "None (admin-only onboarding)", "Anypoint Partner Manager (self-service portal)"],
      ["Pricing", "Enterprise license (high upfront)", "API-call or message-volume based"],
      ["Learning Curve", "Map Editor is proprietary but approachable for BAs", "DataWeave requires coding skills; less BA-friendly"],
      ["Best For", "Large retail/healthcare EDI at scale", "Modern API ecosystems; Salesforce integrations"],
    ],
    [2600, 3380, 3380]
  ),
  blank(),
  tip("If asked 'Can you use MuleSoft?' in an interview: yes — your Sterling knowledge of B2B concepts (EDI standards, trading partner management, AS2, 997 acks) transfers directly. The tool differences are learnable. Your EDI domain expertise is what takes years to build."),
  blank(),

  h2("6.3  Dell Boomi AtomSphere"),
  p("Boomi is a cloud-native iPaaS (Integration Platform as a Service). It is highly regarded for ease of use, fast deployment, and broad connector library. Its B2B/EDI capabilities are solid for mid-market customers."),
  makeTable(
    ["Dimension", "IBM Sterling", "Dell Boomi"],
    [
      ["Deployment", "On-premise or IBM Cloud", "Cloud-native; Atom (local agent) for on-prem data"],
      ["EDI Handling", "Native EDI engine, full lifecycle", "Boomi EDI module handles X12/EDIFACT; solid but less feature-rich"],
      ["Map Design", "Visual Map Editor with extended rules", "Boomi Map with scripting — similar concept, different UI"],
      ["Partner Onboarding", "Manual via Sterling dashboard", "Trading Partner tab with guided configuration"],
      ["Monitoring", "Operations > Business Processes (deep)", "Process Reporting — good but less detailed than Sterling"],
      ["Best For", "Large-scale traditional EDI", "Mid-market; cloud-first organizations; Salesforce, NetSuite integrations"],
    ],
    [2600, 3380, 3380]
  ),
  blank(),

  h2("6.4  TIBCO BusinessConnect / TIBCO B2B"),
  p("TIBCO BusinessConnect is an enterprise B2B gateway historically strong in financial services and manufacturing. It handles EDI, RosettaNet, and SWIFT."),
  makeTable(
    ["Dimension", "IBM Sterling", "TIBCO BusinessConnect"],
    [
      ["Market Focus", "Retail, healthcare, logistics", "Financial services, manufacturing, energy"],
      ["EDI Standards", "X12, EDIFACT, HL7, TRADACOMS", "X12, EDIFACT, SWIFT, RosettaNet, AS4"],
      ["Architecture", "Central hub with perimeter server", "Hub-and-spoke or point-to-point"],
      ["AS2 Support", "Strong — industry standard", "Supported; not as dominant"],
      ["Financial Standards", "Limited SWIFT support", "Native SWIFT adapters; SEPA support"],
    ],
    [2600, 3380, 3380]
  ),
  blank(),

  h2("6.5  SAP PI/PO (Process Integration / Process Orchestration)"),
  p("SAP PI/PO is primarily used by SAP-centric organizations to integrate SAP with external partners and internal systems. If your organization runs SAP ERP, you will likely encounter SAP PI/PO for EDI or IDoc-to-EDI translations."),
  makeTable(
    ["Dimension", "IBM Sterling", "SAP PI/PO"],
    [
      ["Primary Use", "B2B EDI with external trading partners", "SAP-to-SAP and SAP-to-external integrations"],
      ["EDI Handling", "Native EDI engine; full lifecycle", "Via EDI adapter + EDI converter; relies on IDocs internally"],
      ["Mapping Tool", "Sterling Map Editor", "SAP Graphical Message Mapping + XSLT"],
      ["Integration with ERP", "Via SFTP, MQ, JDBC to any ERP", "Tight native integration with SAP ECC/S4HANA"],
      ["Best For", "Any-to-any EDI at scale", "SAP shops needing EDI connectivity for their ERP"],
    ],
    [2600, 3380, 3380]
  ),
  blank(),
  note("Sterling and SAP PI/PO often coexist: Sterling handles external partner B2B (the AS2/SFTP/EDI layer), while SAP PI/PO handles internal SAP-to-Sterling connectivity (translating Sterling XML output into SAP IDocs). Understanding both touchpoints makes you highly valuable in enterprise SAP environments."),
  blank(),

  h2("6.6  Azure Logic Apps + Azure Integration Services"),
  p("Microsoft's Azure Integration Services is a modern cloud platform that includes Logic Apps (workflow orchestration), API Management, Service Bus (messaging), and Event Grid (events). It has growing B2B/EDI capability through Azure Integration Account."),
  makeTable(
    ["Dimension", "IBM Sterling", "Azure Logic Apps"],
    [
      ["Deployment Model", "On-premise or IBM Cloud", "Azure cloud; consumption or standard plan"],
      ["EDI Support", "Native X12, EDIFACT, HL7", "X12, EDIFACT via Integration Account (AS2, X12 built-in)"],
      ["Map/Transform", "Sterling Map Editor (.map files)", "Azure XSLT maps or Liquid templates"],
      ["Monitoring", "Operations > Business Processes", "Azure Monitor + Logic Apps Run History"],
      ["AS2 Support", "Full enterprise AS2", "AS2 built into Logic Apps — handles MDN, signing, encryption"],
      ["Pricing", "Enterprise license", "Pay-per-execution; cost-effective for low volumes"],
      ["Best For", "High-volume, complex EDI at enterprise scale", "Microsoft-centric orgs; low-moderate EDI volume; rapid prototyping"],
    ],
    [2600, 3380, 3380]
  ),
  blank(),

  h2("6.7  Which Platform for Which Scenario?"),
  makeTable(
    ["Scenario", "Best Platform", "Why"],
    [
      ["High-volume retail EDI (Walmart, Target, Kroger)", "IBM Sterling", "Purpose-built; proven at scale; native companion guide compliance"],
      ["Healthcare EDI (837, 835, 834) at scale", "IBM Sterling or Rhapsody", "Native HL7 and X12; HIPAA compliance tooling"],
      ["Salesforce ecosystem integration", "MuleSoft", "Native Salesforce connectivity; API-first design"],
      ["SAP ERP internal + external EDI", "SAP PI/PO + Sterling", "PI handles IDoc; Sterling handles partner EDI — complement each other"],
      ["Start-up or mid-market, cloud-first", "Boomi or Azure Logic Apps", "Lower cost; faster deployment; less infrastructure overhead"],
      ["Financial services (SWIFT, SEPA)", "TIBCO or IBM Sterling", "SWIFT adapters; financial EDI standards support"],
      ["Modern API-to-API integration", "MuleSoft or Azure API Mgmt", "REST/GraphQL native; developer-friendly; API marketplace"],
      ["Manufacturing / RosettaNet", "TIBCO or Sterling", "RosettaNet adapters; XML-based B2B standards"],
    ],
    [3400, 2600, 3360]
  ),
  blank(),
  tip("In interviews: never say 'Sterling is the best'. Say 'Sterling is the right choice when X, Y, Z conditions apply. In other scenarios, platforms like Boomi or MuleSoft may be more appropriate.' Interviewers respect nuanced, business-driven judgment over platform loyalty."),
  blank()
);

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// QUICK REFERENCE / CHEAT SHEET
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
children.push(
  new Paragraph({ children: [new PageBreak()] }),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: noBorder(),
        shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
        margins: { top: 240, bottom: 240, left: 360, right: 360 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "QUICK REFERENCE", font: "Arial", size: 44, bold: true, color: WHITE })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Production-Ready Cheat Sheet", font: "Arial", size: 28, color: BLUE_LIGHT, italics: true })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Everything You Need at Your Fingertips", font: "Arial", size: 22, color: BLUE_LIGHT })] }),
        ]
      })
    ]})]
  }),
  blank(),

  h2("EDI X12 Transaction Set Library"),
  makeTable(
    ["Transaction Set", "Name", "Direction", "Key Use Case"],
    [
      ["810", "Invoice", "Outbound", "You bill your trading partner for goods/services shipped"],
      ["820", "Payment Order / Remittance", "Inbound", "Partner tells you which invoices they are paying and any deductions"],
      ["830", "Planning Schedule", "Inbound", "Forecast demand signals from buyer (used in VMI programs)"],
      ["832", "Price/Sales Catalog", "Both", "Item and pricing data exchange between retailer and vendor"],
      ["834", "Benefit Enrollment", "Both", "Healthcare member enrollment data (insurance plans)"],
      ["835", "Healthcare Claim Payment / ERA", "Inbound", "Electronic Remittance Advice — payer sends payment explanation"],
      ["837P/I/D", "Healthcare Claim — Professional/Institutional/Dental", "Outbound", "Medical claim submission to payer or clearinghouse"],
      ["846", "Inventory Inquiry / Advice", "Both", "Stock levels shared between trading partners"],
      ["850", "Purchase Order", "Inbound", "Retailer sends PO to vendor. Your most common inbound transaction"],
      ["855", "Purchase Order Acknowledgment", "Outbound", "You confirm receipt and acceptance/rejection of PO"],
      ["856", "Ship Notice / Manifest (ASN)", "Outbound", "You notify buyer of shipment before it arrives"],
      ["860", "Purchase Order Change Request", "Inbound", "Buyer modifies an existing PO"],
      ["865", "Purchase Order Change Acknowledgment", "Outbound", "You confirm receipt of PO change"],
      ["940", "Warehouse Shipping Order", "Outbound", "Instructions to 3PL warehouse to ship specific orders"],
      ["943", "Warehouse Stock Transfer Shipping Advice", "Inbound", "3PL confirms what was shipped"],
      ["944", "Warehouse Stock Transfer Receipt Advice", "Inbound", "3PL confirms what was received at warehouse"],
      ["945", "Warehouse Shipping Advice", "Inbound", "3PL confirms warehouse fulfillment"],
      ["997", "Functional Acknowledgment (pre-5010)", "Both", "EDI content validation response (accepted or rejected)"],
      ["999", "Implementation Acknowledgment (5010+)", "Both", "HIPAA 5010 replacement for 997; more detailed error reporting"],
    ],
    [1400, 2800, 1600, 3560]
  ),
  blank(),

  h2("EDI X12 Segment Quick Reference"),
  makeTable(
    ["Segment", "Transaction", "Purpose", "Key Elements"],
    [
      ["ISA/IEA", "All", "Interchange envelope", "ISA05/06=Sender, ISA07/08=Receiver, ISA12=Version, ISA15=P/T"],
      ["GS/GE", "All", "Functional group", "GS01=Type (PO/IN/FA), GS06=Group Ctrl Num"],
      ["ST/SE", "All", "Transaction set", "ST01=TxnSet#, ST02=TxnCtrlNum; SE01=SegmentCount"],
      ["BEG", "850", "PO header", "BEG02=PO type, BEG03=PO#, BEG05=PO Date"],
      ["BIG", "810", "Invoice header", "BIG01=InvDate, BIG02=InvNum, BIG04=POReference"],
      ["BSN", "856", "ASN header", "BSN02=ShipmentID, BSN03=Date, BSN04=Time"],
      ["N1", "All", "Party name", "N101=Qualifier(ST/BT/SF), N102=Name, N103/04=ID type/value"],
      ["PO1", "850", "Line item", "PO101=Line#, PO102=Qty, PO103=UOM, PO104=Price, PO105=PriceBasis"],
      ["IT1", "810", "Invoice line", "IT101=Line#, IT102=Qty, IT103=UOM, IT104=Price"],
      ["HL", "856", "Hierarchy level", "HL01=Level#, HL02=Parent#, HL03=S/O/P/I"],
      ["SN1", "856", "Ship line", "SN102=Qty Shipped, SN103=UOM"],
      ["LIN", "856", "Item ID", "LIN01=Line#, LIN02=Qualifier(UP/IN), LIN03=Value"],
      ["CTT", "850", "Transaction totals", "CTT01=LineCount, CTT02=HashTotal"],
      ["AK1", "997", "Group ack", "AK101=FuncID, AK102=GrpCtrlNum, AK103=Version"],
      ["AK3", "997", "Segment error", "AK301=SegID, AK302=LineNum, AK304=ErrorCode"],
      ["AK4", "997", "Element error", "AK401=ElemPos, AK403=ErrorCode, AK404=BadValue"],
      ["AK5", "997", "Txn disposition", "AK501=A/E/R/M/W/X, AK502=ErrorCode"],
    ],
    [1000, 800, 2200, 5360]
  ),
  blank(),

  h2("AK4 Error Codes — Instant Reference"),
  makeTable(
    ["Code", "Meaning", "Most Common Cause"],
    [
      ["1", "Mandatory element missing — required field left blank", "Missing map rule; conditional suppression removing a mandatory field"],
      ["2", "Conditional required element missing — dependency rule violated", "N102 present but N103 absent when N101 qualifier requires it"],
      ["3", "Too many data elements — extra delimiters in segment", "Extended rule adding extra asterisks or bad delimiter concatenation"],
      ["4", "Data element too short — value below minimum length", "Left-padding stripped; zero-padded field not padded correctly"],
      ["5", "Data element too long — value exceeds maximum length", "String concatenation result exceeds element max; ERP field too long"],
      ["6", "Invalid character — control character or non-printable ASCII", "Carriage return/line feed inside element value; special chars in product description"],
      ["7", "Invalid code value — code not in valid list", "UOM code not in companion guide list; qualifier not recognized"],
      ["8", "Invalid date — date format wrong or non-existent date", "CCYYMMDD where YYMMDD expected; dateconvert format string wrong"],
      ["9", "Invalid time — time format incorrect", "HHMMSS expected but HHMM sent; wrong time zone format"],
      ["10", "Exclusion condition violated — mutually exclusive elements both present", "Two elements that cannot coexist are both populated"],
    ],
    [600, 2800, 6160]
  ),
  blank(),

  h2("Extended Rule Function Library"),
  makeTable(
    ["Function", "Syntax", "Description", "Example"],
    [
      ["concat", "concat(s1, s2, ...)", "Concatenate strings", "concat(first, ' ', last) → 'John Smith'"],
      ["strsub", "strsub(s, start, len)", "Substring (1-based start)", "strsub('ACME001', 1, 4) → 'ACME'"],
      ["strlen", "strlen(s)", "String length", "strlen('EA') → 2"],
      ["ltrim/rtrim", "ltrim(rtrim(s))", "Trim leading/trailing spaces", "ltrim(rtrim(' EA ')) → 'EA'"],
      ["strreplace", "strreplace(s, old, new)", "Replace all occurrences", "strreplace('EA/EA', '/', '-') → 'EA-EA'"],
      ["atof", "atof(s)", "String to decimal", "atof('12.50') → 12.5 (for math)"],
      ["atoi", "atoi(s)", "String to integer", "atoi('42') → 42"],
      ["dateconvert", "dateconvert(s, in_fmt, out_fmt)", "Convert date format", "dateconvert(d, 'CCYYMMDD', 'YYYY-MM-DD')"],
      ["loop_index", "loop_index('SegName')", "Current loop iteration", "loop_index('PO1') → 1, 2, 3..."],
      ["xref_lookup", "xref_lookup('TABLE', key, result)", "Cross-reference table lookup", "xref_lookup('UOM_XREF', 'EA', out_uom)"],
      ["if/else/endif", "if cond then ... else ... endif;", "Conditional logic", "if $N101 = 'ST' then ... endif;"],
    ],
    [1600, 2800, 2600, 2360]
  ),
  blank(),

  h2("Sterling Dashboard Quick Navigation"),
  makeTable(
    ["Task", "Navigate To"],
    [
      ["Monitor BP failures", "Operations > Business Processes > Status=Halted"],
      ["Pull raw EDI document", "Operations > Documents > search by partner/date"],
      ["Check 997 receipt", "Operations > Reports > EDI Correlation"],
      ["Deploy a new map", "Deployment > Maps > New Map"],
      ["Deploy a new BP", "Deployment > Business Processes > Create"],
      ["Create trading partner", "Administration > Trading Partners > New"],
      ["Configure AS2 profile", "Administration > Trading Partners > AS2 > New"],
      ["Import certificate", "Administration > Certificates > Import"],
      ["Configure SFTP adapter", "Administration > Adapter > SFTPClientAdapter"],
      ["Add SSH key", "Administration > Trading Partners > SSH Private Keys"],
      ["Create mailbox", "Mailbox > Manage Mailboxes > New"],
      ["Create routing rule", "Mailbox > Routing Rules > New"],
      ["Schedule a BP", "Deployment > Schedules > New Schedule"],
      ["Create cross-reference", "Administration > Trading Partner > Code Lists > New"],
      ["Check system performance", "Administration > System > Performance Tuning"],
      ["View adapter logs", "SSH: tail -f /opt/IBM/SterlingIntegrator/logs/adapter.log"],
    ],
    [3800, 5560]
  ),
  blank(),

  // CLOSING STATEMENT
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: allBorders("FFD700", 10),
        shading: { fill: "1F3864", type: ShadingType.CLEAR },
        margins: { top: 360, bottom: 360, left: 360, right: 360 },
        width: { size: CONTENT_W, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 }, children: [
            new TextRun({ text: "🏆  YOU ARE PRODUCTION-READY  🏆", font: "Arial", size: 36, bold: true, color: "FFD700" })
          ]}),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 }, children: [
            new TextRun({ text: "Master the labs. Practice the scenarios. Own the troubleshooting checklist.", font: "Arial", size: 24, color: "BDD7EE", italics: true })
          ]}),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 }, children: [
            new TextRun({ text: "The difference between a junior and a senior EDI BA is the confidence to diagnose any situation — without panic.", font: "Arial", size: 22, color: WHITE })
          ]}),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 0 }, children: [
            new TextRun({ text: "IBM Sterling B2B Integrator — Complete Training Guide  ·  2026 Edition  ·  Enhanced for Job Readiness", font: "Arial", size: 18, color: BLUE_LIGHT, italics: true })
          ]}),
        ]
      })
    ]})]
  })
);

// ═══════════════════════════════════════════════════════════════════════════
// BUILD DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }, {
          level: 1, format: LevelFormat.BULLET, text: "◦",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: WHITE },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: BLUE_MID },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
      { id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 3 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children
  }]
});

// Write to current working directory
const outputPath = path.join(process.cwd(), 'EDI_BA_JobReady_Guide_2026.docx');
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outputPath, buf);
  console.log(`✅ Document successfully written to: ${outputPath}`);
}).catch(err => {
  console.error('❌ Error generating document:', err);
  process.exit(1);
});
