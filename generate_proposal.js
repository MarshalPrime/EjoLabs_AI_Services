const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageBreak
} = require("docx");
const fs = require("fs");
const path = require("path");

const ASSETS  = path.join(__dirname, "assets");
const OUTPUT  = path.join(__dirname, "EjoLabs_AI_Services.docx");

// ── Brand colours ────────────────────────────────────────────────────────
const ORANGE      = "E87722";
const NAVY_DARK   = "1A2535";
const ORANGE_SOFT = "F9A85D";
const WHITE       = "FFFFFF";
const GREY        = "6B7280";
const LIGHT_CARD  = "FEF9F4";
const BODY_TEXT   = "2D3748";
const LIGHT_GREY  = "E5E7EB";
const ACCENT_BLUE = "94A3B8";

// ── Load assets ──────────────────────────────────────────────────────────
const logo = fs.readFileSync(path.join(ASSETS, "ejolabs_logo.png"));
const img1 = fs.readFileSync(path.join(ASSETS, "svc01_ejochat.png"));
const img2 = fs.readFileSync(path.join(ASSETS, "svc02_agentic_rag.png"));
const img3 = fs.readFileSync(path.join(ASSETS, "svc03_consultancy.png"));

// ── Helpers ──────────────────────────────────────────────────────────────
const NB = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB, insideH: NB, insideV: NB };

function t(text, opts = {}) {
  return new TextRun({
    text, font: "Arial", size: opts.size || 20,
    bold: opts.bold, italics: opts.italics,
    color: opts.color || BODY_TEXT, break: opts.break,
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before ?? 40, after: opts.after ?? 40 },
    indent: opts.indent,
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    border: opts.borderBottom ? {
      bottom: { style: BorderStyle.SINGLE, size: opts.borderBottom, color: ORANGE, space: 4 }
    } : undefined,
    children: Array.isArray(children) ? children : [children],
  });
}

function spacer(pts = 100) {
  return new Paragraph({ spacing: { before: 0, after: pts }, children: [t("")] });
}

function bullet(text, dark = false) {
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    indent: { left: 360, hanging: 260 },
    children: [
      t("▸  ", { bold: true, color: dark ? ORANGE_SOFT : ORANGE, size: 20 }),
      t(text, { color: dark ? "CBD5E0" : BODY_TEXT, size: 19 }),
    ],
  });
}

function tagLine(arr, dark = false) {
  return new Paragraph({
    spacing: { before: 80, after: 20 },
    children: [t(arr.join("   ·   "), { size: 16, bold: true, color: dark ? ORANGE_SOFT : ORANGE })],
  });
}

// ── Header bar for inner pages ───────────────────────────────────────────
const docHeader = new Header({
  children: [
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [700, 5300, 3360],
      borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 6, color: LIGHT_GREY } },
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 700, type: WidthType.DXA },
          margins: { top: 40, bottom: 60, left: 0, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
          children: [para([new ImageRun({ data: logo, transformation: { width: 40, height: 40 }, type: "png" })])],
        }),
        new TableCell({
          width: { size: 5300, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 40, bottom: 60, left: 0, right: 0 },
          children: [para([t("Ejo Labs", { size: 20, bold: true, color: NAVY_DARK })])],
        }),
        new TableCell({
          width: { size: 3360, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          children: [para([t("AI Services for Rwanda", { size: 16, italics: true, color: GREY })], { align: AlignmentType.RIGHT })],
        }),
      ]})],
    }),
  ],
});

// ── Footer bar ───────────────────────────────────────────────────────────
const docFooter = new Footer({
  children: [
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      borders: { ...noBorders, top: { style: BorderStyle.SINGLE, size: 8, color: ORANGE } },
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 9360, type: WidthType.DXA },
          margins: { top: 80, bottom: 0, left: 0, right: 0 },
          children: [para([t("info@ejolabs.com   ·   ejolabs.com   ·   Kigali, Rwanda", { size: 16, color: GREY })],
            { align: AlignmentType.CENTER })],
        }),
      ]})],
    }),
  ],
});

// =====================================================================
//  PAGE 1 — COVER
// =====================================================================
const coverPage = [
  // Logo — large, centred
  para([new ImageRun({ data: logo, transformation: { width: 150, height: 150 }, type: "png" })],
    { align: AlignmentType.CENTER, before: 1200, after: 200 }),

  para([t("EJO LABS", { size: 52, bold: true, color: NAVY_DARK })],
    { align: AlignmentType.CENTER, before: 0, after: 40 }),

  para([t("Practical AI for Rwanda & Africa", { size: 24, italics: true, color: ORANGE })],
    { align: AlignmentType.CENTER, before: 0, after: 300, borderBottom: 16 }),

  para([t("AI Services That Move", { size: 56, bold: true, color: NAVY_DARK })],
    { align: AlignmentType.CENTER, before: 200, after: 40 }),
  para([t("Rwanda Forward.", { size: 56, bold: true, color: ORANGE })],
    { align: AlignmentType.CENTER, before: 0, after: 300 }),

  para([t("A capabilities overview for partners & clients", { size: 20, italics: true, color: GREY })],
    { align: AlignmentType.CENTER, before: 0, after: 400 }),

  // Three pillars
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    borders: { ...noBorders, insideV: { style: BorderStyle.SINGLE, size: 4, color: "3A4D66" } },
    rows: [new TableRow({ children:
      [
        ["Kinyarwanda-First AI", "Built for Rwanda,\nnot translated for it"],
        ["Production-Ready", "Deployed systems,\nnot prototypes"],
        ["Agentic & Intelligent", "AI that reasons,\nacts, and reports"],
      ].map(([title, desc]) => new TableCell({
        width: { size: 3120, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: NAVY_DARK },
        margins: { top: 140, bottom: 140, left: 180, right: 180 },
        children: [
          para([t(title, { size: 18, bold: true, color: ORANGE_SOFT })], { align: AlignmentType.CENTER, before: 0, after: 40 }),
          para([t(desc, { size: 16, color: ACCENT_BLUE })], { align: AlignmentType.CENTER, before: 0, after: 0 }),
        ],
      }))
    })],
  }),
];

// =====================================================================
//  SERVICE CARD — redesigned: full-width image ABOVE text
// =====================================================================
function serviceCard(num, title, subtitle, desc, bullets, tags, imgData) {
  return [
    // ── Number badge + title bar ──
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [560, 8800],
      borders: noBorders,
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 560, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: ORANGE },
          margins: { top: 60, bottom: 60, left: 0, right: 0 },
          verticalAlign: VerticalAlign.CENTER,
          children: [para([t(num, { size: 20, bold: true, color: WHITE })], { align: AlignmentType.CENTER })],
        }),
        new TableCell({
          width: { size: 8800, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: NAVY_DARK },
          margins: { top: 60, bottom: 60, left: 200, right: 100 },
          verticalAlign: VerticalAlign.CENTER,
          children: [para([t(title, { size: 24, bold: true, color: WHITE })])],
        }),
      ]})],
    }),

    // ── Full-width image ──
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      borders: {
        top: NB, bottom: NB,
        left: { style: BorderStyle.SINGLE, size: 1, color: LIGHT_GREY },
        right: { style: BorderStyle.SINGLE, size: 1, color: LIGHT_GREY },
        insideH: NB, insideV: NB,
      },
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 9360, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [para([new ImageRun({
            data: imgData,
            transformation: { width: 620, height: 340 },
            type: "png",
          })], { align: AlignmentType.CENTER, before: 0, after: 0 })],
        }),
      ]})],
    }),

    // ── Text content card ──
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      borders: {
        top: NB,
        bottom: { style: BorderStyle.SINGLE, size: 2, color: LIGHT_GREY },
        left: { style: BorderStyle.SINGLE, size: 1, color: LIGHT_GREY },
        right: { style: BorderStyle.SINGLE, size: 1, color: LIGHT_GREY },
        insideH: NB, insideV: NB,
      },
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 9360, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: LIGHT_CARD },
          margins: { top: 140, bottom: 140, left: 240, right: 240 },
          children: [
            para([t(subtitle, { size: 20, italics: true, color: ORANGE })], { before: 0, after: 100 }),
            para([t(desc, { size: 19, color: BODY_TEXT })], { align: AlignmentType.JUSTIFIED, before: 0, after: 100 }),
            ...bullets.map(b => bullet(b)),
            spacer(40),
            tagLine(tags),
          ],
        }),
      ]})],
    }),

    spacer(160),
  ];
}

// =====================================================================
//  PAGE 2 — Who We Are + Services 01 & 02
// =====================================================================
const page2 = [
  para([t("WHO WE ARE", { size: 18, bold: true, color: ORANGE })],
    { before: 0, after: 80, borderBottom: 12 }),
  spacer(40),

  para([t(
    "Ejo Labs builds AI systems that speak Kinyarwanda and understand Rwanda. We combine large language models, computer vision, and agentic pipelines with deep local knowledge to deliver production-ready solutions for Rwandan organisations. We build AI for low-resource languages like Kinyarwanda — where big tech isn't looking.",
    { size: 20, color: BODY_TEXT }
  )], { align: AlignmentType.JUSTIFIED, before: 0, after: 80 }),

  // Three mission pillars
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    borders: {
      ...noBorders,
      top: { style: BorderStyle.SINGLE, size: 10, color: ORANGE },
      insideV: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_GREY },
    },
    rows: [new TableRow({ children:
      [
        ["Kinyarwanda-First AI", "Built for Rwandan speech,\nnot translated from English"],
        ["Rwanda-Grounded Context", "Local language, use-cases,\nand regulatory awareness"],
        ["AWS Cloud Infrastructure", "Scalable, secure,\nand deployable in-country"],
      ].map(([title, desc]) => new TableCell({
        width: { size: 3120, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: LIGHT_CARD },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        children: [
          para([t(title, { size: 18, bold: true, color: NAVY_DARK })], { before: 0, after: 40 }),
          para([t(desc, { size: 16, color: GREY })], { before: 0, after: 0 }),
        ],
      }))
    })],
  }),

  spacer(180),

  para([t("OUR SERVICES", { size: 18, bold: true, color: ORANGE })],
    { before: 0, after: 80, borderBottom: 12 }),
  spacer(60),

  ...serviceCard(
    "01", "EjoChat — Kinyarwanda AI Assistant",
    "Rwanda's first production-ready Kinyarwanda conversational AI",
    "EjoChat is fine-tuned on curated Rwandan language data to understand natural Kinyarwanda — informal speech, idioms, and local context. Integrate via REST API into citizen portals, mobile apps, USSD, or WhatsApp.",
    [
      "Kinyarwanda-first, multilingual fallback, domain-tunable",
      "REST API for web, mobile, USSD & WhatsApp",
      "Ideal for: citizen services, NGO outreach, enterprise helpdesks",
    ],
    [],
    img1,
  ),

  ...serviceCard(
    "02", "Agentic AI, RAG & Computer Vision",
    "Intelligent agents that retrieve, reason, act — and see",
    "RAG agents ingest your documents and data feeds, reason over them, and produce structured outputs. For security, this extends to AI-powered video surveillance with real-time object detection and anomaly alerts — powered by Qdrant vector search.",
    [
      "Document RAG: PDFs, databases → cited, structured answers",
      "Video surveillance: object detection, person ID, anomaly alerts",
      "Image analysis: scene classification, evidence documentation",
      "On-premise or AWS (Bedrock, SageMaker) deployment",
    ],
    [],
    img2,
  ),

  new Paragraph({ children: [new PageBreak()] }),
];

// =====================================================================
//  PAGE 3 — Service 03 + Discovery + CTA
// =====================================================================
const page3 = [
  para([t("OUR SERVICES", { size: 18, bold: true, color: ORANGE })],
    { before: 0, after: 80, borderBottom: 12 }),
  spacer(60),

  ...serviceCard(
    "03", "AI Consultancy — We Solve Your Problem",
    "Understanding your world before we build for it",
    "Before any technology, we sit with your team to understand what you do, where friction is, and where AI can create the most value. Then we design and build the right solution — whether that's a custom model, a data pipeline, an agentic system, or something entirely new.",
    [
      "Workshop: map your operations, data, and pain points",
      "Use-case identification: highest-impact, lowest-risk opportunities",
      "Bespoke AI solution design and delivery",
      "Works across: security, agriculture, health, finance, government",
    ],
    [],
    img3,
  ),

  // ── Discovery Session dark block ──
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 14, color: ORANGE },
      bottom: NB, left: NB, right: NB, insideH: NB, insideV: NB,
    },
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 9360, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: NAVY_DARK },
        margins: { top: 220, bottom: 220, left: 300, right: 300 },
        children: [
          para([t("How a Discovery Session Works", { size: 24, bold: true, color: ORANGE_SOFT })],
            { before: 0, after: 100 }),
          para([t(
            "We meet your team, map your workflows, audit your data, and identify where AI can act now. No jargon, completely exploratory — just a clear picture of what's possible and a concrete proposal if you choose to proceed.",
            { size: 19, color: "CBD5E0" }
          )], { align: AlignmentType.JUSTIFIED, before: 0, after: 120 }),
          bullet("Workflow mapping: where time is lost, errors occur, or data sits unused", true),
          bullet("Data audit: what you already have that AI can act on immediately", true),
          bullet("Use-case prioritisation: highest-impact, lowest-risk starting points", true),
          bullet("Feasibility & cost estimate: realistic timelines and investment", true),
          spacer(60),
          tagLine(["BESPOKE", "COLLABORATIVE", "EXPLORATORY"], true),
        ],
      }),
    ]})],
  }),

  spacer(180),

  // ── CTA bar ──
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [5400, 3960],
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 5400, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: ORANGE },
        margins: { top: 180, bottom: 180, left: 260, right: 140 },
        verticalAlign: VerticalAlign.CENTER,
        children: [para([t("Let's build AI that works for you.", { size: 24, bold: true, color: WHITE })])],
      }),
      new TableCell({
        width: { size: 3960, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: NAVY_DARK },
        margins: { top: 180, bottom: 180, left: 260, right: 260 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          para([t("Book a Discovery Session", { size: 22, bold: true, color: ORANGE_SOFT })], { before: 0, after: 50 }),
          para([t("info@ejolabs.com  ·  ejolabs.com", { size: 18, color: ACCENT_BLUE })], { before: 0, after: 0 }),
        ],
      }),
    ]})],
  }),
];

// =====================================================================
//  Build the Document
// =====================================================================
const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 20, color: BODY_TEXT } } } },
  sections: [
    // Section 1: Cover page (no header/footer)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 720, right: 1080, bottom: 720, left: 1080 },
        },
      },
      children: coverPage,
    },
    // Section 2: Content pages (header + footer)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 960, right: 1080, bottom: 900, left: 1080 },
        },
      },
      headers: { default: docHeader },
      footers: { default: docFooter },
      children: [...page2, ...page3],
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUTPUT, buf);
  console.log("✅  Done:", OUTPUT);
  console.log("   Size:", (buf.length / 1024 / 1024).toFixed(1), "MB");
}).catch(err => {
  console.error("❌  Error:", err);
});
