/**
 * AI Takeoff — Proposal Generator
 * 
 * Takes BOQ output (from rules engine or canvas export) and creates
 * a professional .docx proposal with:
 *   - Cover sheet
 *   - Executive summary
 *   - Equipment tables by zone
 *   - Infrastructure summary
 *   - Labor breakdown
 *   - Pricing summary
 *   - Terms & conditions
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, Tab, TabStopType, TabStopPosition
} = require('docx');

// ─── CONSTANTS ───
const PAGE_W = 12240; // US Letter
const PAGE_H = 15840;
const MARGIN = 1440;  // 1 inch
const CONTENT_W = PAGE_W - 2 * MARGIN; // 9360

const COLOR = {
  brand: '1a365d',       // deep navy
  accent: '2b6cb0',      // medium blue
  light: 'ebf4ff',       // light blue bg
  headerBg: '1a365d',
  headerText: 'ffffff',
  altRow: 'f7fafc',
  border: 'cbd5e0',
  text: '2d3748',
  muted: '718096',
  green: '276749',
  money: '1a365d'
};

const THIN_BORDER = { style: BorderStyle.SINGLE, size: 1, color: COLOR.border };
const BORDERS = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'ffffff' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };
const CELL_MARGINS = { top: 60, bottom: 60, left: 100, right: 100 };

// ─── HELPERS ───
function $(amount) {
  return '$' + Number(amount).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function txt(text, opts = {}) {
  return new TextRun({ text, font: 'Arial', size: opts.size || 20, color: opts.color || COLOR.text, bold: opts.bold || false, italics: opts.italics || false });
}

function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 120 },
    ...(opts.heading ? { heading: opts.heading } : {}),
    ...(opts.indent ? { indent: opts.indent } : {}),
    ...(opts.numbering ? { numbering: opts.numbering } : {})
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders: BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: COLOR.headerBg, type: ShadingType.CLEAR },
    margins: CELL_MARGINS,
    children: [para([txt(text, { color: COLOR.headerText, bold: true, size: 18 })], { after: 0 })]
  });
}

function cell(text, width, opts = {}) {
  return new TableCell({
    borders: BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shaded ? { fill: COLOR.altRow, type: ShadingType.CLEAR } : undefined,
    margins: CELL_MARGINS,
    children: [para([txt(text, { size: 18, bold: opts.bold, color: opts.color })], { align: opts.align || AlignmentType.LEFT, after: 0 })]
  });
}

function moneyCell(amount, width, opts = {}) {
  return cell($(amount), width, { align: AlignmentType.RIGHT, color: COLOR.money, ...opts });
}

// ─── GENERATE PROPOSAL ───
async function generateProposal(boqData, config = {}) {
  const {
    companyName = 'RTS Technology Solutions',
    companyTagline = 'Audio Visual | Networking | Security',
    companyAddress = '',
    companyPhone = '',
    companyEmail = '',
    clientName = 'Valued Client',
    clientCompany = '',
    clientAddress = '',
    proposalNumber = `RTS-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 9000) + 1000)}`,
    validDays = 30,
    paymentTerms = 'Net 30',
    warrantyYears = 1
  } = config;

  const projectName = boqData.project || 'Technology Project';
  const date = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const validUntil = new Date(Date.now() + validDays * 86400000).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

  // Flatten all zone items for totals
  const allZones = boqData.zones || [];
  const infrastructure = boqData.infrastructure;
  const summary = boqData.summary || {};

  // ─── COVER PAGE ───
  const coverSection = {
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    children: [
      // Spacer
      para([txt(' ')], { before: 2400 }),
      // Company name
      para([txt(companyName, { size: 56, bold: true, color: COLOR.brand })], { align: AlignmentType.CENTER, after: 60 }),
      para([txt(companyTagline, { size: 22, color: COLOR.muted, italics: true })], { align: AlignmentType.CENTER, after: 600 }),
      // Divider line
      new Table({
        width: { size: 4000, type: WidthType.DXA },
        columnWidths: [4000],
        rows: [new TableRow({
          children: [new TableCell({
            borders: { top: NO_BORDER, bottom: { style: BorderStyle.SINGLE, size: 3, color: COLOR.accent }, left: NO_BORDER, right: NO_BORDER },
            width: { size: 4000, type: WidthType.DXA },
            children: [para([txt(' ', { size: 4 })], { after: 0 })]
          })]
        })]
      }),
      para([txt(' ')], { before: 400 }),
      // Project title
      para([txt('Technology Proposal', { size: 36, color: COLOR.accent })], { align: AlignmentType.CENTER, after: 80 }),
      para([txt(projectName, { size: 28, bold: true, color: COLOR.brand })], { align: AlignmentType.CENTER, after: 600 }),
      // Details grid
      para([txt(`Proposal #: ${proposalNumber}`, { size: 20, color: COLOR.muted })], { align: AlignmentType.CENTER, after: 60 }),
      para([txt(`Date: ${date}`, { size: 20, color: COLOR.muted })], { align: AlignmentType.CENTER, after: 60 }),
      para([txt(`Valid Until: ${validUntil}`, { size: 20, color: COLOR.muted })], { align: AlignmentType.CENTER, after: 200 }),
      // Prepared for
      para([txt('Prepared For', { size: 18, color: COLOR.muted, italics: true })], { align: AlignmentType.CENTER, after: 60 }),
      para([txt(clientCompany || clientName, { size: 24, bold: true, color: COLOR.brand })], { align: AlignmentType.CENTER, after: 40 }),
      ...(clientName && clientCompany ? [para([txt(clientName, { size: 20, color: COLOR.text })], { align: AlignmentType.CENTER, after: 40 })] : []),
      ...(clientAddress ? [para([txt(clientAddress, { size: 18, color: COLOR.muted })], { align: AlignmentType.CENTER })] : []),
    ]
  };

  // ─── CONTENT SECTIONS ───
  const contentChildren = [];

  // --- EXECUTIVE SUMMARY ---
  contentChildren.push(
    para([txt('Executive Summary')], { heading: HeadingLevel.HEADING_1, before: 0, after: 200 }),
    para([txt(`${companyName} is pleased to present this technology proposal for `, { size: 20 }),
      txt(projectName, { size: 20, bold: true }),
      txt(`. This proposal covers the complete design, supply, and installation of audio/visual, networking, and security systems across ${summary.zoneCount || allZones.length} zones totaling approximately ${(summary.totalSqft || 0).toLocaleString()} square feet.`, { size: 20 })
    ], { after: 160 }),
    para([txt('The proposed solution includes:', { size: 20 })], { after: 100 })
  );

  // Count totals by category
  const catCounts = { AV: 0, NET: 0, AC: 0 };
  allZones.forEach(z => {
    (z.items || []).forEach(item => {
      if (item.category) catCounts[item.category] = (catCounts[item.category] || 0) + item.qty;
    });
  });
  if (infrastructure && infrastructure.items) {
    infrastructure.items.forEach(item => {
      if (item.category) catCounts[item.category] = (catCounts[item.category] || 0) + item.qty;
    });
  }

  const bulletItems = [];
  if (catCounts.AV > 0) bulletItems.push(`Audio/Visual systems — ${catCounts.AV} devices including speakers, displays, touch panels, and control processors`);
  if (catCounts.NET > 0) bulletItems.push(`Network infrastructure — ${catCounts.NET} devices including wireless access points, switches, and security appliances`);
  if (catCounts.AC > 0) bulletItems.push(`Access control and surveillance — ${catCounts.AC} devices including cameras, card readers, and recording systems`);

  bulletItems.forEach(item => {
    contentChildren.push(
      para([txt(item, { size: 20 })], { numbering: { reference: 'bullets', level: 0 }, after: 60 })
    );
  });

  contentChildren.push(
    para([txt(' ')], { after: 80 }),
    // Investment summary box
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W / 2, CONTENT_W / 2],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 2, color: COLOR.accent }, bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR.accent }, left: { style: BorderStyle.SINGLE, size: 2, color: COLOR.accent }, right: NO_BORDER },
              width: { size: CONTENT_W / 2, type: WidthType.DXA },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 200, right: 100 },
              children: [para([txt('Total Investment', { size: 22, bold: true, color: COLOR.accent })], { after: 0 })]
            }),
            new TableCell({
              borders: { top: { style: BorderStyle.SINGLE, size: 2, color: COLOR.accent }, bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR.accent }, left: NO_BORDER, right: { style: BorderStyle.SINGLE, size: 2, color: COLOR.accent } },
              width: { size: CONTENT_W / 2, type: WidthType.DXA },
              shading: { fill: COLOR.light, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 100, right: 200 },
              children: [para([txt($(summary.projectTotal || 0), { size: 28, bold: true, color: COLOR.brand })], { align: AlignmentType.RIGHT, after: 0 })]
            })
          ]
        })
      ]
    }),
    para([txt(' ')], { after: 100 })
  );

  // --- EQUIPMENT BY ZONE ---
  contentChildren.push(
    new Paragraph({ children: [new PageBreak()] }),
    para([txt('Equipment Schedule')], { heading: HeadingLevel.HEADING_1, before: 0, after: 200 })
  );

  // Column widths: Qty(600) | Item(3200) | Part#(1800) | Unit(1100) | Ext(1100) | Labor(1100)
  // Wait — that's only 8900. Let me fix: need to sum to 9360
  const EQ_COLS = [600, 3060, 1800, 1300, 1300, 1300]; // = 9360

  allZones.forEach(zone => {
    if (!zone.items || zone.items.length === 0) return;

    contentChildren.push(
      para([txt(`${zone.zone}`, { size: 22, bold: true, color: COLOR.brand })], { before: 240, after: 100 }),
      para([txt(`${zone.label || zone.type} — ${(zone.sqft || 0).toLocaleString()} sq ft`, { size: 18, color: COLOR.muted, italics: true })], { after: 120 })
    );

    const rows = [
      new TableRow({
        children: [
          headerCell('Qty', EQ_COLS[0]),
          headerCell('Equipment', EQ_COLS[1]),
          headerCell('Part Number', EQ_COLS[2]),
          headerCell('Unit Cost', EQ_COLS[3]),
          headerCell('Extended', EQ_COLS[4]),
          headerCell('Labor', EQ_COLS[5])
        ]
      })
    ];

    zone.items.forEach((item, i) => {
      const shaded = i % 2 === 1;
      rows.push(new TableRow({
        children: [
          cell(String(item.qty), EQ_COLS[0], { align: AlignmentType.CENTER, shaded }),
          cell(item.name, EQ_COLS[1], { shaded }),
          cell(item.partNumber, EQ_COLS[2], { shaded, color: COLOR.muted }),
          moneyCell(item.unitCost, EQ_COLS[3], { shaded }),
          moneyCell(item.equipmentCost, EQ_COLS[4], { shaded }),
          moneyCell(item.laborCost, EQ_COLS[5], { shaded })
        ]
      }));
    });

    // Zone subtotal row
    rows.push(new TableRow({
      children: [
        new TableCell({ borders: BORDERS, width: { size: EQ_COLS[0] + EQ_COLS[1] + EQ_COLS[2] + EQ_COLS[3], type: WidthType.DXA }, columnSpan: 4, margins: CELL_MARGINS,
          shading: { fill: COLOR.light, type: ShadingType.CLEAR },
          children: [para([txt('Zone Subtotal', { size: 18, bold: true, color: COLOR.brand })], { align: AlignmentType.RIGHT, after: 0 })] }),
        cell($(zone.totals.equipment), EQ_COLS[4], { bold: true, align: AlignmentType.RIGHT, color: COLOR.brand }),
        cell($(zone.totals.labor), EQ_COLS[5], { bold: true, align: AlignmentType.RIGHT, color: COLOR.brand })
      ]
    }));

    contentChildren.push(
      new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: EQ_COLS, rows })
    );
  });

  // --- INFRASTRUCTURE ---
  if (infrastructure && infrastructure.items && infrastructure.items.length > 0) {
    contentChildren.push(
      new Paragraph({ children: [new PageBreak()] }),
      para([txt('Infrastructure Equipment')], { heading: HeadingLevel.HEADING_1, before: 0, after: 120 }),
      para([txt('The following backbone equipment is automatically sized based on the total device counts across all zones.', { size: 20, color: COLOR.muted })], { after: 160 })
    );

    // Counts summary
    const c = infrastructure.counts;
    if (c) {
      contentChildren.push(
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA },
          columnWidths: [CONTENT_W / 3, CONTENT_W / 3, CONTENT_W / 3],
          rows: [
            new TableRow({ children: [
              new TableCell({ borders: NO_BORDERS, width: { size: CONTENT_W / 3, type: WidthType.DXA }, margins: CELL_MARGINS,
                children: [
                  para([txt(`${c.totalSpeakers}`, { size: 28, bold: true, color: COLOR.brand }), txt(' speakers', { size: 18, color: COLOR.muted })], { after: 0 }),
                  para([txt(`${c.totalDisplays}`, { size: 28, bold: true, color: COLOR.brand }), txt(' displays', { size: 18, color: COLOR.muted })], { after: 0 })
                ] }),
              new TableCell({ borders: NO_BORDERS, width: { size: CONTENT_W / 3, type: WidthType.DXA }, margins: CELL_MARGINS,
                children: [
                  para([txt(`${c.totalCameras}`, { size: 28, bold: true, color: COLOR.brand }), txt(' cameras', { size: 18, color: COLOR.muted })], { after: 0 }),
                  para([txt(`${c.totalReaders}`, { size: 28, bold: true, color: COLOR.brand }), txt(' card readers', { size: 18, color: COLOR.muted })], { after: 0 })
                ] }),
              new TableCell({ borders: NO_BORDERS, width: { size: CONTENT_W / 3, type: WidthType.DXA }, margins: CELL_MARGINS,
                children: [
                  para([txt(`${c.totalPoeDevices}`, { size: 28, bold: true, color: COLOR.brand }), txt(' PoE devices', { size: 18, color: COLOR.muted })], { after: 0 }),
                  para([txt(`${c.totalZones}`, { size: 28, bold: true, color: COLOR.brand }), txt(' audio zones', { size: 18, color: COLOR.muted })], { after: 0 })
                ] })
            ] })
          ]
        }),
        para([txt(' ')], { after: 120 })
      );
    }

    const infraRows = [
      new TableRow({ children: [
        headerCell('Qty', EQ_COLS[0]), headerCell('Equipment', EQ_COLS[1]),
        headerCell('Part Number', EQ_COLS[2]), headerCell('Unit Cost', EQ_COLS[3]),
        headerCell('Extended', EQ_COLS[4]), headerCell('Labor', EQ_COLS[5])
      ] })
    ];

    infrastructure.items.forEach((item, i) => {
      const shaded = i % 2 === 1;
      infraRows.push(new TableRow({ children: [
        cell(String(item.qty), EQ_COLS[0], { align: AlignmentType.CENTER, shaded }),
        cell(item.name, EQ_COLS[1], { shaded }),
        cell(item.partNumber, EQ_COLS[2], { shaded, color: COLOR.muted }),
        moneyCell(item.unitCost, EQ_COLS[3], { shaded }),
        moneyCell(item.equipmentCost, EQ_COLS[4], { shaded }),
        moneyCell(item.laborCost, EQ_COLS[5], { shaded })
      ] }));
    });

    infraRows.push(new TableRow({ children: [
      new TableCell({ borders: BORDERS, width: { size: EQ_COLS[0]+EQ_COLS[1]+EQ_COLS[2]+EQ_COLS[3], type: WidthType.DXA }, columnSpan: 4, margins: CELL_MARGINS,
        shading: { fill: COLOR.light, type: ShadingType.CLEAR },
        children: [para([txt('Infrastructure Subtotal', { size: 18, bold: true, color: COLOR.brand })], { align: AlignmentType.RIGHT, after: 0 })] }),
      cell($(infrastructure.totals.equipment), EQ_COLS[4], { bold: true, align: AlignmentType.RIGHT, color: COLOR.brand }),
      cell($(infrastructure.totals.labor), EQ_COLS[5], { bold: true, align: AlignmentType.RIGHT, color: COLOR.brand })
    ] }));

    contentChildren.push(
      new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: EQ_COLS, rows: infraRows })
    );
  }

  // --- PRICING SUMMARY ---
  contentChildren.push(
    new Paragraph({ children: [new PageBreak()] }),
    para([txt('Investment Summary')], { heading: HeadingLevel.HEADING_1, before: 0, after: 240 })
  );

  const PRICE_COLS = [5460, 3900]; // = 9360
  const priceRows = [];

  const addPriceRow = (label, value, opts = {}) => {
    priceRows.push(new TableRow({ children: [
      new TableCell({
        borders: { top: NO_BORDER, bottom: opts.borderBottom ? { style: BorderStyle.SINGLE, size: 1, color: COLOR.border } : NO_BORDER, left: NO_BORDER, right: NO_BORDER },
        width: { size: PRICE_COLS[0], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [para([txt(label, { size: opts.size || 20, bold: opts.bold, color: opts.labelColor || COLOR.text })], { after: 0 })]
      }),
      new TableCell({
        borders: { top: NO_BORDER, bottom: opts.borderBottom ? { style: BorderStyle.SINGLE, size: 1, color: COLOR.border } : NO_BORDER, left: NO_BORDER, right: NO_BORDER },
        width: { size: PRICE_COLS[1], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [para([txt($(value), { size: opts.size || 20, bold: opts.bold, color: opts.valueColor || COLOR.money })], { align: AlignmentType.RIGHT, after: 0 })]
      })
    ] }));
  };

  addPriceRow('Equipment Subtotal', summary.equipmentSubtotal || 0, { borderBottom: true });
  addPriceRow(`Markup (${summary.markupPercent || 30}%)`, (summary.equipmentWithMarkup || 0) - (summary.equipmentSubtotal || 0), { borderBottom: true });
  addPriceRow('Equipment Total', summary.equipmentWithMarkup || 0, { bold: true, borderBottom: true });
  addPriceRow(`Labor (${(summary.totalLaborHours || 0).toLocaleString()} hours)`, summary.laborSubtotal || 0, { borderBottom: true });

  // Grand total
  priceRows.push(new TableRow({ children: [
    new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 2, color: COLOR.brand }, bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR.brand }, left: { style: BorderStyle.SINGLE, size: 2, color: COLOR.brand }, right: NO_BORDER },
      width: { size: PRICE_COLS[0], type: WidthType.DXA },
      shading: { fill: COLOR.headerBg, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 100 },
      children: [para([txt('PROJECT TOTAL', { size: 24, bold: true, color: COLOR.headerText })], { after: 0 })]
    }),
    new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 2, color: COLOR.brand }, bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR.brand }, left: NO_BORDER, right: { style: BorderStyle.SINGLE, size: 2, color: COLOR.brand } },
      width: { size: PRICE_COLS[1], type: WidthType.DXA },
      shading: { fill: COLOR.headerBg, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 100, right: 200 },
      children: [para([txt($(summary.projectTotal || 0), { size: 28, bold: true, color: COLOR.headerText })], { align: AlignmentType.RIGHT, after: 0 })]
    })
  ] }));

  contentChildren.push(
    new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: PRICE_COLS, rows: priceRows })
  );

  // --- TERMS & CONDITIONS ---
  contentChildren.push(
    new Paragraph({ children: [new PageBreak()] }),
    para([txt('Terms and Conditions')], { heading: HeadingLevel.HEADING_1, before: 0, after: 200 })
  );

  const terms = [
    { title: 'Proposal Validity', body: `This proposal is valid for ${validDays} days from the date of issue. Pricing is subject to change after the expiration date due to manufacturer price adjustments and material availability.` },
    { title: 'Payment Terms', body: `${paymentTerms}. A 50% deposit is required upon acceptance to initiate equipment procurement. Remaining balance is due upon substantial completion of installation. Progress billing may apply for projects exceeding 30 days.` },
    { title: 'Scope of Work', body: 'This proposal includes the supply, delivery, installation, programming, and commissioning of all equipment listed herein. Standard mounting hardware and cabling within 50 feet of each device is included. Structured cabling, conduit, electrical work, and building modifications are not included unless explicitly stated.' },
    { title: 'Warranty', body: `${companyName} provides a ${warrantyYears}-year labor warranty on all installed systems from the date of final acceptance. Manufacturer warranties apply to all equipment and are passed through to the client. Extended warranty and service agreements are available upon request.` },
    { title: 'Change Orders', body: 'Any changes to the scope of work after proposal acceptance will be documented via written change order. Additional equipment, labor, or modifications will be quoted separately and require written approval before proceeding.' },
    { title: 'Exclusions', body: 'Unless otherwise noted, this proposal excludes: electrical work and dedicated circuits, network core infrastructure and ISP services, architectural or structural modifications, permits and inspections, furniture and millwork, and ongoing software licensing fees.' }
  ];

  terms.forEach((term, i) => {
    contentChildren.push(
      para([txt(`${i + 1}. ${term.title}`, { size: 20, bold: true, color: COLOR.brand })], { before: 160, after: 80 }),
      para([txt(term.body, { size: 19, color: COLOR.text })], { after: 120 })
    );
  });

  // --- ACCEPTANCE ---
  contentChildren.push(
    para([txt(' ')], { before: 300 }),
    para([txt('Acceptance')], { heading: HeadingLevel.HEADING_1, before: 0, after: 160 }),
    para([txt('By signing below, the Client accepts this proposal and authorizes the commencement of work as described herein.', { size: 20 })], { after: 300 })
  );

  // Signature lines
  const SIG_COLS = [4380, 600, 4380]; // = 9360

  const sigLine = (label) => new TableRow({ children: [
    new TableCell({ borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 1, color: COLOR.text } }, width: { size: SIG_COLS[0], type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: [para([txt(' ')], { after: 0 })] }),
    new TableCell({ borders: NO_BORDERS, width: { size: SIG_COLS[1], type: WidthType.DXA }, children: [para([txt(' ')], { after: 0 })] }),
    new TableCell({ borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 1, color: COLOR.text } }, width: { size: SIG_COLS[2], type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: [para([txt(' ')], { after: 0 })] })
  ] });

  const sigLabel = (left, right) => new TableRow({ children: [
    new TableCell({ borders: NO_BORDERS, width: { size: SIG_COLS[0], type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: [para([txt(left, { size: 16, color: COLOR.muted })], { after: 0 })] }),
    new TableCell({ borders: NO_BORDERS, width: { size: SIG_COLS[1], type: WidthType.DXA }, children: [para([txt(' ')], { after: 0 })] }),
    new TableCell({ borders: NO_BORDERS, width: { size: SIG_COLS[2], type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 0, right: 0 },
      children: [para([txt(right, { size: 16, color: COLOR.muted })], { after: 0 })] })
  ] });

  contentChildren.push(
    new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: SIG_COLS, rows: [
      sigLine(), sigLabel('Client Signature', 'Date'),
      new TableRow({ children: SIG_COLS.map(w => new TableCell({ borders: NO_BORDERS, width: { size: w, type: WidthType.DXA }, children: [para([txt(' ')], { after: 0 })] })) }),
      sigLine(), sigLabel('Printed Name', 'Title')
    ] })
  );

  // ─── ASSEMBLE DOCUMENT ───
  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 20, color: COLOR.text } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: COLOR.brand },
          paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: COLOR.accent },
          paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } }
      ]
    },
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }]
    },
    sections: [
      coverSection,
      {
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
          }
        },
        headers: {
          default: new Header({
            children: [para([
              txt(companyName, { size: 16, color: COLOR.muted }),
              txt('  |  ', { size: 16, color: COLOR.border }),
              txt(projectName, { size: 16, color: COLOR.muted, italics: true })
            ], { after: 0 })]
          })
        },
        footers: {
          default: new Footer({
            children: [para([
              txt(`Proposal ${proposalNumber}`, { size: 14, color: COLOR.muted }),
              new TextRun({ children: [new Tab()], font: 'Arial', size: 14 }),
              txt('Page ', { size: 14, color: COLOR.muted }),
              new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 14, color: COLOR.muted })
            ], { align: AlignmentType.LEFT, after: 0, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }] })]
          })
        },
        children: contentChildren
      }
    ]
  });

  return Packer.toBuffer(doc);
}

// ─── MAIN ───
async function main() {
  // Load demo BOQ
  const boqPath = path.join(__dirname, '../data/demo-boq-output.json');
  if (!fs.existsSync(boqPath)) {
    console.error('Run processor.js first to generate demo-boq-output.json');
    process.exit(1);
  }

  const boq = JSON.parse(fs.readFileSync(boqPath, 'utf8'));

  const buffer = await generateProposal(boq, {
    companyName: 'RTS Technology Solutions',
    companyTagline: 'Audio Visual | Networking | Security',
    clientName: 'Sarah Chen',
    clientCompany: 'Oceanview Resort & Spa',
    clientAddress: '1200 Pacific Coast Highway, Malibu, CA 90265',
    validDays: 30,
    paymentTerms: 'Net 30',
    warrantyYears: 1
  });

  const outPath = path.join(__dirname, '../output/proposal.docx');
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, buffer);
  console.log(`Proposal generated: ${outPath}`);
  console.log(`File size: ${(buffer.length / 1024).toFixed(1)} KB`);
}

main().catch(console.error);

module.exports = { generateProposal };
