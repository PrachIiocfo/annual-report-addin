/* global Excel, Office, console, document, docx, URL, fetch, FileReader, Image */
/* eslint-disable no-undef */

const BRAND_BLUE = "1E3A8A";
const TABLE_HEADER_BG = "4B3293";
const LIGHT_BLUE = "DBEAFE";
const DARK_BLUE = "1E3A8A";
const TEXT_DARK = "1F2937";
const TEXT_GREY = "646464";
const BORDER_LIGHT = "B4C8DC";
const MISSING_RED = "C00000";
const MISSING_BG = "FFF3CD";

const MISS_OPEN = "\u2620MISS\u2620";
const MISS_CLOSE = "\u2620/MISS\u2620";

const SCALE_MAP = {
  "LAKH": 100000,
  "LAKHS": 100000,
  "CRORE": 10000000,
  "CRORES": 10000000,
  "THOUSAND": 1000,
  "THOUSANDS": 1000,
  "MILLION": 1000000,
  "MILLIONS": 1000000,
  "BILLION": 1000000000,
  "NONE": 1,
  "": 1
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office Add-in ready!");
    const btn = document.getElementById("genBtn");
    if (btn) btn.disabled = false;
  }
});

function getBranding() {
  return {
    COMPANY_NAME: (document.getElementById("companyName") || {}).value || "BLUE STAR LIMITED",
    COMPANY_CIN: (document.getElementById("companyCIN") || {}).value || "",
    HEADER_RIGHT_TOP: (document.getElementById("headerRightTop") || {}).value || "Financial Statements",
    HEADER_RIGHT_BOT: (document.getElementById("headerRightBot") || {}).value || "Consolidated",
    DOC_TITLE: (document.getElementById("docTitle") || {}).value || "Notes to Consolidated Financial Statements",
    DOC_SUBTITLE: (document.getElementById("docSubtitle") || {}).value || "for the year ended March 31, 2025"
  };
}

async function generate() {
  try {
    showStatus("Reading Excel data...", "blue");
    const branding = getBranding();

    const data = await Excel.run(async (context) => {
      const wb = context.workbook;
      const paraSheet = wb.worksheets.getItem("Para_ID");
      const dynSheet = wb.worksheets.getItem("Dynamic Values");
      const tabSheet = wb.worksheets.getItem("Tables");
      const landingSheet = wb.worksheets.getItem("Landing_Page");

      let vmRange = null;
      try {
        const vmSheet = wb.worksheets.getItem("Version Maintenance");
        vmRange = vmSheet.getUsedRange();
        vmRange.load("values");
      } catch (e) { console.log("VM not found:", e.message); }

      const paraRange = paraSheet.getUsedRange();
      const dynRange = dynSheet.getUsedRange();
      const tabRange = tabSheet.getUsedRange();
      const versionCell = landingSheet.getRange("C2");

      wb.load("name");
      paraRange.load("values");
      dynRange.load("values");
      tabRange.load("values");
      versionCell.load("values");

      await context.sync();

      return {
        paragraphs: paraRange.values,
        dynamics: dynRange.values,
        tables: tabRange.values,
        selectedVersion: String((versionCell.values[0][0] || "V1")).trim().toUpperCase(),
        versionMaint: vmRange ? vmRange.values : null,
        workbookName: wb.name || "Unknown.xlsx"
      };
    });

    console.log("Selected version:", data.selectedVersion);
    showStatus("Loading logo...", "blue");

    const logoUrl = lookupLogoUrl(data.versionMaint, data.selectedVersion);
    let logoBase64 = null;
    let logoDims = { w: 0, h: 0 };
    if (logoUrl) {
      logoBase64 = await fetchImageAsBase64(logoUrl);
      if (logoBase64) {
        logoDims = await getImageDimensions(logoBase64);
      }
    }

    const scaleInfo = lookupScale(data.versionMaint, data.selectedVersion);
    console.log("Scale for " + data.selectedVersion + ":", scaleInfo);

    showStatus("Building " + data.selectedVersion + " Word document...", "blue");

    const dynMap = buildDynamicMap(data.dynamics, scaleInfo);
    const tableMap = buildTablesMap(data.tables);
    const versionDetail = lookupVersionDetail(data.versionMaint, data.selectedVersion);
    const userName = (Office.context && Office.context.displayName) || "User";

    const missingTracker = { items: [] };

    const metadata = {
      version: data.selectedVersion,
      versionDetail: versionDetail,
      generatedBy: userName,
      generatedOn: new Date(),
      workbookName: data.workbookName,
      logoBase64: logoBase64,
      logoDims: logoDims,
      scaleLabel: scaleInfo.label,
      scaleDivisor: scaleInfo.divisor
    };

    const blob = await buildWordDoc(data.paragraphs, dynMap, tableMap, branding, data.selectedVersion, metadata, missingTracker);

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Annual_Report_" + data.selectedVersion + "_" + Date.now() + ".docx";
    a.click();
    URL.revokeObjectURL(url);

    const msg = missingTracker.items.length > 0
      ? "\u2713 " + data.selectedVersion + " downloaded \u2014 \u26A0 " + missingTracker.items.length + " missing tag(s)"
      : "\u2713 " + data.selectedVersion + " Word file downloaded!";
    showStatus(msg, missingTracker.items.length > 0 ? "blue" : "green");
  } catch (error) {
    console.error("Error:", error);
    showStatus("Error: " + error.message, "red");
  }
}

function lookupLogoUrl(vmRows, version) {
  if (!vmRows || vmRows.length === 0) return null;
  const header = vmRows[0] || [];
  let colLogo = -1;
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] || "").toLowerCase();
    if (h.indexOf("logo") >= 0 || h.indexOf("url") >= 0 || h.indexOf("image") >= 0) { colLogo = i; break; }
  }
  if (colLogo < 0 && header.length > 5) colLogo = 5;
  if (colLogo < 0) return null;
  for (let i = 1; i < vmRows.length; i++) {
    const row = vmRows[i];
    if (!row) continue;
    if (String(row[0] || "").trim().toUpperCase() === version) {
      const url = String(row[colLogo] || "").trim();
      if (url && (url.startsWith("http://") || url.startsWith("https://"))) return url;
    }
  }
  return null;
}

function lookupScale(vmRows, version) {
  const fallback = { label: "", divisor: 1 };
  if (!vmRows || vmRows.length === 0) return fallback;
  const header = vmRows[0] || [];
  let colScale = -1;
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] || "").toLowerCase().trim();
    if (h === "scale" || h.indexOf("scale") >= 0) { colScale = i; break; }
  }
  if (colScale < 0) return fallback;
  for (let i = 1; i < vmRows.length; i++) {
    const row = vmRows[i];
    if (!row) continue;
    if (String(row[0] || "").trim().toUpperCase() === version) {
      const scaleRaw = String(row[colScale] || "").trim();
      const scaleKey = scaleRaw.toUpperCase();
      const divisor = SCALE_MAP[scaleKey] !== undefined ? SCALE_MAP[scaleKey] : 1;
      return { label: scaleRaw, divisor: divisor };
    }
  }
  return fallback;
}

function normalizeGithubUrl(url) {
  if (!url) return url;
  let u = String(url).trim();
  u = u.replace(/[\?&]raw=true/gi, "");
  if (u.indexOf("github.com/") >= 0 && u.indexOf("/blob/") >= 0) {
    u = u.replace("github.com/", "raw.githubusercontent.com/");
    u = u.replace("/blob/", "/");
  }
  return u;
}

async function fetchImageAsBase64(url) {
  try {
    const cleanUrl = normalizeGithubUrl(url);
    const response = await fetch(cleanUrl);
    if (!response.ok) return null;
    const blob = await response.blob();
    return await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const result = reader.result;
        const idx = result.indexOf(",");
        resolve(idx >= 0 ? result.substring(idx + 1) : result);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch (e) { return null; }
}

function getImageDimensions(base64) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = "data:image/png;base64," + base64;
  });
}

function lookupVersionDetail(vmRows, version) {
  if (!vmRows || vmRows.length === 0) return "-";
  for (let i = 0; i < vmRows.length; i++) {
    const row = vmRows[i];
    if (!row || row.length === 0) continue;
    if (String(row[0] || "").trim().toUpperCase() === version) {
      return String(row[1] || "-").trim();
    }
  }
  return "-";
}

function makeNumberFormatter() {
  const counters = [0, 0, 0, 0, 0];
  const seenNumbers = {};
  function toLetter(n) {
    let result = "", num = n;
    while (num > 0) {
      const rem = (num - 1) % 26;
      result = String.fromCharCode(97 + rem) + result;
      num = Math.floor((num - 1) / 26);
    }
    return result;
  }
  function toRoman(n) {
    const map = ["", "i","ii","iii","iv","v","vi","vii","viii","ix","x","xi","xii","xiii","xiv","xv","xvi","xvii","xviii","xix","xx"];
    if (n <= 20) return map[n];
    const romans = [[1000,"m"],[900,"cm"],[500,"d"],[400,"cd"],[100,"c"],[90,"xc"],[50,"l"],[40,"xl"],[10,"x"],[9,"ix"],[5,"v"],[4,"iv"],[1,"i"]];
    let result = ""; let num = n;
    for (const [val, sym] of romans) { while (num >= val) { result += sym; num -= val; } }
    return result;
  }
  function assign(text, level) {
    const lvlIdx = level - 1;
    const parentContext = counters.slice(0, lvlIdx).join(".");
    const contextKey = parentContext + "|" + text;
    if (seenNumbers[contextKey]) return { number: seenNumbers[contextKey], isNew: false };
    counters[lvlIdx]++;
    for (let l = lvlIdx + 1; l < 5; l++) counters[l] = 0;
    let number = "";
    if (level === 1) number = counters[0] + ".";
    else if (level === 2) number = counters[0] + "." + counters[1];
    else if (level === 3) number = counters[0] + "." + counters[1] + "." + counters[2];
    else if (level === 4) number = "(" + toLetter(counters[3]) + ")";
    else if (level === 5) number = "(" + toRoman(counters[4]) + ")";
    seenNumbers[contextKey] = number;
    return { number: number, isNew: true };
  }
  return { assign: assign };
}

async function buildWordDoc(paraRows, dynMap, tableMap, B, version, metadata, missingTracker) {
  const d = docx;
  const usedTables = {};
  const globalPrintedHeadings = {};
  let tablesIncluded = 0;

  const pageHeader = makePageHeader(d, B, metadata.logoBase64, metadata.logoDims);

  const titleChildren = [
    new d.Paragraph({
      children: [new d.TextRun({ text: B.DOC_TITLE, bold: true, color: BRAND_BLUE, font: "Arial", size: 52 })],
      spacing: { before: 240, after: 80 }
    }),
    new d.Paragraph({
      children: [new d.TextRun({ text: B.DOC_SUBTITLE, color: TEXT_GREY, font: "Arial", size: 22 })],
      spacing: { after: 360 }
    })
  ];

  const headerRow = paraRows[0] || [];
  const colPara = findColIdx(headerRow, "Para");
  const colL6 = findColIdx(headerRow, "Level 6");
  const colLvl = [
    findColIdx(headerRow, "Level 1"),
    findColIdx(headerRow, "Level 2"),
    findColIdx(headerRow, "Level 3"),
    findColIdx(headerRow, "Level 4"),
    findColIdx(headerRow, "Level 5")
  ];
  const colVersion = findExactColIdx(headerRow, version);

  const sections = [];
  sections.push({
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1080, bottom: 1080, left: 1080 } },
      type: d.SectionType.CONTINUOUS
    },
    headers: { default: pageHeader },
    children: titleChildren
  });

  let currentChildren = [];
  function flush2Col() {
    if (currentChildren.length === 0) return;
    sections.push({
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1080, bottom: 1080, left: 1080 } },
        column: { count: 2, space: 432 },
        type: d.SectionType.CONTINUOUS
      },
      headers: { default: pageHeader },
      children: currentChildren
    });
    currentChildren = [];
  }

  let printedCount = 0, skippedCount = 0;
  const numberer = makeNumberFormatter();
  const stickyLevels = ["", "", "", "", ""];

  for (let i = 1; i < paraRows.length; i++) {
    const row = paraRows[i];
    let isCompletelyEmpty = true;
    for (let c = 0; c < headerRow.length; c++) {
      if (String(row[c] || "").trim() !== "") { isCompletelyEmpty = false; break; }
    }
    if (isCompletelyEmpty) continue;

    if (colVersion >= 0) {
      const flag = String(row[colVersion] || "").trim().toUpperCase();
      if (flag !== "Y") { skippedCount++; continue; }
    }
    printedCount++;

    const paraId = normalizeId(String(row[colPara] || ""));

    const rawTexts = [];
    for (let lvl = 1; lvl <= 5; lvl++) {
      const colIdx = colLvl[lvl - 1];
      if (colIdx < 0) { rawTexts.push(""); continue; }
      rawTexts.push(cleanStr(String(row[colIdx] || "")));
    }

    let highestFilledIdx = -1;
    for (let lvl = 0; lvl < 5; lvl++) {
      if (rawTexts[lvl]) { highestFilledIdx = lvl; break; }
    }

    const effectiveTexts = [];
    if (highestFilledIdx === -1) {
      for (let lvl = 0; lvl < 5; lvl++) effectiveTexts.push("");
    } else {
      for (let lvl = 0; lvl < highestFilledIdx; lvl++) effectiveTexts.push(stickyLevels[lvl]);
      effectiveTexts.push(rawTexts[highestFilledIdx]);
      stickyLevels[highestFilledIdx] = rawTexts[highestFilledIdx];
      for (let lvl = highestFilledIdx + 1; lvl < 5; lvl++) stickyLevels[lvl] = "";
      for (let lvl = highestFilledIdx + 1; lvl < 5; lvl++) {
        if (rawTexts[lvl]) { effectiveTexts.push(rawTexts[lvl]); stickyLevels[lvl] = rawTexts[lvl]; }
        else effectiveTexts.push("");
      }
    }

    const seenInRow = {};
    const finalLevelTexts = [];
    for (let lvl = 0; lvl < 5; lvl++) {
      const t = effectiveTexts[lvl];
      if (!t) { finalLevelTexts.push(null); continue; }
      if (seenInRow[t]) finalLevelTexts.push(null);
      else { seenInRow[t] = true; finalLevelTexts.push(t); }
    }

    for (let lvl = 1; lvl <= 5; lvl++) {
      const hText = finalLevelTexts[lvl - 1];
      if (!hText) continue;
      if (globalPrintedHeadings[hText]) continue;
      globalPrintedHeadings[hText] = true;
      const result = numberer.assign(hText, lvl);
      const numberedText = result.number + " " + hText;
      currentChildren.push(makeHeading(d, numberedText, lvl));
    }

    if (colL6 >= 0) {
      let body = cleanStr(String(row[colL6] || ""));
      if (body) {
        body = replaceDynamics(body, dynMap, paraId, missingTracker);
        currentChildren.push(makeBody(d, body));
      }
    }

    if (tableMap[paraId]) {
      const tblKey = tableMap[paraId].key;
      if (!usedTables[tblKey]) {
        usedTables[tblKey] = true;
        tablesIncluded++;
        flush2Col();
        const tblSectionChildren = [];
        tblSectionChildren.push(makeWordTable(d, tableMap[paraId]));
        tblSectionChildren.push(new d.Paragraph({ children: [new d.TextRun({ text: "" })], spacing: { after: 200 } }));
        sections.push({
          properties: {
            page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1080, bottom: 1080, left: 1080 } },
            type: d.SectionType.CONTINUOUS
          },
          headers: { default: pageHeader },
          children: tblSectionChildren
        });
      }
    }
  }

  flush2Col();

  const appendixChildren = buildAppendix(d, B, metadata, printedCount, skippedCount, tablesIncluded, missingTracker);
  sections.push({
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1080, bottom: 1080, left: 1080 } },
      type: d.SectionType.NEXT_PAGE
    },
    headers: { default: pageHeader },
    children: appendixChildren
  });

  const doc = new d.Document({
    styles: { default: { document: { run: { font: "Arial", size: 20 } } } },
    sections: sections
  });
  return await d.Packer.toBlob(doc);
}

function buildAppendix(d, B, m, printedCount, skippedCount, tablesIncluded, missingTracker) {
  const children = [];
  children.push(new d.Paragraph({
    children: [new d.TextRun({ text: "APPENDIX", bold: true, color: BRAND_BLUE, font: "Arial", size: 48 })],
    spacing: { before: 240, after: 80 }
  }));
  children.push(new d.Paragraph({
    children: [new d.TextRun({ text: "Document Generation Information", color: TEXT_GREY, font: "Arial", size: 22, italics: true })],
    spacing: { after: 320 }
  }));
  children.push(makeTableTitleBar(d, "Generation Details"));
  const dateStr = formatDateTime(m.generatedOn);
  const gen = [
    ["Document Title", B.DOC_TITLE], ["Subtitle", B.DOC_SUBTITLE],
    ["Company", B.COMPANY_NAME], ["CIN", B.COMPANY_CIN],
    ["Version Used", m.version], ["Version Detail", m.versionDetail],
    ["Scale Applied", m.scaleLabel || "None"],
    ["Generated By", m.generatedBy], ["Generated On", dateStr],
    ["Source File", m.workbookName], ["Logo Source", m.logoBase64 ? "Loaded from URL" : "Not available"]
  ];
  children.push(buildKeyValueTable(d, gen));
  children.push(new d.Paragraph({ children: [new d.TextRun({ text: "", size: 12 })], spacing: { before: 200 } }));
  children.push(makeTableTitleBar(d, "Content Statistics"));
  const missingCount = missingTracker ? missingTracker.items.length : 0;
  const stats = [
    ["Paragraphs Included", String(printedCount)],
    ["Paragraphs Skipped", String(skippedCount)],
    ["Tables Included", String(tablesIncluded)],
    ["Total Rows in Para_ID", String(printedCount + skippedCount)],
    ["Missing Dynamic Values", String(missingCount) + (missingCount > 0 ? " \u26A0" : "")]
  ];
  children.push(buildKeyValueTable(d, stats));

  if (missingCount > 0) {
    children.push(new d.Paragraph({ children: [new d.TextRun({ text: "", size: 12 })], spacing: { before: 240 } }));
    children.push(makeTableTitleBar(d, "Missing Dynamic Values \u2014 Review Required"));
    children.push(new d.Paragraph({
      children: [new d.TextRun({
        text: "The following dynamic value tags were referenced in paragraphs but not found in the Dynamic Values sheet. Please verify these values before finalising the document.",
        color: TEXT_GREY, italics: true, font: "Arial", size: 18
      })],
      spacing: { before: 100, after: 200 }
    }));
    children.push(buildMissingTagsTable(d, missingTracker.items));
  }

  children.push(new d.Paragraph({ children: [new d.TextRun({ text: "", size: 20 })], spacing: { before: 400 } }));
  children.push(new d.Paragraph({
    alignment: d.AlignmentType.CENTER,
    children: [new d.TextRun({ text: "Generated by Annual Report Auto-Generator  \u2022  iOCFO Consulting", color: TEXT_GREY, font: "Arial", size: 18, italics: true })]
  }));
  children.push(new d.Paragraph({
    alignment: d.AlignmentType.CENTER,
    children: [new d.TextRun({ text: dateStr, color: TEXT_GREY, font: "Arial", size: 16 })],
    spacing: { after: 200 }
  }));
  return children;
}

function buildMissingTagsTable(d, items) {
  const border = { style: d.BorderStyle.SINGLE, size: 4, color: "C8D2E6" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const totalWidth = 9360;
  const agg = {};
  for (const it of items) {
    if (!agg[it.tag]) agg[it.tag] = [];
    if (agg[it.tag].indexOf(it.paraId) === -1) agg[it.tag].push(it.paraId);
  }
  const rows = [];
  rows.push(new d.TableRow({
    cantSplit: true, tableHeader: true,
    children: [
      new d.TableCell({
        width: { size: 3000, type: d.WidthType.DXA }, borders: borders,
        shading: { type: d.ShadingType.CLEAR, fill: TABLE_HEADER_BG },
        margins: { top: 80, bottom: 80, left: 120, right: 100 },
        children: [new d.Paragraph({ children: [new d.TextRun({ text: "Dynamic Tag", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
      }),
      new d.TableCell({
        width: { size: 6360, type: d.WidthType.DXA }, borders: borders,
        shading: { type: d.ShadingType.CLEAR, fill: TABLE_HEADER_BG },
        margins: { top: 80, bottom: 80, left: 120, right: 100 },
        children: [new d.Paragraph({ children: [new d.TextRun({ text: "Used in Paragraph(s)", bold: true, color: "FFFFFF", font: "Arial", size: 18 })] })]
      })
    ]
  }));
  const tagKeys = Object.keys(agg).sort();
  for (let i = 0; i < tagKeys.length; i++) {
    const tag = tagKeys[i];
    const bg = i % 2 === 1 ? MISSING_BG : "FFFFFF";
    rows.push(new d.TableRow({
      cantSplit: true,
      children: [
        new d.TableCell({
          width: { size: 3000, type: d.WidthType.DXA }, borders: borders,
          shading: { type: d.ShadingType.CLEAR, fill: bg },
          margins: { top: 80, bottom: 80, left: 120, right: 100 },
          children: [new d.Paragraph({ children: [new d.TextRun({ text: tag, bold: true, color: MISSING_RED, font: "Arial", size: 18 })] })]
        }),
        new d.TableCell({
          width: { size: 6360, type: d.WidthType.DXA }, borders: borders,
          shading: { type: d.ShadingType.CLEAR, fill: bg },
          margins: { top: 80, bottom: 80, left: 120, right: 100 },
          children: [new d.Paragraph({ children: [new d.TextRun({ text: agg[tag].join(", "), color: TEXT_DARK, font: "Arial", size: 18 })] })]
        })
      ]
    }));
  }
  return new d.Table({
    width: { size: totalWidth, type: d.WidthType.DXA },
    columnWidths: [3000, 6360], rows: rows, layout: d.TableLayoutType.FIXED
  });
}

function buildKeyValueTable(d, rows) {
  const border = { style: d.BorderStyle.SINGLE, size: 4, color: "C8D2E6" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const tableRows = rows.map((kv, idx) => {
    const bg = idx % 2 === 1 ? "F4F7FC" : "FFFFFF";
    return new d.TableRow({
      cantSplit: true,
      children: [
        new d.TableCell({
          width: { size: 3400, type: d.WidthType.DXA }, borders: borders,
          shading: { type: d.ShadingType.CLEAR, fill: bg },
          margins: { top: 100, bottom: 100, left: 140, right: 100 },
          children: [new d.Paragraph({ children: [new d.TextRun({ text: kv[0], bold: true, color: DARK_BLUE, font: "Arial", size: 18 })] })]
        }),
        new d.TableCell({
          width: { size: 5960, type: d.WidthType.DXA }, borders: borders,
          shading: { type: d.ShadingType.CLEAR, fill: bg },
          margins: { top: 100, bottom: 100, left: 140, right: 100 },
          children: [new d.Paragraph({ children: [new d.TextRun({ text: kv[1], color: TEXT_DARK, font: "Arial", size: 18 })] })]
        })
      ]
    });
  });
  return new d.Table({ 
    width: { size: 9360, type: d.WidthType.DXA }, columnWidths: [3400, 5960], 
    rows: tableRows, layout: d.TableLayoutType.FIXED
  });
}

function formatDateTime(d) {
  const date = d || new Date();
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const day = String(date.getDate()).padStart(2, "0");
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12; if (hours === 0) hours = 12;
  return day + " " + months[date.getMonth()] + " " + date.getFullYear() + ", " + hours + ":" + minutes + " " + ampm;
}

function base64ToUint8Array(base64) {
  try {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  } catch (e) { return null; }
}

function makePageHeader(d, B, logoBase64, logoDims) {
  let rightCellChildren;
  if (logoBase64) {
    try {
      const bytes = base64ToUint8Array(logoBase64);
      if (bytes && bytes.length > 0) {
        const MAX_W = 150, MAX_H = 55;
        let imgW = MAX_W, imgH = MAX_H;
        if (logoDims && logoDims.w > 0 && logoDims.h > 0) {
          const ratio = logoDims.w / logoDims.h;
          imgH = MAX_H;
          imgW = Math.round(MAX_H * ratio);
          if (imgW > MAX_W) { imgW = MAX_W; imgH = Math.round(MAX_W / ratio); }
        }
        rightCellChildren = [
          new d.Paragraph({
            alignment: d.AlignmentType.RIGHT, spacing: { after: 100 },
            children: [new d.ImageRun({ data: bytes, transformation: { width: imgW, height: imgH } })]
          }),
          new d.Paragraph({
            alignment: d.AlignmentType.RIGHT, spacing: { after: 20 },
            children: [new d.TextRun({ text: B.HEADER_RIGHT_TOP, bold: true, color: BRAND_BLUE, font: "Arial", size: 18 })]
          }),
          new d.Paragraph({
            alignment: d.AlignmentType.RIGHT,
            children: [new d.TextRun({ text: B.HEADER_RIGHT_BOT, color: TEXT_GREY, font: "Arial", size: 15 })]
          })
        ];
      } else throw new Error("Empty bytes");
    } catch (e) {
      rightCellChildren = [
        new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: B.HEADER_RIGHT_TOP, bold: true, color: BRAND_BLUE, font: "Arial", size: 20 })] }),
        new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: B.HEADER_RIGHT_BOT, color: TEXT_GREY, font: "Arial", size: 16 })] })
      ];
    }
  } else {
    rightCellChildren = [
      new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: B.HEADER_RIGHT_TOP, bold: true, color: BRAND_BLUE, font: "Arial", size: 20 })] }),
      new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: B.HEADER_RIGHT_BOT, color: TEXT_GREY, font: "Arial", size: 16 })] })
    ];
  }

  return new d.Header({
    children: [new d.Table({
      width: { size: 9360, type: d.WidthType.DXA }, columnWidths: [4680, 4680],
      borders: {
        top: { style: d.BorderStyle.NONE },
        bottom: { style: d.BorderStyle.SINGLE, size: 6, color: BORDER_LIGHT },
        left: { style: d.BorderStyle.NONE }, right: { style: d.BorderStyle.NONE },
        insideHorizontal: { style: d.BorderStyle.NONE }, insideVertical: { style: d.BorderStyle.NONE }
      },
      rows: [new d.TableRow({
        children: [
          new d.TableCell({
            width: { size: 4680, type: d.WidthType.DXA }, borders: noBorders(d),
            verticalAlign: d.VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 0, right: 0 },
            children: [
              new d.Paragraph({ children: [new d.TextRun({ text: B.COMPANY_NAME, bold: true, color: BRAND_BLUE, font: "Arial", size: 22 })] }),
              new d.Paragraph({ children: [new d.TextRun({ text: B.COMPANY_CIN, color: TEXT_GREY, font: "Arial", size: 16 })] })
            ]
          }),
          new d.TableCell({
            width: { size: 4680, type: d.WidthType.DXA }, borders: noBorders(d),
            verticalAlign: d.VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 0, right: 0 },
            children: rightCellChildren
          })
        ]
      })]
    })]
  });
}

function noBorders(d) {
  return { top: { style: d.BorderStyle.NONE }, bottom: { style: d.BorderStyle.NONE }, left: { style: d.BorderStyle.NONE }, right: { style: d.BorderStyle.NONE } };
}

function makeHeading(d, text, level) {
  let size, bold, italics, color, indentLeft;
  switch (level) {
    case 1: size = 26; bold = true; italics = false; color = BRAND_BLUE; indentLeft = 0; break;
    case 2: size = 22; bold = true; italics = false; color = TEXT_DARK; indentLeft = 0; break;
    case 3: size = 21; bold = true; italics = true; color = TEXT_DARK; indentLeft = 100; break;
    case 4: size = 20; bold = true; italics = true; color = "404040"; indentLeft = 200; break;
    default: size = 20; bold = true; italics = true; color = "555555"; indentLeft = 300;
  }
  return new d.Paragraph({
    children: [new d.TextRun({ text: text, bold: bold, italics: italics, color: color, font: "Arial", size: size })],
    spacing: { before: level === 1 ? 280 : 200, after: level === 1 ? 140 : 100 },
    indent: { left: indentLeft }, keepNext: true
  });
}

function makeBody(d, text) {
  const runs = parseTextToRuns(d, text);
  return new d.Paragraph({
    alignment: d.AlignmentType.JUSTIFIED,
    children: runs,
    spacing: { after: 120 }
  });
}

function parseTextToRuns(d, text) {
  const runs = [];
  let cursor = 0;
  const re = new RegExp(MISS_OPEN + "(.*?)" + MISS_CLOSE, "g");
  let match;
  while ((match = re.exec(text)) !== null) {
    if (match.index > cursor) {
      runs.push(new d.TextRun({
        text: text.substring(cursor, match.index),
        color: TEXT_DARK, font: "Arial", size: 20
      }));
    }
    runs.push(new d.TextRun({
      text: match[1],
      bold: true, color: MISSING_RED, font: "Arial", size: 20,
      highlight: "yellow"
    }));
    cursor = match.index + match[0].length;
  }
  if (cursor < text.length) {
    runs.push(new d.TextRun({
      text: text.substring(cursor),
      color: TEXT_DARK, font: "Arial", size: 20
    }));
  }
  if (runs.length === 0) {
    runs.push(new d.TextRun({ text: text, color: TEXT_DARK, font: "Arial", size: 20 }));
  }
  return runs;
}

function makeTableTitleBar(d, title) {
  const totalWidth = 9360;
  return new d.Table({
    width: { size: totalWidth, type: d.WidthType.DXA },
    columnWidths: [totalWidth], layout: d.TableLayoutType.FIXED,
    borders: {
      top: { style: d.BorderStyle.NONE }, bottom: { style: d.BorderStyle.NONE },
      left: { style: d.BorderStyle.NONE }, right: { style: d.BorderStyle.NONE },
      insideHorizontal: { style: d.BorderStyle.NONE }, insideVertical: { style: d.BorderStyle.NONE }
    },
    rows: [new d.TableRow({
      cantSplit: true,
      children: [new d.TableCell({
        width: { size: totalWidth, type: d.WidthType.DXA },
        shading: { type: d.ShadingType.CLEAR, fill: TABLE_HEADER_BG },
        margins: { top: 80, bottom: 80, left: 140, right: 100 },
        children: [new d.Paragraph({
          children: [new d.TextRun({ text: title, bold: true, color: "FFFFFF", font: "Arial", size: 20 })],
          keepNext: true
        })]
      })]
    })]
  });
}

// === UPDATED: Header rows keep with next + repeat across pages ===
function makeWordTable(d, info) {
  const numRows = info.endRow - info.headerRow + 1;
  const numCols = info.numCols;
  if (numRows < 1 || numCols < 1) return new d.Paragraph({ children: [new d.TextRun({ text: "" })] });
  
  const totalWidth = 9360;
  
  const colMaxLen = new Array(numCols).fill(1);
  for (let r = 0; r < numRows; r++) {
    const rowData = info.allRows[info.headerRow + r] || [];
    for (let c = 0; c < numCols; c++) {
      const len = String(rowData[c] || "").length;
      if (len > colMaxLen[c]) colMaxLen[c] = len;
    }
  }
  const weights = colMaxLen.map((len, i) => Math.max(len, 5) * (i === 0 ? 1.3 : 1.0));
  const totalWeight = weights.reduce((a, b) => a + b, 0);
  
  const MIN_W = 700;
  let columnWidths = weights.map(w => Math.max(MIN_W, Math.floor((w / totalWeight) * totalWidth)));
  let sum = columnWidths.reduce((a, b) => a + b, 0);
  if (sum > totalWidth) {
    const scale = totalWidth / sum;
    columnWidths = columnWidths.map(w => Math.floor(w * scale));
    sum = columnWidths.reduce((a, b) => a + b, 0);
  }
  columnWidths[0] += (totalWidth - sum);

  const border = { style: d.BorderStyle.SINGLE, size: 4, color: "C8D2E6" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const rows = [];
  for (let r = 0; r < numRows; r++) {
    const sheetRow = info.headerRow + r;
    const rowData = info.allRows[sheetRow] || [];
    const isHeader = (r === 0);
    const firstCol = String(rowData[0] || "").trim();
    let isHighlight = false, isSubHeader = false;
    if (!isHeader) {
      const fc = firstCol.toLowerCase();
      if (fc.startsWith("as at") || fc.startsWith("at march") || fc.startsWith("balance as") || fc.startsWith("total")) isHighlight = true;
      let empty = 0;
      for (let c = 1; c < numCols; c++) if (!rowData[c] || String(rowData[c]).trim() === "") empty++;
      if (empty === numCols - 1 && firstCol) isSubHeader = true;
    }
    const cells = [];
    for (let c = 0; c < numCols; c++) {
      const cellText = String(rowData[c] !== undefined && rowData[c] !== null ? rowData[c] : "").trim();
      let bg = "FFFFFF", txtCol = TEXT_DARK, fSize = 16, isB = false, isI = false;
      let align = (c === 0) ? d.AlignmentType.LEFT : d.AlignmentType.RIGHT;
      if (isHeader) { bg = TABLE_HEADER_BG; txtCol = "FFFFFF"; isB = true; align = d.AlignmentType.CENTER; }
      else if (isSubHeader) { bg = LIGHT_BLUE; txtCol = DARK_BLUE; isB = true; isI = true; align = d.AlignmentType.LEFT; }
      else if (isHighlight) { bg = LIGHT_BLUE; txtCol = DARK_BLUE; isB = true; }
      cells.push(new d.TableCell({
        width: { size: columnWidths[c], type: d.WidthType.DXA },
        borders: borders, shading: { type: d.ShadingType.CLEAR, fill: bg },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new d.Paragraph({
          alignment: align,
          keepNext: isHeader,
          keepLines: true,
          children: [new d.TextRun({ text: cellText, bold: isB, italics: isI, color: txtCol, font: "Arial", size: fSize })]
        })]
      }));
    }
    rows.push(new d.TableRow({ 
      children: cells, 
      cantSplit: true,
      tableHeader: isHeader  // Header repeats on next page if table splits
    }));
  }
  return new d.Table({ 
    width: { size: totalWidth, type: d.WidthType.DXA }, columnWidths: columnWidths, 
    rows: rows, layout: d.TableLayoutType.FIXED
  });
}

function findColIdx(headerRow, keyword) {
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || "").toLowerCase().indexOf(keyword.toLowerCase()) >= 0) return i;
  }
  return -1;
}

function findExactColIdx(headerRow, exactName) {
  const target = String(exactName).trim().toUpperCase();
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || "").trim().toUpperCase() === target) return i;
  }
  return -1;
}

function cleanStr(s) {
  return String(s || "").replace(/\r/g, " ").replace(/\n/g, " ").replace(/\t/g, " ").replace(/`/g, "").replace(/\s+/g, " ").trim();
}

function normalizeId(s) {
  const m = /Para[_\s]*0*(\d+)/i.exec(String(s || ""));
  if (!m) return String(s || "").trim();
  return "Para_" + String(m[1]).padStart(3, "0");
}

function formatIndianNumber(num) {
  const isNeg = num < 0;
  const absNum = Math.abs(num);
  const fixed = absNum.toFixed(2);
  const parts = fixed.split(".");
  let intPart = parts[0];
  const decPart = parts[1];
  if (intPart.length > 3) {
    const last3 = intPart.substring(intPart.length - 3);
    const rest = intPart.substring(0, intPart.length - 3);
    const restFormatted = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ",");
    intPart = restFormatted + "," + last3;
  }
  return (isNeg ? "(" : "") + intPart + "." + decPart + (isNeg ? ")" : "");
}

function applyScaleAndFormat(rawValue, type, scaleDivisor, scaleLabel) {
  const valStr = String(rawValue || "").trim();
  if (!valStr) return "";
  
  const typeNorm = String(type || "").trim().toUpperCase();
  
  if (typeNorm !== "VALUES" && typeNorm !== "VALUE") {
    return valStr;
  }
  
  const cleanNum = valStr.replace(/[,\s\u20B9]/g, "");
  const num = parseFloat(cleanNum);
  if (isNaN(num)) return valStr;
  
  if (scaleDivisor && scaleDivisor > 1) {
    const scaled = num / scaleDivisor;
    return formatIndianNumber(scaled);
  }
  return formatIndianNumber(num);
}

function buildDynamicMap(rows, scaleInfo) {
  const map = {};
  if (!rows || rows.length < 2) return map;
  const header = rows[0];
  let colId = -1, colVal = -1, colType = -1;
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] || "").toLowerCase().trim();
    if (colId < 0 && h.indexOf("dynamic") >= 0 && h.indexOf("para") < 0) colId = i;
    if (colVal < 0 && h.indexOf("para") < 0 &&
        (h === "data" || h === "value" || h.indexOf("example") >= 0)) colVal = i;
    if (colType < 0 && (h === "type" || h.indexOf("type") >= 0) && h.indexOf("para") < 0) colType = i;
  }
  console.log("DynMap cols \u2192 ID:", colId, "Value:", colVal, "Type:", colType);
  if (colId < 0 || colVal < 0) return map;
  
  const divisor = scaleInfo ? scaleInfo.divisor : 1;
  const scaleLabel = scaleInfo ? scaleInfo.label : "";
  
  for (let r = 1; r < rows.length; r++) {
    let id = String(rows[r][colId] || "").trim().toUpperCase().replace(/[\s-]/g, "_");
    let rawVal = String(rows[r][colVal] || "").replace(/[<>`]/g, "").trim();
    let type = colType >= 0 ? String(rows[r][colType] || "").trim() : "";
    if (!id) continue;
    const finalVal = applyScaleAndFormat(rawVal, type, divisor, scaleLabel);
    map[id] = finalVal;
  }
  console.log("Dynamic map built:", Object.keys(map).length, "entries. Scale:", scaleLabel, "/", divisor);
  return map;
}

function replaceDynamics(text, dynMap, paraId, missingTracker) {
  return String(text || "").replace(/<\s*((?:DV|DT|DD)[\s_\-]*\d+)[^>]*>/gi, function(match, id) {
    const key = id.toUpperCase().replace(/[\s-]/g, "_");
    if (dynMap[key] === undefined) {
      if (missingTracker) {
        missingTracker.items.push({ tag: key, paraId: paraId || "?" });
      }
      console.warn("Missing dynamic tag:", key, "in", paraId);
      return MISS_OPEN + match + MISS_CLOSE;
    }
    return dynMap[key];
  });
}

function buildTablesMap(rows) {
  const map = {};
  if (!rows || rows.length === 0) return map;
  const lastRow = rows.length - 1;
  const lastCol = rows[0] ? rows[0].length : 0;
  
  const header = rows[0] || [];
  let colTableId = 0, colParaId = 1, colDataStart = 2;
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] || "").toLowerCase().trim();
    if (h === "table_id" || h === "tableid" || h.indexOf("table id") >= 0 || h.indexOf("table_id") >= 0) colTableId = i;
    if (h === "para_id" || h === "paraid" || h.indexOf("para id") >= 0 || h.indexOf("para_id") >= 0) colParaId = i;
  }
  colDataStart = Math.max(colTableId, colParaId) + 1;
  
  let r = 1;
  while (r <= lastRow) {
    const tableId = String((rows[r] && rows[r][colTableId]) || "").trim();
    
    if (tableId.toLowerCase().startsWith("table_")) {
      const tableStartRow = r;
      const paraId = normalizeId(String((rows[r] && rows[r][colParaId]) || ""));
      
      let endRow = r;
      let lookAhead = r + 1;
      while (lookAhead <= lastRow) {
        const nextTableId = String((rows[lookAhead] && rows[lookAhead][colTableId]) || "").trim();
        if (nextTableId.toLowerCase().startsWith("table_")) break;
        let hasData = false;
        for (let c = colDataStart; c < lastCol; c++) {
          if (rows[lookAhead] && String(rows[lookAhead][c] || "").trim() !== "") { hasData = true; break; }
        }
        if (hasData) endRow = lookAhead;
        lookAhead++;
      }
      
      let numCols = 0;
      for (let rr = tableStartRow; rr <= endRow; rr++) {
        for (let c = colDataStart; c < lastCol; c++) {
          if (rows[rr] && String(rows[rr][c] || "").trim() !== "") {
            if (c - colDataStart + 1 > numCols) numCols = c - colDataStart + 1;
          }
        }
      }
      if (numCols < 2) numCols = 2;
      
      const tableRows = [];
      for (let rr = tableStartRow; rr <= endRow; rr++) {
        const adjustedRow = [];
        for (let c = 0; c < numCols; c++) {
          const sourceCol = colDataStart + c;
          adjustedRow.push((rows[rr] && rows[rr][sourceCol] !== undefined) ? rows[rr][sourceCol] : "");
        }
        tableRows.push(adjustedRow);
      }
      
      const info = {
        key: "T_" + tableId,
        title: tableId.replace(/_/g, " "),
        headerRow: 0,
        endRow: tableRows.length - 1,
        numCols: numCols,
        allRows: tableRows
      };
      
      if (paraId && !map[paraId]) map[paraId] = info;
      r = endRow + 1;
    } else { r++; }
  }
  return map;
}

function showStatus(text, color) {
  const colors = { blue: "#3b82f6", green: "#10b981", red: "#ef4444" };
  const old = document.getElementById("statusBox");
  if (old) old.remove();
  const box = document.createElement("div");
  box.id = "statusBox";
  box.style.cssText = "padding:14px;background:" + (colors[color] || "#6b7280") +
    ";color:white;font-weight:bold;margin:10px 0;border-radius:8px;font-size:13px;" +
    "position:sticky;top:0;z-index:9999;box-shadow:0 2px 8px rgba(0,0,0,.15);";
  box.textContent = text;
  document.body.insertBefore(box, document.body.firstChild);
}

window.generate = generate;