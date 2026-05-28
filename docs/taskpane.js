/* global Excel, Office, console, document, docx, URL, fetch, FileReader, pdfjsLib, ExcelJS */
/* eslint-disable no-undef */

const BRAND_BLUE = "1E3A8A";
const TABLE_HEADER_BG = "4B3293";
const LIGHT_BLUE = "DBEAFE";
const DARK_BLUE = "1E3A8A";
const TEXT_DARK = "1F2937";
const TEXT_GREY = "646464";
const BORDER_LIGHT = "B4C8DC";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office ready!");
    checkRememberedLogin();
  }
});

// ============================================================
// LOGIN SYSTEM
// ============================================================

function checkRememberedLogin() {
  try {
    const saved = localStorage.getItem("iocfo_login");
    if (saved) {
      const data = JSON.parse(saved);
      if (data.expiry && data.expiry > Date.now()) {
        sessionStorage.setItem("iocfo_currentUser", data.username);
        updateUserUI(data.username);
        showMainContent();
        return;
      }
      localStorage.removeItem("iocfo_login");
    }
  } catch (e) { console.warn(e); }
  
  // Allow Enter key on login fields
  const userInput = document.getElementById("loginUsername");
  const passInput = document.getElementById("loginPassword");
  if (userInput) userInput.addEventListener("keypress", (e) => { if (e.key === "Enter") doLogin(); });
  if (passInput) passInput.addEventListener("keypress", (e) => { if (e.key === "Enter") doLogin(); });
}

function showLoginError(msg) {
  const err = document.getElementById("loginError");
  if (err) {
    err.style.display = "block";
    err.innerText = msg;
  }
}

function hideLoginError() {
  const err = document.getElementById("loginError");
  if (err) err.style.display = "none";
}

function showMainContent() {
  const overlay = document.getElementById("loginOverlay");
  const main = document.getElementById("mainContent");
  if (overlay) overlay.style.display = "none";
  if (main) main.style.display = "block";
}

async function doLogin() {
  hideLoginError();
  const username = document.getElementById("loginUsername").value.trim();
  const password = document.getElementById("loginPassword").value.trim();
  const remember = document.getElementById("rememberMe").checked;
  
  console.log("=== LOGIN ATTEMPT ===");
  console.log("Entered Username:", JSON.stringify(username));
  console.log("Entered Password:", JSON.stringify(password));
  
  if (!username || !password) {
    showLoginError("Please enter both username and password");
    return;
  }
  
  const btn = document.getElementById("loginBtn");
  btn.disabled = true;
  btn.innerText = "Signing in...";
  
  try {
    const response = await fetch("assets/Users.xlsx");
    console.log("Fetch status:", response.status, response.ok);
    
    if (!response.ok) {
      showLoginError("Unable to verify credentials. Please try again.");
      btn.disabled = false;
      btn.innerText = "Sign In";
      return;
    }
    
    const buffer = await response.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    
    console.log("Sheets in Users.xlsx:", workbook.worksheets.map(s => s.name));
    
    let usersSheet = workbook.getWorksheet("Users");
    if (!usersSheet) {
      for (const ws of workbook.worksheets) {
        if (ws.name.toLowerCase().includes("user")) { usersSheet = ws; break; }
      }
    }
    
    if (!usersSheet) {
      showLoginError("User database not found");
      btn.disabled = false;
      btn.innerText = "Sign In";
      return;
    }
    
    console.log("Using sheet:", usersSheet.name);
    
    let matched = false;
    usersSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const u = String(row.getCell(1).value || "").trim();
      let pVal = row.getCell(2).value;
if (pVal && typeof pVal === "object") {
  pVal = pVal.text || pVal.hyperlink || pVal.formula || "";
}
const p = String(pVal || "").trim();
      console.log("Row " + rowNumber + " - User:", JSON.stringify(u), "| Pass:", JSON.stringify(p));
      if (u.toLowerCase() === username.toLowerCase() && p === password) {
        matched = true;
        console.log("✓ MATCH FOUND!");
      }
    });
    
    if (matched) {
      if (remember) {
        const expiry = Date.now() + (7 * 24 * 60 * 60 * 1000);
        localStorage.setItem("iocfo_login", JSON.stringify({ username, expiry }));
      }
      sessionStorage.setItem("iocfo_currentUser", username);
      updateUserUI(username);
      showMainContent();
    } else {
      showLoginError("Invalid username or password");
      btn.disabled = false;
      btn.innerText = "Sign In";
    }
  } catch (err) {
    console.error("Login error:", err);
    showLoginError("Login failed: " + err.message);
    btn.disabled = false;
    btn.innerText = "Sign In";
  }
}

// ============================================================
// EXISTING CODE (unchanged)
// ============================================================

async function downloadTemplate() {
  try {
    const response = await fetch("assets/Excel_Template.xlsx");
    if (!response.ok) { console.log("Template file not found!"); return; }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Annual_Report_Template.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) { console.error("Download error:", err); }
}

let selectedPdfFile = null;

function handleFileSelect() {
  const input = document.getElementById("pdfFile");
  const fileName = document.getElementById("fileName");
  const btn = document.getElementById("processPdfBtn");
  if (input.files && input.files[0]) {
    selectedPdfFile = input.files[0];
    if (fileName) {
      fileName.style.display = "block";
      fileName.innerText = "📄 " + selectedPdfFile.name + " (" + (selectedPdfFile.size / 1024 / 1024).toFixed(1) + " MB)";
    }
    if (btn) btn.disabled = false;
  } else {
    selectedPdfFile = null;
    if (fileName) fileName.style.display = "none";
    if (btn) btn.disabled = true;
  }
}

function showPdfStatus(msg, type) {
  const box = document.getElementById("pdfStatus");
  if (!box) return;
  box.style.display = "block";
  box.className = type || "info";
  box.innerText = msg;
}

async function processPDF() {
  if (!selectedPdfFile) { showPdfStatus("Please select a PDF first", "error"); return; }
  const fromPage = parseInt(document.getElementById("fromPage").value) || 1;
  const toPage = parseInt(document.getElementById("toPage").value) || 50;
  if (fromPage > toPage) { showPdfStatus("From page must be less than To page", "error"); return; }

  const btn = document.getElementById("processPdfBtn");
  btn.disabled = true;
  btn.innerText = "⏳ Reading PDF...";

  try {
    showPdfStatus("Loading PDF...", "info");
    const arrayBuffer = await selectedPdfFile.arrayBuffer();
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";

    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const totalPages = pdf.numPages;
    const endPage = Math.min(toPage, totalPages);

    btn.innerText = "⏳ Extracting text...";
    showPdfStatus("Extracting pages " + fromPage + " to " + endPage + "...", "info");

    let allItems = [];
    for (let i = fromPage; i <= endPage; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const viewport = page.getViewport({ scale: 1 });
      const pageWidth = viewport.width;

      const lines = {};
      for (const item of content.items) {
        const y = Math.round(item.transform[5]);
        const x = item.transform[4];
        const col = x < pageWidth / 2 ? "L" : "R";
        const key = y + "_" + col;
        if (!lines[key]) lines[key] = { text: "", fontSize: 0, x: x, y: y, col: col };
        lines[key].text += item.str + " ";
        lines[key].fontSize = Math.max(lines[key].fontSize, item.height);
      }

      const leftKeys = Object.keys(lines).filter(k => k.endsWith("_L")).sort((a, b) => lines[b].y - lines[a].y);
      const rightKeys = Object.keys(lines).filter(k => k.endsWith("_R")).sort((a, b) => lines[b].y - lines[a].y);

      for (const k of leftKeys) {
        const line = lines[k];
        if (line.text.trim().length > 1) allItems.push({ text: line.text.trim(), fontSize: line.fontSize, page: i });
      }
      for (const k of rightKeys) {
        const line = lines[k];
        if (line.text.trim().length > 1) allItems.push({ text: line.text.trim(), fontSize: line.fontSize, page: i });
      }
    }

    btn.innerText = "⏳ Analyzing structure...";
    showPdfStatus("Analyzing paragraphs and headings...", "info");

    const paragraphs = extractParagraphs(allItems);

    if (paragraphs.length === 0) {
      showPdfStatus("No content found. Try a different page range.", "error");
      btn.disabled = false;
      btn.innerText = "🚀 Process & Fill Excel";
      return;
    }

    btn.innerText = "⏳ Building Excel...";
    showPdfStatus("Creating Excel with " + paragraphs.length + " paragraphs...", "info");

    await buildExcelFromParagraphs(paragraphs);

    showPdfStatus("✓ Done! " + paragraphs.length + " paragraphs extracted.", "success");
    btn.disabled = false;
    btn.innerText = "🚀 Process & Fill Excel";
  } catch (err) {
    console.error("PDF error:", err);
    showPdfStatus("Error: " + err.message, "error");
    btn.disabled = false;
    btn.innerText = "🚀 Process & Fill Excel";
  }
}

function classifyLine(text, fontSize, bodyFontSize) {
  if (/^\d+\.\s+[A-Z][A-Z\s]+$/.test(text) && text.length < 150) return "heading_1";
  if (/^\d+\.\s*$/.test(text)) return "heading_1";
  if (/^\d+\.\s+[A-Z]/.test(text) && text.length < 150) return "heading_1";
  if (/^\([a-z]\)\s+[A-Z]/.test(text) && text.length < 200) return "heading_2";
  if (/^\([a-z]\)\s*$/.test(text)) return "heading_2";
  if (/^\(i{1,4}v?\)\s+[A-Z]/.test(text) && text.length < 200) return "heading_3";
  if (/^\(i{1,4}v?\)\s*$/.test(text)) return "heading_3";
  if (text === text.toUpperCase() && text.length > 3 && text.length < 80) {
    const letters = text.replace(/[^A-Z]/g, "");
    if (letters.length > 3) return "heading_1";
  }
  if (fontSize > bodyFontSize * 1.3 && text.length < 150) return "heading_2";
  if (text.startsWith("•") || text.startsWith("●") || /^[\-\*]\s/.test(text)) return "bullet";
  return "body";
}

function extractParagraphs(items) {
  if (items.length === 0) return [];
  const sizeFreq = {};
  for (const it of items) {
    const r = Math.round(it.fontSize * 10) / 10;
    sizeFreq[r] = (sizeFreq[r] || 0) + 1;
  }
  let bodyFontSize = 10;
  let maxFreq = 0;
  for (const s in sizeFreq) {
    if (sizeFreq[s] > maxFreq) { maxFreq = sizeFreq[s]; bodyFontSize = parseFloat(s); }
  }
  const clean = [];
  for (const it of items) {
    const t = it.text.trim();
    if (t.length < 2) continue;
    if (/^\d+$/.test(t)) continue;
    if (/^Page \d+/i.test(t)) continue;
    if (t === "BLUE STAR LIMITED") continue;
    if (/^\(CIN\s*:/.test(t)) continue;
    if (t === "Financial Statements" || t === "Consolidated") continue;
    if (t === "Notes to Consolidated Financial Statements") continue;
    if (/^for the year ended/i.test(t)) continue;
    clean.push(it);
  }
  const merged = [];
  let prev = null;
  for (let i = 0; i < clean.length; i++) {
    const it = clean[i];
    const text = it.text.trim();
    const type = classifyLine(text, it.fontSize, bodyFontSize);
    if (prev && prev.type.startsWith("heading_")) {
      const isAllCaps = text === text.toUpperCase() && /[A-Z]/.test(text);
      const isShort = text.length < 80;
      const noNumbering = !/^\d+\./.test(text) && !/^\([a-z]\)/.test(text) && !/^\(i{1,4}v?\)/.test(text);
      if (prev.text === prev.text.toUpperCase() && isAllCaps && isShort && noNumbering && prev.type === "heading_1") {
        prev.text = prev.text + " " + text;
        prev.fontSize = Math.max(prev.fontSize, it.fontSize);
        continue;
      }
      if (prev.text && !/[.:;?!]$/.test(prev.text) && it.fontSize >= bodyFontSize * 1.2 && isShort && noNumbering) {
        prev.text = prev.text + " " + text;
        prev.fontSize = Math.max(prev.fontSize, it.fontSize);
        continue;
      }
    }
    const newItem = { text: text, type: type, fontSize: it.fontSize };
    merged.push(newItem);
    prev = newItem;
  }
  const paragraphs = [];
  const stack = ["", "", "", "", ""];
  let currentBody = "";
  function flushP() {
    if (currentBody.trim().length < 30) { currentBody = ""; return; }
    paragraphs.push({
      level1: stack[0] || "", level2: stack[1] || "", level3: stack[2] || "",
      level4: stack[3] || "", level5: stack[4] || "",
      level6: currentBody.trim().replace(/\s+/g, " ")
    });
    currentBody = "";
  }
  for (let i = 0; i < merged.length; i++) {
    const it = merged[i];
    if (it.type === "heading_1" || it.type === "heading_2") {
      flushP();
      const level = parseInt(it.type.split("_")[1]);
      stack[level - 1] = it.text;
      for (let l = level; l < 5; l++) stack[l] = "";
      currentBody = "";
    } else if (it.type === "heading_3") {
      currentBody += " " + it.text;
    } else if (it.type === "bullet") {
      currentBody += " • " + it.text;
    } else {
      currentBody += " " + it.text;
    }
  }
  flushP();
  return paragraphs;
}

function cleanHeading(text) {
  if (!text) return "";
  let t = text.trim();
  t = t.replace(/^\d+\.\s*/, "");
  t = t.replace(/^\([a-z]\)\s*/i, "");
  t = t.replace(/^\(i{1,4}v?x?\)\s*/i, "");
  t = t.replace(/^[•●\-\*]\s*/, "");
  return t.trim();
}

async function buildExcelFromParagraphs(paragraphs) {
  const response = await fetch("assets/Excel_Template.xlsx");
  if (!response.ok) throw new Error("Template file not found in assets folder");
  const templateBuffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  let paraSheet = workbook.getWorksheet("Para_ID");
  if (!paraSheet) {
    for (const ws of workbook.worksheets) {
      if (ws.name.toLowerCase().includes("para")) { paraSheet = ws; break; }
    }
  }
  if (!paraSheet) throw new Error("Para_ID sheet not found");

  const lastRow = paraSheet.actualRowCount;
  for (let r = lastRow; r >= 2; r--) paraSheet.spliceRows(r, 1);

  let rowNum = 2;
  let paraCounter = 1;
  for (const p of paragraphs) {
    if (!p.level6 || p.level6.length < 30) continue;
    const numCount = (p.level6.match(/\d/g) || []).length;
    if (numCount / p.level6.length > 0.4) continue;
    const paraId = "Para_" + String(paraCounter).padStart(2, "0");
    paraCounter++;
    const cL1 = cleanHeading(p.level1);
    const cL2 = cleanHeading(p.level2);
    const cL3 = cleanHeading(p.level3);
    const cL4 = cleanHeading(p.level4);
    const cL5 = cleanHeading(p.level5);

    const row = paraSheet.getRow(rowNum);
    row.getCell(1).value = paraId;
    let cascade = cL1 || "";
    row.getCell(2).value = cL1 || "";
    if (cL2) cascade = cL2;
    row.getCell(3).value = cL2 || cascade;
    if (cL3) cascade = cL3;
    row.getCell(4).value = cL3 || cascade;
    if (cL4) cascade = cL4;
    row.getCell(5).value = cL4 || cascade;
    if (cL5) cascade = cL5;
    row.getCell(6).value = cL5 || cascade;
    row.getCell(7).value = p.level6;
    row.getCell(8).value = "Y";
    row.getCell(9).value = "Y";
    row.getCell(10).value = "Y";
    row.commit();
    rowNum++;
  }

  const outBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Annual_Report_Filled_" + Date.now() + ".xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

async function generate() {
  const btn = document.getElementById("genBtn");
  btn.disabled = true;
  btn.querySelector("span:last-child").innerText = "Generating...";

  try {
    let COMPANY_NAME = "Company Name";
    let COMPANY_CIN = "";
    let HEADER_RIGHT_TOP = "Financial Statements";
    let HEADER_RIGHT_BOT = "Consolidated";
    let DOC_TITLE = "Notes to Consolidated Financial Statements";
    let DOC_SUBTITLE = "for the year ended March 31, 2025";
    let YEAR = "";
    let SCALE = "Crore";
    let LOGO_URL = "";
    let selectedVersion = "V1";
    let versionDetail = "-";
    let paragraphs = null, dynamics = null, tables = null;
    let workbookName = "Workbook.xlsx";

    await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const allSheets = wb.worksheets;
      allSheets.load("items/name");
      await ctx.sync();
      const names = allSheets.items.map(s => s.name);
      console.log("Sheets:", names);

      function findSheet(keywords) {
        for (const k of keywords) {
          const m = names.find(n => n.toLowerCase().includes(k.toLowerCase()));
          if (m) return wb.worksheets.getItem(m);
        }
        return null;
      }

      const paraSheet = findSheet(["Para_ID", "Para"]);
      const dynSheet = findSheet(["Dynamic Values", "Dynamic"]);
      const tabSheet = findSheet(["Tables"]);
      const landingSheet = findSheet(["Landing"]);
      const vmSheet = findSheet(["Version"]);

      const paraRange = paraSheet ? paraSheet.getUsedRange() : null;
      const dynRange = dynSheet ? dynSheet.getUsedRange() : null;
      const tabRange = tabSheet ? tabSheet.getUsedRange() : null;
      const vmRange = vmSheet ? vmSheet.getUsedRange() : null;
      const verCell = landingSheet ? landingSheet.getRange("C2") : null;

      wb.load("name");
      if (paraRange) paraRange.load("values");
      if (dynRange) dynRange.load("values");
      if (tabRange) tabRange.load("values");
      if (vmRange) vmRange.load("values");
      if (verCell) verCell.load("values");
      await ctx.sync();

      workbookName = wb.name || "Workbook.xlsx";
      paragraphs = paraRange ? paraRange.values : null;
      dynamics = dynRange ? dynRange.values : null;
      tables = tabRange ? tabRange.values : null;
      if (verCell) selectedVersion = String(verCell.values[0][0] || "V1").trim().toUpperCase();

      if (vmRange) {
        const vmRows = vmRange.values;
        for (let i = 1; i < vmRows.length; i++) {
          if (vmRows[i] && String(vmRows[i][0] || "").trim().toUpperCase() === selectedVersion) {
            const r = vmRows[i];
            versionDetail = String(r[1] || "-").trim();
            YEAR = String(r[2] || "").trim();
            LOGO_URL = String(r[5] || "").trim();
            SCALE = String(r[6] || "Crore").trim();
            COMPANY_NAME = String(r[7] || COMPANY_NAME).trim();
            COMPANY_CIN = String(r[8] || "").trim();
            HEADER_RIGHT_TOP = String(r[9] || HEADER_RIGHT_TOP).trim();
            HEADER_RIGHT_BOT = String(r[10] || HEADER_RIGHT_BOT).trim();
            DOC_TITLE = String(r[11] || DOC_TITLE).trim();
            DOC_SUBTITLE = String(r[12] || DOC_SUBTITLE).trim();
            break;
          }
        }
      }
    });

    console.log("Version:", selectedVersion);
    console.log("Scale:", SCALE);

    let logoBase64 = null;
    let logoSource = "Not available";
    if (LOGO_URL && (LOGO_URL.startsWith("http://") || LOGO_URL.startsWith("https://"))) {
      logoBase64 = await fetchImage(LOGO_URL);
      if (logoBase64) logoSource = "URL from Version Maintenance";
    }

    const B = { COMPANY_NAME, COMPANY_CIN, HEADER_RIGHT_TOP, HEADER_RIGHT_BOT, DOC_TITLE, DOC_SUBTITLE };
    const dynMap = buildDynMap(dynamics, SCALE);
    const tableMap = buildTableMap(tables);
    console.log("Dynamic Map:", dynMap);
    const userName = (Office.context && Office.context.displayName) || "User";

    const missingDVs = new Set();

    const meta = {
      version: selectedVersion, versionDetail, generatedBy: userName,
      generatedOn: new Date(), workbookName, logoBase64, logoSource,
      year: YEAR, scale: SCALE, missingDVs: missingDVs
    };

    const blob = await buildDoc(paragraphs, dynMap, tableMap, B, selectedVersion, meta);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Annual_Report_" + selectedVersion + "_" + Date.now() + ".docx";
    a.click();
    URL.revokeObjectURL(url);

    btn.disabled = false;
    btn.querySelector("span:last-child").innerText = "Generate Word Document";
    console.log("Missing DVs:", Array.from(missingDVs));
  } catch (err) {
    console.error("Generate error:", err);
    btn.disabled = false;
    btn.querySelector("span:last-child").innerText = "Generate Word Document";
  }
}

async function fetchImage(url) {
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const blob = await r.blob();
    return await new Promise((res, rej) => {
      const rd = new FileReader();
      rd.onloadend = () => { const x = rd.result; const i = x.indexOf(","); res(i >= 0 ? x.substring(i + 1) : x); };
      rd.onerror = rej;
      rd.readAsDataURL(blob);
    });
  } catch (e) { return null; }
}

function findColIdx(h, k) {
  for (let i = 0; i < h.length; i++) {
    if (String(h[i] || "").toLowerCase().indexOf(k.toLowerCase()) >= 0) return i;
  }
  return -1;
}

function findExactCol(h, n) {
  const t = String(n).trim().toUpperCase();
  for (let i = 0; i < h.length; i++) {
    if (String(h[i] || "").trim().toUpperCase() === t) return i;
  }
  return -1;
}

function cleanStr(s) {
  return String(s || "").replace(/\r/g, " ").replace(/\n/g, " ").replace(/\t/g, " ").replace(/`/g, "").replace(/\s+/g, " ").trim();
}

function normalizeId(s) {
  const m = /Para[_\s]*0*(\d+)/i.exec(String(s || ""));
  if (!m) return String(s || "").trim();
  return "Para_" + String(m[1]).padStart(2, "0");
}

function getScaleDivisor(scaleStr) {
  const s = String(scaleStr || "").toLowerCase().trim();
  if (s.indexOf("crore") >= 0) return 10000000;
  if (s.indexOf("lakh") >= 0 || s.indexOf("lac") >= 0) return 100000;
  if (s.indexOf("million") >= 0) return 1000000;
  if (s.indexOf("thousand") >= 0) return 1000;
  return 1;
}

function formatScaledNumber(num, divisor) {
  const scaled = num / divisor;
  if (Math.abs(scaled - Math.round(scaled)) < 0.001) {
    return Math.round(scaled).toLocaleString("en-IN");
  }
  return scaled.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function buildDynMap(rows, scaleStr) {
  const map = {};
  if (!rows || rows.length < 2) return map;
  const divisor = getScaleDivisor(scaleStr);

  for (let r = 1; r < rows.length; r++) {
    const id = String(rows[r][0] || "").trim().toUpperCase().replace(/[\s\-]/g, "_");
    if (!id) continue;

    const rawValue = rows[r][1];
    const type = String(rows[r][3] || "").trim().toLowerCase();
    let finalValue;

    if (type === "value" && typeof rawValue === "number" && !isNaN(rawValue)) {
      finalValue = formatScaledNumber(rawValue, divisor);
    } else {
      if (typeof rawValue === "number") {
        finalValue = String(rawValue);
      } else {
        finalValue = String(rawValue || "");
      }
    }

    map[id] = finalValue.replace(/[<>`]/g, "").trim();
  }
  return map;
}

function replaceDynWithMissing(text, map, missingSet) {
  const segments = [];
  const str = String(text || "");
  const re = /<\s*((?:DV|DT|DD)[\s_\-]*\d+)[^>]*>/gi;
  let lastIdx = 0;
  let m;
  while ((m = re.exec(str)) !== null) {
    if (m.index > lastIdx) {
      segments.push({ type: "text", text: str.substring(lastIdx, m.index) });
    }
    const id = m[1].toUpperCase().replace(/[\s\-]/g, "_");
    if (map[id] !== undefined && map[id] !== "") {
      segments.push({ type: "text", text: map[id] });
    } else {
      missingSet.add(id);
      segments.push({ type: "missing", text: "[MISSING: " + id + "]" });
    }
    lastIdx = m.index + m[0].length;
  }
  if (lastIdx < str.length) {
    segments.push({ type: "text", text: str.substring(lastIdx) });
  }
  return segments;
}

function buildTableMap(rows) {
  const map = {};
  if (!rows || rows.length === 0) return map;
  const lr = rows.length - 1;
  const lc = rows[0] ? rows[0].length : 0;
  let r = 1;
  while (r <= lr) {
    const tableId = String((rows[r] && rows[r][0]) || "").trim();
    const paraId = String((rows[r] && rows[r][1]) || "").trim();
    if (tableId.toLowerCase().startsWith("table_") && paraId) {
      const headerRow = r;
      let endRow = r + 1;
      while (endRow <= lr) {
        const nextTableId = String((rows[endRow] && rows[endRow][0]) || "").trim();
        if (nextTableId.toLowerCase().startsWith("table_")) break;
        let allEmpty = true;
        for (let c = 2; c < lc; c++) {
          if (String((rows[endRow] && rows[endRow][c]) || "").trim() !== "") { allEmpty = false; break; }
        }
        if (allEmpty) {
          if (endRow + 1 > lr) break;
          const nx = String((rows[endRow + 1] && rows[endRow + 1][0]) || "").trim();
          if (nx === "" || nx.toLowerCase().startsWith("table_")) break;
        }
        endRow++;
      }
      while (endRow > headerRow) {
        let allEmpty = true;
        for (let c = 2; c < lc; c++) {
          if (String((rows[endRow] && rows[endRow][c]) || "").trim() !== "") { allEmpty = false; break; }
        }
        if (!allEmpty) break;
        endRow--;
      }
      let nc = 0;
      for (let scanRow = headerRow; scanRow <= endRow; scanRow++) {
        for (let c = 2; c < lc; c++) {
          if (String((rows[scanRow] && rows[scanRow][c]) || "").trim() !== "") {
            if (c - 2 + 1 > nc) nc = c - 2 + 1;
          }
        }
      }
      if (nc < 2) nc = 2;
      const pidMatch = /Para[_\s]*0*(\d+)/i.exec(paraId);
      if (pidMatch) {
        const pid = "Para_" + String(pidMatch[1]).padStart(2, "0");
        if (!map[pid]) {
          map[pid] = { key: "T_" + r, title: tableId, headerRow: headerRow, endRow: endRow, numCols: nc, allRows: rows, colOffset: 2 };
        }
      }
      r = endRow + 1;
    } else { r++; }
  }
  return map;
}

async function buildDoc(paraRows, dynMap, tableMap, B, version, meta) {
 const d = docx;
  const usedT = {}, gPH = {};
  let tablesInc = 0;

  // ===== PASS 1: Build Para_ID → Number reference map =====
  const refMap = buildRefMap(paraRows, version);
  const pageHeader = makeHeader(d, B, meta.logoBase64);

  const titleSec = [
    new d.Paragraph({ children: [new d.TextRun({ text: B.DOC_TITLE, bold: true, color: BRAND_BLUE, font: "Arial", size: 52 })], spacing: { before: 200, after: 80 } }),
    new d.Paragraph({ children: [new d.TextRun({ text: B.DOC_SUBTITLE, color: TEXT_GREY, font: "Arial", size: 22 })], spacing: { after: 300 } })
  ];

  const hr = paraRows[0] || [];
  const cP = findColIdx(hr, "Para");
  const cL6 = findColIdx(hr, "Level 6");
  const cL = [findColIdx(hr, "Level 1"), findColIdx(hr, "Level 2"), findColIdx(hr, "Level 3"), findColIdx(hr, "Level 4"), findColIdx(hr, "Level 5")];
  const cV = findExactCol(hr, version);

  const sections = [];
  sections.push({
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1800, right: 1080, bottom: 1080, left: 1080 } }, type: d.SectionType.CONTINUOUS },
    headers: { default: pageHeader }, children: titleSec
  });

  let cur = [];
  function flush() {
    if (cur.length === 0) return;
    sections.push({
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1800, right: 1080, bottom: 1080, left: 1080 } }, column: { count: 2, space: 432 }, type: d.SectionType.CONTINUOUS },
      headers: { default: pageHeader }, children: cur
    });
    cur = [];
  }

  let pc = 0, sc = 0;
  const nf = numberer();
  const stk = ["", "", "", "", ""];

  for (let i = 1; i < paraRows.length; i++) {
    const row = paraRows[i];
    let emp = true;
    for (let c = 0; c < hr.length; c++) if (String(row[c] || "").trim() !== "") { emp = false; break; }
    if (emp) continue;
    if (cV >= 0 && String(row[cV] || "").trim().toUpperCase() !== "Y") { sc++; continue; }
    pc++;
    const pid = normalizeId(String(row[cP] || ""));
    const rt = [];
    for (let l = 1; l <= 5; l++) { const ci = cL[l - 1]; rt.push(ci < 0 ? "" : cleanStr(String(row[ci] || ""))); }
    let hf = -1;
    for (let l = 0; l < 5; l++) if (rt[l]) { hf = l; break; }
    const et = [];
    if (hf === -1) for (let l = 0; l < 5; l++) et.push("");
    else {
      for (let l = 0; l < hf; l++) et.push(stk[l]);
      et.push(rt[hf]); stk[hf] = rt[hf];
      for (let l = hf + 1; l < 5; l++) stk[l] = "";
      for (let l = hf + 1; l < 5; l++) { if (rt[l]) { et.push(rt[l]); stk[l] = rt[l]; } else et.push(""); }
    }
    const seen = {}, ft = [];
    for (let l = 0; l < 5; l++) {
      const t = et[l];
      if (!t) { ft.push(null); continue; }
      if (seen[t]) ft.push(null); else { seen[t] = true; ft.push(t); }
    }
    for (let l = 1; l <= 5; l++) {
      const ht = ft[l - 1];
      if (!ht || gPH[ht]) continue;
      gPH[ht] = true;
      const res = nf.assign(ht, l);
      cur.push(makeHeading(d, res.number + " " + ht, l));
    }
    if (cL6 >= 0) {
      let body = cleanStr(String(row[cL6] || ""));
      if (body) {
        body = replaceParaRefs(body, refMap);
        const segments = replaceDynWithMissing(body, dynMap, meta.missingDVs);
        cur.push(makeBodyWithSegments(d, segments));
      }
    }
    if (tableMap[pid] && !usedT[tableMap[pid].key]) {
      usedT[tableMap[pid].key] = true;
      tablesInc++;
      flush();
      const tc = [];
      tc.push(new d.Paragraph({ children: [new d.TextRun({ text: "" })], spacing: { before: 0, after: 0 }, keepNext: true, keepLines: true }));
      tc.push(makeTable(d, tableMap[pid]));
      tc.push(new d.Paragraph({ children: [new d.TextRun({ text: "" })], spacing: { after: 200 } }));
      sections.push({
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1800, right: 1080, bottom: 1080, left: 1080 } }, type: d.SectionType.CONTINUOUS },
        headers: { default: pageHeader }, children: tc
      });
    }
  }
  flush();

  sections.push({
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1800, right: 1080, bottom: 1080, left: 1080 } }, type: d.SectionType.NEXT_PAGE },
    headers: { default: pageHeader }, children: appendix(d, B, meta, pc, sc, tablesInc)
  });

  const doc = new d.Document({ styles: { default: { document: { run: { font: "Arial", size: 20 } } } }, sections: sections });
  return await d.Packer.toBlob(doc);
}

function numberer() {
  const c = [0, 0, 0, 0, 0];
  const sn = {};
  function tl(n) { let r = "", x = n; while (x > 0) { const m = (x - 1) % 26; r = String.fromCharCode(97 + m) + r; x = Math.floor((x - 1) / 26); } return r; }
  function tr(n) { const m = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii", "xviii", "xix", "xx"]; return n <= 20 ? m[n] : String(n); }
  function assign(t, l) {
    const li = l - 1, k = c.slice(0, li).join(".") + "|" + t;
    if (sn[k]) return { number: sn[k] };
    c[li]++;
    for (let i = li + 1; i < 5; i++) c[i] = 0;
    let n = "";
    if (l === 1) n = c[0] + ".";
    else if (l === 2) n = c[0] + "." + c[1];
    else if (l === 3) n = c[0] + "." + c[1] + "." + c[2];
    else if (l === 4) n = "(" + tl(c[3]) + ")";
    else if (l === 5) n = "(" + tr(c[4]) + ")";
    sn[k] = n;
    return { number: n };
  }
  return { assign };
}

function appendix(d, B, m, pc, sc, ti) {
  const ch = [];
  ch.push(new d.Paragraph({ children: [new d.TextRun({ text: "APPENDIX", bold: true, color: BRAND_BLUE, font: "Arial", size: 48 })], spacing: { before: 240, after: 80 } }));
  ch.push(new d.Paragraph({ children: [new d.TextRun({ text: "Document Generation Information", color: TEXT_GREY, font: "Arial", size: 22, italics: true })], spacing: { after: 320 } }));
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const dt = m.generatedOn;
  let h = dt.getHours();
  const mn = String(dt.getMinutes()).padStart(2, "0");
  const ap = h >= 12 ? "PM" : "AM";
  h = h % 12; if (h === 0) h = 12;
  const ds = String(dt.getDate()).padStart(2, "0") + " " + months[dt.getMonth()] + " " + dt.getFullYear() + ", " + h + ":" + mn + " " + ap;

  const missingArr = m.missingDVs ? Array.from(m.missingDVs) : [];
  const missingCount = missingArr.length;
  const missingList = missingCount > 0 ? missingArr.join(", ") : "None";

  ch.push(kvTable(d, [
    ["Document Title", B.DOC_TITLE], ["Subtitle", B.DOC_SUBTITLE],
    ["Company", B.COMPANY_NAME], ["CIN", B.COMPANY_CIN],
    ["Version", m.version], ["Version Detail", m.versionDetail],
    ["Year", m.year || "-"], ["Scale", m.scale || "-"],
    ["Generated By", m.generatedBy], ["Generated On", ds],
    ["Source File", m.workbookName], ["Logo Source", m.logoSource],
    ["Paragraphs Included", String(pc)], ["Paragraphs Skipped", String(sc)],
    ["Tables Included", String(ti)],
    ["Missing Dynamic Values (Count)", String(missingCount)],
    ["Missing Dynamic Values (List)", missingList]
  ]));
  return ch;
}

function kvTable(d, rows) {
  const b = { style: d.BorderStyle.SINGLE, size: 4, color: "C8D2E6" };
  const bs = { top: b, bottom: b, left: b, right: b };
  const tr = rows.map((kv, i) => {
    const bg = i % 2 === 1 ? "F4F7FC" : "FFFFFF";
    const isMissingRow = String(kv[0]).indexOf("Missing Dynamic Values") >= 0 && kv[1] !== "0" && kv[1] !== "None";
    const rowBg = isMissingRow ? "FEE2E2" : bg;
    const labelColor = isMissingRow ? "991B1B" : DARK_BLUE;
    return new d.TableRow({ cantSplit: true, children: [
      new d.TableCell({ width: { size: 3400, type: d.WidthType.DXA }, borders: bs, shading: { type: d.ShadingType.CLEAR, fill: rowBg }, margins: { top: 100, bottom: 100, left: 140, right: 100 }, children: [new d.Paragraph({ children: [new d.TextRun({ text: kv[0], bold: true, color: labelColor, font: "Arial", size: 18 })] })] }),
      new d.TableCell({ width: { size: 5960, type: d.WidthType.DXA }, borders: bs, shading: { type: d.ShadingType.CLEAR, fill: rowBg }, margins: { top: 100, bottom: 100, left: 140, right: 100 }, children: [new d.Paragraph({ children: [new d.TextRun({ text: kv[1], color: isMissingRow ? "991B1B" : TEXT_DARK, font: "Arial", size: 18, bold: isMissingRow })] })] })
    ]});
  });
  return new d.Table({ width: { size: 9360, type: d.WidthType.DXA }, columnWidths: [3400, 5960], rows: tr });
}

function makeHeader(d, B, logoBase64) {
  const headerChildren = [];
  if (logoBase64) {
    try {
      const bin = atob(logoBase64);
      const bt = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bt[i] = bin.charCodeAt(i);
      headerChildren.push(new d.Paragraph({
        alignment: d.AlignmentType.RIGHT,
        children: [new d.ImageRun({ data: bt, transformation: { width: 140, height: 90 } })],
        spacing: { before: 0, after: 80 }
      }));
    } catch (e) { console.warn("Logo error:", e); }
  }
  const nb = { top: { style: d.BorderStyle.NONE }, bottom: { style: d.BorderStyle.NONE }, left: { style: d.BorderStyle.NONE }, right: { style: d.BorderStyle.NONE } };
  headerChildren.push(new d.Table({
    width: { size: 9360, type: d.WidthType.DXA },
    columnWidths: [4680, 4680],
    borders: {
      top: { style: d.BorderStyle.NONE }, bottom: { style: d.BorderStyle.SINGLE, size: 6, color: BORDER_LIGHT },
      left: { style: d.BorderStyle.NONE }, right: { style: d.BorderStyle.NONE },
      insideHorizontal: { style: d.BorderStyle.NONE }, insideVertical: { style: d.BorderStyle.NONE }
    },
    rows: [new d.TableRow({
      children: [
        new d.TableCell({
          width: { size: 4680, type: d.WidthType.DXA }, borders: nb, verticalAlign: d.VerticalAlign.CENTER,
          children: [
            new d.Paragraph({ children: [new d.TextRun({ text: B.COMPANY_NAME, bold: true, color: BRAND_BLUE, font: "Arial", size: 22 })] }),
            new d.Paragraph({ children: [new d.TextRun({ text: B.COMPANY_CIN, color: TEXT_GREY, font: "Arial", size: 16 })] })
          ]
        }),
        new d.TableCell({
          width: { size: 4680, type: d.WidthType.DXA }, borders: nb, verticalAlign: d.VerticalAlign.CENTER,
          children: [
            new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: B.HEADER_RIGHT_TOP, bold: true, color: BRAND_BLUE, font: "Arial", size: 22 })] }),
            new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: B.HEADER_RIGHT_BOT, color: TEXT_GREY, font: "Arial", size: 16 })] })
          ]
        })
      ]
    })]
  }));
  return new d.Header({ children: headerChildren });
}

function makeHeading(d, t, l) {
  let s, b, it, c, il;
  switch (l) {
    case 1: s = 26; b = true; it = false; c = BRAND_BLUE; il = 0; break;
    case 2: s = 22; b = true; it = false; c = TEXT_DARK; il = 0; break;
    case 3: s = 21; b = true; it = true; c = TEXT_DARK; il = 100; break;
    case 4: s = 20; b = true; it = true; c = "404040"; il = 200; break;
    default: s = 20; b = true; it = true; c = "555555"; il = 300;
  }
  return new d.Paragraph({ children: [new d.TextRun({ text: t, bold: b, italics: it, color: c, font: "Arial", size: s })], spacing: { before: l === 1 ? 280 : 200, after: l === 1 ? 140 : 100 }, indent: { left: il }, keepNext: true });
}

function makeBody(d, t) {
  return new d.Paragraph({ alignment: d.AlignmentType.JUSTIFIED, children: [new d.TextRun({ text: t, color: TEXT_DARK, font: "Arial", size: 20 })], spacing: { after: 120 } });
}

function makeBodyWithSegments(d, segments) {
  const runs = segments.map(seg => {
    if (seg.type === "missing") {
      return new d.TextRun({
        text: seg.text,
        color: "CC0000",
        bold: true,
        font: "Arial",
        size: 20,
        highlight: "yellow"
      });
    }
    return new d.TextRun({ text: seg.text, color: TEXT_DARK, font: "Arial", size: 20 });
  });
  return new d.Paragraph({
    alignment: d.AlignmentType.JUSTIFIED,
    children: runs,
    spacing: { after: 120 }
  });
}

function makeTitleBar(d, t) {
  return new d.Paragraph({ children: [new d.TextRun({ text: t, bold: true, color: "FFFFFF", font: "Arial", size: 20 })], shading: { type: d.ShadingType.CLEAR, fill: TABLE_HEADER_BG }, spacing: { before: 160, after: 0 }, indent: { left: 80 }, keepNext: true, keepLines: true });
}

function makeTable(d, info) {
  const nr = info.endRow - info.headerRow + 1;
  const nc = info.numCols;
  const colOffset = info.colOffset || 0;
  if (nr < 1 || nc < 1) return new d.Paragraph({ children: [new d.TextRun({ text: "" })] });
  const tw = 9360;
  const fw = Math.floor(tw * 0.32);
  const ow = nc > 1 ? Math.floor((tw - fw) / (nc - 1)) : (tw - fw);
  const cw = [fw];
  for (let c = 1; c < nc; c++) cw.push(ow);
  const b = { style: d.BorderStyle.SINGLE, size: 4, color: "C8D2E6" };
  const bs = { top: b, bottom: b, left: b, right: b };
  const rows = [];
  for (let r = 0; r < nr; r++) {
    const sr = info.headerRow + r;
    const rd = info.allRows[sr] || [];
    const ih = r === 0;
    const fc = String(rd[colOffset] || "").trim();
    let hl = false, sh = false;
    if (!ih) {
      const f = fc.toLowerCase();
      if (f.startsWith("as at") || f.startsWith("at march") || f.startsWith("balance as") || f.startsWith("total")) hl = true;
      let e = 0;
      for (let c = 1; c < nc; c++) if (!rd[colOffset + c] || String(rd[colOffset + c]).trim() === "") e++;
      if (e === nc - 1 && fc) sh = true;
    }
    const cells = [];
    for (let c = 0; c < nc; c++) {
      const actualCol = colOffset + c;
      const ct = String(rd[actualCol] != null ? rd[actualCol] : "").trim();
      let bg = "FFFFFF", tc = TEXT_DARK, fs = 14, iB = false, iI = false;
      let al = c === 0 ? d.AlignmentType.LEFT : d.AlignmentType.RIGHT;
      if (ih) { bg = TABLE_HEADER_BG; tc = "FFFFFF"; iB = true; al = d.AlignmentType.CENTER; }
      else if (sh) { bg = LIGHT_BLUE; tc = DARK_BLUE; iB = true; iI = true; al = d.AlignmentType.LEFT; }
      else if (hl) { bg = LIGHT_BLUE; tc = DARK_BLUE; iB = true; }
      cells.push(new d.TableCell({
        width: { size: c === 0 ? fw : ow, type: d.WidthType.DXA },
        borders: bs, shading: { type: d.ShadingType.CLEAR, fill: bg },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new d.Paragraph({ alignment: al, children: [new d.TextRun({ text: ct, bold: iB, italics: iI, color: tc, font: "Arial", size: fs })] })]
      }));
    }
    rows.push(new d.TableRow({ children: cells, cantSplit: true }));
  }
  return new d.Table({ width: { size: tw, type: d.WidthType.DXA }, columnWidths: cw, rows: rows });
}

window.generate = generate;
window.downloadTemplate = downloadTemplate;
window.handleFileSelect = handleFileSelect;
window.processPDF = processPDF;
function doLogout() {
  localStorage.removeItem("iocfo_login");
  const overlay = document.getElementById("loginOverlay");
  const main = document.getElementById("mainContent");
  if (overlay) overlay.style.display = "flex";
  if (main) main.style.display = "none";
  document.getElementById("loginUsername").value = "";
  document.getElementById("loginPassword").value = "";
  document.getElementById("rememberMe").checked = false;
  hideLoginError();
  const btn = document.getElementById("loginBtn");
  if (btn) { btn.disabled = false; btn.innerText = "Sign In"; }
}

window.doLogout = doLogout;
function updateUserUI(username) {
  const initial = document.getElementById("userInitial");
  const loggedInUser = document.getElementById("loggedInUser");
  if (initial) initial.innerText = username.charAt(0).toUpperCase();
  if (loggedInUser) loggedInUser.innerText = username;
}

function toggleUserMenu() {
  const menu = document.getElementById("userMenu");
  const chevron = document.getElementById("chevronIcon");
  if (menu) {
    if (menu.style.display === "none" || menu.style.display === "") {
      menu.style.display = "block";
      if (chevron) chevron.innerText = "▲";
    } else {
      menu.style.display = "none";
      if (chevron) chevron.innerText = "▼";
    }
  }
}

document.addEventListener("click", function(e) {
  const menu = document.getElementById("userMenu");
  const btn = document.getElementById("userIconBtn");
  if (menu && btn && !btn.contains(e.target) && !menu.contains(e.target)) {
    menu.style.display = "none";
    const chevron = document.getElementById("chevronIcon");
    if (chevron) chevron.innerText = "▼";
  }
});

window.toggleUserMenu = toggleUserMenu;
function updateUserUI(username) {
  const initial = document.getElementById("userInitial");
  const loggedInUser = document.getElementById("loggedInUser");
  if (initial) initial.innerText = username.charAt(0).toUpperCase();
  if (loggedInUser) loggedInUser.innerText = username;
}

function toggleUserMenu() {
  const menu = document.getElementById("userMenu");
  const chevron = document.getElementById("chevronIcon");
  if (menu) {
    if (menu.style.display === "none" || menu.style.display === "") {
      menu.style.display = "block";
      if (chevron) chevron.innerText = "▲";
    } else {
      menu.style.display = "none";
      if (chevron) chevron.innerText = "▼";
    }
  }
}

document.addEventListener("click", function(e) {
  const menu = document.getElementById("userMenu");
  const btn = document.getElementById("userIconBtn");
  if (menu && btn && !btn.contains(e.target) && !menu.contains(e.target)) {
    menu.style.display = "none";
    const chevron = document.getElementById("chevronIcon");
    if (chevron) chevron.innerText = "▼";
  }
});

window.toggleUserMenu = toggleUserMenu;
window.updateUserUI = updateUserUI;
// ============================================================
// CROSS-REFERENCE: Build Para_ID → heading number map
// ============================================================
function buildRefMap(paraRows, version) {
  const refMap = {};
  if (!paraRows || paraRows.length < 2) return refMap;

  const hr = paraRows[0] || [];
  const cP = findColIdx(hr, "Para");
  const cL = [findColIdx(hr, "Level 1"), findColIdx(hr, "Level 2"), findColIdx(hr, "Level 3"), findColIdx(hr, "Level 4"), findColIdx(hr, "Level 5")];
  const cV = findExactCol(hr, version);

  const nf = numberer();
  const stk = ["", "", "", "", ""];
  const gPH = {};
  let lastNumber = "";

  for (let i = 1; i < paraRows.length; i++) {
    const row = paraRows[i];
    let emp = true;
    for (let c = 0; c < hr.length; c++) if (String(row[c] || "").trim() !== "") { emp = false; break; }
    if (emp) continue;
    if (cV >= 0 && String(row[cV] || "").trim().toUpperCase() !== "Y") continue;

    const pid = normalizeId(String(row[cP] || ""));

    const rt = [];
    for (let l = 1; l <= 5; l++) { const ci = cL[l - 1]; rt.push(ci < 0 ? "" : cleanStr(String(row[ci] || ""))); }
    let hf = -1;
    for (let l = 0; l < 5; l++) if (rt[l]) { hf = l; break; }
    const et = [];
    if (hf === -1) { for (let l = 0; l < 5; l++) et.push(""); }
    else {
      for (let l = 0; l < hf; l++) et.push(stk[l]);
      et.push(rt[hf]); stk[hf] = rt[hf];
      for (let l = hf + 1; l < 5; l++) stk[l] = "";
      for (let l = hf + 1; l < 5; l++) { if (rt[l]) { et.push(rt[l]); stk[l] = rt[l]; } else et.push(""); }
    }
    const seen = {}, ft = [];
    for (let l = 0; l < 5; l++) {
      const t = et[l];
      if (!t) { ft.push(null); continue; }
      if (seen[t]) ft.push(null); else { seen[t] = true; ft.push(t); }
    }

    const fullPath = [];
    for (let l = 1; l <= 5; l++) {
      const ht = ft[l - 1];
      if (!ht) continue;
      if (gPH[ht]) {
        fullPath[l - 1] = gPH[ht];
      } else {
        const res = nf.assign(ht, l);
        gPH[ht] = res.number;
        fullPath[l - 1] = res.number;
      }
    }

    // Build the active full path from stack tracking
    let paraNumber = "";
    const activePath = [];
    for (let l = 0; l < 5; l++) {
      if (stk[l] && gPH[stk[l]]) {
        activePath.push(gPH[stk[l]].replace(/\.$/, ""));
      }
    }
    if (activePath.length > 0) {
      paraNumber = activePath.join("");
    }

    if (!paraNumber) paraNumber = lastNumber;
    else lastNumber = paraNumber;

    if (pid && paraNumber) {
      refMap[pid] = paraNumber;
    }
  return refMap;
}

// ============================================================
// Replace "Para_XX" references in body text with their number
// ============================================================
function replaceParaRefs(text, refMap) {
  let str = String(text || "");
  str = str.replace(/"?\s*(Para[_\s]*0*\d+)\s*"?/gi, function(match, paraId) {
    const normalized = normalizeId(paraId);
    if (refMap[normalized]) {
      return refMap[normalized];
    }
    return match;
  });
  return str;
}