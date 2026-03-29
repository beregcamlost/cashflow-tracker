/******************************************************************************
 * ██████╗  █████╗ ███████╗██╗  ██╗███████╗██╗      ██████╗ ██╗    ██╗
 * ██╔════╝██╔══██╗██╔════╝██║  ██║██╔════╝██║     ██╔═══██╗██║    ██║
 * ██║     ███████║███████╗███████║█████╗  ██║     ██║   ██║██║ █╗ ██║
 * ██║     ██╔══██║╚════██║██╔══██║██╔══╝  ██║     ██║   ██║██║███╗██║
 * ╚██████╗██║  ██║███████║██║  ██║██║     ███████╗╚██████╔╝╚███╔███╔╝
 *  ╚═════╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚═╝     ╚══════╝ ╚═════╝  ╚══╝╚══╝
 *
 * Cashflow — Bank Statement Import & Financial Dashboard for Google Sheets
 * "Midnight Finance" theme · Preview → Confirm workflow
 *
 * Tabs managed: DASHBOARD · CONFIG · CATEGORIAS · RESUMEN_CATEGORIAS ·
 *               MOV_CC_NACIONAL · MOV_CC_INTL · MOV_BANCO · CUOTAS · PREVIEW
 ******************************************************************************/

/******************************************************************************
 * §1  CONFIG & THEME
 ******************************************************************************/

const PROP_DROP_FOLDER = "DROP_FOLDER_ID";
const PROP_GEMINI_KEY = "GEMINI_API_KEY";
const GEMINI_MODEL = "gemini-2.0-flash";

/**
 * Returns the configured Drop Folder ID from Script Properties.
 * Prompts the user via dialog if not yet configured.
 */
function getDropFolderId_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_DROP_FOLDER);
  if (id) return id;

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Drop Folder Setup",
    "Paste your Google Drive folder URL or ID.\n\n" +
    "Example: https://drive.google.com/drive/folders/1Ab2Cd3... or just the ID",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Drop folder ID is required. Set it in File > Project properties > Script properties.");
  }

  const input = resp.getResponseText().trim();
  const match = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  id = match ? match[1] : input;

  props.setProperty(PROP_DROP_FOLDER, id);
  return id;
}

function getGeminiApiKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(PROP_GEMINI_KEY);
  if (key) return key;

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Gemini API Key Setup",
    "Paste your Gemini API key.\n\nGet one at: https://aistudio.google.com/apikey",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Gemini API key is required for AI-powered parsing.");
  }

  key = resp.getResponseText().trim();
  if (!key) throw new Error("API key cannot be empty.");

  props.setProperty(PROP_GEMINI_KEY, key);
  return key;
}

function getGeminiApiKeyIfSet_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_GEMINI_KEY) || null;
}

function ensureChileanLocale_(ss) {
  if (ss.getSpreadsheetLocale() !== "es_CL") {
    ss.setSpreadsheetLocale("es_CL");
  }
}

const NAC_REGEX   = /^MovimientosNoFacturadosNacionales_.*/i;
const INTL_REGEX  = /^MovimientosNoFacturadosInternacionales_.*/i;
const BANCO_REGEX = /^Ultimos_Movimientos_Cuenta_Corriente_.*/i;
const NAC_FACT_REGEX  = /^MovimientosFacturadosNacionales_.*/i;
const INTL_FACT_REGEX = /^MovimientosFacturadosInternacionales_.*/i;

const TAB_NAC        = "🇨🇱 MOV_CC_NACIONAL";
const TAB_INTL       = "🌎 MOV_CC_INTL";
const DASHBOARD_TAB  = "📊 DASHBOARD";
const PREVIEW_TAB    = "👁️ PREVIEW";
const TAB_CONFIG     = "⚙️ CONFIG";
const TAB_CATEGORIAS = "🏷️ CATEGORIAS";
const TAB_RESUMEN_CAT = "📋 RESUMEN_CATEGORIAS";
const TAB_CUOTAS     = "📅 CUOTAS";
const TAB_BANCO      = "🏦 MOV_BANCO";

/** Legacy tabs — auto-deleted on applyTheme() */
const LEGACY_TABS = [
  "PREVIEW_SUMMARY", "PREVIEW_NAC", "PREVIEW_INTL",
  "RESUMEN", "PROYECCION", "BURN_CURVE",
  "MOV_NAC_AUTO", "MOV_INTL_AUTO", "CHART_DATA",
];

/** Old tab names → new emoji names (for migration) */
const TAB_RENAMES = {
  "DASHBOARD":          DASHBOARD_TAB,
  "CONFIG":             TAB_CONFIG,
  "CATEGORIAS":         TAB_CATEGORIAS,
  "RESUMEN_CATEGORIAS": TAB_RESUMEN_CAT,
  "MOV_CC_NACIONAL":    TAB_NAC,
  "MOV_CC_INTL":        TAB_INTL,
  "MOV_BANCO":          TAB_BANCO,
  "CUOTAS":             TAB_CUOTAS,
  "PREVIEW":            PREVIEW_TAB,
};

const FONT_FAMILY = "Inter";

const FMT_CLP  = "#,##0";
const FMT_USD  = "$#,##0.00";
const FMT_DATE = "dd/MM/yyyy";
const FMT_PCT  = "0.0%";

const MAX_HISTORY = 15;

/** Installment pattern: CC 03-12 or CF 01-06 */
const CUOTA_REGEX = /(?:CC|CF)\s+(\d{2})-(\d{2})/i;

/** Midnight Finance color palette */
const THEME = {
  PRIMARY:        "#1B2A4A",
  SECONDARY:      "#4A6FA5",
  ACCENT:         "#2EC4B6",
  BG_ALT_ROW:     "#F0F4F8",
  BG_DASHBOARD:   "#F8FAFC",
  STATUS_SUCCESS: "#059669",
  STATUS_WARNING: "#D97706",
  STATUS_ERROR:   "#DC2626",
  STATUS_INFO:    "#2563EB",
  POSITIVE_BG:    "#ECFDF5",
  NEGATIVE_BG:    "#FEF2F2",
  POSITIVE_TEXT:  "#065F46",
  NEGATIVE_TEXT:  "#991B1B",
  BORDER_LIGHT:   "#E2E8F0",
  PREVIEW_ACCENT: "#F59E0B",
  PREVIEW_ALT:    "#FFFBEB",
  MUTED_TEXT:     "#64748B",
  WHITE:          "#FFFFFF",
  BLACK:          "#000000",
};

/******************************************************************************
 * §2  MENU
 ******************************************************************************/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("💰 Management")
    .addItem("Preview import",       "previewLatestBankXls")
    .addItem("Confirm import",       "confirmImportFromPreview")
    .addItem("Cancel preview",       "cancelPreview")
    .addSeparator()
    .addItem("Refresh calculations", "refreshCalculations")
    .addItem("Apply theme",          "applyTheme")
    .addSeparator()
    .addSubMenu(ui.createMenu("⚙️ Settings")
      .addItem("Set Drop Folder",    "setDropFolder")
      .addItem("Set Gemini API Key", "setGeminiApiKey"))
    .addToUi();
}

/******************************************************************************
 * §3  ENTRY POINTS (wrapped in safeRun_)
 ******************************************************************************/

function previewLatestBankXls() {
  safeRun_("Preview Import", doPreview_);
}

function confirmImportFromPreview() {
  safeRun_("Confirm Import", doConfirm_);
}

function cancelPreview() {
  safeRun_("Cancel Preview", doCancel_);
}

function refreshCalculations() {
  safeRun_("Refresh Calculations", doRefreshCalculations_);
}

function applyTheme() {
  safeRun_("Apply Theme", doApplyTheme_);
}

function setGeminiApiKey() {
  safeRun_("Set Gemini API Key", function(ss) {
    PropertiesService.getScriptProperties().deleteProperty(PROP_GEMINI_KEY);
    getGeminiApiKey_();
    toast_(ss, "Gemini API key saved!", 3);
  });
}

function setDropFolder() {
  safeRun_("Set Drop Folder", function(ss) {
    PropertiesService.getScriptProperties().deleteProperty(PROP_DROP_FOLDER);
    getDropFolderId_();
    toast_(ss, "Drop folder ID saved!", 3);
  });
}

/******************************************************************************
 * §4  ERROR HANDLING & UX HELPERS
 ******************************************************************************/

/**
 * Wraps a menu-triggered function in try/catch with toast + alert UX.
 * @param {string} label  Human-readable action name for toasts/alerts
 * @param {Function} fn   The actual work function
 */
function safeRun_(label, fn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Starting…", label, 30);
    fn(ss);
    ss.toast("Done!", label, 4);
  } catch (e) {
    ss.toast("Error — see alert", label, 5);
    SpreadsheetApp.getUi().alert(
      label + " — Error",
      e.message || String(e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    const dash = ss.getSheetByName(DASHBOARD_TAB);
    if (dash) {
      updateDashboardStatus_(dash, "ERROR", e.message);
    }
    Logger.log("[%s] ERROR: %s", label, e.message);
  }
}

/** Convenience toast wrapper. */
function toast_(ss, msg) {
  ss.toast(msg, "Management", 30);
}

/******************************************************************************
 * §5  CONFIG & CATEGORY READERS
 ******************************************************************************/

/**
 * Read CONFIG tab → returns {salario, housing, familia, usdClp, claudePlan}.
 * Scans column A for parameter names, column B for values.
 */
function readConfig_(ss) {
  const defaults = { salario: 0, housing: 0, familia: 0, usdClp: 0, claudePlan: 0 };
  const sh = ss.getSheetByName(TAB_CONFIG);
  if (!sh || sh.getLastRow() < 1) return defaults;

  SpreadsheetApp.flush();
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const cfg = Object.assign({}, defaults);
  const map = {
    salario:     "salario",
    housing:     "housing",
    vivienda:    "housing",
    familia:     "familia",
    usd_clp:     "usdClp",
    usdclp:      "usdClp",
    tipo_cambio: "usdClp",
    claude_plan: "claudePlan",
    claudeplan:  "claudePlan",
  };

  for (const [label, val] of data) {
    const key = String(label || "").toLowerCase().replace(/[\s/]/g, "_").trim();
    const mapped = map[key];
    if (mapped) {
      const num = typeof val === "number" ? val : parseFloat(String(val).replace(/[^\d.-]/g, ""));
      if (isFinite(num)) cfg[mapped] = num;
    }
  }

  // Retry for IMPORTRANGE cells that may not have resolved yet
  if (cfg.housing === 0) {
    Utilities.sleep(2000);
    const retryData = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
    for (const [label, val] of retryData) {
      const key = String(label || "").toLowerCase().replace(/[\s/]/g, "_").trim();
      if (key === "housing" || key === "vivienda") {
        const num = typeof val === "number" ? val : parseFloat(String(val).replace(/[^\d.-]/g, ""));
        if (isFinite(num) && num !== 0) cfg.housing = num;
        break;
      }
    }
  }

  return cfg;
}

/**
 * Read CATEGORIAS tab → array of {tipo, keyword, categoria}.
 * Expects columns: Tipo | Keyword | Categoria (row 1 = header, data from row 2).
 */
function readCategories_(ss) {
  const sh = ss.getSheetByName(TAB_CATEGORIAS);
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const categories = [];
  for (const [tipo, keyword, categoria] of data) {
    if (!keyword) continue;
    categories.push({
      tipo:      String(tipo || "").toUpperCase().trim(),
      keyword:   String(keyword).trim(),
      categoria: String(categoria || "Other").trim(),
    });
  }
  return categories;
}

/**
 * Match description against CATEGORIAS keywords (case-insensitive contains).
 * @param {string} desc       Movement description
 * @param {string} tipo       "NAC" or "INTL"
 * @param {Array}  categories From readCategories_
 * @returns {string} Category name or "Other"
 */
function categorizeMovement_(desc, tipo, categories) {
  const d = String(desc).toLowerCase();
  for (const cat of categories) {
    if (cat.tipo && cat.tipo !== "ALL" && cat.tipo !== tipo) continue;
    if (d.includes(cat.keyword.toLowerCase())) return cat.categoria;
  }
  return "Other";
}

/******************************************************************************
 * §6  THEME ENGINE — Formatting Functions
 ******************************************************************************/

/**
 * Master theme application — idempotent, safe to re-run.
 * Deletes legacy tabs, formats remaining 8.
 */
function doApplyTheme_(ss) {
  ensureAllTabs_(ss);
  ensureChileanLocale_(ss);
  toast_(ss, "Cleaning up legacy tabs…");
  removeLegacyTabs_(ss);

  toast_(ss, "Setting tab colors…");
  setTabColors_(ss);

  toast_(ss, "Formatting data tabs…");
  const nacSheet = ss.getSheetByName(TAB_NAC);
  if (nacSheet && nacSheet.getLastRow() > 0) {
    formatDataTab_(nacSheet, "CLP", false);
  }
  const intlSheet = ss.getSheetByName(TAB_INTL);
  if (intlSheet && intlSheet.getLastRow() > 0) {
    formatDataTab_(intlSheet, "USD", false);
  }
  const bancoSheet = ss.getSheetByName(TAB_BANCO);
  if (bancoSheet && bancoSheet.getLastRow() > 0) {
    formatDataTab_(bancoSheet, "CLP", false);
  }

  toast_(ss, "Formatting dashboard…");
  const dash = ss.getSheetByName(DASHBOARD_TAB);
  if (dash) {
    formatDashboardChrome_(dash);
  }
  renderHistory_(ss);

  const prev = ss.getSheetByName(PREVIEW_TAB);
  if (prev && prev.getLastRow() > 0) {
    toast_(ss, "Formatting preview tab…");
  }
}

/** Delete legacy tabs if they exist. */
function removeLegacyTabs_(ss) {
  for (const name of LEGACY_TABS) {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  }
}

/** Assign tab colors. */
function setTabColors_(ss) {
  const map = {
    [DASHBOARD_TAB]:   THEME.ACCENT,
    [TAB_NAC]:         THEME.PRIMARY,
    [TAB_INTL]:        THEME.SECONDARY,
    [TAB_BANCO]:       THEME.STATUS_INFO,
    [PREVIEW_TAB]:     THEME.PREVIEW_ACCENT,
    [TAB_CUOTAS]:      THEME.STATUS_WARNING,
    [TAB_RESUMEN_CAT]: THEME.MUTED_TEXT,
  };
  for (const [name, color] of Object.entries(map)) {
    const sh = ss.getSheetByName(name);
    if (sh) sh.setTabColor(color);
  }
}

/**
 * Style a header row (row 1).
 * @param {Sheet} sheet
 * @param {number} numCols
 * @param {boolean} isPreview  Use amber instead of navy
 */
function styleHeaderRow_(sheet, numCols, isPreview) {
  const hdr = sheet.getRange(1, 1, 1, numCols);
  hdr.setBackground(isPreview ? THEME.PREVIEW_ACCENT : THEME.PRIMARY)
     .setFontColor(THEME.WHITE)
     .setFontWeight("bold")
     .setFontSize(11)
     .setFontFamily(FONT_FAMILY)
     .setHorizontalAlignment("center")
     .setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  hdr.setBorder(null, null, true, null, null, null, THEME.BORDER_LIGHT, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Apply alternating row backgrounds.
 */
function applyZebraStripes_(sheet, startRow, numRows, numCols, isPreview) {
  if (numRows < 1) return;
  const altColor = isPreview ? THEME.PREVIEW_ALT : THEME.BG_ALT_ROW;
  for (let i = 0; i < numRows; i++) {
    if (i % 2 === 1) {
      sheet.getRange(startRow + i, 1, 1, numCols).setBackground(altColor);
    } else {
      sheet.getRange(startRow + i, 1, 1, numCols).setBackground(THEME.WHITE);
    }
  }
}

/**
 * Conditional formatting on amount column: green for positive, red for negative.
 */
function applyAmountConditionalFormatting_(sheet, amountCol, startRow, lastRow) {
  if (lastRow < startRow) return;
  const range = sheet.getRange(startRow, amountCol, lastRow - startRow + 1, 1);

  const rules = sheet.getConditionalFormatRules();

  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground(THEME.POSITIVE_BG)
    .setFontColor(THEME.POSITIVE_TEXT)
    .setRanges([range])
    .build();

  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground(THEME.NEGATIVE_BG)
    .setFontColor(THEME.NEGATIVE_TEXT)
    .setRanges([range])
    .build();

  rules.push(positiveRule, negativeRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Full formatting pass for a data tab (NAC or INTL).
 * Handles 4-column NAC layout (Fecha, Descripcion, Monto, Categoria)
 * and 5-column INTL layout (+ Monto_CLP_est).
 * @param {Sheet} sheet
 * @param {string} currency  "CLP" or "USD"
 * @param {boolean} isPreview
 */
function formatDataTab_(sheet, currency, isPreview) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  // Detect column count from actual data
  const lastCol = sheet.getLastColumn();
  const numCols = Math.max(lastCol, 3);

  // Clear existing conditional format rules
  sheet.setConditionalFormatRules([]);

  // Header
  styleHeaderRow_(sheet, numCols, isPreview);

  // Auto-fit column widths to content
  sheet.autoResizeColumns(1, numCols);

  // Data rows font
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, numCols);
    dataRange.setFontFamily(FONT_FAMILY)
             .setFontSize(10)
             .setVerticalAlignment("middle");

    // Column A: date format, centered
    sheet.getRange(2, 1, lastRow - 1, 1)
         .setNumberFormat(FMT_DATE)
         .setHorizontalAlignment("center");

    // Column B: left-aligned
    sheet.getRange(2, 2, lastRow - 1, 1)
         .setHorizontalAlignment("left");

    // Column C: amount format, right-aligned
    sheet.getRange(2, 3, lastRow - 1, 1)
         .setNumberFormat(currency === "CLP" ? FMT_CLP : FMT_USD)
         .setHorizontalAlignment("right");

    // Column D: Tipo, left-aligned
    if (numCols >= 4) {
      sheet.getRange(2, 4, lastRow - 1, 1)
           .setHorizontalAlignment("left");
    }

    // Column E: Categoria, left-aligned
    if (numCols >= 5) {
      sheet.getRange(2, 5, lastRow - 1, 1)
           .setHorizontalAlignment("left");
    }

    // Column F: Monto_CLP_est (INTL only), right-aligned
    if (numCols >= 6) {
      sheet.getRange(2, 6, lastRow - 1, 1)
           .setNumberFormat(FMT_CLP)
           .setHorizontalAlignment("right");
    }

    // Zebra stripes
    applyZebraStripes_(sheet, 2, lastRow - 1, numCols, isPreview);

    // Conditional formatting on amount column
    applyAmountConditionalFormatting_(sheet, 3, 2, lastRow);
  }
}

/******************************************************************************
 * §7  DASHBOARD
 ******************************************************************************/

/**
 * Writes the static chrome of the dashboard (title bar, section headers, labels).
 * Expanded layout:
 *   Rows 1-7:   Title, status card, file info
 *   Rows 9-13:  Financial overview
 *   Rows 15-18: Savings & indicators
 *   Rows 20-24: CC summary (NAC/INTL cards)
 *   Rows 25-27: CC payment estimates
 *   Rows 29-31: Installments summary
 *   Rows 33-34: How to import
 */
function formatDashboardChrome_(dash) {
  // Ensure minimum size
  if (dash.getMaxColumns() < 8) dash.insertColumnsAfter(dash.getMaxColumns(), 8 - dash.getMaxColumns());

  // Clear all existing content, merges, and formulas before re-applying layout
  const maxRow = Math.max(dash.getMaxRows(), 50);
  dash.getRange(1, 1, maxRow, 8).breakApart().clearContent().clearFormat();

  // Background
  dash.getRange(1, 1, maxRow, 8).setBackground(THEME.BG_DASHBOARD);

  // ── Row 1: Title bar ──────────────────────────────────────────────────
  dash.getRange("A1:H1").merge()
      .setValue("CASHFLOW DASHBOARD")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontSize(18)
      .setFontWeight("bold")
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  dash.setRowHeight(1, 48);

  // Row 2: Teal accent stripe
  dash.getRange("A2:H2").merge()
      .setValue("")
      .setBackground(THEME.ACCENT);
  dash.setRowHeight(2, 4);

  // Row 3: spacer
  dash.setRowHeight(3, 10);

  // ── Rows 4-7: STATUS + FILE INFO ──────────────────────────────────────
  dash.getRange("A4:D4").merge()
      .setValue("STATUS")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  dash.getRange("E4:H4").merge()
      .setValue("FILE INFO")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  // Status labels (rows 5-7, col A)
  const statusLabels = ["Status", "Last updated", "Hint"];
  for (let i = 0; i < statusLabels.length; i++) {
    dash.getRange(5 + i, 1).setValue(statusLabels[i])
        .setFontColor(THEME.MUTED_TEXT)
        .setFontWeight("bold")
        .setFontSize(10)
        .setFontFamily(FONT_FAMILY);
  }

  // File info labels (rows 5-7, col E)
  dash.getRange(5, 5).setValue("NAC file")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(6, 5).setValue("INTL file")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(7, 5).setValue("Banco file")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);

  // Row 8: spacer
  dash.setRowHeight(8, 10);

  // ── Rows 9-13: FINANCIAL OVERVIEW ─────────────────────────────────────
  dash.getRange("A9:H9").merge()
      .setValue("FINANCIAL OVERVIEW")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  const finLabelsLeft  = ["Income (Salario)", "CC Nacional", "Total Expenses", "FSI (margin/income)"];
  const finLabelsRight = ["Housing", "CC Intl (CLP)", "Available Margin", "Stability"];
  for (let i = 0; i < finLabelsLeft.length; i++) {
    dash.getRange(10 + i, 1).setValue(finLabelsLeft[i])
        .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
    dash.getRange(10 + i, 5).setValue(finLabelsRight[i])
        .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  }

  // Row 14: spacer
  dash.setRowHeight(14, 10);

  // ── Rows 15-18: SAVINGS & INDICATORS ──────────────────────────────────
  dash.getRange("A15:H15").merge()
      .setValue("SAVINGS & INDICATORS")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  dash.getRange(16, 1).setValue("Recommended savings")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(16, 5).setValue("Buffer after savings")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(17, 1).setValue("Burn threshold (CLP)")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(17, 5).setValue("Cross day")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);

  // Row 18: Banco summary
  dash.getRange(18, 1).setValue("Banco Balance")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(18, 5).setValue("Debitos Banco")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);

  // Row 19: spacer
  dash.setRowHeight(19, 10);

  // ── Rows 20-24: CC SUMMARY ────────────────────────────────────────────
  dash.getRange("A20:D20").merge()
      .setValue("CC NACIONAL (CLP)")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  dash.getRange("E20:H20").merge()
      .setValue("CC INTERNACIONAL (USD)")
      .setBackground(THEME.SECONDARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  const ccLabels = ["Movements", "Live total", "Preview total", "Delta"];
  for (let i = 0; i < ccLabels.length; i++) {
    dash.getRange(21 + i, 1).setValue(ccLabels[i])
        .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
    dash.getRange(21 + i, 5).setValue(ccLabels[i])
        .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  }

  // Rows 25-27: CC PAYMENT ESTIMATES
  dash.getRange(25, 1).setValue("Cuotas monthly")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(25, 5).setValue("Cuotas monthly")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(26, 1).setValue("Payment est.")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(26, 5).setValue("Payment est.")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange("A27:H27").merge()
      .setValue("")
      .setFontSize(12).setFontWeight("bold").setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  // Row 28: spacer
  dash.setRowHeight(28, 10);

  // ── Rows 29-31: INSTALLMENTS SUMMARY ──────────────────────────────────
  dash.getRange("A29:H29").merge()
      .setValue("INSTALLMENTS")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  dash.getRange(30, 1).setValue("Active plans")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(30, 5).setValue("Monthly payment")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(31, 1).setValue("Total remaining")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  dash.getRange(31, 5).setValue("Finishing soon")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);

  // Row 32: spacer
  dash.setRowHeight(32, 10);

  // ── Rows 33-34: HOW TO IMPORT ─────────────────────────────────────────
  dash.getRange("A33:H33").merge()
      .setValue("HOW TO IMPORT")
      .setBackground(THEME.ACCENT)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  dash.getRange("A34:H34").merge()
      .setValue("1. Drop .xls in Drive → Bank_Drops   2. Finanzas → Preview → review → Confirm")
      .setFontColor(THEME.MUTED_TEXT).setFontSize(10).setFontFamily(FONT_FAMILY);

  // Column widths — matches manual layout for merged value areas
  dash.setColumnWidth(1, 130);  // Labels left
  dash.setColumnWidth(2, 145);  // Values left / Action
  dash.setColumnWidth(3, 265);  // Values left / NAC File
  dash.setColumnWidth(4, 110);  // Values left / NAC Rows
  dash.setColumnWidth(5, 350);  // Labels right / INTL File
  dash.setColumnWidth(6, 60);   // Values right / INTL Rows
  dash.setColumnWidth(7, 70);   // Values right / Status
  dash.setColumnWidth(8, 70);   // Values right
}

/**
 * Update the dashboard status card values.
 * @param {Sheet} dash
 * @param {string} status    e.g. "PREVIEW READY", "SYNC OK", "ERROR"
 * @param {string} [detail]  Additional info for hint row
 */
function updateDashboardStatus_(dash, status, detail) {
  // Break apart rows 5-7 in case of stale merges from a different layout
  dash.getRange(5, 2, 3, 3).breakApart();

  // Status value cell (B5) — color-coded
  const statusCell = dash.getRange(5, 2, 1, 3);
  statusCell.merge()
    .setValue(status)
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontFamily(FONT_FAMILY)
    .setHorizontalAlignment("left");

  // Color based on status keyword
  if (status.includes("OK") || status.includes("SUCCESS") || status.includes("IMPORTED") || status.includes("REFRESHED")) {
    statusCell.setFontColor(THEME.STATUS_SUCCESS).setBackground(THEME.POSITIVE_BG);
  } else if (status.includes("PREVIEW") || status.includes("READY")) {
    statusCell.setFontColor(THEME.STATUS_WARNING).setBackground(THEME.PREVIEW_ALT);
  } else if (status.includes("ERROR") || status.includes("BLOCKED")) {
    statusCell.setFontColor(THEME.STATUS_ERROR).setBackground(THEME.NEGATIVE_BG);
  } else if (status.includes("CLEARED") || status.includes("CANCEL")) {
    statusCell.setFontColor(THEME.STATUS_INFO).setBackground(THEME.BG_DASHBOARD);
  } else {
    statusCell.setFontColor(THEME.PRIMARY).setBackground(THEME.BG_DASHBOARD);
  }

  // Timestamp (B6)
  dash.getRange(6, 2, 1, 3).merge()
      .setValue(new Date())
      .setNumberFormat("dd/MM/yyyy HH:mm:ss")
      .setFontSize(10)
      .setFontFamily(FONT_FAMILY)
      .setFontColor(THEME.BLACK);

  // Hint (B7)
  if (detail) {
    dash.getRange(7, 2, 1, 3).merge()
        .setValue(detail)
        .setFontSize(9)
        .setFontFamily(FONT_FAMILY)
        .setFontColor(THEME.MUTED_TEXT);
  }
}

/**
 * Update dashboard file info section.
 */
function updateDashboardFileInfo_(dash, nacInfo, intlInfo, bancoInfo) {
  // NAC filename (F5:H5)
  const nacDisplay = [nacInfo.name, nacInfo.factName].filter(Boolean).join(" + ") || "—";
  dash.getRange(5, 6, 1, 3).merge()
      .setValue(nacDisplay)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // INTL filename (F6:H6)
  const intlDisplay = [intlInfo.name, intlInfo.factName].filter(Boolean).join(" + ") || "—";
  dash.getRange(6, 6, 1, 3).merge()
      .setValue(intlDisplay)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Banco filename (F7:H7)
  dash.getRange(7, 6, 1, 3).merge()
      .setValue((bancoInfo && bancoInfo.name) || "—")
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.MUTED_TEXT);
}

/**
 * Update dashboard CC summary cards (rows 21-24).
 */
function updateDashboardSummary_(dash, nacStats, intlStats) {
  writeSummaryColumn_(dash, 2, nacStats, "CLP");
  writeSummaryColumn_(dash, 6, intlStats, "USD");
}

function writeSummaryColumn_(dash, startCol, stats, currency) {
  const fmt = currency === "CLP" ? FMT_CLP : FMT_USD;

  // Row 21: Movements count
  dash.getRange(21, startCol, 1, 3).merge()
      .setValue(stats.rowCount != null ? stats.rowCount : "—")
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Row 22: Live total
  dash.getRange(22, startCol, 1, 3).merge()
      .setValue(stats.liveTotal != null ? stats.liveTotal : "—")
      .setNumberFormat(fmt)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Row 23: Preview total
  dash.getRange(23, startCol, 1, 3).merge()
      .setValue(stats.previewTotal != null ? stats.previewTotal : "—")
      .setNumberFormat(fmt)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Row 24: Delta — color-coded
  const deltaCell = dash.getRange(24, startCol, 1, 3);
  deltaCell.merge().setFontSize(10).setFontFamily(FONT_FAMILY).setNumberFormat(fmt);
  if (stats.delta != null) {
    deltaCell.setValue(stats.delta);
    if (stats.delta > 0) {
      deltaCell.setFontColor(THEME.POSITIVE_TEXT).setBackground(THEME.POSITIVE_BG);
    } else if (stats.delta < 0) {
      deltaCell.setFontColor(THEME.NEGATIVE_TEXT).setBackground(THEME.NEGATIVE_BG);
    } else {
      deltaCell.setFontColor(THEME.MUTED_TEXT).setBackground(THEME.BG_DASHBOARD);
    }
  } else {
    deltaCell.setValue("—").setFontColor(THEME.MUTED_TEXT).setBackground(THEME.BG_DASHBOARD);
  }
}

/**
 * Update the FINANCIAL OVERVIEW section (rows 10-13),
 * SAVINGS & INDICATORS section (rows 16-17), and
 * BANCO summary (row 18).
 * @param {Sheet} dash
 * @param {Object} config        From readConfig_
 * @param {number} nacTotal      Sum of positive amounts in MOV_CC_NACIONAL (CLP)
 * @param {number} intlTotalCLP  Sum of positive INTL amounts × USD_CLP
 * @param {number} [bancoDebits] Absolute value of negative banco movements
 * @param {number} [bancoBalance] Saldo disponible from banco file
 */
function updateFinancialOverview_(dash, config, nacTotal, intlTotalCLP, bancoDebits, bancoBalance) {
  const costosFijos = config.housing + config.familia;
  const totalGastar = costosFijos + nacTotal + intlTotalCLP + config.claudePlan;
  const margen = config.salario - totalGastar;
  const fsi = config.salario > 0 ? margen / config.salario : 0;

  // Stability classification
  let stability, stabilityColor, stabilityBg;
  if (fsi > 0.3) {
    stability = "GREEN"; stabilityColor = THEME.STATUS_SUCCESS; stabilityBg = THEME.POSITIVE_BG;
  } else if (fsi > 0.1) {
    stability = "YELLOW"; stabilityColor = THEME.STATUS_WARNING; stabilityBg = THEME.PREVIEW_ALT;
  } else {
    stability = "RED"; stabilityColor = THEME.STATUS_ERROR; stabilityBg = THEME.NEGATIVE_BG;
  }

  // Savings recommendation
  let savings;
  if (margen > 500000) savings = 500000;
  else if (margen > 0) savings = Math.round(margen / 2);
  else savings = 0;

  const bufferAfterSavings = Math.max(0, margen - savings);

  // Burn calculations
  const burnThreshold = 750000;
  const dailyBurn = margen > 0 ? margen / 30 : 0;
  const crossDay = dailyBurn > 0 ? Math.floor(burnThreshold / dailyBurn) : 0;

  // ── FINANCIAL OVERVIEW values (rows 10-13) ────────────────────────────

  // Left side (cols B-D)
  dash.getRange(10, 2, 1, 3).merge()
      .setValue(config.salario).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(11, 2, 1, 3).merge()
      .setValue(nacTotal).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(12, 2, 1, 3).merge()
      .setValue(totalGastar).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  const fsiCell = dash.getRange(13, 2, 1, 3);
  fsiCell.merge().setValue(fsi).setNumberFormat(FMT_PCT)
      .setFontSize(10).setFontWeight("bold").setFontFamily(FONT_FAMILY);
  fsiCell.setFontColor(stabilityColor).setBackground(stabilityBg);

  // Right side (cols F-H)
  dash.getRange(10, 6, 1, 3).merge()
      .setValue(config.housing).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(11, 6, 1, 3).merge()
      .setValue(intlTotalCLP).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Available Margin = min(margen, bancoBalance) if bancoBalance known
  const displayMargen = (bancoBalance && bancoBalance > 0)
    ? Math.min(margen, bancoBalance)
    : margen;

  const margenCell = dash.getRange(12, 6, 1, 3);
  margenCell.merge().setValue(displayMargen).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontWeight("bold").setFontFamily(FONT_FAMILY);
  if (displayMargen >= 0) {
    margenCell.setFontColor(THEME.POSITIVE_TEXT).setBackground(THEME.POSITIVE_BG);
  } else {
    margenCell.setFontColor(THEME.NEGATIVE_TEXT).setBackground(THEME.NEGATIVE_BG);
  }

  const stabCell = dash.getRange(13, 6, 1, 3);
  stabCell.merge().setValue(stability)
      .setFontSize(10).setFontWeight("bold").setFontFamily(FONT_FAMILY);
  stabCell.setFontColor(stabilityColor).setBackground(stabilityBg);

  // ── SAVINGS & INDICATORS values (rows 16-17) ─────────────────────────

  dash.getRange(16, 2, 1, 3).merge()
      .setValue(savings).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(16, 6, 1, 3).merge()
      .setValue(bufferAfterSavings).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(17, 2, 1, 3).merge()
      .setValue(burnThreshold).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(17, 6, 1, 3).merge()
      .setValue(crossDay > 0 ? "Day " + crossDay : "—")
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.MUTED_TEXT);

  // ── BANCO summary (row 18) ───────────────────────────────────────────
  dash.getRange(18, 2, 1, 3).merge()
      .setValue(bancoBalance || 0).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(18, 6, 1, 3).merge()
      .setValue(bancoDebits || 0).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.NEGATIVE_TEXT);
}

/**
 * Update the INSTALLMENTS SUMMARY section (rows 30-31).
 * @param {Sheet} dash
 * @param {Object} cuotasStats  From rebuildCuotas_
 */
function updateInstallmentsSummary_(dash, cuotasStats) {
  dash.getRange(30, 2, 1, 3).merge()
      .setValue(cuotasStats.activePlans)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(30, 6, 1, 3).merge()
      .setValue(cuotasStats.monthlyPayment).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(31, 2, 1, 3).merge()
      .setValue(cuotasStats.totalRemaining).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(31, 6, 1, 3).merge()
      .setValue(cuotasStats.finishingSoon)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);
}

/**
 * Update the CC PAYMENT ESTIMATES section (rows 25-27).
 * @param {Sheet} dash
 * @param {number} nacTotal      Sum of positive NAC amounts (CLP)
 * @param {number} intlTotalUSD  Sum of positive INTL amounts (USD)
 * @param {number} intlTotalCLP  INTL total converted to CLP
 * @param {Object} cuotasStats   From rebuildCuotas_
 */
function updateCCPaymentEstimate_(dash, nacTotal, intlTotalUSD, intlTotalCLP, cuotasStats) {
  // Row 25: Cuotas monthly per card
  dash.getRange(25, 2, 1, 3).merge()
      .setValue(cuotasStats.nacMonthly).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(25, 6, 1, 3).merge()
      .setValue(cuotasStats.intlMonthlyUSD).setNumberFormat(FMT_USD)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Row 26: Payment estimate per card (unbilled total already includes cuotas)
  dash.getRange(26, 2, 1, 3).merge()
      .setValue(nacTotal).setNumberFormat(FMT_CLP)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  dash.getRange(26, 6, 1, 3).merge()
      .setValue(intlTotalUSD).setNumberFormat(FMT_USD)
      .setFontSize(10).setFontFamily(FONT_FAMILY).setFontColor(THEME.BLACK);

  // Row 27: Total CC Payment in CLP (full-width, already merged in chrome)
  const totalCLP = nacTotal + intlTotalCLP;
  dash.getRange(27, 1, 1, 8)
      .setValue(totalCLP).setNumberFormat('"Total CC Payment  " #,##0 " CLP"')
      .setFontSize(12).setFontWeight("bold").setFontFamily(FONT_FAMILY)
      .setFontColor(THEME.PRIMARY).setBackground(THEME.BG_ALT_ROW)
      .setHorizontalAlignment("center");
}

/******************************************************************************
 * §8  IMPORT HISTORY (stored in PropertiesService)
 ******************************************************************************/

/**
 * Append an import log entry. Max 15, FIFO.
 * @param {string} action  "PREVIEW" | "CONFIRM" | "CANCEL"
 * @param {Object} details { nacFile, nacRows, intlFile, intlRows, bancoFile, bancoRows, status }
 */
function logImport_(action, details) {
  const props = PropertiesService.getScriptProperties();
  let history = [];
  try {
    const raw = props.getProperty("IMPORT_HISTORY");
    if (raw) history = JSON.parse(raw);
  } catch (_) { /* corrupt → reset */ }

  history.unshift({
    ts: new Date().toISOString(),
    action: action,
    nacRows: details.nacRows != null ? details.nacRows : "—",
    intlRows: details.intlRows != null ? details.intlRows : "—",
    bancoRows: details.bancoRows != null ? details.bancoRows : "—",
    status: details.status || "OK",
  });

  if (history.length > MAX_HISTORY) history = history.slice(0, MAX_HISTORY);
  props.setProperty("IMPORT_HISTORY", JSON.stringify(history));
}

/**
 * Render import history on the CONFIG tab (below config values).
 */
function renderHistory_(ss) {
  const sh = ss.getSheetByName(TAB_CONFIG);
  if (!sh) return;

  const HIST_GAP = 2;  // blank rows between config values and history
  const configLastRow = 5;  // CONFIG has 5 key-value rows
  const HIST_HEADER = configLastRow + HIST_GAP + 1;  // section header row
  const HIST_COLS_ROW = HIST_HEADER + 1;              // column headers
  const HIST_START = HIST_COLS_ROW + 1;               // first data row
  const NUM_COLS = 6;

  // Clear everything from history header onward
  const lastRow = Math.max(sh.getMaxRows(), HIST_START + MAX_HISTORY);
  if (lastRow >= HIST_HEADER) {
    sh.getRange(HIST_HEADER, 1, lastRow - HIST_HEADER + 1, NUM_COLS)
        .clearContent().clearFormat().setBackground(null);
  }

  // Section header
  sh.getRange(HIST_HEADER, 1, 1, NUM_COLS).merge()
      .setValue("IMPORT HISTORY")
      .setBackground(THEME.SECONDARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  // Column headers
  const histHeaders = ["Timestamp", "Action", "NAC Rows", "INTL Rows", "Banco Rows", "Status"];
  sh.getRange(HIST_COLS_ROW, 1, 1, NUM_COLS).setValues([histHeaders])
      .setBackground(THEME.BG_ALT_ROW)
      .setFontWeight("bold")
      .setFontSize(9)
      .setFontFamily(FONT_FAMILY)
      .setFontColor(THEME.PRIMARY)
      .setHorizontalAlignment("center");

  const props = PropertiesService.getScriptProperties();
  let history = [];
  try {
    const raw = props.getProperty("IMPORT_HISTORY");
    if (raw) history = JSON.parse(raw);
  } catch (_) { /* ignore */ }

  if (history.length === 0) {
    sh.getRange(HIST_START, 1, 1, NUM_COLS).merge()
        .setValue("No imports yet")
        .setFontColor(THEME.MUTED_TEXT)
        .setFontSize(9)
        .setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    return;
  }

  const rows = history.map(h => [
    h.ts ? formatTimestamp_(h.ts) : "—",
    h.action || "—",
    h.nacRows != null ? h.nacRows : "—",
    h.intlRows != null ? h.intlRows : "—",
    h.bancoRows != null ? h.bancoRows : "—",
    h.status || "—",
  ]);

  sh.getRange(HIST_START, 1, rows.length, NUM_COLS).setValues(rows)
      .setFontSize(9)
      .setFontFamily(FONT_FAMILY)
      .setFontColor(THEME.BLACK)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

  for (let i = 0; i < rows.length; i++) {
    const bg = i % 2 === 1 ? THEME.BG_ALT_ROW : THEME.WHITE;
    sh.getRange(HIST_START + i, 1, 1, NUM_COLS).setBackground(bg);
  }

  sh.autoResizeColumns(1, NUM_COLS);
}

/** Format ISO timestamp for display. */
function formatTimestamp_(isoStr) {
  try {
    const d = new Date(isoStr);
    const pad = n => String(n).padStart(2, "0");
    return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + " " +
           pad(d.getHours()) + ":" + pad(d.getMinutes());
  } catch (_) { return isoStr; }
}

/******************************************************************************
 * §9  PREVIEW — Combined Single Tab
 ******************************************************************************/

/**
 * Main preview logic.
 * Shows raw bank data without categories — categorization happens on confirm.
 */
function doPreview_(ss) {
  ensureAllTabs_(ss);
  const props = PropertiesService.getScriptProperties();
  const dash  = ss.getSheetByName(DASHBOARD_TAB);
  const folder = DriveApp.getFolderById(getDropFolderId_());

  const apiKey = getGeminiApiKeyIfSet_();
  let latestNac, latestIntl, latestBanco;
  let nacData = null, intlData = null, bancoResult = null;
  let latestNacFact, latestIntlFact;
  let nacFactData = null, intlFactData = null;

  if (apiKey) {
    // Gemini AI path — supports any bank file format
    toast_(ss, "Scanning files with Gemini AI…");
    try {
      const gemini = findAndExtractAllFilesGemini_(folder, apiKey);
      latestNac   = gemini.nac;
      latestIntl  = gemini.intl;
      latestBanco = gemini.banco;
      nacData  = latestNac  ? latestNac.extractedData  : null;
      intlData = latestIntl ? latestIntl.extractedData : null;
      bancoResult = latestBanco
        ? { data: latestBanco.extractedData, saldoDisponible: latestBanco.saldoDisponible }
        : null;
      latestNacFact  = gemini.nacFact;
      latestIntlFact = gemini.intlFact;
      nacFactData  = latestNacFact  ? latestNacFact.extractedData  : null;
      intlFactData = latestIntlFact ? latestIntlFact.extractedData : null;
    } catch (geminiErr) {
      Logger.log("Gemini failed, falling back to legacy: " + geminiErr.message);
      toast_(ss, "Gemini failed, using legacy parsers…");
      const legacy = findLatestFiles_(folder);
      latestNac = legacy.nac; latestIntl = legacy.intl; latestBanco = legacy.banco;
      nacData  = latestNac  ? extractMovementsFromXls_(latestNac.file, "CLP") : null;
      intlData = latestIntl ? extractMovementsFromXls_(latestIntl.file, "USD") : null;
      bancoResult = latestBanco ? extractMovementsFromBancoXls_(latestBanco.file) : null;
      latestNacFact = legacy.nacFact; latestIntlFact = legacy.intlFact;
      nacFactData  = latestNacFact  ? extractMovementsFromXls_(latestNacFact.file, "CLP") : null;
      intlFactData = latestIntlFact ? extractMovementsFromXls_(latestIntlFact.file, "USD") : null;
    }
  } else {
    // Legacy path — hardcoded BCI/Banco Estado parsers
    toast_(ss, "Scanning Drive folder…");
    const legacy = findLatestFiles_(folder);
    latestNac = legacy.nac; latestIntl = legacy.intl; latestBanco = legacy.banco;
    toast_(ss, "Converting NAC file…");
    nacData  = latestNac  ? extractMovementsFromXls_(latestNac.file, "CLP") : null;
    toast_(ss, "Converting INTL file…");
    intlData = latestIntl ? extractMovementsFromXls_(latestIntl.file, "USD") : null;
    if (latestBanco) {
      toast_(ss, "Converting Banco file…");
      bancoResult = extractMovementsFromBancoXls_(latestBanco.file);
    }
    latestNacFact = legacy.nacFact; latestIntlFact = legacy.intlFact;
    nacFactData  = latestNacFact  ? extractMovementsFromXls_(latestNacFact.file, "CLP") : null;
    intlFactData = latestIntlFact ? extractMovementsFromXls_(latestIntlFact.file, "USD") : null;
  }

  nacData  = mergeWithTipo_(nacData, nacFactData);
  intlData = mergeWithTipo_(intlData, intlFactData);

  if (!latestNac && !latestIntl && !latestBanco && !latestNacFact && !latestIntlFact) {
    const prev = ensureSheet_(ss, PREVIEW_TAB);
    prev.clear();
    prev.getRange(1, 1).setValue("No matching bank files found in Drop folder.")
        .setFontFamily(FONT_FAMILY).setFontSize(11).setFontColor(THEME.STATUS_WARNING);
    if (dash) updateDashboardStatus_(dash, "PREVIEW: NO FILES", "No files found in Drop folder");
    return;
  }

  const nacFP   = latestNac   ? fileFingerprint_(latestNac.file)   : "";
  const intlFP  = latestIntl  ? fileFingerprint_(latestIntl.file)  : "";
  const bancoFP = latestBanco ? fileFingerprint_(latestBanco.file) : "";

  // Compute stats
  const liveNac   = ss.getSheetByName(TAB_NAC);
  const liveIntl  = ss.getSheetByName(TAB_INTL);
  const liveNacSum  = liveNac  ? sumColumn_(liveNac, 3)  : 0;
  const liveIntlSum = liveIntl ? sumColumn_(liveIntl, 3) : 0;
  const prevNacSum   = nacData  ? sumArrayCol_(nacData, 2)  : 0;
  const prevIntlSum  = intlData ? sumArrayCol_(intlData, 2) : 0;
  const prevNacRows  = nacData  ? Math.max(0, nacData.length - 1)  : 0;
  const prevIntlRows = intlData ? Math.max(0, intlData.length - 1) : 0;
  const prevBancoRows = bancoResult ? Math.max(0, bancoResult.data.length - 1) : 0;

  // Store fingerprints
  props.setProperty("PREVIEW_NAC_FP", nacFP);
  props.setProperty("PREVIEW_INTL_FP", intlFP);
  props.setProperty("PREVIEW_BANCO_FP", bancoFP);
  props.setProperty("PREVIEW_NAC_NAME", latestNac ? latestNac.file.getName() : "");
  props.setProperty("PREVIEW_INTL_NAME", latestIntl ? latestIntl.file.getName() : "");
  props.setProperty("PREVIEW_BANCO_NAME", latestBanco ? latestBanco.file.getName() : "");
  const nacFactFP  = latestNacFact  ? fileFingerprint_(latestNacFact.file)  : "";
  const intlFactFP = latestIntlFact ? fileFingerprint_(latestIntlFact.file) : "";
  props.setProperty("PREVIEW_NAC_FACT_FP", nacFactFP);
  props.setProperty("PREVIEW_INTL_FACT_FP", intlFactFP);
  props.setProperty("PREVIEW_NAC_FACT_NAME", latestNacFact ? latestNacFact.file.getName() : "");
  props.setProperty("PREVIEW_INTL_FACT_NAME", latestIntlFact ? latestIntlFact.file.getName() : "");

  // Write combined preview tab
  toast_(ss, "Writing preview…");
  const summaryInfo = {
    nacFile: latestNac ? latestNac.file.getName() : "—",
    nacDate: latestNac ? latestNac.fileDate : "—",
    nacFP: nacFP || "—",
    nacRows: prevNacRows,
    nacPreviewTotal: prevNacSum,
    nacLiveTotal: liveNacSum,
    nacDelta: prevNacSum - liveNacSum,
    intlFile: latestIntl ? latestIntl.file.getName() : "—",
    intlDate: latestIntl ? latestIntl.fileDate : "—",
    intlFP: intlFP || "—",
    intlRows: prevIntlRows,
    intlPreviewTotal: prevIntlSum,
    intlLiveTotal: liveIntlSum,
    intlDelta: prevIntlSum - liveIntlSum,
    bancoFile: latestBanco ? latestBanco.file.getName() : "—",
    bancoRows: prevBancoRows,
    bancoSaldo: bancoResult ? bancoResult.saldoDisponible : 0,
    nacFactFile: latestNacFact ? latestNacFact.file.getName() : "—",
    nacFactFP: nacFactFP || "—",
    intlFactFile: latestIntlFact ? latestIntlFact.file.getName() : "—",
    intlFactFP: intlFactFP || "—",
  };
  const prevSheet = writePreview_(ss, nacData, intlData, bancoResult ? bancoResult.data : null, summaryInfo);

  // Update dashboard
  if (dash) {
    formatDashboardChrome_(dash);
    updateDashboardStatus_(dash, "PREVIEW READY", "Review the PREVIEW tab, then Confirm or Cancel");
    updateDashboardFileInfo_(dash, {
      name: latestNac ? latestNac.file.getName() : null,
      date: latestNac ? String(latestNac.fileDate) : null,
      factName: latestNacFact ? latestNacFact.file.getName() : null,
    }, {
      name: latestIntl ? latestIntl.file.getName() : null,
      date: latestIntl ? String(latestIntl.fileDate) : null,
      factName: latestIntlFact ? latestIntlFact.file.getName() : null,
    }, {
      name: latestBanco ? latestBanco.file.getName() : null,
    });
    updateDashboardSummary_(dash,
      { rowCount: prevNacRows, liveTotal: liveNacSum, previewTotal: prevNacSum, delta: prevNacSum - liveNacSum, status: "Preview" },
      { rowCount: prevIntlRows, liveTotal: liveIntlSum, previewTotal: prevIntlSum, delta: prevIntlSum - liveIntlSum, status: "Preview" }
    );
    renderHistory_(ss);
  }

  logImport_("PREVIEW", {
    nacFile: latestNac ? latestNac.file.getName() : null,
    nacRows: prevNacRows,
    intlFile: latestIntl ? latestIntl.file.getName() : null,
    intlRows: prevIntlRows,
    bancoFile: latestBanco ? latestBanco.file.getName() : null,
    bancoRows: prevBancoRows,
    nacFactFile: latestNacFact ? latestNacFact.file.getName() : null,
    intlFactFile: latestIntlFact ? latestIntlFact.file.getName() : null,
    status: "OK",
  });

  // Navigate to preview tab
  if (prevSheet) ss.setActiveSheet(prevSheet);
}

/**
 * Write all preview content to a single PREVIEW tab.
 */
function writePreview_(ss, nacData, intlData, bancoData, info) {
  const prev = ensureSheet_(ss, PREVIEW_TAB);
  prev.clear();

  let row = 1;

  // ── Section 1: Summary ─────────────────────────────────────────────────
  // Title
  prev.getRange(row, 1, 1, 3).merge()
      .setValue("PREVIEW IMPORT")
      .setBackground(THEME.PREVIEW_ACCENT)
      .setFontColor(THEME.WHITE)
      .setFontSize(14)
      .setFontWeight("bold")
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");
  prev.setRowHeight(row, 36);
  row++;

  // Generated timestamp
  prev.getRange(row, 1).setValue("Generated")
      .setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2, 1, 2).merge().setValue(new Date())
      .setNumberFormat("dd/MM/yyyy HH:mm:ss")
      .setFontSize(9).setFontFamily(FONT_FAMILY);
  row++;

  // NAC info
  prev.getRange(row, 1).setValue("NAC file").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2, 1, 2).merge().setValue(info.nacFile).setFontSize(9).setFontFamily(FONT_FAMILY);
  row++;
  prev.getRange(row, 1).setValue("NAC rows").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2).setValue(info.nacRows).setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 3).setValue("Δ " + (info.nacDelta >= 0 ? "+" : "") + info.nacDelta)
      .setFontSize(9).setFontFamily(FONT_FAMILY)
      .setFontColor(info.nacDelta >= 0 ? THEME.POSITIVE_TEXT : THEME.NEGATIVE_TEXT)
      .setBackground(info.nacDelta >= 0 ? THEME.POSITIVE_BG : THEME.NEGATIVE_BG);
  row++;

  // INTL info
  prev.getRange(row, 1).setValue("INTL file").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2, 1, 2).merge().setValue(info.intlFile).setFontSize(9).setFontFamily(FONT_FAMILY);
  row++;
  prev.getRange(row, 1).setValue("INTL rows").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2).setValue(info.intlRows).setFontSize(9).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 3).setValue("Δ " + (info.intlDelta >= 0 ? "+" : "") + info.intlDelta)
      .setFontSize(9).setFontFamily(FONT_FAMILY)
      .setFontColor(info.intlDelta >= 0 ? THEME.POSITIVE_TEXT : THEME.NEGATIVE_TEXT)
      .setBackground(info.intlDelta >= 0 ? THEME.POSITIVE_BG : THEME.NEGATIVE_BG);
  row++;

  // Banco info
  if (info.bancoFile && info.bancoFile !== "—") {
    prev.getRange(row, 1).setValue("Banco file").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 2, 1, 2).merge().setValue(info.bancoFile).setFontSize(9).setFontFamily(FONT_FAMILY);
    row++;
    prev.getRange(row, 1).setValue("Banco rows").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 2).setValue(info.bancoRows).setFontSize(9).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 3).setValue(info.bancoSaldo || 0)
        .setNumberFormat("\"Saldo: \"" + FMT_CLP)
        .setFontSize(9).setFontFamily(FONT_FAMILY)
        .setFontColor(THEME.STATUS_INFO);
    row++;
  }

  // Facturado file info
  if (info.nacFactFile && info.nacFactFile !== "—") {
    prev.getRange(row, 1).setValue("NAC Fact file").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 2, 1, 2).merge().setValue(info.nacFactFile).setFontSize(9).setFontFamily(FONT_FAMILY);
    row++;
  }
  if (info.intlFactFile && info.intlFactFile !== "—") {
    prev.getRange(row, 1).setValue("INTL Fact file").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(9).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 2, 1, 2).merge().setValue(info.intlFactFile).setFontSize(9).setFontFamily(FONT_FAMILY);
    row++;
  }

  // Checksums
  prev.getRange(row, 1).setValue("NAC MD5").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(8).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2, 1, 2).merge().setValue(info.nacFP).setFontSize(8).setFontFamily(FONT_FAMILY).setFontColor(THEME.MUTED_TEXT);
  row++;
  prev.getRange(row, 1).setValue("INTL MD5").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(8).setFontFamily(FONT_FAMILY);
  prev.getRange(row, 2, 1, 2).merge().setValue(info.intlFP).setFontSize(8).setFontFamily(FONT_FAMILY).setFontColor(THEME.MUTED_TEXT);
  row++;
  if (info.nacFactFP && info.nacFactFP !== "—") {
    prev.getRange(row, 1).setValue("NAC Fact MD5").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(8).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 2, 1, 2).merge().setValue(info.nacFactFP).setFontSize(8).setFontFamily(FONT_FAMILY).setFontColor(THEME.MUTED_TEXT);
    row++;
  }
  if (info.intlFactFP && info.intlFactFP !== "—") {
    prev.getRange(row, 1).setValue("INTL Fact MD5").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(8).setFontFamily(FONT_FAMILY);
    prev.getRange(row, 2, 1, 2).merge().setValue(info.intlFactFP).setFontSize(8).setFontFamily(FONT_FAMILY).setFontColor(THEME.MUTED_TEXT);
    row++;
  }

  // Spacer row before NAC data
  row++;

  // ── Section 2: NAC Movements ───────────────────────────────────────────
  prev.getRange(row, 1, 1, 4).merge()
      .setValue("CC NACIONAL (CLP)")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");
  prev.setRowHeight(row, 30);
  row++;

  if (nacData && nacData.length > 1) {
    // Header row
    prev.getRange(row, 1, 1, 4).setValues([nacData[0]])
        .setBackground(THEME.PREVIEW_ACCENT)
        .setFontColor(THEME.WHITE)
        .setFontWeight("bold")
        .setFontSize(10)
        .setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    row++;

    // Data rows
    const dataOnly = nacData.slice(1);
    prev.getRange(row, 1, dataOnly.length, 4).setValues(dataOnly)
        .setFontSize(10).setFontFamily(FONT_FAMILY);

    // Format columns
    prev.getRange(row, 1, dataOnly.length, 1).setNumberFormat(FMT_DATE).setHorizontalAlignment("center");
    prev.getRange(row, 2, dataOnly.length, 1).setHorizontalAlignment("left");
    prev.getRange(row, 3, dataOnly.length, 1).setNumberFormat(FMT_CLP).setHorizontalAlignment("right");
    prev.getRange(row, 4, dataOnly.length, 1).setHorizontalAlignment("left");

    // Zebra + conditional
    applyZebraStripes_(prev, row, dataOnly.length, 4, true);
    applyAmountConditionalFormatting_(prev, 3, row, row + dataOnly.length - 1);

    row += dataOnly.length;
  } else {
    prev.getRange(row, 1, 1, 4).merge()
        .setValue("No nacional data")
        .setFontColor(THEME.MUTED_TEXT).setFontSize(10).setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    row++;
  }

  // 2 spacer rows before INTL
  row += 2;

  // ── Section 3: INTL Movements ──────────────────────────────────────────
  prev.getRange(row, 1, 1, 4).merge()
      .setValue("CC INTERNACIONAL (USD)")
      .setBackground(THEME.SECONDARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");
  prev.setRowHeight(row, 30);
  row++;

  if (intlData && intlData.length > 1) {
    // Header row
    prev.getRange(row, 1, 1, 4).setValues([intlData[0]])
        .setBackground(THEME.PREVIEW_ACCENT)
        .setFontColor(THEME.WHITE)
        .setFontWeight("bold")
        .setFontSize(10)
        .setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    row++;

    // Data rows
    const dataOnly = intlData.slice(1);
    prev.getRange(row, 1, dataOnly.length, 4).setValues(dataOnly)
        .setFontSize(10).setFontFamily(FONT_FAMILY);

    prev.getRange(row, 1, dataOnly.length, 1).setNumberFormat(FMT_DATE).setHorizontalAlignment("center");
    prev.getRange(row, 2, dataOnly.length, 1).setHorizontalAlignment("left");
    prev.getRange(row, 3, dataOnly.length, 1).setNumberFormat(FMT_USD).setHorizontalAlignment("right");
    prev.getRange(row, 4, dataOnly.length, 1).setHorizontalAlignment("left");

    applyZebraStripes_(prev, row, dataOnly.length, 4, true);
    applyAmountConditionalFormatting_(prev, 3, row, row + dataOnly.length - 1);

    row += dataOnly.length;
  } else {
    prev.getRange(row, 1, 1, 4).merge()
        .setValue("No internacional data")
        .setFontColor(THEME.MUTED_TEXT).setFontSize(10).setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    row++;
  }

  // ── Section 4: Banco Movements ──────────────────────────────────────
  if (bancoData && bancoData.length > 1) {
    // 2 spacer rows before Banco
    row += 2;

    prev.getRange(row, 1, 1, 3).merge()
        .setValue("BANCO ESTADO (CLP)")
        .setBackground(THEME.STATUS_INFO)
        .setFontColor(THEME.WHITE)
        .setFontWeight("bold")
        .setFontSize(11)
        .setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    prev.setRowHeight(row, 30);
    row++;

    // Header row
    prev.getRange(row, 1, 1, 3).setValues([bancoData[0]])
        .setBackground(THEME.PREVIEW_ACCENT)
        .setFontColor(THEME.WHITE)
        .setFontWeight("bold")
        .setFontSize(10)
        .setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
    row++;

    // Data rows
    const dataOnly = bancoData.slice(1);
    prev.getRange(row, 1, dataOnly.length, 3).setValues(dataOnly)
        .setFontSize(10).setFontFamily(FONT_FAMILY);

    prev.getRange(row, 1, dataOnly.length, 1).setNumberFormat(FMT_DATE).setHorizontalAlignment("center");
    prev.getRange(row, 2, dataOnly.length, 1).setHorizontalAlignment("left");
    prev.getRange(row, 3, dataOnly.length, 1).setNumberFormat(FMT_CLP).setHorizontalAlignment("right");

    applyZebraStripes_(prev, row, dataOnly.length, 3, true);
    applyAmountConditionalFormatting_(prev, 3, row, row + dataOnly.length - 1);

    row += dataOnly.length;
  }

  // Auto-fit column widths to content
  prev.autoResizeColumns(1, 4);

  return prev;
}

/******************************************************************************
 * §10  CONFIRM IMPORT
 ******************************************************************************/

function doConfirm_(ss) {
  ensureAllTabs_(ss);
  ensureChileanLocale_(ss);
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const dash  = ss.getSheetByName(DASHBOARD_TAB);

  const previewNacFP   = props.getProperty("PREVIEW_NAC_FP")   || "";
  const previewIntlFP  = props.getProperty("PREVIEW_INTL_FP")  || "";
  const previewBancoFP = props.getProperty("PREVIEW_BANCO_FP") || "";
  const previewNacFactFP  = props.getProperty("PREVIEW_NAC_FACT_FP")  || "";
  const previewIntlFactFP = props.getProperty("PREVIEW_INTL_FACT_FP") || "";

  if (!previewNacFP && !previewIntlFP && !previewBancoFP) {
    throw new Error("No preview found. Run Preview Import first.");
  }

  // Confirmation dialog
  const resp = ui.alert(
    "Confirm Import",
    "This will overwrite the live data tabs with the previewed data.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    toast_(ss, "Import cancelled by user.");
    return;
  }

  toast_(ss, "Verifying file integrity…");
  const folder = DriveApp.getFolderById(getDropFolderId_());
  const apiKey = getGeminiApiKeyIfSet_();

  let latestNac, latestIntl, latestBanco;
  let latestNacFact, latestIntlFact;
  let geminiData = null;

  if (apiKey) {
    try {
      geminiData = findAndExtractAllFilesGemini_(folder, apiKey);
      latestNac = geminiData.nac; latestIntl = geminiData.intl; latestBanco = geminiData.banco;
      latestNacFact = geminiData.nacFact; latestIntlFact = geminiData.intlFact;
    } catch (geminiErr) {
      Logger.log("Gemini failed in confirm, falling back: " + geminiErr.message);
      toast_(ss, "Gemini failed, using legacy parsers…");
      geminiData = null;
      const legacy = findLatestFiles_(folder);
      latestNac = legacy.nac; latestIntl = legacy.intl; latestBanco = legacy.banco;
      latestNacFact = legacy.nacFact; latestIntlFact = legacy.intlFact;
    }
  } else {
    const legacy = findLatestFiles_(folder);
    latestNac = legacy.nac; latestIntl = legacy.intl; latestBanco = legacy.banco;
    latestNacFact = legacy.nacFact; latestIntlFact = legacy.intlFact;
  }

  const latestNacFP   = latestNac   ? fileFingerprint_(latestNac.file)   : "";
  const latestIntlFP  = latestIntl  ? fileFingerprint_(latestIntl.file)  : "";
  const latestBancoFP = latestBanco ? fileFingerprint_(latestBanco.file) : "";
  const latestNacFactFP   = latestNacFact  ? fileFingerprint_(latestNacFact.file)  : "";
  const latestIntlFactFP  = latestIntlFact ? fileFingerprint_(latestIntlFact.file) : "";

  if (latestNac && previewNacFP && latestNacFP !== previewNacFP) {
    throw new Error("NAC file changed since preview. Run Preview Import again.");
  }
  if (latestIntl && previewIntlFP && latestIntlFP !== previewIntlFP) {
    throw new Error("INTL file changed since preview. Run Preview Import again.");
  }
  if (latestBanco && previewBancoFP && latestBancoFP !== previewBancoFP) {
    throw new Error("Banco file changed since preview. Run Preview Import again.");
  }
  if (latestNacFact && previewNacFactFP && latestNacFactFP !== previewNacFactFP) {
    throw new Error("NAC Facturado file changed since preview. Run Preview Import again.");
  }
  if (latestIntlFact && previewIntlFactFP && latestIntlFactFP !== previewIntlFactFP) {
    throw new Error("INTL Facturado file changed since preview. Run Preview Import again.");
  }

  // Read config and categories for auto-categorization
  toast_(ss, "Reading configuration…");
  const config = readConfig_(ss);
  const categories = readCategories_(ss);

  let statusParts = [];
  let nacRows = 0, intlRows = 0, bancoRows = 0;
  let bancoDebits = 0, bancoBalance = 0;

  if (latestNac || latestNacFact) {
    toast_(ss, "Importing NAC movements…");
    const nacNoFact = latestNac ? (geminiData && geminiData.nac ? geminiData.nac.extractedData : extractMovementsFromXls_(latestNac.file, "CLP")) : null;
    const nacFact = latestNacFact ? (geminiData && geminiData.nacFact ? geminiData.nacFact.extractedData : extractMovementsFromXls_(latestNacFact.file, "CLP")) : null;
    const nacData = mergeWithTipo_(nacNoFact, nacFact);
    overwriteTab_(ss, TAB_NAC, nacData, "CLP", categories, config);
    if (latestNac) props.setProperty("LAST_NAC_FP", latestNacFP);
    if (latestNacFact) props.setProperty("LAST_NAC_FACT_FP", latestNacFactFP);
    nacRows = Math.max(0, nacData.length - 1);
    statusParts.push("NAC: IMPORTED");
  } else {
    statusParts.push("NAC: NO FILE");
  }

  if (latestIntl || latestIntlFact) {
    toast_(ss, "Importing INTL movements…");
    const intlNoFact = latestIntl ? (geminiData && geminiData.intl ? geminiData.intl.extractedData : extractMovementsFromXls_(latestIntl.file, "USD")) : null;
    const intlFact = latestIntlFact ? (geminiData && geminiData.intlFact ? geminiData.intlFact.extractedData : extractMovementsFromXls_(latestIntlFact.file, "USD")) : null;
    const intlData = mergeWithTipo_(intlNoFact, intlFact);
    overwriteTab_(ss, TAB_INTL, intlData, "USD", categories, config);
    if (latestIntl) props.setProperty("LAST_INTL_FP", latestIntlFP);
    if (latestIntlFact) props.setProperty("LAST_INTL_FACT_FP", latestIntlFactFP);
    intlRows = Math.max(0, intlData.length - 1);
    statusParts.push("INTL: IMPORTED");
  } else {
    statusParts.push("INTL: NO FILE");
  }

  if (latestBanco) {
    toast_(ss, "Importing Banco movements…");
    let bancoResult;
    if (geminiData && geminiData.banco) {
      bancoResult = { data: geminiData.banco.extractedData, saldoDisponible: geminiData.banco.saldoDisponible };
    } else {
      bancoResult = extractMovementsFromBancoXls_(latestBanco.file);
    }
    overwriteBancoTab_(ss, bancoResult.data, categories);
    props.setProperty("LAST_BANCO_FP", latestBancoFP);
    bancoRows = Math.max(0, bancoResult.data.length - 1);
    bancoBalance = bancoResult.saldoDisponible;
    props.setProperty("BANCO_SALDO_DISPONIBLE", String(bancoBalance));
    for (let i = 1; i < bancoResult.data.length; i++) {
      const m = bancoResult.data[i][2];
      if (typeof m === "number" && m < 0) bancoDebits += Math.abs(m);
    }
    statusParts.push("BANCO: IMPORTED");
  } else {
    statusParts.push("BANCO: NO FILE");
  }

  // Move imported files to _imported/ subfolder
  try {
    toast_(ss, "Moving files to _imported/…");
    const allFiles = folder.getFiles();
    const filesToMove = [];
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      const name = f.getName();
      if (NAC_REGEX.test(name) || INTL_REGEX.test(name) || BANCO_REGEX.test(name)
          || NAC_FACT_REGEX.test(name) || INTL_FACT_REGEX.test(name)) {
        filesToMove.push(f);
      }
    }
    if (filesToMove.length > 0) {
      moveFilesToImported_(folder, filesToMove);
    }
  } catch (moveErr) {
    Logger.log("Failed to move files: " + moveErr.message);
  }

  // Clear preview state
  props.deleteProperty("PREVIEW_NAC_FP");
  props.deleteProperty("PREVIEW_INTL_FP");
  props.deleteProperty("PREVIEW_BANCO_FP");
  props.deleteProperty("PREVIEW_NAC_NAME");
  props.deleteProperty("PREVIEW_INTL_NAME");
  props.deleteProperty("PREVIEW_BANCO_NAME");
  props.deleteProperty("PREVIEW_NAC_FACT_FP");
  props.deleteProperty("PREVIEW_INTL_FACT_FP");
  props.deleteProperty("PREVIEW_NAC_FACT_NAME");
  props.deleteProperty("PREVIEW_INTL_FACT_NAME");

  // Clear preview tab
  const prev = ss.getSheetByName(PREVIEW_TAB);
  if (prev) prev.clear();

  // Rebuild derived data
  toast_(ss, "Rebuilding installments…");
  const cuotasStats = rebuildCuotas_(ss, config);

  toast_(ss, "Rebuilding category summary…");
  rebuildResumenCategorias_(ss, categories, config);

  // Log & update dashboard
  logImport_("CONFIRM", {
    nacFile: latestNac ? latestNac.file.getName() : null,
    nacRows: nacRows,
    nacFactFile: latestNacFact ? latestNacFact.file.getName() : null,
    intlFile: latestIntl ? latestIntl.file.getName() : null,
    intlRows: intlRows,
    intlFactFile: latestIntlFact ? latestIntlFact.file.getName() : null,
    bancoFile: latestBanco ? latestBanco.file.getName() : null,
    bancoRows: bancoRows,
    status: "OK",
  });

  if (dash) {
    formatDashboardChrome_(dash);
    updateDashboardStatus_(dash, "SYNC OK — IMPORTED", statusParts.join(" · "));

    // Update CC summary with new live data
    const liveNac  = ss.getSheetByName(TAB_NAC);
    const liveIntl = ss.getSheetByName(TAB_INTL);
    updateDashboardSummary_(dash,
      { rowCount: nacRows, liveTotal: liveNac ? sumColumn_(liveNac, 3) : 0, previewTotal: null, delta: null },
      { rowCount: intlRows, liveTotal: liveIntl ? sumColumn_(liveIntl, 3) : 0, previewTotal: null, delta: null }
    );
    updateDashboardFileInfo_(dash, {
      name: latestNac ? latestNac.file.getName() : null,
      date: latestNac ? String(latestNac.fileDate) : null,
      factName: latestNacFact ? latestNacFact.file.getName() : null,
    }, {
      name: latestIntl ? latestIntl.file.getName() : null,
      date: latestIntl ? String(latestIntl.fileDate) : null,
      factName: latestIntlFact ? latestIntlFact.file.getName() : null,
    }, {
      name: latestBanco ? latestBanco.file.getName() : null,
    });

    // Financial overview + installments
    const nacPositiveSum = liveNac ? sumPositiveColumn_(liveNac, 3) : 0;
    const intlPositiveSum = liveIntl ? sumPositiveColumn_(liveIntl, 3) : 0;
    const intlTotalCLP = intlPositiveSum * config.usdClp;
    updateFinancialOverview_(dash, config, nacPositiveSum, intlTotalCLP, bancoDebits, bancoBalance);
    updateInstallmentsSummary_(dash, cuotasStats);
    updateCCPaymentEstimate_(dash, nacPositiveSum, intlPositiveSum, intlTotalCLP, cuotasStats);

    renderHistory_(ss);

    // Navigate to dashboard
    ss.setActiveSheet(dash);
  }
}

/******************************************************************************
 * §11  CANCEL PREVIEW
 ******************************************************************************/

function doCancel_(ss) {
  const props = PropertiesService.getScriptProperties();
  const dash  = ss.getSheetByName(DASHBOARD_TAB);

  const prev = ss.getSheetByName(PREVIEW_TAB);
  if (prev) prev.clear();

  props.deleteProperty("PREVIEW_NAC_FP");
  props.deleteProperty("PREVIEW_INTL_FP");
  props.deleteProperty("PREVIEW_BANCO_FP");
  props.deleteProperty("PREVIEW_NAC_NAME");
  props.deleteProperty("PREVIEW_INTL_NAME");
  props.deleteProperty("PREVIEW_BANCO_NAME");
  props.deleteProperty("PREVIEW_NAC_FACT_FP");
  props.deleteProperty("PREVIEW_INTL_FACT_FP");
  props.deleteProperty("PREVIEW_NAC_FACT_NAME");
  props.deleteProperty("PREVIEW_INTL_FACT_NAME");

  logImport_("CANCEL", { status: "Cancelled" });

  if (dash) {
    formatDashboardChrome_(dash);
    updateDashboardStatus_(dash, "PREVIEW CLEARED", "Ready for a new preview");
    renderHistory_(ss);
  }
}

/******************************************************************************
 * §12  CUOTAS — Installment Parsing
 ******************************************************************************/

/**
 * Rebuild the CUOTAS tab by parsing installment patterns from MOV data.
 * Pattern: (CC|CF) NN-NN where NN-NN = cuota_actual-cuotas_totales
 *
 * @param {Spreadsheet} ss
 * @param {Object} config  From readConfig_
 * @returns {Object} { activePlans, monthlyPayment, totalRemaining, finishingSoon }
 */
function rebuildCuotas_(ss, config) {
  const usdClp = config ? config.usdClp : 0;
  const installments = [];

  // Parse NAC movements
  const nacSheet = ss.getSheetByName(TAB_NAC);
  if (nacSheet && nacSheet.getLastRow() > 1) {
    const nacData = nacSheet.getRange(2, 1, nacSheet.getLastRow() - 1, 3).getValues();
    for (const [fecha, desc, monto] of nacData) {
      const m = CUOTA_REGEX.exec(String(desc));
      if (m) {
        const actual = parseInt(m[1], 10);
        const total  = parseInt(m[2], 10);
        const restantes = total - actual;
        installments.push({
          fuente: "NAC", fecha: fecha, desc: String(desc),
          monto: typeof monto === "number" ? monto : 0,
          moneda: "CLP", actual: actual, total: total,
          restantes: restantes,
          restanteEst: (typeof monto === "number" ? monto : 0) * restantes,
          montoCLP: typeof monto === "number" ? monto : 0,
        });
      }
    }
  }

  // Parse INTL movements
  const intlSheet = ss.getSheetByName(TAB_INTL);
  if (intlSheet && intlSheet.getLastRow() > 1) {
    const intlData = intlSheet.getRange(2, 1, intlSheet.getLastRow() - 1, 3).getValues();
    for (const [fecha, desc, monto] of intlData) {
      const m = CUOTA_REGEX.exec(String(desc));
      if (m) {
        const actual = parseInt(m[1], 10);
        const total  = parseInt(m[2], 10);
        const restantes = total - actual;
        const numMonto = typeof monto === "number" ? monto : 0;
        installments.push({
          fuente: "INTL", fecha: fecha, desc: String(desc),
          monto: numMonto,
          moneda: "USD", actual: actual, total: total,
          restantes: restantes,
          restanteEst: numMonto * restantes,
          montoCLP: numMonto * usdClp,
        });
      }
    }
  }

  // Sort by restantes descending (most remaining first)
  installments.sort((a, b) => b.restantes - a.restantes);

  // Compute summary stats (in CLP)
  const activePlans = installments.length;
  const monthlyPayment = installments.reduce((s, i) => s + i.montoCLP, 0);
  const totalRemaining = installments.reduce((s, i) => s + (i.montoCLP * i.restantes), 0);
  const finishingSoon = installments.filter(i => i.restantes <= 1).length;

  // Write CUOTAS tab
  const sh = ensureSheet_(ss, TAB_CUOTAS);
  sh.clear();

  const NUM_COLS = 9;

  // Row 1: Section header
  sh.getRange(1, 1, 1, NUM_COLS).merge()
      .setValue("INSTALLMENTS")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(14)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");
  sh.setRowHeight(1, 36);

  // Row 2: Summary line 1
  sh.getRange(2, 1).setValue("Active plans").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  sh.getRange(2, 2).setValue(activePlans).setFontSize(10).setFontFamily(FONT_FAMILY).setFontWeight("bold");
  sh.getRange(2, 5).setValue("Monthly payment (est.)").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  sh.getRange(2, 7).setValue(monthlyPayment).setNumberFormat(FMT_CLP).setFontSize(10).setFontFamily(FONT_FAMILY).setFontWeight("bold");

  // Row 3: Summary line 2
  sh.getRange(3, 1).setValue("Total remaining (est.)").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  sh.getRange(3, 2).setValue(totalRemaining).setNumberFormat(FMT_CLP).setFontSize(10).setFontFamily(FONT_FAMILY).setFontWeight("bold");
  sh.getRange(3, 5).setValue("Finishing next month").setFontColor(THEME.MUTED_TEXT).setFontWeight("bold").setFontSize(10).setFontFamily(FONT_FAMILY);
  sh.getRange(3, 7).setValue(finishingSoon).setFontSize(10).setFontFamily(FONT_FAMILY).setFontWeight("bold");

  // Row 4: Spacer
  sh.setRowHeight(4, 8);

  // Row 5: Header row
  const headers = ["Fuente", "Fecha", "Descripcion", "Monto", "Moneda", "Cuota", "Total", "Restantes", "Restante_est"];
  sh.getRange(5, 1, 1, NUM_COLS).setValues([headers])
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(10)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");
  sh.setFrozenRows(5);

  // Row 6+: Installment data
  if (installments.length > 0) {
    const dataRows = installments.map(i => [
      i.fuente, i.fecha, i.desc, i.monto, i.moneda,
      i.actual, i.total, i.restantes, i.restanteEst
    ]);
    sh.getRange(6, 1, dataRows.length, NUM_COLS).setValues(dataRows)
        .setFontSize(10).setFontFamily(FONT_FAMILY).setVerticalAlignment("middle");

    // Format columns
    sh.getRange(6, 1, dataRows.length, 1).setHorizontalAlignment("center");  // Fuente
    sh.getRange(6, 2, dataRows.length, 1).setNumberFormat(FMT_DATE).setHorizontalAlignment("center");  // Fecha
    sh.getRange(6, 3, dataRows.length, 1).setHorizontalAlignment("left");  // Descripcion

    // Format Monto per row based on currency
    for (let r = 0; r < dataRows.length; r++) {
      const fmt = installments[r].moneda === "CLP" ? FMT_CLP : FMT_USD;
      sh.getRange(6 + r, 4).setNumberFormat(fmt).setHorizontalAlignment("right");
      sh.getRange(6 + r, 9).setNumberFormat(fmt).setHorizontalAlignment("right");
    }

    sh.getRange(6, 5, dataRows.length, 1).setHorizontalAlignment("center");  // Moneda
    sh.getRange(6, 6, dataRows.length, 1).setHorizontalAlignment("center");  // Cuota
    sh.getRange(6, 7, dataRows.length, 1).setHorizontalAlignment("center");  // Total
    sh.getRange(6, 8, dataRows.length, 1).setHorizontalAlignment("center");  // Restantes

    // Zebra stripes
    applyZebraStripes_(sh, 6, dataRows.length, NUM_COLS, false);

    // Conditional formatting on Restantes: red for high, green for low
    const restRange = sh.getRange(6, 8, dataRows.length, 1);
    const rules = sh.getConditionalFormatRules();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(6)
        .setBackground(THEME.NEGATIVE_BG).setFontColor(THEME.NEGATIVE_TEXT)
        .setRanges([restRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThanOrEqualTo(1)
        .setBackground(THEME.POSITIVE_BG).setFontColor(THEME.POSITIVE_TEXT)
        .setRanges([restRange]).build()
    );
    sh.setConditionalFormatRules(rules);
  } else {
    sh.getRange(6, 1, 1, NUM_COLS).merge()
        .setValue("No installment movements found")
        .setFontColor(THEME.MUTED_TEXT).setFontSize(10).setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
  }

  // Auto-fit column widths to content
  sh.autoResizeColumns(1, NUM_COLS);

  // Per-source monthly breakdown
  const nacMonthly = installments.filter(i => i.fuente === "NAC").reduce((s, i) => s + i.montoCLP, 0);
  const intlMonthly = installments.filter(i => i.fuente === "INTL").reduce((s, i) => s + i.montoCLP, 0);
  const intlMonthlyUSD = installments.filter(i => i.fuente === "INTL").reduce((s, i) => s + i.monto, 0);

  return { activePlans, monthlyPayment, totalRemaining, finishingSoon, nacMonthly, intlMonthly, intlMonthlyUSD };
}

/******************************************************************************
 * §13  RESUMEN_CATEGORIAS — Category Summaries
 ******************************************************************************/

/**
 * Rebuild the RESUMEN_CATEGORIAS tab with aggregated totals by category.
 * Re-categorizes from descriptions (not from stored Categoria column) so
 * it reflects current CATEGORIAS rules.
 *
 * @param {Spreadsheet} ss
 * @param {Array}  categories  From readCategories_
 * @param {Object} config      From readConfig_
 */
function rebuildResumenCategorias_(ss, categories, config) {
  const usdClp = config ? config.usdClp : 0;
  const nacTotals = {};
  const intlTotals = {};

  // Categorize NAC movements
  const nacSheet = ss.getSheetByName(TAB_NAC);
  let nacGrandTotal = 0;
  if (nacSheet && nacSheet.getLastRow() > 1) {
    const nacData = nacSheet.getRange(2, 1, nacSheet.getLastRow() - 1, 3).getValues();
    for (const [, desc, monto] of nacData) {
      const cat = categorizeMovement_(desc, "NAC", categories);
      const num = typeof monto === "number" ? monto : 0;
      nacTotals[cat] = (nacTotals[cat] || 0) + num;
      nacGrandTotal += num;
    }
  }

  // Categorize INTL movements (convert to CLP)
  const intlSheet = ss.getSheetByName(TAB_INTL);
  let intlGrandTotalCLP = 0;
  if (intlSheet && intlSheet.getLastRow() > 1) {
    const intlData = intlSheet.getRange(2, 1, intlSheet.getLastRow() - 1, 3).getValues();
    for (const [, desc, monto] of intlData) {
      const cat = categorizeMovement_(desc, "INTL", categories);
      const num = typeof monto === "number" ? monto : 0;
      const clpVal = num * usdClp;
      intlTotals[cat] = (intlTotals[cat] || 0) + clpVal;
      intlGrandTotalCLP += clpVal;
    }
  }

  // Build sorted rows
  const nacEntries = Object.entries(nacTotals).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
  const intlEntries = Object.entries(intlTotals).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
  const maxRows = Math.max(nacEntries.length, intlEntries.length);

  // Write RESUMEN_CATEGORIAS tab
  const sh = ensureSheet_(ss, TAB_RESUMEN_CAT);
  sh.clear();

  // Row 1: Headers
  sh.getRange(1, 1, 1, 3).merge()
      .setValue("Nacional (CLP)")
      .setBackground(THEME.PRIMARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  sh.getRange(1, 4).setValue("")
      .setBackground(THEME.BG_DASHBOARD);

  sh.getRange(1, 5, 1, 3).merge()
      .setValue("Internacional (CLP est.)")
      .setBackground(THEME.SECONDARY)
      .setFontColor(THEME.WHITE)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily(FONT_FAMILY)
      .setHorizontalAlignment("center");

  // Row 2: Column headers
  const colHeaders = ["Categoria", "Total_CLP", "%", "", "Categoria", "Total_CLP", "%"];
  sh.getRange(2, 1, 1, 7).setValues([colHeaders])
      .setBackground(THEME.BG_ALT_ROW)
      .setFontWeight("bold")
      .setFontSize(9)
      .setFontFamily(FONT_FAMILY)
      .setFontColor(THEME.PRIMARY)
      .setHorizontalAlignment("center");

  // Row 3+: Category data
  if (maxRows > 0) {
    const rows = [];
    for (let i = 0; i < maxRows; i++) {
      const nacCat = nacEntries[i] ? nacEntries[i][0] : "";
      const nacVal = nacEntries[i] ? nacEntries[i][1] : "";
      const nacPct = nacEntries[i] && nacGrandTotal !== 0 ? nacEntries[i][1] / nacGrandTotal : "";

      const intlCat = intlEntries[i] ? intlEntries[i][0] : "";
      const intlVal = intlEntries[i] ? intlEntries[i][1] : "";
      const intlPct = intlEntries[i] && intlGrandTotalCLP !== 0 ? intlEntries[i][1] / intlGrandTotalCLP : "";

      rows.push([nacCat, nacVal, nacPct, "", intlCat, intlVal, intlPct]);
    }

    sh.getRange(3, 1, rows.length, 7).setValues(rows)
        .setFontSize(10).setFontFamily(FONT_FAMILY).setVerticalAlignment("middle");

    // Format NAC side
    sh.getRange(3, 1, rows.length, 1).setHorizontalAlignment("left");
    sh.getRange(3, 2, rows.length, 1).setNumberFormat(FMT_CLP).setHorizontalAlignment("right");
    sh.getRange(3, 3, rows.length, 1).setNumberFormat(FMT_PCT).setHorizontalAlignment("right")
        .setFontColor(THEME.MUTED_TEXT);

    // Format INTL side
    sh.getRange(3, 5, rows.length, 1).setHorizontalAlignment("left");
    sh.getRange(3, 6, rows.length, 1).setNumberFormat(FMT_CLP).setHorizontalAlignment("right");
    sh.getRange(3, 7, rows.length, 1).setNumberFormat(FMT_PCT).setHorizontalAlignment("right")
        .setFontColor(THEME.MUTED_TEXT);

    // Zebra stripes
    for (let i = 0; i < rows.length; i++) {
      const bg = i % 2 === 1 ? THEME.BG_ALT_ROW : THEME.WHITE;
      sh.getRange(3 + i, 1, 1, 3).setBackground(bg);
      sh.getRange(3 + i, 5, 1, 3).setBackground(bg);
    }
  } else {
    sh.getRange(3, 1, 1, 7).merge()
        .setValue("No categorized movements")
        .setFontColor(THEME.MUTED_TEXT).setFontSize(10).setFontFamily(FONT_FAMILY)
        .setHorizontalAlignment("center");
  }

  // Auto-fit column widths to content
  sh.autoResizeColumns(1, 7);
}

/******************************************************************************
 * §14  REFRESH CALCULATIONS
 ******************************************************************************/

/**
 * Recalculates derived data (CUOTAS, RESUMEN_CATEGORIAS, dashboard financials)
 * without re-importing. Useful after editing CONFIG values.
 */
function doRefreshCalculations_(ss) {
  ensureAllTabs_(ss);
  ensureChileanLocale_(ss);
  toast_(ss, "Reading configuration…");
  const config = readConfig_(ss);
  const categories = readCategories_(ss);

  toast_(ss, "Rebuilding installments…");
  const cuotasStats = rebuildCuotas_(ss, config);

  toast_(ss, "Rebuilding category summary…");
  rebuildResumenCategorias_(ss, categories, config);

  toast_(ss, "Updating dashboard…");
  const dash = ss.getSheetByName(DASHBOARD_TAB);
  if (dash) {
    formatDashboardChrome_(dash);

    const liveNac  = ss.getSheetByName(TAB_NAC);
    const liveIntl = ss.getSheetByName(TAB_INTL);

    // Banco data from MOV_BANCO tab
    let bancoDebits = 0;
    let bancoBalance = 0;
    const liveBanco = ss.getSheetByName(TAB_BANCO);
    if (liveBanco && liveBanco.getLastRow() > 1) {
      const bancoVals = liveBanco.getRange(2, 3, liveBanco.getLastRow() - 1, 1).getValues();
      for (const [v] of bancoVals) {
        const n = typeof v === "number" ? v : 0;
        if (n < 0) bancoDebits += Math.abs(n);
      }
    }
    // Read stored saldo from properties
    const saldoProp = PropertiesService.getScriptProperties().getProperty("BANCO_SALDO_DISPONIBLE");
    if (saldoProp) bancoBalance = parseFloat(saldoProp) || 0;

    // Financial overview
    const nacPositiveSum = liveNac ? sumPositiveColumn_(liveNac, 3) : 0;
    const intlPositiveSum = liveIntl ? sumPositiveColumn_(liveIntl, 3) : 0;
    const intlTotalCLP = intlPositiveSum * config.usdClp;
    updateFinancialOverview_(dash, config, nacPositiveSum, intlTotalCLP, bancoDebits, bancoBalance);

    // Installments summary + CC payment estimates
    updateInstallmentsSummary_(dash, cuotasStats);
    updateCCPaymentEstimate_(dash, nacPositiveSum, intlPositiveSum, intlTotalCLP, cuotasStats);

    // CC summary (live data only, no preview)
    const nacRows = liveNac ? Math.max(0, liveNac.getLastRow() - 1) : 0;
    const intlRows = liveIntl ? Math.max(0, liveIntl.getLastRow() - 1) : 0;
    updateDashboardSummary_(dash,
      { rowCount: nacRows, liveTotal: liveNac ? sumColumn_(liveNac, 3) : 0, previewTotal: null, delta: null },
      { rowCount: intlRows, liveTotal: liveIntl ? sumColumn_(liveIntl, 3) : 0, previewTotal: null, delta: null }
    );

    updateDashboardStatus_(dash, "REFRESHED", "Calculations updated from CONFIG");
    renderHistory_(ss);
  }
}

/******************************************************************************
 * §15  CORE HELPERS
 ******************************************************************************/

/**
 * Find the latest NAC, INTL and Banco files in the drop folder.
 * @param {Folder} folder
 * @returns {{ nac: {file, fileDate}|null, intl: {file, fileDate}|null, banco: {file, fileDate}|null }}
 */
function findLatestFiles_(folder) {
  const files = folder.getFiles();
  let latestNac = null;
  let latestIntl = null;
  let latestBanco = null;
  let latestNacFact = null;
  let latestIntlFact = null;

  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();

    let d = extractDateFromName_(name);
    if (d === null) {
      const ts = f.getLastUpdated();
      d = ts.getFullYear() * 10000 + (ts.getMonth() + 1) * 100 + ts.getDate();
    }

    if (NAC_REGEX.test(name)) {
      if (!latestNac || d > latestNac.fileDate) latestNac = { file: f, fileDate: d };
    }
    if (INTL_REGEX.test(name)) {
      if (!latestIntl || d > latestIntl.fileDate) latestIntl = { file: f, fileDate: d };
    }
    if (BANCO_REGEX.test(name)) {
      if (!latestBanco || d > latestBanco.fileDate) latestBanco = { file: f, fileDate: d };
    }
    if (NAC_FACT_REGEX.test(name)) {
      if (!latestNacFact || d > latestNacFact.fileDate) latestNacFact = { file: f, fileDate: d };
    }
    if (INTL_FACT_REGEX.test(name)) {
      if (!latestIntlFact || d > latestIntlFact.fileDate) latestIntlFact = { file: f, fileDate: d };
    }
  }
  return { nac: latestNac, intl: latestIntl, banco: latestBanco, nacFact: latestNacFact, intlFact: latestIntlFact };
}

/** Extract YYYYMMDD integer from filename date patterns. */
function extractDateFromName_(name) {
  const s = String(name);

  let m = s.match(/(\d{2})[-_.](\d{2})[-_.](\d{4})/);
  if (m) {
    const dd = parseInt(m[1], 10), mm = parseInt(m[2], 10), yyyy = parseInt(m[3], 10);
    if (isValidDateParts_(yyyy, mm, dd)) return yyyy * 10000 + mm * 100 + dd;
  }

  m = s.match(/(\d{4})[-_.](\d{2})[-_.](\d{2})/);
  if (m) {
    const yyyy = parseInt(m[1], 10), mm = parseInt(m[2], 10), dd = parseInt(m[3], 10);
    if (isValidDateParts_(yyyy, mm, dd)) return yyyy * 10000 + mm * 100 + dd;
  }

  return null;
}

function isValidDateParts_(yyyy, mm, dd) {
  if (yyyy < 2000 || yyyy > 2100) return false;
  if (mm < 1 || mm > 12) return false;
  if (dd < 1 || dd > 31) return false;
  const d = new Date(Date.UTC(yyyy, mm - 1, dd));
  return d.getUTCFullYear() === yyyy && d.getUTCMonth() === (mm - 1) && d.getUTCDate() === dd;
}

function mergeWithTipo_(noFactData, factData) {
  const source = noFactData || factData;
  if (!source) return null;
  const header = [...source[0], "Tipo"];
  const merged = [header];
  if (noFactData) {
    for (let i = 1; i < noFactData.length; i++) {
      merged.push([...noFactData[i], "No Facturado"]);
    }
  }
  if (factData) {
    for (let i = 1; i < factData.length; i++) {
      merged.push([...factData[i], "Facturado"]);
    }
  }
  return merged;
}

/**
 * Convert .xls to Google Sheets via Drive API v3, extract movements.
 * Uses try/finally to ensure temp file cleanup even on parse failure.
 *
 * @param {File} xlsFile   Google Drive file
 * @param {string} currency "CLP" or "USD"
 * @returns {Array[]} 2D array with header row + data rows
 */
function extractMovementsFromXls_(xlsFile, currency) {
  const blob = xlsFile.getBlob();

  const converted = Drive.Files.create(
    { name: "tmp_convert_" + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );

  try {
    const tmpSS    = SpreadsheetApp.openById(converted.id);
    const tmpSheet = tmpSS.getSheets()[0];
    const values   = tmpSheet.getDataRange().getValues();

    const headerInfo = findHeader_(values);
    if (!headerInfo) {
      throw new Error("Header row not found (Fecha/Descripción/Monto).");
    }

    const { headerRow, colFecha, colDesc, colMonto } = headerInfo;
    const out = [];
    out.push(["Fecha", "Descripcion", currency === "CLP" ? "Monto_CLP" : "Monto_USD"]);

    for (let r = headerRow + 1; r < values.length; r++) {
      const fecha = values[r][colFecha];
      const desc  = values[r][colDesc];
      const monto = values[r][colMonto];

      if (!fecha && !desc && (monto === "" || monto === null)) break;
      if (!desc) continue;

      let numMonto;
      if (typeof monto === "number") {
        numMonto = monto;
      } else {
        const cleaned = String(monto != null ? monto : "").replace(/[^\d.-]/g, "");
        numMonto = parseFloat(cleaned);
      }
      if (!isFinite(numMonto)) continue;

      out.push([fecha, String(desc), numMonto]);
    }

    return out;
  } finally {
    // Always trash the temp file, even if parsing throws
    DriveApp.getFileById(converted.id).setTrashed(true);
  }
}

/** Scan rows for header containing fecha/descripción/monto columns. */
function findHeader_(values) {
  for (let r = 0; r < Math.min(values.length, 250); r++) {
    const row = values[r].map(v => String(v || "").toLowerCase().trim());
    const colFecha = row.findIndex(c => c === "fecha");
    const colDesc  = row.findIndex(c => c.includes("descrip"));
    const colMonto = row.findIndex(c => c.includes("monto"));
    if (colFecha !== -1 && colDesc !== -1 && colMonto !== -1) {
      return { headerRow: r, colFecha, colDesc, colMonto };
    }
  }
  return null;
}

/**
 * Convert Banco Estado .xlsx via Drive API and extract checking account movements.
 * Also reads the "Saldo Disponible" value from near the top of the file.
 *
 * @param {File} xlsFile  Google Drive file (Ultimos_Movimientos_Cuenta_Corriente_*.xlsx)
 * @returns {{ data: Array[], saldoDisponible: number }}
 */
function extractMovementsFromBancoXls_(xlsFile) {
  const blob = xlsFile.getBlob();
  const converted = Drive.Files.create(
    { name: "tmp_convert_banco_" + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );

  try {
    const tmpSS    = SpreadsheetApp.openById(converted.id);
    const tmpSheet = tmpSS.getSheets()[0];
    const values   = tmpSheet.getDataRange().getValues();

    // Read Saldo Disponible — scan first 15 rows for cell containing "saldo disponible"
    let saldoDisponible = 0;
    for (let r = 0; r < Math.min(values.length, 15); r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cellStr = String(values[r][c] || "").toLowerCase();
        if (cellStr.includes("saldo disponible")) {
          // The value is typically in the next column or the one after
          for (let cc = c + 1; cc < Math.min(values[r].length, c + 3); cc++) {
            const val = parseBancoAmount_(values[r][cc]);
            if (val !== 0) { saldoDisponible = val; break; }
          }
          // Also check the row below same column
          if (saldoDisponible === 0 && r + 1 < values.length) {
            saldoDisponible = parseBancoAmount_(values[r + 1][c]);
          }
          break;
        }
      }
      if (saldoDisponible !== 0) break;
    }

    // Find header row
    const headerInfo = findBancoHeader_(values);
    if (!headerInfo) {
      throw new Error("Banco header row not found (Fecha/Descripción/Cargos/Abonos).");
    }

    const { headerRow, colFecha, colDesc, colCargo, colAbono } = headerInfo;
    const out = [];
    out.push(["Fecha", "Descripcion", "Monto_CLP"]);

    for (let r = headerRow + 1; r < values.length; r++) {
      const fecha = values[r][colFecha];
      const desc  = values[r][colDesc];
      const cargo = parseBancoAmount_(values[r][colCargo]);
      const abono = parseBancoAmount_(values[r][colAbono]);

      // Skip subtotal/empty rows
      if (!fecha && !desc) break;
      const descStr = String(desc || "").toLowerCase().trim();
      if (!descStr || descStr.includes("subtotal")) continue;

      // Cargo is negative (debits), Abono is positive (credits)
      const monto = cargo + abono;
      out.push([fecha, String(desc), monto]);
    }

    return { data: out, saldoDisponible: saldoDisponible };
  } finally {
    DriveApp.getFileById(converted.id).setTrashed(true);
  }
}

/**
 * Scan rows for Banco Estado header containing fecha/descripcion/cargos/abonos.
 * @param {Array[]} values  2D array from sheet
 * @returns {{ headerRow, colFecha, colDesc, colCargo, colAbono }|null}
 */
function findBancoHeader_(values) {
  for (let r = 0; r < Math.min(values.length, 30); r++) {
    const row = values[r].map(v => String(v || "").toLowerCase().trim());
    const colFecha = row.findIndex(c => c === "fecha");
    const colDesc  = row.findIndex(c => c.includes("descripci") || c.includes("operaci"));
    const colCargo = row.findIndex(c => c.includes("cargo") || c.includes("cheque"));
    const colAbono = row.findIndex(c => c.includes("abono") || c.includes("dep\u00f3sito") || c.includes("deposito"));
    if (colFecha !== -1 && colDesc !== -1 && colCargo !== -1 && colAbono !== -1) {
      return { headerRow: r, colFecha, colDesc, colCargo, colAbono };
    }
  }
  return null;
}

/**
 * Parse a Banco Estado amount value (handles numbers, "0" strings, negatives).
 * @param {*} val
 * @returns {number}
 */
function parseBancoAmount_(val) {
  if (typeof val === "number") return isFinite(val) ? val : 0;
  const cleaned = String(val != null ? val : "").replace(/[^\d.-]/g, "");
  const num = parseFloat(cleaned);
  return isFinite(num) ? num : 0;
}

/**
 * Overwrite MOV_BANCO tab with banco movements + auto-categorization.
 * @param {Spreadsheet} ss
 * @param {Array[]} data       3-column data [Fecha, Descripcion, Monto_CLP]
 * @param {Array}   categories From readCategories_
 */
function overwriteBancoTab_(ss, data, categories) {
  const sh = ensureSheet_(ss, TAB_BANCO);
  sh.clear();

  const augmented = data.map((row, i) => {
    if (i === 0) return [...row, "Categoria"];
    const cat = categories ? categorizeMovement_(row[1], "BANCO", categories) : "Other";
    return [...row, cat];
  });

  sh.getRange(1, 1, augmented.length, augmented[0].length).setValues(augmented);
  formatDataTab_(sh, "CLP", false);
}

/**
 * Overwrite a live data tab with fresh data + auto-categorization + formatting.
 * Augments 3-column bank data with:
 *   - Column 4: Categoria (from CATEGORIAS keywords)
 *   - Column 5 (INTL only): Monto_CLP_est (= Monto_USD × USD_CLP)
 *
 * @param {Spreadsheet} ss
 * @param {string} tabName
 * @param {Array[]} data       3-column data [Fecha, Descripcion, Monto]
 * @param {string} currency    "CLP" or "USD"
 * @param {Array} categories   From readCategories_ (optional)
 * @param {Object} config      From readConfig_ (optional, needed for INTL CLP estimate)
 */
function overwriteTab_(ss, tabName, data, currency, categories, config) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error("Missing tab: " + tabName);
  sh.clear();

  const isIntl = currency === "USD";
  const usdClp = config ? config.usdClp : 0;
  const tipo = tabName === TAB_NAC ? "NAC" : "INTL";

  // Augment data with categories (and CLP estimate for INTL)
  const augmented = data.map((row, i) => {
    if (i === 0) {
      // Header row
      const hdr = [...row, "Categoria"];
      if (isIntl) hdr.push("Monto_CLP_est");
      return hdr;
    }
    const cat = categories ? categorizeMovement_(row[1], tipo, categories) : "Other";
    const augRow = [...row, cat];
    if (isIntl) augRow.push(row[2] * usdClp);
    return augRow;
  });

  sh.getRange(1, 1, augmented.length, augmented[0].length).setValues(augmented);
  formatDataTab_(sh, currency, false);
}

function fileMd5_(file) {
  const bytes  = file.getBlob().getBytes();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes);
  return digest.map(function(b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); }).join("");
}

function fileFingerprint_(file) {
  return file.getId() + "|" + fileMd5_(file);
}

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Rename old tab names to new emoji-prefixed names.
 * Safe to run multiple times — skips if old name doesn't exist or new name already exists.
 */
function migrateTabs_(ss) {
  for (const [oldName, newName] of Object.entries(TAB_RENAMES)) {
    if (oldName === newName) continue;
    const existing = ss.getSheetByName(oldName);
    if (existing && !ss.getSheetByName(newName)) {
      existing.setName(newName);
    }
  }
}

/**
 * Ensure all managed tabs exist, creating CONFIG and CATEGORIAS with defaults
 * if missing. Other tabs are created empty via ensureSheet_.
 */
function ensureAllTabs_(ss) {
  // Migrate old tab names → emoji names
  migrateTabs_(ss);

  // CONFIG — create with default financial parameters
  if (!ss.getSheetByName(TAB_CONFIG)) {
    const sh = ss.insertSheet(TAB_CONFIG);
    const defaults = [
      ["SALARIO",     0],
      ["HOUSING",     0],
      ["FAMILIA",     0],
      ["USD_CLP",   950],
      ["CLAUDE_PLAN", 0],
    ];
    sh.getRange(1, 1, defaults.length, 2).setValues(defaults);
    sh.getRange(1, 1, defaults.length, 1)
      .setFontWeight("bold")
      .setFontFamily(FONT_FAMILY)
      .setFontSize(10)
      .setBackground(THEME.BG_ALT_ROW);
    sh.autoResizeColumns(1, 2);
  }

  // CATEGORIAS — create with header + sample keywords
  if (!ss.getSheetByName(TAB_CATEGORIAS)) {
    const sh = ss.insertSheet(TAB_CATEGORIAS);
    const rows = [
      ["Tipo", "Keyword", "Categoria"],
      ["ALL",  "UBER",     "Transporte"],
      ["ALL",  "NETFLIX",  "Entretenimiento"],
      ["ALL",  "SPOTIFY",  "Entretenimiento"],
      ["NAC",  "SUPERMERCADO", "Alimentacion"],
      ["NAC",  "FARMACIA", "Salud"],
      ["INTL", "OPENAI",   "Tech"],
    ];
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
    styleHeaderRow_(sh, 3, false);
    sh.autoResizeColumns(1, 3);
  }

  // Remaining tabs — create empty if missing
  ensureSheet_(ss, TAB_NAC);
  ensureSheet_(ss, TAB_INTL);
  ensureSheet_(ss, TAB_BANCO);
  ensureSheet_(ss, TAB_CUOTAS);
  ensureSheet_(ss, TAB_RESUMEN_CAT);
  ensureSheet_(ss, PREVIEW_TAB);
}

function moveFilesToImported_(folder, files) {
  let dest;
  const iter = folder.getFoldersByName("_imported");
  if (iter.hasNext()) {
    dest = iter.next();
  } else {
    dest = folder.createFolder("_imported");
  }
  for (const f of files) {
    if (f) f.moveTo(dest);
  }
}

/** Sum a column (1-indexed) from row 2 onward. */
function sumColumn_(sheet, colNumber) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const vals = sheet.getRange(2, colNumber, last - 1, 1).getValues();
  let s = 0;
  for (const [v] of vals) {
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^\d.-]/g, ""));
    if (isFinite(n)) s += n;
  }
  return s;
}

/** Sum only positive values in a column (1-indexed) from row 2 onward. */
function sumPositiveColumn_(sheet, colNumber) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const vals = sheet.getRange(2, colNumber, last - 1, 1).getValues();
  let s = 0;
  for (const [v] of vals) {
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^\d.-]/g, ""));
    if (isFinite(n) && n > 0) s += n;
  }
  return s;
}

/** Sum a column (0-indexed) from array row 1 onward (skipping header). */
function sumArrayCol_(arr, idx) {
  let s = 0;
  for (let i = 1; i < arr.length; i++) {
    const v = arr[i][idx];
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^\d.-]/g, ""));
    if (isFinite(n)) s += n;
  }
  return s;
}

/******************************************************************************
 * §16  GEMINI AI EXTRACTION
 ******************************************************************************/

function callGemini_(prompt, apiKey) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/"
    + GEMINI_MODEL + ":generateContent?key=" + apiKey;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0,
    },
  };
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error("Gemini API error " + code + ": " + resp.getContentText().substring(0, 200));
  }
  const body = JSON.parse(resp.getContentText());
  const text = body.candidates[0].content.parts[0].text;
  return JSON.parse(text);
}

function convertFileToArray_(file) {
  const blob = file.getBlob();
  const resource = { name: "[temp] " + file.getName(), mimeType: MimeType.GOOGLE_SHEETS };
  const converted = Drive.Files.create(resource, blob, { convert: true });
  try {
    const tmpSS = SpreadsheetApp.openById(converted.id);
    const sheet = tmpSS.getSheets()[0];
    if (sheet.getLastRow() === 0) return [];
    return sheet.getDataRange().getValues();
  } finally {
    DriveApp.getFileById(converted.id).setTrashed(true);
  }
}

function classifyAndExtractWithGemini_(values, apiKey) {
  const maxRows = 500;
  const rows = values.slice(0, maxRows);
  const tsv = rows.map(r => r.map(c => String(c)).join("\t")).join("\n");

  const prompt = `You are analyzing a Chilean bank statement exported as a spreadsheet.

SPREADSHEET DATA (tab-separated):
${tsv}

TASK:
1. Classify this file as exactly one of: "cc_nacional", "cc_internacional", "banco", "unknown"
   - cc_nacional: Chilean credit card statement in CLP (pesos)
   - cc_internacional: International credit card statement in USD or foreign currency
   - banco: Checking/savings account statement (has debits and credits, saldo disponible)
   - unknown: Cannot determine
2. Extract ALL transaction/movement rows as a JSON array
3. For "banco" type: also extract the "Saldo Disponible" (available balance) value

RESPONSE (strict JSON, no markdown):
{
  "type": "cc_nacional",
  "currency": "CLP",
  "saldo_disponible": null,
  "movements": [
    {"fecha": "09/03/2026", "descripcion": "COMPRA SUPERMERCADO", "monto": 45000}
  ]
}

RULES:
- For banco type: negative monto = cargo/debit, positive monto = abono/credit
- For CC types: monto is the charge amount (always positive)
- Dates must be DD/MM/YYYY format
- Clean numeric values: no thousand separators, use period for decimals
- Include ALL data rows, skip headers/totals/subtotals
- saldo_disponible is null for CC types`;

  return callGemini_(prompt, apiKey);
}

function findAndExtractAllFilesGemini_(folder, apiKey) {
  const files = folder.getFiles();
  const results = { nac: null, intl: null, banco: null, nacFact: null, intlFact: null };
  const dates = { nac: 0, intl: 0, banco: 0, nacFact: 0, intlFact: 0 };
  let callCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();
    if (name.startsWith(".") || name.startsWith("~$")) continue;

    const mime = file.getMimeType();
    if (mime.indexOf("spreadsheet") < 0 && mime.indexOf("excel") < 0
        && mime.indexOf("csv") < 0 && !name.endsWith(".xls") && !name.endsWith(".xlsx")) {
      continue;
    }

    const values = convertFileToArray_(file);
    if (!values || values.length < 2) continue;

    // Rate-limit: wait between Gemini calls to avoid 429
    if (callCount > 0) Utilities.sleep(4000);
    callCount++;

    const gemini = classifyAndExtractWithGemini_(values, apiKey);
    if (!gemini || gemini.type === "unknown" || !gemini.movements) continue;

    const fileDate = extractDateFromName_(file.getName()) || file.getLastUpdated().getTime();
    const originalName = file.getName();
    const isFacturado = /Facturados/i.test(originalName) && !/NoFacturados/i.test(originalName);
    const typeKey = gemini.type === "cc_nacional"      ? (isFacturado ? "nacFact" : "nac")
                  : gemini.type === "cc_internacional" ? (isFacturado ? "intlFact" : "intl")
                  : gemini.type === "banco"            ? "banco"
                  : null;
    if (!typeKey) continue;

    if (fileDate >= dates[typeKey]) {
      dates[typeKey] = fileDate;
      const currency = gemini.currency || (typeKey === "intl" ? "USD" : "CLP");
      const montoCol = currency === "USD" ? "Monto_USD" : "Monto_CLP";

      const data = [["Fecha", "Descripcion", montoCol]];
      for (const m of gemini.movements) {
        const fecha = m.fecha || "";
        const desc = m.descripcion || "";
        const monto = typeof m.monto === "number" ? m.monto : parseFloat(String(m.monto).replace(/[^\d.-]/g, "")) || 0;
        data.push([fecha, desc, monto]);
      }

      results[typeKey] = {
        file: file,
        fileDate: fileDate,
        extractedData: data,
        saldoDisponible: gemini.saldo_disponible || 0,
      };
    }
  }
  return results;
}
