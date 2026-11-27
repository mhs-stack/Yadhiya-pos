function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("YADHIYA POS & Manajemen Toko v1.6")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty("SHEET_ID");
  if (sheetId) {
    return SpreadsheetApp.openById(sheetId);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error("SHEET_ID belum diatur di Script Properties.");
}

function ensureSheets() {
  const ss = getSpreadsheet();
  const required = ["Users", "Materials", "BomRecipes", "Sales", "Expenses", "Journals"];
  required.forEach((name) => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });
  return ss;
}

function syncSnapshotToSheets(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = ensureSheets();
    const data = typeof payload === "string" ? JSON.parse(payload) : (payload || {});
    const timestamp = new Date().toISOString();

    const metaSheet = ss.getSheetByName("Journals");
    if (metaSheet && Array.isArray(data.journals)) {
      metaSheet.clear();
      metaSheet.appendRow(["ID", "Date", "Memo", "LinesJSON", "SyncedAt"]);
      data.journals.forEach((j) => {
        metaSheet.appendRow([
          j.id || "",
          j.date || "",
          j.memo || "",
          JSON.stringify(j.lines || []),
          timestamp
        ]);
      });
    }

    const salesSheet = ss.getSheetByName("Sales");
    if (salesSheet && Array.isArray(data.sales)) {
      salesSheet.clear();
      salesSheet.appendRow(["ID", "Date", "Cashier", "Total", "COGS", "MfgElectricity", "MfgOperational", "SyncedAt"]);
      data.sales.forEach((s) => {
        salesSheet.appendRow([
          s.id || "",
          s.date || "",
          s.cashierName || "",
          Number(s.total || 0),
          Number(s.cogs || 0),
          Number(s.mfgElectricity || 0),
          Number(s.mfgOperational || 0),
          timestamp
        ]);
      });
    }

    return { ok: true, syncedAt: timestamp };
  } finally {
    lock.releaseLock();
  }
}
