const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

function normalizeHebrew(value) {
  return String(value ?? "")
    .replace(/[\u0591-\u05C7]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function inferColumnType(rows, header) {
  const sample = rows.slice(0, 300).map(r => r[header]).filter(v => v !== "" && v != null);

  if (!sample.length) return "empty";

  let numeric = 0;
  let dates = 0;

  for (const value of sample) {
    const text = String(value).trim().replace(/,/g, "");
    if (text !== "" && !Number.isNaN(Number(text))) numeric++;
    if (!Number.isNaN(Date.parse(text))) dates++;
  }

  if (numeric / sample.length >= 0.7) return "number";
  if (dates / sample.length >= 0.7) return "date";
  return "text";
}

function buildSheetSummary(rows, sheetName) {
  const headers = Object.keys(rows[0] || {});
  const columnMeta = headers.map((header) => ({
    name: header,
    type: inferColumnType(rows, header)
  }));

  return {
    name: sheetName,
    rowCount: rows.length,
    columnCount: headers.length,
    columns: headers,
    columnMeta,
    preview: rows.slice(0, 30)
  };
}

function readWorkbookFromFile(filePath) {
  const lower = filePath.toLowerCase();

  if (lower.endsWith(".csv")) {
    const text = fs.readFileSync(filePath, "utf8");
    const workbook = XLSX.read(text, { type: "string" });
    const firstSheet = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheet];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map((row) => {
      const normalized = {};
      Object.keys(row).forEach((key) => {
        normalized[normalizeHebrew(key)] =
          typeof row[key] === "string" ? normalizeHebrew(row[key]) : row[key];
      });
      return normalized;
    });

    return {
      fileName: path.basename(filePath),
      sheets: [buildSheetSummary(rows, "CSV")]
    };
  }

  const workbook = XLSX.readFile(filePath, { cellDates: true });

  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map((row) => {
      const normalized = {};
      Object.keys(row).forEach((key) => {
        normalized[normalizeHebrew(key)] =
          typeof row[key] === "string" ? normalizeHebrew(row[key]) : row[key];
      });
      return normalized;
    });

    return buildSheetSummary(rows, sheetName);
  });

  return {
    fileName: path.basename(filePath),
    sheets
  };
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1500,
    height: 980,
    minWidth: 1100,
    minHeight: 760,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });

  win.loadFile("index.html");
}

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

ipcMain.handle("pick-file", async () => {
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [
      { name: "Spreadsheet files", extensions: ["xlsx", "xls", "csv"] }
    ]
  });

  if (result.canceled || !result.filePaths.length) {
    return { canceled: true };
  }

  return {
    canceled: false,
    filePath: result.filePaths[0]
  };
});

ipcMain.handle("read-workbook", async (_event, filePath) => {
  try {
    return {
      ok: true,
      data: readWorkbookFromFile(filePath)
    };
  } catch (error) {
    return {
      ok: false,
      error: error.message || "Failed to read workbook"
    };
  }
});

ipcMain.handle("save-formulas-workbook", async (_event, payload) => {
  try {
    const { sourcePreview, formulas, defaultName } = payload;

    const result = await dialog.showSaveDialog({
      defaultPath: defaultName || "excel_hebrew_desktop_output.xlsx",
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }]
    });

    if (result.canceled || !result.filePath) {
      return { ok: false, canceled: true };
    }

    const wb = XLSX.utils.book_new();

    const previewSheet = XLSX.utils.json_to_sheet(sourcePreview || []);
    XLSX.utils.book_append_sheet(wb, previewSheet, "Preview");

    const formulaRows = (formulas || []).map((item) => ({
      Formula_Name: item.title,
      Formula: item.formula,
      Explanation: item.explanation
    }));
    const formulasSheet = XLSX.utils.json_to_sheet(formulaRows);
    XLSX.utils.book_append_sheet(wb, formulasSheet, "Formulas");

    XLSX.writeFile(wb, result.filePath);

    return { ok: true, savedTo: result.filePath };
  } catch (error) {
    return {
      ok: false,
      error: error.message || "Failed to save workbook"
    };
  }
});
