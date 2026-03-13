self.importScripts("https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js");

function normalizeHebrew(value) {
  return String(value ?? "")
    .replace(/[\u0591-\u05C7]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function inferColumnType(rows, header) {
  const sample = rows.slice(0, 200).map(r => r[header]).filter(v => v !== "" && v != null);

  if (!sample.length) return "empty";

  let numeric = 0;
  let dates = 0;

  for (const val of sample) {
    const text = String(val).trim().replace(/,/g, "");
    if (text !== "" && !Number.isNaN(Number(text))) numeric++;
    if (!Number.isNaN(Date.parse(text))) dates++;
  }

  if (numeric / sample.length >= 0.7) return "number";
  if (dates / sample.length >= 0.7) return "date";
  return "text";
}

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");
  if (!lines.length) return [];

  const rows = lines.map(line => {
    const values = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const next = line[i + 1];

      if (char === '"' && inQuotes && next === '"') {
        current += '"';
        i++;
      } else if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === "," && !inQuotes) {
        values.push(current);
        current = "";
      } else {
        current += char;
      }
    }

    values.push(current);
    return values.map(v => normalizeHebrew(v));
  });

  const headers = rows[0].map((h, i) => h || `Column${i + 1}`);

  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] ?? "";
    });
    return obj;
  });
}

function buildSheetSummary(rows, sheetName) {
  const headers = Object.keys(rows[0] || {});
  const columnMeta = headers.map(header => ({
    name: header,
    type: inferColumnType(rows, header)
  }));

  return {
    name: sheetName,
    rowCount: rows.length,
    columnCount: headers.length,
    columns: headers,
    columnMeta,
    preview: rows.slice(0, 20)
  };
}

self.onmessage = async (event) => {
  const { type, payload } = event.data;

  try {
    if (type === "parse-file") {
      const { fileName, buffer } = payload;
      const lower = fileName.toLowerCase();

      if (lower.endsWith(".csv")) {
        const decoder = new TextDecoder("utf-8");
        const text = decoder.decode(buffer);
        const rows = parseCSV(text);
        const summary = buildSheetSummary(rows, "CSV");
        self.postMessage({
          type: "file-parsed",
          payload: {
            sheets: [summary],
            fileName
          }
        });
        return;
      }

      const workbook = XLSX.read(buffer, { type: "array" });
      const sheets = workbook.SheetNames.map((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" }).map((row) => {
          const normalized = {};
          Object.keys(row).forEach((key) => {
            normalized[normalizeHebrew(key)] = typeof row[key] === "string"
              ? normalizeHebrew(row[key])
              : row[key];
          });
          return normalized;
        });
        return buildSheetSummary(rows, sheetName);
      });

      self.postMessage({
        type: "file-parsed",
        payload: {
          sheets,
          fileName
        }
      });
    }
  } catch (error) {
    self.postMessage({
      type: "parse-error",
      payload: { message: error.message || "שגיאה בקריאת הקובץ" }
    });
  }
};
