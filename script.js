const userPrompt = document.getElementById("userPrompt");
const generateBtn = document.getElementById("generateBtn");
const multiBtn = document.getElementById("multiBtn");
const exportBtn = document.getElementById("exportBtn");
const clearBtn = document.getElementById("clearBtn");
const clearHistoryBtn = document.getElementById("clearHistoryBtn");
const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");

const errorBox = document.getElementById("errorBox");
const statusBox = document.getElementById("statusBox");
const progressBox = document.getElementById("progressBox");

const fileSummarySection = document.getElementById("fileSummarySection");
const fileSummary = document.getElementById("fileSummary");
const columnsSection = document.getElementById("columnsSection");
const previewSection = document.getElementById("previewSection");
const singleResult = document.getElementById("singleResult");
const multiResult = document.getElementById("multiResult");

const formulaText = document.getElementById("formulaText");
const explanationText = document.getElementById("explanationText");
const exampleText = document.getElementById("exampleText");
const tipsList = document.getElementById("tipsList");

const formulaGrid = document.getElementById("formulaGrid");
const copyBtn = document.getElementById("copyBtn");
const copyAllBtn = document.getElementById("copyAllBtn");

const historyList = document.getElementById("historyList");
const examplesList = document.getElementById("examplesList");
const defaultColumnInput = document.getElementById("defaultColumn");
const quickTags = document.getElementById("quickTags");
const columnsList = document.getElementById("columnsList");
const previewTable = document.getElementById("previewTable");
const smartSuggestions = document.getElementById("smartSuggestions");

const EXAMPLES = [
  "חשב ממוצע של סכום",
  "ספור כמה פעמים Approved מופיע בסטטוס",
  "סכם הכנסות רק אם סטטוס הוא Approved",
  "מצא את הערך הגבוה ביותר בעמודת עלות",
  "מצא את הערך הנמוך ביותר בעמודת מחיר",
  "חפש מחיר לפי קוד מוצר",
  "ספור כמה שורות יש שבהן סטטוס הוא Approved וגם פעיל הוא Yes"
];

let workbookState = {
  fileName: "",
  sheets: [],
  activeSheetIndex: -1
};

const worker = new Worker("worker.js");

function showError(message) {
  errorBox.textContent = message;
  errorBox.classList.remove("hidden");
}

function hideError() {
  errorBox.classList.add("hidden");
  errorBox.textContent = "";
}

function showStatus(message) {
  statusBox.textContent = message;
  statusBox.classList.remove("hidden");
}

function hideStatus() {
  statusBox.classList.add("hidden");
  statusBox.textContent = "";
}

function showProgress(message) {
  progressBox.textContent = message;
  progressBox.classList.remove("hidden");
}

function hideProgress() {
  progressBox.classList.add("hidden");
  progressBox.textContent = "";
}

function getActiveSheet() {
  if (workbookState.activeSheetIndex < 0) return null;
  return workbookState.sheets[workbookState.activeSheetIndex] || null;
}

function getColumns() {
  return getActiveSheet()?.columns || [];
}

function getColumnMeta() {
  return getActiveSheet()?.columnMeta || [];
}

function getPreviewRows() {
  return getActiveSheet()?.preview || [];
}

function containsAny(text, words) {
  return words.some(word => text.includes(word));
}

function normalizeText(value) {
  return String(value ?? "")
    .replace(/[\u0591-\u05C7]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeColumnName(value) {
  return String(value ?? "").trim();
}

function detectCell(text) {
  const match = text.match(/([A-Z]+\d+)/i);
  return match ? match[1].toUpperCase() : null;
}

function excelFormula(name, args) {
  return `=${name.toUpperCase()}(${args})`;
}

function clearSingleResult() {
  singleResult.classList.add("hidden");
  formulaText.textContent = "";
  explanationText.textContent = "";
  exampleText.textContent = "";
  tipsList.innerHTML = "";
}

function clearMultiResult() {
  multiResult.classList.add("hidden");
  formulaGrid.innerHTML = "";
}

function clearFileUI() {
  fileSummarySection.classList.add("hidden");
  columnsSection.classList.add("hidden");
  previewSection.classList.add("hidden");
  fileSummary.innerHTML = "";
  columnsList.innerHTML = "";
  previewTable.innerHTML = "";
  smartSuggestions.innerHTML = `<div class="empty-state">העלי קובץ כדי לקבל הצעות חכמות.</div>`;
  sheetSelect.innerHTML = `<option value="">אין גיליון</option>`;
  sheetSelect.disabled = true;
}

function clearAll() {
  userPrompt.value = "";
  hideError();
  hideStatus();
  hideProgress();
  clearSingleResult();
  clearMultiResult();
}

function findColumnByPrompt(prompt) {
  const promptNorm = normalizeText(prompt);
  const columns = getColumns();

  for (const col of columns) {
    if (promptNorm.includes(normalizeText(col))) return col;
  }

  const aliases = [
    { test: /סטטוס|status|מצב/, keys: ["status", "סטטוס", "מצב"] },
    { test: /סכום|amount|total|sum/, keys: ["amount", "total", "sum", "סכום", "סהכ"] },
    { test: /מחיר|price/, keys: ["price", "מחיר"] },
    { test: /עלות|cost/, keys: ["cost", "עלות"] },
    { test: /הכנסה|revenue/, keys: ["revenue", "income", "הכנסה"] },
    { test: /פעיל|active/, keys: ["active", "פעיל"] },
    { test: /ציון|score/, keys: ["score", "ציון"] },
    { test: /קוד|code|מזהה|id/, keys: ["code", "id", "מזהה", "קוד"] }
  ];

  for (const alias of aliases) {
    if (alias.test.test(promptNorm)) {
      const found = columns.find(col => alias.keys.some(key => normalizeText(col).includes(key)));
      if (found) return found;
    }
  }

  return null;
}

function getDefaultColumn() {
  const promptDefault = normalizeColumnName(defaultColumnInput.value);
  return findColumnByPrompt(promptDefault) || promptDefault || "A";
}

function getContext(prompt) {
  const column = findColumnByPrompt(prompt) || getDefaultColumn();
  const cell = detectCell(prompt) || "A2";
  return {
    text: normalizeText(prompt),
    column,
    cell
  };
}

function columnRef(columnName) {
  const clean = String(columnName).trim();
  if (/^[A-Z]{1,3}$/.test(clean.toUpperCase())) {
    const upper = clean.toUpperCase();
    return `${upper}:${upper}`;
  }
  return `[${clean}]`;
}

function detectConditionValue(text) {
  if (containsAny(text, ["approved", "מאושר", "אושר"])) return "Approved";
  if (containsAny(text, ["yes", "כן"])) return "Yes";
  if (containsAny(text, ["no", "לא"])) return "No";
  return "Approved";
}

function makeResult(title, formula, explanation, example, tips) {
  return { title, formula, explanation, example, tips };
}

function getBestNumericColumn() {
  const numeric = getColumnMeta().find(col => col.type === "number");
  return numeric?.name || getColumns()[0] || getDefaultColumn();
}

function getBestStatusColumn() {
  const byName = getColumns().find(c => /status|state|סטטוס|מצב/i.test(c));
  return byName || getColumns()[1] || getDefaultColumn();
}

function getBestIdColumn() {
  const byName = getColumns().find(c => /code|id|מזהה|קוד/i.test(c));
  return byName || getColumns()[0] || "Code";
}

function buildSingleFormula(prompt) {
  const ctx = getContext(prompt);
  const text = ctx.text;
  const conditionValue = detectConditionValue(text);
  const selectedCol = columnRef(ctx.column);
  const numericCol = columnRef(getBestNumericColumn());
  const statusCol = columnRef(getBestStatusColumn());
  const idCol = columnRef(getBestIdColumn());

  if (containsAny(text, ["countifs", "שני תנאים", "כמה תנאים", "וגם", "גם"])) {
    const secondCol = columnRef(getColumns()[2] || getColumns()[1] || "Active");
    return makeResult(
      "COUNTIFS",
      excelFormula("COUNTIFS", `${statusCol},"Approved",${secondCol},"Yes"`),
      "הנוסחה סופרת שורות שעומדות בשני תנאים במקביל.",
      "מתאים לדוחות גדולים של סטטוסים והרשאות.",
      [
        "אפשר להחליף את שמות העמודות לפי הגיליון שנטען.",
        "בקבצים גדולים עדיף לעבוד לפי עמודות מדויקות."
      ]
    );
  }

  if (containsAny(text, ["sumif", "סכם רק אם", "סכום רק אם", "סכום בתנאי"])) {
    return makeResult(
      "SUMIF",
      excelFormula("SUMIF", `${statusCol},"Approved",${numericCol}`),
      "הנוסחה מסכמת ערכים רק בשורות שבהן הסטטוס הוא Approved.",
      "מתאים במיוחד לקבצי חברות עם עמודות סכום וסטטוס.",
      [
        "כדאי לבחור עמודה מספרית כעמודת סכום.",
        "אפשר לשנות את ערך התנאי מ-Approved לכל ערך אחר."
      ]
    );
  }

  if (containsAny(text, ["חבר", "סכום", "סכם", "סכימה", "sum"])) {
    return makeResult(
      "SUM",
      excelFormula("SUM", selectedCol),
      `הנוסחה מחברת את כל הערכים בעמודה ${ctx.column}.`,
      "שימושי לסכומים, הכנסות, שעות, מלאי ותקציבים.",
      [
        "בקבצים גדולים מומלץ לבחור עמודה מספרית בלבד.",
        "אפשר לעבוד עם שם עמודה בעברית."
      ]
    );
  }

  if (containsAny(text, ["ממוצע", "average"])) {
    return makeResult(
      "AVERAGE",
      excelFormula("AVERAGE", selectedCol),
      `הנוסחה מחשבת ממוצע של הערכים בעמודה ${ctx.column}.`,
      "מתאים לדוחות KPI, ציונים, מחירים, עלויות והכנסות.",
      [
        "תאים ריקים בדרך כלל לא נכללים בחישוב.",
        "מומלץ להשתמש בעמודה מספרית."
      ]
    );
  }

  if (containsAny(text, ["הכי גבוה", "גבוה ביותר", "מקסימום", "max"])) {
    return makeResult(
      "MAX",
      excelFormula("MAX", selectedCol),
      `הנוסחה מחזירה את הערך הגבוה ביותר בעמודה ${ctx.column}.`,
      "טוב למציאת שיאים בדוחות ארוכים.",
      [
        "לעמודה מספרית בלבד.",
        "למינימום משתמשים ב-MIN."
      ]
    );
  }

  if (containsAny(text, ["הכי נמוך", "נמוך ביותר", "מינימום", "min"])) {
    return makeResult(
      "MIN",
      excelFormula("MIN", selectedCol),
      `הנוסחה מחזירה את הערך הנמוך ביותר בעמודה ${ctx.column}.`,
      "טוב למציאת ערך תחתון בדוחות ארוכים.",
      [
        "לעמודה מספרית בלבד.",
        "למקסימום משתמשים ב-MAX."
      ]
    );
  }

  if (containsAny(text, ["ספור", "ספירה", "כמה פעמים", "countif", "count"])) {
    return makeResult(
      "COUNTIF",
      excelFormula("COUNTIF", `${selectedCol},"${conditionValue}"`),
      `הנוסחה סופרת כמה פעמים הערך "${conditionValue}" מופיע בעמודה ${ctx.column}.`,
      "מתאים במיוחד לקבצים בעברית עם עמודות סטטוס.",
      [
        "אפשר להחליף את הערך לחיפוש לכל מילה או תנאי אחר.",
        'אפשר גם להשתמש בתנאי כמו ">100".'
      ]
    );
  }

  if (containsAny(text, ["בדוק אם", "אם", "גדול", "קטן", "שווה", "if"])) {
    return makeResult(
      "IF",
      excelFormula("IF", `${ctx.cell}>100,"כן","לא"`),
      `הנוסחה בודקת אם הערך בתא ${ctx.cell} גדול מ-100.`,
      "מתאים לתנאי סף, אישורים, בקרה וחריגים.",
      [
        "אפשר לשנות את המספר 100.",
        "אפשר לשנות את הטקסט כן/לא."
      ]
    );
  }

  if (containsAny(text, ["חפש", "חיפוש", "קוד מוצר", "מצא מחיר", "lookup", "xlookup", "vlookup"])) {
    return makeResult(
      "XLOOKUP",
      excelFormula("XLOOKUP", `A2,${idCol},${numericCol},"לא נמצא"`),
      "הנוסחה מחפשת ערך לפי מזהה ומחזירה ערך מתאים מעמודה אחרת.",
      "מתאים למחיר לפי קוד מוצר, שם לקוח לפי מזהה, או סטטוס לפי מספר הזמנה.",
      [
        "בקבצים גדולים זה שימושי במיוחד לחיפושי רוחב.",
        "אם אין XLOOKUP אצלך, אפשר להחליף ל-VLOOKUP."
      ]
    );
  }

  return makeResult(
    "ברירת מחדל",
    excelFormula("SUM", columnRef(getBestNumericColumn())),
    "לא זוהתה בקשה מדויקת, לכן הוחזרה נוסחת ברירת מחדל לעמודה מספרית.",
    "נסי לכתוב: חשב ממוצע של סכום, או ספור כמה פעמים Approved מופיע בסטטוס.",
    [
      "כתבי את שם העמודה כפי שהוא מופיע בקובץ.",
      "בקבצים ארוכים זה משפר את הדיוק."
    ]
  );
}

function buildMultipleFormulas(prompt) {
  const ctx = getContext(prompt);
  const col = columnRef(ctx.column);
  const statusCol = columnRef(getBestStatusColumn());
  const numericCol = columnRef(getBestNumericColumn());
  const idCol = columnRef(getBestIdColumn());

  return [
    makeResult("SUM", excelFormula("SUM", numericCol), "סכום של עמודה מספרית.", "", []),
    makeResult("AVERAGE", excelFormula("AVERAGE", numericCol), "ממוצע של עמודה מספרית.", "", []),
    makeResult("MAX", excelFormula("MAX", numericCol), "מקסימום של עמודה מספרית.", "", []),
    makeResult("MIN", excelFormula("MIN", numericCol), "מינימום של עמודה מספרית.", "", []),
    makeResult("COUNTIF", excelFormula("COUNTIF", `${statusCol},"Approved"`), "ספירת Approved בעמודת סטטוס.", "", []),
    makeResult("IF", excelFormula("IF", `${ctx.cell}>100,"כן","לא"`), "בדיקת תנאי בסיסית.", "", []),
    makeResult("SUMIF", excelFormula("SUMIF", `${statusCol},"Approved",${numericCol}`), "סכום לפי תנאי.", "", []),
    makeResult("COUNTIFS", excelFormula("COUNTIFS", `${statusCol},"Approved",${statusCol},"Approved"`), "ספירה לפי תנאים.", "", []),
    makeResult("XLOOKUP", excelFormula("XLOOKUP", `A2,${idCol},${col},"לא נמצא"`), "חיפוש לפי מזהה.", "", [])
  ];
}

function showSingleResult(data) {
  clearMultiResult();
  formulaText.textContent = data.formula;
  explanationText.textContent = data.explanation;
  exampleText.textContent = data.example;
  tipsList.innerHTML = "";

  data.tips.forEach(tip => {
    const li = document.createElement("li");
    li.textContent = tip;
    tipsList.appendChild(li);
  });

  singleResult.classList.remove("hidden");
}

function showMultiResults(items) {
  clearSingleResult();
  formulaGrid.innerHTML = "";

  items.forEach(item => {
    const card = document.createElement("div");
    card.className = "formula-card";
    card.innerHTML = `
      <h3>${item.title}</h3>
      <p>${item.explanation}</p>
      <pre>${item.formula}</pre>
    `;
    formulaGrid.appendChild(card);
  });

  multiResult.classList.remove("hidden");
}

function saveHistory(prompt) {
  const history = JSON.parse(localStorage.getItem("excelAssistantHistory") || "[]");
  const updated = [prompt, ...history.filter(item => item !== prompt)].slice(0, 8);
  localStorage.setItem("excelAssistantHistory", JSON.stringify(updated));
  renderHistory();
}

function renderHistory() {
  const history = JSON.parse(localStorage.getItem("excelAssistantHistory") || "[]");
  historyList.innerHTML = "";

  if (!history.length) {
    historyList.innerHTML = `<div class="empty-state">אין עדיין היסטוריה.</div>`;
    return;
  }

  history.forEach(item => {
    const div = document.createElement("div");
    div.className = "history-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    historyList.appendChild(div);
  });
}

function renderExamples() {
  examplesList.innerHTML = "";

  EXAMPLES.forEach(item => {
    const div = document.createElement("div");
    div.className = "example-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    examplesList.appendChild(div);
  });
}

function renderFileSummary() {
  const active = getActiveSheet();
  if (!active) {
    fileSummarySection.classList.add("hidden");
    return;
  }

  fileSummary.innerHTML = `
    <div class="summary-card"><strong>קובץ</strong><span>${workbookState.fileName}</span></div>
    <div class="summary-card"><strong>גיליון</strong><span>${active.name}</span></div>
    <div class="summary-card"><strong>שורות</strong><span>${active.rowCount}</span></div>
    <div class="summary-card"><strong>עמודות</strong><span>${active.columnCount}</span></div>
  `;

  fileSummarySection.classList.remove("hidden");
}

function renderColumns() {
  const columns = getColumns();
  const meta = getColumnMeta();
  columnsList.innerHTML = "";

  if (!columns.length) {
    columnsSection.classList.add("hidden");
    return;
  }

  columns.forEach(col => {
    const info = meta.find(m => m.name === col);
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "column-chip";
    chip.textContent = info ? `${col} (${info.type})` : col;
    chip.addEventListener("click", () => {
      userPrompt.value = `חשב ממוצע של ${col}`;
      defaultColumnInput.value = col;
      userPrompt.focus();
    });
    columnsList.appendChild(chip);
  });

  columnsSection.classList.remove("hidden");
}

function renderPreview() {
  const rows = getPreviewRows();
  const columns = getColumns();

  previewTable.innerHTML = "";

  if (!rows.length || !columns.length) {
    previewSection.classList.add("hidden");
    return;
  }

  const thead = document.createElement("thead");
  const trHead = document.createElement("tr");

  columns.forEach(col => {
    const th = document.createElement("th");
    th.textContent = col;
    trHead.appendChild(th);
  });

  thead.appendChild(trHead);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  rows.slice(0, 10).forEach(row => {
    const tr = document.createElement("tr");
    columns.forEach(col => {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewTable.appendChild(tbody);
  previewSection.classList.remove("hidden");
}

function renderSmartSuggestions() {
  smartSuggestions.innerHTML = "";

  const columns = getColumns();
  const meta = getColumnMeta();

  if (!columns.length) {
    smartSuggestions.innerHTML = `<div class="empty-state">העלי קובץ כדי לקבל הצעות חכמות.</div>`;
    return;
  }

  const suggestions = [];
  const numericColumns = meta.filter(c => c.type === "number").map(c => c.name);
  const statusColumns = columns.filter(c => /status|state|סטטוס|מצב/i.test(c));
  const idColumns = columns.filter(c => /id|code|קוד|מזהה/i.test(c));

  numericColumns.forEach(col => {
    suggestions.push(`חשב ממוצע של ${col}`);
    suggestions.push(`חבר את כל הערכים ב${col}`);
    suggestions.push(`מצא את הערך הגבוה ביותר ב${col}`);
  });

  statusColumns.forEach(col => {
    suggestions.push(`ספור כמה פעמים Approved מופיע ב${col}`);
  });

  if (numericColumns.length && statusColumns.length) {
    suggestions.push(`סכם את ${numericColumns[0]} רק אם ${statusColumns[0]} הוא Approved`);
  }

  if (idColumns.length && numericColumns.length) {
    suggestions.push(`חפש ${numericColumns[0]} לפי ${idColumns[0]}`);
  }

  const unique = [...new Set(suggestions)].slice(0, 12);

  unique.forEach(item => {
    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    smartSuggestions.appendChild(div);
  });

  if (!unique.length) {
    smartSuggestions.innerHTML = `<div class="empty-state">לא נמצאו הצעות אוטומטיות.</div>`;
  }
}

function rebuildSheetSelect() {
  sheetSelect.innerHTML = "";

  if (!workbookState.sheets.length) {
    sheetSelect.innerHTML = `<option value="">אין גיליון</option>`;
    sheetSelect.disabled = true;
    return;
  }

  workbookState.sheets.forEach((sheet, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = sheet.name;
    if (index === workbookState.activeSheetIndex) option.selected = true;
    sheetSelect.appendChild(option);
  });

  sheetSelect.disabled = false;
}

function renderActiveSheet() {
  renderFileSummary();
  renderColumns();
  renderPreview();
  renderSmartSuggestions();

  const firstColumn = getColumns()[0];
  if (firstColumn) {
    defaultColumnInput.value = firstColumn;
  }
}

function handleParsedFile(fileName, sheets) {
  workbookState = {
    fileName,
    sheets,
    activeSheetIndex: sheets.length ? 0 : -1
  };

  rebuildSheetSelect();
  renderActiveSheet();
  showStatus(`הקובץ "${fileName}" נטען בהצלחה. נמצאו ${sheets.length} גיליונות.`);
  hideProgress();
}

function exportWorkbookWithFormulas() {
  const active = getActiveSheet();
  if (!active) {
    showError("צריך קודם לטעון קובץ.");
    return;
  }

  const prompt = userPrompt.value.trim() || "צור נוסחאות";
  const formulas = buildMultipleFormulas(prompt);

  const formulaRows = formulas.map(item => ({
    Formula_Name: item.title,
    Formula: item.formula,
    Explanation: item.explanation
  }));

  const wb = XLSX.utils.book_new();

  const originalSheet = XLSX.utils.json_to_sheet(active.preview);
  XLSX.utils.book_append_sheet(wb, originalSheet, "Preview");

  const formulasSheet = XLSX.utils.json_to_sheet(formulaRows);
  XLSX.utils.book_append_sheet(wb, formulasSheet, "Formulas");

  XLSX.writeFile(wb, "excel_hebrew_pro_output.xlsx");
  showStatus("נוצר קובץ חדש עם גיליון תצוגה מקדימה וגיליון נוסחאות.");
}

worker.onmessage = (event) => {
  const { type, payload } = event.data;

  if (type === "file-parsed") {
    handleParsedFile(payload.fileName, payload.sheets);
  }

  if (type === "parse-error") {
    hideProgress();
    showError(payload.message || "שגיאה בקריאת הקובץ.");
  }
};

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  hideError();
  hideStatus();
  showProgress("קורא את הקובץ ומנתח את המבנה שלו...");

  try {
    const buffer = await file.arrayBuffer();
    worker.postMessage({
      type: "parse-file",
      payload: {
        fileName: file.name,
        buffer
      }
    });
  } catch (error) {
    hideProgress();
    showError("לא ניתן לקרוא את הקובץ.");
  }
});

sheetSelect.addEventListener("change", () => {
  const index = Number(sheetSelect.value);
  if (Number.isNaN(index)) return;
  workbookState.activeSheetIndex = index;
  renderActiveSheet();
});

generateBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  hideError();
  clearSingleResult();
  clearMultiResult();

  if (!prompt) {
    showError("יש לכתוב בקשה לפני יצירת נוסחה.");
    return;
  }

  const result = buildSingleFormula(prompt);
  showSingleResult(result);
  saveHistory(prompt);
});

multiBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  hideError();
  clearSingleResult();
  clearMultiResult();

  if (!prompt) {
    showError("יש לכתוב בקשה לפני יצירת נוסחאות.");
    return;
  }

  const results = buildMultipleFormulas(prompt);
  showMultiResults(results);
  saveHistory(prompt);
});

exportBtn.addEventListener("click", exportWorkbookWithFormulas);

clearBtn.addEventListener("click", clearAll);

copyBtn.addEventListener("click", async () => {
  const text = formulaText.textContent.trim();
  if (!text) return;

  try {
    await navigator.clipboard.writeText(text);
    copyBtn.textContent = "הועתק";
    setTimeout(() => {
      copyBtn.textContent = "העתק";
    }, 1500);
  } catch {
    alert("לא ניתן להעתיק כרגע.");
  }
});

copyAllBtn.addEventListener("click", async () => {
  const allFormulas = Array.from(document.querySelectorAll(".formula-card pre"))
    .map(el => el.textContent.trim())
    .join("\n\n");

  if (!allFormulas) return;

  try {
    await navigator.clipboard.writeText(allFormulas);
    copyAllBtn.textContent = "הועתק";
    setTimeout(() => {
      copyAllBtn.textContent = "העתק הכל";
    }, 1500);
  } catch {
    alert("לא ניתן להעתיק כרגע.");
  }
});

clearHistoryBtn.addEventListener("click", () => {
  localStorage.removeItem("excelAssistantHistory");
  renderHistory();
});

quickTags.addEventListener("click", (event) => {
  const tag = event.target.closest(".tag");
  if (!tag) return;

  const value = tag.textContent.trim();
  const current = userPrompt.value.trim();
  userPrompt.value = current ? `${current} ${value}` : value;
  userPrompt.focus();
});

function renderExamples() {
  examplesList.innerHTML = "";
  EXAMPLES.forEach(item => {
    const div = document.createElement("div");
    div.className = "example-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    examplesList.appendChild(div);
  });
}

function renderHistory() {
  const history = JSON.parse(localStorage.getItem("excelAssistantHistory") || "[]");
  historyList.innerHTML = "";

  if (!history.length) {
    historyList.innerHTML = `<div class="empty-state">אין עדיין היסטוריה.</div>`;
    return;
  }

  history.forEach(item => {
    const div = document.createElement("div");
    div.className = "history-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    historyList.appendChild(div);
  });
}

renderExamples();
renderHistory();
