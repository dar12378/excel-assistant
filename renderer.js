const state = window.ExcelHebrewDesktopState;
const helpers = window.ExcelHebrewDesktopHelpers;

const openFileBtn = document.getElementById("openFileBtn");
const generateBtn = document.getElementById("generateBtn");
const multiBtn = document.getElementById("multiBtn");
const exportBtn = document.getElementById("exportBtn");
const clearHistoryBtn = document.getElementById("clearHistoryBtn");
const sheetSelect = document.getElementById("sheetSelect");

const userPrompt = document.getElementById("userPrompt");
const defaultColumnInput = document.getElementById("defaultColumn");

const statusBox = document.getElementById("statusBox");
const errorBox = document.getElementById("errorBox");

const fileSummarySection = document.getElementById("fileSummarySection");
const fileSummary = document.getElementById("fileSummary");
const columnsSection = document.getElementById("columnsSection");
const columnsList = document.getElementById("columnsList");
const previewSection = document.getElementById("previewSection");
const previewTable = document.getElementById("previewTable");

const singleResult = document.getElementById("singleResult");
const formulaText = document.getElementById("formulaText");
const explanationText = document.getElementById("explanationText");
const exampleText = document.getElementById("exampleText");
const tipsList = document.getElementById("tipsList");
const copyBtn = document.getElementById("copyBtn");

const multiResult = document.getElementById("multiResult");
const formulaGrid = document.getElementById("formulaGrid");
const copyAllBtn = document.getElementById("copyAllBtn");

const examplesList = document.getElementById("examplesList");
const historyList = document.getElementById("historyList");
const smartSuggestions = document.getElementById("smartSuggestions");

const EXAMPLES = [
  "חשב ממוצע של סכום",
  "ספור כמה פעמים Approved מופיע בסטטוס",
  "סכם עלות רק אם סטטוס הוא Approved",
  "מצא את הערך הגבוה ביותר בעמודת הכנסה",
  "חפש מחיר לפי קוד מוצר"
];

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

function getActiveSheet() {
  if (state.activeSheetIndex < 0) return null;
  return state.sheets[state.activeSheetIndex] || null;
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

function findColumnByPrompt(prompt) {
  const promptNorm = helpers.normalizeText(prompt);
  const columns = getColumns();

  for (const col of columns) {
    if (promptNorm.includes(helpers.normalizeText(col))) return col;
  }

  const aliases = [
    { test: /סטטוס|status|מצב/, keys: ["status", "סטטוס", "מצב"] },
    { test: /סכום|amount|total|sum/, keys: ["amount", "total", "sum", "סכום", "סהכ"] },
    { test: /מחיר|price/, keys: ["price", "מחיר"] },
    { test: /עלות|cost/, keys: ["cost", "עלות"] },
    { test: /הכנסה|revenue|income/, keys: ["revenue", "income", "הכנסה"] },
    { test: /פעיל|active/, keys: ["active", "פעיל"] },
    { test: /קוד|code|מזהה|id/, keys: ["code", "id", "מזהה", "קוד"] }
  ];

  for (const alias of aliases) {
    if (alias.test.test(promptNorm)) {
      const found = columns.find(col =>
        alias.keys.some(key => helpers.normalizeText(col).includes(key))
      );
      if (found) return found;
    }
  }

  return null;
}

function getDefaultColumn() {
  const value = String(defaultColumnInput.value || "").trim();
  return findColumnByPrompt(value) || value || "A";
}

function getBestNumericColumn() {
  const numeric = getColumnMeta().find(c => c.type === "number");
  return numeric?.name || getColumns()[0] || getDefaultColumn();
}

function getBestStatusColumn() {
  return getColumns().find(c => /status|state|סטטוס|מצב/i.test(c)) || getColumns()[1] || getDefaultColumn();
}

function getBestIdColumn() {
  return getColumns().find(c => /id|code|מזהה|קוד/i.test(c)) || getColumns()[0] || "Code";
}

function detectConditionValue(text) {
  if (helpers.containsAny(text, ["approved", "מאושר", "אושר"])) return "Approved";
  if (helpers.containsAny(text, ["yes", "כן"])) return "Yes";
  if (helpers.containsAny(text, ["no", "לא"])) return "No";
  return "Approved";
}

function getContext(prompt) {
  const column = findColumnByPrompt(prompt) || getDefaultColumn();
  const cell = helpers.detectCell(prompt) || "A2";
  return {
    text: helpers.normalizeText(prompt),
    column,
    cell
  };
}

function buildSingleFormula(prompt) {
  const ctx = getContext(prompt);
  const text = ctx.text;
  const conditionValue = detectConditionValue(text);

  const selectedCol = helpers.columnRef(ctx.column);
  const numericCol = helpers.columnRef(getBestNumericColumn());
  const statusCol = helpers.columnRef(getBestStatusColumn());
  const idCol = helpers.columnRef(getBestIdColumn());

  if (helpers.containsAny(text, ["countifs", "שני תנאים", "כמה תנאים", "וגם", "גם"])) {
    return helpers.makeResult(
      "COUNTIFS",
      helpers.excelFormula("COUNTIFS", `${statusCol},"Approved",${statusCol},"Approved"`),
      "הנוסחה סופרת שורות שמתאימות לתנאים מרובים.",
      "מתאים לדוחות גדולים של סטטוסים.",
      ["אפשר להתאים עמודות שונות לפי הצורך."]
    );
  }

  if (helpers.containsAny(text, ["sumif", "סכם רק אם", "סכום רק אם", "סכום בתנאי"])) {
    return helpers.makeResult(
      "SUMIF",
      helpers.excelFormula("SUMIF", `${statusCol},"Approved",${numericCol}`),
      "הנוסחה מסכמת רק שורות עם סטטוס Approved.",
      "שימושי לקבצים גדולים של חברות.",
      ["מומלץ לבחור עמודה מספרית לסכום."]
    );
  }

  if (helpers.containsAny(text, ["חבר", "סכום", "סכם", "sum"])) {
    return helpers.makeResult(
      "SUM",
      helpers.excelFormula("SUM", selectedCol),
      `הנוסחה מחברת את כל הערכים בעמודה ${ctx.column}.`,
      "שימושי לסכומים, הכנסות, מלאי ושעות.",
      ["תאים ריקים לא מפריעים בדרך כלל."]
    );
  }

  if (helpers.containsAny(text, ["ממוצע", "average"])) {
    return helpers.makeResult(
      "AVERAGE",
      helpers.excelFormula("AVERAGE", selectedCol),
      `הנוסחה מחשבת ממוצע של הערכים בעמודה ${ctx.column}.`,
      "מתאים ל-KPI, מחירים, ציונים ועלויות.",
      ["כדאי להשתמש בעמודה מספרית."]
    );
  }

  if (helpers.containsAny(text, ["מקסימום", "גבוה ביותר", "הכי גבוה", "max"])) {
    return helpers.makeResult(
      "MAX",
      helpers.excelFormula("MAX", selectedCol),
      `הנוסחה מחזירה את הערך הגבוה ביותר בעמודה ${ctx.column}.`,
      "מתאים לשיאים בדוחות ארוכים.",
      ["מומלץ לעמודה מספרית."]
    );
  }

  if (helpers.containsAny(text, ["מינימום", "נמוך ביותר", "הכי נמוך", "min"])) {
    return helpers.makeResult(
      "MIN",
      helpers.excelFormula("MIN", selectedCol),
      `הנוסחה מחזירה את הערך הנמוך ביותר בעמודה ${ctx.column}.`,
      "מתאים לערכים תחתונים בדוחות ארוכים.",
      ["מומלץ לעמודה מספרית."]
    );
  }

  if (helpers.containsAny(text, ["ספור", "ספירה", "כמה פעמים", "countif", "count"])) {
    return helpers.makeResult(
      "COUNTIF",
      helpers.excelFormula("COUNTIF", `${selectedCol},"${conditionValue}"`),
      `הנוסחה סופרת כמה פעמים "${conditionValue}" מופיע בעמודה ${ctx.column}.`,
      "מתאים מאוד לעמודות סטטוס בעברית.",
      ["אפשר להחליף את ערך החיפוש."]
    );
  }

  if (helpers.containsAny(text, ["אם", "בדוק אם", "גדול", "קטן", "if"])) {
    return helpers.makeResult(
      "IF",
      helpers.excelFormula("IF", `${ctx.cell}>100,"כן","לא"`),
      `הנוסחה בודקת אם הערך בתא ${ctx.cell} גדול מ-100.`,
      "מתאים לחריגים ובקרה.",
      ["אפשר לשנות את הסף 100."]
    );
  }

  if (helpers.containsAny(text, ["חפש", "חיפוש", "lookup", "xlookup", "vlookup", "קוד מוצר"])) {
    return helpers.makeResult(
      "XLOOKUP",
      helpers.excelFormula("XLOOKUP", `A2,${idCol},${numericCol},"לא נמצא"`),
      "הנוסחה מחפשת ערך לפי מזהה ומחזירה ערך מתאים.",
      "מתאים למחיר לפי קוד או נתון לפי מזהה.",
      ["אם אין XLOOKUP אפשר להחליף ל-VLOOKUP."]
    );
  }

  return helpers.makeResult(
    "ברירת מחדל",
    helpers.excelFormula("SUM", helpers.columnRef(getBestNumericColumn())),
    "לא זוהתה בקשה מדויקת, לכן הוחזרה נוסחת ברירת מחדל.",
    "כתבי שם עמודה מדויק כדי לקבל תוצאה טובה יותר.",
    ["מומלץ לכתוב את שם העמודה כפי שמופיע בקובץ."]
  );
}

function buildMultipleFormulas(prompt) {
  const ctx = getContext(prompt);
  const selectedCol = helpers.columnRef(ctx.column);
  const numericCol = helpers.columnRef(getBestNumericColumn());
  const statusCol = helpers.columnRef(getBestStatusColumn());
  const idCol = helpers.columnRef(getBestIdColumn());

  return [
    helpers.makeResult("SUM", helpers.excelFormula("SUM", numericCol), "סכום של עמודה מספרית.", "", []),
    helpers.makeResult("AVERAGE", helpers.excelFormula("AVERAGE", numericCol), "ממוצע של עמודה מספרית.", "", []),
    helpers.makeResult("MAX", helpers.excelFormula("MAX", numericCol), "מקסימום של עמודה מספרית.", "", []),
    helpers.makeResult("MIN", helpers.excelFormula("MIN", numericCol), "מינימום של עמודה מספרית.", "", []),
    helpers.makeResult("COUNTIF", helpers.excelFormula("COUNTIF", `${statusCol},"Approved"`), "ספירת Approved.", "", []),
    helpers.makeResult("IF", helpers.excelFormula("IF", `${ctx.cell}>100,"כן","לא"`), "בדיקת תנאי.", "", []),
    helpers.makeResult("SUMIF", helpers.excelFormula("SUMIF", `${statusCol},"Approved",${numericCol}`), "סכום לפי תנאי.", "", []),
    helpers.makeResult("XLOOKUP", helpers.excelFormula("XLOOKUP", `A2,${idCol},${selectedCol},"לא נמצא"`), "חיפוש לפי מזהה.", "", [])
  ];
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

function showSingleResult(data) {
  clearMultiResult();
  formulaText.textContent = data.formula;
  explanationText.textContent = data.explanation;
  exampleText.textContent = data.example;
  tipsList.innerHTML = "";

  data.tips.forEach((tip) => {
    const li = document.createElement("li");
    li.textContent = tip;
    tipsList.appendChild(li);
  });

  singleResult.classList.remove("hidden");
}

function showMultiResults(items) {
  clearSingleResult();
  formulaGrid.innerHTML = "";

  items.forEach((item) => {
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

function renderFileSummary() {
  const active = getActiveSheet();
  if (!active) {
    fileSummarySection.classList.add("hidden");
    return;
  }

  fileSummary.innerHTML = `
    <div class="summary-card"><strong>קובץ</strong><span>${state.fileName}</span></div>
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

  columns.forEach((col) => {
    const info = meta.find((m) => m.name === col);
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "column-chip";
    chip.textContent = info ? `${col} (${info.type})` : col;
    chip.addEventListener("click", () => {
      defaultColumnInput.value = col;
      userPrompt.value = `חשב ממוצע של ${col}`;
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

  columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    trHead.appendChild(th);
  });

  thead.appendChild(trHead);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  rows.slice(0, 10).forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((col) => {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewTable.appendChild(tbody);
  previewSection.classList.remove("hidden");
}

function renderSuggestions() {
  smartSuggestions.innerHTML = "";
  const columns = getColumns();
  const meta = getColumnMeta();

  if (!columns.length) {
    smartSuggestions.innerHTML = `<div class="empty-state">פתחי קובץ כדי לקבל הצעות.</div>`;
    return;
  }

  const numeric = meta.filter((m) => m.type === "number").map((m) => m.name);
  const statusCols = columns.filter((c) => /status|state|סטטוס|מצב/i.test(c));
  const suggestions = [];

  numeric.forEach((col) => {
    suggestions.push(`חשב ממוצע של ${col}`);
    suggestions.push(`מצא את הערך הגבוה ביותר בעמודת ${col}`);
    suggestions.push(`חבר את כל הערכים בעמודת ${col}`);
  });

  statusCols.forEach((col) => {
    suggestions.push(`ספור כמה פעמים Approved מופיע ב${col}`);
  });

  const unique = [...new Set(suggestions)].slice(0, 10);

  unique.forEach((item) => {
    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    smartSuggestions.appendChild(div);
  });
}

function saveHistory(prompt) {
  const history = JSON.parse(localStorage.getItem("excelDesktopHistory") || "[]");
  const updated = [prompt, ...history.filter((item) => item !== prompt)].slice(0, 8);
  localStorage.setItem("excelDesktopHistory", JSON.stringify(updated));
  renderHistory();
}

function renderHistory() {
  const history = JSON.parse(localStorage.getItem("excelDesktopHistory") || "[]");
  historyList.innerHTML = "";

  if (!history.length) {
    historyList.innerHTML = `<div class="empty-state">אין עדיין היסטוריה.</div>`;
    return;
  }

  history.forEach((item) => {
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
  EXAMPLES.forEach((item) => {
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

function rebuildSheetSelect() {
  sheetSelect.innerHTML = "";

  if (!state.sheets.length) {
    sheetSelect.innerHTML = `<option value="">אין גיליון</option>`;
    sheetSelect.disabled = true;
    return;
  }

  state.sheets.forEach((sheet, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = sheet.name;
    if (index === state.activeSheetIndex) option.selected = true;
    sheetSelect.appendChild(option);
  });

  sheetSelect.disabled = false;
}

function renderActiveSheet() {
  const firstColumn = getColumns()[0];
  if (firstColumn) defaultColumnInput.value = firstColumn;
  renderFileSummary();
  renderColumns();
  renderPreview();
  renderSuggestions();
}

openFileBtn.addEventListener("click", async () => {
  hideError();
  hideStatus();

  const pick = await window.excelDesktopAPI.pickFile();
  if (pick.canceled) return;

  showStatus("טוען את הקובץ...");
  const result = await window.excelDesktopAPI.readWorkbook(pick.filePath);

  if (!result.ok) {
    showError(result.error || "שגיאה בקריאת הקובץ.");
    return;
  }

  state.fileName = result.data.fileName;
  state.sheets = result.data.sheets;
  state.activeSheetIndex = result.data.sheets.length ? 0 : -1;

  rebuildSheetSelect();
  renderActiveSheet();
  showStatus(`הקובץ "${state.fileName}" נטען בהצלחה.`);
});

sheetSelect.addEventListener("change", () => {
  const index = Number(sheetSelect.value);
  if (Number.isNaN(index)) return;
  state.activeSheetIndex = index;
  renderActiveSheet();
});

generateBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();
  hideError();

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

  if (!prompt) {
    showError("יש לכתוב בקשה לפני יצירת נוסחאות.");
    return;
  }

  const results = buildMultipleFormulas(prompt);
  showMultiResults(results);
  saveHistory(prompt);
});

exportBtn.addEventListener("click", async () => {
  const active = getActiveSheet();
  if (!active) {
    showError("צריך קודם לפתוח קובץ.");
    return;
  }

  const formulas = buildMultipleFormulas(userPrompt.value.trim() || "צור נוסחאות");
  const result = await window.excelDesktopAPI.saveFormulasWorkbook({
    sourcePreview: active.preview,
    formulas,
    defaultName: "excel_hebrew_desktop_output.xlsx"
  });

  if (!result.ok && !result.canceled) {
    showError(result.error || "שגיאה בשמירת הקובץ.");
    return;
  }

  if (result.ok) {
    showStatus(`הקובץ נשמר בהצלחה: ${result.savedTo}`);
  }
});

copyBtn.addEventListener("click", async () => {
  const text = formulaText.textContent.trim();
  if (!text) return;
  await navigator.clipboard.writeText(text);
  copyBtn.textContent = "הועתק";
  setTimeout(() => {
    copyBtn.textContent = "העתק";
  }, 1200);
});

copyAllBtn.addEventListener("click", async () => {
  const formulas = Array.from(document.querySelectorAll(".formula-card pre"))
    .map((el) => el.textContent.trim())
    .join("\n\n");

  if (!formulas) return;

  await navigator.clipboard.writeText(formulas);
  copyAllBtn.textContent = "הועתק";
  setTimeout(() => {
    copyAllBtn.textContent = "העתק הכל";
  }, 1200);
});

clearHistoryBtn.addEventListener("click", () => {
  localStorage.removeItem("excelDesktopHistory");
  renderHistory();
});

renderExamples();
renderHistory();
