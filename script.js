const userPrompt = document.getElementById("userPrompt");
const generateBtn = document.getElementById("generateBtn");
const multiBtn = document.getElementById("multiBtn");
const clearBtn = document.getElementById("clearBtn");
const clearHistoryBtn = document.getElementById("clearHistoryBtn");
const fileInput = document.getElementById("fileInput");

const errorBox = document.getElementById("errorBox");
const fileStatus = document.getElementById("fileStatus");

const singleResult = document.getElementById("singleResult");
const multiResult = document.getElementById("multiResult");
const columnsSection = document.getElementById("columnsSection");
const previewSection = document.getElementById("previewSection");

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

let uploadedData = [];
let uploadedColumns = [];
let uploadedFileName = "";

const EXAMPLES = [
  "חבר את כל הערכים בעמודה Amount",
  "חשב ממוצע של עמודה Score",
  "ספור כמה פעמים Approved מופיע בעמודה Status",
  "בדוק אם A2 גדול מ-100",
  "מצא את הערך הגבוה ביותר בעמודה Revenue",
  "מצא את הערך הנמוך ביותר בעמודה Cost",
  "חפש מחיר לפי קוד מוצר",
  "סכם את עמודה Amount רק אם בעמודה Status כתוב Approved",
  "ספור כמה שורות יש שבהן Status הוא Approved וגם Active הוא Yes"
];

function showError(message) {
  errorBox.textContent = message;
  errorBox.classList.remove("hidden");
}

function hideError() {
  errorBox.classList.add("hidden");
  errorBox.textContent = "";
}

function showFileStatus(message) {
  fileStatus.textContent = message;
  fileStatus.classList.remove("hidden");
}

function hideFileStatus() {
  fileStatus.classList.add("hidden");
  fileStatus.textContent = "";
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

function clearPreview() {
  columnsSection.classList.add("hidden");
  previewSection.classList.add("hidden");
  columnsList.innerHTML = "";
  previewTable.innerHTML = "";
  smartSuggestions.innerHTML = `<div class="empty-state">העלי קובץ כדי לקבל הצעות חכמות לפי העמודות.</div>`;
}

function clearAll() {
  userPrompt.value = "";
  hideError();
  hideFileStatus();
  clearSingleResult();
  clearMultiResult();
}

function containsAny(text, words) {
  return words.some(word => text.includes(word));
}

function normalizeColumnName(value) {
  if (!value) return "A";
  return String(value).trim();
}

function normalizeExcelLetter(value) {
  if (!value) return "A";
  const clean = String(value).trim().toUpperCase().replace(/[^A-Z]/g, "");
  return clean || "A";
}

function detectColumnByPhrase(text) {
  const match = text.match(/עמודה\s*([A-Z]{1,3}|[A-Za-z_][A-Za-z0-9 _-]*)/i);
  return match ? normalizeColumnName(match[1]) : null;
}

function detectCell(text) {
  const match = text.match(/([A-Z]+\d+)/i);
  return match ? match[1].toUpperCase() : null;
}

function detectRange(text) {
  const match = text.match(/([A-Z]+\d+)\s*[:-]\s*([A-Z]+\d+)/i);
  if (!match) return null;
  return `${match[1].toUpperCase()}:${match[2].toUpperCase()}`;
}

function excelFormula(name, args) {
  return `=${name.toUpperCase()}(${args})`;
}

function findUploadedColumnByName(prompt) {
  const lower = prompt.toLowerCase();
  for (const col of uploadedColumns) {
    if (lower.includes(String(col).toLowerCase())) {
      return col;
    }
  }
  return null;
}

function getContext(prompt) {
  const range = detectRange(prompt);
  const promptColumn = detectColumnByPhrase(prompt);
  const uploadedColumn = findUploadedColumnByName(prompt);
  const fallback = normalizeColumnName(defaultColumnInput.value) || "A";
  const column = uploadedColumn || promptColumn || fallback;
  const cell = detectCell(prompt) || "A2";

  return {
    text: prompt.toLowerCase(),
    column,
    cell,
    range,
    area: range || column
  };
}

function makeResult(title, formula, explanation, example, tips) {
  return { title, formula, explanation, example, tips };
}

function detectConditionValue(text) {
  if (containsAny(text, ["approved", "אושר", "מאושר"])) return "Approved";
  if (containsAny(text, ["yes", "כן"])) return "Yes";
  if (containsAny(text, ["no", "לא"])) return "No";
  return "Yes";
}

function columnRef(columnName) {
  if (/^[A-Z]{1,3}$/.test(String(columnName).trim().toUpperCase())) {
    return `${String(columnName).trim().toUpperCase()}:${String(columnName).trim().toUpperCase()}`;
  }
  return `[${columnName}]`;
}

function singleCellRef(cell) {
  return cell;
}

function buildSingleFormula(prompt) {
  const ctx = getContext(prompt);
  const text = ctx.text;
  const conditionValue = detectConditionValue(text);
  const col = columnRef(ctx.column);

  if (containsAny(text, ["countifs", "שני תנאים", "כמה תנאים", "וגם", "גם"])) {
    const first = uploadedColumns[0] || "Status";
    const second = uploadedColumns[1] || "Active";
    return makeResult(
      "COUNTIFS",
      excelFormula("COUNTIFS", `${columnRef(first)},"Approved",${columnRef(second)},"Yes"`),
      "הנוסחה סופרת שורות שמתאימות לשני תנאים במקביל.",
      "מתאים לדוחות סטטוסים, אישורים ומעקבים.",
      [
        "אם טעון קובץ, אפשר להחליף את שמות העמודות לפי הכותרות האמיתיות.",
        "אם אין קובץ, אפשר להשתמש בעמודות רגילות כמו C:C ו-D:D."
      ]
    );
  }

  if (containsAny(text, ["sumif", "סכם רק אם", "סכום רק אם", "סכום בתנאי"])) {
    const sumColumn = uploadedColumns.find(c => /amount|sum|price|cost|revenue|total/i.test(String(c))) || uploadedColumns[0] || ctx.column;
    const conditionColumn = uploadedColumns.find(c => /status|state|approved/i.test(String(c))) || uploadedColumns[1] || "Status";
    return makeResult(
      "SUMIF",
      excelFormula("SUMIF", `${columnRef(conditionColumn)},"Approved",${columnRef(sumColumn)}`),
      "הנוסחה מסכמת ערכים רק בשורות שבהן מתקיים תנאי מסוים.",
      "מעולה לסיכום סכומים רק עבור רשומות מאושרות או לפי סטטוס.",
      [
        "החלק הראשון הוא עמודת התנאי.",
        "החלק האחרון הוא עמודת הסכומים."
      ]
    );
  }

  if (containsAny(text, ["חבר", "סכום", "סכם", "סכימה", "sum"])) {
    return makeResult(
      "SUM",
      excelFormula("SUM", col),
      `הנוסחה מחברת את כל הערכים בעמודה או בטווח ${ctx.column}.`,
      "שימושי לסכומים, מכירות, שעות, תקציבים ועוד.",
      [
        "אם הקובץ מכיל עמודת Amount או Revenue, כדאי לבחור בה.",
        "אפשר גם לעבוד על טווח כמו A2:A50."
      ]
    );
  }

  if (containsAny(text, ["ממוצע", "average"])) {
    return makeResult(
      "AVERAGE",
      excelFormula("AVERAGE", col),
      `הנוסחה מחשבת ממוצע של הערכים בעמודה או בטווח ${ctx.column}.`,
      "מתאים לציונים, מחירים, עלויות וביצועים.",
      [
        "תאים ריקים בדרך כלל לא נכללים בממוצע.",
        "כדאי לבחור עמודה מספרית."
      ]
    );
  }

  if (containsAny(text, ["הכי גבוה", "גבוה ביותר", "מקסימום", "max"])) {
    return makeResult(
      "MAX",
      excelFormula("MAX", col),
      `הנוסחה מחזירה את הערך הגבוה ביותר בעמודה או בטווח ${ctx.column}.`,
      "מתאים למציאת שיא של מחיר, ציון או הכנסה.",
      [
        "השתמשי בעמודה מספרית בלבד.",
        "לערך הנמוך ביותר השתמשי ב-MIN."
      ]
    );
  }

  if (containsAny(text, ["הכי נמוך", "נמוך ביותר", "מינימום", "min"])) {
    return makeResult(
      "MIN",
      excelFormula("MIN", col),
      `הנוסחה מחזירה את הערך הנמוך ביותר בעמודה או בטווח ${ctx.column}.`,
      "מתאים למציאת מינימום של מחיר, ציון או עלות.",
      [
        "השתמשי בעמודה מספרית בלבד.",
        "לערך הגבוה ביותר השתמשי ב-MAX."
      ]
    );
  }

  if (containsAny(text, ["ספור", "ספירה", "כמה פעמים", "countif", "count"])) {
    return makeResult(
      "COUNTIF",
      excelFormula("COUNTIF", `${col},"${conditionValue}"`),
      `הנוסחה סופרת כמה פעמים הערך "${conditionValue}" מופיע בעמודה ${ctx.column}.`,
      "מתאים לספירת סטטוסים, תשובות, ערכים חוזרים ועוד.",
      [
        "אפשר להחליף את ערך החיפוש לכל טקסט אחר.",
        'אפשר גם להשתמש בתנאים כמו ">100".'
      ]
    );
  }

  if (containsAny(text, ["בדוק אם", "אם", "גדול", "קטן", "שווה", "if"])) {
    return makeResult(
      "IF",
      excelFormula("IF", `${singleCellRef(ctx.cell)}>100,"כן","לא"`),
      `הנוסחה בודקת אם הערך בתא ${ctx.cell} גדול מ-100.`,
      "מתאים לבדיקות סטטוס, חריגות, ספים ותנאים עסקיים.",
      [
        "אפשר לשנות את 100 לכל מספר אחר.",
        "אפשר לשנות את התוצאה כן/לא לכל טקסט אחר."
      ]
    );
  }

  if (containsAny(text, ["חפש", "חיפוש", "קוד מוצר", "מצא מחיר", "lookup", "xlookup", "vlookup"])) {
    const first = uploadedColumns[0] || "Code";
    const second = uploadedColumns[1] || "Price";
    return makeResult(
      "XLOOKUP",
      excelFormula("XLOOKUP", `A2,${columnRef(first)},${columnRef(second)},"לא נמצא"`),
      "הנוסחה מחפשת ערך לפי מפתח ומחזירה ערך מתאים מעמודה אחרת.",
      "מתאים לחיפוש מחיר, שם מוצר, לקוח, סטטוס או כל מידע לפי מזהה.",
      [
        "אם נטען קובץ, כדאי לבחור עמודת מזהה ועמודת תוצאה מהקובץ.",
        "אם אין אצלך XLOOKUP, אפשר להחליף ל-VLOOKUP."
      ]
    );
  }

  return makeResult(
    "ברירת מחדל",
    excelFormula("SUM", col),
    `לא זוהתה בקשה מדויקת, לכן הוחזרה נוסחת ברירת מחדל לעמודה ${ctx.column}.`,
    "נסי לכתוב בקשה ברורה יותר, למשל: חשב ממוצע של עמודה Amount.",
    [
      "כדאי לציין שם עמודה שמופיע בקובץ.",
      "כדאי לציין גם תנאי אם יש צורך."
    ]
  );
}

function buildMultipleFormulas(prompt) {
  const ctx = getContext(prompt);
  const col = columnRef(ctx.column);
  const statusCol = uploadedColumns.find(c => /status|state/i.test(String(c))) || "Status";
  const valueCol = uploadedColumns.find(c => /amount|price|revenue|cost|score|total/i.test(String(c))) || ctx.column;

  return [
    makeResult("SUM", excelFormula("SUM", col), `סכום של ${ctx.column}.`, "", []),
    makeResult("AVERAGE", excelFormula("AVERAGE", col), `ממוצע של ${ctx.column}.`, "", []),
    makeResult("MAX", excelFormula("MAX", col), `מקסימום בעמודה ${ctx.column}.`, "", []),
    makeResult("MIN", excelFormula("MIN", col), `מינימום בעמודה ${ctx.column}.`, "", []),
    makeResult("COUNTIF", excelFormula("COUNTIF", `${col},"Yes"`), `ספירת Yes בעמודה ${ctx.column}.`, "", []),
    makeResult("IF", excelFormula("IF", `${ctx.cell}>100,"כן","לא"`), `בדיקה אם ${ctx.cell} גדול מ-100.`, "", []),
    makeResult("SUMIF", excelFormula("SUMIF", `${columnRef(statusCol)},"Approved",${columnRef(valueCol)}`), "סכום לפי תנאי.", "", []),
    makeResult("COUNTIFS", excelFormula("COUNTIFS", `${columnRef(statusCol)},"Approved",${columnRef(statusCol)},"Approved"`), "ספירה לפי תנאים.", "", []),
    makeResult("XLOOKUP", excelFormula("XLOOKUP", `A2,${columnRef(uploadedColumns[0] || "Code")},${columnRef(uploadedColumns[1] || "Value")},"לא נמצא"`), "חיפוש ערך לפי מפתח.", "", [])
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

function renderColumns() {
  columnsList.innerHTML = "";
  if (!uploadedColumns.length) {
    columnsSection.classList.add("hidden");
    return;
  }

  uploadedColumns.forEach(col => {
    const chip = document.createElement("div");
    chip.className = "column-chip";
    chip.textContent = col;
    chip.addEventListener("click", () => {
      userPrompt.value = `חשב ממוצע של עמודה ${col}`;
      userPrompt.focus();
    });
    columnsList.appendChild(chip);
  });

  columnsSection.classList.remove("hidden");
}

function renderPreview() {
  previewTable.innerHTML = "";

  if (!uploadedData.length) {
    previewSection.classList.add("hidden");
    return;
  }

  const headers = uploadedColumns;
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  headers.forEach(header => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  uploadedData.slice(0, 5).forEach(row => {
    const tr = document.createElement("tr");
    headers.forEach(header => {
      const td = document.createElement("td");
      td.textContent = row[header] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewTable.appendChild(tbody);
  previewSection.classList.remove("hidden");
}

function renderSmartSuggestions() {
  smartSuggestions.innerHTML = "";

  if (!uploadedColumns.length) {
    smartSuggestions.innerHTML = `<div class="empty-state">העלי קובץ כדי לקבל הצעות חכמות לפי העמודות.</div>`;
    return;
  }

  const suggestions = [];

  const numericLike = uploadedColumns.filter(col =>
    /amount|price|cost|revenue|total|score|qty|quantity|hours|sum|value/i.test(String(col))
  );

  const statusLike = uploadedColumns.filter(col =>
    /status|state|approved|active|flag|result/i.test(String(col))
  );

  numericLike.forEach(col => {
    suggestions.push(`חשב ממוצע של עמודה ${col}`);
    suggestions.push(`מצא את הערך הגבוה ביותר בעמודה ${col}`);
    suggestions.push(`חבר את כל הערכים בעמודה ${col}`);
  });

  statusLike.forEach(col => {
    suggestions.push(`ספור כמה פעמים Approved מופיע בעמודה ${col}`);
  });

  if (uploadedColumns.length >= 2) {
    suggestions.push(`חפש ערך לפי ${uploadedColumns[0]}`);
    suggestions.push(`סכם את עמודה ${uploadedColumns[1]} רק אם בעמודה ${uploadedColumns[0]} כתוב Approved`);
  }

  const unique = [...new Set(suggestions)].slice(0, 10);

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
    smartSuggestions.innerHTML = `<div class="empty-state">לא נמצאו עדיין הצעות אוטומטיות.</div>`;
  }
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
    return values.map(v => v.trim());
  });

  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header || `Column${index + 1}`] = row[index] ?? "";
    });
    return obj;
  });
}

function loadFromObjects(data, fileName) {
  if (!Array.isArray(data) || !data.length) {
    showError("הקובץ נטען אבל לא נמצאו בו נתונים.");
    return;
  }

  uploadedData = data.filter(row => row && typeof row === "object");
  uploadedColumns = Object.keys(uploadedData[0] || {});
  uploadedFileName = fileName || "";

  showFileStatus(`הקובץ "${uploadedFileName}" נטען בהצלחה. זוהו ${uploadedColumns.length} עמודות ו-${uploadedData.length} שורות.`);
  renderColumns();
  renderPreview();
  renderSmartSuggestions();

  if (uploadedColumns.length) {
    defaultColumnInput.value = uploadedColumns[0];
  }
}

function handleFileUpload(file) {
  if (!file) return;

  hideError();
  hideFileStatus();

  const name = file.name.toLowerCase();

  if (name.endsWith(".csv")) {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const text = event.target.result;
        const data = parseCSV(text);
        loadFromObjects(data, file.name);
      } catch (error) {
        showError("לא ניתן לקרוא את קובץ ה-CSV.");
      }
    };
    reader.readAsText(file, "utf-8");
    return;
  }

  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        loadFromObjects(json, file.name);
      } catch (error) {
        showError("לא ניתן לקרוא את קובץ האקסל.");
      }
    };
    reader.readAsArrayBuffer(file);
    return;
  }

  showError("יש להעלות קובץ CSV, XLSX או XLS בלבד.");
}

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
  } catch (error) {
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
  } catch (error) {
    alert("לא ניתן להעתיק כרגע.");
  }
});

clearHistoryBtn.addEventListener("click", () => {
  localStorage.removeItem("excelAssistantHistory");
  renderHistory();
});

quickTags.addEventListener("click", event => {
  const tag = event.target.closest(".tag");
  if (!tag) return;

  const value = tag.textContent.trim();
  const current = userPrompt.value.trim();
  userPrompt.value = current ? `${current} ${value}` : value;
  userPrompt.focus();
});

fileInput.addEventListener("change", event => {
  const file = event.target.files[0];
  handleFileUpload(file);
});

renderExamples();
renderHistory();
renderSmartSuggestions();
