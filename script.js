const userPrompt = document.getElementById("userPrompt");
const defaultColumnInput = document.getElementById("defaultColumn");
const generateBtn = document.getElementById("generateBtn");
const analyzeBtn = document.getElementById("analyzeBtn");
const reviewFileBtn = document.getElementById("reviewFileBtn");
const multiBtn = document.getElementById("multiBtn");
const fixFormulaBtn = document.getElementById("fixFormulaBtn");
const clearBtn = document.getElementById("clearBtn");
const uploadBtn = document.getElementById("uploadBtn");
const voiceBtn = document.getElementById("voiceBtn");
const fileInput = document.getElementById("fileInput");
const clearHistoryBtn = document.getElementById("clearHistoryBtn");
const languageSelect = document.getElementById("languageSelect");

const statusBox = document.getElementById("statusBox");
const formulaText = document.getElementById("formulaText");
const explanationText = document.getElementById("explanationText");
const exampleText = document.getElementById("exampleText");
const tipsList = document.getElementById("tipsList");

const assistantAnswer = document.getElementById("assistantAnswer");
const copyAnswerBtn = document.getElementById("copyAnswerBtn");

const singleResult = document.getElementById("singleResult");
const multiResult = document.getElementById("multiResult");
const formulaGrid = document.getElementById("formulaGrid");

const copyBtn = document.getElementById("copyBtn");
const copyAllBtn = document.getElementById("copyAllBtn");

const columnsSection = document.getElementById("columnsSection");
const columnsList = document.getElementById("columnsList");
const previewSection = document.getElementById("previewSection");
const previewTable = document.getElementById("previewTable");

const fileInsightSection = document.getElementById("fileInsightSection");
const fileInsightText = document.getElementById("fileInsightText");

const fileIssuesSection = document.getElementById("fileIssuesSection");
const fileIssuesText = document.getElementById("fileIssuesText");

const suggestionsList = document.getElementById("suggestionsList");
const historyList = document.getElementById("historyList");

let excelColumns = [];
let excelRows = [];
let uploadedFileName = "";
let currentLang = "he";

const UI_TEXT = {
  he: {
    title: "Helper Excel Reading",
    subtitle: "כתבי או אמרי מה את רוצה לעשות באקסל, העלי קובץ, וקבלי עזרה חכמה בעברית או באנגלית.",
    language: "שפה",
    promptLabel: "מה תרצי לעשות באקסל?",
    promptPlaceholder: "לדוגמה: תקן לי את הנוסחה, סכם את הקובץ, מצא שגיאות, או חשב ממוצע של סכום",
    defaultColumn: "עמודת ברירת מחדל",
    generate: "צור נוסחה",
    analyze: "נתח בקשה",
    reviewFile: "בדוק את הקובץ",
    multi: "כמה נוסחאות",
    fixFormula: "תקן נוסחה",
    upload: "העלאת קובץ",
    voice: "דיבור",
    clear: "נקה",
    clearHistory: "נקה היסטוריה",
    assistantAnswer: "תשובת האפליקציה",
    formula: "נוסחה",
    explanation: "הסבר",
    example: "דוגמה",
    tips: "טיפים",
    multiTitle: "נוסחאות אפשריות",
    fileInsight: "הבנת הקובץ",
    fileIssues: "שגיאות ותיקונים מומלצים",
    columns: "עמודות שזוהו בקובץ",
    preview: "תצוגה מקדימה של הקובץ",
    suggestions: "הצעות מהירות",
    history: "היסטוריה",
    copy: "העתק",
    copyAll: "העתק הכל",
    copied: "הועתק",
    noHistory: "אין עדיין היסטוריה.",
    noFileYet: "לא נטען עדיין קובץ.",
    assistantDefault: "כאן יופיעו הסברים, סיכומים, ותשובות של האפליקציה."
  },
  en: {
    title: "Helper Excel Reading",
    subtitle: "Write or say what you want to do in Excel, upload a file, and get smart help in Hebrew or English.",
    language: "Language",
    promptLabel: "What would you like to do in Excel?",
    promptPlaceholder: "For example: fix my formula, summarize the file, find issues, or calculate an average of amount",
    defaultColumn: "Default column",
    generate: "Create formula",
    analyze: "Analyze request",
    reviewFile: "Review file",
    multi: "Multiple formulas",
    fixFormula: "Fix formula",
    upload: "Upload file",
    voice: "Voice",
    clear: "Clear",
    clearHistory: "Clear history",
    assistantAnswer: "App answer",
    formula: "Formula",
    explanation: "Explanation",
    example: "Example",
    tips: "Tips",
    multiTitle: "Possible formulas",
    fileInsight: "File understanding",
    fileIssues: "Issues and suggested fixes",
    columns: "Detected columns",
    preview: "File preview",
    suggestions: "Quick suggestions",
    history: "History",
    copy: "Copy",
    copyAll: "Copy all",
    copied: "Copied",
    noHistory: "No history yet.",
    noFileYet: "No file uploaded yet.",
    assistantDefault: "Explanations, summaries, and answers from the app will appear here."
  }
};

const DEFAULT_SUGGESTIONS = {
  he: [
    "סכם את הקובץ",
    "מצא שגיאות בקובץ",
    "תקן לי את הנוסחה =sum(a:a",
    "חשב ממוצע של סכום",
    "ספור כמה פעמים Approved מופיע בסטטוס",
    "עזור לי להבין למה הקובץ משמש",
    "השלם לי נוסחה לטבלת מעקב",
    "בדוק אם יש עמודות חסרות"
  ],
  en: [
    "Summarize the file",
    "Find issues in the file",
    "Fix this formula =sum(a:a",
    "Calculate average of amount",
    "Count how many Approved values appear in status",
    "Help me understand what the file is for",
    "Complete a formula for a tracking table",
    "Check whether any columns are missing"
  ]
};

const COMMON_CORRECTIONS = {
  he: {
    "עיברית": "עברית",
    "אקסאל": "אקסל",
    "אקסאלים": "אקסלים",
    "נסחא": "נוסחה",
    "נסחאות": "נוסחאות",
    "ממוצה": "ממוצע",
    "העלהת": "העלאת",
    "תאויות": "טעויות",
    "שגיעות": "שגיאות",
    "תיכון": "תיקון"
  },
  en: {
    "formla": "formula",
    "colum": "column",
    "avrage": "average",
    "uplod": "upload",
    "fomula": "formula",
    "spredsheet": "spreadsheet",
    "erors": "errors"
  }
};

function t(key) {
  return UI_TEXT[currentLang][key];
}

function normalizeText(text) {
  return String(text || "").toLowerCase().trim();
}

function containsAny(text, words) {
  return words.some(word => text.includes(word));
}

function showStatus(message, type = "info") {
  statusBox.textContent = message;
  statusBox.className = "status-line " + (type === "error" ? "status-error" : "status-info");
}

function clearStatus() {
  statusBox.textContent = "";
  statusBox.className = "status-line";
}

function applySpellingCorrections(text) {
  let fixed = String(text || "");
  const map = COMMON_CORRECTIONS[currentLang];
  Object.keys(map).forEach((wrong) => {
    const regex = new RegExp(wrong, "gi");
    fixed = fixed.replace(regex, map[wrong]);
  });
  return fixed;
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

function detectColumnByPhrase(text) {
  const match = text.match(/(?:עמודה|column)\s*([A-Z]{1,3}|[A-Za-zא-ת_][A-Za-zא-ת0-9 _-]*)/i);
  return match ? match[1].trim() : null;
}

function normalizeColumn(value) {
  if (!value) return "A";
  return String(value).trim();
}

function findColumnFromUploadedData(prompt) {
  const lower = normalizeText(prompt);
  for (const col of excelColumns) {
    if (lower.includes(normalizeText(col))) {
      return col;
    }
  }
  return null;
}

function excelFormula(name, args) {
  return `=${name.toUpperCase()}(${args})`;
}

function makeResult(title, formula, explanation, example, tips) {
  return { title, formula, explanation, example, tips };
}

function detectConditionValue(text) {
  if (containsAny(text, ["approved", "אושר", "מאושר"])) return "Approved";
  if (containsAny(text, ["yes", "כן"])) return "Yes";
  if (containsAny(text, ["no", "לא"])) return "No";
  if (containsAny(text, ["active", "פעיל"])) return "Active";
  return "Yes";
}

function getContext(prompt) {
  const fixedPrompt = applySpellingCorrections(prompt);
  const range = detectRange(fixedPrompt);
  const uploadedColumn = findColumnFromUploadedData(fixedPrompt);
  const phraseColumn = detectColumnByPhrase(fixedPrompt);
  const fallback = normalizeColumn(defaultColumnInput.value) || "A";
  const column = uploadedColumn || phraseColumn || fallback;
  const cell = detectCell(fixedPrompt) || "A2";
  const area = range || `${column}:${column}`;

  return {
    text: normalizeText(fixedPrompt),
    fixedPrompt,
    column,
    cell,
    area
  };
}

function getNumericColumns() {
  if (!excelRows.length || !excelColumns.length) return [];
  return excelColumns.filter((col) => {
    const sample = excelRows.slice(0, 80).map(row => row[col]).filter(v => v !== "" && v != null);
    if (!sample.length) return false;
    const numericCount = sample.filter(v => !isNaN(Number(String(v).replace(/,/g, "")))).length;
    return numericCount / sample.length >= 0.6;
  });
}

function getDateColumns() {
  if (!excelRows.length || !excelColumns.length) return [];
  return excelColumns.filter((col) => {
    const sample = excelRows.slice(0, 50).map(row => row[col]).filter(v => v !== "" && v != null);
    if (!sample.length) return false;
    const dateCount = sample.filter(v => !isNaN(Date.parse(v))).length;
    return dateCount / sample.length >= 0.6;
  });
}

function guessFilePurpose() {
  if (!excelColumns.length) {
    return currentLang === "he"
      ? "עדיין לא נטען קובץ, לכן אין מספיק מידע להבין את מטרת הקובץ."
      : "No file has been uploaded yet, so there is not enough information to understand the file purpose.";
  }

  const lowerCols = excelColumns.map(c => normalizeText(c));

  if (
    lowerCols.some(c => c.includes("status") || c.includes("סטטוס")) &&
    lowerCols.some(c => c.includes("amount") || c.includes("סכום") || c.includes("price") || c.includes("מחיר"))
  ) {
    return currentLang === "he"
      ? "זה נראה כמו קובץ מעקב עסקי או דוח ביצועים עם סטטוסים וסכומים."
      : "This looks like a business tracking or performance report file with statuses and amounts.";
  }

  if (
    lowerCols.some(c => c.includes("date") || c.includes("תאריך")) &&
    lowerCols.some(c => c.includes("name") || c.includes("שם"))
  ) {
    return currentLang === "he"
      ? "זה נראה כמו קובץ רשומות, משימות או פעילות לפי תאריכים."
      : "This looks like a records, tasks, or activity file organized by dates.";
  }

  if (
    lowerCols.some(c => c.includes("inventory") || c.includes("מלאי") || c.includes("qty") || c.includes("quantity"))
  ) {
    return currentLang === "he"
      ? "זה נראה כמו קובץ מלאי או ניהול כמויות."
      : "This looks like an inventory or quantity management file.";
  }

  return currentLang === "he"
    ? "זה נראה כמו קובץ נתונים כללי. האפליקציה יכולה לעזור בנוסחאות, בדיקת שגיאות, וסיכום מבנה הקובץ."
    : "This looks like a general data file. The app can help with formulas, error checks, and file structure summaries.";
}

function summarizeFile() {
  if (!excelRows.length || !excelColumns.length) {
    return currentLang === "he"
      ? "לא נטען עדיין קובץ. אחרי העלאת קובץ אוכל לסכם את המבנה שלו, לזהות שגיאות, ולהציע תיקונים."
      : "No file has been uploaded yet. After uploading a file, I can summarize its structure, detect issues, and suggest fixes.";
  }

  const numericColumns = getNumericColumns();
  const dateColumns = getDateColumns();

  if (currentLang === "he") {
    return [
      `שם הקובץ: ${uploadedFileName}`,
      `שורות שזוהו: ${excelRows.length}`,
      `עמודות שזוהו: ${excelColumns.join(", ")}`,
      `מטרת הקובץ המשוערת: ${guessFilePurpose()}`,
      numericColumns.length ? `עמודות שנראות מספריות: ${numericColumns.join(", ")}` : "לא זוהו עמודות מספריות באופן ברור.",
      dateColumns.length ? `עמודות שנראות תאריכים: ${dateColumns.join(", ")}` : "לא זוהו עמודות תאריך באופן ברור."
    ].join("\n");
  }

  return [
    `File name: ${uploadedFileName}`,
    `Detected rows: ${excelRows.length}`,
    `Detected columns: ${excelColumns.join(", ")}`,
    `Estimated file purpose: ${guessFilePurpose()}`,
    numericColumns.length ? `Columns that look numeric: ${numericColumns.join(", ")}` : "No clearly numeric columns were detected.",
    dateColumns.length ? `Columns that look like dates: ${dateColumns.join(", ")}` : "No clearly date-based columns were detected."
  ].join("\n");
}

function reviewFileIssues() {
  if (!excelRows.length || !excelColumns.length) {
    return currentLang === "he"
      ? "לא נטען קובץ ולכן עדיין אי אפשר לבדוק שגיאות."
      : "No file has been uploaded yet, so file issues cannot be checked.";
  }

  const messages = [];
  const numericColumns = getNumericColumns();

  excelColumns.forEach((col) => {
    const values = excelRows.map(row => row[col]);
    const emptyCount = values.filter(v => v === "" || v == null).length;
    const duplicateCount = values.filter((v, i, arr) => v !== "" && arr.indexOf(v) !== i).length;

    if (emptyCount > 0) {
      messages.push(currentLang === "he"
        ? `בעמודה "${col}" יש ${emptyCount} תאים ריקים.`
        : `Column "${col}" has ${emptyCount} empty cells.`);
    }

    if (duplicateCount > 0 && duplicateCount >= Math.max(3, Math.floor(values.length * 0.15))) {
      messages.push(currentLang === "he"
        ? `בעמודה "${col}" יש הרבה ערכים כפולים.`
        : `Column "${col}" has many duplicate values.`);
    }
  });

  numericColumns.forEach((col) => {
    const values = excelRows.map(row => row[col]).filter(v => v !== "" && v != null);
    const invalid = values.filter(v => isNaN(Number(String(v).replace(/,/g, ""))));
    if (invalid.length) {
      messages.push(currentLang === "he"
        ? `בעמודה מספרית "${col}" זוהו ערכים שלא נראים מספריים.`
        : `In numeric-looking column "${col}", some values do not look numeric.`);
    }
  });

  if (excelColumns.length !== new Set(excelColumns.map(c => normalizeText(c))).size) {
    messages.push(currentLang === "he"
      ? "יש שמות עמודות כפולים או כמעט כפולים."
      : "There are duplicate or nearly duplicate column names.");
  }

  if (!messages.length) {
    return currentLang === "he"
      ? "לא זוהו שגיאות בולטות בבדיקה המהירה. הקובץ נראה תקין יחסית."
      : "No major issues were found in the quick review. The file looks relatively clean.";
  }

  const recommendation = currentLang === "he"
    ? "\n\nהמלצה: בדקי עמודות עם תאים ריקים, ערכים כפולים, או עמודות מספריות עם טקסט, ותקני אותן לפני חישובים חשובים."
    : "\n\nRecommendation: review columns with empty cells, duplicate values, or numeric columns containing text before important calculations.";

  return messages.join("\n") + recommendation;
}

function fixFormula(raw) {
  let formula = String(raw || "").trim();

  if (!formula) return "";

  if (!formula.startsWith("=")) {
    formula = "=" + formula;
  }

  formula = formula.replace(/;/g, ",");
  formula = formula.replace(/\s+/g, "");
  formula = formula.replace(/סאם/gi, "SUM");
  formula = formula.replace(/אברג|averagee/gi, "AVERAGE");
  formula = formula.replace(/קאונטיף|countiff/gi, "COUNTIF");
  formula = formula.replace(/xlokup|xlookupp/gi, "XLOOKUP");
  formula = formula.replace(/וילוקאפ|vlookupp/gi, "VLOOKUP");
  formula = formula.replace(/ifif/gi, "IF");
  formula = formula.replace(/^=sum\(/i, "=SUM(");
  formula = formula.replace(/^=average\(/i, "=AVERAGE(");
  formula = formula.replace(/^=countif\(/i, "=COUNTIF(");

  const openCount = (formula.match(/\(/g) || []).length;
  const closeCount = (formula.match(/\)/g) || []).length;
  if (openCount > closeCount) {
    formula += ")".repeat(openCount - closeCount);
  }

  return formula;
}

function buildSingleFormula(prompt) {
  const ctx = getContext(prompt);
  const text = ctx.text;
  const conditionValue = detectConditionValue(text);

  if (containsAny(text, ["fixformula", "תקן נוסחה", "תקן לי את הנוסחה", "fix formula"])) {
    const repaired = fixFormula(ctx.fixedPrompt.replace(/תקן לי את הנוסחה|תקן נוסחה|fix formula/gi, "").trim());
    return makeResult(
      currentLang === "he" ? "תיקון נוסחה" : "Formula fix",
      repaired || "=SUM(A:A)",
      currentLang === "he" ? "בוצע תיקון בסיסי לנוסחה לפי כללים מקומיים." : "A basic local repair was applied to the formula.",
      currentLang === "he" ? "אפשר עכשיו לבדוק אם היא באמת תואמת למבנה הקובץ." : "You can now verify whether it matches the file structure.",
      currentLang === "he"
        ? ["בדקי שהתאים והעמודות באמת קיימים.", "אפשר להדביק גם נוסחה שבורה חלקית."]
        : ["Check that the cells and columns really exist.", "You can also paste a partially broken formula."]
    );
  }

  if (containsAny(text, ["countifs", "שני תנאים", "כמה תנאים", "וגם", "גם"])) {
    const firstCol = excelColumns[0] || "C";
    const secondCol = excelColumns[1] || "D";
    return makeResult(
      "COUNTIFS",
      excelFormula("COUNTIFS", `${firstCol}:${firstCol},"Approved",${secondCol}:${secondCol},"Yes"`),
      currentLang === "he" ? "הנוסחה סופרת שורות שמתאימות לשני תנאים יחד." : "The formula counts rows matching two conditions together.",
      currentLang === "he" ? "מתאים לדוחות עם שני תנאים במקביל." : "Useful for reports with two conditions.",
      currentLang === "he"
        ? ["אפשר לשנות את שתי העמודות.", "אפשר לשנות את ערכי התנאי לפי הצורך."]
        : ["You can change both columns.", "You can change the condition values as needed."]
    );
  }

  if (containsAny(text, ["sumifs", "כמה תנאי סכום", "סכום עם כמה תנאים"])) {
    const sumCol = excelColumns[0] || "B";
    const cond1 = excelColumns[1] || "C";
    const cond2 = excelColumns[2] || "D";
    return makeResult(
      "SUMIFS",
      excelFormula("SUMIFS", `${sumCol}:${sumCol},${cond1}:${cond1},"Approved",${cond2}:${cond2},"Yes"`),
      currentLang === "he" ? "הנוסחה מסכמת ערכים לפי כמה תנאים יחד." : "The formula sums values using multiple conditions.",
      currentLang === "he" ? "שימושי לסכומי דוחות מורכבים." : "Useful for complex report totals.",
      currentLang === "he"
        ? ["העמודה הראשונה היא עמודת הסכום.", "אחריה באים זוגות של עמודות תנאי וערכי תנאי."]
        : ["The first column is the sum column.", "Then come condition columns and condition values."]
    );
  }

  if (containsAny(text, ["sumif", "סכם רק אם", "סכום רק אם", "סכום בתנאי"])) {
    const sumCol = findColumnFromUploadedData(ctx.fixedPrompt) || excelColumns[0] || "B";
    const condCol = excelColumns[1] || "C";
    return makeResult(
      "SUMIF",
      excelFormula("SUMIF", `${condCol}:${condCol},"Approved",${sumCol}:${sumCol}`),
      currentLang === "he" ? "הנוסחה מסכמת ערכים רק בשורות שבהן מתקיים תנאי מסוים." : "The formula sums values only in rows that meet a condition.",
      currentLang === "he" ? "מתאים לסכומים לפי סטטוס." : "Useful for totals by status.",
      currentLang === "he"
        ? ["החלק הראשון הוא עמודת התנאי.", "החלק האחרון הוא עמודת הסכום."]
        : ["The first part is the condition column.", "The last part is the sum column."]
    );
  }

  if (containsAny(text, ["חבר", "סכום", "סכם", "סכימה", "sum"])) {
    return makeResult(
      "SUM",
      excelFormula("SUM", ctx.area),
      currentLang === "he" ? `הנוסחה מחברת את כל הערכים בטווח ${ctx.area}.` : `The formula adds all values in the range ${ctx.area}.`,
      currentLang === "he" ? "שימושי לסכומים, מכירות, שעות ותקציבים." : "Useful for totals, sales, hours, and budgets.",
      currentLang === "he"
        ? ["אפשר גם לעבוד על טווח כמו A1:A20.", "בקובץ נטען אפשר לכתוב את שם העמודה עצמה."]
        : ["You can also use a range like A1:A20.", "With an uploaded file you can use the actual column name."]
    );
  }

  if (containsAny(text, ["ממוצע", "average"])) {
    return makeResult(
      "AVERAGE",
      excelFormula("AVERAGE", ctx.area),
      currentLang === "he" ? `הנוסחה מחשבת ממוצע של הערכים בטווח ${ctx.area}.` : `The formula calculates the average of values in ${ctx.area}.`,
      currentLang === "he" ? "מתאים לציונים, מחירים ועלויות." : "Useful for scores, prices, and costs.",
      currentLang === "he"
        ? ["Excel מתעלם מתאים ריקים בדרך כלל.", "כדאי להשתמש בעמודה מספרית."]
        : ["Excel usually ignores empty cells.", "Prefer a numeric column."]
    );
  }

  if (containsAny(text, ["הכי גבוה", "גבוה ביותר", "מקסימום", "max"])) {
    return makeResult(
      "MAX",
      excelFormula("MAX", ctx.area),
      currentLang === "he" ? `הנוסחה מחזירה את הערך הגבוה ביותר בטווח ${ctx.area}.` : `The formula returns the highest value in ${ctx.area}.`,
      currentLang === "he" ? "מתאים למציאת שיאים." : "Useful for finding peaks.",
      currentLang === "he"
        ? ["מומלץ לעבוד עם עמודה מספרית.", "לערך הנמוך ביותר משתמשים ב-MIN."]
        : ["Prefer a numeric column.", "Use MIN for the lowest value."]
    );
  }

  if (containsAny(text, ["הכי נמוך", "נמוך ביותר", "מינימום", "min"])) {
    return makeResult(
      "MIN",
      excelFormula("MIN", ctx.area),
      currentLang === "he" ? `הנוסחה מחזירה את הערך הנמוך ביותר בטווח ${ctx.area}.` : `The formula returns the lowest value in ${ctx.area}.`,
      currentLang === "he" ? "מתאים למציאת מינימום." : "Useful for finding minimum values.",
      currentLang === "he"
        ? ["מומלץ לעבוד עם עמודה מספרית.", "לערך הגבוה ביותר משתמשים ב-MAX."]
        : ["Prefer a numeric column.", "Use MAX for the highest value."]
    );
  }

  if (containsAny(text, ["ספור", "ספירה", "כמה פעמים", "countif", "count"])) {
    return makeResult(
      "COUNTIF",
      excelFormula("COUNTIF", `${ctx.column}:${ctx.column},"${conditionValue}"`),
      currentLang === "he"
        ? `הנוסחה סופרת כמה פעמים הערך "${conditionValue}" מופיע בעמודה ${ctx.column}.`
        : `The formula counts how many times "${conditionValue}" appears in column ${ctx.column}.`,
      currentLang === "he" ? "מתאים לסטטוסים, תשובות וערכים חוזרים." : "Useful for statuses, responses, and repeated values.",
      currentLang === "he"
        ? ["אפשר להחליף את הטקסט לכל ערך אחר.", 'אפשר גם להשתמש בתנאי כמו ">100".']
        : ["You can replace the text with any other value.", 'You can also use conditions like ">100".']
    );
  }

  if (containsAny(text, ["בדוק אם", "אם", "גדול", "קטן", "שווה", "if"])) {
    return makeResult(
      "IF",
      excelFormula("IF", `${ctx.cell}>100,"כן","לא"`),
      currentLang === "he"
        ? `הנוסחה בודקת אם הערך בתא ${ctx.cell} גדול מ-100.`
        : `The formula checks whether the value in ${ctx.cell} is greater than 100.`,
      currentLang === "he" ? "מתאים לבדיקות תנאי פשוטות." : "Useful for simple condition checks.",
      currentLang === "he"
        ? ["אפשר לשנות את 100.", "אפשר לשנות את התוצאות כן/לא."]
        : ["You can change 100.", "You can change the Yes/No results."]
    );
  }

  if (containsAny(text, ["vlookup", "חפש אנכי"])) {
    return makeResult(
      "VLOOKUP",
      '=VLOOKUP(A2,A:B,2,FALSE)',
      currentLang === "he"
        ? "הנוסחה מחפשת ערך בעמודה הראשונה ומחזירה ערך מעמודה אחרת."
        : "The formula looks up a value in the first column and returns a value from another column.",
      currentLang === "he" ? "מתאים לטבלאות ישנות בלי XLOOKUP." : "Useful for older tables without XLOOKUP.",
      currentLang === "he"
        ? ["החיפוש נעשה רק מהעמודה הראשונה.", "אם יש XLOOKUP עדיף להשתמש בו."]
        : ["The lookup works only from the first column.", "If XLOOKUP is available, it is usually better."]
    );
  }

  if (containsAny(text, ["חפש", "חיפוש", "קוד מוצר", "מצא מחיר", "lookup", "xlookup"])) {
    const firstCol = excelColumns[0] || "A";
    const secondCol = excelColumns[1] || "B";
    return makeResult(
      "XLOOKUP",
      excelFormula("XLOOKUP", `A2,${firstCol}:${firstCol},${secondCol}:${secondCol},"לא נמצא"`),
      currentLang === "he"
        ? "הנוסחה מחפשת ערך לפי מפתח ומחזירה ערך מתאים מעמודה אחרת."
        : "The formula looks up a key and returns a matching value from another column.",
      currentLang === "he" ? "מתאים לחיפוש מחיר לפי קוד מוצר או נתון לפי מזהה." : "Useful for price by product code or value by ID.",
      currentLang === "he"
        ? ["בקובץ נטען המערכת משתמשת בעמודות הראשונות כברירת מחדל.", "אפשר להתאים ידנית את עמודות החיפוש וההחזרה."]
        : ["With an uploaded file the app uses the first columns by default.", "You can manually adjust the lookup and return columns."]
    );
  }

  return makeResult(
    currentLang === "he" ? "ברירת מחדל" : "Default",
    excelFormula("SUM", `${ctx.column}:${ctx.column}`),
    currentLang === "he"
      ? `לא זוהתה בקשה מדויקת, לכן הוחזרה נוסחת ברירת מחדל לעמודה ${ctx.column}.`
      : `No exact request was detected, so a default formula was returned for column ${ctx.column}.`,
    currentLang === "he"
      ? "נסי לכתוב בקשה ברורה יותר, למשל: חשב ממוצע של עמודה B."
      : "Try a clearer request, for example: calculate the average of column B.",
    currentLang === "he"
      ? ["אפשר לציין עמודה, תא או טווח.", "אפשר גם לציין שם עמודה מהקובץ שנטען."]
      : ["You can specify a column, cell, or range.", "You can also mention a column name from the uploaded file."]
  );
}

function buildMultipleFormulas(prompt) {
  const ctx = getContext(prompt);
  const col = ctx.column;
  const cell = ctx.cell;
  const firstCol = excelColumns[0] || "A";
  const secondCol = excelColumns[1] || "B";
  const thirdCol = excelColumns[2] || "C";

  return [
    makeResult("SUM", excelFormula("SUM", `${col}:${col}`), currentLang === "he" ? "סכום של העמודה או הטווח." : "Sum of the column or range.", "", []),
    makeResult("AVERAGE", excelFormula("AVERAGE", `${col}:${col}`), currentLang === "he" ? "ממוצע של העמודה או הטווח." : "Average of the column or range.", "", []),
    makeResult("MAX", excelFormula("MAX", `${col}:${col}`), currentLang === "he" ? "הערך הגבוה ביותר." : "Highest value.", "", []),
    makeResult("MIN", excelFormula("MIN", `${col}:${col}`), currentLang === "he" ? "הערך הנמוך ביותר." : "Lowest value.", "", []),
    makeResult("COUNTIF", excelFormula("COUNTIF", `${col}:${col},"Yes"`), currentLang === "he" ? "ספירת Yes." : "Count of Yes.", "", []),
    makeResult("IF", excelFormula("IF", `${cell}>100,"כן","לא"`), currentLang === "he" ? "בדיקת תנאי." : "Condition check.", "", []),
    makeResult("SUMIF", excelFormula("SUMIF", `${secondCol}:${secondCol},"Approved",${firstCol}:${firstCol}`), currentLang === "he" ? "סכום לפי תנאי." : "Conditional sum.", "", []),
    makeResult("COUNTIFS", excelFormula("COUNTIFS", `${secondCol}:${secondCol},"Approved",${thirdCol}:${thirdCol},"Yes"`), currentLang === "he" ? "ספירה לפי שני תנאים." : "Count with two conditions.", "", []),
    makeResult("XLOOKUP", excelFormula("XLOOKUP", `A2,${firstCol}:${firstCol},${secondCol}:${secondCol},"לא נמצא"`), currentLang === "he" ? "חיפוש לפי מפתח." : "Lookup by key.", "", [])
  ];
}

function showSingleResult(data) {
  formulaText.textContent = data.formula;
  explanationText.textContent = data.explanation;
  exampleText.textContent = data.example;

  tipsList.innerHTML = "";
  (data.tips || []).forEach((tip) => {
    const li = document.createElement("li");
    li.textContent = tip;
    tipsList.appendChild(li);
  });

  multiResult.classList.add("hidden");
  singleResult.classList.remove("hidden");
}

function showMultiResults(items) {
  formulaGrid.innerHTML = "";
  items.forEach((item) => {
    const div = document.createElement("div");
    div.className = "multi-item";
    div.innerHTML = `
      <h3>${item.title}</h3>
      <p>${item.explanation}</p>
      <pre>${item.formula}</pre>
    `;
    formulaGrid.appendChild(div);
  });
  multiResult.classList.remove("hidden");
}

function readCsvText(text) {
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");
  if (!lines.length) return [];

  const rows = lines.map((line) => {
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
  return rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header || `Column${index + 1}`] = row[index] ?? "";
    });
    return obj;
  });
}

function loadRows(rows, fileName) {
  excelRows = Array.isArray(rows) ? rows : [];
  excelColumns = Object.keys(excelRows[0] || {});
  uploadedFileName = fileName || "";

  renderColumns();
  renderPreview();
  renderSuggestions();
  renderFileInsight();
  renderFileReview();

  if (excelColumns.length) {
    defaultColumnInput.value = excelColumns[0];
  }

  showStatus(
    excelColumns.length
      ? (currentLang === "he"
          ? `הקובץ "${uploadedFileName}" נטען בהצלחה. זוהו ${excelColumns.length} עמודות ו-${excelRows.length} שורות.`
          : `The file "${uploadedFileName}" was loaded successfully. ${excelColumns.length} columns and ${excelRows.length} rows were detected.`)
      : (currentLang === "he"
          ? `הקובץ "${uploadedFileName}" נטען, אבל לא זוהו עמודות.`
          : `The file "${uploadedFileName}" was loaded, but no columns were detected.`)
  );
}

function renderColumns() {
  columnsList.innerHTML = "";

  if (!excelColumns.length) {
    columnsSection.classList.add("hidden");
    return;
  }

  excelColumns.forEach((col) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "column-chip";
    chip.textContent = col;
    chip.addEventListener("click", () => {
      userPrompt.value = currentLang === "he" ? `חשב ממוצע של ${col}` : `Calculate average of ${col}`;
      defaultColumnInput.value = col;
      userPrompt.focus();
    });
    columnsList.appendChild(chip);
  });

  columnsSection.classList.remove("hidden");
}

function renderPreview() {
  previewTable.innerHTML = "";

  if (!excelRows.length || !excelColumns.length) {
    previewSection.classList.add("hidden");
    return;
  }

  const thead = document.createElement("thead");
  const trHead = document.createElement("tr");

  excelColumns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    trHead.appendChild(th);
  });

  thead.appendChild(trHead);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  excelRows.slice(0, 5).forEach((row) => {
    const tr = document.createElement("tr");
    excelColumns.forEach((col) => {
      const td = document.createElement("td");
      td.textContent = row[col] ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewTable.appendChild(tbody);
  previewSection.classList.remove("hidden");
}

function renderFileInsight() {
  fileInsightText.textContent = summarizeFile();
  fileInsightSection.classList.remove("hidden");
}

function renderFileReview() {
  fileIssuesText.textContent = reviewFileIssues();
  fileIssuesSection.classList.remove("hidden");
}

function renderSuggestions() {
  suggestionsList.innerHTML = "";
  const source = excelColumns.length
    ? [
        currentLang === "he" ? `חשב ממוצע של ${excelColumns[0]}` : `Calculate average of ${excelColumns[0]}`,
        currentLang === "he" ? "מצא שגיאות בקובץ" : "Find issues in the file",
        currentLang === "he" ? "סכם את הקובץ" : "Summarize the file",
        currentLang === "he" ? "עזור לי להבין למה הקובץ משמש" : "Help me understand what the file is for",
        currentLang === "he" ? "תקן לי את הנוסחה =sum(a:a" : "Fix this formula =sum(a:a"
      ]
    : DEFAULT_SUGGESTIONS[currentLang];

  source.forEach((item) => {
    const div = document.createElement("div");
    div.className = "suggestion-item";
    div.textContent = item;
    div.addEventListener("click", () => {
      userPrompt.value = item;
      userPrompt.focus();
    });
    suggestionsList.appendChild(div);
  });
}

function renderHistory() {
  const key = currentLang === "he" ? "helper_excel_reading_history" : "helper_excel_reading_history_en";
  const history = JSON.parse(localStorage.getItem(key) || "[]");
  historyList.innerHTML = "";

  if (!history.length) {
    historyList.innerHTML = `<div class="muted">${t("noHistory")}</div>`;
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

function saveHistory(text) {
  const key = currentLang === "he" ? "helper_excel_reading_history" : "helper_excel_reading_history_en";
  const history = JSON.parse(localStorage.getItem(key) || "[]");
  const updated = [text, ...history.filter(item => item !== text)].slice(0, 10);
  localStorage.setItem(key, JSON.stringify(updated));
  renderHistory();
}

function clearAll() {
  userPrompt.value = "";
  defaultColumnInput.value = excelColumns[0] || "A";
  formulaText.textContent = "";
  explanationText.textContent = "";
  exampleText.textContent = "";
  tipsList.innerHTML = "";
  formulaGrid.innerHTML = "";
  assistantAnswer.textContent = t("assistantDefault");
  multiResult.classList.add("hidden");
  clearStatus();
}

function startVoiceInput() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    showStatus(
      currentLang === "he"
        ? "הדפדפן הזה לא תומך בזיהוי דיבור."
        : "This browser does not support speech recognition.",
      "error"
    );
    return;
  }

  const recognition = new SpeechRecognition();
  recognition.lang = currentLang === "he" ? "he-IL" : "en-US";
  recognition.interimResults = false;
  recognition.maxAlternatives = 1;

  showStatus(currentLang === "he" ? "מקשיב..." : "Listening...");
  recognition.start();

  recognition.onresult = function (event) {
    const text = event.results[0][0].transcript;
    userPrompt.value = text;
    showStatus(currentLang === "he" ? "הטקסט נקלט מהדיבור." : "Voice text captured.");
  };

  recognition.onerror = function () {
    showStatus(currentLang === "he" ? "לא הצלחתי לקלוט דיבור." : "Could not capture speech.", "error");
  };
}

function updateLanguageUI() {
  document.documentElement.lang = currentLang;
  document.documentElement.dir = currentLang === "he" ? "rtl" : "ltr";

  document.getElementById("appTitle").textContent = t("title");
  document.getElementById("appSubtitle").textContent = t("subtitle");
  document.getElementById("languageLabel").textContent = t("language");
  document.getElementById("promptLabel").textContent = t("promptLabel");
  userPrompt.placeholder = t("promptPlaceholder");
  document.getElementById("defaultColumnLabel").textContent = t("defaultColumn");
  generateBtn.textContent = t("generate");
  analyzeBtn.textContent = t("analyze");
  reviewFileBtn.textContent = t("reviewFile");
  multiBtn.textContent = t("multi");
  fixFormulaBtn.textContent = t("fixFormula");
  uploadBtn.textContent = t("upload");
  voiceBtn.textContent = t("voice");
  clearBtn.textContent = t("clear");
  clearHistoryBtn.textContent = t("clearHistory");
  document.getElementById("assistantAnswerTitle").textContent = t("assistantAnswer");
  document.getElementById("formulaTitle").textContent = t("formula");
  document.getElementById("explanationTitle").textContent = t("explanation");
  document.getElementById("exampleTitle").textContent = t("example");
  document.getElementById("tipsTitle").textContent = t("tips");
  document.getElementById("multiTitle").textContent = t("multiTitle");
  document.getElementById("fileInsightTitle").textContent = t("fileInsight");
  document.getElementById("fileIssuesTitle").textContent = t("fileIssues");
  document.getElementById("columnsTitle").textContent = t("columns");
  document.getElementById("previewTitle").textContent = t("preview");
  document.getElementById("suggestionsTitle").textContent = t("suggestions");
  document.getElementById("historyTitle").textContent = t("history");
  copyBtn.textContent = t("copy");
  copyAllBtn.textContent = t("copyAll");
  copyAnswerBtn.textContent = t("copy");

  if (!assistantAnswer.textContent.trim() || assistantAnswer.textContent === UI_TEXT.he.assistantDefault || assistantAnswer.textContent === UI_TEXT.en.assistantDefault) {
    assistantAnswer.textContent = t("assistantDefault");
  }

  renderSuggestions();
  renderHistory();
  if (excelRows.length) {
    renderFileInsight();
    renderFileReview();
  }
}

uploadBtn.addEventListener("click", () => {
  fileInput.click();
});

fileInput.addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const name = file.name.toLowerCase();
  uploadedFileName = file.name;

  if (name.endsWith(".csv")) {
    const reader = new FileReader();
    reader.onload = function (evt) {
      try {
        const rows = readCsvText(evt.target.result);
        loadRows(rows, file.name);
      } catch {
        showStatus(currentLang === "he" ? "לא ניתן לקרוא את קובץ ה-CSV." : "Could not read the CSV file.", "error");
      }
    };
    reader.readAsText(file, "utf-8");
    return;
  }

  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const reader = new FileReader();
    reader.onload = function (evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        loadRows(rows, file.name);
      } catch {
        showStatus(currentLang === "he" ? "לא ניתן לקרוא את קובץ האקסל." : "Could not read the Excel file.", "error");
      }
    };
    reader.readAsArrayBuffer(file);
    return;
  }

  showStatus(currentLang === "he" ? "יש לבחור קובץ Excel או CSV בלבד." : "Please choose an Excel or CSV file only.", "error");
});

generateBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  if (!prompt) {
    showStatus(currentLang === "he" ? "יש לכתוב בקשה לפני יצירת נוסחה." : "Please enter a request before creating a formula.", "error");
    return;
  }

  const result = buildSingleFormula(prompt);
  showSingleResult(result);
  assistantAnswer.textContent = buildAssistantAnswer(prompt, result);
  saveHistory(prompt);
  showStatus(currentLang === "he" ? "הנוסחה נוצרה בהצלחה." : "The formula was created successfully.");
});

analyzeBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  if (!prompt && !excelRows.length) {
    showStatus(currentLang === "he" ? "כתבי בקשה או העלי קובץ כדי שאוכל לנתח." : "Write a request or upload a file so I can analyze.", "error");
    return;
  }

  const result = buildSingleFormula(prompt || (currentLang === "he" ? "סכם את הקובץ" : "Summarize the file"));
  assistantAnswer.textContent = buildAssistantAnswer(prompt || "", result);

  if (excelRows.length) {
    renderFileInsight();
  }

  showStatus(currentLang === "he" ? "הבקשה נותחה." : "The request was analyzed.");
});

reviewFileBtn.addEventListener("click", () => {
  if (!excelRows.length) {
    showStatus(currentLang === "he" ? "צריך קודם להעלות קובץ." : "Please upload a file first.", "error");
    return;
  }

  renderFileReview();
  assistantAnswer.textContent = reviewFileIssues();
  showStatus(currentLang === "he" ? "בדיקת הקובץ הושלמה." : "File review completed.");
});

multiBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  if (!prompt) {
    showStatus(currentLang === "he" ? "יש לכתוב בקשה לפני יצירת נוסחאות." : "Please enter a request before creating formulas.", "error");
    return;
  }

  const results = buildMultipleFormulas(prompt);
  showMultiResults(results);
  assistantAnswer.textContent = currentLang === "he"
    ? "יצרתי כמה נוסחאות אפשריות לפי הבקשה שלך. בחרי את זו שהכי מתאימה."
    : "I created multiple possible formulas based on your request. Choose the one that fits best.";
  saveHistory(prompt);
  showStatus(currentLang === "he" ? "נוצרו כמה נוסחאות אפשריות." : "Multiple possible formulas were created.");
});

fixFormulaBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  if (!prompt) {
    showStatus(currentLang === "he" ? "יש להדביק נוסחה או טקסט לתיקון." : "Please paste a formula or text to fix.", "error");
    return;
  }

  const fixed = fixFormula(prompt);
  formulaText.textContent = fixed || "=SUM(A:A)";
  explanationText.textContent = currentLang === "he"
    ? "בוצע תיקון בסיסי לנוסחה: הוספת '=', ניקוי רווחים, החלפת ';' ב-',' ואיזון סוגריים."
    : "A basic formula repair was applied: adding '=', removing spaces, replacing ';' with ',', and balancing parentheses.";
  exampleText.textContent = currentLang === "he"
    ? "אם עדיין זו לא הנוסחה שרצית, כתבי במילים מה היא אמורה לעשות."
    : "If this is still not the formula you wanted, describe in words what it should do.";
  tipsList.innerHTML = "";

  [
    currentLang === "he" ? "בדקי שהתאים והעמודות באמת קיימים." : "Check that the cells and columns really exist.",
    currentLang === "he" ? "אפשר לתקן גם נוסחה שבורה חלקית." : "You can also repair a partially broken formula."
  ].forEach((tip) => {
    const li = document.createElement("li");
    li.textContent = tip;
    tipsList.appendChild(li);
  });

  assistantAnswer.textContent = currentLang === "he"
    ? `נוסחה מתוקנת:\n${fixed}`
    : `Fixed formula:\n${fixed}`;

  showStatus(currentLang === "he" ? "בוצע תיקון לנוסחה." : "The formula was repaired.");
});

voiceBtn.addEventListener("click", startVoiceInput);
clearBtn.addEventListener("click", clearAll);

clearHistoryBtn.addEventListener("click", () => {
  const key = currentLang === "he" ? "helper_excel_reading_history" : "helper_excel_reading_history_en";
  localStorage.removeItem(key);
  renderHistory();
  showStatus(currentLang === "he" ? "ההיסטוריה נוקתה." : "History was cleared.");
});

copyBtn.addEventListener("click", async () => {
  const text = formulaText.textContent.trim();
  if (!text) return;

  try {
    await navigator.clipboard.writeText(text);
    copyBtn.textContent = t("copied");
    setTimeout(() => {
      copyBtn.textContent = t("copy");
    }, 1200);
  } catch {
    showStatus(currentLang === "he" ? "לא ניתן להעתיק כרגע." : "Cannot copy right now.", "error");
  }
});

copyAllBtn.addEventListener("click", async () => {
  const all = Array.from(document.querySelectorAll(".multi-item pre"))
    .map(el => el.textContent.trim())
    .join("\n\n");

  if (!all) return;

  try {
    await navigator.clipboard.writeText(all);
    copyAllBtn.textContent = t("copied");
    setTimeout(() => {
      copyAllBtn.textContent = t("copyAll");
    }, 1200);
  } catch {
    showStatus(currentLang === "he" ? "לא ניתן להעתיק כרגע." : "Cannot copy right now.", "error");
  }
});

copyAnswerBtn.addEventListener("click", async () => {
  const text = assistantAnswer.textContent.trim();
  if (!text) return;

  try {
    await navigator.clipboard.writeText(text);
    copyAnswerBtn.textContent = t("copied");
    setTimeout(() => {
      copyAnswerBtn.textContent = t("copy");
    }, 1200);
  } catch {
    showStatus(currentLang === "he" ? "לא ניתן להעתיק כרגע." : "Cannot copy right now.", "error");
  }
});

languageSelect.addEventListener("change", () => {
  currentLang = languageSelect.value;
  updateLanguageUI();
});

updateLanguageUI();
renderSuggestions();
renderHistory();
assistantAnswer.textContent = t("assistantDefault");
