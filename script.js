const userPrompt = document.getElementById("userPrompt");
const generateBtn = document.getElementById("generateBtn");
const multiBtn = document.getElementById("multiBtn");
const clearBtn = document.getElementById("clearBtn");
const clearHistoryBtn = document.getElementById("clearHistoryBtn");

const errorBox = document.getElementById("errorBox");
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

const EXAMPLES = [
  "חבר את כל הערכים בעמודה B",
  "חשב ממוצע של עמודה F",
  "ספור כמה פעמים Approved מופיע בעמודה C",
  "בדוק אם A2 גדול מ-100",
  "מצא את הערך הגבוה ביותר בעמודה G",
  "מצא את הערך הנמוך ביותר בעמודה H",
  "חפש מחיר לפי קוד מוצר",
  "סכם את עמודה B רק אם בעמודה C כתוב Approved",
  "ספור כמה שורות יש שבהן C הוא Approved וגם D הוא Yes"
];

function showError(message) {
  errorBox.textContent = message;
  errorBox.classList.remove("hidden");
}

function hideError() {
  errorBox.classList.add("hidden");
  errorBox.textContent = "";
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

function clearAll() {
  userPrompt.value = "";
  hideError();
  clearSingleResult();
  clearMultiResult();
}

function containsAny(text, words) {
  return words.some(word => text.includes(word));
}

function normalizeColumn(value) {
  if (!value) return "A";
  const clean = value.trim().toUpperCase().replace(/[^A-Z]/g, "");
  return clean || "A";
}

function detectColumn(text) {
  const match = text.match(/עמודה\s*([A-Z]{1,3})/i);
  return match ? normalizeColumn(match[1]) : null;
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

function getContext(prompt) {
  const range = detectRange(prompt);
  const column = detectColumn(prompt) || normalizeColumn(defaultColumnInput.value) || "A";
  const cell = detectCell(prompt) || `${column}2`;
  const area = range || `${column}:${column}`;

  return {
    text: prompt.toLowerCase(),
    column,
    cell,
    range,
    area
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

function buildSingleFormula(prompt) {
  const ctx = getContext(prompt);
  const text = ctx.text;
  const conditionValue = detectConditionValue(text);

  if (containsAny(text, ["countifs", "שני תנאים", "כמה תנאים", "וגם", "גם"])) {
    return makeResult(
      "COUNTIFS",
      excelFormula("COUNTIFS", `C:C,"Approved",D:D,"Yes"`),
      "הנוסחה סופרת שורות שמתאימות לשני תנאים במקביל: Approved בעמודה C וגם Yes בעמודה D.",
      "מתאים לדוחות סטטוסים, אישורים ומעקבים.",
      [
        "אפשר לשנות את עמודות התנאים.",
        "אפשר לשנות את ערכי התנאי לפי הצורך."
      ]
    );
  }

  if (containsAny(text, ["sumif", "סכם רק אם", "סכום רק אם", "סכום בתנאי"])) {
    return makeResult(
      "SUMIF",
      excelFormula("SUMIF", `C:C,"Approved",B:B`),
      "הנוסחה מסכמת את עמודה B רק בשורות שבהן בעמודה C מופיע Approved.",
      "מעולה לסיכום סכומים רק עבור רשומות מאושרות.",
      [
        "העמודה הראשונה היא עמודת התנאי.",
        "העמודה האחרונה היא עמודת הסכום."
      ]
    );
  }

  if (containsAny(text, ["חבר", "סכום", "סכם", "סכימה", "sum"])) {
    return makeResult(
      "SUM",
      excelFormula("SUM", ctx.area),
      `הנוסחה מחברת את כל הערכים בטווח ${ctx.area}.`,
      "שימושי לסכומים, תקציבים, שעות, מכירות ועוד.",
      [
        "אפשר לעבוד על עמודה שלמה או על טווח מוגדר.",
        "Excel מתעלם מתאים ריקים."
      ]
    );
  }

  if (containsAny(text, ["ממוצע", "average"])) {
    return makeResult(
      "AVERAGE",
      excelFormula("AVERAGE", ctx.area),
      `הנוסחה מחשבת ממוצע של הערכים בטווח ${ctx.area}.`,
      "מתאים לציונים, עלויות, ביצועים ועוד.",
      [
        "אפשר לצמצם לטווח מסוים כמו B2:B20.",
        "תאים ריקים בדרך כלל לא ייכללו בחישוב."
      ]
    );
  }

  if (containsAny(text, ["הכי גבוה", "גבוה ביותר", "מקסימום", "max"])) {
    return makeResult(
      "MAX",
      excelFormula("MAX", ctx.area),
      `הנוסחה מחזירה את הערך הגבוה ביותר בטווח ${ctx.area}.`,
      "מתאים למציאת שיא של מחיר, ציון, כמות או זמן.",
      [
        "לערך הנמוך ביותר משתמשים ב-MIN.",
        "אפשר לעבוד גם על טווח מצומצם."
      ]
    );
  }

  if (containsAny(text, ["הכי נמוך", "נמוך ביותר", "מינימום", "min"])) {
    return makeResult(
      "MIN",
      excelFormula("MIN", ctx.area),
      `הנוסחה מחזירה את הערך הנמוך ביותר בטווח ${ctx.area}.`,
      "מתאים למציאת מינימום של מחיר, ציון או כל ערך מספרי.",
      [
        "לערך הגבוה ביותר משתמשים ב-MAX.",
        "אפשר לעבוד גם על טווח מצומצם."
      ]
    );
  }

  if (containsAny(text, ["ספור", "ספירה", "כמה פעמים", "countif", "count"])) {
    return makeResult(
      "COUNTIF",
      excelFormula("COUNTIF", `${ctx.column}:${ctx.column},"${conditionValue}"`),
      `הנוסחה סופרת כמה פעמים הערך "${conditionValue}" מופיע בעמודה ${ctx.column}.`,
      "מתאים לספירת סטטוסים, תשובות, ערכים חוזרים ועוד.",
      [
        "אפשר להחליף את ערך החיפוש לכל טקסט אחר.",
        'אפשר להשתמש גם בתנאים כמו ">100".'
      ]
    );
  }

  if (containsAny(text, ["בדוק אם", "אם", "גדול", "קטן", "שווה", "if"])) {
    return makeResult(
      "IF",
      excelFormula("IF", `${ctx.cell}>100,"כן","לא"`),
      `הנוסחה בודקת אם הערך בתא ${ctx.cell} גדול מ-100. אם כן, יוחזר "כן", אחרת "לא".`,
      "מתאים לבדיקות סטטוס, חריגות, ספים ותנאים עסקיים.",
      [
        "אפשר לשנות את 100 לכל מספר אחר.",
        "אפשר לשנות את התוצאה כן/לא לכל טקסט אחר."
      ]
    );
  }

  if (containsAny(text, ["חפש", "חיפוש", "קוד מוצר", "מצא מחיר", "lookup", "xlookup", "vlookup"])) {
    return makeResult(
      "XLOOKUP",
      excelFormula("XLOOKUP", `A2,Products!A:A,Products!C:C,"לא נמצא"`),
      "הנוסחה מחפשת את הערך בתא A2 בעמודה A של הגיליון Products ומחזירה את הערך המתאים מעמודה C.",
      "מתאים לחיפוש מחיר, שם מוצר, לקוח, סטטוס או כל מידע לפי מזהה.",
      [
        "אפשר לשנות את שם הגיליון Products.",
        "אם אין אצלך XLOOKUP, אפשר להחליף ל-VLOOKUP."
      ]
    );
  }

  return makeResult(
    "ברירת מחדל",
    excelFormula("SUM", `${ctx.column}:${ctx.column}`),
    `לא זוהתה בקשה מדויקת, לכן הוחזרה נוסחת ברירת מחדל לעמודה ${ctx.column}.`,
    "נסי לכתוב בקשה ברורה יותר, למשל: חשב ממוצע של עמודה B.",
    [
      "כדאי לציין עמודה, תא או טווח.",
      "כדאי לציין גם תנאי אם יש צורך."
    ]
  );
}

function buildMultipleFormulas(prompt) {
  const ctx = getContext(prompt);

  return [
    makeResult("SUM", excelFormula("SUM", ctx.area), `סכום של ${ctx.area}.`, "", []),
    makeResult("AVERAGE", excelFormula("AVERAGE", ctx.area), `ממוצע של ${ctx.area}.`, "", []),
    makeResult("MAX", excelFormula("MAX", ctx.area), `מקסימום בטווח ${ctx.area}.`, "", []),
    makeResult("MIN", excelFormula("MIN", ctx.area), `מינימום בטווח ${ctx.area}.`, "", []),
    makeResult("COUNTIF", excelFormula("COUNTIF", `${ctx.column}:${ctx.column},"Yes"`), `ספירת Yes בעמודה ${ctx.column}.`, "", []),
    makeResult("IF", excelFormula("IF", `${ctx.cell}>100,"כן","לא"`), `בדיקה אם ${ctx.cell} גדול מ-100.`, "", []),
    makeResult("SUMIF", excelFormula("SUMIF", `C:C,"Approved",B:B`), "סכום לפי תנאי.", "", []),
    makeResult("COUNTIFS", excelFormula("COUNTIFS", `C:C,"Approved",D:D,"Yes"`), "ספירה לפי שני תנאים.", "", []),
    makeResult("XLOOKUP", excelFormula("XLOOKUP", `A2,Products!A:A,Products!C:C,"לא נמצא"`), "חיפוש ערך לפי מפתח.", "", [])
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

quickTags.addEventListener("click", (event) => {
  const tag = event.target.closest(".tag");
  if (!tag) return;

  const value = tag.textContent.trim();
  const current = userPrompt.value.trim();
  userPrompt.value = current ? `${current} ${value}` : value;
  userPrompt.focus();
});

renderExamples();
renderHistory();
