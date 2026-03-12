const userPrompt = document.getElementById("userPrompt");
const generateBtn = document.getElementById("generateBtn");
const clearBtn = document.getElementById("clearBtn");
const errorBox = document.getElementById("errorBox");
const result = document.getElementById("result");
const formulaText = document.getElementById("formulaText");
const explanationText = document.getElementById("explanationText");
const exampleText = document.getElementById("exampleText");
const tipsList = document.getElementById("tipsList");
const copyBtn = document.getElementById("copyBtn");
const exampleItems = document.querySelectorAll(".example-item");

function showError(message) {
  errorBox.textContent = message;
  errorBox.classList.remove("hidden");
}

function hideError() {
  errorBox.classList.add("hidden");
  errorBox.textContent = "";
}

function showResult(data) {
  formulaText.textContent = data.formula || "";
  explanationText.textContent = data.explanation || "";
  exampleText.textContent = data.example || "";

  tipsList.innerHTML = "";

  if (Array.isArray(data.tips)) {
    data.tips.forEach((tip) => {
      const li = document.createElement("li");
      li.textContent = tip;
      tipsList.appendChild(li);
    });
  }

  result.classList.remove("hidden");
}

function clearAll() {
  userPrompt.value = "";
  hideError();
  result.classList.add("hidden");
  formulaText.textContent = "";
  explanationText.textContent = "";
  exampleText.textContent = "";
  tipsList.innerHTML = "";
}

function containsAny(text, words) {
  return words.some(word => text.includes(word));
}

function detectColumn(text) {
  const match = text.match(/עמודה\\s*([A-Zא-ת])/i);
  if (!match) return null;
  const value = match[1].toUpperCase();
  if (/^[A-Z]$/.test(value)) return value;
  return null;
}

function detectCell(text) {
  const match = text.match(/([A-Z]+\\d+)/i);
  return match ? match[1].toUpperCase() : null;
}

function buildFormula(prompt) {
  const text = prompt.trim().toLowerCase();
  const column = detectColumn(prompt) || "A";
  const cell = detectCell(prompt) || "A2";

  if (containsAny(text, ["חבר", "סכום", "סכם", "סכימה"])) {
    return {
      formula: `=SUM(${column}:${column})`,
      explanation: `הנוסחה מחברת את כל הערכים בעמודה ${column}.`,
      example: `אם יש מספרים בעמודה ${column}, תקבלי את הסכום הכולל שלהם.`,
      tips: [
        "אפשר להחליף את האות של העמודה לפי הצורך.",
        "אפשר גם לעבוד על טווח מסוים, למשל =SUM(A1:A20)."
      ]
    };
  }

  if (containsAny(text, ["ממוצע", "ממוצעים", "average"])) {
    return {
      formula: `=AVERAGE(${column}:${column})`,
      explanation: `הנוסחה מחשבת ממוצע של כל הערכים בעמודה ${column}.`,
      example: `אם בעמודה ${column} יש ציונים או מחירים, תקבלי את הממוצע שלהם.`,
      tips: [
        "אם יש תאים ריקים, Excel מתעלם מהם בדרך כלל.",
        "אפשר לצמצם לטווח מסוים כמו =AVERAGE(B2:B10)."
      ]
    };
  }

  if (containsAny(text, ["גדול", "קטן", "בדוק אם", "אם"])) {
    return {
      formula: `=IF(${cell}>100,"כן","לא")`,
      explanation: `הנוסחה בודקת אם הערך בתא ${cell} גדול מ-100. אם כן, יוחזר \"כן\", אחרת \"לא\".`,
      example: `אפשר לשנות את 100 לכל מספר אחר שתרצי לבדוק.`,
      tips: [
        "אפשר לשנות את התנאי לקטן מ, שווה ל, או טקסט.",
        "אפשר להחליף את 'כן' ו'לא' לכל תוצאה אחרת."
      ]
    };
  }

  if (containsAny(text, ["ספור", "ספירה", "כמה פעמים", "count"])) {
    return {
      formula: `=COUNTIF(${column}:${column},"Yes")`,
      explanation: `הנוסחה סופרת כמה פעמים הערך \"Yes\" מופיע בעמודה ${column}.`,
      example: `אם תרצי ערך אחר, אפשר להחליף את Yes ל-Approved או לכל טקסט אחר.`,
      tips: [
        "חשוב שהטקסט יהיה כתוב בדיוק כמו באקסל.",
        "אפשר לספור גם מספרים או תנאים אחרים."
      ]
    };
  }

  if (containsAny(text, ["הכי גבוה", "מקסימום", "ערך גבוה ביותר", "max"])) {
    return {
      formula: `=MAX(${column}:${column})`,
      explanation: `הנוסחה מחזירה את הערך הגבוה ביותר בעמודה ${column}.`,
      example: `שימושי למציאת המחיר הכי גבוה או הציון הכי גבוה.`,
      tips: [
        "אפשר גם להשתמש בטווח מוגדר במקום עמודה שלמה.",
        "לתוצאה הנמוכה ביותר משתמשים ב-MIN."
      ]
    };
  }

  if (containsAny(text, ["הכי נמוך", "מינימום", "ערך נמוך ביותר", "min"])) {
    return {
      formula: `=MIN(${column}:${column})`,
      explanation: `הנוסחה מחזירה את הערך הנמוך ביותר בעמודה ${column}.`,
      example: `שימושי למציאת המחיר הכי נמוך או הציון הכי נמוך.`,
      tips: [
        "אפשר גם להשתמש בטווח מוגדר במקום עמודה שלמה.",
        "לתוצאה הגבוהה ביותר משתמשים ב-MAX."
      ]
    };
  }

  if (containsAny(text, ["חפש", "חיפוש", "מצא מחיר", "קוד מוצר", "lookup"])) {
    return {
      formula: `=XLOOKUP(A2,Products!A:A,Products!C:C,"לא נמצא")`,
      explanation: `הנוסחה מחפשת את הערך שבתא A2 בגיליון Products בעמודה A, ומחזירה את הערך המתאים מעמודה C.`,
      example: `אם A2 מכיל קוד מוצר, אפשר להחזיר מחיר, שם מוצר או כל עמודה אחרת.`,
      tips: [
        "Products הוא שם גיליון לדוגמה ואפשר לשנות אותו.",
        "אם אין XLOOKUP אצלך, אפשר להשתמש ב-VLOOKUP."
      ]
    };
  }

  return {
    formula: `=SUM(A:A)`,
    explanation: "לא זוהתה בקשה מדויקת, לכן הוחזרה נוסחת ברירת מחדל של סכום עמודה A.",
    example: "נסי לכתוב בקשה ברורה יותר כמו: חבר את כל הערכים בעמודה B.",
    tips: [
      "כתבי מה בדיוק את רוצה לחשב.",
      "אפשר לציין עמודה כמו B או תא כמו A2."
    ]
  };
}

generateBtn.addEventListener("click", () => {
  const prompt = userPrompt.value.trim();

  hideError();
  result.classList.add("hidden");

  if (!prompt) {
    showError("יש לכתוב בקשה לפני יצירת נוסחה.");
    return;
  }

  const data = buildFormula(prompt);
  showResult(data);
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

exampleItems.forEach((item) => {
  item.addEventListener("click", () => {
    userPrompt.value = item.textContent.trim();
    userPrompt.focus();
  });
});
