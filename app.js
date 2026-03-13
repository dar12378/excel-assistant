window.ExcelHebrewDesktopState = {
  fileName: "",
  sheets: [],
  activeSheetIndex: -1
};

window.ExcelHebrewDesktopHelpers = {
  normalizeText(value) {
    return String(value ?? "")
      .replace(/[\u0591-\u05C7]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  },

  containsAny(text, words) {
    return words.some((word) => text.includes(word));
  },

  detectCell(text) {
    const match = text.match(/([A-Z]+\d+)/i);
    return match ? match[1].toUpperCase() : null;
  },

  excelFormula(name, args) {
    return `=${name.toUpperCase()}(${args})`;
  },

  columnRef(columnName) {
    const clean = String(columnName).trim();
    if (/^[A-Z]{1,3}$/.test(clean.toUpperCase())) {
      const upper = clean.toUpperCase();
      return `${upper}:${upper}`;
    }
    return `[${clean}]`;
  },

  makeResult(title, formula, explanation, example, tips) {
    return { title, formula, explanation, example, tips };
  }
};
