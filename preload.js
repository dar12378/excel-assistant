const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("excelDesktopAPI", {
  pickFile: () => ipcRenderer.invoke("pick-file"),
  readWorkbook: (filePath) => ipcRenderer.invoke("read-workbook", filePath),
  saveFormulasWorkbook: (payload) => ipcRenderer.invoke("save-formulas-workbook", payload)
});
