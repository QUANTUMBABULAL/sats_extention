const fileInput = document.getElementById("fileInput");
const loadButton = document.getElementById("loadButton");
const startButton = document.getElementById("startButton");
const statusElement = document.getElementById("status");
const fileNameElement = document.getElementById("fileName");

let workbookRows = null;

fileInput.addEventListener("change", handleFileChange);
loadButton.addEventListener("click", () => fileInput.click());
startButton.addEventListener("click", handleStartClick);

async function handleFileChange(event) {
  const [file] = event.target.files || [];

  if (!file) {
    resetState();
    return;
  }

  if (!/\.xlsx$/i.test(file.name)) {
    showStatus("Only .xlsx files are supported by this extension.", true);
    resetState(false);
    return;
  }

  try {
    showStatus("Loading workbook...");
    const buffer = await file.arrayBuffer();
    workbookRows = parseWorkbook(buffer);
    updateSummary(file.name);
    showStatus("File loaded");
    startButton.disabled = false;
  } catch (error) {
    console.error(error);
    resetState(false);
    showStatus(`Could not parse workbook: ${error.message}`, true);
  }
}

async function handleStartClick() {
  if (!workbookRows) {
    showStatus("Upload and parse a workbook before sending.", true);
    return;
  }

  try {
    showStatus("Automation started");
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab?.id) {
      throw new Error("No active tab was found.");
    }

    const response = await chrome.tabs.sendMessage(tab.id, {
      type: "EXCEL_WORKBOOK_DATA",
      payload: workbookRows,
    });

    if (!response?.ok) {
      throw new Error("The content script did not acknowledge the message.");
    }

    showStatus("Completed");
  } catch (error) {
    console.error(error);
    showStatus(`Send failed: ${error.message}`, true);
  }
}

function resetState(updateStatus = true) {
  workbookRows = null;
  startButton.disabled = true;
  fileNameElement.textContent = "No file selected";

  if (updateStatus) {
    showStatus("Choose an Excel file to start.");
  }
}

function updateSummary(fileName) {
  fileNameElement.textContent = fileName;
}

function showStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.classList.toggle("error", isError);
}

function parseWorkbook(arrayBuffer) {
  if (typeof XLSX === "undefined") {
    throw new Error("SheetJS failed to load.");
  }

  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    return [];
  }

  const worksheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(worksheet, { defval: "" });
}

resetState(false);
