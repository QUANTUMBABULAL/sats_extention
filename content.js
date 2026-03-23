(() => {
  const MESSAGE_TYPE = "EXCEL_WORKBOOK_DATA";
  const ROW_SELECTOR = "#example tbody tr";
  const ROW_DELAY_MS = 300;
  const FIELD_SELECTORS = [
    { key: "Lan1", selector: '[name="registrationCustomDTO.result.firstl_grade"]' },
    { key: "Lan2", selector: '[name="registrationCustomDTO.result.secondl_grade"]' },
    { key: "Sub1", selector: '[name="registrationCustomDTO.result.sub0_grade"]' },
    { key: "Sub2", selector: '[name="registrationCustomDTO.result.sub1_grade"]' },
  ];

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message?.type !== MESSAGE_TYPE) {
      return;
    }

    runAutomation(Array.isArray(message.payload) ? message.payload : [])
      .then((result) => sendResponse({ ok: true, ...result }))
      .catch((error) => {
        console.error("Automation failed:", error);
        sendResponse({ ok: false, error: error.message });
      });

    return true;
  });

  async function runAutomation(rowsData) {
    const pageRows = Array.from(document.querySelectorAll(ROW_SELECTOR));
    const rowCount = Math.min(pageRows.length, rowsData.length);

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const rowElement = pageRows[rowIndex];
      const rowData = rowsData[rowIndex] || {};

      if (!rowElement) {
        continue;
      }

      updateRow(rowElement, rowData, rowIndex);

      if (rowIndex < rowCount - 1) {
        await sleep(ROW_DELAY_MS);
      }
    }

    window.__excelExtensionData = rowsData;

    return {
      processedRows: rowCount,
      availableRows: pageRows.length,
      receivedRows: rowsData.length,
    };
  }

  function updateRow(rowElement, rowData, rowIndex) {
    for (const field of FIELD_SELECTORS) {
      const dropdown = rowElement.querySelector(field.selector);
      if (!dropdown) {
        continue;
      }

      const excelValue = normalizeExcelValue(rowData?.[field.key]);
      if (!excelValue) {
        continue;
      }

      setDropdownValue(dropdown, excelValue);
    }

    rowElement.dataset.excelRowIndex = String(rowIndex);
  }

  function setDropdownValue(dropdown, excelValue) {
    const desiredValue = `1,${excelValue}`;
    const optionIndex = findMatchingOptionIndex(dropdown, desiredValue, excelValue);
    const matchedOption = optionIndex >= 0 ? dropdown.options[optionIndex] : null;
    const valueToSet = matchedOption?.value || desiredValue;

    setNativeValue(dropdown, valueToSet);
    dropdown.selectedIndex = optionIndex;

    if (dropdown.value !== valueToSet && optionIndex >= 0) {
      setNativeValue(dropdown, dropdown.options[optionIndex].value);
    }

    dropdown.dispatchEvent(new Event("input", { bubbles: true }));
    dropdown.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function normalizeExcelValue(value) {
    if (value === undefined || value === null) {
      return "";
    }

    return String(value).trim();
  }

  function findMatchingOptionIndex(dropdown, desiredValue, excelValue) {
    const options = Array.from(dropdown.options || []);

    return options.findIndex((option) => {
      const optionValue = String(option.value || "").trim();
      const optionLabel = String(option.label || option.textContent || "").trim();

      return (
        optionValue === desiredValue ||
        optionValue === excelValue ||
        optionValue.endsWith(`,${excelValue}`) ||
        optionLabel === desiredValue ||
        optionLabel === excelValue ||
        optionLabel.endsWith(`,${excelValue}`)
      );
    });
  }

  function setNativeValue(dropdown, value) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value");
    if (descriptor?.set) {
      descriptor.set.call(dropdown, value);
      return;
    }

    dropdown.value = value;
  }

  function sleep(milliseconds) {
    return new Promise((resolve) => {
      setTimeout(resolve, milliseconds);
    });
  }
})();