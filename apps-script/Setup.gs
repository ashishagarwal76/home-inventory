function setupInventory() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error(
      "Open your inventory spreadsheet, choose Extensions > Apps Script, and run setupInventory there."
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetID = spreadsheet.getId();
    const savedID = properties.getProperty("INVENTORY_SPREADSHEET_ID");
    const configuredID = String(savedID || "").trim();
    if (configuredID && configuredID !== spreadsheetID) {
      throw new Error(
        "This project is configured for a different spreadsheet. Check INVENTORY_SPREADSHEET_ID before running setupInventory. No setup changes were made."
      );
    }

    const schema = inventoryTemplateSchema_();
    if (!Array.isArray(schema) || !schema.length) {
      throw new Error("The inventory template schema is missing. Check Schema.gs before running setupInventory.");
    }
    const names = new Set();
    const plans = schema.map(function (table) {
      if (!table || typeof table.name !== "string" || !table.name.trim() ||
          names.has(table.name) || !Array.isArray(table.headers) || !table.headers.length ||
          table.headers.some(header => typeof header !== "string" || !header) ||
          new Set(table.headers).size !== table.headers.length) {
        throw new Error("The inventory template schema is invalid. Check Schema.gs before running setupInventory.");
      }
      names.add(table.name);
      const sheet = spreadsheet.getSheetByName(table.name);
      if (!sheet || sheet.getDataRange().isBlank()) {
        return { name: table.name, headers: table.headers, sheet: sheet };
      }

      const width = Math.min(table.headers.length, sheet.getMaxColumns());
      const actual = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
      table.headers.forEach(function (header, index) {
        if (actual[index] !== header) {
          throw new Error(
            'Sheet "' + table.name + '": column ' + (index + 1) +
            ' must be "' + header + '". No setup changes were made.'
          );
        }
      });
      return null;
    });

    // All existing tables have passed validation before the first write.
    plans.filter(Boolean).forEach(function (plan) {
      const sheet = plan.sheet || spreadsheet.insertSheet(plan.name);
      if (sheet.getMaxColumns() < plan.headers.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), plan.headers.length - sheet.getMaxColumns());
      }
      sheet.getRange(1, 1, 1, plan.headers.length).setValues([plan.headers.slice()]);
    });
    SpreadsheetApp.flush();
    if (savedID !== spreadsheetID) {
      properties.setProperty("INVENTORY_SPREADSHEET_ID", spreadsheetID);
    }
  } finally {
    lock.releaseLock();
  }

  // Photo setup acquires the same script lock itself.
  setupPhotoStorage();
  const message = "Inventory setup is complete." +
    (inventoryScannerUrl_() ? "" : " Live QR scanning is optional; set INVENTORY_SCANNER_URL when you are ready.");
  console.log(message);
  return message;
}
