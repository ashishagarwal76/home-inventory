function inventoryConfiguredSpreadsheetID_() {
  const id = String(PropertiesService.getScriptProperties()
    .getProperty("INVENTORY_SPREADSHEET_ID") || "").trim();
  if (!id) {
    throw new Error(
      "Run setupInventory from the Apps Script editor attached to your inventory spreadsheet first."
    );
  }
  if (!/^[A-Za-z0-9_-]+$/.test(id)) {
    throw new Error(
      "INVENTORY_SPREADSHEET_ID must contain a spreadsheet ID. Run setupInventory from your inventory spreadsheet's Apps Script editor."
    );
  }
  return id;
}

function inventoryOptionalForm_(key) {
  const allowed = [
    "INVENTORY_ADD_FORM_ID",
    "INVENTORY_CONSUME_FORM_ID",
    "INVENTORY_MOVE_FORM_ID",
    "INVENTORY_MOVE_BOX_FORM_ID",
    "INVENTORY_NEW_BOX_FORM_ID"
  ];
  if (!allowed.includes(key)) {
    throw new Error("Unknown inventory Form setting: " + String(key));
  }
  const id = String(PropertiesService.getScriptProperties()
    .getProperty(key) || "").trim();
  if (!id) return null;
  if (!/^[A-Za-z0-9_-]+$/.test(id)) {
    throw new Error(key + " must contain a Google Form ID, not a URL.");
  }
  return FormApp.openById(id);
}

function inventoryScannerUrl_() {
  return String(PropertiesService.getScriptProperties()
    .getProperty("INVENTORY_SCANNER_URL") || "").trim();
}
