function initializeCreationTransactions() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getInventorySpreadsheet_();
    inventoryRequireNestedBoxes_(ss);
    const items = ss.getSheetByName("NewItemRequests");
    const boxes = ss.getSheetByName("NewBoxRequests");
    if (!items || !boxes) throw new Error("Missing NewItemRequests or NewBoxRequests sheet.");
    inventoryBoxAddHeaders_(items, ["Initial Quantity", "Initial StorageID", "InitialTransactionID"]);
    inventoryBoxAddHeaders_(boxes, ["CreatedPlacementID", "InitialPlacementLocation", "InitialPlacementDescription", "InitialPlacementAt"]);
    SpreadsheetApp.flush();
    return "Creation transactions are ready.";
  } finally { lock.releaseLock(); }
}

function getNewItemPlacementOptions() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return { placements: inventoryStorageTree_(getInventorySpreadsheet_()).choices() }; }
  finally { lock.releaseLock(); }
}

function creationHeaders_(sheet, required) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  required.forEach(name => {
    if (headers.filter(value => value === name).length !== 1) {
      throw new Error("Run initializeCreationTransactions once in Apps Script. Missing or duplicated " + sheet.getName() + " header: " + name);
    }
  });
  return headers;
}

function creationOpeningInput_(input) {
  const quantityText = input.initialQuantity == null ? "" : String(input.initialQuantity).trim();
  const storageID = String(input.initialStorageID || "").trim();
  if (!quantityText && !storageID) return {quantity: null, storageID: "", error: ""};
  if (!quantityText || !storageID) return {error: "For initial stock, enter both a quantity and a placement, or leave both blank."};
  if (!["string", "number"].includes(typeof input.initialQuantity)) return {error: "Initial quantity must be a number greater than zero."};
  const quantity = Number(quantityText);
  const error = creationQuantityError_(quantity, String(input.unit || ""));
  return {quantity: quantity, storageID: storageID, error: error};
}

function creationQuantityError_(quantity, unit) {
  if (!Number.isFinite(quantity) || quantity <= 0 || quantity > Number.MAX_SAFE_INTEGER) return "Initial quantity must be a number greater than zero.";
  if (["each", "ea", "pack", "packs", "piece", "pieces", "pcs"].includes(unit.trim().toLowerCase())) {
    return Number.isSafeInteger(quantity) ? "" : "This unit requires a whole-number initial quantity.";
  }
  const scaled = quantity * 1000;
  const tolerance = Number.EPSILON * Math.max(1, Math.abs(scaled)) * 16;
  return quantity < 0.001 || !Number.isSafeInteger(Math.round(scaled)) || Math.abs(scaled - Math.round(scaled)) > tolerance
    ? "Enter an initial quantity of at least 0.001, with up to 3 decimal places." : "";
}

function recordInitialItemStock_(ss, requests, requestRow, requestID, itemID) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const headers = creationHeaders_(requests, ["Initial Quantity", "Initial StorageID", "InitialTransactionID"]);
    const row = requests.getRange(requestRow, 1, 1, headers.length).getValues()[0];
    const value = name => row[headers.indexOf(name)];
    if (String(value("RequestID")).trim() !== requestID || String(value("CreatedItemID")).trim() !== itemID ||
        String(value("Status")).trim() !== "Processed") throw new Error("The saved item request is not confirmed. Retry the same request.");
    const quantityValue = value("Initial Quantity");
    const storageID = String(value("Initial StorageID") || "").trim();
    if (String(quantityValue).trim() === "" && !storageID) return {recorded: false, message: "No initial stock recorded."};
    const quantity = Number(quantityValue);
    const transactionID = "INITIAL-" + requestID;
    const transactions = ss.getSheetByName("Transactions");
    if (!transactions) throw new Error("Missing Transactions sheet.");
    inventoryBoxHeaders_(transactions, ["TransactionID", "Timestamp", "Action", "ItemID", "Quantity", "FromStorageID", "ToStorageID", "Notes"]);
    const saved = transactions.getDataRange().getValues().slice(1).filter(r => String(r[0]).trim() === transactionID);
    if (saved.length > 1) throw new Error("Duplicate initial-stock transaction reference.");
    if (saved.length) {
      const r = saved[0];
      if (r[2] !== "Add" || String(r[3]).trim() !== itemID || Number(r[4]) !== quantity ||
          String(r[5] || "").trim() !== "" || String(r[6] || "").trim() !== storageID) throw new Error("The saved initial-stock transaction has different details.");
      requests.getRange(requestRow, headers.indexOf("InitialTransactionID") + 1).setValue(transactionID);
      return {recorded: true, transactionID: transactionID, message: "Initial stock was already recorded. No duplicate added."};
    }
    const matches = ss.getSheetByName("Items").getDataRange().getValues().slice(1).filter(r => String(r[0]).trim() === itemID);
    if (matches.length !== 1) throw new Error("The created item is missing or duplicated.");
    const unit = String(matches[0][3] || "").trim();
    if (unit.toLowerCase() !== String(value("Unit") || "").trim().toLowerCase()) throw new Error("The item's unit changed before initial stock was recorded.");
    const error = creationQuantityError_(quantity, unit);
    if (error || !storageID) throw new Error(error || "The initial placement is missing.");
    const tree = inventoryStorageTree_(ss);
    if (!tree.boxes.has(storageID) && !tree.locations.has(storageID)) throw new Error("The initial placement no longer exists. Restore it and retry this saved request.");
    const placement = tree.resolve(storageID).path;
    transactions.appendRow([transactionID, new Date(), "Add", itemID, quantity, "", storageID, "Initial stock recorded during item creation."]);
    SpreadsheetApp.flush();
    requests.getRange(requestRow, headers.indexOf("InitialTransactionID") + 1).setValue(transactionID);
    SpreadsheetApp.flush();
    return {recorded: true, transactionID: transactionID, message: "Initial stock: " + quantity + " " + unit + " in " + placement + "."};
  } finally { lock.releaseLock(); }
}

function recordInitialBoxPlacement_(ss, requests, requestRow, boxID) {
  // Called by processNewBoxRequest while it holds the script lock.
  const headers = creationHeaders_(requests, ["CreatedPlacementID", "InitialPlacementLocation", "InitialPlacementDescription", "InitialPlacementAt"]);
  const row = requests.getRange(requestRow, 1, 1, headers.length).getValues()[0];
  const value = name => row[headers.indexOf(name)];
  const id = String(value("CreatedPlacementID") || "").trim();
  if (!id) return; // Requests created before this update have no creation event to replay.
  const destination = String(value("Location") || "").split(" — ")[0].trim();
  const history = inventoryBoxMoveHistory_(ss);
  const saved = history.entries.find(entry => entry.id === id);
  if (saved) {
    if (saved.boxID !== boxID || saved.toStorage !== destination || saved.status !== "Processed") throw new Error("The initial box placement has different saved details.");
    return;
  }
  const values = {
    MovementID: id, Timestamp: value("InitialPlacementAt") || new Date(), BoxID: boxID,
    FromLocationID: "", ToLocationID: value("InitialPlacementLocation"),
    Notes: "Box created and placed during creation.", FromStorageID: "", ToStorageID: destination,
    Status: "Processed", Message: "Box created in " + String(value("InitialPlacementDescription") || "the selected placement") + ".", ProcessedAt: new Date()
  };
  history.sheet.appendRow(history.headers.map(name => inventoryBoxSafeText_(values[name] == null ? "" : values[name])));
  SpreadsheetApp.flush();
}
