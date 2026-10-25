function submitWebNewItem(input) {
  if (!input || typeof input !== "object") {
    throw new Error("Missing New Item request.");
  }

  const requestID = String(input.requestID || "").trim();
  const uuidPattern = /^NEWITEM-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  if (!uuidPattern.test(requestID)) {
    throw new Error("Invalid New Item request reference.");
  }

  const opening = creationOpeningInput_(input);
  if (opening.error) return {accepted: false, requestID: requestID, itemID: "", message: opening.error};

  const requests = getInventorySpreadsheet_()
    .getSheetByName("NewItemRequests");
  if (!requests) throw new Error("Missing NewItemRequests sheet.");

  const headers = requests
    .getRange(1, 1, 1, requests.getLastColumn())
    .getDisplayValues()[0];

  [
    "Item Name", "Category", "Unit", "Minimum Stock", "Notes",
    "RequestID", "Status", "Message", "ProcessedAt", "CreatedItemID"
  ].forEach(header => {
    if (headers.filter(value => value === header).length !== 1) {
      throw new Error("Missing or duplicated NewItemRequests header: " + header);
    }
  });
  creationHeaders_(requests, ["Initial Quantity", "Initial StorageID", "InitialTransactionID"]);

  function findRequestRow() {
    const count = requests.getLastRow() - 1;
    if (count < 1) return 0;

    const ids = requests
      .getRange(2, headers.indexOf("RequestID") + 1, count, 1)
      .getDisplayValues();
    const matches = [];

    ids.forEach((row, index) => {
      if (row[0].trim() === requestID) matches.push(index + 2);
    });

    if (matches.length > 1) {
      throw new Error("Duplicate request reference in NewItemRequests.");
    }
    return matches[0] || 0;
  }

  function text(value) {
    return value == null ? "" : String(value).trim();
  }

  function safeText(value) {
    return /^[=+\-@']/.test(value) ? "'" + value : value;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let requestRow;

  try {
    requestRow = findRequestRow();
    if (requestRow) {
      const saved = requests.getRange(requestRow, 1, 1, headers.length).getValues()[0];
      const savedQuantity = saved[headers.indexOf("Initial Quantity")];
      const savedStorage = String(saved[headers.indexOf("Initial StorageID")] || "").trim();
      const quantity = String(savedQuantity).trim() === "" ? null : Number(savedQuantity);
      if (quantity !== opening.quantity || savedStorage !== opening.storageID) {
        return {accepted: false, requestID: requestID, itemID: "", message: "This request reference was already used for different initial-stock details."};
      }
    }

    if (!requestRow) {
      if (opening.quantity !== null) {
        const tree = inventoryStorageTree_(getInventorySpreadsheet_());
        if (!tree.boxes.has(opening.storageID) && !tree.locations.has(opening.storageID)) {
          return {accepted: false, requestID: requestID, itemID: "", message: "Select an existing placement for the initial stock."};
        }
      }
      const minimumText = text(input.minimumStock);
      const minimum = Number(minimumText);
      const values = {
        "Timestamp": new Date(),
        "Item Name": safeText(text(input.itemName)),
        "Category": safeText(text(input.category)),
        "Unit": safeText(text(input.unit)),
        "Minimum Stock": minimumText !== "" && Number.isFinite(minimum)
          ? minimum : safeText(minimumText),
        "Notes": safeText(text(input.notes)),
        "RequestID": requestID,
        "Initial Quantity": opening.quantity === null ? "" : opening.quantity,
        "Initial StorageID": opening.storageID
      };

      requests.appendRow(headers.map(header =>
        Object.prototype.hasOwnProperty.call(values, header)
          ? values[header] : ""
      ));
      SpreadsheetApp.flush();
      requestRow = findRequestRow();
      if (!requestRow) throw new Error("Could not locate the saved request.");
    }
  } finally {
    lock.releaseLock();
  }

  // The processor acquires the same lock, so call it after releasing ours.
  processNewItemRequest({
    range: requests.getRange(requestRow, 1, 1, headers.length)
  });

  requestRow = findRequestRow();
  if (!requestRow) throw new Error("The saved request could not be found.");

  const result = requests
    .getRange(requestRow, 1, 1, headers.length)
    .getDisplayValues()[0];
  const status = result[headers.indexOf("Status")].trim();

  if (!["Processed", "Rejected"].includes(status)) {
    throw new Error("The result is not confirmed. Retry the same request.");
  }

  const response = {
    accepted: status === "Processed",
    requestID: requestID,
    itemID: result[headers.indexOf("CreatedItemID")].trim(),
    message: result[headers.indexOf("Message")]
  };
  if (response.accepted) {
    const initial = recordInitialItemStock_(getInventorySpreadsheet_(), requests, requestRow, requestID, response.itemID);
    response.initialStockRecorded = initial.recorded;
    response.initialStockMessage = initial.message;
    response.message = "Item created: " + result[headers.indexOf("Item Name")] + ". " + initial.message;
    if (initial.recorded) {
      try { refreshStock(); response.stockRefreshed = true; }
      catch (error) {
        console.error(error); response.stockRefreshed = false;
        response.initialStockMessage += " Stock was saved, but its display could not refresh. Use Refresh balances or run refreshStock.";
        response.message = "Item created: " + result[headers.indexOf("Item Name")] + ". " + response.initialStockMessage;
      }
    }
  }
  return response;
}
