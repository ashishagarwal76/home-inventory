function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  // Run only when an edit includes column B on the Items tab.
  if (
    sheet.getName() !== "Items" ||
    e.range.getColumn() > 2 ||
    e.range.getLastColumn() < 2 ||
    e.range.getLastRow() < 2
  ) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const properties = PropertiesService.getScriptProperties();
    let lastNumber = Number(
      properties.getProperty("LAST_ITEM_NUMBER") || 0
    );

    // Account for existing IDs, including older IDs such as ITEM-001.
    const existingIDs = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getDisplayValues();

    for (const [id] of existingIDs) {
      const match = id.trim().match(/^ITEM-(\d+)$/);
      if (match) {
        lastNumber = Math.max(lastNumber, Number(match[1]));
      }
    }

    const firstRow = Math.max(2, e.range.getRow());
    const lastRow = e.range.getLastRow();

    for (let row = firstRow; row <= lastRow; row++) {
      const idCell = sheet.getRange(row, 1);
      const itemName = sheet.getRange(row, 2).getDisplayValue().trim();

      if (itemName && idCell.isBlank()) {
        lastNumber++;

        // Reserve the number before writing it to the sheet.
        properties.setProperty(
          "LAST_ITEM_NUMBER",
          String(lastNumber)
        );

        idCell.setValue(
          "ITEM-" + String(lastNumber).padStart(5, "0")
        );
      }
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function refreshAddInventoryForm() {
  const form = inventoryOptionalForm_("INVENTORY_ADD_FORM_ID");
  if (!form) return;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function readChoices(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error("Missing sheet: " + sheetName);
    }
    if (sheet.getLastRow() < 2) return [];

    const rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getDisplayValues();

    const seen = new Set();

    return rows
      .filter(([id, name]) => id.trim() && name.trim())
      .map(([id, name]) => {
        id = id.trim();

        if (seen.has(id)) {
          throw new Error("Duplicate ID in " + sheetName + ": " + id);
        }
        seen.add(id);

        return id + " — " + name.trim();
      });
  }

  const items = readChoices("Items");
  const destinations = [
    ...readChoices("Boxes"),
    ...readChoices("Locations")
  ];

  if (!items.length || !destinations.length) {
    throw new Error(
      "Add at least one item and one box or location first."
    );
  }

  function findDropdown(title) {
    const matches = form.getItems(FormApp.ItemType.LIST)
      .filter(question => question.getTitle().trim() === title);

    if (matches.length !== 1) {
      throw new Error(
        'The form needs exactly one dropdown named "' + title + '".'
      );
    }
    return matches[0].asListItem();
  }

  const itemDropdown = findDropdown("Item");
  const destinationDropdown = findDropdown("Destination");

  itemDropdown.setChoiceValues(items);
  destinationDropdown.setChoiceValues(destinations);

  console.log("Item and Destination dropdowns updated.");
}


function processAddRequest(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const requests = e.range.getSheet();
  if (requests.getName() !== "AddRequests") return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = requests.getParent();
    const row = e.range.getRow();
    const headers = requests
      .getRange(1, 1, 1, requests.getLastColumn())
      .getDisplayValues()[0];

    function cell(header) {
      const column = headers.indexOf(header);
      if (column === -1) {
        throw new Error("Missing AddRequests header: " + header);
      }
      return requests.getRange(row, column + 1);
    }

    // Check required columns before processing.
    [
      "Item", "Quantity", "Destination", "Notes",
      "RequestID", "Status", "Message", "ProcessedAt"
    ].forEach(cell);

    if (["Processed", "Rejected"].includes(
      cell("Status").getDisplayValue()
    )) return;

    let requestID = cell("RequestID").getDisplayValue();
    if (!requestID) {
      requestID = Utilities.getUuid();
      cell("RequestID").setValue(requestID);
      SpreadsheetApp.flush();
    }

    function finish(status, message) {
      cell("Message").setValue(message);
      cell("ProcessedAt").setValue(new Date());
      cell("Status").setValue(status);
      SpreadsheetApp.flush();
    }

    function readRows(sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) throw new Error("Missing sheet: " + sheetName);
      return sheet.getDataRange().getValues().slice(1);
    }

    const transactions = spreadsheet.getSheetByName("Transactions");
    if (!transactions) throw new Error("Missing Transactions sheet.");

    // A retry must not record the same request twice.
    if (readRows("Transactions").some(r => r[0] === requestID)) {
      finish("Processed", "Inventory already added.");
      return;
    }

    // Dropdown labels contain ID followed by an em dash and name.
    const getID = value => String(value).split(" — ")[0].trim();

    const itemID = getID(cell("Item").getDisplayValue());
    const destination = getID(cell("Destination").getDisplayValue());
    const quantity = Number(cell("Quantity").getValue());

    const items = readRows("Items");
    const boxes = readRows("Boxes");
    const locations = readRows("Locations");
    const matchingItems = items.filter(r => String(r[0]).trim() === itemID);
    const matchingBoxes = boxes.filter(r => String(r[0]).trim() === destination);
    const matchingLocations = locations.filter(
      r => String(r[0]).trim() === destination
    );

    let error = "";

    if (matchingItems.length !== 1) {
      error = "Item ID is missing or duplicated in Items.";
    } else if (!Number.isFinite(quantity) || quantity <= 0) {
      error = "Quantity must be a number greater than zero.";
    } else if (matchingBoxes.length + matchingLocations.length !== 1) {
      error = "Destination is missing or duplicated.";
    } else if (
      matchingBoxes.length === 1 &&
      locations.filter(r =>
        String(r[0]).trim() === String(matchingBoxes[0][2]).trim()
      ).length !== 1
    ) {
      error = "The selected box needs a valid location.";
    }

    if (error) {
      finish("Rejected", error);
      return;
    }

    let notes = cell("Notes").getDisplayValue();
    // Store formula-like notes as text.
    if (/^[=+\-@]/.test(notes)) notes = "'" + notes;

    transactions.appendRow([
      requestID,
      new Date(),
      "Add",
      itemID,
      quantity,
      "",
      destination,
      notes
    ]);

    SpreadsheetApp.flush();
    finish("Processed", "Inventory added successfully.");
  } finally {
    lock.releaseLock();
  }
}

//refreshStock
function refreshStock() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getInventorySpreadsheet_();

    function rows(name) {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error("Missing sheet: " + name);
      return sheet.getDataRange().getValues().slice(1)
        .filter(row => row[0] !== "");
    }

    const items = new Map(
      rows("Items").map(r => [String(r[0]).trim(), r])
    );
    const storage = inventoryStorageTree_(ss);

    const balances = new Map();
    function change(itemID, storageID, quantity) {
      if (!storageID) return;
      const key = JSON.stringify([itemID, storageID]);
      const next = (balances.get(key) || 0) + quantity;
      if (!Number.isFinite(next)) throw new Error("The inventory balance exceeds the supported range.");
      balances.set(key, next);
    }

    for (const transaction of rows("Transactions")) {
      const itemID = String(transaction[3]).trim();
      const quantity = Number(transaction[4]);
      const source = String(transaction[5] || "").trim();
      const destination = String(transaction[6] || "").trim();
      if (!Number.isFinite(quantity) || quantity <= 0) {
        throw new Error("Invalid quantity in transaction: " + transaction[0]);
      }
      change(itemID, source, -quantity);
      change(itemID, destination, quantity);
    }

    const output = [];
    for (const [key, balance] of balances) {
      const quantity = Math.round(balance * 1000000) / 1000000;
      if (quantity === 0) continue;
      const [itemID, storageID] = JSON.parse(key);
      const item = items.get(itemID);
      if (!item) throw new Error("Missing item for: " + itemID + " / " + storageID);

      // Member storage is still owned inventory, held outside a box/location.
      if (storageID.startsWith("MEMBER:")) {
        const email = storageID.slice(7);
        if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
            email !== email.trim().toLowerCase()) {
          throw new Error("Invalid family-member storage ID: " + storageID);
        }
        output.push([
          itemID, item[1], storageID, "", "With " + email, quantity, item[3]
        ]);
        continue;
      }

      const box = storage.boxes.get(storageID);
      const place = storage.resolve(storageID);
      output.push([
        itemID, item[1], storageID, box ? box.name : "",
        place.locationName, quantity, item[3]
      ]);
    }

    output.sort((a, b) =>
      String(a[0]).localeCompare(String(b[0])) ||
      String(a[2]).localeCompare(String(b[2]))
    );
    const safeOutput = output.map(row => row.map(value =>
      typeof value === "string" && /^[=+\-@']/.test(value) ? "'" + value : value
    ));

    const stockSheet = ss.getSheetByName("Stock");
    if (!stockSheet) throw new Error("Missing Stock sheet.");
    if (stockSheet.getLastRow() > 1) {
      stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 7).clearContent();
    }
    if (safeOutput.length) {
      if (stockSheet.getMaxRows() < safeOutput.length + 1) {
        stockSheet.insertRowsAfter(
          stockSheet.getMaxRows(),
          safeOutput.length + 1 - stockSheet.getMaxRows()
        );
      }
      stockSheet.getRange(2, 1, safeOutput.length, 7).setValues(safeOutput);
      stockSheet.getRange(2, 6, safeOutput.length, 1).setNumberFormat("0.######");
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function handleInventorySubmission(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const sheetName = e.range.getSheet().getName();

  if (sheetName === "AddRequests") {
    processAddRequest(e);
  } else if (sheetName === "ConsumeRequests") {
    processConsumeRequest(e);
  } else if (sheetName === "MoveRequests") {
    processMoveRequest(e);
  } else if (sheetName === "BoxMoveRequests") {
    processBoxMoveRequest(e);
  } else if (sheetName === "NewItemRequests") {
    processNewItemRequest(e);
    refreshAllInventoryForms();
  } else if (sheetName === "NewLocationRequests") {
    processNewLocationRequest(e);
    refreshAllInventoryForms();
  } else if (sheetName === "NewBoxRequests") {
    processNewBoxRequest(e);
    refreshAllInventoryForms();
  } else {
    return;
  }

  refreshStock();
}


function refreshConsumeInventoryForm() {
  const form = inventoryOptionalForm_("INVENTORY_CONSUME_FORM_ID");
  if (!form) return;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function readChoices(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error("Missing sheet: " + sheetName);
    }
    if (sheet.getLastRow() < 2) return [];

    const rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getDisplayValues();

    const seen = new Set();

    return rows
      .filter(([id, name]) => id.trim() && name.trim())
      .map(([id, name]) => {
        id = id.trim();

        if (seen.has(id)) {
          throw new Error("Duplicate ID in " + sheetName + ": " + id);
        }
        seen.add(id);

        return id + " — " + name.trim();
      });
  }

  const items = readChoices("Items");
  const sources = [
    ...readChoices("Boxes"),
    ...readChoices("Locations")
  ];

  if (!items.length || !sources.length) {
    throw new Error(
      "Add at least one item and one box or location first."
    );
  }

  function findDropdown(title) {
    const matches = form.getItems(FormApp.ItemType.LIST)
      .filter(question => question.getTitle().trim() === title);

    if (matches.length !== 1) {
      throw new Error(
        'The form needs exactly one dropdown named "' + title + '".'
      );
    }
    return matches[0].asListItem();
  }

  const itemDropdown = findDropdown("Item");
  const sourceDropdown = findDropdown("Source");

  itemDropdown.setChoiceValues(items);
  sourceDropdown.setChoiceValues(sources);

  console.log("Consume form dropdowns updated.");
}

function processConsumeRequest(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const requests = e.range.getSheet();
  if (requests.getName() !== "ConsumeRequests") return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = requests.getParent();
    const row = e.range.getRow();
    const headers = requests
      .getRange(1, 1, 1, requests.getLastColumn())
      .getDisplayValues()[0];

    function cell(header) {
      const column = headers.indexOf(header);
      if (column === -1) {
        throw new Error("Missing ConsumeRequests header: " + header);
      }
      return requests.getRange(row, column + 1);
    }

    [
      "Item", "Quantity", "Source", "Notes",
      "RequestID", "Status", "Message", "ProcessedAt"
    ].forEach(cell);

    if (["Processed", "Rejected"].includes(
      cell("Status").getDisplayValue()
    )) return;

    let requestID = cell("RequestID").getDisplayValue();
    if (!requestID) {
      requestID = Utilities.getUuid();
      cell("RequestID").setValue(requestID);
      SpreadsheetApp.flush();
    }

    function finish(status, message) {
      cell("Message").setValue(message);
      cell("ProcessedAt").setValue(new Date());
      cell("Status").setValue(status);
      SpreadsheetApp.flush();
    }

    function readRows(sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) throw new Error("Missing sheet: " + sheetName);
      return sheet.getDataRange().getValues().slice(1);
    }

    const transactions = spreadsheet.getSheetByName("Transactions");
    if (!transactions) throw new Error("Missing Transactions sheet.");

    const history = readRows("Transactions");

    // Avoid posting the same request twice if processing is retried.
    if (history.some(r => r[0] === requestID)) {
      finish("Processed", "Consumption already recorded.");
      return;
    }

    const getID = value => String(value).split(" — ")[0].trim();

    const itemID = getID(cell("Item").getDisplayValue());
    const source = getID(cell("Source").getDisplayValue());
    const quantity = Number(cell("Quantity").getValue());

    const items = readRows("Items");
    const boxes = readRows("Boxes");
    const locations = readRows("Locations");

    const matchingItems = items.filter(
      r => String(r[0]).trim() === itemID
    );
    const matchingBoxes = boxes.filter(
      r => String(r[0]).trim() === source
    );
    const matchingLocations = locations.filter(
      r => String(r[0]).trim() === source
    );

    let error = "";

    if (matchingItems.length !== 1) {
      error = "Item ID is missing or duplicated in Items.";
    } else if (!Number.isFinite(quantity) || quantity <= 0) {
      error = "Quantity must be a number greater than zero.";
    } else if (matchingBoxes.length + matchingLocations.length !== 1) {
      error = "Source is missing or duplicated.";
    } else if (
      matchingBoxes.length === 1 &&
      locations.filter(r =>
        String(r[0]).trim() === String(matchingBoxes[0][2]).trim()
      ).length !== 1
    ) {
      error = "The selected box needs a valid location.";
    }

    if (error) {
      finish("Rejected", error);
      return;
    }

    // Calculate directly from Transactions, not the Stock snapshot.
    let available = 0;

    for (const transaction of history) {
      if (String(transaction[3]).trim() !== itemID) continue;

      const from = String(transaction[5] || "").trim();
      const to = String(transaction[6] || "").trim();

      if (from !== source && to !== source) continue;

      const amount = Number(transaction[4]);
      if (!Number.isFinite(amount) || amount <= 0) {
        throw new Error(
          "Invalid quantity in transaction: " + transaction[0]
        );
      }

      if (to === source) available += amount;
      if (from === source) available -= amount;
    }

    // Allow only a tiny floating-point rounding difference.
    const tolerance = Number.EPSILON *
      Math.max(1, Math.abs(available), quantity) * 16;

    if (available <= 0 || quantity > available + tolerance) {
      finish(
        "Rejected",
        "Insufficient stock. Available: " + available +
        "; requested: " + quantity + "."
      );
      return;
    }

    let notes = cell("Notes").getDisplayValue();
    if (/^[=+\-@]/.test(notes)) notes = "'" + notes;

    transactions.appendRow([
      requestID,
      new Date(),
      "Consume",
      itemID,
      quantity,
      source,
      "",
      notes
    ]);

    SpreadsheetApp.flush();
    finish("Processed", "Consumption recorded successfully.");
  } finally {
    lock.releaseLock();
  }
}

function refreshMoveInventoryForm() {
  const form = inventoryOptionalForm_("INVENTORY_MOVE_FORM_ID");
  if (!form) return;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function readChoices(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error("Missing sheet: " + sheetName);
    }
    if (sheet.getLastRow() < 2) return [];

    const rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getDisplayValues();

    const seen = new Set();

    return rows
      .filter(([id, name]) => id.trim() && name.trim())
      .map(([id, name]) => {
        id = id.trim();

        if (seen.has(id)) {
          throw new Error("Duplicate ID in " + sheetName + ": " + id);
        }
        seen.add(id);

        return id + " — " + name.trim();
      });
  }

  const items = readChoices("Items");
  const storage = [
    ...readChoices("Boxes"),
    ...readChoices("Locations")
  ];

  if (!items.length || !storage.length) {
    throw new Error(
      "Add at least one item and one box or location first."
    );
  }

  function findDropdown(title) {
    const matches = form.getItems(FormApp.ItemType.LIST)
      .filter(question => question.getTitle().trim() === title);

    if (matches.length !== 1) {
      throw new Error(
        'The form needs exactly one dropdown named "' + title + '".'
      );
    }
    return matches[0].asListItem();
  }

  const itemDropdown = findDropdown("Item");
  const sourceDropdown = findDropdown("Source");
  const destinationDropdown = findDropdown("Destination");

  itemDropdown.setChoiceValues(items);
  sourceDropdown.setChoiceValues(storage);
  destinationDropdown.setChoiceValues(storage);

  console.log("Move form dropdowns updated.");
}

function processMoveRequest(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const requests = e.range.getSheet();
  if (requests.getName() !== "MoveRequests") return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = requests.getParent();
    const row = e.range.getRow();
    const headers = requests
      .getRange(1, 1, 1, requests.getLastColumn())
      .getDisplayValues()[0];

    function cell(header) {
      const column = headers.indexOf(header);
      if (column === -1) {
        throw new Error("Missing MoveRequests header: " + header);
      }
      return requests.getRange(row, column + 1);
    }

    [
      "Item", "Quantity", "Source", "Destination", "Notes",
      "RequestID", "Status", "Message", "ProcessedAt"
    ].forEach(cell);

    if (["Processed", "Rejected"].includes(
      cell("Status").getDisplayValue()
    )) return;

    let requestID = cell("RequestID").getDisplayValue();
    if (!requestID) {
      requestID = Utilities.getUuid();
      cell("RequestID").setValue(requestID);
      SpreadsheetApp.flush();
    }

    function finish(status, message) {
      cell("Message").setValue(message);
      cell("ProcessedAt").setValue(new Date());
      cell("Status").setValue(status);
      SpreadsheetApp.flush();
    }

    function readRows(sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) throw new Error("Missing sheet: " + sheetName);
      return sheet.getDataRange().getValues().slice(1);
    }

    const transactions = spreadsheet.getSheetByName("Transactions");
    if (!transactions) throw new Error("Missing Transactions sheet.");

    const history = readRows("Transactions");

    if (history.some(r => r[0] === requestID)) {
      finish("Processed", "Movement already recorded.");
      return;
    }

    const getID = value => String(value).split(" — ")[0].trim();

    const itemID = getID(cell("Item").getDisplayValue());
    const source = getID(cell("Source").getDisplayValue());
    const destination = getID(cell("Destination").getDisplayValue());
    const quantity = Number(cell("Quantity").getValue());

    const items = readRows("Items");
    const boxes = readRows("Boxes");
    const locations = readRows("Locations");

    function validStorage(id) {
      const matchingBoxes = boxes.filter(
        r => String(r[0]).trim() === id
      );
      const matchingLocations = locations.filter(
        r => String(r[0]).trim() === id
      );

      if (matchingBoxes.length + matchingLocations.length !== 1) {
        return false;
      }

      if (matchingBoxes.length === 1) {
        const locationID = String(matchingBoxes[0][2]).trim();
        return locations.filter(
          r => String(r[0]).trim() === locationID
        ).length === 1;
      }

      return true;
    }

    let error = "";

    if (items.filter(
      r => String(r[0]).trim() === itemID
    ).length !== 1) {
      error = "Item ID is missing or duplicated in Items.";
    } else if (!Number.isFinite(quantity) || quantity <= 0) {
      error = "Quantity must be a number greater than zero.";
    } else if (!validStorage(source)) {
      error = "Source is invalid or its box needs a valid location.";
    } else if (!validStorage(destination)) {
      error = "Destination is invalid or its box needs a valid location.";
    } else if (source === destination) {
      error = "Source and destination must be different.";
    }

    if (error) {
      finish("Rejected", error);
      return;
    }

    // Calculate available stock directly from Transactions.
    let available = 0;

    for (const transaction of history) {
      if (String(transaction[3]).trim() !== itemID) continue;

      const from = String(transaction[5] || "").trim();
      const to = String(transaction[6] || "").trim();

      if (from !== source && to !== source) continue;

      const amount = Number(transaction[4]);
      if (!Number.isFinite(amount) || amount <= 0) {
        throw new Error(
          "Invalid quantity in transaction: " + transaction[0]
        );
      }

      if (to === source) available += amount;
      if (from === source) available -= amount;
    }

    const tolerance = Number.EPSILON *
      Math.max(1, Math.abs(available), quantity) * 16;

    if (available <= 0 || quantity > available + tolerance) {
      finish(
        "Rejected",
        "Insufficient stock. Available: " + available +
        "; requested: " + quantity + "."
      );
      return;
    }

    let notes = cell("Notes").getDisplayValue();
    if (/^[=+\-@]/.test(notes)) notes = "'" + notes;

    // One row records both sides of the transfer.
    transactions.appendRow([
      requestID,
      new Date(),
      "Move",
      itemID,
      quantity,
      source,
      destination,
      notes
    ]);

    SpreadsheetApp.flush();
    finish("Processed", "Movement recorded successfully.");
  } finally {
    lock.releaseLock();
  }
}

function refreshMoveBoxForm() {
  const form = inventoryOptionalForm_("INVENTORY_MOVE_BOX_FORM_ID");
  if (!form) return;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function readChoices(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error("Missing sheet: " + sheetName);
    }
    if (sheet.getLastRow() < 2) return [];

    const rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getDisplayValues();

    const seen = new Set();

    return rows
      .filter(([id, name]) => id.trim() && name.trim())
      .map(([id, name]) => {
        id = id.trim();

        if (seen.has(id)) {
          throw new Error("Duplicate ID in " + sheetName + ": " + id);
        }
        seen.add(id);

        return id + " — " + name.trim();
      });
  }

  const boxes = readChoices("Boxes");
  const locations = readChoices("Locations");

  if (!boxes.length || !locations.length) {
    throw new Error("Add at least one box and one location first.");
  }

  function findDropdown(title) {
    const matches = form.getItems(FormApp.ItemType.LIST)
      .filter(question => question.getTitle().trim() === title);

    if (matches.length !== 1) {
      throw new Error(
        'The form needs exactly one dropdown named "' + title + '".'
      );
    }
    return matches[0].asListItem();
  }

  const boxDropdown = findDropdown("Box");
  const locationDropdown = findDropdown("Destination Location");

  boxDropdown.setChoiceValues(boxes);
  locationDropdown.setChoiceValues(locations);

  console.log("Move Box dropdowns updated.");
}

function processBoxMoveRequest(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const requests = e.range.getSheet();
  if (requests.getName() !== "BoxMoveRequests") return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = requests.getParent();
    const row = e.range.getRow();
    const headers = requests
      .getRange(1, 1, 1, requests.getLastColumn())
      .getDisplayValues()[0];

    function cell(header) {
      const column = headers.indexOf(header);
      if (column === -1) {
        throw new Error("Missing BoxMoveRequests header: " + header);
      }
      return requests.getRange(row, column + 1);
    }

    [
      "Box", "Destination Location", "Notes",
      "RequestID", "Status", "Message", "ProcessedAt"
    ].forEach(cell);

    if (["Processed", "Rejected"].includes(
      cell("Status").getDisplayValue()
    )) return;

    let requestID = cell("RequestID").getDisplayValue();
    if (!requestID) {
      requestID = Utilities.getUuid();
      cell("RequestID").setValue(requestID);
      SpreadsheetApp.flush();
    }

    function finish(status, message) {
      cell("Message").setValue(message);
      cell("ProcessedAt").setValue(new Date());
      cell("Status").setValue(status);
      SpreadsheetApp.flush();
    }

    const boxesSheet = spreadsheet.getSheetByName("Boxes");
    const locationsSheet = spreadsheet.getSheetByName("Locations");
    const historySheet = spreadsheet.getSheetByName("BoxMoves");

    if (!boxesSheet || !locationsSheet || !historySheet) {
      throw new Error("Missing Boxes, Locations, or BoxMoves sheet.");
    }

    const boxes = boxesSheet.getDataRange().getValues();
    const locations = locationsSheet.getDataRange().getValues().slice(1);
    const history = historySheet.getDataRange().getValues().slice(1);

    const getID = value => String(value).split(" — ")[0].trim();
    const existingMove = history.find(r => r[0] === requestID);

    // If retrying, use the box from the recorded movement.
    const boxID = existingMove
      ? String(existingMove[2]).trim()
      : getID(cell("Box").getDisplayValue());

    const destination = getID(
      cell("Destination Location").getDisplayValue()
    );

    const matchingRows = [];
    for (let index = 1; index < boxes.length; index++) {
      if (String(boxes[index][0]).trim() === boxID) {
        matchingRows.push(index + 1);
      }
    }

    if (matchingRows.length !== 1) {
      if (existingMove) {
        throw new Error("Recorded movement needs a unique box: " + boxID);
      }
      finish("Rejected", "Box ID is missing or duplicated in Boxes.");
      return;
    }

    const locationCell = boxesSheet.getRange(matchingRows[0], 3);

    // Recover the latest recorded location if an earlier update stopped.
    const boxHistory = history.filter(
      r => String(r[2]).trim() === boxID
    );

    if (boxHistory.length) {
      const latest = boxHistory[boxHistory.length - 1];
      locationCell.setValue(String(latest[4]).trim());
      SpreadsheetApp.flush();
    }

    if (existingMove) {
      finish("Processed", "Box movement already recorded.");
      return;
    }

    const destinationMatches = locations.filter(
      r => String(r[0]).trim() === destination
    );

    if (destinationMatches.length !== 1) {
      finish("Rejected", "Destination location is missing or duplicated.");
      return;
    }

    const source = locationCell.getDisplayValue().trim();

    if (source === destination) {
      finish("Rejected", "The box is already at that location.");
      return;
    }

    let notes = cell("Notes").getDisplayValue();
    if (/^[=+\-@]/.test(notes)) notes = "'" + notes;

    // Record history first so a retry can recover an interrupted update.
    historySheet.appendRow([
      requestID,
      new Date(),
      boxID,
      source,
      destination,
      notes
    ]);
    SpreadsheetApp.flush();

    locationCell.setValue(destination);
    SpreadsheetApp.flush();

    finish("Processed", "Box moved successfully.");
  } finally {
    lock.releaseLock();
  }
}

function refreshAllInventoryForms() {
  refreshAddInventoryForm();
  refreshConsumeInventoryForm();
  refreshMoveInventoryForm();
  refreshMoveBoxForm();
  refreshNewBoxForm();

  console.log("All inventory form dropdowns updated.");
}


function processNewItemRequest(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const requests = e.range.getSheet();
  if (requests.getName() !== "NewItemRequests") return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = requests.getParent();
    const row = e.range.getRow();
    const headers = requests
      .getRange(1, 1, 1, requests.getLastColumn())
      .getDisplayValues()[0];

    function cell(header) {
      const column = headers.indexOf(header);
      if (column === -1) {
        throw new Error("Missing NewItemRequests header: " + header);
      }
      return requests.getRange(row, column + 1);
    }

    [
      "Item Name", "Category", "Unit", "Minimum Stock", "Notes",
      "RequestID", "Status", "Message", "ProcessedAt", "CreatedItemID"
    ].forEach(cell);

    if (["Processed", "Rejected"].includes(
      cell("Status").getDisplayValue()
    )) return;

    if (!cell("RequestID").getDisplayValue()) {
      cell("RequestID").setValue(Utilities.getUuid());
      SpreadsheetApp.flush();
    }

    function finish(status, message) {
      cell("Message").setValue(message);
      cell("ProcessedAt").setValue(new Date());
      cell("Status").setValue(status);
      SpreadsheetApp.flush();
    }

    const itemsSheet = spreadsheet.getSheetByName("Items");
    if (!itemsSheet) throw new Error("Missing Items sheet.");

    const items = itemsSheet.getDataRange().getValues().slice(1);
    let itemID = cell("CreatedItemID").getDisplayValue().trim();

    // Recover a request whose item was saved before processing stopped.
    if (itemID && items.some(r => String(r[0]).trim() === itemID)) {
      finish("Processed", "Item already created: " + itemID);
      return;
    }

    const name = cell("Item Name").getDisplayValue().trim();
    const category = cell("Category").getDisplayValue().trim();
    const unit = cell("Unit").getDisplayValue().trim();
    const minimumText = cell("Minimum Stock").getDisplayValue().trim();
    const minimum = Number(cell("Minimum Stock").getValue());
    const notes = cell("Notes").getDisplayValue().trim();

    if (!name || !unit) {
      finish("Rejected", "Item Name and Unit are required.");
      return;
    }

    if (
      minimumText === "" ||
      !Number.isFinite(minimum) ||
      minimum < 0
    ) {
      finish("Rejected", "Minimum Stock must be zero or greater.");
      return;
    }

    const properties = PropertiesService.getScriptProperties();
    let lastNumber = Number(
      properties.getProperty("LAST_ITEM_NUMBER") || 0
    );

    if (!Number.isSafeInteger(lastNumber) || lastNumber < 0) {
      throw new Error("The item number counter is invalid.");
    }

    for (const item of items) {
      const match = String(item[0]).trim().match(/^ITEM-(\d+)$/);
      if (match) {
        lastNumber = Math.max(lastNumber, Number(match[1]));
      }
    }

    if (!itemID) {
      lastNumber++;

      // Share the counter used by manual ItemName entry.
      properties.setProperty("LAST_ITEM_NUMBER", String(lastNumber));

      itemID = "ITEM-" + String(lastNumber).padStart(5, "0");
      cell("CreatedItemID").setValue(itemID);
      SpreadsheetApp.flush();
    } else {
      const match = itemID.match(/^ITEM-(\d+)$/);
      if (!match) throw new Error("Invalid reserved ItemID.");

      properties.setProperty(
        "LAST_ITEM_NUMBER",
        String(Math.max(lastNumber, Number(match[1])))
      );
    }

    function safeText(value) {
      return /^[=+\-@]/.test(value) ? "'" + value : value;
    }

    itemsSheet.appendRow([
      itemID,
      safeText(name),
      safeText(category),
      safeText(unit),
      minimum,
      safeText(notes)
    ]);

    SpreadsheetApp.flush();
    finish("Processed", "Item created: " + itemID);
  } finally {
    lock.releaseLock();
  }
}

function processNewLocationRequest(e) {
  if (!e || !e.range) {
    throw new Error("This function runs from a form submission.");
  }

  const requests = e.range.getSheet();
  if (requests.getName() !== "NewLocationRequests") return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = requests.getParent();
    const row = e.range.getRow();
    const headers = requests
      .getRange(1, 1, 1, requests.getLastColumn())
      .getDisplayValues()[0];

    function cell(header) {
      const column = headers.indexOf(header);
      if (column === -1) {
        throw new Error(
          "Missing NewLocationRequests header: " + header
        );
      }
      return requests.getRange(row, column + 1);
    }

    [
      "Location Name", "Notes", "RequestID",
      "Status", "Message", "ProcessedAt", "CreatedLocationID"
    ].forEach(cell);

    if (["Processed", "Rejected"].includes(
      cell("Status").getDisplayValue()
    )) return;

    if (!cell("RequestID").getDisplayValue()) {
      cell("RequestID").setValue(Utilities.getUuid());
      SpreadsheetApp.flush();
    }

    function finish(status, message) {
      cell("Message").setValue(message);
      cell("ProcessedAt").setValue(new Date());
      cell("Status").setValue(status);
      SpreadsheetApp.flush();
    }

    const locationsSheet = spreadsheet.getSheetByName("Locations");
    if (!locationsSheet) throw new Error("Missing Locations sheet.");

    const locations = locationsSheet
      .getDataRange().getValues().slice(1);

    let locationID = cell("CreatedLocationID")
      .getDisplayValue().trim();

    // Avoid creating another location when a request is retried.
    if (
      locationID &&
      locations.some(r => String(r[0]).trim() === locationID)
    ) {
      finish("Processed", "Location already created: " + locationID);
      return;
    }

    const name = cell("Location Name").getDisplayValue().trim();
    const notes = cell("Notes").getDisplayValue().trim();

    if (!name) {
      finish("Rejected", "Location Name is required.");
      return;
    }

    const properties = PropertiesService.getScriptProperties();
    let lastNumber = Number(
      properties.getProperty("LAST_LOCATION_NUMBER") || 0
    );

    if (!Number.isSafeInteger(lastNumber) || lastNumber < 0) {
      throw new Error("The location number counter is invalid.");
    }

    // Continue after the highest existing location number.
    for (const location of locations) {
      const match = String(location[0]).trim().match(/^LOC-(\d+)$/);
      if (match) {
        const number = Number(match[1]);
        if (!Number.isSafeInteger(number)) {
          throw new Error("Invalid location number: " + location[0]);
        }
        lastNumber = Math.max(lastNumber, number);
      }
    }

    if (!locationID) {
      lastNumber++;

      if (!Number.isSafeInteger(lastNumber)) {
        throw new Error("Location number exceeds the supported range.");
      }

      properties.setProperty(
        "LAST_LOCATION_NUMBER",
        String(lastNumber)
      );

      locationID = "LOC-" + String(lastNumber).padStart(3, "0");
      cell("CreatedLocationID").setValue(locationID);
      SpreadsheetApp.flush();
    } else {
      const match = locationID.match(/^LOC-(\d+)$/);
      if (!match || !Number.isSafeInteger(Number(match[1]))) {
        throw new Error("Invalid reserved LocationID.");
      }

      properties.setProperty(
        "LAST_LOCATION_NUMBER",
        String(Math.max(lastNumber, Number(match[1])))
      );
    }

    function safeText(value) {
      return /^[=+\-@]/.test(value) ? "'" + value : value;
    }

    locationsSheet.appendRow([
      locationID,
      safeText(name),
      safeText(notes)
    ]);

    SpreadsheetApp.flush();
    finish("Processed", "Location created: " + locationID);
  } finally {
    lock.releaseLock();
  }
}

function refreshNewBoxForm() {
  const form = inventoryOptionalForm_("INVENTORY_NEW_BOX_FORM_ID");
  if (!form) return;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Locations");

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error("Add at least one location in the Locations tab.");
  }

  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getDisplayValues();

  const seen = new Set();
  const choices = rows
    .filter(([id, name]) => id.trim() && name.trim())
    .map(([id, name]) => {
      id = id.trim();

      if (seen.has(id)) {
        throw new Error("Duplicate LocationID: " + id);
      }
      seen.add(id);

      return id + " — " + name.trim();
    });

  if (!choices.length) {
    throw new Error("Each location needs an ID and a name.");
  }

  const matches = form.getItems(FormApp.ItemType.LIST)
    .filter(question => question.getTitle().trim() === "Location");

  if (matches.length !== 1) {
    throw new Error(
      'The form needs exactly one dropdown named "Location".'
    );
  }

  matches[0].asListItem().setChoiceValues(choices);
  console.log("New Box location dropdown updated.");
}

//processNewBoxRequest
function processNewBoxRequest(e) {
  if (!e || !e.range) throw new Error("This function runs from a form submission.");
  const requests = e.range.getSheet();
  if (requests.getName() !== "NewBoxRequests") return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = requests.getParent(), row = e.range.getRow();
    const headers = requests.getRange(1, 1, 1, requests.getLastColumn()).getDisplayValues()[0];
    function cell(name) {
      if (headers.filter(header => header === name).length !== 1) throw new Error("Missing or duplicated NewBoxRequests header: " + name);
      return requests.getRange(row, headers.indexOf(name) + 1);
    }
    ["Box Name", "Location", "Notes", "RequestID", "Status", "Message", "ProcessedAt", "CreatedBoxID"].forEach(cell);
    creationHeaders_(requests, ["CreatedPlacementID", "InitialPlacementLocation", "InitialPlacementDescription", "InitialPlacementAt"]);
    if (["Processed", "Rejected"].includes(cell("Status").getDisplayValue().trim())) return;
    if (!cell("RequestID").getDisplayValue()) { cell("RequestID").setValue(Utilities.getUuid()); SpreadsheetApp.flush(); }
    function finish(status, message) {
      cell("Message").setValue(inventoryBoxSafeText_(message));
      cell("ProcessedAt").setValue(new Date()); cell("Status").setValue(status); SpreadsheetApp.flush();
    }
    const boxes = inventoryRequireNestedBoxes_(ss);
    inventoryRecoverBoxMoves_(ss);
    let boxID = cell("CreatedBoxID").getDisplayValue().trim();
    const existing = boxes.getDataRange().getValues().slice(1).filter(r => String(r[0]).trim() === boxID);
    if (boxID && existing.length > 1) throw new Error("Duplicated reserved BoxID: " + boxID);
    if (boxID && existing.length === 1) {
      recordInitialBoxPlacement_(ss, requests, row, boxID);
      finish("Processed", "Box already created: " + String(existing[0][1] || "the selected box")); return;
    }
    const name = cell("Box Name").getDisplayValue().trim();
    const destination = cell("Location").getDisplayValue().split(" — ")[0].trim();
    const notes = cell("Notes").getDisplayValue().trim();
    if (!name) { finish("Rejected", "Box Name is required."); return; }
    const locations = ss.getSheetByName("Locations");
    if (!locations) throw new Error("Missing Locations sheet.");
    const destinationCount = [boxes, locations].reduce((count, sheet) => count +
      sheet.getDataRange().getValues().slice(1).filter(value => String(value[0]).trim() === destination).length, 0);
    if (destinationCount !== 1) {
      finish("Rejected", "Select an existing location or parent box."); return;
    }
    const tree = inventoryStorageTree_(ss);
    if (destination === boxID) { finish("Rejected", "A box cannot be its own parent."); return; }
    const props = PropertiesService.getScriptProperties();
    let last = Number(props.getProperty("LAST_BOX_NUMBER") || 0);
    if (!Number.isSafeInteger(last) || last < 0) throw new Error("The box number counter is invalid.");
    for (const id of tree.boxes.keys()) {
      const match = id.match(/^BOX-(\d+)$/);
      if (!match) continue;
      const number = Number(match[1]);
      if (!Number.isSafeInteger(number)) throw new Error("Invalid box number: " + id);
      last = Math.max(last, number);
    }
    if (!boxID) {
      last++;
      if (!Number.isSafeInteger(last)) throw new Error("Box number exceeds the supported range.");
      props.setProperty("LAST_BOX_NUMBER", String(last));
      boxID = "BOX-" + String(last).padStart(3, "0");
      cell("CreatedBoxID").setValue(boxID); SpreadsheetApp.flush();
    } else {
      const match = boxID.match(/^BOX-(\d+)$/);
      if (!match || !Number.isSafeInteger(Number(match[1]))) throw new Error("Invalid reserved BoxID.");
      props.setProperty("LAST_BOX_NUMBER", String(Math.max(last, Number(match[1]))));
    }
    const fields = {
      BoxID: boxID, BoxName: name, LocationID: tree.resolve(destination).locationID,
      Notes: notes, PhotoURL: "", ParentBoxID: tree.boxes.has(destination) ? destination : ""
    };
    cell("CreatedPlacementID").setValue("CREATE-" + cell("RequestID").getDisplayValue().trim());
    cell("InitialPlacementLocation").setValue(tree.resolve(destination).locationID);
    cell("InitialPlacementDescription").setValue(inventoryBoxSafeText_(tree.resolve(destination).path));
    cell("InitialPlacementAt").setValue(new Date());
    SpreadsheetApp.flush();
    boxes.appendRow(tree.headers.map(header => inventoryBoxSafeText_(fields[header] == null ? "" : fields[header])));
    SpreadsheetApp.flush();
    recordInitialBoxPlacement_(ss, requests, row, boxID);
    finish("Processed", "Box created: " + name);
  } finally { lock.releaseLock(); }
}
function getLabelRecords_() {
  const ss = getInventorySpreadsheet_();
  const records = [];
  const seen = new Set();

  for (const sheetName of ["Items", "Boxes", "Locations"]) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Missing sheet: " + sheetName);
    if (sheet.getLastRow() < 2) continue;

    const rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getDisplayValues();

    for (const [rawID, rawName] of rows) {
      const id = rawID.trim();
      const name = rawName.trim();
      if (!id || !name) continue;

      if (seen.has(id)) {
        throw new Error("Duplicate label ID: " + id);
      }
      seen.add(id);

      records.push({
        id: id,
        name: name,
        type: sheetName
      });
    }
  }

  return records;
}
