function getNewBoxOptions() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getInventorySpreadsheet_();
    inventoryRequireNestedBoxes_(ss);
    inventoryRecoverBoxMoves_(ss);
    const tree = inventoryStorageTree_(ss);
    inventoryWriteBoxRoots_(tree);
    // Keep the existing response key and field name for saved New Box retries.
    return { locations: tree.choices() };
  } finally { lock.releaseLock(); }
}

function submitWebNewLocation(input) {
  return newStorageSubmit_(input, {
    requestPrefix: "NEWLOC",
    requestSheet: "NewLocationRequests",
    recordSheet: "Locations",
    recordHeaders: ["LocationID", "LocationName", "Notes"],
    fieldHeaders: ["Location Name", "Notes"],
    createdHeader: "CreatedLocationID",
    resultKey: "locationID",
    idPattern: /^LOC-\d+$/,
    processor: processNewLocationRequest,
    fields: {
      "Location Name": input && input.locationName,
      "Notes": input && input.notes
    }
  });
}

function submitWebNewBox(input) {
  return newStorageSubmit_(input, {
    requestPrefix: "NEWBOX",
    requestSheet: "NewBoxRequests",
    recordSheet: "Boxes",
    recordHeaders: ["BoxID", "BoxName", "LocationID", "Notes", "PhotoURL"],
    fieldHeaders: ["Box Name", "Location", "Notes"],
    createdHeader: "CreatedBoxID",
    resultKey: "boxID",
    idPattern: /^BOX-\d+$/,
    processor: processNewBoxRequest,
    fields: {
      "Box Name": input && input.boxName,
      "Location": input && input.locationID,
      "Notes": input && input.notes
    },
    validateNew: () => newStoragePhotoError_(input && input.photoRequest)
  });
}

function newStorageCheckHeaders_(sheet, expected) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  expected.forEach((header, index) => {
    if (headers[index] !== header) {
      throw new Error(sheet.getName() + " column " + (index + 1) +
        " must be " + header + ".");
    }
  });
}

function newStoragePhotoError_(photo) {
  if (photo == null) return "";
  const uuid = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";
  if (typeof photo !== "object" || photo.recordType !== "Box" || photo.recordID !== "" ||
      !new RegExp("^PHOTO-" + uuid + "$", "i").test(photo.requestID) ||
      typeof photo.base64 !== "string" || !photo.base64 ||
      photo.base64.length > 1300000 || !/^[A-Za-z0-9+/]+={0,2}$/.test(photo.base64)) {
    return "The cropped box photo is invalid. Choose the photo again.";
  }
  try {
    const bytes = Utilities.base64Decode(photo.base64);
    if (bytes.length < 4 || bytes.length > 1048576 ||
        (bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 ||
        (bytes[bytes.length - 2] & 255) !== 255 ||
        (bytes[bytes.length - 1] & 255) !== 217) {
      return "Choose a photo and crop it before creating the box.";
    }
  } catch (error) {
    return "The cropped photo could not be read. Choose the photo again.";
  }
  return "";
}

function newStorageSubmit_(input, config) {
  if (!input || typeof input !== "object") throw new Error("Missing request.");
  const requestID = String(input.requestID || "").trim();
  const uuid = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";
  if (!new RegExp("^" + config.requestPrefix + "-" + uuid + "$", "i").test(requestID)) {
    throw new Error("Invalid request reference.");
  }

  const spreadsheet = getInventorySpreadsheet_();
  const requests = spreadsheet.getSheetByName(config.requestSheet);
  const records = spreadsheet.getSheetByName(config.recordSheet);
  if (!requests || !records) {
    throw new Error("Missing " + config.requestSheet + " or " + config.recordSheet + " sheet.");
  }
  newStorageCheckHeaders_(records, config.recordHeaders);

  const headers = requests
    .getRange(1, 1, 1, requests.getLastColumn()).getDisplayValues()[0];
  const required = config.fieldHeaders.concat([
    "RequestID", "Status", "Message", "ProcessedAt", config.createdHeader
  ]);
  required.forEach(header => {
    if (headers.filter(value => value === header).length !== 1) {
      throw new Error("Missing or duplicated " + config.requestSheet + " header: " + header);
    }
  });

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
    if (matches.length > 1) throw new Error("Duplicate request reference.");
    return matches[0] || 0;
  }

  function response(accepted, recordID, message) {
    const result = { accepted: accepted, requestID: requestID, message: message };
    result[config.resultKey] = recordID;
    return result;
  }

  let requestRow;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    requestRow = findRequestRow();
    if (!requestRow) {
      // Validate a chosen photo before any new box request or box is written.
      if (config.validateNew) {
        const error = config.validateNew();
        if (error) return response(false, "", error);
      }
      const values = { "Timestamp": new Date(), "RequestID": requestID };
      for (const header of config.fieldHeaders) {
        const value = config.fields[header];
        const text = value == null ? "" : String(value).trim();
        values[header] = /^[=+\-@']/.test(text) ? "'" + text : text;
      }
      requests.appendRow(headers.map(header =>
        Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ""
      ));
      SpreadsheetApp.flush();
      requestRow = findRequestRow();
      if (!requestRow) throw new Error("Could not locate the saved request.");
    }
  } finally {
    lock.releaseLock();
  }

  // Each existing processor acquires its own script lock.
  config.processor({ range: requests.getRange(requestRow, 1, 1, headers.length) });

  requestRow = findRequestRow();
  if (!requestRow) throw new Error("The saved request could not be found.");
  const row = requests
    .getRange(requestRow, 1, 1, headers.length).getDisplayValues()[0];
  const status = row[headers.indexOf("Status")].trim();
  if (!["Processed", "Rejected"].includes(status)) {
    throw new Error("The result is not confirmed. Retry the same request.");
  }
  const recordID = row[headers.indexOf(config.createdHeader)].trim();
  if (status === "Processed") {
    if (!config.idPattern.test(recordID) ||
        records.getDataRange().getDisplayValues().slice(1)
          .filter(value => value[0].trim() === recordID).length !== 1) {
      throw new Error("The created record is missing or duplicated.");
    }
  }
  return response(status === "Processed", recordID, row[headers.indexOf("Message")]);
}