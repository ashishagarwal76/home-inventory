function getMoveBoxOptions() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getInventorySpreadsheet_();
    inventoryRequireNestedBoxes_(ss);
    inventoryRecoverBoxMoves_(ss);
    const tree = inventoryStorageTree_(ss);
    inventoryWriteBoxRoots_(tree);
    const destinations = tree.choices();
    const boxes = Array.from(tree.boxes.values()).map(box => {
      const place = tree.resolve(box.id);
      return {
        id: box.id, label: place.path,
        locationID: place.locationID, locationName: place.locationName,
        parentBoxID: box.parentBoxID, currentStorageID: box.parentBoxID || place.locationID,
        path: place.path,
        descendantIDs: Array.from(tree.boxes.keys()).filter(id => tree.resolve(id).boxIDs.includes(box.id))
      };
    }).sort((a, b) => a.id.localeCompare(b.id));
    return {boxes: boxes, destinations: destinations, locations: destinations};
  } finally {
    lock.releaseLock();
  }
}

function submitWebMoveBox(input) {
  const requestID = String(input && input.requestID || "").trim();
  const boxID = String(input && input.boxID || "").trim();
  const destination = String(input && input.destination || "").trim();
  const notes = String(input && input.notes || "").trim();
  const expectedFrom = String(input && input.expectedFromStorageID || "").trim();
  function response(accepted, message) { return {accepted: accepted, requestID: requestID, boxID: boxID, message: message}; }
  if (!input || typeof input !== "object" || Array.isArray(input)) return response(false, "No request was received.");
  if (!/^WEB-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(requestID)) {
    return response(false, "Invalid movement reference. Reload the page and try again.");
  }
  if (!boxID || !destination) return response(false, "Select a box and a destination box or location.");
  if (notes.length > 2000) return response(false, "Please keep Notes to 2,000 characters or fewer.");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  try {
    const ss = getInventorySpreadsheet_();
    inventoryRequireNestedBoxes_(ss);
    inventoryRecoverBoxMoves_(ss);
    let history = inventoryBoxMoveHistory_(ss);
    let saved = history.entries.find(entry => entry.id === requestID);
    if (saved) {
      const to = saved.toStorage || saved.toLocation;
      if (saved.boxID !== boxID || to !== destination || saved.notes !== notes ||
          (expectedFrom && saved.fromStorage && expectedFrom !== saved.fromStorage)) {
        return response(false, "This movement reference was already used for different details.");
      }
      if (!saved.status) {
        // Before this update, history was appended before LocationID was saved.
        const tree = inventoryStorageTree_(ss), box = tree.boxes.get(boxID);
        const later = history.entries.some(entry => entry.row > saved.row && entry.boxID === boxID);
        if (!later && box && !box.parentBoxID && box.locationID === saved.fromLocation && tree.locations.has(destination)) {
          history.sheet.getRange(saved.row, history.columns.FromStorageID + 1).setValue(saved.fromLocation);
          history.sheet.getRange(saved.row, history.columns.ToStorageID + 1).setValue(destination);
          history.sheet.getRange(saved.row, history.columns.Status + 1).setValue("Pending");
          inventoryRecoverBoxMoves_(ss);
        } else {
          inventoryBoxMoveFinish_(history, saved, "Processed", "Movement already recorded. The current box placement was preserved.");
        }
        history = inventoryBoxMoveHistory_(ss);
        saved = history.entries.find(entry => entry.id === requestID);
      }
      result = response(saved.status === "Processed", saved.message ||
        (saved.status === "Processed" ? "Movement was already saved. No duplicate recorded." : "The previous movement was rejected. Refresh the choices."));
    } else {
      const tree = inventoryStorageTree_(ss), box = tree.boxes.get(boxID);
      if (!box) return response(false, "The selected box no longer exists. Refresh the choices.");
      if (!tree.boxes.has(destination) && !tree.locations.has(destination)) return response(false, "The selected destination no longer exists. Refresh the choices.");
      const from = box.parentBoxID || tree.resolve(boxID).locationID;
      if (expectedFrom && expectedFrom !== from) return response(false, "The box placement changed. Refresh the choices before moving it.");
      if (destination === boxID || tree.resolve(destination).boxIDs.includes(boxID)) {
        return response(false, "A box cannot go inside itself or one of its descendants.");
      }
      if (from === destination) return response(false, "This box is already directly in the selected destination.");
      const values = {
        MovementID: requestID, Timestamp: new Date(), BoxID: boxID,
        FromLocationID: tree.resolve(boxID).locationID,
        ToLocationID: tree.resolve(destination).locationID,
        Notes: notes, FromStorageID: from, ToStorageID: destination, Status: "Pending"
      };
      history.sheet.appendRow(history.headers.map(header => inventoryBoxSafeText_(values[header] == null ? "" : values[header])));
      SpreadsheetApp.flush();
      inventoryRecoverBoxMoves_(ss);
      saved = inventoryBoxMoveHistory_(ss).entries.find(entry => entry.id === requestID);
      result = response(saved.status === "Processed", saved.message);
    }
  } finally {
    lock.releaseLock();
  }
  if (result.accepted) {
    try { refreshStock(); result.stockRefreshed = true; }
    catch (error) {
      console.error(error); result.stockRefreshed = false;
      result.message += " The move was saved, but Stock could not refresh. Run refreshStock; do not submit another move.";
    }
  }
  return result;
}

function inventoryBoxMoveHistory_(ss) {
  const sheet = ss.getSheetByName("BoxMoves");
  if (!sheet) throw new Error("Run initializeNestedBoxes once in Apps Script.");
  const headers = inventoryBoxHeaders_(sheet, ["MovementID", "Timestamp", "BoxID", "FromLocationID", "ToLocationID", "Notes"]);
  const columns = {};
  ["FromStorageID", "ToStorageID", "Status", "Message", "ProcessedAt"].forEach(name => {
    if (headers.filter(header => header === name).length !== 1) throw new Error("Run initializeNestedBoxes once in Apps Script.");
  });
  headers.forEach((name, index) => { columns[name] = index; });
  const seen = new Set(), entries = [];
  sheet.getDataRange().getValues().slice(1).forEach((row, index) => {
    const id = String(row[0] || "").trim();
    if (!id) return;
    if (seen.has(id)) throw new Error("Duplicate box movement reference: " + id);
    seen.add(id);
    function value(name) { return String(row[columns[name]] || "").trim(); }
    entries.push({id: id, row: index + 2, boxID: value("BoxID"),
      fromLocation: value("FromLocationID"), toLocation: value("ToLocationID"),
      fromStorage: value("FromStorageID"), toStorage: value("ToStorageID"),
      notes: value("Notes"), status: value("Status"), message: value("Message")});
  });
  return {sheet: sheet, headers: headers, columns: columns, entries: entries};
}

function inventoryBoxMoveFinish_(history, entry, status, message) {
  history.sheet.getRange(entry.row, history.columns.Message + 1).setValue(inventoryBoxSafeText_(message));
  history.sheet.getRange(entry.row, history.columns.ProcessedAt + 1).setValue(new Date());
  history.sheet.getRange(entry.row, history.columns.Status + 1).setValue(status);
  SpreadsheetApp.flush();
}

function inventoryRecoverBoxMoves_(ss) {
  const history = inventoryBoxMoveHistory_(ss);
  for (const entry of history.entries.filter(row => row.status === "Pending")) {
    const tree = inventoryStorageTree_(ss), box = tree.boxes.get(entry.boxID);
    const from = box && (box.parentBoxID || tree.resolve(box.id).locationID);
    if (!box || !entry.fromStorage || !entry.toStorage ||
        (!tree.boxes.has(entry.toStorage) && !tree.locations.has(entry.toStorage))) {
      inventoryBoxMoveFinish_(history, entry, "Rejected", "The box or destination changed before the saved move could complete. Refresh the choices.");
      continue;
    }
    if (from !== entry.fromStorage && from !== entry.toStorage) {
      inventoryBoxMoveFinish_(history, entry, "Rejected", "The box placement changed before this saved move could complete. Its current placement was preserved.");
      continue;
    }
    if (from === entry.fromStorage) {
      if (entry.toStorage === box.id || tree.resolve(entry.toStorage).boxIDs.includes(box.id)) {
        inventoryBoxMoveFinish_(history, entry, "Rejected", "This move would put a box inside itself or one of its descendants.");
        continue;
      }
      inventoryWriteBoxRoots_(tree, {boxID: box.id, destination: entry.toStorage});
      SpreadsheetApp.flush();
    } else {
      inventoryWriteBoxRoots_(tree);
    }
    const path = inventoryStorageTree_(ss).resolve(box.id).path;
    inventoryBoxMoveFinish_(history, entry, "Processed", "Box moved: " + path + ". All contained boxes and their inventory follow this placement.");
  }
}

function getMoveBoxThumbnail(boxID) {
  try {
    const sheet = getInventorySpreadsheet_().getSheetByName("Boxes");
    const rows = sheet.getDataRange().getValues().slice(1).filter(row => String(row[0]).trim() === String(boxID || "").trim());
    if (rows.length !== 1) return "";
    const url = String(rows[0][4] || "");
    if (!/^https:\/\/drive\.google\.com\//i.test(url)) return "";
    const match = url.match(/\/d\/([A-Za-z0-9_-]+)/) || url.match(/[?&]id=([A-Za-z0-9_-]+)/);
    const folderID = PropertiesService.getScriptProperties().getProperty("INVENTORY_PHOTO_FOLDER_ID");
    if (!match || !folderID) return "";
    const file = DriveApp.getFileById(match[1]), parents = file.getParents();
    let allowed = false;
    while (parents.hasNext()) if (parents.next().getId() === folderID) allowed = true;
    if (!allowed) return "";
    const blob = file.getThumbnail() || file.getBlob(), mime = blob.getContentType(), bytes = blob.getBytes();
    if (!/^image\/(jpeg|png|webp|gif)$/.test(mime) || bytes.length > 1048576) return "";
    return "data:" + mime + ";base64," + Utilities.base64Encode(bytes);
  } catch (error) { console.warn(error); return ""; }
}
