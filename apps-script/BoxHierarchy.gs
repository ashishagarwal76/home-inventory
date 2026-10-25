function initializeNestedBoxes() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getInventorySpreadsheet_();
    const boxes = ss.getSheetByName("Boxes");
    if (!boxes) throw new Error("Missing Boxes sheet.");
    inventoryBoxHeaders_(boxes, ["BoxID", "BoxName", "LocationID", "Notes", "PhotoURL"]);
    inventoryBoxAddHeaders_(boxes, ["ParentBoxID"]);
    let moves = ss.getSheetByName("BoxMoves");
    if (!moves) {
      moves = ss.insertSheet("BoxMoves");
      moves.getRange(1, 1, 1, 6).setValues([[
        "MovementID", "Timestamp", "BoxID", "FromLocationID", "ToLocationID", "Notes"
      ]]);
    }
    inventoryBoxHeaders_(moves, [
      "MovementID", "Timestamp", "BoxID", "FromLocationID", "ToLocationID", "Notes"
    ]);
    inventoryBoxAddHeaders_(moves, [
      "FromStorageID", "ToStorageID", "Status", "Message", "ProcessedAt"
    ]);
    inventoryRecoverBoxMoves_(ss);
    inventoryWriteBoxRoots_(inventoryStorageTree_(ss));
    SpreadsheetApp.flush();
    return "Nested boxes are ready. ParentBoxID was added; existing boxes keep their locations.";
  } finally {
    lock.releaseLock();
  }
}

function inventoryBoxHeaders_(sheet, expected) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  expected.forEach((name, index) => {
    if (headers[index] !== name) throw new Error(sheet.getName() + " column " + (index + 1) + " must be " + name + ".");
  });
  return headers;
}

function inventoryBoxAddHeaders_(sheet, names) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (const name of names) {
    const count = headers.filter(value => value === name).length;
    if (count > 1) throw new Error("Duplicated " + sheet.getName() + " header: " + name);
    if (count) continue;
    if (sheet.getMaxColumns() < headers.length + 1) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length + 1 - sheet.getMaxColumns());
    }
    headers.push(name);
    sheet.getRange(1, headers.length).setValue(name);
  }
  return headers;
}

function inventoryRequireNestedBoxes_(ss) {
  const sheet = ss.getSheetByName("Boxes");
  if (!sheet) throw new Error("Missing Boxes sheet.");
  const headers = inventoryBoxHeaders_(sheet, ["BoxID", "BoxName", "LocationID", "Notes", "PhotoURL"]);
  if (headers.filter(name => name === "ParentBoxID").length !== 1) {
    throw new Error("Run initializeNestedBoxes once in Apps Script before using nested boxes.");
  }
  return sheet;
}

function inventoryStorageTree_(ss) {
  const boxesSheet = ss.getSheetByName("Boxes");
  const locationsSheet = ss.getSheetByName("Locations");
  if (!boxesSheet || !locationsSheet) throw new Error("Missing Boxes or Locations sheet.");
  const data = boxesSheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const parentColumn = headers.indexOf("ParentBoxID");
  if (headers.filter(name => name === "ParentBoxID").length > 1) throw new Error("Duplicated ParentBoxID header.");
  const boxes = new Map(), locations = new Map(), used = new Set();
  for (const row of locationsSheet.getDataRange().getValues().slice(1)) {
    const id = String(row[0] || "").trim();
    if (!id) continue;
    if (used.has(id) || id.startsWith("MEMBER:")) throw new Error("Duplicated or reserved storage ID: " + id);
    used.add(id);
    locations.set(id, String(row[1] || "Unnamed location").trim());
  }
  data.slice(1).forEach((row, index) => {
    const id = String(row[0] || "").trim();
    if (!id) return;
    if (used.has(id) || id.startsWith("MEMBER:")) throw new Error("Duplicated or reserved storage ID: " + id);
    used.add(id);
    boxes.set(id, {
      id: id, name: String(row[1] || "Unnamed box").trim(), locationID: String(row[2] || "").trim(),
      parentBoxID: parentColumn < 0 ? "" : String(row[parentColumn] || "").trim(), row: index + 2
    });
  });
  const resolved = new Map();
  function resolve(id) {
    if (locations.has(id)) return {
      locationID: id, locationName: locations.get(id), boxIDs: [], boxPath: "",
      path: locations.get(id)
    };
    if (resolved.has(id)) return resolved.get(id);
    const chain = [], seen = new Set();
    let current = id, base;
    while (true) {
      if (resolved.has(current)) { base = resolved.get(current); break; }
      if (seen.has(current)) throw new Error("Box containment loop: " + chain.map(box => box.id).concat(current).join(" > "));
      const box = boxes.get(current);
      if (!box) throw new Error("Missing box or parent box: " + current);
      seen.add(current); chain.push(box);
      if (box.parentBoxID) { current = box.parentBoxID; continue; }
      if (!locations.has(box.locationID)) throw new Error("Missing location for box " + box.id + ": " + box.locationID);
      base = resolve(box.locationID); break;
    }
    for (let index = chain.length - 1; index >= 0; index--) {
      const box = chain[index];
      const label = box.name;
      base = {
        locationID: base.locationID, locationName: base.locationName,
        boxIDs: base.boxIDs.concat(box.id),
        boxPath: base.boxPath ? base.boxPath + " > " + label : label,
        path: base.path + " > " + label
      };
      resolved.set(box.id, base);
    }
    return resolved.get(id);
  }
  boxes.forEach(box => resolve(box.id));
  function choices() {
    const entries = Array.from(locations, ([id, name]) => ({id: id, label: "Location: " + name, kind: "location"}));
    boxes.forEach(box => entries.push({id: box.id, label: "Box: " + resolve(box.id).path, kind: "box"}));
    return entries.sort((a, b) => a.id.localeCompare(b.id));
  }
  return {ss: ss, sheet: boxesSheet, data: data, headers: headers, parentColumn: parentColumn,
    boxes: boxes, locations: locations, resolve: resolve, choices: choices};
}

function inventoryBoxSafeText_(value) {
  return typeof value === "string" && /^[=+\-@']/.test(value) ? "'" + value : value;
}

function inventoryWriteBoxRoots_(tree, placement) {
  if (placement) {
    if (tree.parentColumn < 0) throw new Error("Run initializeNestedBoxes first.");
    const box = tree.boxes.get(placement.boxID);
    if (!box) throw new Error("The selected box no longer exists.");
    if (placement.destination === box.id) throw new Error("A box cannot be placed inside itself.");
    if (!tree.boxes.has(placement.destination) && !tree.locations.has(placement.destination)) {
      throw new Error("The selected destination no longer exists.");
    }
    const target = tree.resolve(placement.destination);
    if (target.boxIDs.includes(box.id)) throw new Error("A box cannot be placed inside one of its descendants.");
    box.parentBoxID = tree.boxes.has(placement.destination) ? placement.destination : "";
    box.locationID = target.locationID;
    // Rebuild the proposed tree using a small in-memory spreadsheet view.
    const changed = tree.data.map(row => row.slice());
    changed[box.row - 1][2] = target.locationID;
    changed[box.row - 1][tree.parentColumn] = box.parentBoxID;
    const virtual = {getSheetByName(name) {
      if (name !== "Boxes") return tree.ss.getSheetByName(name);
      return {getDataRange: () => ({getValues: () => changed})};
    }};
    const proposed = inventoryStorageTree_(virtual);
    return inventoryCommitBoxRows_(tree, proposed, changed);
  }
  return inventoryCommitBoxRows_(tree, tree, tree.data.map(row => row.slice()));
}

function inventoryCommitBoxRows_(original, proposed, values) {
  let changed = false;
  proposed.boxes.forEach(box => {
    const root = proposed.resolve(box.id).locationID;
    if (String(values[box.row - 1][2] || "").trim() !== root) changed = true;
    values[box.row - 1][2] = root;
  });
  if (!changed) changed = values.some((row, r) => row.some((value, c) => value !== original.data[r][c]));
  if (!changed || values.length < 2) return;
  const width = original.headers.length;
  const range = original.sheet.getRange(2, 1, values.length - 1, width);
  const formulas = range.getFormulas();
  const safeRows = values.slice(1).map((row, r) => Array.from({length: width}, (_, c) => {
    if (c !== 2 && c !== original.parentColumn && formulas[r] && formulas[r][c]) return formulas[r][c];
    return inventoryBoxSafeText_(row[c] == null ? "" : row[c]);
  }));
  range.setValues(safeRows);
}
