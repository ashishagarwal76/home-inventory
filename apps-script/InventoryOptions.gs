function getAddInventoryOptions() {
  const spreadsheet = getInventorySpreadsheet_();
  function readCatalog(name) {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) throw new Error("Missing sheet: " + name);
    const seen = new Set();
    return sheet.getDataRange().getDisplayValues().slice(1)
      .filter(row => row.some(value => value.trim() !== ""))
      .map(row => {
        const id = row[0].trim(), description = row[1].trim();
        if (!id || !description) throw new Error(name + ": each record needs an ID and description.");
        if (seen.has(id)) throw new Error("Duplicate ID in " + name + ": " + id);
        seen.add(id);
        return {id: id, name: description, row: row};
      });
  }
  const items = readCatalog("Items");
  readCatalog("Locations"); readCatalog("Boxes");
  const labels = inventoryItemDescriptionLabels_(items.map(item => ({
    id: item.id, name: item.name, category: item.row[2] || "", notes: item.row[5] || ""
  })));
  return {
    items: items.map(item => ({id: item.id, label: labels.get(item.id), unit: String(item.row[3] || "").trim()})),
    destinations: inventoryStorageTree_(spreadsheet).choices()
  };
}

function getPhotoUploadOptions() {
  const spreadsheet = getInventorySpreadsheet_();
  function readRecords(sheetName, idHeader, nameHeader) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error("Missing sheet: " + sheetName);
    const data = sheet.getDataRange().getDisplayValues(), headers = data[0].map(value => value.trim());
    function column(name) {
      const index = headers.indexOf(name);
      if (index < 0) throw new Error(sheetName + " needs a " + name + " column.");
      return index;
    }
    const idColumn = column(idHeader), nameColumn = column(nameHeader), photoColumn = column("PhotoURL"), seen = new Set();
    return data.slice(1).filter(row => row.some(value => value.trim() !== "")).map(row => {
      const id = row[idColumn].trim(), name = row[nameColumn].trim();
      if (!id || !name) throw new Error(sheetName + ": each record needs an ID and description.");
      if (seen.has(id)) throw new Error("Duplicate ID in " + sheetName + ": " + id);
      seen.add(id);
      return {id: id, name: name, category: row[headers.indexOf("Category")] || "", notes: row[headers.indexOf("Notes")] || "", hasPhoto: Boolean(row[photoColumn].trim())};
    });
  }
  const items = readRecords("Items", "ItemID", "ItemName"), boxes = readRecords("Boxes", "BoxID", "BoxName");
  const labels = inventoryItemDescriptionLabels_(items), tree = inventoryStorageTree_(spreadsheet);
  return {
    items: items.map(item => ({id: item.id, label: labels.get(item.id), hasPhoto: item.hasPhoto})),
    boxes: boxes.map(box => ({id: box.id, label: tree.resolve(box.id).path, hasPhoto: box.hasPhoto}))
  };
}

function inventoryItemDescriptionLabels_(items) {
  const counts = new Map();
  items.forEach(item => { const name = item.name.toLowerCase(); counts.set(name, (counts.get(name) || 0) + 1); });
  const preliminary = items.map(item => ({item: item, label: item.name +
    (counts.get(item.name.toLowerCase()) > 1 && item.category ? " — " + item.category : "")}));
  const duplicates = new Map();
  preliminary.forEach(entry => duplicates.set(entry.label.toLowerCase(), (duplicates.get(entry.label.toLowerCase()) || 0) + 1));
  return new Map(preliminary.map(entry => [entry.item.id, entry.label +
    (duplicates.get(entry.label.toLowerCase()) > 1 && entry.item.notes ? " — " + entry.item.notes : "")]));
}
