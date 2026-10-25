function searchInventory(query) {
  if (typeof query !== "string" || !query.trim()) {
    throw new Error("Enter an item name or description to search.");
  }
  query = query.trim().replace(/\s+/g, " ");
  if (query.length > 200) {
    throw new Error("Please keep the search to 200 characters or fewer.");
  }
  const terms = Array.from(new Set(findInventoryText_(query).split(" ").filter(Boolean)));
  if (!terms.length) throw new Error("Enter an item name or description to search.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = getInventorySpreadsheet_();
    function rows(name) {
      const sheet = spreadsheet.getSheetByName(name);
      if (!sheet) throw new Error("Missing sheet: " + name);
      return sheet.getDataRange().getValues().slice(1)
        .filter(row => String(row[0] || "").trim() !== "");
    }

    const ids = new Set();
    const matches = [];
    for (const row of rows("Items")) {
      const id = String(row[0]).trim();
      if (ids.has(id)) throw new Error("Duplicate ItemID in Items: " + id);
      ids.add(id);
      const entry = {
        id: id, name: String(row[1] || "").trim(),
        category: String(row[2] || "").trim(), unit: String(row[3] || "").trim(),
        notes: String(row[5] || "").trim()
      };
      const text = findInventoryText_([
        entry.id, entry.name, entry.category, entry.notes
      ].join(" "));
      if (!terms.every(term => text.includes(term))) continue;
      if (!entry.name || !entry.unit) {
        throw new Error("Missing item name or unit in Items: " + id);
      }
      matches.push(entry);
    }
    matches.sort((a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
    const report = { query: query, generatedAt: new Date().toISOString(), items: [] };
    if (!matches.length) return report;

    const storage = inventoryStorageTree_(spreadsheet);

    const transactions = spreadsheet.getSheetByName("Transactions");
    if (!transactions) throw new Error("Missing sheet: Transactions");
    const expected = [
      "TransactionID", "Timestamp", "Action", "ItemID",
      "Quantity", "FromStorageID", "ToStorageID", "Notes"
    ];
    const headers = transactions.getRange(1, 1, 1, 8).getDisplayValues()[0];
    if (!expected.every((name, index) => String(headers[index] || "").trim() === name)) {
      throw new Error("Transactions column headers do not match.");
    }
    const matchedIDs = new Set(matches.map(entry => entry.id));
    const balances = new Map(matches.map(entry => [entry.id, new Map()]));
    function change(itemID, storageID, quantity) {
      if (!storageID) return;
      const positions = balances.get(itemID);
      const total = (positions.get(storageID) || 0) + quantity;
      if (!Number.isFinite(total)) throw new Error("The stock balance exceeds the supported range.");
      positions.set(storageID, total);
    }
    const transactionIDs = new Set();
    for (const transaction of rows("Transactions")) {
      const transactionID = String(transaction[0]).trim();
      if (transactionIDs.has(transactionID)) {
        throw new Error("Duplicate transaction reference: " + transactionID);
      }
      transactionIDs.add(transactionID);
      const quantity = Number(transaction[4]);
      if (!Number.isFinite(quantity) || quantity <= 0) {
        throw new Error("Invalid quantity in transaction: " + transactionID);
      }
      const itemID = String(transaction[3]).trim();
      if (!matchedIDs.has(itemID)) continue;
      change(itemID, String(transaction[5] || "").trim(), -quantity);
      change(itemID, String(transaction[6] || "").trim(), quantity);
    }

    for (const entry of matches) {
      entry.positions = [];
      entry.availableQuantity = 0;
      entry.withFamilyQuantity = 0;
      for (const [storageID, rawQuantity] of balances.get(entry.id)) {
        const quantity = findInventoryQuantity_(rawQuantity);
        if (quantity === 0) continue;
        if (quantity < 0) {
          throw new Error("Negative stock for " + entry.id + " at " + storageID +
            ". Review Transactions before using this report.");
        }
        const position = {
          storageID: storageID, kind: "", boxName: "", locationID: "",
          locationName: "", memberEmail: "", quantity: quantity
        };
        if (storageID.startsWith("MEMBER:")) {
          const email = storageID.slice(7);
          if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
              email !== email.trim().toLowerCase()) {
            throw new Error("Invalid family-member storage ID: " + storageID);
          }
          position.kind = "member";
          position.memberEmail = email;
          entry.withFamilyQuantity += quantity;
        } else {
          const box = storage.boxes.get(storageID);
          const place = storage.resolve(storageID);
          position.kind = box ? "box" : "location";
          position.boxName = box ? box.name : "";
          position.locationID = place.locationID;
          position.locationName = place.locationName;
          position.boxPath = place.boxPath;
          position.storagePath = place.path;
          entry.availableQuantity += quantity;
        }
        entry.positions.push(position);
      }
      entry.availableQuantity = findInventoryQuantity_(entry.availableQuantity);
      entry.withFamilyQuantity = findInventoryQuantity_(entry.withFamilyQuantity);
      entry.totalQuantity = findInventoryQuantity_(entry.availableQuantity + entry.withFamilyQuantity);
      entry.positions.sort((a, b) =>
        Number(a.kind === "member") - Number(b.kind === "member") ||
        a.locationName.localeCompare(b.locationName) ||
        a.locationID.localeCompare(b.locationID) ||
        a.storageID.localeCompare(b.storageID)
      );
      report.items.push(entry);
    }
    report.generatedAt = new Date().toISOString();
    return report;
  } finally {
    lock.releaseLock();
  }
}

function findInventoryText_(value) {
  return String(value).normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

function findInventoryQuantity_(value) {
  if (!Number.isFinite(value)) throw new Error("The stock total exceeds the supported range.");
  return Number(value.toFixed(6));
}