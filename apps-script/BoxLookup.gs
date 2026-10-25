function getScannedBoxContents_(boxID) {
  if (typeof boxID !== "string" || !boxID.trim() || boxID.length > 100) {
    throw new Error("Enter a valid Box ID.");
  }
  boxID = boxID.trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getInventorySpreadsheet_();
    const tree = inventoryStorageTree_(ss);
    const box = tree.boxes.get(boxID);
    if (!box) throw new Error("This box no longer exists.");

    const placement = tree.resolve(boxID);
    const includedBoxes = new Set([boxID]);
    const nestedBoxes = [];

    function relativeBoxPath(id) {
      const ids = tree.resolve(id).boxIDs;
      return ids.slice(ids.indexOf(boxID))
        .map(value => tree.boxes.get(value).name).join(" > ");
    }

    tree.boxes.forEach(function (child) {
      if (child.id === boxID) return;
      const childPlacement = tree.resolve(child.id);
      if (!childPlacement.boxIDs.includes(boxID)) return;
      includedBoxes.add(child.id);
      nestedBoxes.push({
        id: child.id,
        name: child.name,
        parentBoxID: child.parentBoxID,
        path: relativeBoxPath(child.id)
      });
    });
    nestedBoxes.sort((a, b) => a.path.localeCompare(b.path) || a.id.localeCompare(b.id));

    const itemsSheet = ss.getSheetByName("Items");
    const transactionsSheet = ss.getSheetByName("Transactions");
    if (!itemsSheet || !transactionsSheet) {
      throw new Error("Missing Items or Transactions sheet.");
    }

    const items = new Map();
    itemsSheet.getDataRange().getValues().slice(1).forEach(function (row) {
      const id = String(row[0] || "").trim();
      if (!id) return;
      if (items.has(id)) throw new Error("Duplicate ItemID in Items: " + id);
      items.set(id, {
        id: id,
        name: String(row[1] || "").trim(),
        unit: String(row[3] || "").trim()
      });
    });

    const expectedHeaders = [
      "TransactionID", "Timestamp", "Action", "ItemID",
      "Quantity", "FromStorageID", "ToStorageID", "Notes"
    ];
    const headers = transactionsSheet.getRange(1, 1, 1, 8).getDisplayValues()[0];
    if (!expectedHeaders.every((name, index) => String(headers[index] || "").trim() === name)) {
      throw new Error("Transactions column headers do not match.");
    }

    const balances = new Map();
    const transactionIDs = new Set();

    function change(itemID, storageID, amount) {
      if (!includedBoxes.has(storageID)) return;
      const key = JSON.stringify([itemID, storageID]);
      const next = (balances.get(key) || 0) + amount;
      if (!Number.isFinite(next)) throw new Error("The stock balance exceeds the supported range.");
      balances.set(key, next);
    }

    transactionsSheet.getDataRange().getValues().slice(1).forEach(function (row) {
      const transactionID = String(row[0] || "").trim();
      if (!transactionID) return;
      if (transactionIDs.has(transactionID)) {
        throw new Error("Duplicate transaction reference: " + transactionID);
      }
      transactionIDs.add(transactionID);

      const amount = Number(row[4]);
      if (!Number.isFinite(amount) || amount <= 0) {
        throw new Error("Invalid quantity in transaction: " + transactionID);
      }
      const itemID = String(row[3] || "").trim();
      change(itemID, String(row[5] || "").trim(), -amount);
      change(itemID, String(row[6] || "").trim(), amount);
    });

    const contents = [];
    const itemTotals = new Map();
    for (const [key, rawQuantity] of balances) {
      const quantity = findInventoryQuantity_(rawQuantity);
      if (quantity === 0) continue;
      const [itemID, storageID] = JSON.parse(key);
      if (quantity < 0) {
        throw new Error("Negative stock for " + itemID + " at " + storageID + ". Review Transactions.");
      }

      const item = items.get(itemID);
      if (!item || !item.name || !item.unit) {
        throw new Error("Missing item name or unit for: " + itemID);
      }
      contents.push({
        itemID: item.id,
        itemName: item.name,
        unit: item.unit,
        quantity: quantity,
        boxID: storageID,
        boxName: tree.boxes.get(storageID).name,
        path: relativeBoxPath(storageID)
      });

      const total = itemTotals.get(itemID) || {
        itemID: item.id, itemName: item.name, unit: item.unit, quantity: 0
      };
      total.quantity = findInventoryQuantity_(total.quantity + quantity);
      itemTotals.set(itemID, total);
    }
    contents.sort((a, b) => a.path.localeCompare(b.path) ||
      a.itemName.localeCompare(b.itemName) || a.itemID.localeCompare(b.itemID));

    return {
      id: box.id,
      name: box.name,
      locationID: placement.locationID,
      locationName: placement.locationName,
      parentBoxID: box.parentBoxID,
      parentBoxName: box.parentBoxID ? tree.boxes.get(box.parentBoxID).name : "",
      path: placement.path,
      nestedBoxes: nestedBoxes,
      contents: contents,
      itemTotals: Array.from(itemTotals.values()).sort((a, b) =>
        a.itemName.localeCompare(b.itemName) || a.itemID.localeCompare(b.itemID)),
      generatedAt: new Date().toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}
