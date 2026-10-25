function getFamilyInventoryOptions() {
  let stockWarning = "";
  try {
    refreshStock();
  } catch (error) {
    console.error(error);
    stockWarning = "The Stock display could not refresh. The balances below are read from saved transactions.";
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = getInventorySpreadsheet_();
    const catalog = familyInventoryCatalog_(spreadsheet);
    const history = familyInventoryHistory_(spreadsheet);
    const balances = familyInventoryBalances_(history);
    const members = familyInventoryMembers_(spreadsheet);
    const active = new Set(members.map(member => member.id));
    const itemIDs = new Set(catalog.items.map(item => item.id));
    const holdings = [];

    for (const [key, value] of balances) {
      const pair = JSON.parse(key);
      const memberEmail = familyInventoryMemberEmail_(pair[1]);
      const quantity = familyInventoryRounded_(value);
      if (!memberEmail || quantity === 0) continue;
      if (quantity < 0) {
        throw new Error("Issued stock is negative for " + pair[0] + " / " + memberEmail + ". Review Transactions.");
      }
      if (!itemIDs.has(pair[0])) {
        throw new Error("An issued item is missing from Items: " + pair[0]);
      }
      holdings.push({ itemID: pair[0], memberEmail: memberEmail, quantity: quantity });
    }
    holdings.sort((a, b) => a.memberEmail.localeCompare(b.memberEmail) ||
      a.itemID.localeCompare(b.itemID));

    const returnEmails = Array.from(new Set(holdings.map(holding => holding.memberEmail))).sort();
    return {
      items: catalog.items,
      storages: catalog.storages,
      members: members,
      returnMembers: returnEmails.map(email => ({
        id: email,
        label: email + (active.has(email) ? "" : " (not currently shared)"),
        active: active.has(email)
      })),
      balances: Array.from(balances, ([key, quantity]) => {
        const pair = JSON.parse(key);
        return { itemID: pair[0], storageID: pair[1], quantity: familyInventoryRounded_(quantity) };
      }),
      holdings: holdings,
      stockWarning: stockWarning
    };
  } finally {
    lock.releaseLock();
  }
}

function submitWebFamilyIssue(input) {
  return familyInventorySubmit_(input, "issue");
}

function submitWebFamilyReturn(input) {
  return familyInventorySubmit_(input, "return");
}

function familyInventorySubmit_(input, mode) {
  const prefix = mode === "issue" ? "ISSUE" : "RETURN";
  const requestID = String(input && input.requestID || "").trim();
  function rejected(message) {
    return { accepted: false, requestID: requestID, mode: mode, message: message };
  }
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    return rejected("No request was received.");
  }
  const uuid = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";
  if (!new RegExp("^" + prefix + "-" + uuid + "$").test(requestID)) {
    return rejected("Invalid request reference. Reload the page and try again.");
  }
  const itemID = String(input.itemID || "").trim();
  const storageID = String(input.storageID || "").trim();
  const memberEmail = String(input.memberEmail || "").trim().toLowerCase();
  const notes = String(input.notes || "").trim();
  if (!itemID || !storageID || !familyInventoryValidEmail_(memberEmail)) {
    return rejected("Select an item, storage and family member.");
  }
  if (!["number", "string"].includes(typeof input.quantity) ||
      String(input.quantity).trim() === "") {
    return rejected("Quantity must be a number greater than zero.");
  }
  const quantity = Number(input.quantity);
  if (!Number.isFinite(quantity) || quantity <= 0 || quantity > Number.MAX_SAFE_INTEGER) {
    return rejected("Quantity must be a supported number greater than zero.");
  }
  if (notes.length > 2000) {
    return rejected("Please keep Notes to 2,000 characters or fewer.");
  }

  const memberStorage = "MEMBER:" + memberEmail;
  const source = mode === "issue" ? storageID : memberStorage;
  const destination = mode === "issue" ? memberStorage : storageID;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  try {
    const spreadsheet = getInventorySpreadsheet_();
    const history = familyInventoryHistory_(spreadsheet);
    const duplicates = history.filter(row => String(row[0]).trim() === requestID);
    if (duplicates.length > 1) throw new Error("Duplicate transaction reference: " + requestID);

    if (duplicates.length) {
      const saved = duplicates[0];
      if (String(saved[2]).trim() !== "Move" ||
          String(saved[3]).trim() !== itemID ||
          Number(saved[4]) !== quantity ||
          String(saved[5] || "").trim() !== source ||
          String(saved[6] || "").trim() !== destination ||
          String(saved[7] || "").trim() !== notes) {
        return rejected("This request reference was already used for different details.");
      }
      result = {
        accepted: true, requestID: requestID, mode: mode,
        message: "Accepted: this " + (mode === "issue" ? "issue" : "return") +
          " was already saved. No duplicate recorded."
      };
    } else {
      const catalog = familyInventoryCatalog_(spreadsheet);
      const item = catalog.items.find(entry => entry.id === itemID);
      const storage = catalog.storages.find(entry => entry.id === storageID);
      if (!item) return rejected("The selected item no longer exists. Refresh the choices.");
      if (!storage) return rejected("The selected box or location no longer exists. Refresh the choices.");
      const quantityError = familyInventoryQuantityError_(quantity, item.unit);
      if (quantityError) return rejected(quantityError);

      if (mode === "issue" &&
          !familyInventoryMembers_(spreadsheet).some(member => member.id === memberEmail)) {
        return rejected("This account no longer has shared editing access. Refresh the family members.");
      }

      const balances = familyInventoryBalances_(history);
      const available = balances.get(JSON.stringify([itemID, source])) || 0;
      const tolerance = Number.EPSILON *
        Math.max(1, Math.abs(available), quantity) * 16;
      if (available <= 0 || quantity > available + tolerance) {
        return rejected(mode === "issue"
          ? "Insufficient stock at " + storage.label + ". Available: " +
            familyInventoryRounded_(Math.max(0, available)) + " " + item.unit + "."
          : memberEmail + " holds only " +
            familyInventoryRounded_(Math.max(0, available)) + " " + item.unit +
            " of this item. A return cannot exceed that quantity.");
      }

      // The existing ledger records both sides atomically in one row.
      // Distinct ISSUE/RETURN references identify these movements in history.
      spreadsheet.getSheetByName("Transactions").appendRow([
        requestID, new Date(), "Move", familyInventorySafeText_(itemID),
        quantity, familyInventorySafeText_(source),
        familyInventorySafeText_(destination), familyInventorySafeText_(notes)
      ]);
      SpreadsheetApp.flush();

      const remaining = familyInventoryRounded_(Math.max(0, available - quantity));
      result = {
        accepted: true, requestID: requestID, mode: mode,
        message: mode === "issue"
          ? "Issued " + quantity + " " + item.unit + " of " + item.label +
            " to " + memberEmail + " from " + storage.label +
            ". Remaining at source: " + remaining + " " + item.unit + "."
          : "Returned " + quantity + " " + item.unit + " of " + item.label +
            " from " + memberEmail + " to " + storage.label +
            ". Still with this member: " + remaining + " " + item.unit + "."
      };
    }
  } finally {
    lock.releaseLock();
  }

  // refreshStock acquires its own script lock.
  try {
    refreshStock();
    result.stockRefreshed = true;
  } catch (error) {
    console.error(error);
    result.stockRefreshed = false;
    result.message += " The movement was saved, but the Stock display could not refresh." +
      " Use Refresh balances to try again.";
  }
  return result;
}

function familyInventoryHistory_(spreadsheet) {
  const transactions = spreadsheet.getSheetByName("Transactions");
  if (!transactions) throw new Error("Missing Transactions sheet.");
  const expected = [
    "TransactionID", "Timestamp", "Action", "ItemID",
    "Quantity", "FromStorageID", "ToStorageID", "Notes"
  ];
  const headers = transactions.getRange(1, 1, 1, 8).getDisplayValues()[0];
  if (!expected.every((name, index) => String(headers[index] || "").trim() === name)) {
    throw new Error("Transactions column headers do not match.");
  }
  const history = transactions.getDataRange().getValues().slice(1)
    .filter(row => String(row[0] || "").trim() !== "");
  const ids = new Set();
  for (const row of history) {
    const id = String(row[0]).trim();
    if (ids.has(id)) throw new Error("Duplicate transaction reference: " + id);
    ids.add(id);
  }
  return history;
}

function familyInventoryBalances_(history) {
  const balances = new Map();
  function change(itemID, storageID, quantity) {
    if (!storageID) return;
    const key = JSON.stringify([itemID, storageID]);
    const next = (balances.get(key) || 0) + quantity;
    if (!Number.isFinite(next)) throw new Error("The inventory balance exceeds the supported range.");
    balances.set(key, next);
  }
  for (const row of history) {
    const quantity = Number(row[4]);
    if (!Number.isFinite(quantity) || quantity <= 0) {
      throw new Error("Invalid quantity in transaction: " + row[0]);
    }
    const itemID = String(row[3]).trim();
    change(itemID, String(row[5] || "").trim(), -quantity);
    change(itemID, String(row[6] || "").trim(), quantity);
  }
  return balances;
}

function familyInventoryCatalog_(spreadsheet) {
  const options = getMoveInventoryOptions();
  if (!options || !Array.isArray(options.items)) {
    throw new Error("The inventory choices could not be read.");
  }
  const storages = inventoryStorageTree_(spreadsheet).choices();
  const seen = new Set();
  const items = options.items.map(item => {
    if (!item || typeof item.id !== "string" || !item.id.trim() ||
        typeof item.label !== "string" || typeof item.unit !== "string" ||
        seen.has(item.id.trim())) {
      throw new Error("An item choice is missing or duplicated.");
    }
    const id = item.id.trim();
    seen.add(id);
    return { id: id, label: item.label, unit: item.unit.trim() };
  });
  // Always offer actual storage records, independently of other forms' choices.
  storages.sort((a, b) => a.id.localeCompare(b.id));
  return { items: items, storages: storages };
}

function familyInventoryMembers_(spreadsheet) {
  const users = spreadsheet.getEditors();
  const owner = spreadsheet.getOwner();
  if (owner) users.push(owner);
  const emails = new Set();
  for (const user of users) {
    const email = String(user.getEmail() || "").trim().toLowerCase();
    if (familyInventoryValidEmail_(email)) emails.add(email);
  }
  return Array.from(emails).sort().map(email => ({ id: email, label: email, active: true }));
}

function familyInventoryValidEmail_(email) {
  return typeof email === "string" && email.length <= 254 &&
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function familyInventoryMemberEmail_(storageID) {
  if (!String(storageID).startsWith("MEMBER:")) return "";
  const email = String(storageID).slice(7);
  if (!familyInventoryValidEmail_(email) || email !== email.trim().toLowerCase()) {
    throw new Error("Invalid family-member storage ID: " + storageID);
  }
  return email;
}

function familyInventoryQuantityError_(quantity, unit) {
  const wholeOnly = [
    "each", "ea", "pack", "packs", "piece", "pieces", "pcs"
  ].includes(unit.trim().toLowerCase());
  if (wholeOnly) {
    return Number.isSafeInteger(quantity) ? "" : unit + " requires a whole-number quantity.";
  }
  const scaled = quantity * 1000;
  const tolerance = Number.EPSILON * Math.max(1, Math.abs(scaled)) * 16;
  if (quantity < 0.001 || !Number.isSafeInteger(Math.round(scaled)) ||
      Math.abs(scaled - Math.round(scaled)) > tolerance) {
    return "Enter at least 0.001, with up to 3 decimal places.";
  }
  return "";
}

function familyInventoryRounded_(quantity) {
  return Math.round(quantity * 1000000) / 1000000;
}

function familyInventorySafeText_(value) {
  return /^[=+\-@']/.test(value) ? "'" + value : value;
}

function getFamilyItemThumbnail(itemID) {
  try {
    const spreadsheet = getInventorySpreadsheet_();
    const sheet = spreadsheet.getSheetByName("Items");
    if (!sheet) return "";
    const matches = sheet.getDataRange().getValues().slice(1)
      .filter(row => String(row[0]).trim() === String(itemID || "").trim());
    if (matches.length !== 1) return "";
    const url = String(matches[0][6] || "").trim();
    if (!/^https:\/\/drive\.google\.com\//i.test(url)) return "";
    const match = url.match(/\/d\/([A-Za-z0-9_-]+)/) ||
      url.match(/[?&]id=([A-Za-z0-9_-]+)/);
    if (!match) return "";

    const folderID = PropertiesService.getScriptProperties()
      .getProperty("INVENTORY_PHOTO_FOLDER_ID");
    if (!folderID) return "";
    const file = DriveApp.getFileById(match[1]);
    const parents = file.getParents();
    let inPhotoFolder = false;
    while (parents.hasNext()) {
      if (parents.next().getId() === folderID) inPhotoFolder = true;
    }
    if (!inPhotoFolder) return "";

    const blob = file.getThumbnail() || file.getBlob();
    const mime = blob.getContentType();
    const bytes = blob.getBytes();
    if (!/^image\/(jpeg|png|webp|gif)$/.test(mime) || bytes.length > 1048576) return "";
    return "data:" + mime + ";base64," + Utilities.base64Encode(bytes);
  } catch (error) {
    console.warn(error);
    return "";
  }
}