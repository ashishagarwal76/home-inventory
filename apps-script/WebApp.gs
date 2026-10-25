function getInventorySpreadsheet_() {
  return SpreadsheetApp.openById(inventoryConfiguredSpreadsheetID_());
}

function testInventoryConnection() {
  const spreadsheet = getInventorySpreadsheet_();

  const requiredSheets = [
    "Items",
    "Boxes",
    "Locations",
    "Transactions",
    "Stock"
  ];

  const missing = requiredSheets.filter(
    name => !spreadsheet.getSheetByName(name)
  );

  if (missing.length) {
    throw new Error("Missing tabs: " + missing.join(", "));
  }

  console.log(
    "Connected to " + spreadsheet.getName() +
    ". All required tabs found."
  );
}

//getAddInventoryOptions

function testAddInventoryOptions() {
  const options = getAddInventoryOptions();

  console.log("Items loaded: " + options.items.length);
  console.log("Storage choices loaded: " + options.destinations.length);
}

function submitWebAddInventory(input) {
  if (!input || typeof input !== "object") {
    return { accepted: false, message: "No request was received." };
  }

  const requestID = String(input.requestID || "").trim();
  const itemID = String(input.itemID || "").trim();
  const destination = String(input.destination || "").trim();
  const quantity = Number(input.quantity);
  const notes = String(input.notes || "").trim();

  // The page will generate one unique ID per submission.
  if (!/^WEB-[0-9a-f-]{36}$/i.test(requestID)) {
    return {
      accepted: false,
      message: "Invalid request ID. Reload the page and try again."
    };
  }

  if (!Number.isFinite(quantity) || quantity <= 0) {
    return {
      accepted: false,
      message: "Quantity must be a number greater than zero."
    };
  }

  if (notes.length > 2000) {
    return {
      accepted: false,
      message: "Please keep Notes to 2,000 characters or fewer."
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let result;

  try {
    const spreadsheet = getInventorySpreadsheet_();
    const transactions = spreadsheet.getSheetByName("Transactions");

    if (!transactions) {
      throw new Error("Missing Transactions sheet.");
    }

    const expectedHeaders = [
      "TransactionID", "Timestamp", "Action", "ItemID",
      "Quantity", "FromStorageID", "ToStorageID", "Notes"
    ];

    const headers = transactions
      .getRange(1, 1, 1, 8).getDisplayValues()[0];

    if (!expectedHeaders.every((name, index) =>
      headers[index].trim() === name
    )) {
      throw new Error("Transactions column headers do not match.");
    }

    const history = transactions.getDataRange().getValues().slice(1);
    const existing = history.find(row => row[0] === requestID);

    if (existing) {
      // A retry returns the original result without adding stock again.
      if (
        existing[2] !== "Add" ||
        String(existing[3]).trim() !== itemID ||
        Number(existing[4]) !== quantity ||
        String(existing[6]).trim() !== destination
      ) {
        return {
          accepted: false,
          message: "This request ID was already used for different details."
        };
      }

      result = {
        accepted: true,
        requestID,
        message: "Accepted: this addition was already saved. No duplicate added."
      };
    } else {
      // Reload current catalogs while holding the shared inventory lock.
      const options = getAddInventoryOptions();
      const item = options.items.find(entry => entry.id === itemID);
      const storage = options.destinations.find(
        entry => entry.id === destination
      );

      if (!item) {
        return {
          accepted: false,
          message: "The selected item no longer exists. Refresh the choices."
        };
      }
const wholeNumberUnits = [
  "each", "pack"
];

const wholeOnly = wholeNumberUnits.includes(
  item.unit.trim().toLowerCase()
);

if (wholeOnly && !Number.isInteger(quantity)) {
  return {
    accepted: false,
    message: item.unit + " requires a whole-number quantity."
  };
}

if (!wholeOnly) {
  const scaled = quantity * 1000;
  const tolerance =
    Number.EPSILON * Math.max(1, Math.abs(scaled)) * 16;

  if (
    quantity < 0.001 ||
    Math.abs(scaled - Math.round(scaled)) > tolerance
  ) {
    return {
      accepted: false,
      message: "Enter at least 0.001, with up to 3 decimal places."
    };
  }
}
      if (!storage) {
        return {
          accepted: false,
          message: "The selected destination no longer exists. Refresh the choices."
        };
      }

      const safeNotes = /^[=+\-@]/.test(notes) ? "'" + notes : notes;

      transactions.appendRow([
        requestID,
        new Date(),
        "Add",
        itemID,
        quantity,
        "",
        destination,
        safeNotes
      ]);

      SpreadsheetApp.flush();

      result = {
        accepted: true,
        requestID,
        message: "Accepted: added " + quantity + " " + item.unit +
          " of " + item.label + " to " + storage.label + "."
      };
    }
  } finally {
    lock.releaseLock();
  }

  // Refresh separately because refreshStock takes its own lock.
  try {
    refreshStock();
  } catch (error) {
    console.error(error);
    result.message +=
      " Inventory was saved, but the Stock display could not refresh." +
      " Run refreshStock from Apps Script; do not submit a new addition.";
  }

  return result;
}
function doGet(e) {
  const page = e && e.parameter
    ? String(e.parameter.page || "").trim().toLowerCase()
    : "";

  const pages = {
    scan: {
      file: "ScanLabel",
      title: "Home Inventory – Scan QR Label",
      template: true
    },
    labels: {
      file: "PrintLabels",
      title: "Home Inventory – Print QR Labels"
    },
    home: {
      file: "Home",
      title: "Home Inventory",
      template: true
    },
    create: {
      file: "CreateRecords",
      title: "Home Inventory – Create Item, Box or Location",
      template: true
    },
    family: {
      file: "FamilyInventory",
      title: "Home Inventory – Issue / Return Items"
    },
    find: {
      file: "FindInventory",
      title: "Home Inventory – Find Inventory"
    },
    add: {
      file: "AddInventory",
      title: "Home Inventory – Add"
    },
    consume: {
      file: "ConsumeInventory",
      title: "Home Inventory – Consume"
    },
    move: {
      file: "MoveInventory",
      title: "Home Inventory – Move"
    },
    movebox: {
      file: "MoveBox",
      title: "Home Inventory – Move Box"
    },
    photo: {
      file: "UploadPhoto",
      title: "Home Inventory – Upload Photo"
    },
    newbox: {
      file: "NewBox",
      title: "Home Inventory – New Box"
    },
    newlocation: {
      file: "NewLocation",
      title: "Home Inventory – New Location"
    },
    newitem: {
      file: "NewItem",
      title: "Home Inventory – New Item"
    }
  };

  const selected = Object.prototype.hasOwnProperty.call(pages, page)
    ? pages[page]
    : pages.home;

  const appUrl = ScriptApp.getService().getUrl();
  if (!appUrl) {
    throw new Error("Deploy the web app before opening its pages.");
  }

  let output;
  if (selected.template) {
    const template = HtmlService.createTemplateFromFile(selected.file);
    template.appUrl = appUrl;
    if (selected.file === "ScanLabel") template.scannerUrl = inventoryScannerUrl_();
    output = template.evaluate();
  } else {
    output = HtmlService.createHtmlOutputFromFile(selected.file);
  }

  // Other pages link to Home; creation forms also link to their parent menu.
  function navigationUrl(route) {
    return (appUrl + "?page=" + route)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  const linkStyle =
    'display:inline-flex;align-items:center;gap:8px;min-height:44px;' +
    'max-width:100%;box-sizing:border-box;padding:10px 14px;' +
    'border:1px solid #b7c9d1;border-radius:8px;background:#fff;color:#165d78;' +
    'font-size:16px;font-weight:600;line-height:1.4;text-decoration:none;';
  const backLink = ["NewItem", "NewBox", "NewLocation"].includes(selected.file)
    ? '<a id="inventoryCreateBackLink" href="' + navigationUrl("create") +
      '" target="_top" aria-label="Back to Create Item, Box or Location" ' +
      'style="' + linkStyle + '">' +
      '<span aria-hidden="true">←</span> Back to Create</a>'
    : '';
  const navigation = selected.file === "Home" ? '' :
    '<nav id="inventoryHomeNavigation" aria-label="Main navigation" ' +
      'style="display:flex;flex-wrap:wrap;gap:8px;max-width:520px;' +
        'margin:0 auto 8px;font-family:system-ui,sans-serif;">' +
      backLink +
      '<a href="' + navigationUrl("home") + '" target="_top" ' +
        'style="' + linkStyle + '">' +
        '<span aria-hidden="true">←</span> Home' +
      '</a>' +
    '</nav>';

  if (navigation) {
    const content = output.getContent();
    const bodyTag = /<body\b[^>]*>/i;
    output.setContent(bodyTag.test(content)
      ? content.replace(bodyTag, tag => tag + "\n" + navigation)
      : navigation + content);
  }
  if ([
    "AddInventory",
    "ConsumeInventory",
    "MoveInventory",
    "FamilyInventory",
    "MoveBox"
  ].includes(selected.file)) {
    const prefill = HtmlService
      .createHtmlOutputFromFile("ScanPrefill")
      .getContent();

    const content = output.getContent();

    output.setContent(
      /<\/body\s*>/i.test(content)
        ? content.replace(
            /<\/body\s*>/i,
            () => prefill + "\n</body>"
          )
        : content + "\n" + prefill
    );
  }
  return output
    .setTitle(selected.title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}



function getConsumeInventoryOptions() {
  const options = getAddInventoryOptions();

  return {
    items: options.items,
    sources: options.destinations
  };
}

function testConsumeInventoryOptions() {
  const options = getConsumeInventoryOptions();

  console.log("Items loaded: " + options.items.length);
  console.log("Source choices loaded: " + options.sources.length);
}


function submitWebConsumeInventory(input) {
  if (!input || typeof input !== "object") {
    return { accepted: false, message: "No request was received." };
  }

  const requestID = String(input.requestID || "").trim();
  const itemID = String(input.itemID || "").trim();
  const source = String(input.source || "").trim();
  const quantity = Number(input.quantity);
  const notes = String(input.notes || "").trim();

  if (!/^WEB-[0-9a-f-]{36}$/i.test(requestID)) {
    return {
      accepted: false,
      message: "Invalid request ID. Reload the page and try again."
    };
  }

  if (!Number.isFinite(quantity) || quantity <= 0) {
    return {
      accepted: false,
      message: "Quantity must be a number greater than zero."
    };
  }

  if (notes.length > 2000) {
    return {
      accepted: false,
      message: "Please keep Notes to 2,000 characters or fewer."
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let result;

  try {
    const spreadsheet = getInventorySpreadsheet_();
    const transactions = spreadsheet.getSheetByName("Transactions");

    if (!transactions) {
      throw new Error("Missing Transactions sheet.");
    }

    const expectedHeaders = [
      "TransactionID", "Timestamp", "Action", "ItemID",
      "Quantity", "FromStorageID", "ToStorageID", "Notes"
    ];

    const headers = transactions
      .getRange(1, 1, 1, 8).getDisplayValues()[0];

    if (!expectedHeaders.every((name, index) =>
      headers[index].trim() === name
    )) {
      throw new Error("Transactions column headers do not match.");
    }

    const history = transactions.getDataRange().getValues().slice(1);
    const existing = history.find(row => row[0] === requestID);

    if (existing) {
      // A retry must not consume the same quantity twice.
      if (
        existing[2] !== "Consume" ||
        String(existing[3]).trim() !== itemID ||
        Number(existing[4]) !== quantity ||
        String(existing[5]).trim() !== source
      ) {
        return {
          accepted: false,
          message: "This request ID was already used for different details."
        };
      }

      result = {
        accepted: true,
        requestID,
        message: "Accepted: this consumption was already saved. No duplicate recorded."
      };
    } else {
      const options = getConsumeInventoryOptions();
      const item = options.items.find(entry => entry.id === itemID);
      const storage = options.sources.find(entry => entry.id === source);

      if (!item) {
        return {
          accepted: false,
          message: "The selected item no longer exists. Refresh the choices."
        };
      }

      if (!storage) {
        return {
          accepted: false,
          message: "The selected source no longer exists. Refresh the choices."
        };
      }

      const wholeNumberUnits = [
        "each", "ea", "pack", "packs",
        "piece", "pieces", "pcs"
      ];

      const wholeOnly = wholeNumberUnits.includes(
        item.unit.trim().toLowerCase()
      );

      if (wholeOnly && !Number.isInteger(quantity)) {
        return {
          accepted: false,
          message: item.unit + " requires a whole-number quantity."
        };
      }

      if (!wholeOnly) {
        const scaled = quantity * 1000;
        const tolerance =
          Number.EPSILON * Math.max(1, Math.abs(scaled)) * 16;

        if (
          quantity < 0.001 ||
          Math.abs(scaled - Math.round(scaled)) > tolerance
        ) {
          return {
            accepted: false,
            message: "Enter at least 0.001, with up to 3 decimal places."
          };
        }
      }

      // Read the transaction history, not the Stock display.
      // The shared lock keeps validation and posting together.
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

      const stockTolerance = Number.EPSILON *
        Math.max(1, Math.abs(available), quantity) * 16;

      if (available <= 0 || quantity > available + stockTolerance) {
        return {
          accepted: false,
          message: "Insufficient stock at " + storage.label +
            ". Available: " + Number(available.toFixed(6)) +
            " " + item.unit + "; requested: " + quantity + "."
        };
      }

      const safeNotes = /^[=+\-@]/.test(notes) ? "'" + notes : notes;

      transactions.appendRow([
        requestID,
        new Date(),
        "Consume",
        itemID,
        quantity,
        source,
        "",
        safeNotes
      ]);

      SpreadsheetApp.flush();

      const remaining = Number(
        Math.max(0, available - quantity).toFixed(6)
      );

      result = {
        accepted: true,
        requestID,
        message: "Accepted: consumed " + quantity + " " + item.unit +
          " of " + item.label + " from " + storage.label +
          ". Remaining at this source: " + remaining +
          " " + item.unit + "."
      };
    }
  } finally {
    lock.releaseLock();
  }

  // refreshStock acquires its own lock.
  try {
    refreshStock();
  } catch (error) {
    console.error(error);
    result.message +=
      " Consumption was saved, but the Stock display could not refresh." +
      " Run refreshStock from Apps Script; do not submit a new consumption.";
  }

  return result;
}

function getMoveInventoryOptions() {
  const options = getAddInventoryOptions();

  return {
    items: options.items,
    sources: options.destinations
  };
}

function testMoveInventoryOptions() {
  const options = getMoveInventoryOptions();

  console.log("Items loaded: " + options.items.length);
  console.log("Storage choices loaded: " + options.sources.length);
}

function submitWebMoveInventory(input) {
  if (!input || typeof input !== "object") {
    return { accepted: false, message: "No request was received." };
  }

  const requestID = String(input.requestID || "").trim();
  const itemID = String(input.itemID || "").trim();
  const source = String(input.source || "").trim();
  const destination = String(input.destination || "").trim();
  const quantity = Number(input.quantity);
  const notes = String(input.notes || "").trim();

  if (!/^WEB-[0-9a-f-]{36}$/i.test(requestID)) {
    return {
      accepted: false,
      message: "Invalid request ID. Reload the page and try again."
    };
  }

  if (!Number.isFinite(quantity) || quantity <= 0) {
    return {
      accepted: false,
      message: "Quantity must be a number greater than zero."
    };
  }

  if (!source || !destination) {
    return {
      accepted: false,
      message: "Select both a source and destination."
    };
  }

  if (source === destination) {
    return {
      accepted: false,
      message: "Source and destination must be different."
    };
  }

  if (notes.length > 2000) {
    return {
      accepted: false,
      message: "Please keep Notes to 2,000 characters or fewer."
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let result;

  try {
    const spreadsheet = getInventorySpreadsheet_();
    const transactions = spreadsheet.getSheetByName("Transactions");

    if (!transactions) {
      throw new Error("Missing Transactions sheet.");
    }

    const expectedHeaders = [
      "TransactionID", "Timestamp", "Action", "ItemID",
      "Quantity", "FromStorageID", "ToStorageID", "Notes"
    ];

    const headers = transactions
      .getRange(1, 1, 1, 8).getDisplayValues()[0];

    if (!expectedHeaders.every((name, index) =>
      headers[index].trim() === name
    )) {
      throw new Error("Transactions column headers do not match.");
    }

    const history = transactions.getDataRange().getValues().slice(1);
    const existing = history.find(row => row[0] === requestID);

    if (existing) {
      // A retry must not move the same quantity twice.
      if (
        existing[2] !== "Move" ||
        String(existing[3]).trim() !== itemID ||
        Number(existing[4]) !== quantity ||
        String(existing[5]).trim() !== source ||
        String(existing[6]).trim() !== destination
      ) {
        return {
          accepted: false,
          message: "This request ID was already used for different details."
        };
      }

      result = {
        accepted: true,
        requestID,
        message: "Accepted: this movement was already saved. No duplicate recorded."
      };
    } else {
      const options = getMoveInventoryOptions();
      const item = options.items.find(entry => entry.id === itemID);
      const fromStorage = options.sources.find(
        entry => entry.id === source
      );
      const toStorage = options.sources.find(
        entry => entry.id === destination
      );

      if (!item) {
        return {
          accepted: false,
          message: "The selected item no longer exists. Refresh the choices."
        };
      }

      if (!fromStorage || !toStorage) {
        return {
          accepted: false,
          message: "The source or destination no longer exists. Refresh the choices."
        };
      }

      const wholeNumberUnits = [
        "each", "ea", "pack", "packs",
        "piece", "pieces", "pcs"
      ];

      const wholeOnly = wholeNumberUnits.includes(
        item.unit.trim().toLowerCase()
      );

      if (wholeOnly && !Number.isInteger(quantity)) {
        return {
          accepted: false,
          message: item.unit + " requires a whole-number quantity."
        };
      }

      if (!wholeOnly) {
        const scaled = quantity * 1000;
        const tolerance =
          Number.EPSILON * Math.max(1, Math.abs(scaled)) * 16;

        if (
          quantity < 0.001 ||
          Math.abs(scaled - Math.round(scaled)) > tolerance
        ) {
          return {
            accepted: false,
            message: "Enter at least 0.001, with up to 3 decimal places."
          };
        }
      }

      // Validate against transaction history while holding the lock.
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

      const stockTolerance = Number.EPSILON *
        Math.max(1, Math.abs(available), quantity) * 16;

      if (available <= 0 || quantity > available + stockTolerance) {
        return {
          accepted: false,
          message: "Insufficient stock at " + fromStorage.label +
            ". Available: " + Number(available.toFixed(6)) +
            " " + item.unit + "; requested: " + quantity + "."
        };
      }

      const safeNotes = /^[=+\-@]/.test(notes) ? "'" + notes : notes;

      // One transaction records both sides of the movement.
      transactions.appendRow([
        requestID,
        new Date(),
        "Move",
        itemID,
        quantity,
        source,
        destination,
        safeNotes
      ]);

      SpreadsheetApp.flush();

      const remaining = Number(
        Math.max(0, available - quantity).toFixed(6)
      );

      result = {
        accepted: true,
        requestID,
        message: "Accepted: moved " + quantity + " " + item.unit +
          " of " + item.label +
          " from " + fromStorage.label +
          " to " + toStorage.label +
          ". Remaining at source: " + remaining +
          " " + item.unit + "."
      };
    }
  } finally {
    lock.releaseLock();
  }

  try {
    refreshStock();
  } catch (error) {
    console.error(error);
    result.message +=
      " Movement was saved, but the Stock display could not refresh." +
      " Run refreshStock from Apps Script; do not submit a new movement.";
  }

  return result;
}

//getMoveBoxOptions 

function testMoveBoxOptions() {
  const options = getMoveBoxOptions();

  console.log("Boxes loaded: " + options.boxes.length);
  console.log("Locations loaded: " + options.locations.length);
}

//submitWebMoveBox

//getPhotoUploadOptions


function testPhotoUploadOptions() {
  const options = getPhotoUploadOptions();

  console.log("Items loaded: " + options.items.length);
  console.log("Boxes loaded: " + options.boxes.length);
}


function setupPhotoStorage() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const properties = PropertiesService.getScriptProperties();
    const savedID = properties.getProperty("INVENTORY_PHOTO_FOLDER_ID");

    let folder;

    if (savedID) {
      folder = DriveApp.getFolderById(savedID);

      if (folder.isTrashed()) {
        throw new Error(
          "The photo folder is in Trash. Restore it before continuing."
        );
      }
    } else {
      folder = DriveApp.createFolder("Home Inventory Photos");

      properties.setProperty(
        "INVENTORY_PHOTO_FOLDER_ID",
        folder.getId()
      );
    }

    console.log("Photo folder ready: " + folder.getUrl());
  } finally {
    lock.releaseLock();
  }
}

function uploadInventoryPhoto(input) {
  if (!input || typeof input !== "object") {
    return { accepted: false, message: "No photo was received." };
  }

  const requestID = String(input.requestID || "").trim();
  const recordType = String(input.recordType || "");
  const recordID = String(input.recordID || "").trim();
  const base64 = String(input.base64 || "");

  if (!/^PHOTO-[0-9a-f-]{36}$/i.test(requestID)) {
    return { accepted: false, message: "Invalid upload request ID." };
  }

  const settings = {
    Item: { sheet: "Items", idHeader: "ItemID" },
    Box: { sheet: "Boxes", idHeader: "BoxID" }
  };

  if (!Object.prototype.hasOwnProperty.call(settings, recordType)) {
    return { accepted: false, message: "Select Item or Box." };
  }

  // Limit the resized image to approximately 1 MB.
  if (!base64 || base64.length > 1400000) {
    return {
      accepted: false,
      message: "The photo is empty or too large. Choose a smaller image."
    };
  }

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (error) {
    return { accepted: false, message: "The photo data is invalid." };
  }

  if (
    bytes.length > 1048576 ||
    bytes.length < 4 ||
    (bytes[0] & 255) !== 255 ||
    (bytes[1] & 255) !== 216 ||
    (bytes[2] & 255) !== 255
  ) {
    return {
      accepted: false,
      message: "Upload a JPEG photo smaller than 1 MB."
    };
  }

  const fingerprint = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    bytes
  ).map(value => (value & 255).toString(16).padStart(2, "0")).join("");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const config = settings[recordType];
    const sheet = getInventorySpreadsheet_().getSheetByName(config.sheet);
    if (!sheet) throw new Error("Missing sheet: " + config.sheet);

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0].map(value => value.trim());
    const idColumn = headers.indexOf(config.idHeader);
    const photoColumn = headers.indexOf("PhotoURL");

    if (idColumn === -1 || photoColumn === -1) {
      throw new Error(
        config.sheet + " needs " + config.idHeader + " and PhotoURL."
      );
    }

    const matches = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColumn].trim() === recordID) matches.push(i + 1);
    }

    if (matches.length !== 1) {
      return {
        accepted: false,
        message: "The selected record is missing or duplicated."
      };
    }

    const folderID = PropertiesService.getScriptProperties()
      .getProperty("INVENTORY_PHOTO_FOLDER_ID");

    if (!folderID) {
      throw new Error("Run setupPhotoStorage first.");
    }

    const folder = DriveApp.getFolderById(folderID);
    if (folder.isTrashed()) {
      throw new Error("Restore the photo folder from Trash first.");
    }

    const photoCell = sheet.getRange(matches[0], photoColumn + 1);
    const currentURL = photoCell.getDisplayValue().trim();
    const filename = requestID + ".jpg";

    // Reuse the uploaded file if the same request is retried.
    const files = folder.getFilesByName(filename);
    let file;
    let metadata;

    if (files.hasNext()) {
      file = files.next();

      if (files.hasNext()) {
        throw new Error("Duplicate upload files found for this request.");
      }

      const description = file.getDescription();

      if (description) {
        metadata = JSON.parse(description);

        if (
          metadata.recordType !== recordType ||
          metadata.recordID !== recordID ||
          metadata.fingerprint !== fingerprint
        ) {
          return {
            accepted: false,
            message: "This upload request was already used for different details."
          };
        }
      } else {
        // Recover a file created before its metadata could be saved.
        const savedFingerprint = Utilities.computeDigest(
          Utilities.DigestAlgorithm.SHA_256,
          file.getBlob().getBytes()
        ).map(value =>
          (value & 255).toString(16).padStart(2, "0")
        ).join("");

        if (savedFingerprint !== fingerprint) {
          throw new Error("The saved upload does not match this photo.");
        }
      }
    } else {
      file = folder.createFile(
        Utilities.newBlob(bytes, "image/jpeg", filename)
      );
    }

    if (!metadata) {
      metadata = {
        recordType,
        recordID,
        fingerprint,
        previousURL: currentURL,
        saved: false
      };
      file.setDescription(JSON.stringify(metadata));
    }

    const photoURL = file.getUrl();

    if (metadata.saved) {
      return {
        accepted: true,
        requestID,
        photoURL,
        message: currentURL === photoURL
          ? "Photo already saved for " + recordID + "."
          : "This upload was saved earlier. A newer photo is now assigned."
      };
    }

    if (currentURL !== photoURL) {
      if (currentURL !== metadata.previousURL) {
        return {
          accepted: false,
          message: "Another photo was assigned during this upload." +
            " Refresh and start a new upload if you want to replace it."
        };
      }

      photoCell.setValue(photoURL);
      SpreadsheetApp.flush();
    }

    metadata.saved = true;
    file.setDescription(JSON.stringify(metadata));

    return {
      accepted: true,
      requestID,
      photoURL,
      message: "Photo saved for " + recordID + "."
    };
  } finally {
    lock.releaseLock();
  }
}

function getInventoryThumbnail(recordType, recordID) {
  const settings = {
    Item: { sheet: "Items", idHeader: "ItemID" },
    Box: { sheet: "Boxes", idHeader: "BoxID" }
  };

  if (!Object.prototype.hasOwnProperty.call(settings, recordType)) {
    throw new Error("Select Item or Box.");
  }

  const config = settings[recordType];
  const sheet = getInventorySpreadsheet_().getSheetByName(config.sheet);

  if (!sheet) throw new Error("Missing sheet: " + config.sheet);

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0].map(value => value.trim());
  const idColumn = headers.indexOf(config.idHeader);
  const photoColumn = headers.indexOf("PhotoURL");

  if (idColumn === -1 || photoColumn === -1) {
    throw new Error("Missing ID or PhotoURL column in " + config.sheet);
  }

  const matches = data.slice(1).filter(
    row => row[idColumn].trim() === String(recordID).trim()
  );

  if (matches.length !== 1) {
    throw new Error("Record is missing or duplicated.");
  }

  const photoURL = matches[0][photoColumn].trim();
  if (!photoURL) return { hasPhoto: false };

  // Accept Google Drive file links only.
  const match = photoURL.match(
    /^https:\/\/drive\.google\.com\/file\/d\/([A-Za-z0-9_-]+)(?:\/|$|\?)/
  );

  if (!match) {
    throw new Error("PhotoURL must be a Google Drive file link.");
  }

  const file = DriveApp.getFileById(match[1]);

  if (file.isTrashed()) {
    throw new Error("This photo is in Drive Trash.");
  }

  const mimeType = file.getMimeType();

  if (!["image/jpeg", "image/png"].includes(mimeType)) {
    throw new Error("The photo must be JPEG or PNG.");
  }

  if (file.getSize() > 1048576) {
    throw new Error("The photo is too large. Upload it through the photo page.");
  }

  return {
    hasPhoto: true,
    dataURL: "data:" + mimeType + ";base64," +
      Utilities.base64Encode(file.getBlob().getBytes())
  };
}
function getPrintLabelRecords() {
  return getLabelRecords_();
}
function lookupScannedInventoryRecord(recordID) {
  if (typeof recordID !== "string") {
    throw new Error("Scan an inventory QR label.");
  }

  const id = recordID.trim();

  if (!id || id.length > 100) {
    throw new Error("This is not a valid inventory label.");
  }

  const record = getLabelRecords_().find(function (entry) {
    return entry.id === id;
  });

  if (!record) {
    throw new Error("No inventory record matches this label.");
  }
  if (record.type === "Items") {
    const report = searchInventory(record.id);

    record.availability = report.items.find(function (item) {
      return item.id === record.id;
    });

    if (!record.availability) {
      throw new Error("Could not load this item's stock.");
    }
  }
  if (record.type === "Boxes") {
    record.boxDetails = getScannedBoxContents_(record.id);
  }
  return record;
}
//submitWebNewItem
