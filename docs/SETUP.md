# Set up your own inventory app

Use a new spreadsheet for your first installation. You need a Google account with permission to use Sheets, Drive, and Apps Script. A Cloudflare account is needed only for live camera scanning; entering record IDs works without it.

## 1. Create the spreadsheet

Upload `template/HomeInventory-template.xlsx` to your Google Drive, open it with Google Sheets, and save/convert it as a **Google Sheets** file. The Excel file supplies blank tabs and headers; it does not contain Apps Script code.

You may instead start with a new blank Google Sheet. The setup helper creates the same required tabs.

In your new Google Sheet, choose **Extensions → Apps Script**. This creates a script bound to that spreadsheet. Keep this relationship: setup and the optional Forms helpers rely on it. See Google's [bound script guide](https://developers.google.com/apps-script/guides/bound#create_a_bound_script).

## 2. Copy the app files

Copy every file from `apps-script/` into that bound project, keeping each filename and its contents. For a `.gs` file, choose **+ → Script**; for a `.html` file, choose **+ → HTML**. Enter the base name without adding a second extension. Replace the starter `Code.gs` contents with the supplied `Code.gs`.

There are **14 script files**, **16 HTML files**, and **1 manifest**:

| Script files | HTML files |
|---|---|
| BoxHierarchy.gs | AddInventory.html |
| BoxLookup.gs | BoxItemPhotos.html |
| Code.gs | ConsumeInventory.html |
| Configuration.gs | CreateRecords.html |
| CreationTransactions.gs | FamilyInventory.html |
| FamilyInventoryWeb.gs | FindInventory.html |
| FindInventoryWeb.gs | Home.html |
| InventoryOptions.gs | MoveBox.html |
| MoveBoxWeb.gs | MoveInventory.html |
| NewItemWeb.gs | NewBox.html |
| NewStorageWeb.gs | NewItem.html |
| Schema.gs | NewLocation.html |
| Setup.gs | PrintLabels.html |
| WebApp.gs | ScanLabel.html |
|  | ScanPrefill.html |
|  | UploadPhoto.html |

In **Project Settings**, enable showing `appsscript.json` in the editor. Replace its contents with the supplied manifest. Set the project time zone to your preferred time zone and use the same time zone in the spreadsheet's settings. Save the project.

Install all files before running any helper. `BoxItemPhotos` and `ScanPrefill` are included script fragments even though they have no menu page of their own.

For file management details, see Google's [script editor guide](https://developers.google.com/apps-script/guides/projects#manage_files_in_a_project). Experienced users may use [clasp](https://github.com/google/clasp) to synchronize an existing bound project; manual copying is the standard path here, and a clasp sign-in is not required.

## 3. Initialize this installation

From the Apps Script editor's function selector, choose **setupInventory** and click **Run**. Use the Google account that owns this installation and review the permissions requested by your copy of the code.

The helper:

- Checks existing core headers before creating or filling missing blank tabs.
- Sets `INVENTORY_SPREADSHEET_ID` in this project's Script Properties to the bound spreadsheet's ID.
- Creates or reuses this installation's **Home Inventory Photos** Drive folder through `setupPhotoStorage()`.
- Does not install triggers or configure Google Forms.

Setup may be run again on a compatible workbook without clearing its records. If it reports a conflicting configured spreadsheet or incompatible header, stop and inspect the named setting/tab. Do not remove records to force setup to pass.

The nine core tabs are:

| Tab | Holds |
|---|---|
| Items | Item names, units, minimum stock, notes, and photo links |
| Boxes | Box names, placements, parent box IDs, notes, and photo links |
| Locations | Storage location names |
| Transactions | Stock changes and family issue/return movements |
| Stock | Derived current quantities |
| BoxMoves | Box placement history and recovery state |
| NewItemRequests | Item creation and initial-stock request state |
| NewLocationRequests | Location creation request state |
| NewBoxRequests | Box creation and initial-placement request state |

Leave their names and required headers unchanged. The exact schema is in `apps-script/Schema.gs`. Stock's seven columns are `ItemID`, `ItemName`, `StorageID`, `BoxName`, `LocationName`, `Quantity`, and `Unit`; quantities are calculated from the transaction ledger.

`setupPhotoStorage()` writes the folder URL to the execution log. Save that URL for sharing the folder in the next step.

## 4. Choose who can use the installation

For a personal installation, keep the spreadsheet and photo folder private to your Google account.

For shared use, share **your spreadsheet and your Home Inventory Photos folder** with the intended Google accounts. Give trusted users editing permission to both if they will create records, change stock, or upload photos. Family-member choices are drawn from the spreadsheet's owner and editors. Spreadsheet editors can also access its bound script, so these should be people you trust to manage the installation.

The supplied manifest uses `USER_ACCESSING` and `ANYONE`: the app runs as the signed-in user opening it, and `ANYONE` means any logged-in user rather than anonymous access. Each user must authorize the app and have the necessary access to your files. This setting does not itself publish your spreadsheet or Drive photos. Google's [manifest reference](https://developers.google.com/apps-script/manifest/web-app-api-executable#webapp) defines these values.

Keep inventory records, photos, credentials, and actual project configuration out of GitHub. A public scanner page contains no inventory catalog and does not grant access to the spreadsheet.

## 5. Deploy the web app

In Apps Script, choose **Deploy → New deployment → Web app**. Select **User accessing the web app** for the execution identity. For access, use the signed-in Google-account option corresponding to the supplied manifest; Google Workspace policies can limit the options available. Do not change this installation to run as the deployer as a shortcut for file sharing.

Deploy and copy the web app URL ending in **`/exec`**. Open that URL while signed in as an intended user. Each new user may be asked to authorize access. See Google's [web app deployment and permission guide](https://developers.google.com/apps-script/guides/web#deploy_a_script_as_a_web_app).

No installable trigger is needed for the core web app. The existing `onEdit(e)` function is a simple sheet-edit trigger; do not run it manually. The optional Google Forms trigger is described separately in [Optional Forms](OPTIONAL_FORMS.md).

## 6. Add your first records

Open the deployed app and choose **Create Item, Box or Location**. Start with a location, then a box and an item. An item can have initial stock in a selected placement and an optional cropped photo. Existing item or box photos can be added through **Upload Photo**.

Use the web app for stock changes so the transaction history remains consistent. The Stock tab is a derived display; editing it is not a stock transaction.

## 7. Optional: connect the camera scanner

The inventory app can be used now with manually entered IDs. To enable its **Start Camera Scan** link:

1. In your own Cloudflare account, create a **Pages** project using **Direct Upload / Drag and drop**.
2. Upload the contents of `scanner/`, including `index.html`, both `.mjs` files, `styles.css`, `_headers`, and `robots.txt`. The uploaded site's root must contain `index.html`. These files are already static; no build command is needed.
3. Deploy and copy your production URL, such as `https://your-project.pages.dev/`. Cloudflare's [Direct Upload guide](https://developers.cloudflare.com/pages/get-started/direct-upload/#drag-and-drop) covers this workflow.
4. In Apps Script **Project Settings → Script Properties**, add **INVENTORY_SCANNER_URL** with your own HTTPS `pages.dev` URL, then save. See Google's [Script Properties guide](https://developers.google.com/apps-script/guides/properties#manage_script_properties_manually).
5. Reload **Scan QR Label** in the inventory app. Launch the camera from that page, allow camera access, scan a label, and tap **Open inventory record**.

The current app accepts HTTPS `*.pages.dev` scanner addresses without embedded credentials or a custom port. Custom domains and GitHub Pages scanner URLs are outside this setup; do not weaken the existing validator to work around an incorrect URL.

Launch scanning through the inventory app, so it supplies the return address for **your** Apps Script deployment. The scanner validates that address as an Apps Script web app URL and returns the scanned record ID to its scan page. It does not contain a spreadsheet ID, catalog, or Google credentials. Camera frames are decoded in the browser rather than uploaded by this scanner.

Leave `INVENTORY_SCANNER_URL` absent or blank to keep camera scanning unconfigured. Record lookup by typing the printed ID still works.

## Configuration reference

These are **Script Properties** in your own project, not values to edit into the repository:

| Property | How it is set |
|---|---|
| INVENTORY_SPREADSHEET_ID | Written by `setupInventory()` for the bound spreadsheet |
| INVENTORY_PHOTO_FOLDER_ID | Written by photo storage setup |
| INVENTORY_SCANNER_URL | Optional; your Cloudflare Pages scanner URL |
| INVENTORY_ADD_FORM_ID | Optional; your Add Google Form |
| INVENTORY_CONSUME_FORM_ID | Optional; your Consume Google Form |
| INVENTORY_MOVE_FORM_ID | Optional; your Move Inventory Google Form |
| INVENTORY_MOVE_BOX_FORM_ID | Optional; your Move Box Google Form |
| INVENTORY_NEW_BOX_FORM_ID | Optional; your New Box Google Form |

Leave all Forms properties blank unless you set up that separate workflow. Runtime ID counters are maintained by the app; do not copy another installation's properties.

## Print labels

Choose **Print QR Labels**, select records, choose 48, 24, or 8 labels per sheet, and preview them. Print on **US Letter**, portrait, at **100% / actual size**, with browser headers and footers disabled. The layouts use a half-inch page margin and dotted cut lines for plain paper cut by hand. Test one sheet with your printer before printing a large batch.

## Update your installation

Back up the spreadsheet and script before updating an installation with records. Copy the updated files into the same bound project without duplicating functions. Keep your existing Script Properties.

Save, then use **Deploy → Manage deployments → Edit → New version → Deploy** to update the existing web app URL. Upload the complete `scanner/` folder again for scanner updates. Run the checks in [Validation](VALIDATION.md) in an isolated test copy before relying on an updated installation.

## If something does not load

| Symptom | Check |
|---|---|
| Missing file or undefined photo helper | Copy all files, including BoxItemPhotos and ScanPrefill, with exact names |
| Missing sheet/header | Run setup from the bound spreadsheet and compare with Schema.gs |
| Permission or photo error | Sign in with the intended account; check spreadsheet and photo-folder sharing |
| Scanner link unavailable | Check INVENTORY_SCANNER_URL is your HTTPS pages.dev production URL |
| QR library did not load | Check the browser can reach jsDelivr, then reload |
| Changes do not appear | Save the files and deploy a new version of the existing web app |

The fresh-install tests, including box-item photos, are listed in [Validation](VALIDATION.md).
