# Home Inventory

A browser app for keeping track of household items, the boxes they are stored in, and items lent to family members. Each installation uses its own Google Sheet and Google Drive photo folder.

## What it does

- Create items, boxes, and storage locations, with optional item and box photos.
- Add, consume, and move stock; issue items to family members and record returns.
- Nest boxes inside other boxes and move a whole box to a new placement.
- Search item descriptions and view quantities and recorded positions.
- Print genuine QR labels in 48, 24, or 8-label layouts on US Letter paper.
- Scan an item, box, or location ID. A box record shows its location, nested boxes, contents, and saved item photos.
- Open stock actions from a scanned item with that item already selected.

## How it runs

The inventory app runs on **Google Apps Script**, with Google Sheets holding the records and Google Drive holding photos. The optional camera scanner is a separate static page hosted on your own **Cloudflare Pages** site. QR labels contain record IDs; the app looks up the corresponding data after a scan.

This repository supplies the source and a blank template. Posting the source on GitHub does not deploy the inventory app. Follow [Setup](docs/SETUP.md) to create your own installation. No command-line tools are required.

## Included files

| Folder or file | Purpose |
|---|---|
| `apps-script/` | Apps Script code, HTML pages, setup helpers, and manifest |
| `scanner/` | Ready-to-upload static camera scanner |
| `template/HomeInventory-template.xlsx` | Empty workbook to upload and convert to Google Sheets |
| [Setup](docs/SETUP.md) | Installation, sharing, scanner connection, and updates |
| [Post on GitHub](docs/PUBLISH_TO_GITHUB.md) | Upload this reusable project to your repository |
| [Optional Forms](docs/OPTIONAL_FORMS.md) | Separate Google Forms workflow |
| [Validation](docs/VALIDATION.md) | Checks performed and tests for a new installation |
| [Third-party notices](THIRD_PARTY_NOTICES.md) | QR library versions and licenses |

The spreadsheet template contains headers, not inventory records. Account credentials, installed Script Property values, deployment IDs, response rows, and photos do not belong in the repository.

## Start here

1. Create your Google Sheets copy and open its bound Apps Script editor.
2. Copy in the files from `apps-script/`, then run `setupInventory()` once from that editor.
3. Deploy the web app, keeping the supplied run-as-accessing-user model.
4. Optionally deploy `scanner/` to your own Cloudflare Pages project and configure its URL.

The complete instructions are in [Setup](docs/SETUP.md). A fresh Google installation and the box-item photo display still need the live acceptance checks in [Validation](docs/VALIDATION.md); local source checks do not establish those results.

## Development checks

With Node.js 20 or newer, run `node --test tests/setup.test.cjs tests/scanner.test.mjs`. No package installation is needed for these tests.

## License

The app source is available under the [MIT license](LICENSE). QR libraries loaded by the app and scanner retain their own licenses; see [Third-party notices](THIRD_PARTY_NOTICES.md).
