# Validation

## Local checks performed

- The collected Apps Script functions and HTML scripts parse, and the reviewed page routes, included fragments, and browser-to-server method names resolve.
- Setup and configuration have 16 mocked checks covering a fresh nine-tab setup, reruns retaining existing records/formulas, malformed headers, conflicting configuration, optional Forms, and error/lock handling.
- Eight scanner tests cover accepted IDs, valid return addresses, exact return IDs, duplicate frames, permission failures, and camera cleanup. All 24 included tests passed.
- The empty template's nine header rows match the supplied schema, and its saved workbook is reopened to confirm the headers and blank data area.
- All nine template sheets were rendered and visually checked. The saved workbook contains 77 header cells, zero populated data rows, zero formulas, and frozen header rows.
- The source package is checked for the original installation's spreadsheet/form/project IDs and scanner URL before it is zipped. No inventory rows, photos, account tokens, or real local clasp configuration are included.

Run the included checks with Node.js 20 or newer:

```text
node --test tests/setup.test.cjs tests/scanner.test.mjs
```

These local checks use mocks and source parsing. They do not confirm Google authorization, Google Sheets conversion, file sharing, a deployed camera, or physical printer alignment.

## Fresh-install acceptance check

Use a separate, disposable installation following [Setup](SETUP.md). Its Google deployment has not been verified as part of this local package build.

1. Run `setupInventory()` and confirm it creates/reuses the nine tabs and its own photo folder. Run it again and confirm the headers remain unchanged.
2. Deploy with the supplied accessing-user execution model. Open the Home page while signed in.
3. Create a location, an outer box, and a box inside it. Create an item with a unit, initial quantity, placement, and optional photo.
4. Find the item and confirm its total and recorded path. Add, consume, move, issue, and return disposable stock; check that totals follow the saved Transactions rows.
5. Open a scanned box record and confirm its location, nested boxes, contained items, and saved photos. Also check an item with no photo. Move the box and confirm its contained boxes/items follow it.
6. Print a few genuine QR labels on US Letter at actual size. Scan the printed labels using your own Cloudflare scanner and open the matching record.
7. Follow an action link from a scanned item/box and confirm the correct selection, unit, photo, and placement.
8. If sharing the installation, test with another intended account that has access to your spreadsheet/photo folder. Test an unshared account separately to verify Google rejects access to your records.

Check the transaction history before repeating any request whose result was uncertain. The forms retain request references to resolve retries without creating a duplicate movement.

The box-item photo enhancement was present in the collected source and passed local callback/queue checks during development. Its deployed photo display still needs the live check above.
