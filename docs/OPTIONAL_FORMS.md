# Optional Google Forms workflow

The web app works without Google Forms or an installed form-submission trigger. This optional legacy path is for people who want to maintain their own Forms in addition to the app.

Create your own forms and link their responses to your inventory spreadsheet. Each response tab must have the exact name and headers expected by the processors. Google creates response tabs when linking a form; it does not automatically reuse the empty request tabs from this template. Reconcile those names in a new test installation before using Forms with real records. Keep a backup of any existing records.

## Configured forms

Set the following Script Properties to the IDs of your own forms. Use the ID, not its full URL. Unconfigured refresh helpers return without opening any form.

| Script Property | Form | Exact dropdown titles |
|---|---|---|
| INVENTORY_ADD_FORM_ID | Add stock | Item, Destination |
| INVENTORY_CONSUME_FORM_ID | Consume stock | Item, Source |
| INVENTORY_MOVE_FORM_ID | Move stock | Item, Source, Destination |
| INVENTORY_MOVE_BOX_FORM_ID | Move a box to a location | Box, Destination Location |
| INVENTORY_NEW_BOX_FORM_ID | Create a box at a location | Location |

Create at least one item and storage location before running `refreshAllInventoryForms()`, since Forms dropdowns need actual choices. The refresh helpers generate the `ID — name` labels used by the processors.

## Response tabs

`template/schema.json` contains the exact headers for the nine core tabs and the four additional optional response tabs: `AddRequests`, `ConsumeRequests`, `MoveRequests`, and `BoxMoveRequests`. Keep Google's Timestamp column and append the processing columns named there. Creation Forms use the core `NewItemRequests`, `NewBoxRequests`, and `NewLocationRequests` tabs.

The form question titles must match the request field headers, including spaces. Processing columns such as RequestID, Status, Message, ProcessedAt, and CreatedItemID are app metadata, not questions to show in a form. Newly created response sheets need these metadata columns added after the response fields.

After checking the response-tab names and headers, install one trigger in the spreadsheet-bound Apps Script project:

- Function: `handleInventorySubmission`
- Event source: **From spreadsheet**
- Event type: **On form submit**

Do not create a separate trigger for each processor. The dispatcher selects the processor from the response-tab name. [Google's installable trigger guide](https://developers.google.com/apps-script/guides/triggers/installable) explains the trigger's execution identity and authorization.

## Scope of the legacy forms

Use the web app's **Move Box** page for nested-box movements. The legacy Move Box Form processor changes a box's location and does not provide the web page's parent-box movement controls. The New Box Form refresh helper offers locations; nested placement is available in the web app.

Use the web app's **New Item** page for initial stock and photo cropping. The optional form dispatcher does not run that page's separately tracked initial-stock/photo workflow.

Test each configured form with disposable records, check the request status and ledger, and retry the same submission event before using it with real inventory. The reusable package's local setup tests do not validate your form links, question definitions, sharing, or installed trigger.
