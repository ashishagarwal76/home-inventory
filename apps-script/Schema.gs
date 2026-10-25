// Sheet names, header order, and PhotoURL positions are part of the app format.
function inventoryTemplateSchema_() {
  return [
    {
      "name": "Stock",
      "headers": [
        "ItemID",
        "ItemName",
        "StorageID",
        "BoxName",
        "LocationName",
        "Quantity",
        "Unit"
      ]
    },
    {
      "name": "Items",
      "headers": [
        "ItemID",
        "ItemName",
        "Category",
        "Unit",
        "MinimumStock",
        "Notes",
        "PhotoURL"
      ]
    },
    {
      "name": "Boxes",
      "headers": [
        "BoxID",
        "BoxName",
        "LocationID",
        "Notes",
        "PhotoURL",
        "ParentBoxID"
      ]
    },
    {
      "name": "Locations",
      "headers": [
        "LocationID",
        "LocationName",
        "Notes"
      ]
    },
    {
      "name": "Transactions",
      "headers": [
        "TransactionID",
        "Timestamp",
        "Action",
        "ItemID",
        "Quantity",
        "FromStorageID",
        "ToStorageID",
        "Notes"
      ]
    },
    {
      "name": "BoxMoves",
      "headers": [
        "MovementID",
        "Timestamp",
        "BoxID",
        "FromLocationID",
        "ToLocationID",
        "Notes",
        "FromStorageID",
        "ToStorageID",
        "Status",
        "Message",
        "ProcessedAt"
      ]
    },
    {
      "name": "NewItemRequests",
      "headers": [
        "Timestamp",
        "Item Name",
        "Category",
        "Unit",
        "Minimum Stock",
        "Notes",
        "RequestID",
        "Status",
        "Message",
        "ProcessedAt",
        "CreatedItemID",
        "Initial Quantity",
        "Initial StorageID",
        "InitialTransactionID"
      ]
    },
    {
      "name": "NewBoxRequests",
      "headers": [
        "Timestamp",
        "Box Name",
        "Location",
        "Notes",
        "RequestID",
        "Status",
        "Message",
        "ProcessedAt",
        "CreatedBoxID",
        "CreatedPlacementID",
        "InitialPlacementLocation",
        "InitialPlacementDescription",
        "InitialPlacementAt"
      ]
    },
    {
      "name": "NewLocationRequests",
      "headers": [
        "Timestamp",
        "Location Name",
        "Notes",
        "RequestID",
        "Status",
        "Message",
        "ProcessedAt",
        "CreatedLocationID"
      ]
    }
  ];
}
