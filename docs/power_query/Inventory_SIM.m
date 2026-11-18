let
    // Inventory_SIM: SIMULATED daily stock ledger per SKU, 1-31 Oct 2026 (not real stock).
    // Closing_Stock = Opening_Stock + Units_Received - Units_Fulfilled. Stock never goes negative.
    // Units_Ordered_Real = real units sold that day (blank for 29-31 Oct: no sales data). Backorders reduce stock only when dispatched.
    // Replace the path below with the location of 02_Simulated_Data on your computer.
    Source = Excel.Workbook(File.Contents("C:\Users\Admin\OneDrive\Desktop\Marketplace_Ops_Tracker\02_Simulated_Data\Synthetic_Ops_Data_SIMULATED.xlsx"), null, true),
    Sheet = Source{[Item = "Inventory_SIM", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Date", type date}, {"SKU_ID", type text}, {"Opening_Stock", Int64.Type}, {"Units_Received", Int64.Type}, {"Units_Ordered_Real", Int64.Type}, {"Units_Fulfilled", Int64.Type}, {"Closing_Stock", Int64.Type}, {"Backorder_Units_EOD", Int64.Type}, {"Reorder_Point", Int64.Type}, {"PO_Raised_Qty", Int64.Type}, {"Planned_Arrival", type date}, {"Data_Type", type text}}),
    Merged = Table.NestedJoin(Typed, {"SKU_ID"}, dim_SKU, {"SKU_ID"}, "SKU", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "SKU", {"Product_Name"}),
    Reordered = Table.SelectColumns(Expanded, {"Date", "SKU_ID", "Product_Name", "Opening_Stock", "Units_Received", "Units_Ordered_Real", "Units_Fulfilled", "Closing_Stock", "Backorder_Units_EOD", "Reorder_Point", "PO_Raised_Qty", "Planned_Arrival", "Data_Type"}),
    Sorted = Table.Sort(Reordered, {{"SKU_ID", Order.Ascending}, {"Date", Order.Ascending}})
in
    Sorted
