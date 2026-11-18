let
    // Sales_Data: REAL daily sales and traffic for Amazon, Flipkart and Myntra (1-28 Oct 2026), one row per Date x Channel x SKU.
    // Figures are not changed. Columns are renamed to one common structure and the three channels are stacked (appended).
    // Traffic is channel-specific: Amazon = Clicks, Flipkart = Product Page Views, Myntra = PDP Views. Do not compare it directly across channels.
    // Net_Sales_exGST: Amazon and Flipkart = reported sales / 1.12; Myntra reported sales already equal MRP, so used as reported.
    // Values are NOT rounded here so totals reconcile exactly with the source; rounding is applied only in number formatting.

    // 1. Amazon
    Amz_Selected = Table.SelectColumns(src_Amazon, {"Clean_Date", "SKU_ID", "Impressions", "Clicks", "Add to Cart", "Units Ordered", "Ordered Product Sales (INR)", "Sales ex-GST"}),
    Amz_Renamed = Table.RenameColumns(Amz_Selected, {{"Clean_Date", "Date"}, {"Clicks", "Traffic"}, {"Add to Cart", "Add_to_Cart"}, {"Units Ordered", "Units"}, {"Ordered Product Sales (INR)", "Reported_Sales"}, {"Sales ex-GST", "Net_Sales_exGST"}}),
    Amz_Channel = Table.AddColumn(Amz_Renamed, "Channel", each "Amazon", type text),
    Amz_Final = Table.AddColumn(Amz_Channel, "Traffic_Measure", each "Clicks", type text),

    // 2. Flipkart (no impressions or wishlist data in this report)
    Fk_Selected = Table.SelectColumns(src_Flipkart, {"Clean Date", "SKU ID", "Product Page Views", "Add to Cart", "Units Sold", "Sales (INR)", "Sales ex-GST"}),
    Fk_Renamed = Table.RenameColumns(Fk_Selected, {{"Clean Date", "Date"}, {"SKU ID", "SKU_ID"}, {"Product Page Views", "Traffic"}, {"Add to Cart", "Add_to_Cart"}, {"Units Sold", "Units"}, {"Sales (INR)", "Reported_Sales"}, {"Sales ex-GST", "Net_Sales_exGST"}}),
    Fk_Channel = Table.AddColumn(Fk_Renamed, "Channel", each "Flipkart", type text),
    Fk_Final = Table.AddColumn(Fk_Channel, "Traffic_Measure", each "Product Page Views", type text),

    // 3. Myntra (26 Oct is missing in the source report and stays missing; it is NOT filled with zeros)
    My_Selected = Table.SelectColumns(src_Myntra, {"Clean_Date", "SKU_ID", "Impressions (num)", "PDP Views", "Added to Bag", "Wishlist Adds", "Units", "Sales (INR)"}),
    My_Renamed = Table.RenameColumns(My_Selected, {{"Clean_Date", "Date"}, {"Impressions (num)", "Impressions"}, {"PDP Views", "Traffic"}, {"Added to Bag", "Add_to_Cart"}, {"Wishlist Adds", "Wishlist_Adds"}, {"Sales (INR)", "Reported_Sales"}}),
    My_NetSales = Table.AddColumn(My_Renamed, "Net_Sales_exGST", each [Reported_Sales], type number),
    My_Channel = Table.AddColumn(My_NetSales, "Channel", each "Myntra", type text),
    My_Final = Table.AddColumn(My_Channel, "Traffic_Measure", each "PDP Views", type text),

    // 4. Stack the three channels. Columns a channel does not report become blank (null), not zero.
    Combined = Table.Combine({Amz_Final, Fk_Final, My_Final}),
    Trimmed = Table.TransformColumns(Combined, {{"SKU_ID", Text.Trim, type text}}),
    Typed = Table.TransformColumnTypes(Trimmed, {{"Date", type date}, {"Impressions", Int64.Type}, {"Traffic", Int64.Type}, {"Add_to_Cart", Int64.Type}, {"Wishlist_Adds", Int64.Type}, {"Units", Int64.Type}, {"Reported_Sales", type number}, {"Net_Sales_exGST", type number}}),

    // 5. Add product details from dim_SKU (left join, so no sales rows can be lost)
    Merged = Table.NestedJoin(Typed, {"SKU_ID"}, dim_SKU, {"SKU_ID"}, "SKU", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "SKU", {"Product_Name", "Brand", "Category_Original", "Product_Type_Derived"}),

    // 6. Helper columns for reporting
    AddedWeek = Table.AddColumn(Expanded, "Week_Start", each Date.StartOfWeek([Date], Day.Monday), type date),
    AddedDataType = Table.AddColumn(AddedWeek, "Data_Type", each "Real", type text),
    Reordered = Table.SelectColumns(AddedDataType, {"Date", "Week_Start", "Channel", "SKU_ID", "Product_Name", "Brand", "Product_Type_Derived", "Category_Original", "Traffic_Measure", "Impressions", "Traffic", "Add_to_Cart", "Wishlist_Adds", "Units", "Reported_Sales", "Net_Sales_exGST", "Data_Type"}),
    Sorted = Table.Sort(Reordered, {{"Date", Order.Ascending}, {"Channel", Order.Ascending}, {"SKU_ID", Order.Ascending}})
in
    Sorted
