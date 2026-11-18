let
    // Ads_Data: REAL advertising data for the 6 advertised SKUs. Ad sales are as attributed by each platform.
    // Channel codes are mapped to the names used in Sales_Data so the two tables can be compared.
    Source = src_Ads,
    Renamed = Table.RenameColumns(Source, {{"SKU ID", "SKU_ID"}, {"Ad Spend (INR)", "Ad_Spend"}, {"Ad Sales (INR)", "Ad_Sales_Platform"}, {"Channel", "Channel_Code"}}),
    Mapped = Table.AddColumn(Renamed, "Channel", each if [Channel_Code] = "AMZ" then "Amazon" else if [Channel_Code] = "FK" then "Flipkart" else if [Channel_Code] = "MY" then "Myntra" else null, type text),
    Typed = Table.TransformColumnTypes(Mapped, {{"Date", type date}, {"SKU_ID", type text}, {"Ad_Spend", type number}, {"Ad_Sales_Platform", type number}}),
    Selected = Table.SelectColumns(Typed, {"Date", "Channel", "SKU_ID", "Ad_Spend", "Ad_Sales_Platform"}),
    AddedDataType = Table.AddColumn(Selected, "Data_Type", each "Real", type text)
in
    AddedDataType
