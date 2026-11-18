let
    // dim_SKU: product reference table built from SKU_Master. Original Category kept for traceability.
    // Product_Type_Derived added because 10 of 14 original categories do not match the product names.
    Source = src_SKU_Master,
    Renamed = Table.RenameColumns(Source, {{"SKU ID", "SKU_ID"}, {"Product Name", "Product_Name"}, {"Category", "Category_Original"}, {"MRP (INR)", "MRP"}}),
    Trimmed = Table.TransformColumns(Renamed, {{"SKU_ID", Text.Trim, type text}}),
    // Product type = product name without its first word, e.g. Coral Cushion Covers becomes Cushion Covers
    AddedProductType = Table.AddColumn(Trimmed, "Product_Type_Derived", each Text.AfterDelimiter([Product_Name], " "), type text),
    Selected = Table.SelectColumns(AddedProductType, {"SKU_ID", "Product_Name", "Brand", "Category_Original", "Product_Type_Derived", "MRP"}),
    Typed = Table.TransformColumnTypes(Selected, {{"MRP", Int64.Type}})
in
    Typed
