let
    // Orders_SIM: SIMULATED test order data (not real orders). Generated from the real Oct 2026 sales.
    // Rows with Included_In_Recorded_Sales = "Yes" reconcile exactly to real units per Date x Channel x SKU.
    // Rows with "No" are cancelled test orders, kept separate and excluded from recorded sales and stock.
    // Replace the path below with the location of 02_Simulated_Data on your computer.
    Source = Excel.Workbook(File.Contents("C:\Users\Admin\OneDrive\Desktop\Marketplace_Ops_Tracker\02_Simulated_Data\Synthetic_Ops_Data_SIMULATED.xlsx"), null, true),
    Sheet = Source{[Item = "Orders_SIM", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"Order_ID", type text}, {"Order_Date", type date}, {"Channel", type text}, {"SKU_ID", type text}, {"Qty", Int64.Type}, {"Unit_Net_Price", type number}, {"Order_Value", type number}, {"Warehouse", type text}, {"Included_In_Recorded_Sales", type text}, {"Order_Status", type text}, {"SLA_Due_Date", type date}, {"Dispatch_Date", type date}, {"Delivery_Date", type date}, {"Delay_Reason", type text}, {"Returned_Units", Int64.Type}, {"Return_Date", type date}, {"Data_Type", type text}}),
    Merged = Table.NestedJoin(Typed, {"SKU_ID"}, dim_SKU, {"SKU_ID"}, "SKU", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "SKU", {"Product_Name", "Product_Type_Derived"}),
    Reordered = Table.SelectColumns(Expanded, {"Order_ID", "Order_Date", "Channel", "SKU_ID", "Product_Name", "Product_Type_Derived", "Qty", "Unit_Net_Price", "Order_Value", "Warehouse", "Included_In_Recorded_Sales", "Order_Status", "SLA_Due_Date", "Dispatch_Date", "Delivery_Date", "Delay_Reason", "Returned_Units", "Return_Date", "Data_Type"}),
    Sorted = Table.Sort(Reordered, {{"Order_Date", Order.Ascending}, {"Channel", Order.Ascending}, {"Order_ID", Order.Ascending}})
in
    Sorted
// In the workbook, three calculated columns are added beside this table (Orders_Fulfilment!AM:AO):
//   Days_to_Dispatch  =IF(ISBLANK(AG4),"-",AG4-U4)
//   SLA_Result        =IF(AD4="No","Excluded - cancelled",IF(ISBLANK(AG4),IF(Report_Date>AF4,"Breach - still pending","Pending - within SLA"),IF(AG4<=AF4,"Within SLA","Breach - late dispatch")))
//   Pending_Age_Days  =IF(AND(AD4="Yes",ISBLANK(AG4)),Report_Date-U4,"-")
