# Power Query code

These are the queries inside `Marketplace_Operations_Tracker.xlsx`, exported so they can be reviewed without opening Excel.
In Excel: **Data > Queries & Connections**, then right-click a query > **Edit** > **Advanced Editor**.

| Query | Loads to | Purpose |
| --- | --- | --- |
| `src_Amazon`, `src_Flipkart`, `src_Myntra`, `src_SKU_Master`, `src_Ads` | Connection only | Read the cleaned sheets from `01_Source`. `src_Flipkart` also removes 12 empty rows left over from duplicate removal. |
| `dim_SKU` | Connection only | Product reference; keeps the original Category and adds `Product_Type_Derived` from the product name. |
| `Sales_Data` | `Sales_Performance!T3` | Appends the three channels into one REAL table (1,162 rows). |
| `Ads_Data` | `Sales_Performance!AL3` | REAL ads data with channel names matched to `Sales_Data`. |
| `Orders_SIM` | `Orders_Fulfilment!T3` | SIMULATED order data (4,118 rows, incl. 124 cancelled test orders). |
| `Inventory_SIM` | `Inventory!T3` | SIMULATED daily stock ledger (14 SKUs x 31 days). |

The `src_*` queries were created with **Get Data > From File > From Excel Workbook** (Source > Navigation > Promoted Headers > Changed Type).
File paths are absolute; after downloading, update them with **Data > Get Data > Data Source Settings > Change Source**.
