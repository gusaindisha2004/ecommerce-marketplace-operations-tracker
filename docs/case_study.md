# Case study: E-commerce Marketplace Operations Tracker

*Self-initiated portfolio project by Disha Gusain. Sales, traffic and ads data are real; order, dispatch and stock data are simulated and labelled as such.*

## Situation

A home-decor seller lists 14 SKUs on Amazon, Flipkart and Myntra. Each marketplace exports a different report: different product IDs, date formats, GST treatment and traffic definitions. There was no single view of sales, ad efficiency, fulfilment and stock risk.

## Task

Build a weekly operations review in Excel that a manager can understand in one minute, without changing any real figure and without presenting simulated results as real findings.

## Approach

1. **Understand the data first.** I profiled the cleaned source before any modelling. I found a missing Myntra day (26 Oct), wrong SKU-master categories for 10 of 14 SKUs, different GST bases, and ad sales that exceed recorded sales on 339 of 498 rows.
2. **Agree the rules before simulating.** I wrote the reconciliation rules for the test data down first: recorded orders must equal real units per Date x Channel x SKU, cancellations are kept separate, stock is reduced only on dispatch, the missing day stays missing, and the report date is 31 Oct with a 2-day SLA.
3. **Build a repeatable pipeline.** Ten Power Query queries combine the three channels into one REAL sales table (1,162 rows) and load the SIMULATED orders and inventory.
4. **Calculate KPIs with transparent formulas.** SUMIFS, COUNTIFS, XLOOKUP and FILTER drive the KPIs, and three named parameters (`Report_Date`, `SLA_Days`, `Cover_Days`) control them.
5. **Design for a one-minute read.** The dashboard has KPI cards, red/amber/green flags, a live action list, and a PivotTable with slicers for drilling down.
6. **Verify.** Nine live reconciliation checks sit in the workbook, 12 scripted checks run on the simulated data, and 169 KPI values were recalculated independently. All match.

## Results

- Net sales (real): ₹27.8 lakh. Amazon 41.1%, Flipkart 33.7%, Myntra 25.3%.
- Concentration risk (real): one SKU (Ivory Candle Holders) makes up 32.3% of net sales.
- Ad efficiency (real): TACoS 33.6% overall. Platform ROAS of 4.64x is overstated by attribution.
- Data issues logged (real): the missing Myntra day, a 7-day zero-sales run on a Flipkart SKU despite traffic, and wrong categories.
- Monitoring demonstrated (simulated): 92.6% dispatch SLA against a 95% target. The Bengaluru warehouse is the bottleneck, and a stock-out on SKU028 is driving 30 pending orders.

## What I would improve next

- Store the 95% SLA target as a named parameter on `Assumptions_DQ`.
- Show product names instead of SKU IDs in the PivotTable.
- Tidy the weekly chart's title and legend.
- Add more months of data, and replace the simulated layer with real order and inventory exports.

## Skills shown

Data cleaning and reconciliation, Power Query, Excel formulas and PivotTables, KPI design (TACoS, SLA, days of cover), dashboard design, documenting assumptions, and separating real data from simulated data honestly.
