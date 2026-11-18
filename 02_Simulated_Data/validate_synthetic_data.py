"""
Validate the SIMULATED order and inventory data against the REAL sales data.

Usage (from the repository root):
    python 02_Simulated_Data/validate_synthetic_data.py
or with explicit paths:
    python 02_Simulated_Data/validate_synthetic_data.py <source_cleaned.xlsx> <simulated.xlsx>

Every check prints PASS or FAIL, followed by a short summary of what the tracker will show.
"""
import os
import sys
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from generate_synthetic_ops_data import load_real  # noqa: E402

ROOT = os.path.dirname(HERE)
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "01_Source", "Source_Marketplace_Data_Cleaned.xlsx")
sim = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "Synthetic_Ops_Data_SIMULATED.xlsx")

real = load_real(src)
real["Date"] = pd.to_datetime(real["Date"])
o = pd.read_excel(sim, "Orders_SIM")
inv = pd.read_excel(sim, "Inventory_SIM")
for c in ["Order_Date", "SLA_Due_Date", "Dispatch_Date", "Delivery_Date", "Return_Date"]:
    o[c] = pd.to_datetime(o[c])
inv["Date"] = pd.to_datetime(inv["Date"])
rec = o[o.Included_In_Recorded_Sales == "Yes"]
can = o[o.Included_In_Recorded_Sales == "No"]
checks = {}

# 1. Recorded orders reconcile to real units and real net sales per Date x Channel x SKU
g = rec.groupby(["Order_Date", "Channel", "SKU_ID"]).agg(Q=("Qty", "sum"), V=("Order_Value", "sum")).reset_index()
r = real[real.Units > 0].rename(columns={"Date": "Order_Date"})
m = r.merge(g, on=["Order_Date", "Channel", "SKU_ID"], how="outer")
checks["Units match every Date x Channel x SKU"] = bool((m.Units == m.Q).all())
checks["Order value matches real net sales (within 0.01)"] = bool(((m.Net_Sales - m.V).abs() < 0.01).all())
checks["Total recorded units equal real units"] = int(rec.Qty.sum()) == int(real.Units.sum())
# 2. Cancelled test orders stay separate
checks["Cancelled orders are never dispatched"] = bool(can.Dispatch_Date.isna().all())
# 3. Missing Myntra day stays missing
checks["No Myntra orders on 26 Oct"] = int(((o.Channel == "Myntra") & (o.Order_Date == "2026-10-26")).sum()) == 0
# 4. Inventory logic
inv = inv.sort_values(["SKU_ID", "Date"])
checks["Closing = Opening + Received - Fulfilled"] = bool((inv.Closing_Stock == inv.Opening_Stock + inv.Units_Received - inv.Units_Fulfilled).all())
checks["Stock never negative"] = bool((inv.Closing_Stock >= 0).all())
checks["Next day opening = previous closing"] = bool(
    (inv.groupby("SKU_ID").Closing_Stock.shift(1).dropna().values == inv[inv.Date > "2026-10-01"].Opening_Stock.values).all())
disp = rec.dropna(subset=["Dispatch_Date"]).groupby(["Dispatch_Date", "SKU_ID"]).Qty.sum().sort_index()
led = inv.set_index(["Date", "SKU_ID"]).Units_Fulfilled
led = led[led > 0].rename("Qty").rename_axis(["Dispatch_Date", "SKU_ID"]).sort_index()
checks["Ledger units fulfilled = order units dispatched"] = bool(disp.equals(led))
# 5. Date logic
d = rec.dropna(subset=["Dispatch_Date"])
checks["Dispatch after order date and by 31 Oct"] = bool(((d.Dispatch_Date > d.Order_Date) & (d.Dispatch_Date <= "2026-10-31")).all())
dl = rec.dropna(subset=["Delivery_Date"])
checks["Delivery after dispatch and by 31 Oct"] = bool(((dl.Delivery_Date > dl.Dispatch_Date) & (dl.Delivery_Date <= "2026-10-31")).all())
checks["Returned units never exceed order qty"] = bool((rec.Returned_Units <= rec.Qty).all())

for k, v in checks.items():
    print(("PASS  " if v else "FAIL  ") + k)

due = rec[rec.SLA_Due_Date <= "2026-10-31"]
within = due.Dispatch_Date.notna() & (due.Dispatch_Date <= due.SLA_Due_Date)
delivered_units = rec[rec.Order_Status.isin(["Delivered", "Returned"])].Qty.sum()
print("\nSummary (simulated test data)")
print(f"  Recorded orders {len(rec):,} | cancelled test orders {len(can)} ({len(can) / len(o):.1%})")
print(f"  Dispatch SLA {within.mean():.1%} | pending orders {rec.Dispatch_Date.isna().sum()}")
print(f"  Return rate {rec.Returned_Units.sum() / delivered_units:.1%} of delivered units")
sys.exit(0 if all(checks.values()) else 1)
