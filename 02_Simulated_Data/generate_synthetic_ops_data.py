"""
Generate SIMULATED (synthetic test) order and inventory data for the
E-commerce Marketplace Operations Tracker.

Source of truth: the REAL cleaned sales data (Amazon, Flipkart, Myntra, 1-28 Oct 2026).
Rules (agreed before generation):
  1. Recorded orders reconcile EXACTLY to real units for every Date x Channel x SKU.
     Order value = qty x real ex-GST unit price, so values reconcile to real net sales.
  2. Cancelled test orders are separate rows (Included_In_Recorded_Sales = "No") and
     never touch recorded sales or inventory.
  3. Returns are tracked on units (Returned_Units), not orders.
  4. Inventory: Closing = Opening + Received - Units Fulfilled (dispatched).
     Backordered orders do not reduce stock until fulfilled. Stock never goes negative.
  5. Backorders wait for the next restock and are then dispatched (linked by date).
  6. Myntra 26 Oct is missing in the real data and stays missing (no orders invented).
  7. Reporting date 31 Oct 2026; dispatch SLA 2 days; pending orders past SLA are breaches.
Fixed random seed -> the output is reproducible.
Usage: python generate_synthetic_ops_data.py <source_cleaned.xlsx> <output.xlsx>
"""
import sys
import random
from datetime import date, timedelta
import pandas as pd

SEED = 2026
REPORT_DATE = date(2026, 10, 31)
SALES_END = date(2026, 10, 28)
SLA_DAYS = 2
LEAD_TIME_DAYS = 6          # supplier lead time from reorder to receipt
SAFETY_DAYS = 1             # safety stock, in days of average demand
ORDER_QTY_DAYS = 10         # reorder quantity, in days of average demand
QTY2_SHARE = 0.15           # share of orders with 2 units (rest 1 unit)
CANCEL_RATE = 0.03          # cancelled orders as share of all orders placed
RETURN_RATE_UNITS = 0.06    # returned units as share of delivered units
WAREHOUSE_SPLIT = {"Delhi": 0.55, "Bengaluru": 0.45}
PROCESSING_DAYS = {         # days from order to ready-to-dispatch
    "Delhi": [(1, 0.85), (2, 0.12), (3, 0.03)],
    "Bengaluru": [(1, 0.60), (2, 0.28), (3, 0.12)],
}
TRANSIT_DAYS = [(2, 0.30), (3, 0.35), (4, 0.25), (5, 0.10)]
# Documented supplier delay scenario: POs for these SKUs raised on/after the date arrive late
SUPPLIER_DELAYS = {"SKU028": (date(2026, 10, 15), 6), "SKU038": (date(2026, 10, 20), 4)}
CH_CODE = {"Amazon": "AMZ", "Flipkart": "FK", "Myntra": "MY"}


def pick(rng, weighted):
    r, acc = rng.random(), 0.0
    for val, w in weighted:
        acc += w
        if r <= acc:
            return val
    return weighted[-1][0]


def load_real(path):
    a = pd.read_excel(path, "Amazon_Clean")
    f = pd.read_excel(path, "Flipkart_Clean").dropna(how="all")
    m = pd.read_excel(path, "Myntra_Clean")
    rows = []
    for _, r in a.iterrows():
        rows.append(("Amazon", r["Clean_Date"], r["SKU_ID"], int(r["Units Ordered"]), float(r["Sales ex-GST"])))
    for _, r in f.iterrows():
        rows.append(("Flipkart", r["Clean Date"], str(r["SKU ID"]).strip(), int(r["Units Sold"]), float(r["Sales ex-GST"])))
    for _, r in m.iterrows():
        rows.append(("Myntra", r["Clean_Date"], str(r["SKU_ID"]).strip(), int(r["Units"]), float(r["Sales (INR)"])))
    df = pd.DataFrame(rows, columns=["Channel", "Date", "SKU_ID", "Units", "Net_Sales"])
    df["Date"] = pd.to_datetime(df["Date"]).dt.date
    return df


def main(src, out):
    rng = random.Random(SEED)
    real = load_real(src)
    skus = sorted(real["SKU_ID"].unique())
    days = [date(2026, 10, 1) + timedelta(d) for d in range(31)]

    # ---------- 1. Recorded orders (reconcile exactly to real units) ----------
    orders, seq = [], {}
    for _, r in real.sort_values(["Date", "Channel", "SKU_ID"]).iterrows():
        if r.Units == 0:
            continue
        unit_price = r.Net_Sales / r.Units
        left = r.Units
        while left > 0:
            q = 2 if (left >= 2 and rng.random() < QTY2_SHARE) else 1
            left -= q
            key = (r.Channel, r.Date)
            seq[key] = seq.get(key, 0) + 1
            wh = pick(rng, list(WAREHOUSE_SPLIT.items()))
            orders.append({
                "Order_ID": f"{CH_CODE[r.Channel]}-{r.Date:%m%d}-{seq[key]:03d}",
                "Order_Date": r.Date, "Channel": r.Channel, "SKU_ID": r.SKU_ID, "Qty": q,
                "Unit_Net_Price": unit_price, "Order_Value": q * unit_price,
                "Warehouse": wh, "Ready_Date": r.Date + timedelta(pick(rng, PROCESSING_DAYS[wh])),
                "Included_In_Recorded_Sales": "Yes",
            })

    # ---------- 2. Inventory policy per SKU (from October average demand) ----------
    avg = real.groupby("SKU_ID")["Units"].sum() / 28.0
    policy = {}
    for s in skus:
        d = avg[s]
        policy[s] = {
            "avg_daily": d,
            "reorder_point": round(d * (LEAD_TIME_DAYS + SAFETY_DAYS)),
            "order_qty": max(5, round(d * ORDER_QTY_DAYS)),
            "opening": round(d * rng.uniform(5, 9)),
        }

    # ---------- 3. Day-by-day fulfilment + inventory simulation ----------
    on_hand = {s: policy[s]["opening"] for s in skus}
    pipeline = {s: [] for s in skus}            # (arrival_date, qty)
    open_orders = sorted(orders, key=lambda o: (o["Ready_Date"], o["Order_Date"], o["Order_ID"]))
    for o in orders:
        o["Dispatch_Date"] = None
        o["Waited_For_Stock"] = False
    ordered_real = real.groupby(["Date", "SKU_ID"])["Units"].sum().to_dict()
    ledger = []
    for day in days:
        for s in skus:
            opening = on_hand[s]
            received = sum(q for (ad, q) in pipeline[s] if ad == day)
            pipeline[s] = [(ad, q) for (ad, q) in pipeline[s] if ad != day]
            on_hand[s] += received
            fulfilled = 0
            for o in open_orders:
                if o["SKU_ID"] != s or o["Dispatch_Date"] is not None or o["Ready_Date"] > day:
                    continue
                if on_hand[s] >= o["Qty"]:
                    on_hand[s] -= o["Qty"]
                    fulfilled += o["Qty"]
                    o["Dispatch_Date"] = day
                else:
                    o["Waited_For_Stock"] = True
            backorder = sum(o["Qty"] for o in open_orders if o["SKU_ID"] == s and o["Dispatch_Date"] is None
                            and o["Ready_Date"] <= day)
            # reorder check: stock position = on hand + on order - backorders
            position = on_hand[s] + sum(q for _, q in pipeline[s]) - backorder
            po_qty = 0
            if day <= SALES_END and position <= policy[s]["reorder_point"]:
                po_qty = policy[s]["order_qty"] + backorder
                delay = 0
                if s in SUPPLIER_DELAYS and day >= SUPPLIER_DELAYS[s][0]:
                    delay = SUPPLIER_DELAYS[s][1]
                pipeline[s].append((day + timedelta(LEAD_TIME_DAYS + delay), po_qty))
            assert on_hand[s] >= 0
            ledger.append({
                "Date": day, "SKU_ID": s, "Opening_Stock": opening, "Units_Received": received,
                "Units_Ordered_Real": ordered_real.get((day, s)) if day <= SALES_END else None,
                "Units_Fulfilled": fulfilled, "Closing_Stock": on_hand[s],
                "Backorder_Units_EOD": backorder, "Reorder_Point": policy[s]["reorder_point"],
                "PO_Raised_Qty": po_qty if po_qty else None,
                "Planned_Arrival": (day + timedelta(LEAD_TIME_DAYS)) if po_qty else None,
            })

    # ---------- 4. Status, delivery, delay reason ----------
    for o in orders:
        due = o["Order_Date"] + timedelta(SLA_DAYS)
        o["SLA_Due_Date"] = due
        if o["Dispatch_Date"] is None:
            o["Delivery_Date"] = None
            o["Order_Status"] = "Pending - awaiting stock" if o["Waited_For_Stock"] else "Pending - processing"
        else:
            dd = o["Dispatch_Date"] + timedelta(pick(rng, TRANSIT_DAYS))
            o["Delivery_Date"] = dd if dd <= REPORT_DATE else None
            o["Order_Status"] = "Delivered" if dd <= REPORT_DATE else "Shipped - in transit"
        late = (o["Dispatch_Date"] or REPORT_DATE + timedelta(1)) > due
        if late and REPORT_DATE >= due:
            o["Delay_Reason"] = "Stock shortage" if o["Waited_For_Stock"] else "Warehouse processing"
        else:
            o["Delay_Reason"] = None
        o["Returned_Units"] = 0
        o["Return_Date"] = None

    # Returns: 6% of delivered UNITS, only where the return date falls on/before the reporting date
    delivered = [o for o in orders if o["Order_Status"] == "Delivered"]
    target = round(RETURN_RATE_UNITS * sum(o["Qty"] for o in delivered))
    eligible = [o for o in delivered if o["Delivery_Date"] + timedelta(2) <= REPORT_DATE]
    rng.shuffle(eligible)
    got = 0
    for o in eligible:
        if got >= target:
            break
        ru = min(o["Qty"], target - got)
        rd = o["Delivery_Date"] + timedelta(rng.randint(2, 7))
        if rd > REPORT_DATE:
            rd = REPORT_DATE
        o["Returned_Units"], o["Return_Date"], o["Order_Status"] = ru, rd, "Returned"
        got += ru

    # ---------- 5. Cancelled test orders (separate; excluded from recorded sales) ----------
    n_cancel = round(len(orders) * CANCEL_RATE / (1 - CANCEL_RATE))
    weights = real[real.Units > 0]
    picks = weights.sample(n=n_cancel, weights="Units", replace=True, random_state=SEED)
    cancelled = []
    for i, (_, r) in enumerate(picks.sort_values(["Date", "Channel"]).iterrows(), start=1):
        price = r.Net_Sales / r.Units
        cancelled.append({
            "Order_ID": f"{CH_CODE[r.Channel]}-{r.Date:%m%d}-C{i:03d}", "Order_Date": r.Date,
            "Channel": r.Channel, "SKU_ID": r.SKU_ID, "Qty": 1, "Unit_Net_Price": price,
            "Order_Value": price, "Warehouse": pick(rng, list(WAREHOUSE_SPLIT.items())),
            "Included_In_Recorded_Sales": "No", "Order_Status": "Cancelled",
            "SLA_Due_Date": None, "Dispatch_Date": None, "Delivery_Date": None,
            "Delay_Reason": None, "Returned_Units": 0, "Return_Date": None,
        })

    cols = ["Order_ID", "Order_Date", "Channel", "SKU_ID", "Qty", "Unit_Net_Price", "Order_Value",
            "Warehouse", "Included_In_Recorded_Sales", "Order_Status", "SLA_Due_Date", "Dispatch_Date",
            "Delivery_Date", "Delay_Reason", "Returned_Units", "Return_Date"]
    odf = pd.DataFrame(orders + cancelled)[cols].sort_values(["Order_Date", "Channel", "Order_ID"])
    odf["Data_Type"] = "Simulated"
    ldf = pd.DataFrame(ledger)
    ldf["Data_Type"] = "Simulated"

    pol = pd.DataFrame([{"SKU_ID": s, "Avg_Daily_Units_Oct": round(policy[s]["avg_daily"], 2),
                         "Opening_Stock_1_Oct": policy[s]["opening"], "Reorder_Point": policy[s]["reorder_point"],
                         "Reorder_Qty_Base": policy[s]["order_qty"]} for s in skus])
    params = pd.DataFrame([
        ("Data type", "SIMULATED test data. Not real orders or stock. Generated from real Oct 2026 sales."),
        ("Random seed", SEED), ("Reporting date", str(REPORT_DATE)), ("Real sales period", "1-28 Oct 2026"),
        ("Dispatch SLA", f"{SLA_DAYS} days from order date"),
        ("Order size", f"{int((1-QTY2_SHARE)*100)}% 1 unit, {int(QTY2_SHARE*100)}% 2 units"),
        ("Warehouse split", "Delhi 55%, Bengaluru 45% of orders"),
        ("Processing time Delhi", "1 day 85%, 2 days 12%, 3 days 3%"),
        ("Processing time Bengaluru", "1 day 60%, 2 days 28%, 3 days 12%"),
        ("Transit time", "2 days 30%, 3 days 35%, 4 days 25%, 5 days 10%"),
        ("Cancellations", f"{int(CANCEL_RATE*100)}% of orders placed; separate rows, excluded from recorded sales and stock"),
        ("Returns", f"{int(RETURN_RATE_UNITS*100)}% of delivered units; returned units not restocked"),
        ("Inventory", "One stock pool per SKU shared by all channels and both warehouses"),
        ("Opening stock", "5-9 days of October average demand per SKU (random)"),
        ("Supplier delay scenario", "POs for SKU028 raised from 15 Oct arrive 6 days late; POs for SKU038 raised from 20 Oct arrive 4 days late (test scenario)"),
        ("Reorder rule", f"Raise PO when stock position <= reorder point; reorder point = avg daily demand x ({LEAD_TIME_DAYS} lead + {SAFETY_DAYS} safety) days"),
        ("Reorder quantity", f"{ORDER_QTY_DAYS} days of avg demand + current backorders"),
        ("Supplier lead time", f"{LEAD_TIME_DAYS} days"),
        ("Stock formula", "Closing = Opening + Received - Units Fulfilled (dispatched). Never negative."),
        ("Backorders", "Orders wait for stock; dispatched first-in-first-out when the next restock arrives"),
        ("Myntra 26 Oct", "Missing in real data; no orders simulated for that day"),
        ("29-31 Oct", "No real sales data; Units_Ordered_Real left blank. Only dispatch/receipts continue."),
    ], columns=["Assumption", "Value"])

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        odf.to_excel(w, sheet_name="Orders_SIM", index=False)
        ldf.to_excel(w, sheet_name="Inventory_SIM", index=False)
        pol.to_excel(w, sheet_name="Inventory_Policy_SIM", index=False)
        params.to_excel(w, sheet_name="Assumptions_SIM", index=False)
    return real, odf, ldf


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
