# manufacturing-bom-cost-optimizer
Multi-vendor Bill of Materials costing tool built in Excel (.xlsm) with VBA automation. Features dynamic price lookup, tiered discount logic, vendor optimization macros, and JIT inventory tracking - designed for robotics/manufacturing supply chain workflows.

---

## What This Tool Does

Given a product Bill of Materials (BOM), this workbook:

- Pulls unit prices **dynamically** from vendor pricing sheets based on your selected vendor per part
- Applies **tiered volume discounts** automatically (10% at qty ≥ 10, 15% at qty ≥ 25)
- Calculates total build cost in real time as you change vendors or production quantities
- Flags inventory reorder needs based on current stock levels and lead times
- Uses **4 VBA macros** to automate vendor optimization, cost analysis, and reporting

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| `Welcome` | Quick-start guide and workbook overview |
| `VBA Setup Guide` | Step-by-step macro installation instructions |
| `Vendor_A` | Full pricing catalog for Vendor A (14 parts) |
| `Vendor_B` | Full pricing catalog for Vendor B (14 parts) |
| `Vendor_C` | Full pricing catalog for Vendor C (14 parts) |
| `Product_LinearActuator` | BOM costing sheet for a Precision Linear Actuator (11 parts) |
| `Inventory Tracking` | JIT reorder tracker with live status alerts |

---

## How to Use It

**Step 1** — Open `Product_LinearActuator` (or any product BOM sheet)

**Step 2** — Set your production quantity in the yellow cell at row 2 (cell `I2`)

**Step 3** — In the `Selected_Vendor` column (Column C), use the dropdown to choose Vendor_A, Vendor_B, or Vendor_C for each part

**Step 4** — All prices, discounts, and total costs calculate automatically

That's it. No manual lookups, no hardcoded prices.

---

## Vendor Catalog Schema

Each vendor sheet (`Vendor_A`, `Vendor_B`, `Vendor_C`) contains:

| Column | Description |
|---|---|
| `Part_ID` | Unique part identifier (e.g. P001) |
| `Part_Name` | Descriptive part name |
| `Unit_Price` | Base price per unit ($) |
| `Lead_Time_Days` | Vendor fulfillment time |
| `Min_Order_Qty` | Minimum units per order |
| `Discount_10%_Qty` | Quantity threshold for 10% discount |
| `Discount_15%_Qty` | Quantity threshold for 15% discount |

Not every vendor stocks every part — the BOM formulas handle missing parts gracefully using `IFERROR`.

---

## BOM Sheet Formula Logic

The `Product_LinearActuator` sheet uses `INDEX/MATCH` with `INDIRECT` to dynamically look up pricing from whichever vendor is selected per row:

```
Unit_Price  = INDEX(INDIRECT(Selected_Vendor & "!C:C"), MATCH(Part_ID, INDIRECT(Selected_Vendor & "!A:A"), 0))
Discount_%  = IF(Total_Qty >= 15% threshold, 15%, IF(Total_Qty >= 10% threshold, 10%, 0%))
Discounted_Price = Unit_Price × (1 − Discount_%)
Total_Cost  = Total_Qty × Discounted_Price
```

Changing the vendor in Column C instantly recalculates all downstream values for that part row.

---

## Discount Structure

| Total Qty Ordered | Discount Applied |
|---|---|
| < 10 units | 0% |
| ≥ 10 units | 10% |
| ≥ 25 units | 15% |

Discount thresholds are read from each vendor's own sheet, so different vendors can have different break points — the formula accounts for this.

---

## VBA Macros

The file includes 4 macros (see `Costing_Sheet_VBA.txt` for full source code):

### `OptimizeLinearActuator`
Loops through every part in the BOM, looks up prices from all three vendor sheets, applies the correct discount for the current quantity, and automatically sets each row to the cheapest available vendor. Displays a summary popup showing:
- Number of parts changed
- Original total cost vs. optimized total cost
- Total dollar savings and % reduction

### `ResetLinearActuatorToVendorA`
Resets all parts to Vendor A. Useful for establishing a cost baseline before running the optimizer, or for scenario comparisons.

### `ShowVendorBreakdown`
Displays a summary of how many parts are currently sourced from each vendor, plus the current total BOM cost.

### `HighlightExpensiveParts`
Scans the BOM and color-highlights the top 3 most expensive parts by total cost:
- 🔴 #1 most expensive — red highlight
- 🟡 #2 — yellow highlight
- 🟨 #3 — light yellow highlight

---

## Inventory Tracking Sheet

The `Inventory Tracking` sheet provides a just-in-time (JIT) reorder view. For each part it calculates:

- **Reorder Point** = (Monthly Usage / 30) × Lead Time Days + Safety Stock
- **Safety Stock** = 50% of monthly usage
- **Status** = `HEALTHY` / `LOW STOCK` / `REORDER NOW` based on current stock vs. reorder point
- **Action** = recommended order quantity if stock is below the reorder point

A summary row at the bottom counts parts in each status category at a glance.

---

## Setting Up the Macros

1. Open the file — it is already saved as `.xlsm` (macro-enabled)
2. Click **Enable Content** when prompted
3. Press `ALT + F11` to open the VBA Editor
4. Go to **Insert → Module**
5. Open `Costing_Sheet_VBA.txt`, copy the full contents, and paste into the module
6. Close the VBA Editor (`ALT + Q`)
7. Run macros via **Developer → Macros** or assign them to buttons in the sheet

---

## Parts Covered (Linear Actuator BOM)

| Part ID | Part Name |
|---|---|
| P001 | Stepper Motor NEMA 23 |
| P002 | Linear Guide Rail 500mm |
| P003 | Ball Screw Assembly |
| P004 | Motor Coupling 8mm |
| P005 | Limit Switch (Roller) |
| P006 | Aluminum Mounting Bracket |
| P007 | Deep Groove Bearing 608 |
| P008 | M5 Fastener Kit (100pc) |
| P009 | Motor Controller Board |
| P010 | 24V Power Supply 5A |
| P014 | Cable Harness 1m |

---

## Files in This Repo

```
├── Costing_Sheet.xlsm       # Main Excel workbook (macro-enabled)
├── Costing_Sheet_VBA.txt    # VBA source code (plain text, for reference)
└── README.md                # This file
```

---

## Skills Demonstrated

- Excel formula design: `INDEX`, `MATCH`, `INDIRECT`, `IFERROR`, `IF` with nested logic
- Dynamic cross-sheet data retrieval without hardcoded cell references
- Tiered discount logic driven by vendor-specific thresholds
- VBA automation: vendor optimization loop, cost comparison, conditional highlighting
- UX-first design: dropdown menus, clear sheet naming, Welcome tab, and in-sheet instructions
- Inventory management logic: reorder points, safety stock, JIT alerts
- Scalable architecture: adding a new product only requires a new BOM sheet; vendor data stays centralized

---

*Built by Jui Mathuria · [Portfolio](https://juimathuria.vercel.app) · [LinkedIn](https://linkedin.com/in/juimathuria)*

