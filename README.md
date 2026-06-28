# Part 1: Business Data Cleaning, Validation & Excel Reporting

**Student:** Gurbani Bal
**Student ID:** BITSOM_BA_25111041
**Program:** PGDM General Management — XLRI Jamshedpur (Batch 2026–27)
**Repository:** `gurbanibal_bitsom_ba_25111041_part1_data_cleaning`

---

## Problem Summary

A retail company exports order-level sales data from multiple internal systems. The raw dataset contained a range of data quality issues including inconsistent date formats, text case mismatches, duplicate records, missing values, invalid discounts, and order status inconsistencies. The objective of this assignment was to clean and validate the dataset, apply business rules, document all findings, and produce summary reports ready for business review.

---

## Dataset Description

| Attribute | Detail |
|---|---|
| File | `raw_orders.xlsx` |
| Sheet | `raw_orders` |
| Raw records | 932 rows |
| Columns | 21 (order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city, category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status) |
| Date range | 2024–2025 |
| Regions covered | North, South, East, West |
| Categories | Furniture, Office Supplies, Technology |

---

## Tools Used

- **Microsoft Excel** — manual inspection, screenshots
- **Python 3** (pandas, openpyxl, dateutil) — programmatic cleaning, calculated columns, and report generation
- **GitHub** — version control and submission

---

## Cleaning Steps Performed

### Step 1: Text Field Standardisation
Applied to fields: `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`.

- Removed all leading and trailing whitespace
- Applied Title Case to resolve all-caps variants (e.g., "NORTH" → "North", "STANDARD CLASS" → "Standard Class", "LABELS" → "Labels")
- Corrected mixed-case values such as "storage" → "Storage" and "FIRST CLASS" → "First Class"

### Step 2: Date Cleaning and Standardisation
The raw data contained five distinct date formats across `order_date` and `ship_date`:

- `21 Jul 2024` — Day MonthName Year
- `07/27/2024` — MM/DD/YYYY
- `28-11-2024` — DD-MM-YYYY
- `2024-05-24` — ISO (YYYY-MM-DD)
- `05 Sep 2024` — Day MonthName Year (variant)

All dates were parsed and standardised to `YYYY-MM-DD`. A new column `shipping_delay_days` was computed as the difference between `ship_date` and `order_date`.

### Step 3: Duplicate Handling
- Identified and removed **20 exact duplicate rows** (all 21 columns identical). First occurrence was kept.
- After removal, **12 order IDs** were found to appear more than once with conflicting field values (different `sales`, `order_status`, or `payment_status`). These were **flagged, not deleted**, as silent removal could mask data entry errors.

### Step 4: Missing Value Treatment
- `region` (25 records): filled with "Unknown" and flagged
- `ship_mode` (21 records): filled with "Unknown" and flagged
- `discount` (22 records): treated as 0 where all other sales fields were valid; flagged

### Step 5: Business Rule Application
All business rules from the `business_rules` sheet were applied. See section below for details.

### Step 6: Calculated Columns
Eight new columns were added to `cleaned_orders.xlsx`:

| Column | Logic |
|---|---|
| `cleaned_discount` | `discount` with nulls replaced by 0 |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)` |
| `calculated_profit` | `calculated_sales − cost` |
| `profit_margin` | `calculated_profit ÷ calculated_sales` |
| `shipping_delay_days` | `ship_date − order_date` in days |
| `order_month` | Month number from `order_date` |
| `order_year` | Year from `order_date` |
| `data_quality_flag` | "Clean" or "WARNING: issue1 \| issue2 ..." |

---

## Business Rules Applied

| Rule Area | Action Taken |
|---|---|
| Missing `region` | Filled as "Unknown"; flagged in `data_quality_flag` |
| Missing `ship_mode` | Filled as "Unknown"; flagged in `data_quality_flag` |
| Missing `discount` | Treated as 0 if other sales fields valid; flagged |
| Negative discount | Flagged as `negative_discount`; original value preserved |
| Discount > 50% | Flagged as `high_discount` |
| Cancelled orders | Excluded from completed sales summaries |
| Failed payments | Excluded from completed sales summaries |
| Refunded / Returned orders | Summarised separately in pivot report |
| Ship date before order date | Flagged as `ship_before_order` |
| Exact duplicate rows | Removed (kept first occurrence) |
| Conflicting duplicate order IDs | Both records retained and flagged for manual review |
| Failed payment + Completed order | Flagged as `failed_payment_completed_order` |

---

## Summary of Data Quality Issues Found

| Issue | Count |
|---|---|
| Raw records | 932 |
| Exact duplicates removed | 20 |
| Records in cleaned file | 912 |
| Clean records | 672 |
| Total flagged records | 240 |
| Ship date before order date | 126 |
| Missing discount | 22 |
| Missing region (filled Unknown) | 25 |
| Missing ship_mode (filled Unknown) | 21 |
| Conflicting duplicate order IDs | 12 unique IDs (24 records flagged) |
| Negative discounts | 11 |
| High discounts (>50%) | 7 |
| Failed payment + Completed status | ~5 |

---

## Summary of Final Pivot Reports

**pivot_summary.xlsx** contains six sheets:

| Sheet | Description |
|---|---|
| Sales by Region | Total sales, profit, and profit margin by region — sorted by sales descending |
| Sales by Category | Sales and profit broken down by category and sub-category |
| Orders by Ship Mode | Order count and total sales by shipping method |
| Margin by Segment | Profit margin comparison across Consumer, Corporate, Small Business, Home Office |
| Exceptions by Region | Count and sales value of Cancelled, Returned, and Failed Payment orders by region |
| Monthly Sales Trend | Month-by-month sales, profit, and order count for completed paid orders |

---

## Key Business Insights

1. **South and West regions** generate the highest sales volumes among completed paid orders.
2. **Technology** is the highest-revenue category; within it, **Phones** and **Copiers** dominate.
3. **Standard Class** and **Second Class** shipping are the most frequently used modes, while **Same Day** is used for a smaller but significant share.
4. **Home Office** and **Consumer** segments show marginally higher profit margins compared to Corporate and Small Business in the valid completed orders.
5. **126 records** have ship dates that precede order dates — this is the single largest data quality issue and warrants investigation with the logistics team.
6. A meaningful share of orders with **Failed** payment status were also recorded as **Completed** in `order_status`, suggesting a sync issue between the payment and order management systems.
7. **Refunded and returned orders** are distributed across all regions, with no single region showing a disproportionately high return rate based on the available data.

---

## Assumptions and Limitations

### Assumptions
- Dates in ambiguous formats (e.g., `06/08/2024`) were parsed with `dayfirst=True` as the default; ISO-formatted dates (`YYYY-MM-DD`) were treated as unambiguous.
- Discounts above 50% were treated as invalid based on standard retail norms. The actual business threshold was not specified in the brief.
- Missing discounts were treated as 0 only where all other order fields (quantity, unit_price, cost, profit) were present and non-null.
- Only records with `order_status = Completed` AND `payment_status = Paid` were included in sales pivot summaries.
- `calculated_sales` (derived from quantity, unit_price, and cleaned_discount) was used in pivot summaries rather than the reported `sales` column, as it is reproducible from first principles.

### Limitations
- Some of the 126 "ship before order" flags may be false positives caused by date format ambiguity (MM/DD vs DD/MM parsing).
- Conflicting duplicate records cannot be resolved without access to the source ERP system; they have been flagged for stakeholder review.
- No product master data was available to validate unit prices or cost figures.
- The same customer name sometimes appears with different `customer_id` values; customer-level deduplication was outside the scope of this assignment.

---

## Screenshots

| File | Contents |
|---|---|
| `screenshots/raw_data_preview.png` | Raw dataset before cleaning (first rows visible in Excel) |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset with all calculated columns |
| `screenshots/pivot_summary_1.png` | Sales & Profit by Region pivot |
| `screenshots/pivot_summary_2.png` | Monthly Sales Trend pivot |

---

## Repository Structure

```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx          # Original dataset — unchanged
│   └── cleaned_orders.xlsx      # Cleaned dataset with calculated columns
├── outputs/
│   ├── data_quality_report.xlsx # 8-sheet quality report
│   ├── pivot_summary.xlsx       # 6-sheet pivot summary
│   └── cleaning_log.md          # Full documentation of cleaning decisions
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
└── README.md
```
