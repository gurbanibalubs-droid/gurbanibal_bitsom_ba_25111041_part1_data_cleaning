# Cleaning Log — Part 1: Business Data Cleaning, Validation & Excel Reporting

**Student:** Gurbani Bal | **ID:** BITSOM_BA_25111041
**Dataset:** raw_orders.xlsx (932 raw records)
**Cleaned File:** cleaned_orders.xlsx (912 records after deduplication)

---

## 1. Issues Found

### 1.1 Text Field Issues
- **Extra/leading/trailing spaces** found in: `region` ("  North ", "  East "), `segment` ("Corporate ", "  Small Business "), `customer_name` ("Ananya Rao "), `payment_status` ("Paid "), `order_status` ("Completed "), `category` ("  Furniture ", "Technology "), `ship_mode` ("Second Class ").
- **Case inconsistencies** found in: `region` ("NORTH" instead of "North"), `sub_category` ("LABELS", "storage"), `ship_mode` ("STANDARD CLASS", "FIRST CLASS"), `category` ("Technology " with wrong case).

### 1.2 Date Format Issues
Five different date formats were found within the same columns:
- `21 Jul 2024` (Day MonthName Year)
- `07/27/2024` (MM/DD/YYYY)
- `28-11-2024` (DD-MM-YYYY)
- `2024-05-24` (YYYY-MM-DD — ISO)
- `05 Sep 2024` (Day MonthName Year variant)

**Ship date before order date** — 126 records flagged where ship_date is earlier than order_date.

### 1.3 Duplicate Records
- **Exact duplicate rows:** 20 records were identical copies of existing rows.
- **Conflicting duplicate order IDs:** 12 order IDs appeared more than once after exact duplicates were removed, with differing values in fields such as `sales`, `order_status`, or `payment_status`. These were flagged, not silently removed.

Conflicting IDs flagged for review:
`ORD-2025-10091`, `ORD-2024-10124`, `ORD-2024-10143`, `ORD-2025-10171`, `ORD-2025-10225`, `ORD-2024-10273`, `ORD-2025-10315`, `ORD-2024-10332`, `ORD-2025-10336`, `ORD-2025-10374`, `ORD-2024-10424`, `ORD-2025-10572`

### 1.4 Missing Values
| Field | Count Missing | Action |
|---|---|---|
| region | 25 | Filled with "Unknown" |
| ship_mode | 21 | Filled with "Unknown" |
| discount | 22 | Treated as 0 (if sales fields valid) |
| order_date | 0 | — |
| ship_date | 0 | — |

### 1.5 Invalid Discount Values
- **Negative discounts** found in 11 records (e.g., -0.19, -0.23, -0.09, -0.14, -0.16). Flagged as `negative_discount`.
- **High discounts (>50%)** found in 7 records (e.g., discount = 0.55). Flagged as `high_discount`.

### 1.6 Payment / Order Status Mismatches
- Records with `payment_status = Failed` and `order_status = Completed`: flagged as `failed_payment_completed_order`. These are internally inconsistent and excluded from completed sales summaries.

---

## 2. Cleaning Actions Performed

### 2.1 Text Standardisation
- Applied `.strip()` to all text fields to remove leading and trailing spaces.
- Applied `.title()` (Title Case) to: `segment`, `region`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`, `state`, `city`, `customer_name`.
- This resolved all-caps variants ("NORTH" → "North", "STANDARD CLASS" → "Standard Class", "LABELS" → "Labels", "storage" → "Storage").

### 2.2 Date Cleaning
- Parsed all date values using Python's `dateutil.parser` with both `dayfirst=True` and `dayfirst=False` attempts to handle the variety of formats.
- Standardised all dates to `YYYY-MM-DD` format in `order_date` and `ship_date` columns.
- Computed `shipping_delay_days` as `ship_date - order_date` in days.
- Records where `ship_date_parsed < order_date_parsed` were flagged as `ship_before_order`.

### 2.3 Duplicate Handling
- Exact duplicate rows (all 21 columns identical) were removed — kept first occurrence, removed 20 duplicate rows.
- Conflicting duplicate order IDs (same `order_id`, different field values) were **not removed**. Both records were retained and flagged with `conflicting_duplicate_id` in `data_quality_flag`. Reason: silent deletion of conflicting records could hide data entry errors; they require manual review.

### 2.4 Missing Values
- `region`: filled as "Unknown" where blank or null; flagged in `data_quality_flag`.
- `ship_mode`: filled as "Unknown" where blank or null; flagged in `data_quality_flag`.
- `discount`: filled as 0 where missing (stored in `cleaned_discount`); flagged in `data_quality_flag`.

### 2.5 Discount Cleaning
- Created `cleaned_discount` column: takes `discount` value, replacing nulls with 0.
- Negative discounts and discounts >0.5 are flagged but their original values are preserved in `cleaned_discount` for transparency.
- For `calculated_sales`, discounts were clipped to [0, 1] range to prevent nonsensical calculations.

### 2.6 Calculated Columns Added
| Column | Formula / Logic |
|---|---|
| `cleaned_discount` | `discount` with nulls replaced by 0 |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)` |
| `calculated_profit` | `calculated_sales − cost` |
| `profit_margin` | `calculated_profit ÷ calculated_sales` |
| `shipping_delay_days` | `ship_date − order_date` in days |
| `order_month` | Month extracted from `order_date` |
| `order_year` | Year extracted from `order_date` |
| `data_quality_flag` | "Clean" or "WARNING: <issue list>" |

---

## 3. Business Rules Applied

| Rule | Action Taken |
|---|---|
| Missing `region` | Filled as "Unknown"; flagged in `data_quality_flag` |
| Missing `ship_mode` | Filled as "Unknown"; flagged in `data_quality_flag` |
| Missing `discount` | Treated as 0 in `cleaned_discount`; flagged |
| Negative discount | Flagged as `negative_discount`; value preserved |
| Discount > 50% | Flagged as `high_discount`; value preserved |
| Cancelled orders | Excluded from completed sales pivot summaries |
| Failed payments | Excluded from completed sales pivot summaries |
| Refunded/Returned orders | Summarised separately in pivot report (Sheet: Exceptions by Region) |
| Ship date before order date | Flagged as `ship_before_order`; record retained |
| Exact duplicate rows | First occurrence kept; duplicates removed |
| Conflicting duplicate order IDs | Both records retained; flagged for manual review |
| Failed payment + Completed order | Flagged as `failed_payment_completed_order` |

---

## 4. Assumptions Made

1. **Date parsing ambiguity:** For dates like `06/08/2024`, the parser was run with `dayfirst=True` first. Where the ISO format (`YYYY-MM-DD`) was present, it took precedence as it is unambiguous. Some ship-before-order flags may arise from dates parsed in the wrong order due to this ambiguity; these records are flagged rather than corrected.
2. **Discount threshold:** Discounts above 50% (>0.5) were treated as "high/invalid" based on typical retail business norms. The specific allowed maximum was not stated in the brief, so 50% was used as a reasonable threshold.
3. **Missing discount with valid sales fields:** If quantity, unit_price, sales, cost, and profit were all present and non-null, missing discount was treated as 0 (implying no discount was applied). If other fields were also missing, the record was flagged as invalid.
4. **Conflicting duplicates:** Where the same `order_id` appeared with different values, both records are retained in `cleaned_orders.xlsx` and flagged. The decision of which to keep should be made by a business stakeholder with access to the source systems.
5. **Completed sales definition:** Only records where `order_status = Completed` AND `payment_status = Paid` are included in the main sales and profit pivot summaries.
6. **`calculated_sales` vs reported `sales`:** In many records, the reported `sales` value differs from `quantity × unit_price × (1−discount)`. This may be due to rounding, promotional pricing, or bundled deals. Both values are retained; `calculated_sales` is used for pivot summaries as it is reproducible from first principles.

---

## 5. Records Removed

- **20 exact duplicate rows** were removed (kept first occurrence of each).
- No other records were deleted. All flagged records (invalid discounts, ship-before-order, conflicting duplicates) were retained in `cleaned_orders.xlsx` with appropriate flags.

---

## 6. Records Flagged

| Flag | Count |
|---|---|
| ship_before_order | 126 |
| missing_discount | 22 |
| missing_region | 19 |
| conflicting_duplicate_id | 17 |
| missing_ship_mode | 14 |
| negative_discount | 11 |
| high_discount | 7 |
| failed_payment_completed_order | ~5 |
| Combinations of above | ~19 |
| **Total flagged records** | **240** |

---

## 7. Limitations of the Cleaning Process

1. **Date parsing ambiguity:** Formats like `06/08/2024` could be June 8 (MM/DD) or August 6 (DD/MM). The parser makes a best-effort guess. A significant portion of the 126 "ship before order" flags may be due to date format misinterpretation rather than genuine data errors.
2. **No source system cross-check:** Without access to the original ERP or order management system, it is impossible to verify which of the conflicting duplicate records is correct.
3. **Discount validity range:** The 50% threshold is an assumption. The actual business-allowed discount range was not specified.
4. **Product-level cost data:** No product master was available to validate `unit_price` or `cost` values against reference data.
5. **Customer deduplication:** The same customer name (e.g., "Isha Nair") appears with different `customer_id` values in some cases. This was not resolved as the scope was order-level cleaning, not CRM deduplication.
