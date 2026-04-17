# Order Management Spreadsheet

A YAML-driven Excel workbook for managing customers, suppliers, products,
sales orders, purchase orders, payments, and inventory — with VAT support
and a dashboard.

The workbook is **generated** from a `schema.yml` file by a Python build
script. The schema is the source of truth. The `.xlsx` is disposable —
regenerate it whenever the schema changes.

-----

## Quick start

```bash
# 1. Install dependencies
pip install openpyxl pyyaml

# 2. Generate the workbook
python build.py

# 3. Open OrderManagement.xlsx in Excel, LibreOffice, or Google Sheets
```

Optional arguments:

```bash
python build.py my_schema.yml my_output.xlsx
```

-----

## What’s in the workbook

|Sheet                 |Role                                                        |Type    |
|----------------------|------------------------------------------------------------|--------|
|**Dashboard**         |KPIs and status breakdowns                                  |View    |
|**Customers**         |People and companies you sell to                            |Table   |
|**Suppliers**         |People and companies you buy from                           |Table   |
|**Products**          |Shared catalogue with `kind` = Sellable / Purchasable / Both|Table   |
|**SalesOrders**       |Customer order headers (auto-totals from lines)             |Table   |
|**SalesOrderLines**   |Line items on sales orders                                  |Table   |
|**PurchaseOrders**    |Supplier order headers (auto-totals from lines)             |Table   |
|**PurchaseOrderLines**|Line items on purchase orders                               |Table   |
|**Payments**          |Incoming and outgoing payments                              |Table   |
|**InventoryMovements**|Stock receipts and dispatches                               |Table   |
|**StockLevels**       |Current stock per product (formula-driven)                  |View    |
|`_VatRates`, `_Enums` |Hidden reference sheets                                     |Internal|

Each data sheet is a real **Excel Table** (`Ctrl+T` style), which gives you:

- Auto-filter dropdowns on every column
- Named ranges for use in formulas (`Customers[name]`, etc.)
- Row striping and automatic formatting
- Structured references when entering your own formulas

-----

## The relational model

### Cell-colour convention

|Colour          |Meaning                                   |
|----------------|------------------------------------------|
|**Blue text**   |Input — type your values here             |
|**Black text**  |Formula — don’t edit                      |
|**Green italic**|Lookup — auto-populated from related table|

### Foreign key behaviour

When you enter a row in `SalesOrders`:

1. Click the `customer_id` cell → a dropdown appears listing every `customer_id` from the `Customers` table
1. Pick one → the `customer_name` column auto-populates via an `INDEX/MATCH` formula
1. The order’s `subtotal`, `vat_total`, `grand_total`, `amount_paid`, and
   `balance_due` are auto-calculated from `SalesOrderLines` and `Payments`

The same pattern applies to:

- `SalesOrderLines.order_id` → `SalesOrders.order_id`
- `SalesOrderLines.product_id` → `Products.product_id`
- `PurchaseOrders.supplier_id` → `Suppliers.supplier_id`
- `PurchaseOrderLines.po_id` → `PurchaseOrders.po_id`
- `PurchaseOrderLines.product_id` → `Products.product_id`
- `Payments.reference_id` → either `SalesOrders.order_id` *or* `PurchaseOrders.po_id` (polymorphic — see below)
- `InventoryMovements.product_id` → `Products.product_id`

### Polymorphic reference — Payments

The `Payments` table has a `reference_id` column that can hold **either** a
sales order ID (`SO-0001`) **or** a purchase order ID (`PO-0001`). The
`direction` column disambiguates: `In` means it’s a customer payment
against a sales order; `Out` means it’s a payment to a supplier against a
purchase order.

The order headers’ `amount_paid` formula filters Payments by both the ID
and the direction:

```
=SUMIFS(Payments[amount],
        Payments[reference_id], [@order_id],
        Payments[direction], "In")
```

### VAT handling

Line items carry a `vat_code` (STD / RED / ZER / EXM / OSC / RCG / RCS).
The `line_vat` formula looks up the rate from the named range `VatRates`:

```
=line_net * VLOOKUP(vat_code, VatRates, 2, FALSE)
```

Rates live in the hidden `_VatRates` sheet, which is populated from the
`vat_rates:` section of `schema.yml`. Change a rate in the YAML and
regenerate.

The Dashboard shows a **VAT owed to HMRC** figure computed as
`sales VAT − purchase VAT`. This is a net-VAT position, not a filed
return.

-----

## How to modify the model

All structural changes go in `schema.yml`. Don’t edit the `.xlsx` for
anything schema-related — your changes will be lost next time you
regenerate. Use the workbook for **data** entry only.

### 1. Add a field to an existing table

Open `schema.yml`, find the table, and add a line under `fields:`:

```yaml
- name: Customers
  primary_key: customer_id
  fields:
    - {name: customer_id, type: id, prefix: "CUST-", width: 12}
    - {name: name, type: text, required: true, width: 30}
    - {name: tax_exempt, type: bool, default: "No", width: 12}   # ← new
```

Save, run `python build.py`, done. The new column appears on the
Customers sheet with a Yes/No dropdown.

**Watch out:** if you have existing data in `OrderManagement.xlsx`,
regenerating will overwrite it — the generator produces a fresh workbook.
Keep your data in YAML sample data for rebuilds, or edit the generator to
merge with an existing file (not implemented by default).

### 2. Add a whole new table

```yaml
tables:
  # ...existing tables...
  
  - name: Contacts
    description: Additional contacts at customers and suppliers
    primary_key: contact_id
    tab_color: "3B82F6"
    fields:
      - {name: contact_id, type: id, prefix: "CON-", width: 12}
      - {name: customer_id, type: fk, references: "Customers.customer_id", width: 12}
      - {name: customer_name, type: lookup, source: "Customers", key: customer_id, field: name, width: 28}
      - {name: name, type: text, width: 25}
      - {name: role, type: text, width: 20}
      - {name: email, type: text, width: 28}
      - {name: phone, type: text, width: 16}
```

### 3. Add an enum (dropdown values)

Under the `enums:` section:

```yaml
enums:
  # ...existing...
  contact_role:
    - Buyer
    - Accounts
    - Technical
    - Other
```

Then reference it in a field:

```yaml
- {name: role, type: enum, enum: contact_role, width: 14}
```

### 4. Change a VAT rate

```yaml
vat_rates:
  STD: 0.22    # changed from 0.20
  RED: 0.05
  ZER: 0.00
```

All formulas update on regeneration.

### 5. Change the number of rows per table

```yaml
workbook:
  default_rows: 1000    # was 500
```

### 6. Disable sample data

```yaml
workbook:
  sample_data: false
```

The tables will be empty but the structure remains.

-----

## Field type reference

|Type      |Purpose                        |Notes                                                                |
|----------|-------------------------------|---------------------------------------------------------------------|
|`id`      |Primary key                    |Sample data provides values; you type your own for new rows          |
|`fk`      |Foreign key                    |Renders as a dropdown of valid IDs from the referenced table         |
|`text`    |Free text                      |                                                                     |
|`number`  |Numeric                        |                                                                     |
|`currency`|GBP amount                     |Formatted as `£1,234.56`                                             |
|`percent` |Percentage                     |Enter `0.1` for 10%; formatted as `10.0%`                            |
|`date`    |Date                           |DD/MM/YYYY format                                                    |
|`enum`    |Fixed dropdown                 |References a list under `enums:`                                     |
|`bool`    |Yes/No dropdown                |                                                                     |
|`formula` |Excel formula                  |Uses `[@field]` for same-row refs; `TableName[field]` for column refs|
|`lookup`  |Auto-display from related table|Specify `source`, `key`, `field`                                     |

### Formula syntax

Formulas in the YAML use two kinds of placeholder:

- `[@field]` — the value of `field` in the current row (same table)
- `TableName[field]` — the entire column in another table

The build script translates these into concrete Excel cell references
(like `E5` and `SalesOrderLines!$E$5:$E$504`) when writing the workbook.

Example — computing the net amount on a sales order line:

```yaml
- {name: line_net, type: formula,
   expr: '=[@quantity]*[@unit_price]*(1-[@discount_pct])',
   format: currency}
```

Example — summing child rows back to parent:

```yaml
- {name: subtotal, type: formula,
   expr: '=SUMIFS(SalesOrderLines[line_net],SalesOrderLines[order_id],[@order_id])',
   format: currency}
```

-----

## Editing the workbook directly (data entry)

Once generated, use the `.xlsx` as a normal workbook:

1. Go to the relevant table sheet
1. Type into the first blank row (default-valued cells are pre-filled in blue)
1. Use dropdowns for foreign keys, enums, and Yes/No flags
1. Watch the Dashboard update automatically

**Don’t** add columns or rename them in Excel — those changes vanish on
the next regeneration. Structural changes live in YAML.

-----

## Architecture

```
schema.yml          ← Source of truth (edit this)
    │
    ▼
build.py            ← Reads YAML, writes Excel
    │
    ▼
OrderManagement.xlsx  ← Generated output (don't edit structure)
```

The generator:

1. Loads `schema.yml`
1. Computes metadata for every table (field positions, row ranges)
1. Creates hidden reference sheets (`_VatRates`, `_Enums`)
1. For each table, writes a worksheet with headers, sample data,
   per-row formulas and lookups, an Excel Table definition, and data
   validation rules
1. Writes the Dashboard with cross-table KPI formulas
1. Writes the StockLevels view with on-hand calculations

All formulas use direct cell references (`=E5*F5`) rather than structured
references (`=[@quantity]*[@unit_price]`). This makes the workbook
portable across Excel, LibreOffice, and Google Sheets — structured
references don’t always render consistently when a file is generated
outside Excel itself.

-----

## Limitations

- **Regeneration overwrites data.** The script produces a clean workbook
  each time. If you enter data directly in `.xlsx` and then regenerate,
  your data is lost unless the sample data in the YAML includes it.
  For production use, consider a merge step that preserves existing rows
  from the previous workbook — not included in this version.
- **No enforced referential integrity.** Excel’s data validation warns
  against invalid foreign keys but doesn’t prevent them if dropdown
  checking is disabled. This is an Excel limitation, not a code issue.
- **No cascading updates.** If you rename a customer’s name, existing
  orders still show the old name (because `customer_name` on
  `SalesOrders` is a lookup, it will refresh automatically — but
  any text you’ve typed referencing the name stays as-is).
- **Default 500 rows per table.** Change `default_rows` in YAML if
  you expect more.
- **No audit trail.** Edits to rows don’t leave a history.

-----

## File layout

```
ordermgmt/
├── schema.yml              # Relational model definition
├── build.py                # Generator script
├── OrderManagement.xlsx    # Generated workbook (output)
└── README.md               # This file
```
