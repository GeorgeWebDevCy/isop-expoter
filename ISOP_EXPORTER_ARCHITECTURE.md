# ISOP Summer Camp Exporter - Architecture and Dependency Documentation

Generated on: 2026-02-27
Repository path: d:\GitHub Projects\isop-expoter

## 1. Scope

This document explains, in detail, how the custom plugin `ISOP Summer Camp Exporter` works, how it depends on `WooCommerce Extra Product Options` (ThemeComplete EPO), and how the current product options CSV (`current-product-options.csv`) relates to the hardcoded exporter IDs.

This is based on:
- Local plugin code in this repository.
- Dependency plugin code at:
  `C:\Users\georg\Downloads\codecanyon-o5qu7zYy-woocommerce-extra-product-options-wordpress-plugin\woocommerce-tm-extra-product-options`

## 2. Plugin Purpose

`ISOP Summer Camp Exporter` adds an admin page that exports WooCommerce orders into an XLSX file.

The export is designed for a specific summer camp product form structure:
- Parent-level information.
- Up to 6 children in one order.
- Field values stored by ThemeComplete EPO as order item metadata.

The exported file has one row per child (not one row per order).

## 3. High-Level Architecture

### 3.1 Main components

1. WordPress plugin bootstrap and constants:
- `isop-summer-camp-exporter.php`

2. Boilerplate core class and loader:
- `includes/class-isop-summer-camp-exporter.php`
- `includes/class-isop-summer-camp-exporter-loader.php`
- `includes/class-isop-summer-camp-exporter-i18n.php`

3. Admin/public enqueue wrappers (boilerplate):
- `admin/class-isop-summer-camp-exporter-admin.php`
- `public/class-isop-summer-camp-exporter-public.php`

4. Real business logic (all export logic):
- `isop-summer-camp-exporter.php`
  (functions: `get_epo_data`, `get_epo_checkbox`, `get_current_child_data`, `insert_child_into_sheet`, `isop_summer_camp_callback`)

5. Dependency API:
- `THEMECOMPLETE_EPO_API()->get_saved_addons_from_order(...)` from ThemeComplete EPO plugin
  (with fallback to `get_option(...)` for backward compatibility).

### 3.2 Runtime shape

On plugin load:
- Registers activation/deactivation hooks.
- Initializes a boilerplate core class (mostly enqueue + i18n).
- Registers an admin menu page.
- Export logic runs only when the admin page form is submitted.

## 4. Bootstrap and Hook Flow

### 4.1 Bootstrap file behavior (`isop-summer-camp-exporter.php`)

Main bootstrap responsibilities:
- Defines plugin version and many string constants.
- Loads plugin-update-checker and points to GitHub repo/branch.
- Registers activate/deactivate handlers.
- Loads boilerplate core class and runs it.
- Adds admin menu entry: `Isop Summer Camp Exporter`.

Important note:
- Previously this file started with two literal `}` lines before `<?php`.
- This has now been fixed; file start is clean and no early output is emitted from that issue.

### 4.2 Admin menu

The callback is `isop_summer_camp_callback()`.

This function:
- Renders a form with `start-year` and `end-year`.
- On POST, queries WooCommerce orders and streams XLSX download.

## 5. Export Pipeline (Detailed)

### Step 1: Form submit gate

Condition:
- `$_POST['export_orders']`
- `$_POST['start-year']`
- `$_POST['end-year']`

### Step 2: WooCommerce dependency check

- `is_plugin_active('woocommerce/woocommerce.php')`

### Step 3: Query orders

- Uses `wc_get_orders` with:
  - statuses: `completed`, `processing`, `on-hold`
  - sorted by `ID ASC`
  - `date_created`: `startYear-01-01...endYear-12-31`
  - `posts_per_page`: `-1`

### Step 4: Spreadsheet setup

Uses PhpSpreadsheet:
- `Spreadsheet`
- `Writer\Xlsx`

Creates `Orders` sheet with headers `A1..AD1`:
- A Order ID
- B Date
- C Status
- D Customer
- E Total
- F Programme
- G ISOP Student
- H Year Group
- I Name
- J Surname
- K DOB
- L Nationallity (typo in header text)
- M Languages
- N Allergies
- O Swimming allowed
- P Athletic activities allowed
- Q-V Week 1..6
- W Parent Name
- X Parent Phone
- Y Parent Email
- Z Parent Address
- AA Parent Signature
- AB Photo Consent
- AC Marketing Source
- AD Marketing Source Set to Other

### Step 5: Read EPO values for each order

For every order:
- Reads parent + marketing fields by hardcoded EPO IDs.
- Reads child 1..6 fields by hardcoded EPO IDs.
- Builds normalized child arrays via `get_current_child_data`.
- Inserts each child into sheet via `insert_child_into_sheet`.

### Step 6: Child row insertion logic

`insert_child_into_sheet(...)` behavior:
- If `programme` is null, skips row.
- Initializes defaults:
  - all weeks = `No`
  - isop student = `No`
  - swimming = `No`
  - photo consent = `No`
  - year group = `N/A`
- Sets order/customer values.
- Sets child field values when present.
- Sets week flags based on checkbox arrays and `ALL_WEEKS` constant.
- Writes parent/marketing values.
- Increments row pointer.

### Step 7: Output download

- Applies widths and style.
- Sends headers:
  - `Content-Type: application/vnd.ms-excel`
  - `Content-Disposition: attachment;filename="isop-summer-camp-orders.xlsx"`
- Writes to `php://output` and exits.

## 6. Internal Data Model in Exporter

`get_current_child_data(...)` returns:
- programme
- is_isop
- year_group
- weeks_non_isop
- weeks_is_isop
- name
- surname
- dob
- nationality
- langs_spoken
- health
- swimming
- consent
- add
- parent_name
- parent_phone
- parent_address
- parent_email
- parent_sig
- photo

### Value normalization

`strip_euro_recursive` runs on each value:
- Recurses arrays.
- Applies regex to strip text starting at a misencoded Euro marker (`β‚¬...`).

## 7. ThemeComplete EPO Dependency Architecture

## 7.1 Dependency plugin reviewed

Plugin header reviewed:
- Name: Extra Product Options & Add-Ons for WooCommerce
- Version: 7.5.6

## 7.2 Setup and service registration

`Themecomplete_Extra_Product_Options_Setup`:
- defines constants
- includes autoloader and helper functions
- initializes services including main EPO interfaces

`THEMECOMPLETE_EPO_API()` is a singleton accessor returning `THEMECOMPLETE_EPO_API_Base`.

## 7.3 How exporter call resolves

Exporter calls:
- `THEMECOMPLETE_EPO_API()->get_saved_addons_from_order($orderId, $optionId)`
- Fallback path: `get_option(...)` only if older dependency versions do not expose the new method.

In dependency plugin:
- `get_option(...)` is deprecated wrapper.
- `get_saved_addons_from_order(...)` is the current primary method.

`get_saved_addons_from_order(...)`:
- Requires `woocommerce_init` to have fired.
- Iterates order line items.
- Reads `_tmcartepo_data` and `_tmcartfee_data` from order item meta.
- Filters and parses EPO structures.
- Returns grouped data by line item.

## 7.4 Where option `section` IDs come from

When product is added to cart (EPO field classes):
- Field data includes `'section' => $this->element['uniqid']`.

Then at checkout/order creation:
- `_tmcartepo_data` is written to each order line item.

So, the IDs your exporter uses are EPO element `uniqid` values saved as `section` in order meta.

## 8. `current-product-options.csv` Architecture

This CSV is not a simple "one row = one field" file.

It is EPO builder export format:
- Header has many columns for all element prefixes (`selectbox_*`, `radiobuttons_*`, `checkboxes_*`, etc.).
- Rows are index-aligned across those arrays.
- Import logic reconstructs arrays from columns and row indexes.

Repository file stats:
- `current-product-options.csv` has 111 data rows (+1 header).

Detected element counts by `*_uniqid` in this CSV:
- sections: 7
- header: 8
- selectbox: 13
- radiobuttons: 29
- checkboxes: 18
- textfield: 30
- date: 6
- textarea: 6

Sections found:
- Child 1 Section: `69a1695a7c0d34.05018140`
- Child 2 Section: `69a1695a7c0dc3.62297026`
- Child 3 Section: `69a1695a7c0dd1.95839351`
- Child 4 Section: `69a1695a7c0de8.60671284`
- Child 5 Section: `69a1695a7c0df7.84748842`
- Child 6 Section: `69a1695a7c0e04.58855085`
- Parent Info Section: `69a1695a7c0e15.80571741`

## 9. Field ID Inventory From `current-product-options.csv`

The current product option IDs (new schema) for export-relevant fields are:

### 9.1 Child repeated fields (1..6)

- Programme (selectbox):
  - `69a1695a7c0ea9.02082713`
  - `69a1695a7c0ec4.78903872`
  - `69a1695a7c0ee4.83165602`
  - `69a1695a7c0f06.33469928`
  - `69a1695a7c0f24.07286904`
  - `69a1695a7c0f42.44973218`

- Is ISOP (radiobuttons):
  - `69a1695a7c0f74.09504718`
  - `69a1695a7c0fc5.02663835`
  - `69a1695a7c1018.73562668`
  - `69a1695a7c1066.17656952`
  - `69a1695a7c10b6.93091538`
  - `69a1695a7c1106.86287243`

- Year group (selectbox):
  - `69a1695a7c0eb2.21953864`
  - `69a1695a7c0ed1.17217141`
  - `69a1695a7c0ef3.65612751`
  - `69a1695a7c0f17.97629543`
  - `69a1695a7c0f36.81403704`
  - `69a1695a7c0f57.90637022`

- Weeks non-ISOP (checkboxes):
  - `69a1695a7c1140.48717701`
  - `69a1695a7c1172.05618072`
  - `69a1695a7c11a9.13800066`
  - `69a1695a7c1242.51721986`
  - `69a1695a7c1275.63528494`
  - `69a1695a7c12a8.83399846`

- Weeks ISOP (checkboxes):
  - `69a1695a7c1152.89440043`
  - `69a1695a7c1183.12477560`
  - `69a1695a7c11b0.81646841`
  - `69a1695a7c1250.73158171`
  - `69a1695a7c1282.56556619`
  - `69a1695a7c12e8.78249298`

- Name (textfield):
  - `69a1695a7c1308.02829879`
  - `69a1695a7c1344.61271865`
  - `69a1695a7c1385.92716432`
  - `69a1695a7c13c6.55032705`
  - `69a1695a7c1434.04451151`
  - `69a1695a7c1478.76545373`

- Surname (textfield):
  - `69a1695a7c1313.60288492`
  - `69a1695a7c1351.91965440`
  - `69a1695a7c1399.66233074`
  - `69a1695a7c13d4.53869052`
  - `69a1695a7c1441.68392177`
  - `69a1695a7c1489.91474275`

- DOB (date):
  - `69a1695a7c1562.12839538`
  - `69a1695a7c1579.01169817`
  - `69a1695a7c1583.88760547`
  - `69a1695a7c1598.89701528`
  - `69a1695a7c15a7.11185236`
  - `69a1695a7c15b1.51032037`

- Nationality (textfield):
  - `69a1695a7c1326.64112164`
  - `69a1695a7c1369.37220497`
  - `69a1695a7c13a6.45871572`
  - `69a1695a7c1416.46492015`
  - `69a1695a7c1455.71124343`
  - `69a1695a7c14b4.93312802`

- Languages (textfield):
  - `69a1695a7c1332.20935890`
  - `69a1695a7c1370.75713671`
  - `69a1695a7c13b0.07881027`
  - `69a1695a7c1426.13625078`
  - `69a1695a7c1466.28004610`
  - `69a1695a7c14c0.71052850`

- Health/allergies (textarea):
  - `69a1695a7c15c2.52819341`
  - `69a1695a7c15d1.39282728`
  - `69a1695a7c15e7.91607207`
  - `69a1695a7c15f6.30984382`
  - `69a1695a7c1603.23082201`
  - `69a1695a7c1611.87468054`

- Swimming permission (radiobuttons):
  - `69a1695a7c0f89.88311589`
  - `69a1695a7c0fd6.82927620`
  - `69a1695a7c1025.39379772`
  - `69a1695a7c1078.87025028`
  - `69a1695a7c10c6.92621074`
  - `69a1695a7c1119.17109144`

- Parental consent (radiobuttons):
  - `69a1695a7c0f98.23082060`
  - `69a1695a7c0fe0.86844026`
  - `69a1695a7c1039.44641730`
  - `69a1695a7c1088.25742100`
  - `69a1695a7c10d0.21965373`
  - `69a1695a7c1121.51828903`

- Photo consent (radiobuttons):
  - `69a1695a7c0fa6.37872831`
  - `69a1695a7c0ff4.85291424`
  - `69a1695a7c1049.19243106`
  - `69a1695a7c1097.60866743`
  - `69a1695a7c10e0.10325022`
  - `69a1695a7c1130.56580189`

- Add Another Child (radiobuttons, chain control):
  - `69a1695a7c0fb3.44723221`
  - `69a1695a7c1009.11213078`
  - `69a1695a7c1055.75282022`
  - `69a1695a7c10a9.74429260`
  - `69a1695a7c10f6.23402097`

### 9.2 Parent and marketing fields

- Parent name: `69a1695a7c14d6.23089076`
- Parent phone: `69a1695a7c14e4.02544636`
- Parent email: `69a1695a7c14f5.02100431`
- Parent address: `69a1695a7c1503.88503410`
- Parent signature: `69a1695a7c1511.62189457`
- Marketing source: `69a1695a7c0f64.05526525`
- Marketing source other: `69a1695a7c1527.20035702`

## 10. Hardcoded Exporter IDs vs Current CSV IDs

### 10.1 Hardcoded IDs currently in exporter code

Parent + marketing (single IDs):
- `parent_name`: `63c796ae351489.63307542`
- `parent_phone`: `63c796ae351491.52493353`
- `parent_email`: `63c796ae3514a5.64335462`
- `parent_address`: `63c796ae3514b1.17358693`
- `parent_sig`: `63c796ae3514c9.80400881`
- `marketing_source`: `67d085c0924cb1.29085079`
- `marketing_source_other`: `67d0864c924cd2.46491338`

Child repeated hardcoded IDs (1..6):

- Programme:
  - `63c796ae350fd4.47899329`
  - `63c796ae350ff4.86528396`
  - `63c796ae351014.63152225`
  - `63c796ae351037.35378954`
  - `63c796ae351056.35349346`
  - `63c796ae351073.56974688`

- Is ISOP:
  - `63c796ae351098.29269019`
  - `63c796ae3510d6.49644554`
  - `63c796ae351112.16466643`
  - `63c796ae351159.16308749`
  - `63c796ae351191.05095302`
  - `63c796ae3511d9.66370336`

- Year group:
  - `63c796ae350fe9.67195496`
  - `63c796ae351002.93863775`
  - `63c796ae351025.83062078`
  - `63c796ae351042.98702374`
  - `63c796ae351068.59422766`
  - `63c796ae351081.10354514`

- Weeks non-ISOP:
  - `63c796ae351207.25731871`
  - `63c796ae351226.05609817`
  - `63c796ae351244.09025481`
  - `63c796ae351269.16023819`
  - `63c796ae3512c5.48807658`
  - `63c796ae3512e7.89146672`

- Weeks ISOP:
  - `63c796ae351213.50136665`
  - `63c796ae351233.01383284`
  - `63c796ae351255.96533613`
  - `63c796ae351272.45449040`
  - `63c796ae3512d6.33127856`
  - `63c796ae3512f6.58501165`

- Name:
  - `63c796ae351307.70122204`
  - `63c796ae351341.89993507`
  - `63c796ae351384.61975114`
  - `63c796ae3513c2.35501410`
  - `63c796ae351403.28834192`
  - `63c796ae351446.60946329`

- Surname:
  - `63c796ae351316.45003654`
  - `63c796ae351356.44394816`
  - `63c796ae351394.09162182`
  - `63c796ae3513d7.04441495`
  - `63c796ae351418.19452819`
  - `63c796ae351452.03336171`

- DOB:
  - `63c796ae3514d2.17093747`
  - `63c796ae3514e6.76479703`
  - `63c796ae3514f3.26171042`
  - `63c796ae351509.09606580`
  - `63c796ae351513.50958059`
  - `63c796ae351524.28460941`

- Nationality:
  - `63c796ae351320.19034332`
  - `63c796ae351363.69267898`
  - `63c796ae3513a8.83702350`
  - `63c796ae3513e6.31276199`
  - `63c796ae351427.49487978`
  - `63c796ae351467.65291553`

- Languages:
  - `63c796ae351333.14471075`
  - `63c796ae351375.65384919`
  - `63c796ae3513b2.97668616`
  - `63c796ae3513f8.02759799`
  - `63c796ae351435.87962631`
  - `63c796ae351479.57300118`

- Health/allergies:
  - `63c796ae351538.87430147`
  - `63c796ae351544.70768613`
  - `63c796ae351557.74051345`
  - `63c796ae351561.12618878`
  - `63c796ae351572.36626393`
  - `63c796ae351581.38636241`

- Swimming:
  - `63c796ae3510a2.56869018`
  - `63c796ae3510e4.01393417`
  - `63c796ae351129.76215516`
  - `63c796ae351162.38108754`
  - `63c796ae3511a1.96267578`
  - `63c796ae3511e3.43617115`

- Parental consent:
  - `63c796ae3510b4.32887095`
  - `63c796ae3510f5.73100731`
  - `63c796ae351138.74682524`
  - `63c796ae351175.59176972`
  - `63c796ae3511b3.92669862`
  - `63c796ae3511f6.41303666`

- Add child:
  - `63c796ae3510c2.48147773`
  - `63c796ae351105.12345352`
  - `63c796ae351149.05024047`
  - `63c796ae351188.63717837`
  - `63c796ae3511c2.34129641`

- Photo consent:
  - `63cd127a292407.37927849`
  - `63cd1657292438.67347362`
  - `63cd1669292448.40299089`
  - `63cd1681292452.72686890`
  - `63cd1698292467.84694245`
  - `63cd16a9292474.15688992`

Hardcoded IDs in exporter (`isop-summer-camp-exporter.php`) were extracted and compared to IDs in `current-product-options.csv`.

Result:
- Hardcoded exporter IDs: 96
- IDs found in current CSV: 0
- IDs missing from current CSV: 96

Meaning:
- The current exporter is wired to an older EPO schema.
- The current CSV represents a different set of element `uniqid` values.
- Without remapping IDs, this exporter cannot reliably read current form data.

Expected runtime effect of this mismatch:
- `get_epo_data(...)` often returns null.
- Child `programme` null triggers row skip.
- Parent/marketing columns will be blank.
- Export may produce empty or partial sheets even when orders have EPO data.

## 11. Additional Behavior and Risks Found

1. Output before PHP tag:
- Resolved: historical `}` preamble issue at top of plugin file has been removed.

2. Debug output in export path:
- Resolved: debug `echo` and `var_dump` output has been removed from runtime export flow.

3. Week header style range:
- Resolved: header style range now covers all output columns (`A1:AD1`).

4. No nonce/capability hardening on export POST:
- Resolved: export POST now validates capability (`manage_woocommerce`) and nonce.

5. Input sanitization:
- Resolved: years are sanitized, format-validated (`YYYY`), cast to int, and range-checked.

6. Dependency assumptions:
- Resolved: export path now checks ThemeComplete EPO API availability before processing.

7. Deprecated dependency API method:
- Resolved: exporter now uses `get_saved_addons_from_order(...)` with backward-compatible fallback.

8. Potential undefined variable:
- Resolved: `ch6_add` is now explicitly initialized.

9. Data retrieval scope:
- Resolved: `get_epo_checkbox` now iterates all order items before returning results.

## 12. File-Level Architecture Map

Primary custom plugin files:
- `isop-summer-camp-exporter.php`: all export business logic.
- `includes/class-isop-summer-camp-exporter.php`: boilerplate bootstrap class.
- `includes/class-isop-summer-camp-exporter-loader.php`: action/filter registration helper.
- `admin/class-isop-summer-camp-exporter-admin.php`: admin CSS/JS enqueue.
- `public/class-isop-summer-camp-exporter-public.php`: public CSS/JS enqueue.
- `plugin-update-checker/*`: GitHub update mechanism.
- `vendor/*`: PhpSpreadsheet and dependencies.

Dependency files reviewed (ThemeComplete EPO):
- `tm-woo-extra-product-options.php`
- `includes/class-themecomplete-extra-product-options-setup.php`
- `includes/functions/epo-functions.php`
- `includes/classes/class-themecomplete-epo-api-base.php`
- `includes/classes/class-themecomplete-epo-order.php`
- `includes/fields/class-themecomplete-epo-fields.php`
- `admin/class-themecomplete-epo-admin-csv.php`

## 13. Practical Maintenance Notes

To align exporter with `current-product-options.csv`:

1. Replace all hardcoded EPO IDs in exporter with current IDs listed in Section 9.
2. Confirm each child index mapping (child 1..6) exactly.
3. Update parent and marketing IDs.
4. Keep nonce/capability/year validation checks in place during future edits.
5. Keep export output path free from debug output (`echo`, `var_dump`, `print_r`).
6. Keep EPO API access on `get_saved_addons_from_order(...)`; retain fallback only for compatibility.

## 14. Bottom Line

Architecture is straightforward:
- WordPress admin menu -> WooCommerce order query -> EPO option fetch by hardcoded IDs -> row-per-child XLSX output.

The key architectural constraint is ID coupling:
- Exporter correctness depends entirely on exact EPO `uniqid` matches.
- Current product CSV and exporter are currently out of sync (0/96 ID overlap).
