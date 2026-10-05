# AI Agent Instructions for KashLeo Data Migration

## Context
You are helping migrate KashLeo MFI's data from scattered Excel files to a modern MFS system. The data is in various states of quality and needs to be consolidated into standardized templates.

## Critical Rules

### 1. NEVER Modify Source Files
- Source CSV files in the root directory are READ-ONLY
- Only write to the `final data/` directory
- Preserve original data - never delete or overwrite source files

### 2. Output Location
- ALL transformed/consolidated data MUST be saved to: `C:\Users\admin\Desktop\KashLeo data\final data\`
- Use exact filenames from templates:
  - `mfi-import-clients.csv`
  - `mfi-import-loans.csv`
  - `mfi-import-products.csv`

### 3. Referential Integrity
- Every loan MUST have a matching client in the clients file (via client_id_number)
- Every loan MUST have a matching product in the products file (via product_name)
- If a loan references a missing client/product: NOTE it in a `data_issues.md` file in `final data/`

### 4. Data Validation Before Saving
- Check for duplicate id_numbers in clients file
- Ensure all required fields are populated (use "UNKNOWN" only if truly missing)
- Validate phone numbers (should be numeric, 9-10 digits for Kenya)
- Validate dates (YYYY-MM-DD format preferred)
- Check for empty/null critical fields

## Field Mapping Guidelines

### Client Mapping
| Source Field Pattern | Target Field | Notes |
|---------------------|--------------|-------|
| Id Number / ID No. | id_number | Primary key, clean spaces |
| Client Name (split) | first_name, last_name | Split full name, middle name → last_name |
| Contact / Phone | phone | Clean to digits only |
| Branch | branch | Standardize to mapping table |
| Loan officer | loan_officer_employee_number | Map to employee number |
| Status | client_status | Map: Active→active, Dormant→dormant |
| Residential address | address | Use if available |
| Business type/location | notes | Add if relevant |

### Loan Mapping
| Source Field Pattern | Target Field | Notes |
|---------------------|--------------|-------|
| Loan No / Loan Number | loan_number | Generate if missing: L-{branch_code}-{seq} |
| Id Number / Client ID | client_id_number | MUST match clients.id_number |
| Principal / Amount | principal | Numeric only |
| OLB / Balance | balance | Outstanding loan balance |
| Loan Amount | principal | If principal missing |
| Product / Loan Type | product_name | Map to products.name |
| Interest Rate | interest_rate | Remove %, convert to decimal |
| Duration / Term | term_value | Numeric value |
| Term Unit | term_unit | Map: weeks→weeks, months→months |
| Disbursement Date | disbursed_at | Format as YYYY-MM-DD |
| Branch | branch | Standardize to mapping table |
| Status | status | Map: Active→active, Paid→paid |

### Product Mapping
| Source Field Pattern | Target Field | Notes |
|---------------------|--------------|-------|
| Product Name | name | Use exact name for loan matching |
| Product Code | product_code | Generate if missing: P-{product_name_abbr} |
| Min Loan Amount | min_loan_amount | Numeric |
| Max Loan Amount | max_loan_amount | Numeric |
| Duration (Weeks) | default_term_months | Convert weeks→months (÷4.33) |
| Interest Rate (%) | default_interest_rate | Remove % |
| Interest Type | interest_type | Map: Percentage→flat_rate, Reducing→reducing_balance |
| Application Fee | penalty_amount | Map if applicable |
| Payment Interval | payment_interval_days | Map: weekly→7, monthly→30 |

## Data Quality Rules

### Client Data
- **id_number**: Must be unique, remove spaces/hyphens, keep digits only
- **phone**: Format as 10 digits (e.g., 7123456789), add country code if needed
- **email**: Validate format, leave empty if invalid
- **gender**: Map to male/female/other (case-insensitive)
- **branch**: Use standardized branch names from mapping table
- **client_status**: Only allow: active, dormant, closed

### Loan Data
- **loan_number**: Must be unique across all loans
- **client_id_number**: MUST exist in clients file
- **product_name**: MUST exist in products file
- **principal**: Must be > 0
- **balance**: Must be >= 0 and <= principal
- **interest_rate**: Store as decimal (e.g., 12% → 12.0 or 0.12 depending on type)
- **disbursed_at**: Must be valid date, before maturity_date
- **status**: Only allow: active, paid, closed, written_off

### Product Data
- **name**: Must be unique (used for loan matching)
- **product_code**: Must be unique
- **default_interest_rate**: Must be numeric
- **is_active**: Set to 1 for all products unless explicitly marked inactive

## Standardization Steps

### 1. Create Mapping Tables
Before transformation, create these mapping tables in `final data/`:
- `branch_mapping.csv` - source_branch_name → standardized_branch_name
- `officer_mapping.csv` - source_officer_name → employee_number
- `product_mapping.csv` - source_product_name → target_product_name

### 2. Handle Missing Data
- **Missing clients in loan data**: Create placeholder client records with notes: "Created from loan data - verify"
- **Missing products in loan data**: Create placeholder product records with notes: "Created from loan data - verify"
- **Missing required fields**: Use "UNKNOWN" and note in `data_issues.md`

### 3. Deduplication
- Use id_number as primary key for clients
- If duplicate id_numbers exist: merge records, keep most complete data, note conflicts in `data_issues.md`
- Use loan_number as primary key for loans

## Workflow for Each File Type

### Processing Client Files
1. Read source CSV
2. Extract and clean id_number (primary key)
3. Split full name into first_name/last_name
4. Map branch name using branch_mapping.csv
5. Map loan officer using officer_mapping.csv
6. Fill missing fields where possible (notes, next_of_kin → "UNKNOWN")
7. Append to consolidated clients file
8. Track source file in notes field

### Processing Loan Files
1. Read source CSV
2. Extract client_id_number
3. Verify client exists in consolidated clients (if not, create placeholder)
4. Map product name using product_mapping.csv
5. Calculate derived fields (interest_rate decimal, term conversions)
6. Generate loan_number if missing
7. Append to consolidated loans file
8. Track source file in notes field

### Processing Product Files
1. Read source CSV
2. Standardize product name (remove variations)
3. Generate product_code if missing
4. Convert units (weeks→months, etc.)
5. Set default values for missing fields
6. Append to consolidated products file

## Error Handling
- Create `final data/data_issues.md` to document:
  - Files with structural issues
  - Records with missing critical data
  - Mapping conflicts
  - Data quality concerns
  - Assumptions made during transformation

## Final Validation Checklist
Before considering the migration data complete:
- [ ] All clients have unique id_number
- [ ] All loans have matching client_id_number in clients file
- [ ] All loans have matching product_name in products file
- [ ] No duplicate loan numbers
- [ ] All required fields populated (or marked as UNKNOWN)
- [ ] All dates in valid format
- [ ] All phone numbers standardized
- [ ] Branch names consistent across all files
- [ ] data_issues.md created if any issues found
- [ ] Three final CSV files exist in `final data/`

## Communication
- Always report progress after each major step
- Highlight any data quality issues that require human review
- Ask for clarification on ambiguous mappings
- Provide summary statistics (e.g., "Processed 500 clients, found 23 duplicates")
