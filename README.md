# KashLeo MFI Data Migration Project

## Overview
This project aims to migrate data from KashLeo Microfinance Institution's legacy Excel-based system to a modern MFS (Microfinance System). The data is currently scattered across multiple Excel files with inconsistent formatting, missing client records, and disconnected loan data.

## Project Goals
1. **Consolidate scattered data** - Merge data from multiple sources (client lists, loan disbursements, product definitions, etc.)
2. **Map to standardized format** - Transform data to match the target MFS import templates
3. **Ensure data integrity** - Identify and resolve missing/inconsistent client records
4. **Clean and validate** - Standardize formats, fix data quality issues, and validate relationships

## Data Sources
- Client lists from all branches (multiple files)
- Loan disbursement reports (monthly by branch)
- Active client reports
- Dormant client reports
- Loan product definitions
- Employee data (loan officers)

## Target Structure
All final consolidated data will be saved in the `final data/` folder with the following standardized CSV files:

### 1. Clients (`mfi-import-clients-template.csv`)
Required fields:
- id_number (unique identifier)
- first_name
- last_name
- phone
- email
- gender
- address
- branch
- loan_officer_employee_number
- next_of_kin_name
- next_of_kin_contact
- client_status (active/dormant/closed)
- notes

### 2. Loans (`mfi-import-loans-template.csv`)
Required fields:
- loan_number
- client_id_number (must match clients.id_number)
- client_phone
- client_number
- product_name (must match products.name)
- principal
- balance
- fees_outstanding
- interest_rate
- term_value
- term_unit (months/weeks)
- status (active/paid/closed)
- disbursed_at
- maturity_date
- branch
- collection_agent_employee_number
- notes

### 3. Products (`mfi-import-products-template.csv`)
Required fields:
- name
- product_code
- description
- default_interest_rate
- default_interest_rate_type (percent/flat)
- default_term_months
- default_term_unit (monthly/weekly)
- interest_type (flat_rate/reducing_balance)
- min_loan_amount
- max_loan_amount
- penalty_amount
- penalty_amount_type
- arrears_penalty_scope
- payment_interval_days
- is_active (1/0)

## Data Challenges
- **Missing clients**: Some clients appear in loan data but not in client lists
- **Inconsistent naming**: Branch names, officer names, product names vary across files
- **Data quality**: Phone numbers, ID numbers may have formatting issues
- **Orphaned records**: Loans referencing non-existent clients or products
- **Duplicate records**: Same client may appear in multiple files

## Workflow
1. **Source Analysis** - Examine each source file to understand structure and content
2. **Data Cleaning** - Standardize formats, remove duplicates, fix inconsistencies
3. **Entity Mapping** - Create mapping tables for branches, officers, products
4. **Data Transformation** - Convert source fields to target template format
5. **Validation** - Ensure referential integrity (loans → clients, loans → products)
6. **Consolidation** - Merge transformed data into final CSV files
7. **Review** - Manual review of consolidated data before import

## Directory Structure
```
KashLeo data/
├── README.md                 # This file
├── AGENT.md                  # AI agent instructions
├── template/                 # Target import templates
│   ├── mfi-import-clients-template.csv
│   ├── mfi-import-loans-template.csv
│   └── mfi-import-products-template.csv
├── draft/                    # Original Excel files (backup)
├── final data/               # Consolidated, migration-ready data (OUTPUT)
│   ├── mfi-import-clients.csv
│   ├── mfi-import-loans.csv
│   └── mfi-import-products.csv
└── [source CSV files]       # Converted source files
```

## Progress
- [x] Convert all Excel files to CSV format
- [x] Move original Excel files to draft/ folder
- [x] Analyze source data structure
- [x] Identify missing clients from loan files
- [x] Create combined client list (kashleo_clients.csv)
- [ ] Create entity mappings (branches, officers, products)
- [ ] Transform and consolidate client data to template format
- [ ] Transform and consolidate loan data
- [ ] Transform and consolidate product data
- [ ] Validate referential integrity
- [ ] Generate final import files in `final data/`
