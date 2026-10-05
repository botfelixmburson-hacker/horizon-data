# Data Migration Issues Log

This file tracks all data quality issues, conflicts, and assumptions made during the KashLeo MFI data migration process.

## Format
- **Date**: YYYY-MM-DD
- **Source File**: Name of file where issue was found
- **Issue Type**: [Missing Data] [Duplicate] [Mapping Conflict] [Data Quality] [Assumption]
- **Description**: Detailed description of the issue
- **Resolution**: How it was handled or action required

---

## Issues Log

### 2026-10-05

**Issue Type**: [Missing Data]
**Source File**: Multiple loan disbursement files
**Description**: Found 214 clients in loan files who are not present in the main client list (all_clients_from__all_branches_145338.csv)
**Resolution**: Created "missing clients from list.csv" in final data/ folder. Contains client names, contacts, IDs, and source files for manual review and addition to client database.

**Issue Type**: [Data Quality]
**Source File**: ENGINEER BRANCH DISBURSEMENT files
**Description**: 114 of the 214 missing clients are from Engineer branch, confirming the suspicion that Engineer branch clients are not in the main system
**Resolution**: These clients need to be added to the consolidated client list. Many have ID numbers but are missing from the master client file.

**Issue Type**: [Duplicate]
**Source File**: ENGINEER BRANCH DISBURSEMENT & OLB Report (1) (3) (1).csv
**Description**: Found duplicate entries in Engineer branch loan files (e.g., "Mose Bahati Bevlyne" appears multiple times with same ID 32739955)
**Resolution**: Need to deduplicate during consolidation. Keep the most complete record.

**Analysis Summary**:
- Total clients in main list: 790
- Total unique clients in loan files: 221
- Missing clients (not in main list): 214
- After deduplication: 153 unique missing clients added (61 duplicates merged)
- Engineer branch missing clients: 114 (53% of missing)
- Final combined client list: 943 clients

**Combined Client List Created**:
- File: `kashleo_clients.csv`
- Total clients: 943
- From main list: 790
- From loan files (new additions): 153
- Engineer branch in combined list: 53 clients

**Guarantor Information Added**:
- Extracted guarantor data from 6 loan disbursement files
- Added guarantor columns: Guarantor Name, Guarantor Contact, Guarantor ID
- Clients with guarantor information: 153 out of 943 (16%)
- Note: Clients without active loans in the analyzed files have no guarantor data

**Phone Number Standardization**:
- Renamed "Contact" column to "Phone"
- Standardized phone formats to Kenyan format (0XXXXXXXXX):
  - 9-digit numbers → prefixed with "0"
  - 10-digit numbers starting with "254" → converted to "0XXXXXXXXX"
  - 12-digit numbers starting with "254" → converted to "0XXXXXXXXX"
- Final distribution: 909 clients with 10-digit format, 1 with 7-digit format
- Guarantor contact numbers also standardized

**Missing ID Numbers**:
- Found 100 clients missing ID numbers (10.6% of total)
- Created `clients_missing_id.csv` file listing these clients
- These clients need manual ID number collection before system import
- Most missing IDs are from "Unknown" branch clients added from loan files

**Engineer Branch Data Enhancement**:
- Extracted complete data from ENGINEER BRANCH DISBURSEMENT & OLB Report (1) (3) (1).csv
- Updated all 53 Engineer branch clients with missing phone numbers
- Added guarantor information for all Engineer branch clients
- Engineer branch now has 100% phone number coverage and guarantor data
- One ID mismatch found (315103 vs 31510377) - manually corrected

**Joseph's Client List Created**:
- File: `joseph_clients.csv`
- Total clients: 30
- Loan officers identified: Joseph Ngondi (Karatina), Joseph (Nyahururu)
- Branch distribution: Karatina (29 clients), Nyahururu (1 client)
- Status distribution: Dormant (29), Active (1)
- Clients with guarantor information: 1 (only active client from loan files)
- **Critical limitation**: Most clients (29/30) are dormant and have no recent loans in available files
- Available loan files only cover July-October 2026 disbursements
- Historical loan data with guarantor information for dormant clients is not available in current files
- dormant_clients_140854.csv is empty/invalid - cannot extract historical data
- Need to obtain historical loan records to populate guarantor data for dormant clients

**Guarantor Data Status in kashleo_clients_complete.csv**:
- Total clients: 943
- Clients with guarantor information: 154 (16.3%)
- Clients without guarantor information: 789 (83.7%)
- Guarantor data extracted from 6 loan disbursement files (July-October 2026)
- All 154 clients with guarantor info are those with recent active loans
- 789 dormant clients lack guarantor data because historical loan files are not available
- Current files only contain recent disbursement data (4 months)
- **Root cause**: Historical loan disbursement files containing guarantor information for dormant clients are missing from the dataset
