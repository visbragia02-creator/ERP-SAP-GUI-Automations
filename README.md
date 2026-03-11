# ERP-SAP-GUI-Automations
High-efficiency SAP automation scripts for enterprise process optimization.
# SAP Pre-Upload Validation Engine 🚀

## 📌 The Business Problem
Uploading mass financial data (Journal Entries, Accruals, Allocations) into complex ERPs like SAP is highly prone to human error. A single missing Cost Object, imbalanced debit/credit, or incorrect posting key can cause upload failures, system errors, and hours of manual reconciliation during critical month-end closing periods.

## 💡 The Solution
I developed this VBA-based Validation Engine to act as a strict gatekeeper between raw financial data and the ERP system. It automatically scans thousands of rows against dozens of complex corporate business rules, highlighting errors before the data ever reaches SAP.

## ⚙️ Core Technical Features
* **Modular Architecture:** The main controller (`ValidationTool`) orchestrates over 25 distinct validation subroutines, making the code scalable and easy to maintain.
* **Double-Entry Verification:** Automatically identifies blocks of journal entries and ensures debits (40) and credits (50) balance perfectly.
* **Dynamic Cost Object Logic:** Verifies the existence and uniqueness of Cost Centers and Internal Orders based on specific GL account prefixes (e.g., Operating Category accounts).
* **Conditional Formatting & UI:** Visually flags errors (using RGB color coding) and generates specific alert messages to guide the end-user on how to fix the data.
* **Data Sanitization:** Enforces string formatting, uppercase conversion, zero-padding for partner codes, and date validation.

## 📈 Business Impact
* **Risk Mitigation:** Prevents invalid data from contaminating the corporate ERP.
* **Time-Saving:** Reduces hours of manual data checking and failed SAP upload troubleshooting into a process that takes seconds.
* **Process Standardization:** Enforces strict adherence to corporate accounting structures and reporting guidelines.

---
*Note: Sensitive corporate data, network paths, and specific company codes have been redacted for confidentiality.*

# Corporate Template Auto-Converter 🔄

## 📌 The Business Problem
When corporate reporting standards change or systems are upgraded, migrating historical data and daily reports into new templates becomes a massive operational bottleneck. 
Manually copying, pasting, and reformatting data from legacy structures into new standard layouts drains analytical resources and introduces a high risk of human error.

## 💡 The Solution
I developed a "One-Click" Template Converter that seamlessly transforms deprecated legacy formats into the new corporate standard. 
Instead of manual data entry, the user simply clicks a button, and the tool extracts the raw data, maps it to the correct new fields, and applies all required structural rules instantly.

## ⚙️ Core Technical Features
* **Automated Data Mapping:** Intelligently parses data from the legacy file and routes it to the exact new cell locations, dynamically handling structural changes between the old and new versions.
* **One-Click Execution:** Designed with the end-user in mind, providing a frictionless UI that allows any team member to convert complex files without technical knowledge.
* **Format Standardization:** Automatically applies the new corporate formatting (fonts, number formats, date structures, and cell alignments) to ensure 100% compliance with new guidelines.
* **Data Integrity & Cleanup:** Cleanses legacy data during the transfer, ensuring no information is lost or incorrectly formatted during the structural shift.

## 📈 Business Impact
* **Massive Time Savings:** Reduced a tedious, multi-minute manual process per file down to a fraction of a second.
* **Zero Error Margin:** Eliminated copy-paste errors and structural misalignments associated with manual transitions.
* **Frictionless Change Management:** Enabled the operations team to adopt the new reporting standard immediately without a painful transition period or a backlog of conversion work.
---
*Note: Specific corporate templates and proprietary data structures have been abstracted for confidentiality.*

 
