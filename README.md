# Logistics Shipment Data Validator & Dashboard

This repository contains an advanced Excel VBA macro designed to automate the reconciliation of global shipment data. It compares internal tracking data (DCT/API) against Logistics Service Provider (LSP) KPI reports, utilizing a standardized Rate Card to calculate expected transit times and flag delays. 

The tool processes raw data, dynamically maps columns, identifies discrepancies, and generates a fully formatted, interactive Excel dashboard complete with charts and toggle filters.

## ✨ Key Features

*   **Dynamic Column Mapping:** Uses fuzzy logic and alias matching to automatically locate required data columns (HAWB, POD, Start Clock, ATD, ATA, Delay Codes) regardless of column order or minor naming variations.
*   **Three-Way Data Reconciliation:** Cross-references data across three distinct files:
    *   Internal System Data (DCT/API)
    *   LSP KPI Reports
    *   Master Rate Card (Transit Times & Destination Mapping)
*   **Intelligent Discrepancy Detection:** Automatically identifies and categorizes:
    *   Missing Shipments (HBL/HAWB) in either system.
    *   Service level mismatches.
    *   Date discrepancies (Start Clock, APD, ATD, ATA, POD).
*   **Automated Delay Dashboard:** Generates a clean, presentation-ready dashboard featuring:
    *   Summary metrics (Match %, Severity Level).
    *   Proportional pie charts for delay scenarios and LSP chargeability.
    *   Delay Code frequency distributions.
*   **Interactive UI:** Includes synchronous toggle buttons built directly into the generated Excel sheets to filter data by specific discrepancy types (e.g., hiding/showing KPI-only delays or Start Clock mismatches).
*   **Lane Mapping Audit:** Compiles a transparency report showing how origin/destination regions and country codes were derived and matched against the master rate card.

## 🚀 How It Works

1.  **Execution:** Run the `CompareShipmentData` subroutine.
2.  **Date Selection:** The user is prompted to input a target month and year (MM/YYYY) to filter the validation scope.
3.  **File Selection:** The script opens standard file dialogs requesting the user to locate the DCT File, the KPI Report, and the Rate Card file.
4.  **Processing:** The macro loads the data into memory using dictionaries for high-speed comparison, skipping regional date bugs, and actively comparing timeline milestones.
5.  **Output:** A brand-new workbook is generated containing:
    *   `Discrepancies`: Tabular breakdown of all mismatched data points.
    *   `Delay Dashboard`: High-level graphical summary.
    *   `Delay Data`: Raw delay logs with interactive filtering buttons.
    *   `Lane Mapping Audit`: Breakdown of transit time logic.

## 🛠️ Requirements

*   Microsoft Excel for Windows (Macro-Enabled)
*   Source files must be standard Excel formats (`.xls`, `.xlsx`, `.xlsm`, `.xlsb`).
*   Requires standard header naming conventions (aliases are built-in, but highly abnormal headers may require updating the `FindSmartHeaderCol` function).

## 💡 Code Structure Overview

*   `CompareShipmentData()`: The main execution subroutine orchestrating file prompts, data extraction, and UI generation.
*   `UpdateDashboardTables()`: Engine that dynamically recalculates and sorts chart data based on user interactions with the toggle buttons.
*   **Helper Functions:**
    *   `FindSmartHeaderCol()`: Intelligent column indexer using string cleaning and aliases.
    *   `ParseSafeDate()`: Custom date parser to bypass Windows regional formatting bugs.
    *   `GetOriginRegion()`: Maps raw postcodes and country codes to predefined logistics zones.
    *   `ToggleKPIFilter()` / `ToggleSCFilter()` / `TogglePODFilter()`: Subroutines attached to the dashboard buttons to control visible table rows and update charts synchronously.
