# Residential CRM

This repository contains Google Apps Script modules used to automate the offer generation and document management workflow for Myenergy's residential CRM. The scripts interact with Google Drive, Google Sheets, and Gmail to create folders, generate documents from templates, and send emails.

## Recent Updates:

- **Fingerprint-based Fast-Clone:** Optimized offer processing by introducing a fingerprint-based "fast-clone" mechanism. This allows quick reuse of previous offer calculations and charts, saving computational time and speeding up the offer creation process.

- **API Integration (PVGIS):** Added integration with the PVGIS API to fetch solar panel data, including monthly production and energy savings.

- **New Data Normalization:** Ensured that values like "prezzo energia" (energy price) are consistently normalized using the appropriate locale and formatting functions.

- **Error Handling Improvements:** Robust error handling for invalid inputs and missing data, improving the reliability of the automation process.

- **Optimized Spreadsheet Operations:** Reduced script execution times by optimizing the way data is written to and retrieved from Google Sheets.

## Repository structure

- **gScripts/** – Collection of `.js` files containing Apps Script functions.
  - `main.js` – Main entry point that orchestrates creation of technical data spreadsheets, generation of offer documents, and insertion of charts.
  - `docTemplates.js` – Defines document template IDs and logic to choose which template to use for each opportunity.
  - `docPlaceholders.js` – Builds the placeholder map and replaces placeholders in generated documents.
  - `datiTecnici.js` – Functions for updating the technical data sheet and retrieving calculated results (including handling PVGIS API data).
  - `formatting.js` – Helpers for number, currency, and percentage formatting, as well as hyperlink utilities (including consistent locale formatting for energy price and other values).
  - `insertCharts.js` – Replaces chart placeholders in documents with inline images.
  - `newCharts.js` – Exports charts from the technical spreadsheet to Drive and returns their blobs (also checks if the chart already exists to avoid unnecessary regeneration).
  - `newClientFolders.js` – Creates client folders and stores their URLs in a Google Sheet.
  - `newSubfolders.js` – Utility for creating or retrieving subfolders by name.
  - `duplicateCheck.js` – Scans the CRM spreadsheet for duplicate opportunities and sends a warning email if needed.
  - `WebRouter.js` – Single Web App router (doGet) that routes requests to handlers based on mode (e.g., export, email).
  - `exportBudget.js` – Budget export handler and utilities; reads named ranges (offerta_analizzata, budget) from the technical sheet, generates a CSV (UTF-8 with BOM), saves it to Drive, and triggers a direct browser download.
  - `newEmailOffer.js` – Email handler and helpers for composing personalized offer drafts in Gmail; invoked by WebRouter (mode=email), no standalone doGet.
  - `newld.js` – Handles additional logic for special cases in the offer workflow.
  - `newSubfolders.js` – Utility for creating or retrieving subfolders by name.
  - `offerteOutput.js` – Manages the output of the offer generation process, including storing results and managing offer-related data.
  - `pool_alloc_manager.js` – Manages the allocation of resources and their optimization for offer generation.
  - `precompute.js` – Main module for offer precomputation, including calculation of energy savings, return on investment, and other relevant metrics.
  - `prewarm_offerte.js` – Prewarms the offer generation process by precomputing values for offers that are likely to be needed.
  - `print_result.html` – Template file for rendering the final offer result in HTML format.
  - `yousign.js` – Handles Yousign API integration for e-signature functionality.

## Usage

These scripts are intended for use inside the Google Apps Script environment. To deploy:

1. Create a new Apps Script project and copy the files from `gScripts/` into it.
2. Adjust the hardcoded IDs in `main.js` and related modules (spreadsheet IDs, template IDs, Drive folder IDs) to match your environment.
3. Upload the HTML files from `mail templates/` as HTML templates if you want to send mails using Gmail.
4. Invoke the `main` function (or other utilities) with the required parameters when an offer needs to be generated.

The scripts assume the presence of a spreadsheet acting as the CRM database with sheets named `offerte` and `cronologia`.

### Key Functions:

- **Fast-Clone (Fingerprint-based):** The `fastCloneFromFingerprint` function checks if the offer already exists with the same fingerprint. If found, it clones the values from the existing offer, avoiding recalculating or regenerating data and charts.
  
- **Data Entry and Calculation:** Once an offer is cloned or generated, it stores results in the `offerte_output` sheet, including calculated metrics like energy savings and return on investment.
  
- **PVGIS API Integration:** Uses the `processDatiTecnici` function to fetch solar data from PVGIS, calculating monthly energy production and adjusting the offer accordingly.

- **Document Generation:** Uses placeholders to generate offer documents and presentations, integrating technical data and calculations.

## Developer Guidelines

- **Test Pull Request for Significant Changes:**  
  For major modifications, please create a test pull request. This allows for thorough review and testing before the changes are merged into the main branch.

- **Hiding API Keys:**  
  Ensure that all API keys and sensitive credentials are removed or hidden (for example, using environment variables or configuration files) before pushing your changes.

- **Performance Optimizations:**  
  If you're modifying any parts of the script, particularly around Google Sheets or Drive file manipulations, ensure to use batch operations (`setValues`, `getValues`) and avoid excessive calls to the API.

## License

This repository does not currently include an explicit license.
