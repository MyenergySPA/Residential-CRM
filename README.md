# Residential CRM

This repository contains Google Apps Script modules used to automate the offer generation and document management workflow for Myenergy's residential CRM. The scripts interact with Google Drive, Google Sheets and Gmail to create folders, generate documents from templates and send e‑mails.

## Repository structure

- **gScripts/** – collection of `.js` files containing Apps Script functions.
  - `main.js` – main entry point that orchestrates creation of technical data spreadsheets, generation of offer documents and insertion of charts.
  - `docTemplates.js` – defines document template IDs and logic to choose which template to use for each opportunity.
  - `docPlaceholders.js` – builds the placeholder map and replaces placeholders in generated documents.
  - `datiTecnici.js` – functions for updating the technical data sheet and retrieving calculated results.
  - `formatting.js` – helpers for number, currency and percentage formatting as well as hyperlink utilities.
  - `insertCharts.js` – replaces chart placeholders in documents with inline images.
  - `newCharts.js` – exports charts from the technical spreadsheet to Drive and returns their blobs.
  - `newClientFolders.js` – creates client folders and stores their URLs in a Google Sheet.
  - `newSubfolders.js` – utility for creating or retrieving subfolders by name.
  - `duplicateCheck.js` – scans the CRM spreadsheet for duplicate opportunities and sends a warning email if needed.
  - `newEmailOffer.js` – small webapp and helper function for composing personalised offer emails.
- **mail templates/** – HTML files used as email bodies (`offerta.html`, `recensione.html`, etc.).

## Usage

These scripts are intended for use inside the Google Apps Script environment. To deploy:

1. Create a new Apps Script project and copy the files from `gScripts/` into it.
2. Adjust the hardcoded IDs in `main.js` and related modules (spreadsheet IDs, template IDs, Drive folder IDs) to match your environment.
3. Upload the HTML files from `mail templates/` as HTML templates if you want to send mails using Gmail.
4. Invoke the `main` function (or other utilities) with the required parameters when an offer needs to be generated.

The scripts assume the presence of a spreadsheet acting as the CRM database with sheets named `offerte` and `cronologia`.

## Developer Guidelines

- **Test Pull Request for Significant Changes:**  
  For major modifications, please create a test pull request. This allows for thorough review and testing before the changes are merged into the main branch.

- **Hiding API Keys:**  
  Ensure that all API keys and sensitive credentials are removed or hidden (for example, using environment variables or configuration files) before pushing your changes.

## License

This repository does not currently include an explicit license.
