# Tiller Quick Search

Google Apps Script sidebar for **Tiller** spreadsheets that adds a Quick Search panel to filter transactions by date, amount, description, account, and category. Uses a basic filter and an ARRAYFORMULA-based Match column on the Transactions sheet.

## Requirements

- Google Sheets with a **Transactions** sheet (Tiller-style column headers: Date, Description, Category, Amount, Account)
- Apps Script project with **Sheets API v4** enabled (Advanced Service)

## Installation

1. In your Google Sheet: **Extensions → Apps Script**
2. Create or open a project and add the script files (or use [clasp](https://github.com/google/clasp) to push from this repo)
3. Enable the **Sheets** advanced service: Project Settings → Advanced Google Services → Sheets API v4 → On
4. Save and reload the sheet; use **Tiller Tools → Quick Search** to open the sidebar

## Repo contents

- **Code.js** – Menu entry (Tiller Tools → Quick Search)
- **QuickSearchSidebar.js** – Server-side logic (criteria, filter, helper columns)
- **QuickSearch.html** – Sidebar UI
- **appsscript.json** – Project config (Sheets API, scopes)

## Usage

Open **Tiller Tools → Quick Search**, set filters (date range, amount, description regex, account, category), then click **Search**. Use **Reset All** or the × buttons to clear sections. Description supports regex and ` but not ` (e.g. `gas but not chevron`).
