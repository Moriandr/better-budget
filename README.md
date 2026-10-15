# Better Budget

A personal budgeting tool that runs inside a Google Sheet, built with Google Apps Script.
Record what you spend, earn and move in plain English; it keeps proper double-entry books
behind the scenes and builds your monthly statements, an executive dashboard and an optional
weekly budget email. No accounting knowledge required.

## Features

- **Plain-English entry form**: spent money, received money, moved money, got money back, opening balance (an Advanced mode is there for bookkeepers)
- **Monthly reports**: T-accounts, trial balance, profit and loss, balance sheet and cash flow statement
- **Executive dashboard**: net worth, revenue vs expenses, expense breakdown by group
- **Weekly budget email**: flags budgets that are over, or on pace to go over
- **Start New Month**: rolls balances forward into the next ledger tab
- **Generic Chart of Accounts** with dropdowns, plus a Settings table (currency, main bank account, report email)
- Your data stays in your own Google Sheet; the only thing ever sent out is the optional weekly email to yourself

## Install (about 15 minutes)

1. Download the files in [`dist/`](dist): `Better Budget.gs`, `SidebarHTML.html` and `README.txt`.
2. Follow the step-by-step guide in [`dist/README.txt`](dist/README.txt) (also at [`docs/INSTALL.txt`](docs/INSTALL.txt)).

In short: create a Google Sheet, open **Extensions > Apps Script**, paste `Better Budget.gs` into `Code.gs`,
add an HTML file named exactly `SidebarHTML`, reload the sheet and use the **Better Budget** menu.

Google will show an "unverified app" screen the first time, as it does for any script you add yourself.
See the install guide for exactly what to click.

## Repository layout

```
src/modules/     the tool, one file per area (edit these)
src/             SidebarHTML.html (entry form) and appsscript.json (manifest)
tools/build.py   joins the modules into dist/Better Budget.gs
dist/            ready-to-install files (generated; commit after building)
docs/            install guide and publishing checklist
```

## Development

```
python3 tools/build.py
```

The Apps Script runtime loads all files into one shared scope, so top-level code in a module
must not depend on constants defined in another module (use functions instead).
Bump `BETTER_BUDGET_VERSION` in `src/modules/About.gs` for each release and add a line to `CHANGELOG.md`.

## Disclaimer

Better Budget is a personal budgeting aid. It is not professional financial, tax or legal advice,
and it comes with no warranty. Check important figures yourself and keep your own backups.
