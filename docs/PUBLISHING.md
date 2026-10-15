# Better Budget - publishing checklist

## 1. Build a clean template (do NOT copy your personal workbook)
1. Create a brand-new Google Sheet named "Better Budget (Template)".
2. Extensions > Apps Script. Delete the default Code.gs.
3. Add one file per script, pasting the latest versions:
   Custom Menu, Settings and Setup, Styling, First Run, Record sidebar,
   SidebarHTML (type: HTML), Monthly Reports, New Month, Dashboard,
   Weekly Report, About.
   (Do not add Clear table or any old copies.)
4. Project Settings > tick "Show appsscript.json manifest file" and paste in
   appsscript.json. Set the time zone there if you want the weekly email at 8am
   in another zone (the email time follows the script time zone).
5. Reload the sheet, then Better Budget > Set Up / Upgrade Tool and walk
   through the first run. Confirm: Start Here, Chart of Accounts (starter), one
   ledger tab. Delete that test ledger's entries (or the tab) and the Financials
   tabs so the template ships empty.

## 2. Test as a stranger
- Open the template in a different Google account (or private window), choose
  File > Make a copy, authorise, and run through the Start Here steps.
- Expect Google's "unverified app" screen on first authorisation. That is normal
  for a copied script: Advanced > Go to Better Budget (unsafe). Say so in your
  instructions.

## 3. Share
- Share > "Anyone with the link" as Viewer, then give out the link with
  /copy at the end instead of /edit, so people are asked to make their own copy.
- Put the version (About.gs) in the file description. Bump it on every release.

## 4. Before every release
- Update BETTER_BUDGET_VERSION in About.gs.
- Re-test: first run, entry sidebar, reports, new month, dashboard, weekly email.
