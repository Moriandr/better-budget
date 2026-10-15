BETTER BUDGET - SET-UP GUIDE
============================

Better Budget is a budgeting tool that runs inside a Google Sheet. You tell it
what you spent, earned or moved, and it builds your monthly reports, a
dashboard, and an optional weekly email about your budgets.
No accounting knowledge needed. Free. Your data stays in your own Google Sheet.

This guide has 3 parts:
   PART 1  Put the tool into a Google Sheet      (about 15 minutes, once)
   PART 2  First-time set-up                     (about 5 minutes, once)
   PART 3  Using it                              (every day / every month)

If something goes wrong, jump to PART 4 at the end.


=============================================================================
PART 1 - PUT THE TOOL INTO A GOOGLE SHEET
=============================================================================

WHAT YOU NEED BEFORE STARTING
  - A Google account (a free Gmail account is fine).
  - The 2 tool files that came with this guide:

      [ ] Better Budget.gs        the whole tool in one script file
      [ ] SidebarHTML.html        the entry form (an HTML file, see Step 4)

  You will copy the text inside each file and paste it into Google's code
  editor. Open each file with Notepad (Windows) or TextEdit (Mac), or any
  text editor.

-----------------------------------------------------------------------------
STEP 1.  Create the sheet
-----------------------------------------------------------------------------
  1. Go to  sheets.google.com  and click "Blank spreadsheet".
  2. Click the title at the top left and rename it "Better Budget".

-----------------------------------------------------------------------------
STEP 2.  Open the code editor
-----------------------------------------------------------------------------
  1. In the menu at the top of the sheet, click  Extensions > Apps Script.
  2. A new browser tab opens. On the left, under "Files", you will see one
     file called  Code.gs.

  IMPORTANT: always open the editor this way, from inside your sheet.
  If you create a project at script.google.com instead, the menu will never
  appear.

-----------------------------------------------------------------------------
STEP 3.  Paste in the script
-----------------------------------------------------------------------------
  1. Open  Better Budget.gs  in your text editor. Press Ctrl+A (select all),
     then Ctrl+C (copy). On a Mac use Cmd instead of Ctrl.
  2. In the Google editor, click  Code.gs  in the left list.
  3. Click in the code area, press Ctrl+A to select the old text, then
     Ctrl+V to paste. The old text is replaced.
  4. Optional: hover over Code.gs, click the three dots (...), choose
     "Rename" and call it   Better Budget.   (Do not type ".gs".)

-----------------------------------------------------------------------------
STEP 4.  Add the HTML file (this one is different)
-----------------------------------------------------------------------------
  1. Click the "+" next to "Files" and choose  "HTML"  (not "Script").
  2. Type exactly:   SidebarHTML     (no ".html"; the editor adds it).
  3. The editor fills it with a few lines of sample text. Click in the code
     area, press Ctrl+A, and paste the full contents of SidebarHTML.html.

  If the name is not exactly "SidebarHTML", the entry form will not open.

-----------------------------------------------------------------------------
STEP 5.  Save and check
-----------------------------------------------------------------------------
  1. Click the Save icon (the floppy disk) at the top, or press Ctrl+S.
  2. Check the left list shows exactly 2 files:
       Code (or Better Budget), and SidebarHTML
     If there are other files, you may have added something by mistake.

-----------------------------------------------------------------------------
STEP 6.  Make the menu appear
-----------------------------------------------------------------------------
  1. Go back to the browser tab with your spreadsheet.
  2. Reload the page (press F5, or click the browser's reload button).
  3. Wait about 10 seconds. A new menu called  Better Budget  appears at the
     top, next to "Help".

  If it does not appear, go to PART 4.


=============================================================================
PART 2 - FIRST-TIME SET-UP
=============================================================================

STEP 7.  Run the set-up
  1. Click  Better Budget > Set Up / Upgrade Tool.
  2. Google will ask for permission. Do this:
       - Click "Continue" and choose your Google account.
       - A screen says "Google hasn't verified this app". This is normal for
         any script you add yourself. Click "Advanced" (small, bottom left),
         then "Go to Better Budget (unsafe)", then "Allow".
  3. Nothing may happen after you allow it. Click  Better Budget >
     Set Up / Upgrade Tool  again.
  4. A box asks which month to start with. Type it like  October 2026,
     or leave it blank to use this month, then click OK.
  5. Wait. The tool creates three tabs along the bottom:
       - Start Here          short instructions
       - Chart of Accounts   your list of accounts, plus your settings
       - (a month tab)       your ledger for the first month, e.g. "October 2026"
  6. Click OK on the final message.

STEP 8.  Check your settings
  1. Click the "Chart of Accounts" tab.
  2. To the right of the list (columns H to J) is the Settings table:
       Currency symbol           type yours, for example  $  or  J$  or  EUR
       Report email              where the weekly email goes (blank = your own)
     Leave the other settings alone.
  3. In the account list (columns A to F), change the account names to match
     your real life, for example your real bank or credit card. Use the
     drop-down arrows in the Category column; do not type there.
  4. Do not rename the "Chart of Accounts" tab.

STEP 9.  Enter starting balances
  1. Click your month tab (for example "October 2026").
  2. Click  Better Budget > Make Transaction Entry.  A form opens on the right.
  3. Choose "Opening balance". For each account, enter what it held (or what
     you owed) on the day you want to start. One entry per account.
  4. (Optional) In the Chart of Accounts tab, type a number in the
     "Monthly Budget" column next to any spending category you want to
     track, for example 400 next to Groceries.

You are ready to use the tool.


=============================================================================
PART 3 - USING IT
=============================================================================

EVERY DAY: record what happens
  1. Make sure the correct month tab is open.
  2. Click  Better Budget > Make Transaction Entry.
  3. Choose what happened:
       Spent money       you bought something
       Received money    you were paid, or money came in
       Moved money       you moved money between your own accounts, or paid a
                         card or loan
       Got money back    a refund or reimbursement (it reduces that
                         spending category)
       Opening balance   a starting balance (see Step 9)
       Advanced          only if you know double-entry bookkeeping
  4. Fill in the date, amount and a short note, then save.
  Always use the form instead of typing into the ledger by hand.

EVERY MONTH: reports and rolling forward
  1. Open the month tab. Click  Better Budget > Generate Monthly Reports.
     A new tab "Financials - Month Year" is created with your statements.
  2. Click  Better Budget > Generate Executive Dashboard  for the big picture.
  3. When the month is over: generate that month's reports FIRST, then click
     Better Budget > Start New Month  while you are on the last month tab.
     This creates next month's tab and carries your balances forward.
  Reports are snapshots. If you change an old entry, run the reports again.

OPTIONAL: weekly email
  Better Budget > Weekly Email > Send now     see what it looks like
  Better Budget > Weekly Email > Turn on      get it every Monday morning
  Better Budget > Weekly Email > Turn off     stop it
  Better Budget > Weekly Email > Status       check whether it is on
  The email shows your budgets and warns about any that are over or on pace
  to go over. It goes to the "Report email" in Settings, or to your own Google
  account if that is blank. It is the only thing the tool ever sends out.
  To set which time zone "Monday morning" uses: in the code editor click the
  gear icon (Project Settings) and set the Time zone.

HABITS THAT KEEP IT WORKING
  - Make a backup now and then: File > Make a copy.
  - Do not delete or rename the tabs the tool creates.
  - Do not delete an account that already has entries.

UPDATING LATER
  Better Budget > About shows your version. To update, replace the text in
  the script (and SidebarHTML) with the newer version, save, reload the sheet,
  then click Set Up / Upgrade Tool once. Your data is kept.


=============================================================================
PART 4 - IF SOMETHING GOES WRONG
=============================================================================

The "Better Budget" menu does not appear
  1. Reload the spreadsheet and wait 10 seconds.
  2. Make sure you opened the editor from the sheet itself
     (Extensions > Apps Script), not from script.google.com.
  3. In the code editor, check the script was pasted in full: press Ctrl+End
     inside it and compare the last lines with the original file. Also check
     that SidebarHTML is there.
  4. In the editor, click "Executions" (left side, list icon), then reload the
     sheet. If a row says "Failed", click it and read the message. Send that
     message to whoever gave you this tool.
  5. If you use several Google accounts in one browser, try a private
     (incognito) window signed into one account only.
  6. You need edit access to the sheet. People with view-only access do not
     get the menu.

"X is not defined" or "Script function not found"
  The script or the HTML file was only partly pasted. Select all in the
  editor (Ctrl+A) and paste the whole file again.

The entry form does not open
  The HTML file must be named exactly  SidebarHTML  and pasted in full.

"Cannot call SpreadsheetApp.getUi() from this context"
  You pressed Run inside the code editor. Use the Better Budget menu in the
  spreadsheet instead.

"Incompatible sheet selected"
  You are on the wrong tab. Click a month tab and try again.

Something looks odd in a sheet you already had
  Better Budget > Set Up / Upgrade Tool is safe to run again. It adds what is
  missing and never deletes your data.


=============================================================================
IMPORTANT
=============================================================================
Better Budget is a personal budgeting aid. It is not professional financial,
tax or legal advice, and it comes with no warranty. Check important figures
yourself and keep your own backups.
