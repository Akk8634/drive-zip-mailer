# Drive → Zip → Email Mailer

A Google Apps Script that **zips the documents in a Google Drive folder and emails them automatically**. Everything is controlled from a Google Sheet, so non-technical people can add or change recipients without touching code.

No server, no hosting cost: it runs inside your Google account.

## What it does
- Reads rows from a Google Sheet: one row = one folder → one recipient
- Zips every file in the folder (optionally including sub-folders)
- Exports Google Docs, Sheets and Slides to PDF automatically
- Emails the zip as an attachment (falls back to a Drive link if it is over ~24 MB)
- Repeats on a schedule: `once`, `daily`, `weekly`, `monthly` or `off`
- Writes **Last sent** and **Status** (or the error message) back to the sheet
- Adds a **Zip Mailer** menu in the sheet: send now, send selected row, install trigger

## Setup (5 minutes)
1. Create a new Google Sheet.
2. Open **Extensions → Apps Script** and paste the contents of [`Code.gs`](Code.gs).
3. (Optional) In project settings enable "Show appsscript.json" and paste [`appsscript.json`](appsscript.json).
4. Reload the sheet, then run **Zip Mailer → Create sample sheet**.
5. Fill in the "Mailer" sheet (see [`sample-sheet.csv`](sample-sheet.csv) for an example).
6. Run **Zip Mailer → Send now** once and approve the permissions.
7. Run **Zip Mailer → Install daily trigger** to automate it.

## Sheet columns
| Column | Meaning |
|---|---|
| Name | Label used in the zip file name and email |
| Drive folder | Folder URL or ID |
| To / CC | Recipient email(s), comma separated |
| Subject | Optional custom subject |
| Repeat | `once`, `daily`, `weekly`, `monthly`, `off` |
| Include sub-folders | `yes` or `no` |
| Last sent / Status | Filled in by the script |

## Notes
- Google limits the number of emails per day (about 100 on free accounts, 1,500 on Workspace).
- Apps Script runs are limited to 6 minutes. For very large folders, split them across rows.
- The script runs as the account that owns the sheet, and that account must have access to the folders.

## Need something similar?
I build custom automations like this for Google Sheets, Drive, Gmail and Telegram. [Get in touch](https://github.com/Akk8634).
