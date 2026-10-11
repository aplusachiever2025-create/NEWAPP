# SG Tutor Match — NEWAPP

This repository is being repurposed as the **independent SG Tutor Match system**. The original `aplusachiever2025-create/app` repository is not modified.

## Files
- `index.html`: parent and tutor intake interface. It uses `google.script.run`, so it is intended to be served by Apps Script HTML Service. GitHub Pages alone cannot persist submissions.
- `admin.html`: admin interface source, intended to be added to Apps Script as an HTML file named `admin`.
- `Code.gs`: Google Apps Script backend for Sheets setup, submissions, admin authentication, records, and subject matching.

## Required setup (not yet deployed)
1. Open [Google Apps Script](https://script.google.com/) and create a project.
2. Add `Code.gs` and HTML files named `index` and `admin` using the matching files in this repository.
3. In **Project Settings → Script properties**, set `ADMIN_PASSWORD` to a strong, unique password. Do not store this password in GitHub.
4. Run `setupSheets()` from the Apps Script editor and approve the requested Google permissions. Keep the created spreadsheet private.
5. Deploy as a Web App, execute as the owner, and review the access setting carefully. Share the Web App URL only with intended users. Do not expose the underlying spreadsheet.
6. Test with dummy data first: parent submission, tutor submission, admin login, match generation, removal, and private contact visibility. Only after tests pass should you use it for real contact information.

## Important limitations
- A GitHub Pages site cannot use `google.script.run`; Apps Script must serve the application for the included client calls to work. The Pages version is source/preview only until the app is deployed through Apps Script.
- No existing Supabase data is imported or deleted. No Supabase project is changed by this repository.
- The matching function generates suggestions only; it does not send contacts to parents or tutors and does not confirm a match.
- This initial version does not include resume upload. Add private Drive upload only after access controls and file-handling rules are tested.
