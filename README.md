# Shared Inbox Email Router (Google Apps Script)

This repository contains a privacy-safe, reusable template for a Gmail routing automation built with Google Apps Script.

## What it does

- Scans recent inbox messages on a schedule.
- Classifies messages using subject keyword rules.
- Routes matched emails to destination inboxes.
- Applies configurable blockers on both subject and body per destination.
- Prevents duplicate forwarding with a processed-message store.
- Logs each run, forward action, and error in Google Sheets.
- Sends a daily error digest by email.

## Project structure

- `shared-inbox-router-template/src/Code.gs`: Router logic.
- `shared-inbox-router-template/appsscript.json`: Apps Script manifest (V8 + Gmail advanced service).

## Setup

1. Create a new standalone Apps Script project.
2. Copy `Code.gs` into the project.
3. Copy `appsscript.json` into the project manifest.
4. Enable the Gmail advanced service (`Gmail API v1`).
5. Create log sheets and update `CONFIG.LOG_SPREADSHEET_ID`.
6. Fill in your mailbox list, destinations, keywords, and blockers in `CONFIG`.
7. Run `createTriggers()` once to install scheduled triggers.

## Configuration highlights

- `ALLOWED_MAILBOXES`: permitted sender mailbox identities.
- `ROUTE_A_KEYWORDS` / `ROUTE_B_KEYWORDS`: routing rules by subject.
- `ROUTE_A_SUBJECT_BLOCKERS`, `ROUTE_A_BODY_BLOCKERS`, `ROUTE_B_SUBJECT_BLOCKERS`, `ROUTE_B_BODY_BLOCKERS`: block lists by destination.
- `MAX_FORWARDS_PER_RUN`: forwarding safety cap per execution.

## Notes

- This template intentionally uses placeholder emails, IDs, and generic keywords.
- Keep secrets and real identifiers out of source control.
- Review and test routing rules in a safe mailbox before production rollout.
