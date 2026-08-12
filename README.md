# Budget Inbox Sync

A Google Apps Script that turns Gmail into a lightweight financial data pipeline — no bank API, no third-party budgeting app.

## What it does

- Parses transaction notification emails (major banks + payment apps) straight out of Gmail and logs them to a Google Sheet, deduplicated
- Auto-categorizes transactions against a keyword-rule sheet, without ever overwriting a manual correction
- Tracks account balances over time from balance-alert emails
- Sends a daily HTML digest of new transactions + current balances
- Generates a weekly AI financial report (Gemini API) — budget-vs-actual, spending trends, and concrete next-week actions, written in plain language instead of a wall of numbers

## How it works

Everything lives in Google Sheets as the "database" — a Config sheet for settings, a Categories sheet for keyword rules, and auto-populated Transactions/BalanceHistory sheets. Apps Script triggers run the Gmail search, parse, and categorize pipeline on a schedule; a separate trigger sends the digest email.

**Tools:** Google Apps Script, Gmail API, Sheets API, Gemini API

## Setup

1. Run `bootstrapAll()` once from the Apps Script editor — creates the Config, Categories, and BalanceHistory sheets
2. Seed the Categories sheet with your own keyword → category rules
3. Set your Gemini API key in Script Properties (`GEMINI_API_KEY`) if you want the weekly AI report
4. Set up time-driven triggers for `fetchTransactionEmails`, `sendDailyDigest`, and `sendComprehensiveWeeklyReport`
