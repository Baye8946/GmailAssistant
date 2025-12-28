# PowerShell Gmail Assistant

A personal PowerShell tool to automate Gmail and:

- Fetch new emails
- Filter and categorize depending on your preferences
- Generate daily summaries for console and in CSV format

Perfect for freelancers, founders, business owners, and anyone who wants to bring some sanity to their inbox.

## Features
- Gmail API integration (secure OAuth)
- Rule-based categorization
- Daily report with counts
- CSV export

## Setup
1. Enable Gmail API and get `client_secret.json` (Google Cloud Console)
2. Run `Get-GmailToken.ps1` for first-time authorization
3. Run `GmailDailySummary.ps1` for daily summary

**Warning**: Never commit `client_secret.json` or `gmail_token.json`

## Usage
```powershell
.\GmailDailySummary.ps1
Manages your gmail messages, filters through for important messages, responses from cold emails, meeting invites, and categorizes these messages to help your inbox stay clutter free and make sure you don't miss important messages
