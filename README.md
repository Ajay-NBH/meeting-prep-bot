Meeting Prep Bot
Automated Pre-Meeting Brief Generator for NoBrokerHood Sales Team

This Python script automates the preparation and delivery of comprehensive, data-driven pre-meeting briefs for upcoming client meetings. It integrates with Google Workspace (Calendar, Drive, Sheets, Gmail), uses Gemini LLM for natural language tasks, and aggregates internal data to empower your sales and leadership teams.

Features
Calendar Integration: Scans upcoming events from Google Calendar, focusing on meetings where the brand.vmeet agent is invited.
Brand & Industry Extraction: Uses Gemini LLM to intelligently extract brand and industry information from meeting titles.
Internal Data Aggregation: Pulls and summarizes relevant decks, case studies, campaign sheets, and past meeting notes from Google Drive and Sheets.
Follow-Up Detection: Detects direct follow-up meetings (based on attendee overlap) and flags other past brand interactions.
Brief Generation: Crafts a detailed, structured meeting brief using Gemini LLM, adapting format based on meeting type (follow-up vs. new).
Automated Email Delivery: Emails the pre-meeting brief to all NBH internal attendees (excluding service accounts).
Leadership Alerts: Sends special notifications to leadership if complex engagement patterns are detected (e.g., multiple threads, follow-ups).
Event Tagging & Reminders: Marks processed events and sets 1-hour email reminders before meetings.

Setup
1. Google Cloud Platform
APIs Required:
Google Calendar
Google Drive
Google Sheets
Gmail
Credentials:
Download your credentials.json OAuth client file from GCP.
Place it in the project root.

3. Environment Variables
GEMINI_API_KEY – Your Gemini LLM API key.
NBH_GDRIVE_FOLDER_ID – (Optional) ID of your Google Drive folder for NBH data.
(Optional) GOOGLE_TOKEN_JSON_* variables for running in CI environments.

5. Python Requirements
Install dependencies (recommend using a virtual environment):
bash
pip install -r requirements.txt
Requirements include:
google-api-python-client, google-auth-oauthlib, fitz (PyMuPDF), openpyxl, python-pptx, markdown, google-generativeai, etc.

4. First Run
On first execution, the script will prompt for Google OAuth consent and save token files locally.
Subsequent runs will use the saved tokens.
Usage
Run the script directly:

bash
python meeting-prep_running_version.py
What it does:

Authenticates with Google Calendar, Drive, Sheets, and Gmail.
Checks upcoming meetings (default: next 96 hours).
For each event:
Extracts meeting, attendee, and brand details.
Aggregates and summarizes internal context from relevant files.
Determines if the meeting is a direct follow-up.
Crafts a detailed pre-meeting brief (markdown converted to HTML).
Emails the brief to NBH attendees.
Sets reminders and notifies leadership if warranted.
Tags the event and records it as processed.
Configuration
File and Column Names: If your Drive/Sheets file names or column headers differ from defaults, update the relevant variables at the top of the script.
Leadership Emails: Set the leadership_emails list to control who receives alert notifications.
Excluded Accounts: Adjust the EXCLUDED_NBH_PSEUDO_NAMES_FOR_FOLLOWUP and NBH_SERVICE_ACCOUNTS_TO_EXCLUDE as needed.
Customization
The script is modular. You can:
Adapt file naming conventions.
Modify the LLM prompt templates in the code.
Adjust logic for attendee/brand parsing, follow-up detection, or email formatting.
Troubleshooting
OAuth Issues: Delete token files and re-run to re-authenticate.
API Errors: Check scopes, quotas, and credentials.
Gemini API Errors: Ensure your API key is correct and quotas are not exceeded.
Security Notes
Treat your credentials and token files securely. Do not commit them to public repositories.
Use environment variables for sensitive keys and IDs.
License
MIT License (or specify your license here)

Credits
Developed by the NoBrokerHood Monetization Team, with Gemini LLM integration.

