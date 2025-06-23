# NBH Pre-Meeting Briefing Agent

This repository contains the Python script for an automated agent that prepares and emails comprehensive pre-meeting briefs for the NoBrokerHood (NBH) monetization team. The agent integrates with Google Workspace (Calendar, Drive, Sheets, Gmail) and the Google Gemini API to deliver data-driven, context-aware briefings.

## Overview

The agent's primary goal is to empower the NBH sales team by providing them with all relevant internal and external context before they meet with a potential client. It automates the tedious process of gathering information, allowing the team to focus on strategy and execution.

The workflow is as follows:
1.  **Monitor Calendar:** The agent scans the specified Google Calendar for upcoming meetings.
2.  **Extract Brand Info:** For each relevant meeting, it uses the Gemini API to parse the meeting title and identify the client's **Brand Name** and **Industry**.
3.  **Retrieve Internal Data:** It queries a dedicated Google Drive folder to retrieve and parse various internal documents:
    *   The NBH Pitch Deck
    *   Case Study documents (PDFs)
    *   Historical campaign data (Google Sheets)
    *   Overall platform metrics (Google Sheets)
    *   A database of all previous NBH meetings (Google Sheets)
4.  **Analyze & Synthesize:** The agent analyzes the retrieved data, looking for past interactions with the brand, relevant case studies from the same industry, and attendee overlap to determine if a meeting is a direct follow-up.
5.  **Send Leadership Alerts:** If a meeting is scheduled with an existing brand but involves a new, separate team (i.e., not a direct follow-up), a special alert email is sent to leadership to ensure internal alignment.
6.  **Generate Brief:** It combines all the synthesized information into a detailed context and uses a sophisticated prompt to ask the Gemini API to generate a structured, professional pre-meeting brief. The prompt dynamically changes the brief's structure based on whether the meeting is a first-time interaction or a direct follow-up.
7.  **Email & Tag:** The final brief is emailed to all internal NBH attendees. The agent then tags the calendar event as processed and sets a 1-hour email reminder for the attendees.

## Features

- **Automated Calendar Monitoring:** Runs on a schedule to find new meetings automatically.
- **AI-Powered Brand Extraction:** Uses Gemini to intelligently identify brand names and industries from messy meeting titles.
- **Multi-Format Document Parsing:** Reads and extracts text from PDFs, Google Slides, Google Sheets, Excel files, and Google Docs.
- **Sophisticated Follow-up Detection:** Goes beyond simple name matching by normalizing attendee names to accurately detect when a meeting is a direct continuation of a previous discussion.
- **Context-Aware Brief Generation:** The generated brief's format changes to be more effective for first-time meetings vs. follow-up discussions.
- **Proactive Leadership Alerts:** Identifies potentially siloed engagements with existing clients and notifies leadership for better coordination.
- **Robust Authentication:** Securely handles Google Cloud authentication for both local development (interactive flow) and CI/CD environments (via environment variables).

## Setup and Configuration

### 1. Prerequisites

- Python 3.8+
- A Google Cloud Platform (GCP) project.

### 2. GCP Configuration

1.  **Enable APIs:** In your GCP project, enable the following APIs:
    *   Google Calendar API
    *   Gmail API
    *   Google Drive API
    *   Google Sheets API
    *   Vertex AI API (for Gemini)

2.  **Create OAuth Credentials:**
    *   Go to "APIs & Services" -> "Credentials".
    *   Click "Create Credentials" -> "OAuth client ID".
    *   Select "Desktop app" as the application type.
    *   Download the JSON file and rename it to `credentials.json`. Place this file in the root of the project directory.

### 3. Python Environment

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd <your-repository-name>
    ```

2.  **Create a virtual environment:**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    *(Note: You will need to create a `requirements.txt` file. See the section below.)*

### 4. Create `requirements.txt`

Create a file named `requirements.txt` in the root of your project with the following content:
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
google-generativeai
python-pptx
openpyxl
PyMuPDF
markdown

### 5. Environment Variables

The script relies on environment variables for configuration. You can set these in your shell or use a `.env` file with a library like `python-dotenv`.

- **`GEMINI_API_KEY`**: Your API key for the Gemini/Vertex AI API.
- **`NBH_GDRIVE_FOLDER_ID`**: The ID of the Google Drive folder containing all the NBH data files.
- **`GOOGLE_TOKEN_JSON_...`** (For CI/CD): If running in a CI environment like GitHub Actions, you need to store the generated `token_..._...json` file content as a secret. The script looks for secrets named `GOOGLE_TOKEN_JSON_CALENDAR`, `GOOGLE_TOKEN_JSON_GMAIL`, etc.

## Running the Script

### First-Time Run (Local)

The first time you run the script locally, it will open a web browser to ask for your consent to access your Google account.

1.  Ensure `credentials.json` is in the project root.
2.  Set the required environment variables.
3.  Run the script:
    ```bash
    python meeting-prep_running_version.py
    ```
4.  Follow the on-screen prompts in your browser to grant permissions.
5.  After successful authentication, the script will generate `token_brandvmeet_...json` files. These files store your refresh tokens so you won't have to authenticate through the browser on subsequent runs. **Do not commit these token files to version control.** Add them to your `.gitignore` file.

### Subsequent Runs

As long as the `token_...json` files are present and valid, the script will run without requiring browser interaction.

```bash
python meeting-prep_running_version.py
