# Gmail Attachment Monitor

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

The Gmail Attachment Monitor is a desktop application designed specifically for students during placement season (or anyone who needs to monitor emails for specific content within attachments). It automatically checks your Gmail account for new emails containing specific keywords (like your name) within attached documents (PDF, DOCX, TXT, CSV, XLSX, ZIP contents listing) and notifies you if a match is found.

This helps you stay updated on shortlisting notifications or other important information hidden inside email attachments without manually checking every single one.

## Features

*   **Automated Background Monitoring:** Runs checks at a user-defined interval.
*   **Attachment Content Scanning:** Searches for keywords *inside* common attachment types:
    *   Plain Text (.txt)
    *   PDF (.pdf) - Extracts text content.
    *   Microsoft Word (.docx)
    *   CSV (.csv)
    *   Excel (.xlsx) - Extracts text from all sheets.
    *   ZIP (.zip) - Lists contents (future enhancement: search inside zipped files).
*   **Keyword Specific:** Focuses searches on a user-provided keyword (e.g., your name).
*   **Date Range Filtering:** Limits searches to recent emails (e.g., last 1 day, 7 days).
*   **Desktop Notifications:** Uses system notifications (via Plyer) for new matches.
*   **User-Friendly Interface (PyQt6):**
    *   Clear display of matched emails.
    *   Details view showing sender, subject, date, matched filenames, and email body preview.
    *   Filtering capabilities for matched emails.
    *   Activity log.
    *   Manual check option.
    *   Theming (Light, Dark, Blue).
*   **Secure Authentication:** Uses Google OAuth 2.0 for secure access to your Gmail account (read-only permission requested). Your credentials are stored locally.

## Requirements

*   Python 3.8+
*   Pip (Python package installer)
*   A Google Account (Gmail)
*   Enabled Gmail API for your Google Account
*   `credentials.json` file obtained from Google Cloud Console

## Setup

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/your-username/gmail-attachment-monitor.git
    cd gmail-attachment-monitor
    ```

2.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    # On Windows
    .\venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

3.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Enable Gmail API & Download Credentials:**
    *   Go to the [Google Cloud Console](https://console.cloud.google.com/).
    *   Create a new project or select an existing one.
    *   Navigate to "APIs & Services" > "Library".
    *   Search for "Gmail API" and enable it.
    *   Navigate to "APIs & Services" > "Credentials".
    *   Click "Create Credentials" > "OAuth client ID".
    *   If prompted, configure the OAuth consent screen (select "External" user type, provide app name, user support email, developer contact info). You don't need to submit for verification for personal use with a limited number of users. Add your own Google account email as a Test User during configuration or later in the "OAuth consent screen" settings.
    *   Choose "Desktop app" as the Application type.
    *   Give it a name (e.g., "Gmail Monitor Desktop App").
    *   Click "Create".
    *   A window will pop up showing your Client ID and Client Secret. Click **"DOWNLOAD JSON"**.
    *   **Rename the downloaded file to `credentials.json`**.
    *   **Place this `credentials.json` file in the root directory** of the cloned project (`gmail-attachment-monitor/`). 

## Usage

1.  **Ensure your virtual environment is active.**
2.  **Run the application:**
    ```bash
    python main.py
    ```
3.  **First Run - Authorization:**
    *   On the first run, a browser window will open asking you to log in to your Google Account and authorize the application to access your Gmail (read-only).
    *   Grant the permissions.
    *   A `token.pickle` file will be created in the project's root directory to store your authorization token for future runs.
4.  **Configure Settings:**
    *   In the "Monitoring Settings" tab:
        *   Enter the **Username/Keyword** you want to find within attachments (e.g., "Varun venkat", "StudentID123"). The search is case-insensitive.
        *   Set how many **past days** of emails to check initially and during manual checks.
        *   Set the **Check Interval** (in seconds) for automatic background monitoring.
5.  **Start Monitoring:**
    *   Click "Start Monitoring". The application will now check your emails periodically in the background.
    *   The status bar will show monitoring activity.
6.  **Check Now:**
    *   Click "Check Now" to perform an immediate manual scan based on the current settings.
7.  **Review Matches:**
    *   Go to the "Matched Emails" tab to see a list of emails where the keyword was found in an attachment.
    *   Click on an email in the list to see its details (sender, subject, date, matched attachments, body preview).
    *   Use the filter bar to search within the matched emails.
8.  **Stop Monitoring:**
    *   Click "Stop Monitoring" to disable automatic background checks.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an Issue.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
