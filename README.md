# Attendance System Setup

## Prerequisites

1.  **Google Cloud Credentials**
    *   Go to [Google Cloud Console](https://console.cloud.google.com/).
    *   Create a new project.
    *   Enable **Google Sheets API** and **Google Drive API**.
    *   Create a **Service Account**.
    *   Download the JSON key file and rename it to `credentials.json`.
    *   Place `credentials.json` in this folder (`attendance_system`).
    *   **Share your Google Sheet** with the `client_email` address found in `credentials.json`.

2.  **Twilio Setup**
    *   Sign up for [Twilio](https://www.twilio.com/).
    *   Set up the **WhatsApp Sandbox**.
    *   Get your **Account SID** and **Auth Token**.
    *   Update the `.env` file with these values.

3.  **Google Sheet Setup**
    *   Create a new Google Sheet named `Attendance_Tracker`.
    *   **Tab 1 (Sheet1)**: This will store the attendance logs.
        *   Columns: `Date`, `Name`, `Phone`, `Time`, `Status`
    *   **Tab 2 (Employees)**: Create a new sheet named `Employees`.
        *   Columns: `Name`, `Phone`
        *   Add your name and phone number (in the format `+1234567890`) to test.

## Running the App

1.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
2.  Run the server:
    ```bash
    python app.py
    ```
3.  Expose the server to the internet (using ngrok or similar) and configure the Twilio Webhook URL to point to `your-url/bot`.
