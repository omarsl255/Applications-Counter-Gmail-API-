# Job Application Tracker — Gmail Automation and Analysis

This Python script connects to Gmail via the Google API, finds job application confirmation emails, computes key metrics (unique application count, monthly trend, day-of-week and hourly distributions), exports results to CSV for Power BI, and produces PNG visualizations.

## Features
- Non-redundant total count of unique application emails
- Keyword breakdown (e.g., "thank you for applying", "recruiting team")
- Date analysis:
    - Month (YYYY-MM)
    - Day of week (Monday, Tuesday, ...)
    - Hour of day (0–23)
    - Cumulative totals
- Exports a single CSV: `job_application_data_[timestamp].csv` (Analysis_Type, Time_Period, Count)
- Saves PNG charts to the project directory

## Prerequisites
- Python 3.6+
- Required libraries:
    - google-api-python-client
    - google-auth-oauthlib
    - matplotlib
    - numpy

## Installation
Clone repository and install dependencies:
```bash
git clone https://github.com/omarsl255/Applications-Counter-Gmail-API-.git
cd job-application-tracker
pip install google-api-python-client google-auth-oauthlib matplotlib numpy
```

## Google Gmail API Setup (required)
1. Open Google Cloud Console and create a new project.
2. Enable the Gmail API (APIs & Services > Library > Gmail API).
3. Configure OAuth consent screen:
     - User Type: External
     - Fill required fields
     - Add scope: `https://www.googleapis.com/auth/gmail.readonly`
     - Add test user(s): the Gmail address(es) you will use
4. Create OAuth credentials:
     - APIs & Services > Credentials > Create Credentials > OAuth client ID
     - Application type: Desktop app
     - Download the JSON and save as `credentials.json` in the project root (same folder as `job_application_counter.py`).

## How to run
```bash
python job_application_counter.py
```
On first run the script will:
- Print a URL — open it in your browser
- Sign in and grant read-only Gmail access
- Copy the authorization code from Google and paste it into the console prompt

A `token.json` file will be saved for future runs.

## Outputs
- CSV: `job_application_data_[timestamp].csv` — columns: Analysis_Type, Time_Period, Count
    - Example rows:
        - Analysis_Type: Monthly, Time_Period: 2024-06, Count: 12
        - Analysis_Type: DayOfWeek, Time_Period: Monday, Count: 25
        - Analysis_Type: Hourly, Time_Period: 14, Count: 8
- PNG charts: saved to project directory (monthly_trend.png, weekday_distribution.png, hourly_distribution.png, etc.)

## Using in Power BI
1. Open Power BI Desktop → Get Data → Text/CSV
2. Select the generated `job_application_data_[timestamp].csv`
3. Use Analysis_Type to filter and build visuals (monthly, weekday, hourly, totals)

## Notes
- Keep `credentials.json` private.
- The script uses read-only Gmail scope; it does not modify mail.
- If you change scopes or delete `token.json`, you must re-authenticate.

License and contribution guidance can be added to this README as needed.