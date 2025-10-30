import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta
from collections import defaultdict
import csv # Import the CSV library

# --- Visualization Imports (Needed for plotting results) ---
try:
    import matplotlib.pyplot as plt
    import numpy as np
except ImportError:
    # Fallback to allow the script to run without plotting if the dependencies are missing.
    plt = None 
    np = None 
    print("Warning: matplotlib or numpy not installed. Visualization will be skipped. Run 'pip install matplotlib numpy' to enable plotting.")
# -----------------------------------------------------------

# --- Configuration ---
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'

# Days to look back for the search query (e.g., 365 for the last year).
DAYS_TO_LOOK_BACK = 365 

# List of individual core search phrases (English and German) to count.
# These will be combined to form the full query.
CORE_SEARCH_PHRASES = [
    # General English Confirmations (Primary candidates for SUBJECT search)
    "application received",
    "thank you for applying",
    "thank you for Your application",
    "application confirmation",
    "application submitted",
    "candidate profile",
    "Your application",
    "we received your",
    "your application to",
    
    # Generic Automated/System & Team Phrases (These should usually search the whole message)
    "application portal",
    "Greenhouse", 
    "recruiting team", 
    "talent team", 
    "talent acquisition team", 
    
    # German Confirmations (Primary candidates for SUBJECT search)
    "Bewerbung erhalten",
    "Vielen Dank für Ihre Bewerbung",
    "Vielen Dank für deine Bewerbung",
    "Ihre Bewerbung",
    "deine Bewerbung",
    "Eingangsbestätigung",
]

# --- Query Refinement ---
# Separate phrases into those best searched in the subject vs. those searched everywhere.
# Terms with 0 count have been removed from the original lists.
SUBJECT_ONLY_PHRASES = [
    # Confirmation terms must be in the subject for high confidence
    "application received", "thank you for applying", "thank you for Your application",
    "application confirmation", "application submitted", "candidate profile",
    "Your application", "we received your",
    "your application to", "Bewerbung erhalten", "Vielen Dank für Ihre Bewerbung",
    "Vielen Dank für deine Bewerbung", "Ihre Bewerbung", "deine Bewerbung",
    "Eingangsbestätigung"
]

BODY_OR_SENDER_PHRASES = [
    # System/Team names can appear in the body, sender name, or sender email.
    "application portal", "Greenhouse", 
    "recruiting team", "talent team", "talent acquisition team"
]

# 1. Build the SUBJECT-only part of the full query
SUBJECT_QUERIES = [f'subject:"{p}"' for p in SUBJECT_ONLY_PHRASES]

# 2. Build the BODY/SENDER part of the full query (searches anywhere)
BODY_QUERIES = [f'"{p}"' for p in BODY_OR_SENDER_PHRASES]

# Combine both parts for the final non-redundant search
ALL_QUERIES_COMBINED = SUBJECT_QUERIES + BODY_QUERIES
FULL_JOB_APPLICATION_QUERY = f'({" OR ".join(ALL_QUERIES_COMBINED)}) -is:draft'

# --- End Query Refinement ---


def authenticate_gmail():
    """Shows user authentication flow using console and returns a Gmail API service object."""
    creds = None
    # The token.json file stores the user's access and refresh tokens.
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    # If there are no (valid) credentials available, or they are expired, handle login/refresh.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing existing token...")
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"Error: {CREDENTIALS_FILE} not found.")
                print("Please download your credentials file from the Google API Console.")
                return None
            
            # Use 'urn:ietf:wg:oauth:2.0:oob' (Out-of-Band) for reliable manual console flow.
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            
            # Set the redirect URI for the Out-of-Band flow (copy/paste from browser).
            flow.redirect_uri = 'urn:ietf:wg:oauth:2.0:oob'

            auth_url, _ = flow.authorization_url(prompt='consent')
            
            print("Starting manual console authentication flow...")
            print("\nPlease visit this URL in your browser to authorize:")
            print(auth_url)
            
            code = input("\nEnter the authorization code displayed in your browser: ").strip()

            try:
                # Exchange the authorization code for credentials
                flow.fetch_token(code=code)
                creds = flow.credentials
            except Exception as e:
                print(f"Error exchanging code for token. Please ensure the code is correct: {e}")
                return None
        
        # Save the credentials for the next run
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
            print(f"Token saved to {TOKEN_FILE}.")
            
    try:
        # Build the Gmail service
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        print(f"An HTTP error occurred: {error}")
        return None

def get_messages_count(service, user_id='me', search_query=''):
    """
    Counts the EXACT number of emails matching a specific search query 
    by iterating through all pages.
    """
    try:
        messages = []
        page_token = None
        
        while True:
            # Call the list method with the query. 
            response = service.users().messages().list(
                userId=user_id, 
                q=search_query, 
                pageToken=page_token
            ).execute()

            # Extend the list with the messages from the current page.
            messages.extend(response.get('messages', [])) 

            # Get the token for the next page. If no token, we are done.
            page_token = response.get('nextPageToken')
            if not page_token:
                break
        
        return len(messages)

    except HttpError as error:
        print(f"An error occurred during the API call: {error}")
        return 0

def get_message_dates(service, full_query):
    """
    Fetches all messages matching the full query and extracts their internal dates.
    Returns a list of datetime objects.
    """
    print("\nStarting analysis of message dates (this may take a moment)...")
    
    messages = []
    page_token = None
    
    # 1. Fetch all message IDs
    while True:
        try:
            response = service.users().messages().list(
                userId='me', 
                q=full_query, 
                pageToken=page_token
            ).execute()
            messages.extend(response.get('messages', [])) 
            page_token = response.get('nextPageToken')
            if not page_token:
                break
        except HttpError as error:
            print(f"Error fetching message IDs for date analysis: {error}")
            return []

    if not messages:
        return []

    date_objects = []
    
    # 2. Process messages to get internalDate
    for i, msg in enumerate(messages):
        try:
            msg_detail = service.users().messages().get(
                userId='me', 
                id=msg['id'], 
                format='metadata'
            ).execute()
            
            internal_date_ms = msg_detail.get('internalDate')
            
            if internal_date_ms:
                # Convert milliseconds since epoch to datetime object
                dt_object = datetime.fromtimestamp(int(internal_date_ms) / 1000) 
                date_objects.append(dt_object)
                
        except HttpError as e:
            # Silently fail on individual message fetch error
            continue 

        # Print progress update
        if (i + 1) % 50 == 0 or (i + 1) == len(messages):
            print(f"Progress: {i + 1}/{len(messages)} messages analyzed...", end='\r', flush=True)

    print("\nDate analysis complete.")
    return date_objects

# --- Analysis Functions ---

def get_monthly_counts(date_objects):
    """Returns a dictionary of monthly counts {YYYY-MM: count}."""
    monthly_counts = defaultdict(int)
    for dt in date_objects:
        monthly_counts[dt.strftime("%Y-%m")] += 1
    return dict(monthly_counts)

def get_day_of_week_counts(date_objects):
    """Returns a dictionary of day of week counts {DayName: count}."""
    day_names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    day_counts = defaultdict(int)
    for dt in date_objects:
        # weekday() returns 0 for Monday, 6 for Sunday
        day_name = day_names[dt.weekday()]
        day_counts[day_name] += 1
    
    # Ensure all days are present in the correct order for visualization
    return {day: day_counts.get(day, 0) for day in day_names}

def get_hourly_counts(date_objects):
    """Returns a dictionary of hourly counts {Hour: count}."""
    hourly_counts = defaultdict(int)
    for dt in date_objects:
        # Hour is 0-23
        hourly_counts[dt.hour] += 1
    
    # Ensure all 24 hours are present in order
    return {hour: hourly_counts.get(hour, 0) for hour in range(24)}

# --- End Analysis Functions ---

def save_to_csv(monthly_counts, day_of_week_counts, hourly_counts, total_count, days_back):
    """Saves all date analysis results to a single CSV file for Power BI."""
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"job_application_data_{timestamp}.csv"
    
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # 1. Write Metadata Summary
        writer.writerow(["Metric", "Value", "Notes"])
        writer.writerow(["Total Applications", total_count, f"Unique application emails found over the last {days_back} days."])
        writer.writerow(["Analysis Start Date", (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d"), ""])
        writer.writerow([])
        
        # 2. Write Monthly Data
        writer.writerow(["Analysis_Type", "Time_Period", "Count"])
        for month, count in sorted(monthly_counts.items()):
            # Time_Period format YYYY-MM
            writer.writerow(["Monthly", month, count])

        # 3. Write Day-of-Week Data
        writer.writerow([])
        writer.writerow(["Analysis_Type", "Time_Period", "Count"])
        for day, count in day_of_week_counts.items():
            # Time_Period format DayName
            writer.writerow(["DayOfWeek", day, count])

        # 4. Write Hourly Data
        writer.writerow([])
        writer.writerow(["Analysis_Type", "Time_Period", "Count"])
        for hour, count in sorted(hourly_counts.items()):
            # Time_Period format 0-23
            writer.writerow(["Hourly", hour, count])

    print(f"\n[Data Exported] Application data successfully saved for Power BI as: {filename}")
    return filename

def create_date_query(base_query, days_back):
    """Generates the full query restricted by a look-back period (e.g., 365 days)."""
    
    start_date = datetime.now() - timedelta(days=days_back)
    
    # Format the start date for the 'after' operator (YYYY/MM/DD)
    formatted_date = start_date.strftime("%Y/%m/%d")
    
    # Combine the base query with the date filter
    full_query = f"({base_query}) after:{formatted_date}"
    return full_query

def visualize_results(phrase_counts, days_back):
    """Generates and saves a horizontal bar chart of the individual phrase counts."""
    
    if plt is None:
        print("\nVisualization requires 'matplotlib' and 'numpy'. Please install them.")
        return

    # Filter out phrases with zero counts for cleaner visualization
    data = {k: v for k, v in phrase_counts.items() if v > 0}
    
    if not data:
        print("\nNo applications found to visualize based on individual phrases.")
        return

    phrases = list(data.keys())
    counts = list(data.values())
    y_pos = np.arange(len(phrases))

    # Set up the plot aesthetics
    plt.style.use('seaborn-v0_8-darkgrid')
    fig, ax = plt.subplots(figsize=(10, 6))

    # Create the horizontal bars
    bars = ax.barh(y_pos, counts, color=plt.cm.cividis(np.linspace(0, 1, len(phrases))))

    # Add labels and title
    ax.set_yticks(y_pos)
    ax.set_yticklabels(phrases, fontsize=10)
    ax.invert_yaxis()  # Labels read top-to-bottom
    ax.set_xlabel('Number of Emails Matched', fontsize=12)
    ax.set_title(f'Job Application Email Matches by Keyword (Last {days_back} Days)', fontsize=14, pad=20)
    ax.tick_params(axis='x', labelsize=10)
    
    # Add the count labels inside the bars
    for bar in bars:
        width = bar.get_width()
        ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, 
                f'{int(width)}',
                va='center', ha='left', fontsize=10, weight='bold')

    # Remove the top and right spines
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.set_xlim(right=max(counts) * 1.1) # Extend x-limit for labels

    plt.tight_layout()
    
    # --- SAVE THE PLOT TO FILE ---
    # Generate a unique filename based on the current timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"job_application_count_breakdown_{timestamp}.png"
    plt.savefig(filename)
    print(f"\n[Visualization Saved] The bar chart has been saved as: {filename}")
    plt.close(fig) # Close the figure to free up memory

def visualize_monthly_results(monthly_counts, days_back):
    """Generates and saves a line plot of the monthly application trend."""

    if plt is None:
        return

    if not monthly_counts:
        return

    # Sort the dictionary keys (YYYY-MM) chronologically
    sorted_keys = sorted(monthly_counts.keys())
    
    # Prepare data for plotting
    dates = [datetime.strptime(k, "%Y-%m") for k in sorted_keys]
    counts = [monthly_counts[k] for k in sorted_keys]

    # Use Month/Year labels for x-axis
    labels = [date.strftime('%b %Y') for date in dates]

    # Set up the plot aesthetics
    plt.style.use('seaborn-v0_8-darkgrid')
    fig, ax = plt.subplots(figsize=(12, 6))

    # Create the line plot
    ax.plot(labels, counts, marker='o', linestyle='-', color='#0077B6', linewidth=2)
    
    # Add data labels for each point
    for i, count in enumerate(counts):
        ax.annotate(str(count), (labels[i], counts[i] + 0.5), ha='center', fontsize=9, color='#333333')

    # Add labels and title
    ax.set_ylabel('Number of Applications', fontsize=12)
    ax.set_title(f'Monthly Job Application Trend (Last {days_back} Days)', fontsize=16, pad=20)
    
    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45, ha='right')
    ax.tick_params(axis='x', labelsize=10)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    ax.set_ylim(bottom=0) # Start y-axis at 0

    plt.tight_layout()
    
    # --- SAVE THE PLOT TO FILE ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"job_application_monthly_trend_{timestamp}.png"
    plt.savefig(filename)
    print(f"[Visualization Saved] The monthly trend chart has been saved as: {filename}")
    plt.close(fig)

def visualize_day_of_week_results(day_counts, days_back):
    """Generates and saves a bar chart of the day of week application breakdown."""
    
    if plt is None:
        return

    days = list(day_counts.keys())
    counts = list(day_counts.values())

    if not any(counts):
        print("No day-of-week data found to visualize.")
        return

    plt.style.use('seaborn-v0_8-darkgrid')
    fig, ax = plt.subplots(figsize=(9, 5))

    ax.bar(days, counts, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2'])
    
    for i, count in enumerate(counts):
        if count > 0:
            ax.text(i, count + max(counts) * 0.02, str(count), ha='center', va='bottom', fontsize=10)

    ax.set_title(f'Applications by Day of Week (Last {days_back} Days)', fontsize=14, pad=15)
    ax.set_ylabel('Number of Applications', fontsize=12)
    ax.set_ylim(top=max(counts) * 1.1) # Extend y-limit for labels
    
    plt.tight_layout()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"job_application_day_of_week_{timestamp}.png"
    plt.savefig(filename)
    print(f"[Visualization Saved] The day-of-week chart has been saved as: {filename}")
    plt.close(fig)

def visualize_hourly_results(hourly_counts, days_back):
    """Generates and saves a bar chart of the hourly application breakdown."""
    
    if plt is None:
        return

    hours = list(hourly_counts.keys())
    counts = list(hourly_counts.values())

    if not any(counts):
        print("No hourly data found to visualize.")
        return

    plt.style.use('seaborn-v0_8-darkgrid')
    fig, ax = plt.subplots(figsize=(12, 6))

    ax.bar(hours, counts, color='#39A78E')
    
    ax.set_xticks(hours)
    ax.set_xticklabels([f'{h:02d}:00' for h in hours], rotation=45, ha='right')

    ax.set_title(f'Applications by Time of Day (UTC, Last {days_back} Days)', fontsize=14, pad=15)
    ax.set_xlabel('Hour of Day', fontsize=12)
    ax.set_ylabel('Number of Applications', fontsize=12)
    ax.set_ylim(bottom=0)
    
    plt.tight_layout()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"job_application_hourly_trend_{timestamp}.png"
    plt.savefig(filename)
    print(f"[Visualization Saved] The hourly trend chart has been saved as: {filename}")
    plt.close(fig)

def visualize_cumulative_results(date_objects, days_back):
    """Generates and saves a line plot of the cumulative application total."""
    
    if plt is None:
        return

    if not date_objects:
        print("No application data found for cumulative visualization.")
        return

    # Sort dates chronologically
    sorted_dates = sorted(date_objects)
    
    # Create cumulative counts
    cumulative_counts = np.cumsum(np.ones(len(sorted_dates)))

    plt.style.use('seaborn-v0_8-darkgrid')
    fig, ax = plt.subplots(figsize=(12, 6))

    # Plot dates vs cumulative count
    ax.plot(sorted_dates, cumulative_counts, marker='.', linestyle='-', color='#765D98', linewidth=2)
    
    ax.set_title(f'Cumulative Job Applications Over Time (Last {days_back} Days)', fontsize=16, pad=20)
    ax.set_xlabel('Date', fontsize=12)
    ax.set_ylabel('Cumulative Total Applications', fontsize=12)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    
    # Format x-axis dates
    fig.autofmt_xdate(rotation=45) 
    
    plt.tight_layout()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"job_application_cumulative_total_{timestamp}.png"
    plt.savefig(filename)
    print(f"[Visualization Saved] The cumulative chart has been saved as: {filename}")
    plt.close(fig)

def main():
    """Authenticates, constructs the query, calculates counts, and prints/visualizes results."""
    
    # 1. Authenticate and get the service object
    service = authenticate_gmail()
    if not service:
        print("\nCould not initialize Gmail service. Check 'credentials.json' and network.")
        return

    phrase_counts = {}
    
    # 2. Get individual keyword counts
    print("\n--- Individual Term Counts (Searches performed on whole message where appropriate) ---")
    
    for phrase in CORE_SEARCH_PHRASES:
        # Determine the search scope: 'subject:' or general (whole message)
        if phrase in SUBJECT_ONLY_PHRASES:
            search_scope = f'subject:"{phrase}"'
        else:
            search_scope = f'"{phrase}"'
            
        individual_base_query = f'{search_scope} -is:draft'
        individual_full_query = create_date_query(individual_base_query, DAYS_TO_LOOK_BACK)
        
        # NOTE: Using a single query to show what it is searching for
        print(f"Searching for: {search_scope} ... ", end="", flush=True)

        count = get_messages_count(service, search_query=individual_full_query)
        phrase_counts[phrase] = count
        
        # Overwrite the previous print to show the count result
        print(f"\r[{count:^5}] matches for phrase: '{phrase}'{' ' * 30}", flush=True)


    # 3. Get the overall query (for total count and monthly analysis)
    full_query = create_date_query(FULL_JOB_APPLICATION_QUERY, DAYS_TO_LOOK_BACK)
    
    # 4. Get the total (non-redundant) count 
    total_count = get_messages_count(service, search_query=full_query)

    # 5. Get the dates for ALL matching emails (needed for all advanced date analyses)
    date_objects = get_message_dates(service, full_query)

    # 6. Perform advanced date analyses
    monthly_counts = get_monthly_counts(date_objects)
    day_of_week_counts = get_day_of_week_counts(date_objects)
    hourly_counts = get_hourly_counts(date_objects)
    
    # 7. Output the console result
    start_date = (datetime.now() - timedelta(days=DAYS_TO_LOOK_BACK)).strftime("%Y/%m/%d")
    print("\n--- JOB APPLICATION COUNT SUMMARY ---")
    print(f"Search Period: Emails received after {start_date} (Last {DAYS_TO_LOOK_BACK} days)")
    print(f"Total applications found (Non-Redundant): {total_count}")
    print("-------------------------------------")
    
    # 8. Export to CSV for Power BI
    save_to_csv(monthly_counts, day_of_week_counts, hourly_counts, total_count, DAYS_TO_LOOK_BACK)

    # 9. Visualize all results
    visualize_results(phrase_counts, DAYS_TO_LOOK_BACK)
    visualize_monthly_results(monthly_counts, DAYS_TO_LOOK_BACK)
    visualize_day_of_week_results(day_of_week_counts, DAYS_TO_LOOK_BACK)
    visualize_hourly_results(hourly_counts, DAYS_TO_LOOK_BACK)
    visualize_cumulative_results(date_objects, DAYS_TO_LOOK_BACK)

if __name__ == '__main__':
    main()
