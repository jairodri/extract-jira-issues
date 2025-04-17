# JIRA Issues Extractor

A professional automation tool that extracts JIRA issues based on predefined filters, generates formatted Excel reports, and creates email drafts with the reports attached.

## Features

- Automated extraction of JIRA issues using configurable JQL filters
- Excel report generation with professional formatting:
  - Multiple sheets (one per filter)
  - Formatted headers
  - Clickable hyperlinks
  - Properly formatted dates
  - Auto-adjusted column widths
  - Auto-filtering capabilities
- Outlook email draft creation with:
  - Configurable recipients (TO and CC)
  - Dynamic content based on extracted data
  - Excel report attachment
  - HTML formatted body

## Requirements

- Python 3.6+
- Chrome browser
- Microsoft Outlook

## Dependencies

```
selenium
webdriver-manager
pandas
python-dotenv
openpyxl
dateutil
pywin32
```

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/extract-jira-issues.git
   cd extract-jira-issues
   ```

2. Install required packages:
   ```
   pip install -r requirements.txt
   ```

3. Create and configure your `.env` file (see Configuration section).

## Configuration

Create a `.env` file in the project directory with the following parameters:

```
# JIRA Configuration
JIRA_URL_BASE="https://your-jira-instance.com/issues/"

# JIRA filters with JQL queries
FILTER_PATTERN="JIRA_FILTER_"
JIRA_FILTER_UNASSIGNED="category = YourCategory AND assignee is EMPTY AND status = \"To Do\""
JIRA_FILTER_IN_PROGRESS="category = YourCategory AND status not in (Closed, Cancelled, \"To Do\")"

# Browser Configuration
HEADLESS_MODE=True
WAIT_TIME=10
WAIT_ELEMENT=.issue-table-wrapper

# Email Configuration
MAIL_PATTERN="MAIL_TO_"
MAIL_TO_01="user1@company.com"
MAIL_TO_02="user2@company.com"

MAIL_PATTERN_CC="MAIL_CC_"
MAIL_CC_01="manager@company.com"

MAIL_SUBJECT="JIRA Issues Report - "
MAIL_BODY_TEMPLATE="<html><body><p>Hello team,</p><p>Attached is the updated JIRA issues report generated on {FECHA}.</p><p>The report contains {NUM_ISSUES} issues distributed across {NUM_PESTANAS} sheets:</p>{LISTA_PESTANAS}<p>Regards,</p></body></html>"
```

## Usage

Run the script with:

```
python jira_issues.py
```

## Workflow

The script executes the following workflow:

1. **Environment Setup**:
   - Loads environment variables from `.env`
   - Configures execution parameters (headless mode, timeouts)

2. **JIRA Filters Loading**:
   - Retrieves JIRA filters from environment variables
   - Formats sheet names for readability

3. **Browser Initialization**:
   - Creates Chrome WebDriver instance
   - Configures browser options (headless mode if enabled)

4. **Data Extraction**:
   - For each JIRA filter:
     - Builds URL with encoded JQL
     - Navigates to the URL
     - Waits for issue table to load
     - Extracts issue data (keys, types, links, summaries, statuses, etc.)
     - Stores data in a DataFrame
     - Adds DataFrame to a dictionary using sheet name as key

5. **Excel Report Generation**:
   - Creates an Excel file with current timestamp
   - For each DataFrame:
     - Creates a sheet with formatted name
     - Applies professional formatting (bold headers, color fills)
     - Converts issue links to clickable hyperlinks
     - Formats date columns properly
     - Adjusts column widths based on content
     - Enables filtering for all columns

6. **Email Draft Creation**:
   - Loads recipient lists (TO and CC) from environment variables
   - Generates dynamic content based on extracted data
   - Creates an Outlook draft with:
     - Configured recipients
     - Subject line with timestamp
     - HTML formatted body
     - Excel report attached

7. **Cleanup**:
   - Removes temporary Excel file (if email draft created successfully)
   - Closes browser and releases resources

## Core Functions

- `create_chrome_driver()`: Initializes and configures Chrome WebDriver
- `navigate_to_url()`: Navigates to JIRA URL and waits for elements to load
- `load_jira_filters()`: Retrieves JIRA filters from environment variables
- `extract_jira_issues()`: Extracts issue data from JIRA tables
- `generate_excel_report()`: Creates and formats Excel report with multiple sheets
- `generate_email_draft()`: Creates Outlook email draft with dynamic content
- `format_sheet_name()`: Formats sheet names for better readability

## Customization

You can customize the tool by:

- Adding more JIRA filters in the `.env` file
- Modifying email templates
- Extending the data extraction to include additional fields
- Changing Excel formatting options

## License

MIT License
