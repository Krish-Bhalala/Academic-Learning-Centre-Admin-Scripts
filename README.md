# Acadmeic-Learning-Centre-Admin-Scripts
## Welcome to the Automated Timesheet Generator! 🎉 
Picture this: every week, you sit down to look at tutoring schedules and fill out timesheets. It feels like a boring chore that never ends. You open the schedule, write down the hours, and hope you don’t miss anything. 
I thought, “This is too much work!” So, I created this fun Python script to help. It fetches your tutoring appointment data from a scheduling system and fills out your Excel timesheet for you. 
Now, instead of getting lost in paperwork, you can relax and let the script do the hard work. It gathers your appointment details, organizes them, and updates your timesheet—all while you enjoy your coffee (or tea, no judgment here). ☕✨

### Features

- Fetches tutor appointment data from a web-based scheduling system
- Parses the raw HTML data to extract relevant information
- Automatically updates an Excel timesheet template with the fetched data
- Handles date ranges automatically, defaulting to the previous Saturday through Friday
- Saves the updated timesheet with a new filename

### Prerequisites

- Python 3.x
- Required Python libraries:
  - openpyxl
  - beautifulsoup4
  - requests (implicit in the http.client usage)

You can install the required libraries using pip:

```
pip install openpyxl beautifulsoup4 requests
```

### Configuration

Before running the script, you need to set up a few configuration variables:

1. `directory_path`: Set this to the path of the folder containing your timesheet Excel template.
2. `file_path`: Set this to the full path of your timesheet Excel template.
3. `TUTOR_ID`: Replace with your actual tutor ID.
4. `EMPLOYEE_ENDPOINT`: Set this to the correct endpoint URL for fetching employee data.

### Usage

1. Ensure all configuration variables are set correctly.
2. Run the script:

```
python timesheet_generator.py
```

The script will:
- Fetch your appointment data for the previous week (Saturday to Friday)
- Parse the data and extract relevant information
- Update the Excel template with the fetched data
- Save a new Excel file with the updated information

### Important Notes

- This script uses an unverified SSL context. In a production environment, proper SSL verification should be implemented.
- The script assumes a specific structure (issued by academic learning center) for the Excel template. Ensure your template matches the expected format.
- Sensitive information like cookies and employee-only endpoints have been removed from the provided code for privacy reasons. You'll need to add these back in for the script to function properly.

### Customization

You can customize various aspects of the script, such as:
- The date range for fetching appointments (modify the `updateDate` function)
- The Excel template structure (adjust the `COLUMN_SEQUENCE` and cell references)
- The formatting of the fetched data (modify the `format_data` function)

### Troubleshooting

If you encounter any issues:
1. Check that all required libraries are installed
2. Verify that all configuration variables are set correctly
3. Ensure you have the necessary permissions to access the scheduling system and write to the specified directory

### Disclaimer

This script is provided as-is, without any guarantees. Always verify the generated timesheet for accuracy before submission.
