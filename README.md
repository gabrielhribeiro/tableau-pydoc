# Tableau Report Automation with Python

This repository contains a Python script that automates the process of generating reports using data from Tableau Server. The script streamlines the extraction, processing, and insertion of data into a Word document template, making the reporting process more efficient and less time-consuming.

## Features

- **Essential Libraries**: The script utilizes libraries such as `tableauserverclient`, `pandas`, `io`, `BytesIO`, `Document` from `docx`, and `datetime`.
- **Tableau Server Connection**: Establishes a connection with Tableau Server using the provided URL and authentication token.
- **Auxiliary Functions**: Two helper functions, `locate_workbook_id` and `locate_view_id`, are used to obtain the IDs of workbooks and views.
- **User Inputs**: Collects user inputs for the date range and reporting target.
- **Data Extraction**: Retrieves data from various views as CSV files based on the specified date range.
- **Data Processing**: Imports CSV data into pandas DataFrames for further manipulation.
- **Word Document Template**: Replaces placeholders in the Word template with values from the DataFrames.
- **Image Insertion**: Embeds an image from Tableau Server into the Word document.
- **Exception Handling**: Implements exception management to catch and display any errors that occur during execution.
- **Report Generation**: Creates a Word document containing the processed Tableau data and saves it with a filename that includes the current date.

## Usage

1. Clone the repository to your local machine.
2. Install the required Python libraries.
3. Set up your Tableau Developer account and obtain API credentials.
4. Update the script with your Tableau Server URL and authentication token.
5. Customize the script to match your specific reporting requirements.
6. Run the script to generate a report based on the provided date range and target.

## Output

The generated Word document will contain the processed Tableau data, including description, measurements, categories, form data, network destinations, vendor name, sensitive data, and tables. Upon successful completion, the script will notify the user and provide the path to the generated report.

## Contributing

Feel free to submit issues, fork the repository, and create pull requests to contribute to the project. Your feedback and contributions are greatly appreciated!
