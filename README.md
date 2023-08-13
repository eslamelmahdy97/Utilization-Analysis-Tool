# Utilization-Analysis-Tool
This repository contains a Python script that processes and analyzes data from a CSV file, generates insights, and creates a detailed report in a Word document format. The script has been optimized to reduce the processing time for the given task and has significantly increased the number of accounts benefiting from this service. As a result of its effectiveness, the company has designated this script as the primary tool for generating insights and reports.

## Prerequisites

To run the script successfully, ensure you have the following dependencies installed:

- Python (version 3.x)
- pandas library
- docx library
- datetime module

## Usage

1. Run the script by executing the following command:
   ```
   python Tool.py
   ```

2. You will be prompted to enter the month of choice for analysis.

3. Enter the account name to select the appropriate CSV file for analysis.

4. The script will process the data, generate insights, and create a Word document named `accountname.docx` in the repository directory.

## Analysis and Reporting

The script performs the following tasks:

- Converts CSV data into a pandas DataFrame.
- Prepares the data by converting date columns and extracting claim-related information.
- Analyzes various parameters related to claims and member behavior.
- Generates a detailed report with the following insights:
  - Total consumption (in EGP) and total claims count.
  - Average claim frequency and average claim severity.
  - Total number of using members.
  - Top diseases based on total cost and count of claims.
  - Top providers based on total cost and count of claims.
  - Breakdown of subcategories and insights for each subcategory.
  - Special insights for "Prescription Medicine" subcategory, distinguishing between chronic and acute cases.
  - Insights for the last month's consumption based on user input.
  
The generated Word document will include these insights along with relevant tables and headers for easy readability.

## Customization
This Tool was built based on the data I used so feel free to customize the script according to your specific needs. You can modify the formatting of the generated report, add new analysis parameters, or adjust the way data is processed and visualized.

