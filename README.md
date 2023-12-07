# H1 Keywords Checker
This script automates the process of verifying if the primary keyword for a webpage is present in the H1 tag. It reads a list of URLs and corresponding keywords from an Excel file and outputs the results to a CSV file.
# Prerequisites
Before running the script, make sure you have the following installed:
* Python 3
* Selenium WebDriver
* ChromeDriver
* openpyxl
* A rank report from Ahrefs in CSV format

⠀Setup
**1** **Install Python Dependencies**:
```bash
pip install selenium openpyxl
```

**2** **ChromeDriver**: Ensure you have ChromeDriver installed and its path correctly set up. The ChromeDriver version should be compatible with your Chrome browser version.
**3** **Prepare the Excel File**: Convert your Ahrefs rank report from CSV to XLSX format. You can do this using a spreadsheet program like Microsoft Excel or Google Sheets:
- Open the CSV file in your spreadsheet program.
- Save the file as h1keywords.xlsx.

**4** **Place the Excel File**: Make sure h1keywords.xlsx is in the same directory as your script, or update the input_file_path variable in the script to match the file's location.
## Running the Script
To run the script, execute the following command in your terminal:
```bash
python h1keys.py
```

The script will process each URL from the Excel file, check for the presence of the primary keyword in the H1 tag, and then save the results to output.csv.
# Output
After completion, the script will display a message indicating that the processing is complete and will provide the path to the saved output file. The output.csv file will contain the following columns:
* **URL**: The webpage URL.
* **Keyword**: The primary keyword associated with the URL.
* **Pass/Fail**: Indicates whether the keyword was found in the H1 tag of the webpage.
