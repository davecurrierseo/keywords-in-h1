# Selenium SEO Analysis Script

## Description

This Python script automates the process of analyzing SEO-related attributes of web pages listed in an Excel file. It checks for the presence of specified keywords in various parts of a web page, including the title, H1 tag, URL, paragraph tags, image URLs, and alt attributes. It operates in headless mode, meaning it doesn't open an actual browser window. The script generates a CSV file with the results of the analysis.

## Features

- **Keyword Analysis**: Checks if a given keyword is present in the page title, H1 tag, URL, paragraph tags, and image attributes.
- **Document Analysis**: Determines if a URL points to a document such as a PDF or Word file.
- **Progress Tracking**: Updates a text-based progress bar in the terminal to indicate the processing status.
- **Output Generation**: Outputs the results into a CSV file with a unique name to avoid overwriting existing files.

## Dependencies

- **Python 3**: The script is written in Python 3 and requires it to run.
- **Selenium**: For automating web browser interaction.
- **openpyxl**: For reading from and writing to Excel files (`.xlsx`).
- **Chromedriver**: The Chrome WebDriver for Selenium.

## Setup

1. **Install Python 3**: Download and install Python 3 from the [official website](https://www.python.org/).
2. **Install Dependencies**: Run the following command to install the necessary Python packages:

    ```sh
    pip install selenium openpyxl
    ```

3. **Chromedriver**: Download the appropriate version of Chromedriver for your system from the [Chromedriver download page](https://sites.google.com/a/chromium.org/chromedriver/downloads). Ensure it's in your PATH or specify its location in the script.

4. **Input File**: Prepare an Excel file (`rank.xlsx`) with two columns: one for keywords and the other for URLs.

## Usage

1. **Configure Script**: Ensure that the path to `chromedriver` and the input/output file paths in the script are correct.
2. **Run Script**: Execute the script with the following command:

    ```sh
    python3 seo_analysis.py
    ```

3. **Results**: After completion, the script will provide the path to the generated CSV file containing the analysis results.

## Notes

- This script is designed for macOS. If using a different OS, modifications might be necessary.
- The script prevents the system from sleeping during execution. Ensure that this behavior is acceptable for your use case.
- The script does not handle all possible exceptions and edge cases. Use it as a starting point and modify it according to your needs.

## Contributing

Contributions to enhance the functionality, improve the efficiency, or fix issues in the script are welcome.

## License

Specify your license or if the script is open for public use.

