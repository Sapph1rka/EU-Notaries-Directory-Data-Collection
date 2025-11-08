# EU Notaries Directory Data Collection

A Python web scraping tool designed to collect and organize notary information from the [EU Notaries Directory](https://notaries-directory.eu/). The tool extracts detailed information about notaries across European countries and saves the data in a formatted Excel file.

## Features

- **Automated Web Scraping**: Efficiently scrapes notary data from multiple pages
- **Parallel Processing**: Uses multi-threading to process multiple notaries simultaneously
- **Incremental Data Saving**: Saves data incrementally to prevent data loss
- **Resume Capability**: Can resume interrupted scraping sessions by merging with existing data
- **Formatted Excel Output**: Automatically formats the output Excel file with styled headers and optimized column widths
- **Error Handling**: Robust error handling with comprehensive logging
- **Duplicate Prevention**: Automatically removes duplicate entries

## Data Collected

The tool collects the following information for each notary:

- **Full Name**: Complete name of the notary
- **First Name**: First name extracted from the full name
- **Email**: Email address (if available)
- **Country**: Country where the notary is located

### Example Output Data

The following table shows an example of the collected data structure:

| Full Name | First Name | Email | Country |
|-----------|------------|-------|---------|
| Dr. Maria Schmidt | Dr. | maria.schmidt@notary.de | Germany |
| Jean-Pierre Dubois | Jean-Pierre | jp.dubois@notaire.fr | France |
| Alessandro Rossi | Alessandro | a.rossi@notaio.it | Italy |
| Ana García López | Ana | ana.garcia@notario.es | Spain |
| Jan Kowalski | Jan | j.kowalski@notariusz.pl | Poland |
| Sophie Van der Berg | Sophie | s.vanderberg@notaris.nl | Netherlands |
| Thomas Andersen | Thomas | thomas@notar.dk | Denmark |

> **Note**: The header row in the actual Excel file has a yellow background with bold text, matching the formatting shown above.

## Output Format

Data is saved to `notaries_data.xlsx` with the following characteristics:

- **Header Row**: Yellow background with bold text
- **Column Widths**: Optimized for readability
  - Full Name: 60 characters
  - First Name: 40 characters
  - Email: 60 characters
  - Country: 40 characters
- **Data Format**: Excel (.xlsx) format compatible with Microsoft Excel, Google Sheets, and other spreadsheet applications

## Requirements

- Python 3.8 or higher
- Chrome browser installed
- ChromeDriver (automatically managed by Selenium)

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd EU-Notaries-Directory-Data-Collection
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
```

3. Activate the virtual environment:
   - On Windows:
   ```bash
   venv\Scripts\activate
   ```
   - On macOS/Linux:
   ```bash
   source venv/bin/activate
   ```

4. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Ensure Chrome browser is installed on your system
2. Run the script:
```bash
python main.py
```

3. The script will:
   - Start scraping from page 0
   - Process notaries in batches of 7 (configurable)
   - Save data incrementally to `notaries_data.xlsx`
   - Continue until all pages are processed

4. Monitor progress through console logs

5. The output file `notaries_data.xlsx` will be created/updated in the same directory

## How It Works

### Scraping Process

1. **Page Navigation**: The script navigates through paginated search results (`/en/search?page={page}`)

2. **Element Extraction**: For each page, it identifies all notary listing elements using the `list-element` class

3. **Parallel Processing**: 
   - Groups notaries into batches (default: 7 per batch)
   - Creates separate threads for each notary in the batch
   - Each thread opens a new browser instance to scrape the notary's detail page

4. **Data Extraction**: For each notary profile page, the script extracts:
   - Country from the profile header
   - Email address from the contact information section
   - Full name (already available from the listing page)

5. **Data Storage**:
   - Data is collected in a queue
   - After each batch, data is saved to Excel
   - If the file already exists, new data is merged with existing data
   - Duplicates are automatically removed

6. **Formatting**: After saving, the Excel file is formatted with:
   - Yellow header background
   - Bold header text
   - Optimized column widths

### Data Storage Details

- **File Format**: Excel (.xlsx) using openpyxl library
- **Data Structure**: 
  - Columns: Full Name, First Name, Email, Country
  - Rows: One per notary
- **Incremental Updates**: Each batch of processed notaries is immediately saved
- **Duplicate Handling**: Uses pandas `drop_duplicates()` to ensure unique records
- **Sorting**: Data is sorted by the original index before saving

## Configuration

You can modify the following constants in `main.py` to customize behavior:

- `BATCH_SIZE`: Number of notaries to process in parallel (default: 7)
- `PAGE_LOAD_DELAY`: Seconds to wait for page load (default: 1.5)
- `THREAD_START_DELAY`: Delay between thread starts in seconds (default: 0.1)
- `WAIT_TIMEOUT`: Selenium wait timeout in seconds (default: 10)
- `OUTPUT_FILENAME`: Output Excel filename (default: "notaries_data.xlsx")
- `COLUMN_WIDTHS`: Dictionary of column widths for Excel formatting

## Error Handling

The tool includes comprehensive error handling:

- **Timeout Errors**: Logged and skipped, allowing the script to continue
- **Missing Elements**: Gracefully handles missing email or country information
- **Network Issues**: Logs errors and continues with next notary
- **File Errors**: Logs file operation errors without crashing
- **Keyboard Interrupt**: Gracefully handles Ctrl+C to stop scraping

## Logging

The script uses Python's logging module with the following log levels:

- **INFO**: General progress information (pages processed, records saved)
- **WARNING**: Non-critical issues (missing optional data)
- **ERROR**: Critical errors that prevent data collection
- **DEBUG**: Detailed debugging information

Logs are displayed in the console with timestamps.

## Limitations

- Requires stable internet connection
- Dependent on website structure (XPath selectors may need updates if site changes)
- Chrome browser must be installed
- Processing time depends on number of notaries and network speed

## Troubleshooting

### ChromeDriver Issues
If you encounter ChromeDriver errors, ensure:
- Chrome browser is up to date
- Selenium version is compatible with your Chrome version
- Consider using `webdriver-manager` for automatic driver management

### No Data Collected
- Check internet connection
- Verify the website is accessible
- Check console logs for error messages
- Ensure XPath selectors are still valid (website structure may have changed)

### Excel File Not Opening
- Ensure openpyxl is properly installed
- Check file permissions
- Verify the file isn't open in another program

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is provided as-is for educational and research purposes.

## Disclaimer

This tool is for data collection purposes only. Please ensure you comply with:
- The website's Terms of Service
- Robots.txt guidelines
- Applicable data protection regulations (GDPR, etc.)
- Rate limiting and respectful scraping practices

Always use web scraping tools responsibly and ethically.
