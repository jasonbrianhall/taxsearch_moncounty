# Monongalia County Tax Search Tool

A command-line utility for searching and retrieving tax records from the Monongalia County Sheriff's Office website.

## Features

- Search tax records by various criteria:
  - Taxpayer name
  - Account number
  - Tax ticket number
  - District/Map/Parcel information
- Export results to multiple formats (Excel, CSV, JSON, TXT)
- Automatic pagination handling for large result sets
- Detailed logging for troubleshooting
- Multiple filtering options (by year, property type, payment status)
- Support for custom URLs to work with different county systems

## Requirements

### Basic Requirements
- Python 3.6+
- Required Python packages:
  - requests
  - urllib3

### Optional Dependencies
- openpyxl (for Excel export)
- tqdm (for progress bars)

## Installation

1. Clone this repository or download the script:
   ```
   git clone https://github.com/your-username/moncounty-tax-search.git
   cd moncounty-tax-search
   ```

2. Install required packages:
   ```
   pip install requests urllib3
   ```

3. Install optional dependencies:
   ```
   pip install openpyxl tqdm
   ```

## Usage

### Basic Command Structure

```
python moncountysearch.py [search_type] [search_parameters] [options]
```

### Search Types

#### Search by Name
```
python moncountysearch.py name "DOE JOHN" --output results.xlsx
```

#### Search by Account Number
```
python moncountysearch.py account "12345" --output results.csv
```

#### Search by Ticket Number
```
python moncountysearch.py ticket "2023" "67890" --suffix "A" --output results.json
```

#### Search by Map/Parcel
```
python moncountysearch.py map --district "01" --map "123" --parcel "456" --output results.txt
```

### Common Options

| Option | Description |
|--------|-------------|
| `--domain`, `-dm` | Domain to search (default: monongalia.softwaresystems.com) |
| `--url`, `-u` | Full URL to search (default: https://monongalia.softwaresystems.com/SEARCH.html) |
| `--output`, `-o` | Output file (supports .xlsx, .json, .csv, .txt) |
| `--limit-year`, `-ly` | Limit search to specific tax year |
| `--prop-type`, `-pt` | Property type: B=Both, R=Real, P=Personal (default: B) |
| `--status`, `-st` | Payment status: B=Both, P=Paid, U=Unpaid (default: B) |
| `--district`, `-d` | Filter by district code |
| `--max-pages`, `-mp` | Maximum number of pages to retrieve |
| `--debug`, `-v` | Increase verbosity (use -v, -vv for more detail) |
| `--timeout`, `-to` | Request timeout in seconds (default: 30) |

### Advanced Usage Examples

#### Limit search to specific tax year and property type
```
python moncountysearch.py name "SMITH JANE" --limit-year "2023" --prop-type "R" --output results.xlsx
```

#### Search for unpaid taxes only
```
python moncountysearch.py name "SMITH JANE" --status "U" --output unpaid.xlsx
```

#### Limit number of pages retrieved
```
python moncountysearch.py name "SMITH" --max-pages 5 --output smith_partial.xlsx
```

#### Use a different county system
```
python moncountysearch.py name "SMITH" --url "https://othercounty.softwaresystems.com/SEARCH.html" --output smith_other.xlsx
```

#### Use custom domain and URL path
```
python moncountysearch.py name "SMITH" --domain "othercounty.gov" --url "https://othercounty.gov/tax/SEARCH.html" --output smith_custom.xlsx
```

#### Increase logging verbosity
```
python moncountysearch.py name "SMITH" -vv --output smith_debug.xlsx
```

#### Inspect log files (useful for debugging)
```
python moncountysearch.py --inspect
```

## Output Formats

### Excel (.xlsx)
- Creates a formatted spreadsheet with two sheets:
  - "Tax Records" containing all tax records found
  - "Search Info" containing metadata about the search

### CSV (.csv)
- Simple comma-separated values file containing tax records

### JSON (.json)
- Structured JSON output with headers, data rows, and pagination info

### Text (.txt)
- Tab-delimited text file with basic formatting

## Logging

The tool automatically creates detailed logs in the `logs` directory:
- `tax_search_YYYYMMDD_HHMMSS.log`: Log messages for each search session
- `response_YYYYMMDD_HHMMSS.html`: Raw HTML responses (useful for debugging)

## Troubleshooting

### Common Issues

1. **SSL Certificate Verification Failed**
   - The tool disables SSL verification by default to work with the county's website

2. **No Results Found**
   - Verify your search parameters
   - Try using a more general search term
   - Check the log files for details

3. **Excel Output Fails**
   - Make sure openpyxl is installed: `pip install openpyxl`

4. **Slow Performance**
   - Large result sets can take time to process
   - Use `--max-pages` to limit the number of pages retrieved

5. **URL/Domain Issues**
   - If connecting to a different county system:
     - Ensure the URL parameter ends with `SEARCH.html` or the tool will append it
     - Make sure the domain parameter matches the actual domain for cookie setting
     - Check logs for any connection errors

### Inspecting Log Files

Use the inspect feature to check the contents of response logs:
```
python moncountysearch.py --inspect
```

## Using with Multiple County Systems

This tool can now work with different county systems that use the same software platform:

1. **Default Configuration**: 
   - Works with Monongalia County (WV) by default
   - URL: https://monongalia.softwaresystems.com/SEARCH.html

2. **Other Counties**:
   - Use the `--url` parameter to specify a different county's search page
   - Use the `--domain` parameter if cookie handling is required with the custom URL
   - Example: `python moncountysearch.py name "SMITH" --url "https://othercounty.gov/tax/SEARCH.html" --domain "othercounty.gov"`

3. **URL Handling**:
   - If the provided URL doesn't end with "SEARCH.html", the tool will append it automatically
   - The domain is extracted from the URL for cookie handling purposes

## License

[MIT License](LICENSE)

## Disclaimer

This tool is not affiliated with or endorsed by the Monongalia County Sheriff's Office or any other county government. It is provided as-is with no warranty. Use at your own risk.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
