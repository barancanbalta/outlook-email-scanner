# Outlook Email Address Scanner

A comprehensive Python tool to extract and collect unique email addresses from Microsoft Outlook mailboxes. This tool scans sent items and archive folders across all configured Outlook accounts to build a complete list of email contacts.

## Features

- **Comprehensive Scanning**: Scans sent items and archive folders from all Outlook accounts
- **Dual-Mode Processing**: Fast table mode for quick scanning + detailed iterative mode for complete recipient information
- **Multiple Export Formats**: Exports to both Excel (.xlsx) and CSV formats
- **Checkpoint System**: Automatic backups every 1,000 records to prevent data loss
- **Progress Tracking**: Real-time progress logs with ETA calculations
- **Error Handling**: Robust error handling with detailed logging
- **Date Filtering**: Optional date range filtering to scan recent emails only
- **Multi-Account Support**: Automatically processes all configured Outlook accounts

## Requirements

- Windows OS (required for Outlook COM interface)
- Microsoft Outlook installed and configured
- Python 3.8 or higher

## Installation

1. Clone this repository:
```bash
git clone https://github.com/barancanbalta/outlook-email-scanner.git
cd outlook-email-scanner
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Run the script with default settings (scans all emails from all accounts):
```bash
python outlook_email_scanner.py
```

### Advanced Usage

You can customize the script by modifying the parameters in the `if __name__ == "__main__"` section:

```python
run(
    restrictdays=365,          # Scan only last 365 days (None = all emails)
    output=Path("mycontacts.xlsx"),  # Custom output file path
    verbose=True,              # Enable detailed debug logging
    maxitems=1000,             # Limit items per folder (for testing)
    usetablemode=True,         # Use fast GetTable API
    savecheckpoints=True,      # Save checkpoint backups
    logfile=Path("scan.log"),  # Log file path
    exportcsv=True             # Also export CSV format
)
```

### Parameters

- `restrictdays` (int, optional): Number of days to scan back. `None` scans all emails.
- `output` (Path, optional): Output file path. `None` generates automatic filename with timestamp.
- `verbose` (bool): Enable detailed debug logging.
- `maxitems` (int, optional): Maximum items to process per folder (useful for testing).
- `usetablemode` (bool): Use fast GetTable API for initial scanning.
- `savecheckpoints` (bool): Save checkpoint backups every 1,000 records.
- `logfile` (Path, optional): Path to log file. `None` logs only to console.
- `exportcsv` (bool): Also export results as CSV file.

## Output

The script generates:

1. **Excel File (.xlsx)**: Contains two columns:
   - E-posta: Email address
   - Ad: Contact name (if available)

2. **CSV File (.csv)**: Same data in CSV format for easy import

3. **Checkpoint Files (optional)**: Backup files saved every 1,000 records as `checkpoint_*.csv`

4. **Log File (optional)**: Detailed execution log with timestamps

## How It Works

1. **Connection**: Connects to Outlook using COM interface
2. **Account Discovery**: Finds all configured Outlook accounts
3. **Folder Scanning**: Scans sent items and archive folders for each account
4. **Dual-Mode Processing**:
   - **Table Mode**: Fast scanning using GetTable API (email addresses only)
   - **Iterative Mode**: Detailed scanning for complete recipient information including names
5. **Email Normalization**: Validates and normalizes email addresses
6. **Deduplication**: Removes duplicate email addresses
7. **Export**: Saves results to Excel and CSV formats

## Notes

- The script requires Outlook to be installed and configured on your system
- Large mailboxes may take significant time to process
- Checkpoint files help prevent data loss during long scans
- The script handles errors gracefully and continues processing

## Author

For questions or contributions, please open an issue on GitHub.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the issues page.

## Disclaimer

This tool is for personal use and legitimate business purposes. Ensure you have proper authorization before scanning email accounts. The author is not responsible for any misuse of this software.
