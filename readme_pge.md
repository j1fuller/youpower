# YouPower PG&E Green Button Data Tool

## Overview

This tool automates the process of downloading Green Button Data (GBD) from PG&E's online portal and processes it into a formatted Excel workbook for billing calculations. The tool supports Time-of-Use C (TOU-C) rate plans and can be adapted for other rate structures.

## Features

- Automated login to PG&E's customer portal
- Selection of date ranges for data retrieval
- Support for multiple utility accounts
- Conversion of GBD data to formatted Excel workbooks
- Calculation of electricity usage by time period (On-Peak, Off-Peak)
- TOU-C rate calculation with tiered pricing
- GUI interface for easy operation by non-technical users

## Installation

### Prerequisites

- Windows 10 or above
- Chrome browser installed (for the Selenium automation)

### Setup

1. Download the latest `youpower_pge.exe` file from the releases page
2. Place it in a designated folder with read/write permissions
3. Ensure you have the following files in the same directory:
   - `icon.ico`
   - `logo.png`

## Usage

1. Launch the application by double-clicking `youpower_pge.exe`
2. Select "PG&E" from the utility provider dropdown
3. Enter your PG&E username and password
4. Select the start and end dates for the data you wish to retrieve
5. Choose a download folder for the GBD files and Excel output
6. Check "Process to Excel after download" to automatically create the formatted Excel file
7. Click "Start Automation" to begin the process

## Excel Output Structure

The generated Excel workbook contains the following sheets:

- **Data**: Contains the main GBD data and calculations
- **Pricing Variables**: TOU-C rate information for On-Peak and Off-Peak periods
- **Baseline Allowances**: Climate zone baseline allocations
- **Weekday Time Table**: Hourly time period designations for weekdays
- **Weekend & Holiday Time Table**: Hourly time period designations for weekends and holidays

## Building from Source

To build the executable from source:

1. Ensure you have Python 3.8+ installed
2. Install required packages:
   ```
   pip install pyinstaller selenium webdriver-manager pandas openpyxl PyQt5
   ```
3. Run the build script:
   ```
   python build_pge_scraper.py
   ```

## Development Notes

- The TOU-C rates and climate zone baseline allowances are placeholders and should be updated with current values
- The time period definitions (4PM-9PM peak) are based on current PG&E TOU-C structure
- Additional utility providers (SCE) can be implemented following the same pattern

## Troubleshooting

- If the application cannot find the Green Button on PG&E's website, try manually logging in to verify the correct navigation path
- For login issues, ensure your credentials are correct and that your account doesn't have multi-factor authentication enabled
- If the Excel processing fails, check that the downloaded GBD file is in the expected format (.xml or .csv)

## License

Copyright © 2025 YouPower Pvt Ltd. All rights reserved.
