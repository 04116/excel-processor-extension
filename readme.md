# PowerBI Data Export Extension

## Overview

Chrome extension that downloads data from PowerBI reports and creates a new Excel file with separate sheets for each time period.

## Features

- Downloads multiple PowerBI reports with different time periods
- Creates a new Excel workbook with each PowerBI file as a separate sheet
- Automatically unmerges cells for better data processing
- Includes a summary sheet with export details
- Handles download failures with placeholder sheets

## Time Periods

- **Day_Le**: Yesterday vs Same Day Last Year
- **MTD_Le**: Month to Date vs Same Period Last Year
- **09.06-now**: June 9 to Today vs Same Period Last Year

## Installation

### For Users (Chrome Extension)

1. Download the release zip file
2. Extract it to a folder on your computer
3. Open Chrome and go to `chrome://extensions/`
4. Enable "Developer mode" (toggle in top right)
5. Click "Load unpacked" and select the extracted folder
6. The extension icon will appear in your toolbar

### For Developers

1. Clone this repository
2. Run `./build-release.sh` to create a release package
3. Follow the user installation steps above

## Usage

1. Navigate to app.powerbi.com and open a report
2. Click the extension icon
3. Wait for time periods to load
4. Click "Download PowerBI Data & Create Excel File"
5. The extension will download all PowerBI files and create a new Excel file

## Output

- New Excel file named `PowerBI_Data_[timestamp].xlsx`
- Each PowerBI report becomes a separate sheet
- Summary sheet with export statistics
- All cells are unmerged for easy data processing
