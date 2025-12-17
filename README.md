# Eligibility-Checking

A tool that accepts an Excel Sheet from Openjet and another from specific facilities. When used, it checks the eligibility of the files from these facilities and return errors if they exist.

## Recent Improvements (December 2024)

### 🚀 Performance Enhancements
- **Web Workers for Parallel Processing**: Claims are now processed in parallel using Web Workers, significantly improving performance for large datasets (1000+ claims)
- **Real-time Progress Updates**: See live progress percentage and elapsed time during processing
- **Intelligent Worker Distribution**: Automatically uses optimal number of workers based on available CPU cores
- **Graceful Fallback**: Falls back to single-threaded processing if Web Workers are unavailable or encounter errors

### 🐛 Critical Bugs Fixed
1. **VVIP Member Check**: Fixed bug where VVIP members were incorrectly validated instead of auto-approved
2. **Worker Distribution**: Fixed critical bug where claims could be silently dropped during parallel processing
3. **File Upload UI**: Removed misleading "multiple" attribute from file input

### 🔍 Enhanced Debug Logging
- Debug logs now include detailed ineligibility reasons for each eligibility record:
  - Date mismatches with specific dates
  - Clinician name mismatches
  - Service category validation failures
  - Status check failures
- Generate comprehensive debug reports from the modal interface

## Features
- Upload and validate patient eligibility reports
- Support for multiple report formats (Clinicpro, Odoo, InstaHMS)
- Real-time validation against eligibility sheets
- Filter by insurance provider (Daman/Thiqa)
- Export invalid claims to Excel
- Detailed modal views for eligibility records
- Debug log generation for troubleshooting

## Usage
1. Upload your patient report (Clinicpro, Odoo, or InstaHMS format)
2. Upload the Eligibility XLSX file
3. Click **Process** to check eligibility
4. Review results in the table
5. Click **Export Invalid Rows** to download patients not eligible

## Performance
- Small datasets (< 100 claims): Instant processing
- Medium datasets (100-1000 claims): Uses Web Workers for parallel processing
- Large datasets (1000+ claims): Efficient parallel processing with progress tracking

## Browser Compatibility
- Modern browsers with Web Worker support (Chrome, Firefox, Safari, Edge)
- Falls back to single-threaded processing on older browsers

## Technical Details
- Client-side processing (no server required)
- Uses XLSX.js for Excel file parsing
- Web Workers API for parallel processing
- Bootstrap 5 for UI components

## Known Limitations
- All processing happens in the browser (requires sufficient memory for large files)
- Date parsing supports common formats but may not handle all international formats
- See BUGS_AND_IMPROVEMENTS.md for detailed technical information

## Development
- Main application: `elig.html` and `elig.js`
- Worker script: `eligibility-worker.js`
- Styles: `tables.css`
- Entry point: `index.html`

## Bug Reports
Found an issue? Please report it to the [developer](https://github.com/EvacionSaraak/Submission-Checker-Tools).

