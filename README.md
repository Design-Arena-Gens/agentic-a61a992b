# PolicyBoss Car Insurance Scraper

Production-ready Python scraping agent that fetches car insurance expiry dates using **backend API calls** (not HTML/XPath scraping).

## Features

✅ **API-Based Extraction** - Uses PolicyBoss backend APIs, not DOM scraping
✅ **Persistent Chrome Profile** - Reuses logged-in session cookies
✅ **Auto API Detection** - Monitors network requests to detect API endpoints
✅ **Retry Logic** - 3 retries per vehicle with exponential backoff
✅ **Error Handling** - Graceful handling of timeouts, invalid numbers, session expiry
✅ **Excel I/O** - Reads vehicle numbers from Excel, outputs results to Excel
✅ **Production Ready** - Console logging, status tracking, comprehensive error messages

## Requirements

- Python 3.10+
- Google Chrome browser
- ChromeDriver (auto-installed via webdriver-manager)

## Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

## Setup

### 1. Prepare Input File

Create `input.xlsx` with vehicle registration numbers in Column A:

| Vehicle Number |
|----------------|
| RJ45CR3119     |
| DL3CAA1234     |
| MH12AB5678     |

Or use the provided sample:
```bash
python create_sample_input.py
```

### 2. Run the Scraper

```bash
python scraper.py
```

### 3. First Run - Manual Login

On first run:
1. Chrome will open automatically
2. **Login to PolicyBoss** manually in the browser
3. Press ENTER in the terminal when logged in
4. The script will capture your session cookies

### 4. API Detection

The scraper will ask you to:
1. Search for a vehicle number in the browser
2. This helps detect the backend API endpoint
3. Press ENTER after submitting the search

Once detected, the API endpoint will be reused for all vehicles.

### 5. Automatic Processing

The scraper will:
- Use the detected API to fetch data for each vehicle
- Parse JSON responses (not HTML)
- Extract expiry dates and format them as dd/mm/yyyy
- Save results to `output.xlsx`

## Output Format

`output.xlsx` contains:

| Vehicle Number | Expiry Date | Status     |
|----------------|-------------|------------|
| RJ45CR3119     | 15/06/2024  | FOUND      |
| DL3CAA1234     |             | NOT_FOUND  |
| MH12AB5678     | 20/12/2024  | FOUND      |

**Status Values:**
- `FOUND` - Expiry date successfully retrieved
- `NOT_FOUND` - Vehicle not found or no expiry date in API response
- `ERROR` - API timeout or request failed
- `SESSION_EXPIRED` - Need to re-login

## How It Works

### 1. Session Management
- Chrome launches with persistent profile (`./chrome_profile/`)
- Session cookies are captured and reused
- No need to login on subsequent runs

### 2. API Detection
The scraper monitors browser network traffic to detect PolicyBoss API endpoints:
- Analyzes Chrome performance logs
- Identifies JSON API calls related to vehicle/insurance/policy
- Falls back to common API patterns if detection fails

### 3. API Request Flow
```
Vehicle Number → API Request → JSON Response → Parse Expiry Date → Excel Output
```

### 4. Common API Patterns Tested
If auto-detection fails, the scraper tries these patterns:
```
/api/v1/vehicle/search?regNo={vehicle_number}
/api/vehicle/details?registration={vehicle_number}
/api/insurance/vehicle?regNo={vehicle_number}
/api/car-insurance/vehicle-details?regNo={vehicle_number}
```

## Detected API Structure

After running, the scraper will display the detected API endpoint and its structure.

**Example API Endpoint:**
```
https://www.policyboss.com/api/v1/vehicle/search?regNo=RJ45CR3119
```

**Example API Response:**
```json
{
  "status": "success",
  "data": {
    "vehicleNumber": "RJ45CR3119",
    "expiryDate": "2024-06-15T00:00:00Z",
    "insurerName": "ICICI Lombard",
    "policyNumber": "POL123456"
  }
}
```

The scraper automatically parses nested JSON structures to find:
- Expiry date fields: `expiryDate`, `policy_expiry_date`, `insuranceExpiryDate`, etc.
- Insurer fields: `insurerName`, `insurer`, `companyName`, etc.

## Edge Cases Handled

✅ **Vehicle Not Found** - Marked as NOT_FOUND
✅ **API Timeout** - Retries 3 times with delays
✅ **Invalid Registration** - Logged as ERROR
✅ **Session Expired** - Prompts for re-login
✅ **Empty Cells** - Skipped gracefully
✅ **Malformed Dates** - Attempts multiple date formats

## Retry Logic

- **Max Retries:** 3 per vehicle
- **Delay:** 2 seconds between retries
- **Timeout:** 15 seconds per API request

## Code Structure

```python
PolicyBossScraper
├── setup_chrome()          # Launch Chrome with persistent profile
├── manual_login()          # Prompt user for manual login
├── capture_cookies()       # Extract session cookies
├── detect_api_endpoint()   # Monitor network to find API
├── try_common_api_patterns() # Fallback API patterns
├── fetch_vehicle_data()    # Make API request per vehicle
├── parse_api_response()    # Extract expiry date from JSON
├── format_date()           # Convert to dd/mm/yyyy
└── process_excel()         # Orchestrate full workflow
```

## Troubleshooting

### Session Expired
```bash
# If you see SESSION_EXPIRED status:
# 1. Delete chrome_profile folder
# 2. Run scraper again
# 3. Login manually when prompted
rm -rf chrome_profile
python scraper.py
```

### API Not Detected
```bash
# The scraper will automatically try common patterns
# If still failing, check console logs for:
# - Network request URLs
# - API response samples
# Then update API patterns in try_common_api_patterns()
```

### No Expiry Date in Response
```bash
# If status is NOT_FOUND but vehicle exists:
# 1. Check console logs for API response structure
# 2. Add new date field keys to parse_api_response()
# 3. Update date_keys list with PolicyBoss-specific field names
```

## Security Notes

- Chrome profile stored locally in `./chrome_profile/`
- Session cookies captured but not logged
- No credentials stored in code
- API requests use same origin headers

## Limitations

- Requires manual login on first run
- API endpoint detection needs user interaction
- Rate limiting may apply (1-second delay between requests)
- PolicyBoss API structure may change

## Production Deployment

For production use:
1. Store chrome_profile in persistent storage
2. Implement automatic session refresh
3. Add webhook/email notifications
4. Scale with multi-threading for large Excel files
5. Add monitoring and alerting

## Support

If the scraper fails:
1. Check console logs for error messages
2. Verify PolicyBoss website structure hasn't changed
3. Test API endpoint manually with captured cookies
4. Update API patterns if PolicyBoss changed their backend

---

**Built with:** Python, Selenium, Requests, Pandas
**Target:** PolicyBoss Car Insurance (https://www.policyboss.com/car-insurance)
**Method:** Backend API extraction (no HTML scraping)
