"""
PolicyBoss Car Insurance Scraper
Uses backend API calls to fetch car insurance expiry dates
"""

import time
import json
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import pandas as pd
from datetime import datetime
import os
import sys

class PolicyBossScraper:
    def __init__(self, chrome_profile_dir="./chrome_profile"):
        """Initialize scraper with persistent Chrome profile"""
        self.chrome_profile_dir = os.path.abspath(chrome_profile_dir)
        self.base_url = "https://www.policyboss.com"
        self.driver = None
        self.session = requests.Session()
        self.api_endpoint = None
        self.cookies = {}

    def setup_chrome(self):
        """Setup Chrome with persistent profile"""
        print(f"[INFO] Setting up Chrome with profile: {self.chrome_profile_dir}")

        chrome_options = Options()
        chrome_options.add_argument(f"--user-data-dir={self.chrome_profile_dir}")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        # Enable performance logging to capture network requests
        chrome_options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})

        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            print("[SUCCESS] Chrome launched successfully")
        except Exception as e:
            print(f"[ERROR] Failed to launch Chrome: {e}")
            sys.exit(1)

    def manual_login(self):
        """Prompt user to login manually"""
        print("\n" + "="*60)
        print("[ACTION REQUIRED] Please login to PolicyBoss manually")
        print("="*60)

        self.driver.get("https://www.policyboss.com/car-insurance")

        print("\nSteps:")
        print("1. Complete the login process in the browser")
        print("2. Navigate to car insurance page if needed")
        print("3. Press ENTER here when done...")

        input("\nPress ENTER after you've logged in: ")
        print("[INFO] Proceeding with logged-in session...")

    def capture_cookies(self):
        """Capture session cookies from browser"""
        print("[INFO] Capturing session cookies...")

        try:
            selenium_cookies = self.driver.get_cookies()

            for cookie in selenium_cookies:
                self.cookies[cookie['name']] = cookie['value']
                self.session.cookies.set(cookie['name'], cookie['value'])

            print(f"[SUCCESS] Captured {len(self.cookies)} cookies")

            # Print important cookies (masked)
            important_cookies = ['JSESSIONID', 'sessionid', 'token', 'auth']
            for key in important_cookies:
                if key in self.cookies:
                    masked_value = self.cookies[key][:8] + "..." if len(self.cookies[key]) > 8 else self.cookies[key]
                    print(f"  - {key}: {masked_value}")

        except Exception as e:
            print(f"[ERROR] Failed to capture cookies: {e}")

    def detect_api_endpoint(self, vehicle_number):
        """
        Detect the API endpoint by monitoring network requests
        User should perform a search action while this monitors
        """
        print(f"\n[INFO] Detecting API endpoint for vehicle: {vehicle_number}")
        print("[ACTION REQUIRED] Please search for a vehicle in the browser to help detect the API")

        # Navigate to car insurance page
        self.driver.get("https://www.policyboss.com/car-insurance")
        time.sleep(2)

        print(f"\nPlease enter vehicle number '{vehicle_number}' in the form and submit")
        print("Press ENTER here after submitting the form...")
        input()

        # Capture network logs
        print("[INFO] Analyzing network requests...")
        time.sleep(3)

        try:
            logs = self.driver.get_log('performance')

            api_candidates = []

            for entry in logs:
                try:
                    log_data = json.loads(entry['message'])
                    message = log_data.get('message', {})

                    if message.get('method') == 'Network.responseReceived':
                        response = message.get('params', {}).get('response', {})
                        url = response.get('url', '')

                        # Look for API endpoints related to vehicle/insurance/policy
                        if any(keyword in url.lower() for keyword in ['vehicle', 'insurance', 'policy', 'quote', 'api', 'search']):
                            if 'policyboss.com' in url and response.get('mimeType') == 'application/json':
                                api_candidates.append(url)

                except Exception as e:
                    continue

            if api_candidates:
                print(f"\n[SUCCESS] Found {len(api_candidates)} potential API endpoints:")
                for i, url in enumerate(set(api_candidates), 1):
                    print(f"  {i}. {url}")

                # Use the most relevant endpoint
                self.api_endpoint = api_candidates[0]
                print(f"\n[SELECTED] Using: {self.api_endpoint}")
                return True
            else:
                print("[WARNING] No API endpoints detected from network logs")
                print("[INFO] Will attempt common PolicyBoss API patterns...")
                return False

        except Exception as e:
            print(f"[ERROR] Failed to capture network logs: {e}")
            return False

    def try_common_api_patterns(self, vehicle_number):
        """
        Try common API endpoint patterns used by insurance websites
        """
        print(f"[INFO] Trying common API patterns for: {vehicle_number}")

        # Common PolicyBoss API patterns
        api_patterns = [
            f"https://www.policyboss.com/api/v1/vehicle/search?regNo={vehicle_number}",
            f"https://www.policyboss.com/api/vehicle/details?registration={vehicle_number}",
            f"https://api.policyboss.com/vehicle/search?regNo={vehicle_number}",
            f"https://www.policyboss.com/api/insurance/vehicle?regNo={vehicle_number}",
            f"https://www.policyboss.com/api/car-insurance/vehicle-details?regNo={vehicle_number}",
        ]

        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/plain, */*',
            'Referer': 'https://www.policyboss.com/car-insurance',
            'Origin': 'https://www.policyboss.com'
        }

        for pattern in api_patterns:
            try:
                print(f"[TRYING] {pattern}")
                response = self.session.get(pattern, headers=headers, timeout=10)

                if response.status_code == 200:
                    try:
                        data = response.json()
                        print(f"[SUCCESS] Got response: {json.dumps(data, indent=2)[:200]}...")
                        return data, pattern
                    except:
                        continue

            except Exception as e:
                continue

        return None, None

    def fetch_vehicle_data(self, vehicle_number, retry_count=3):
        """
        Fetch vehicle insurance data via API
        Returns: (expiry_date, status, insurer_name)
        """
        print(f"\n[PROCESSING] Vehicle: {vehicle_number}")

        for attempt in range(1, retry_count + 1):
            try:
                print(f"  [Attempt {attempt}/{retry_count}]")

                # If we have detected API endpoint, use it
                if self.api_endpoint:
                    # Construct API URL (replace vehicle number in the endpoint)
                    api_url = self.api_endpoint.replace(vehicle_number, vehicle_number)
                else:
                    # Try common patterns
                    data, endpoint = self.try_common_api_patterns(vehicle_number)
                    if data and endpoint:
                        self.api_endpoint = endpoint
                        return self.parse_api_response(data, vehicle_number)

                # Make API request
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    'Accept': 'application/json',
                    'Referer': 'https://www.policyboss.com/car-insurance',
                }

                response = self.session.get(self.api_endpoint, headers=headers, timeout=15)

                if response.status_code == 200:
                    data = response.json()
                    return self.parse_api_response(data, vehicle_number)
                elif response.status_code == 401 or response.status_code == 403:
                    print(f"  [ERROR] Session expired (Status: {response.status_code})")
                    return None, "SESSION_EXPIRED", None
                else:
                    print(f"  [ERROR] API returned status: {response.status_code}")

            except requests.exceptions.Timeout:
                print(f"  [ERROR] Request timeout")
                if attempt < retry_count:
                    time.sleep(2)
                    continue
            except Exception as e:
                print(f"  [ERROR] {str(e)}")
                if attempt < retry_count:
                    time.sleep(2)
                    continue

        return None, "ERROR", None

    def parse_api_response(self, data, vehicle_number):
        """
        Parse API response to extract expiry date and insurer
        Handles multiple response formats
        """
        print(f"  [PARSING] API Response structure: {list(data.keys()) if isinstance(data, dict) else type(data)}")

        expiry_date = None
        insurer_name = None

        # Try to find expiry date in common response formats
        # Format 1: Direct keys
        date_keys = ['expiryDate', 'expiry_date', 'policyExpiryDate', 'policy_expiry_date',
                     'insuranceExpiryDate', 'insurance_expiry_date', 'expiresOn', 'expires_on']

        insurer_keys = ['insurerName', 'insurer_name', 'insurer', 'companyName', 'company_name']

        def search_nested(obj, keys):
            """Recursively search for keys in nested structure"""
            if isinstance(obj, dict):
                for key in keys:
                    if key in obj:
                        return obj[key]
                for value in obj.values():
                    result = search_nested(value, keys)
                    if result:
                        return result
            elif isinstance(obj, list):
                for item in obj:
                    result = search_nested(item, keys)
                    if result:
                        return result
            return None

        expiry_date = search_nested(data, date_keys)
        insurer_name = search_nested(data, insurer_keys)

        if expiry_date:
            # Format the date to dd/mm/yyyy
            formatted_date = self.format_date(expiry_date)
            if formatted_date:
                print(f"  [SUCCESS] Expiry Date: {formatted_date}, Insurer: {insurer_name or 'N/A'}")
                return formatted_date, "FOUND", insurer_name

        print(f"  [NOT FOUND] No expiry date in API response")
        return None, "NOT_FOUND", None

    def format_date(self, date_str):
        """
        Convert various date formats to dd/mm/yyyy
        """
        if not date_str:
            return None

        date_formats = [
            '%Y-%m-%d',
            '%d-%m-%Y',
            '%Y/%m/%d',
            '%d/%m/%Y',
            '%Y-%m-%dT%H:%M:%S',
            '%Y-%m-%dT%H:%M:%S.%fZ',
            '%d %b %Y',
            '%d %B %Y',
        ]

        for fmt in date_formats:
            try:
                dt = datetime.strptime(str(date_str), fmt)
                return dt.strftime('%d/%m/%Y')
            except:
                continue

        # If no format matches, return as-is
        print(f"  [WARNING] Could not parse date format: {date_str}")
        return str(date_str)

    def process_excel(self, input_file='input.xlsx', output_file='output.xlsx'):
        """
        Process Excel file with vehicle numbers
        """
        print(f"\n[INFO] Processing Excel file: {input_file}")

        try:
            # Read input Excel
            df = pd.read_excel(input_file)

            if df.empty or df.columns[0] not in df:
                print("[ERROR] Excel file must have vehicle numbers in Column A")
                return

            vehicle_numbers = df.iloc[:, 0].tolist()
            print(f"[INFO] Found {len(vehicle_numbers)} vehicle numbers")

            # Prepare output data
            results = []

            # First vehicle - use for API detection if needed
            if vehicle_numbers and not self.api_endpoint:
                print("\n[SETUP] Need to detect API endpoint...")
                self.detect_api_endpoint(vehicle_numbers[0])

            # Process each vehicle
            for i, vehicle_number in enumerate(vehicle_numbers, 1):
                print(f"\n--- Processing {i}/{len(vehicle_numbers)} ---")

                if pd.isna(vehicle_number):
                    results.append({
                        'Vehicle Number': '',
                        'Expiry Date': '',
                        'Status': 'EMPTY'
                    })
                    continue

                vehicle_number = str(vehicle_number).strip()

                expiry_date, status, insurer = self.fetch_vehicle_data(vehicle_number)

                results.append({
                    'Vehicle Number': vehicle_number,
                    'Expiry Date': expiry_date or '',
                    'Status': status
                })

                # Small delay between requests
                time.sleep(1)

            # Create output DataFrame
            output_df = pd.DataFrame(results)
            output_df.to_excel(output_file, index=False)

            print(f"\n[SUCCESS] Results saved to: {output_file}")
            print(f"\nSummary:")
            print(f"  Total: {len(results)}")
            print(f"  Found: {sum(1 for r in results if r['Status'] == 'FOUND')}")
            print(f"  Not Found: {sum(1 for r in results if r['Status'] == 'NOT_FOUND')}")
            print(f"  Errors: {sum(1 for r in results if r['Status'] == 'ERROR')}")

        except FileNotFoundError:
            print(f"[ERROR] Input file not found: {input_file}")
        except Exception as e:
            print(f"[ERROR] Failed to process Excel: {e}")

    def close(self):
        """Cleanup resources"""
        if self.driver:
            print("\n[INFO] Closing browser...")
            self.driver.quit()

def main():
    """Main execution function"""
    print("="*60)
    print("PolicyBoss Car Insurance Scraper")
    print("API-based Vehicle Insurance Expiry Date Extractor")
    print("="*60)

    scraper = PolicyBossScraper()

    try:
        # Setup Chrome with persistent profile
        scraper.setup_chrome()

        # Check if already logged in
        scraper.driver.get("https://www.policyboss.com/car-insurance")
        time.sleep(3)

        # Prompt for manual login (first run or if session expired)
        print("\n[INFO] If you're not logged in, please complete login now")
        response = input("Are you logged in? (y/n): ").strip().lower()

        if response != 'y':
            scraper.manual_login()

        # Capture session cookies
        scraper.capture_cookies()

        # Process Excel file
        scraper.process_excel('input.xlsx', 'output.xlsx')

    except KeyboardInterrupt:
        print("\n[INFO] Interrupted by user")
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        scraper.close()
        print("\n[DONE] Scraper finished")

if __name__ == "__main__":
    main()
