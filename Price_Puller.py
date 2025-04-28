import requests
import pandas as pd
from datetime import datetime, date
import os
import logging
import sys
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from urllib3.util import Retry
from requests.adapters import HTTPAdapter

# Setup logging
logging.basicConfig(
    filename='price_puller.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def load_urls(file_path):
    urls = {}
    try:
        with open(file_path, 'r') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or ',' not in line:
                    continue
                sheet_name, url = line.split(',', 1)
                urls[sheet_name.strip()] = url.strip()
        logging.info(f"✅ Loaded {len(urls)} URLs from {file_path}")
    except Exception as e:
        logging.exception(f"🚨 Failed to load URLs: {e}")
    return urls

def print_progress_bar(iteration, total, prefix='', suffix='', length=50, fill='█', start_time=None):
    percent = f"{100 * (iteration / float(total)):.1f}"
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)

    if start_time and iteration > 0:
        elapsed = time.time() - start_time
        eta_seconds = (elapsed / iteration) * (total - iteration)
        eta_formatted = time.strftime("%M:%S", time.gmtime(eta_seconds))
        eta_display = f" | ETA: {eta_formatted}"
    else:
        eta_display = ""

    sys.stdout.write(f'\r{prefix} |{bar}| {percent}% {suffix}{eta_display}')
    sys.stdout.flush()

def get_avg_price(url, sheet_name, output_file='Price_Puller.xlsx'):
    headers = {
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9",
        "content-type": "application/json",
        "priority": "u=1, i",
        "x-fwd-svc": "atc"
    }

    session = requests.Session()
    retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))

    try:
        response = session.get(url, headers=headers)
        response.raise_for_status()

        data = response.json()
        links = data.get('links', [])

        today_date = date.today()
        today_str = today_date.strftime('%d%b%Y').upper()

        row_data = {'Date': today_str}
        price_columns = []

        for link in links:
            year = link.get('value')
            avg_price = link.get('avgPrice')
            if year and avg_price:
                year = int(year)
                row_data[year] = avg_price
                price_columns.append(year)

        # Calculate row number based on anchor date
        anchor_date = datetime.strptime("22Apr2025", "%d%b%Y").date()
        delta_days = (today_date - anchor_date).days + 2

        # Load or create workbook
        if os.path.exists(output_file):
            wb = load_workbook(output_file)
        else:
            wb = Workbook()

        # Load or create sheet
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)

        # Setup headers
        existing_headers = [cell.value for cell in ws[1] if cell.value is not None]
        if not existing_headers:
            all_headers = ['Date'] + sorted(price_columns)
            for idx, header in enumerate(all_headers, start=1):
                cell = ws.cell(row=1, column=idx)
                cell.value = header
                cell.font = Font(bold=True)
        else:
            # Add new headers if necessary
            missing_headers = [year for year in price_columns if year not in existing_headers]
            if missing_headers:
                for year in sorted(missing_headers):
                    new_col = ws.max_column + 1
                    ws.cell(row=1, column=new_col).value = year
                    ws.cell(row=1, column=new_col).font = Font(bold=True)

            # Update headers and map
            all_headers = [cell.value for cell in ws[1] if cell.value is not None]

        header_map = {header: idx + 1 for idx, header in enumerate(all_headers)}

        # Write data
        ws.cell(row=delta_days, column=header_map['Date']).value = today_str

        for year in price_columns:
            col_index = header_map.get(year)
            if col_index:
                value = row_data[year]
                cell = ws.cell(row=delta_days, column=col_index)
                cell.value = value
                if isinstance(value, (int, float)):
                    cell.number_format = '"$"#,##0.00'

        wb.save(output_file)
        logging.info(f"✅ Successfully updated sheet '{sheet_name}' for {today_str}")

    except Exception as e:
        logging.exception(f"🚨 Error processing {sheet_name}: {e}")

def main():
    url_file = 'car_urls.txt'
    urls = load_urls(url_file)

    if not urls:
        logging.error("🚨 No URLs to process. Exiting.")
        return

    total = len(urls)
    completed = 0
    start_time = time.time()

    print("Starting price pulling...\n")
    print_progress_bar(0, total, prefix='Progress', suffix='Complete', length=50, start_time=start_time)

    for sheet_name, url in urls.items():
        get_avg_price(url, sheet_name)
        completed += 1
        print_progress_bar(completed, total, prefix='Progress', suffix='Complete', length=50, start_time=start_time)

    print("\n✅ All done!")

if __name__ == '__main__':
    main()