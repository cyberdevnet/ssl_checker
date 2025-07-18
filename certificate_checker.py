
import openpyxl
from datetime import datetime
import subprocess
import re
from urllib.parse import urlparse
import socket

# Load workbook and sheet
workbook = openpyxl.load_workbook('certificati_iso_date_with_status.xlsx')
sheet = workbook['Sheet1']

############ Find column indices #################
header = [cell.value for cell in sheet[1]]

try:
    url_col_index = header.index("URL") + 1
    cert_col = header.index("cert scadenza") + 1
    online_col = header.index("sito online") + 1
    ssl_error_col = header.index("SSL error") + 1
except ValueError as e:
    raise ValueError(f"Missing expected column: {e}")

############ Extract all URLs #################
urls = []
for row in sheet.iter_rows(min_row=2, min_col=url_col_index, max_col=url_col_index, values_only=True):
    urls.append(row[0])

############ Cert verification and online check #################

def get_hostname(url):
    parsed = urlparse(url)
    return parsed.hostname

def get_cert_info(hostname, port=443):
    try:
        socket.gethostbyname(hostname)
        online_status = "yes"

        result = subprocess.run(
            ["openssl", "s_client", "-connect", f"{hostname}:{port}", "-servername", hostname, "-showcerts"],
            input="",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=10,
            text=True
        )

        output = result.stdout
        if result.returncode != 0 or not output:
            return "Connection failed", "no", ""

        # Parse expiry date
        match = re.search(r"NotAfter: (.+)", output)
        if match:
            expiry_str = match.group(1).strip()
            try:
                expiry_date = datetime.strptime(expiry_str, "%b %d %H:%M:%S %Y GMT").date().isoformat()
            except ValueError:
                expiry_date = f"Unparsable date: {expiry_str}"
        else:
            expiry_date = "NotAfter not found"

        # Extract all verify errors
        verify_errors = re.findall(r"verify error:[^\n]+", result.stderr)

        ssl_error_msg = "; ".join(verify_errors) if verify_errors else ""

        return expiry_date, online_status, ssl_error_msg

    except socket.gaierror:
        return "DNS resolution failed", "no", ""
    except subprocess.TimeoutExpired:
        return "Timeout", "no", ""
    except Exception as e:
        return f"Error: {str(e)}", "no", ""

############ Process each URL #################

for i, url in enumerate(urls, start=2):
    hostname = get_hostname(url)
    if hostname:
        expiry, online, ssl_error = get_cert_info(hostname)
    else:
        expiry, online, ssl_error = "Invalid URL", "no", ""

    sheet.cell(row=i, column=cert_col).value = expiry
    sheet.cell(row=i, column=online_col).value = online
    sheet.cell(row=i, column=ssl_error_col).value = ssl_error

    print(f"{hostname} => {expiry}, online: {online}, ssl error: {ssl_error or 'none'}")

# Save workbook
workbook.save("certificati_iso_date_with_status.xlsx")
