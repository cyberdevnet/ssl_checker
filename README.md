# Certificati SSL Checker

This Python script reads a list of URLs from an Excel file, checks whether the websites are online, extracts the SSL certificate expiration date, and any SSL verification errors. Results are written back to the Excel file.

---

## 🧾 Requirements

- Python 3.7+
- `openssl` command-line tool (must be available in your system path)
- Dependencies listed in `requirements.txt`

---

## 📦 Installation

1. **Clone or download this repo**

2. **Install dependencies**

pip install -r requirements.txt

# Ensure `openssl` is installed

`openssl` version

If it's not installed, on Linux/MacOS use your package manager (e.g., apt, brew). On Windows, install via https://slproweb.com/products/Win32`openssl`.html
📄 Input File

The Excel file must be named:
certificati_iso_date_with_status.xlsx

It must contain at least the following columns in the first row:

    URL – the website to check

    cert scadenza – where the cert expiration date will be written

    sito online – status ("yes"/"no") based on DNS and connection

    SSL error – SSL verify errors (if any)

Each URL should start with https:// or http://.
🚀 Usage

# Run the script:

python ssl_check.py

# This will:

    Connect to each URL in the list

    Check if it's online

    Extract SSL certificate expiration date

    Capture any SSL errors

    Save all results back into the same Excel file

🛠 Notes

    The script uses `openssl` s_client to extract certificate details.

    It has a timeout of 10 seconds per URL.

    DNS failures, SSL errors, and unreachable hosts are handled gracefully.
