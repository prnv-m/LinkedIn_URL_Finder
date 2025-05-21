# LinkedIn URL Finder

This project allows you to clean a CSV file with lead information (names and emails), split full names into first and last names, and then find corresponding LinkedIn URLs based on the cleaned data.

---

## 🧰 Prerequisites

Before you begin, make sure:

- You have **Python 3.12** installed.
- You have `pip` available for Python 3.12.

---

## 📦 Install Required Libraries

Open your terminal and run the following command:

```
python3.12 -m pip install pandas stealth_requests beautifulsoup4
```
# Clean the CSV/XLSX
## Run the cleaning script to process your raw leads after changing input file to desired input file
```
python clean_csv_leadfile.py 
```
# Find LinkedIn URLs
## Run the cleaning script to process your raw leads after changing input file to desired input file and output files

```
python find_linkedin_fromcsv.py
```




