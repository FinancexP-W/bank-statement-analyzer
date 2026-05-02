# 💰 Bank Statement Analyzer

**Automatically categorise your transactions and generate a financial dashboard — runs 100% on your computer. No internet. No subscriptions. No data uploaded anywhere.**

> Built for Indian bank statements (HDFC, ICICI, SBI, Axis, Kotak and more), works with any CSV-format statement worldwide.

---

## 📸 What it produces

| Sheet | What you get |
|-------|-------------|
| **Dashboard** | Dark-themed visual overview — category bars, monthly spend vs income, top 10 spends |
| **Summary** | KPI cards + full category breakdown for both spend and income |
| **Monthly Breakdown** | Every month × every category as a pivot table |
| **All Transactions** | Filterable list of every transaction with auto-assigned category |
| **Top 10 Spends** | Your 10 largest transactions across the entire period |
| **Extraordinary Transactions** | Any transaction > 30% of your average monthly spend |

---

## ⚡ Quick Start

### 1. Install Python
Download from [python.org](https://python.org/downloads) — **tick "Add Python to PATH"** during installation.

### 2. Install libraries
```bash
pip install pandas openpyxl
```

### 3. Download this repo
Click the green **Code** button → **Download ZIP** → extract to a folder on your PC.

### 4. Run
```bash
python bank_analyzer.py
```

The tool walks you through everything interactively. No config files to edit before you start.

---

## 🏦 Supported Banks (out of the box)

| Bank / Card | Format |
|-------------|--------|
| HDFC Bank | Separate Debit / Credit columns |
| ICICI Bank | Separate columns |
| SBI | Separate columns |
| Axis Bank | Separate columns |
| Kotak Bank | Separate columns |
| HDFC Credit Card | Single signed amount column |
| ICICI Credit Card | Single signed amount column |

**Any other bank?** Choose `NEW` when prompted and enter your column names — the tool remembers them for future runs.

---

## 🗂️ How it works

```
Your CSVs  →  bank_analyzer.py  →  Excel Dashboard
               │
               ├─ Reads category_keywords.csv  (your rules)
               ├─ Auto-detects date formats     (no manual config)
               ├─ Skips total/summary rows      (no inflated numbers)
               └─ Remembers your files          (session.json)
```

**First run:** add your CSV files one by one, pick the bank type, done.  
**Every run after:** press Enter — all previous files reload automatically. Just add the new month's file.

---

## 📂 Files in this repo

```
bank_analyzer.py          ← Main tool (run this)
category_keywords.csv     ← Category rules (edit this to customise)
README.md                 ← This file
```

`statement_configs.json` and `session.json` are created automatically on first run.

---

## 🏷️ Categories (auto-detected)

The tool matches transaction descriptions against keywords to assign categories:

| Category | Example keywords matched |
|----------|--------------------------|
| Food & Dining | swiggy, zomato, restaurant, blinkit, bigbasket |
| Shopping | amazon, flipkart, myntra, ajio, dmart |
| Transport | uber, ola, rapido, metro, petrol, fastag |
| Entertainment | netflix, hotstar, spotify, bookmyshow |
| Utilities | airtel, jio, bescom, gas bill, broadband |
| Health | apollo, practo, pharmacy, 1mg, hospital |
| EMI / Loan | emi, nach, loan repayment, bajaj finance |
| Investment | groww, zerodha, mutual fund, sip, ppf |
| Rent | rent, house rent, pg rent |
| Salary / Income | salary, credited, freelance, stipend |
| + 9 more | Education, Subscriptions, Travel, ATM, Charity… |

### Adding your own keywords

Open `category_keywords.csv` in Excel. Find the category row and add your keyword to the Keywords column, separated by a comma:

```
Food & Dining,  swiggy,zomato,...,my_local_restaurant_name
```

Matching is **case-insensitive** and **partial** — `apollo` matches `APOLLO PHARMACY KORAMANGALA 2ND FLOOR`.

---

## 💾 Session memory

After your first run, a `session.json` file is created. Next time you run:

```
  Last session: 10 file(s)
  ────────────────────────────────────────────────────
   1.  ✓  hdfc_jan_2024.csv         [HDFC Bank]
   2.  ✓  hdfc_feb_2024.csv         [HDFC Bank]
   3.  ✓  icici_cc_jan.csv          [ICICI Credit Card]
   ...
  ────────────────────────────────────────────────────
  Options:
    a        = load ALL (default)
    n        = start fresh
    1,3,5…   = load specific files only
```

No need to re-enter 10 file paths every month. Add only the new file.

---

## 🛡️ Privacy

- **All processing happens on your machine.** The script has no network calls.
- Your CSV files and the generated Excel file never leave your computer.
- No account, no login, no telemetry.

---

## 🐛 Common issues

| Problem | Fix |
|---------|-----|
| `'python' is not recognised` | Reinstall Python and tick **Add to PATH** |
| `Permission denied: '...\Bank statement'` | You typed a folder path — add the filename: `...\Bank statement\hdfc.csv` |
| `0 transactions parsed` | Column names don't match preset — choose `NEW` and enter exact column names shown |
| Months show as `Unknown` | Tool auto-detects dates; if it fails, choose `NEW` and set date format manually |
| `ModuleNotFoundError: pandas` | Run `pip install pandas openpyxl` |

---

## 📋 Requirements

- Python 3.8 or higher
- `pandas`
- `openpyxl`

```bash
pip install pandas openpyxl
```

---

## 🤝 Contributing

Found a bug? Bank format not working? Open an **Issue** and paste:
1. The column names from your CSV (first row)
2. A sample row (remove real amounts/names if you prefer)
3. The error message from the terminal

PRs welcome for new bank presets, keyword additions, or bug fixes.

---

## 📄 Licence

MIT — free to use, modify, and share. Credit appreciated but not required.

---

*Made with Python • Runs locally • Your data stays yours*
