# 💰 Bank Statement Analyzer

**Automatically categorise your transactions and generate a financial dashboard — runs 100% on your computer. No internet. No subscriptions. No data uploaded anywhere.**

> Built for Indian bank statements (HDFC, ICICI, SBI, Axis, Kotak and more), works with any CSV-format statement worldwide.

---

## 📸 What it produces

A 7-sheet Excel dashboard, generated entirely on your machine:

| Sheet | What you get |
|-------|-------------|
| **Summary** | KPI cards (total spend, income, savings rate, avg monthly) + full category breakdowns for spend and income. Contra entries excluded from all totals. |
| **Monthly Breakdown** | Every month × every spend category as a pivot table. Contra excluded. |
| **All Transactions** | Filterable list of every transaction with date, description, category, amount. Contra rows highlighted in purple italic. |
| **Top 10 Spends** | Your 10 largest real transactions. Contra entries never appear here. |
| **Extraordinary Transactions** | Any transaction > 30% of your average monthly spend — catches unusual activity. |
| **Contra Entries** | Dedicated sheet for internal transfers, CC bill payments, inter-account moves. Shows balance check, category summary, and full list. |
| **Dashboard** | Dark-themed visual overview — category bars, monthly spend vs income, net savings trend, top 10 spend table. |

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

The tool walks you through everything interactively — no config files to edit before you start.

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

**Any other bank?** Choose `NEW` when prompted and enter your column names — the tool saves them for future runs.

---

## 🗂️ How it works

```
Your CSVs  →  bank_analyzer.py  →  7-sheet Excel Dashboard
               │
               ├─ Reads category_keywords.csv    (your keyword rules)
               ├─ Reads category_overrides.csv   (per-transaction fixes)
               ├─ Auto-detects date formats       (no manual config needed)
               ├─ Skips total/balance/footer rows (no inflated numbers)
               ├─ Excludes Contra entries         (clean income & spend totals)
               └─ Remembers your files            (session.json)
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

These are created automatically on first run — **do not upload them to GitHub** (add to `.gitignore`):

```
category_overrides.csv    ← Your per-transaction category fixes
statement_configs.json    ← Your saved bank column settings
session.json              ← Remembered file list from last run
```

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

Open `category_keywords.csv` in Excel, find the category row, and add your keyword to the Keywords column separated by a comma:

```
Food & Dining,  swiggy,zomato,...,my_local_restaurant_name
```

Matching is **case-insensitive** and **partial** — `apollo` matches `APOLLO PHARMACY KORAMANGALA 2ND FLOOR`.

---

## 🔄 Contra Entries — tracking internal transfers

When you transfer money between your own accounts or pay a credit card bill, the same amount appears as both a debit and a credit across your statements. Without handling this, your total spend and income are both inflated by the same amount.

**How to tag entries as Contra:** in `category_keywords.csv`, name any category starting with the word `Contra` and add the relevant keywords:

```
Category                      Keywords
------------------------------------------------------------
Contra - CC Payment           hdfc credit card payment,icici cc bill,bob credit
Contra - Transfer Axis A/c    imps to axis,neft axis savings,transfer axis
Contra - Investment           groww imps,zerodha imps,coin purchase
```

The naming is flexible — `Contra - CC Payment`, `Contra: CC Payment`, and `Contra CC Payment` all work. The word `Contra` at the start is what matters.

**What happens to Contra entries:**

| Where | Behaviour |
|-------|-----------|
| Summary, Monthly Breakdown, Dashboard | Fully excluded — totals reflect real income and spend only |
| Top 10 Spends | Never appears |
| All Transactions | Still visible, highlighted in purple italic |
| Contra Entries sheet | Dedicated view with outflow total, inflow total, balance check (✅ Balanced / ⚠ Imbalance if one side of a transfer is missing), category summary, and full list |

---

## 📅 Credit card billing cycle grouping

By default every transaction is grouped into its calendar month. For credit cards whose billing cycle doesn't align with the calendar (e.g. 21st–20th), you can group by billing period instead.

When adding a credit card file the tool asks:

```
  1 = By transaction date  (default)
  2 = By billing cycle
  Choose grouping mode [1]: 2

  Billing cycle start day (1-28) [21]: 21

  ✓ Confirmed: a transaction on 25 Mar (cycle start 21st) → Apr 2024
```

**The logic:** whichever calendar month has more days in the cycle is the billing month. A 21 Mar–20 Apr cycle has 11 days in March and 20 in April → all those transactions are grouped under **April**.

The setting is saved per file — each credit card can have its own cycle. Files with billing cycle mode enabled show `[cycle:21]` in the session display.

---

## ✏️ Category overrides

Sometimes a specific transaction is categorised wrongly and adding a keyword to the CSV would cause false matches elsewhere. `category_overrides.csv` lets you fix individual transactions by their exact description — no keyword changes needed.

```
Description                              Category
-------------------------------------------------------
HDFC CREDIT CARD BILL PAYMENT            Contra - CC Payment
MARUTI INSURANCE ANNUAL PREMIUM          Insurance
UPI/TRANSFER/123456/SAVINGS              Contra - Transfer Axis A/c
```

- Matching is **case-insensitive** and **partial** (description contains the key)
- Overrides take **priority over keyword rules** — applied first, every run
- Applied consistently across **every sheet** — Summary, Monthly Breakdown, Dashboard, Contra Entries all reflect the corrected category

The All Transactions sheet has a reminder banner at the top pointing you to this file whenever you spot a wrong category.

---

## 💾 Session memory

After your first run, a `session.json` file is created. Next time you run:

```
  Last session: 10 file(s)
  ────────────────────────────────────────────────────────
   1.  ✓  hdfc_jan_2024.csv              [HDFC Bank]
   2.  ✓  hdfc_feb_2024.csv              [HDFC Bank]
   3.  ✓  icici_cc_jan.csv               [ICICI Credit Card]  [cycle:21]
   4.  ✗ NOT FOUND  old_statement.csv    [HDFC Bank]
  ────────────────────────────────────────────────────────
  Options:
    a        = load ALL found files  (default)
    n        = start fresh
    1,3,5…   = load specific files only
```

No need to re-enter paths every month — just add the new file when prompted. Files marked ✗ NOT FOUND were moved or renamed and are skipped automatically.

---

## 🛡️ Privacy

- **All processing happens on your machine.** The script has zero network calls.
- Your CSV files and generated Excel file never leave your computer.
- No account, no login, no telemetry.

---

## 🐛 Common issues

| Problem | Fix |
|---------|-----|
| `'python' is not recognised` | Reinstall Python and tick **Add Python to PATH** |
| `Permission denied: '...\Bank statement'` | You typed a folder path — add the filename: `...\Bank statement\hdfc.csv` |
| `0 transactions parsed` | Column names don't match preset — choose `NEW` and enter the exact column names shown |
| Months show as `Unknown` | Date format not detected — choose `NEW` and set date format manually (e.g. `%d/%m/%Y`) |
| Savings rate shows `4918%` | You are running an old version — replace `bank_analyzer.py` with the latest and re-run |
| Contra entries appear in Top 10 | Same — replace `bank_analyzer.py` with the latest version |
| Contra Entries sheet shows ⚠ Imbalance | One side of a transfer is missing — make sure both the sending and receiving bank CSVs are loaded |
| Billing cycle option not shown | File configured as a bank type, not credit card — re-add it and choose the correct credit card type |
| Override not working | Check spelling in `category_overrides.csv` — open in Notepad and compare against the exact description in All Transactions |
| `ModuleNotFoundError: pandas` | Run `pip install pandas openpyxl` |
| Excel file won't open / corrupted | File was open in Excel when the tool ran — close Excel, delete the old file, re-run |

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
