# 📊 Stock Portfolio Tracker

A Python automation tool that fetches **live stock market data** using `yfinance` and generates a professional, multi-sheet **Excel report** — no manual data entry required.

Built as a finance + Python skills showcase for resume and GitHub.

---

## 🖼️ What It Produces

A formatted `.xlsx` workbook with 4 sheets:

| Sheet | Contents |
|---|---|
| 📊 Summary | KPI tiles (total value, gain/loss, day change) + allocation table + bar chart |
| 📋 Holdings | Detailed position table with Excel formulas (cost basis, P&L, return %) |
| 📈 Performance | 1D / 1W / 1M / 1Y returns, color-coded green/red |
| 🔍 Fundamentals | P/E ratio, dividend yield, beta, 52-week range |

---

## 🛠️ Tech Stack

- **Python 3.8+**
- [`yfinance`](https://github.com/ranaroussi/yfinance) — live price & fundamentals data
- [`pandas`](https://pandas.pydata.org/) — data manipulation
- [`openpyxl`](https://openpyxl.readthedocs.io/) — Excel generation, formatting, charts

---

## 🚀 Getting Started

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/stock-portfolio-tracker.git
cd stock-portfolio-tracker
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Edit your portfolio
Open `portfolio_tracker.py` and update the `PORTFOLIO` list at the top:
```python
PORTFOLIO = [
    ("AAPL",  10,  150.00),  # (ticker, shares, avg_cost_per_share)
    ("MSFT",   5,  280.00),
    ("NVDA",   6,  500.00),
    # Add your own positions...
]
```

### 4. Run it
```bash
python portfolio_tracker.py
```

This generates `portfolio_report.xlsx` in the same directory.

---

## 📁 Project Structure

```
stock-portfolio-tracker/
├── portfolio_tracker.py    # Main script
├── requirements.txt        # Python dependencies
├── README.md               # This file
└── portfolio_report.xlsx   # Sample output (simulated data)
```

---

## 💡 Key Finance Concepts Demonstrated

- **Cost Basis** — total capital deployed per position
- **Unrealized P&L** — mark-to-market gain/loss
- **Portfolio Weight** — % allocation per holding
- **1D/1W/1M/1Y Returns** — time-series performance
- **Beta** — market sensitivity measure
- **Dividend Yield** — income return on investment
- **P/E Ratio** — valuation multiple

---

## 🔧 Possible Extensions

- [ ] Add email delivery of the report (via `smtplib`)
- [ ] Schedule daily runs with `cron` or Task Scheduler
- [ ] Add a sector diversification analysis sheet
- [ ] Integrate options positions
- [ ] Add Monte Carlo simulation for portfolio risk

---

## 📄 License

MIT — free to use and modify.
