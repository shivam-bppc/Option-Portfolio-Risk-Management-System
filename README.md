# Option Portfolio Risk Management System

> A Python-based quantitative finance project implementing option pricing, Greeks calculation, volatility surface modeling, portfolio hedging, and Value-at-Risk analysis on NSE-listed stocks.

---

## Overview

This project builds a complete **end-to-end options risk management framework** using real market data from the NSE. Starting from raw stock price data, it prices a full options chain using the **Black-Scholes-Merton (BSM) model**, computes option sensitivities (Greeks), constructs a volatility surface, implements **delta and gamma hedging strategies**, and finally quantifies portfolio risk using **Value-at-Risk (VaR)**.

Built as part of the Derivatives and Risk Management course at **BITS Pilani, Pilani Campus**.

---

## Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.x | Core language |
| `yfinance` | Real-time & historical NSE stock data |
| `numpy` / `scipy` | Mathematical computations, BSM model |
| `pandas` | Data manipulation and tabulation |
| `matplotlib` / `plotly` | Volatility surface and PnL visualizations |
| Excel (`.xlsx`) | Output tables for Greeks, VaR, and hedging |

---

## Project Structure

```
option-portfolio-risk/
│
├── data/
│   └── stock_prices.csv          # Raw downloaded price data
│
├── src/
│   ├── data_collection.py        # Stock data extraction via yfinance
│   ├── bsm_pricing.py            # Black-Scholes-Merton pricing engine
│   ├── greeks.py                 # Delta, Gamma, Vega, Theta, Rho
│   ├── implied_volatility.py     # IV calculation from NSE option chain
│   ├── volatility_surface.py     # 3D vol surface construction
│   ├── portfolio.py              # Portfolio construction & hedging
│   └── var.py                    # VaR: parametric & historical simulation
│
├── outputs/
│   └── results.xlsx              # Full results: pricing, Greeks, VaR tables
│
├── main.py                       # Run full pipeline end-to-end
├── requirements.txt
└── README.md
```

---

## Methodology

### Part A — Data Collection & Return Statistics
- Extracted **3 months of daily closing prices** for a NIFTY 200 stock via `yfinance`
- Computed:
  - Daily log returns
  - Annualized volatility (`√252 × daily std`)
  - Skewness and Kurtosis of return distribution

---

### Part B — Option Pricing (BSM Model)
- Set **5 strike prices**: ATM, ATM ± 2%, ATM ± 5%
- Set **3 maturities**: 30, 60, 90 days
- Priced **call and put options** for each strike × maturity combination → **30 option prices total**

BSM Formula used:

```
C = S·N(d1) - K·e^(-rT)·N(d2)
P = K·e^(-rT)·N(-d2) - S·N(-d1)

where:
  d1 = [ln(S/K) + (r + σ²/2)·T] / (σ·√T)
  d2 = d1 - σ·√T
```

---

### Part C — Greeks & Implied Volatility
Calculated for each option using **both historical volatility and IV**:

| Greek | Measures |
|---|---|
| Delta (Δ) | Sensitivity to stock price change |
| Gamma (Γ) | Rate of change of Delta |
| Vega (ν) | Sensitivity to volatility change |
| Theta (Θ) | Time decay of option value |
| Rho (ρ) | Sensitivity to interest rate change |

- Downloaded **live NSE option chain data** to back-calculate **Implied Volatility (IV)** via Newton-Raphson root-finding
- Constructed a **3D Volatility Surface** (Strike × Maturity × IV) to visualize the volatility smile/skew

---

### Part D — Portfolio Construction & Hedging

Constructed a sample options portfolio and computed:
- Portfolio **Delta, Gamma, Vega**
- **Delta Hedge**: Used underlying stock positions to neutralize portfolio delta
- **Gamma Hedge**: Used additional options/futures to neutralize gamma
- Simulated **PnL under ±1% and ±2% stock price moves** to validate hedging effectiveness

Hedged and unhedged PnL values in the report

*(See `outputs/results.xlsx` for actual values)*

---

### Part E — Value-at-Risk (VaR)

Computed **1-day VaR** at both **95% and 99% confidence levels** using two methods:

| Method | Description |
|---|---|
| Parametric (Variance-Covariance) | Assumes normal distribution of returns |
| Historical Simulation | Uses last 60 days of actual portfolio returns |

Compared **hedged vs unhedged** portfolio VaR to demonstrate risk reduction from hedging.

---

##  Sample Outputs

<img width="987" height="491" alt="Screenshot 2026-02-28 at 10 23 23" src="https://github.com/user-attachments/assets/b3677e45-7d51-49d0-b071-89bbc12fe37d" />

<img width="753" height="613" alt="Screenshot 2026-02-28 at 10 24 49" src="https://github.com/user-attachments/assets/da406da0-c35d-492e-852c-82e0f909ebb5" />


### Volatility Surface
> *(Add a screenshot of your 3D volatility surface plot here)*

### Greeks Table (Sample)
| Strike | Maturity | Delta | Gamma | Vega | Theta | Rho |
|---|---|---|---|---|---|---|
| ATM | 30d | 0.52 | 0.03 | 0.18 | -0.04 | 0.07 |
| ATM+5% | 60d | 0.38 | 0.02 | 0.21 | -0.03 | 0.06 |
| ATM-5% | 90d | 0.64 | 0.02 | 0.19 | -0.02 | 0.09 |

*(Full table in `outputs/results.xlsx`)*

---

## 🚀 How to Run

```bash
# 1. Clone the repo
git clone https://github.com/shivam-bppc/option-portfolio-risk.git
cd option-portfolio-risk

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the full pipeline
python main.py
```

---

## 📦 Requirements

```
yfinance
numpy
scipy
pandas
matplotlib
plotly
openpyxl
```

---

## 📚 References

- Black, F., & Scholes, M. (1973). *The Pricing of Options and Corporate Liabilities*
- Hull, J. C. — *Options, Futures, and Other Derivatives*
- NSE India Option Chain Data — [nseindia.com](https://www.nseindia.com)

---

## 👤 Author

**Shivam Singla**  
B.E. Computer Science Engineering, BITS Pilani (2027)  
📧 F20230576@pilani.bits-pilani.ac.in  
🔗 [LinkedIn](https://www.linkedin.com/in/shivam-singla-043ab53a8/) · [GitHub](https://github.com/shivam-bppc)

---

> *Built for academic purposes as part of ECON/FIN coursework at BITS Pilani. Not financial advice.*
