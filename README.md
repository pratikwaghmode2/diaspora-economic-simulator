# 🇮🇳 Global Indian Diaspora: Economic Footprint & Remittance Simulation Engine (GID-EFRE)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Data: MEA Jan 2026](https://img.shields.io/badge/Data-MEA%20Govt%20of%20India%20(Jan%202026)-orange.svg)](https://www.mea.gov.in)
[![Benchmarks: RBI & World Bank](https://img.shields.io/badge/Benchmarks-RBI%20%26%20World%20Bank-emerald.svg)](https://www.rbi.org.in)
[![Deploy: GitHub Pages Ready](https://img.shields.io/badge/Deploy-GitHub%20Pages%20Ready-purple.svg)](https://pages.github.com)

> An open-source interactive macroeconomic simulator and geospatial explorer visualizing **37.28 Million Overseas Indians** across **205 territories**, calibrated against **$126.5+ Billion in annual inward remittances** to India and mapped to **Indian recipient states**.

---

## 🌟 Key Features

1. **Interactive Global Choropleth Map**:
   - Visualizes all 205 countries/territories by remittance volume, total diaspora size, or NRI proportion.
   - Click-to-inspect tooltips displaying regional classification, currency, and migration history.

2. **⚡ Real-Time Economic Stress-Testing Lab**:
   - **Multi-Currency Sensitivity Sliders**: Simulate currency fluctuations across **USD/INR**, **EUR/INR**, and **GBP/INR**.
   - **Workforce Shocks**: Model the effect of Gulf employment cycles (±20%) and Western high-skilled tech wage shifts.
   - **Windfall Calculator**: Computes the net purchasing power gain/loss (in ₹ Crores) flowing directly to Indian households.
   - **1-Click Share Button**: `📋 Copy My Simulation for LinkedIn` formats and copies your custom scenario to your clipboard in 1 click!

3. **🇮🇳 "Where the Money Lands in India" (State-Wise Footprint)**:
   - Calibrated against the official **Reserve Bank of India (RBI) Inward Remittance Survey**.
   - Dynamically calculates the exact ₹ Crores flowing into each Indian state (Maharashtra ~35.2%, Kerala ~10.2%, Tamil Nadu ~8.8%, Karnataka ~7.6%, Punjab ~5.4%, Gujarat ~5.8%, UP ~4.2%, etc.) as you adjust global sliders!
   - State inspector card mapping each Indian state to its primary international diaspora corridor.

4. **Multi-Currency Global Switcher**:
   - Toggle instantly between **USD ($)**, **EUR (€)**, **INR (₹ Lakh Crores / Crores)**, and **GBP (£)** across all metrics, charts, and tables.
   - **Live Foreign Exchange API**: Auto-fetches live market rates on page load with graceful offline fallback.

5. **🎓 Student & Researcher Analytical Spotlight**:
   - *Paradox 1: The Tale of Two Diasporas* — Why historical indentured settlements (Mauritius, Fiji, Trinidad, Guyana with 70%+ Indian populations) send minimal remittances, while Gulf nations (UAE, Saudi, Kuwait with 99% NRIs) send tens of billions.
   - *Macro Fact: The $126B Lifeline* — How remittances surpass India's total Foreign Direct Investment (FDI) and stabilize the Current Account Deficit (CAD).
   - *Trend: The Medical Student & Central European Corridor* — Exploring Indian population clusters in Armenia (7.5k), Georgia (600+), Kazakhstan (8k), Cyprus (14k), and Croatia (12k).
   - *Mechanics: The Rupee Windfall* — The mathematical relationship between currency depreciation and surges in NRE fixed deposits.

6. **Head-to-Head Country Comparison**:
   - Select any two nations (e.g. *USA vs. UAE* or *Canada vs. Australia*) for instant side-by-side metric ratios.

7. **🧮 Personal Remittance Impact Calculator**:
   - Select your earning country (UAE, USA, UK, Germany, Canada, Australia, Saudi Arabia, Singapore, Kuwait, Qatar, Oman).
   - Enter your monthly transfer amount or click quick presets ($500, $1,000, $2,500, $5,000).
   - Choose your recipient Indian State or Union Territory (all 36 entities supported).
   - Calculates monthly Rupee realization, annual family contribution, corridor comparison multiplier, and simulated **+5% Rupee Windfall bonus** flowing to your household!

8. **📄 Executive 1-Click Print / PDF Report**:
   - Built-in `@media print` CSS engine formats the dashboard into clean, publication-ready multi-page briefing papers.
   - Cleanly hides interactive sliders and UI toggles during print, preserving data tables, KPIs, and analytical spotlights.

9. **Searchable & Filterable Registry**:
   - Search by name or ISO-3 code.
   - Filter by region or diaspora profile (NRI-dominated >75%, PIO-dominated >75%, or balanced).
   - **1-Click Export to CSV and Excel**.

---

## 🚀 Quick Start (Run Locally)

This repository is **zero-dependency** for the frontend. You can run it locally in seconds:

### Option A: Double-Click
Simply open `index.html` in any modern web browser (Google Chrome, Microsoft Edge, Mozilla Firefox, Brave, Safari).

### Option B: Local Python Server
```bash
# Clone the repository
git clone https://github.com/<your-username>/diaspora-economic-simulator.git
cd diaspora-economic-simulator

# Launch local preview server (auto-opens browser at http://localhost:8000)
python scripts/preview_server.py
```

---

## 🌐 Deploy to GitHub Pages (1-Click Free Hosting)

You can host this live on the web for free using GitHub Pages:

1. Push this repository to your GitHub account:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: Global Indian Diaspora Economic Simulator"
   git branch -M main
   git remote add origin https://github.com/<your-username>/diaspora-economic-simulator.git
   git push -u origin main
   ```
2. On GitHub, navigate to **Settings** → **Pages**.
3. Under **Branch**, select `main` and root folder `/`, then click **Save**.
4. Your site will be live at: `https://<your-username>.github.io/diaspora-economic-simulator/`!

---

## 📁 Repository Structure

```
diaspora-economic-simulator/
├── index.html                   # Master single-page interactive dashboard (Tailwind + Plotly.js)
├── data/
│   ├── overseas_indians_2026.json # Enriched dataset (ISO3, currencies, remittances, DEPI scores)
│   ├── overseas_indians_2026.csv  # Clean CSV table
│   └── Population_of_Overseas_Indians_Jan_2026.xlsx # Master formatted Excel workbook
├── scripts/
│   ├── generate_data.py          # Calibration & enrichment pipeline (Python)
│   └── preview_server.py         # 1-click local HTTP server
├── LINKEDIN_POST.md              # Viral LinkedIn launch post template with hooks & insights
├── README.md                     # Project documentation & guide
└── LICENSE                       # MIT Open Source License
```

---

## 📊 Data Sources & Methodology

1. **Demographics**: Ministry of External Affairs (MEA), Government of India — *Population of Overseas Indians / Data on Indian Diaspora Abroad (as of January 2026)*.
2. **Macroeconomic Baselines**:
   - **Reserve Bank of India (RBI)**: *Inward Remittances Survey & Annual Report* (Corridor distribution: USA ~23.4%, UAE ~18.0%, UK ~6.8%, Singapore ~5.7%, Saudi Arabia ~5.1%, Eurozone, etc.).
   - **Reserve Bank of India (RBI)**: *State-wise Inward Remittance Survey* (Maharashtra ~35.2%, Kerala ~10.2%, Tamil Nadu ~8.8%, Karnataka ~7.6%, Gujarat ~5.8%, Punjab ~5.4%, UP ~4.2%).
   - **World Bank**: *Migration and Development Brief* (Total annual remittance baseline: $125B–$130B).
3. **Geospatial Standards**: United Nations Statistics Division (M49) & ISO 3166-1 alpha-3 codes.

---

## 📄 License

Distributed under the **MIT License**. See `LICENSE` for more information.
