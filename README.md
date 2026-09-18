# ST Dashboard — Supertails Delivery Performance

Professional dashboard that applies the **exact Excel SOP logic** for:
- Model classification (15 Min / 30 Min / SDD / Overnight / Warehouse)
- TAT calculation
- SLA Status
- Breach Delay & Final Verdict
- EOD Report
- Pendency Report
- Breach Report

**New in this version:**
- Manual **Rider Count** input for each store (sidebar)
- **Avg Order / Rider** calculated automatically
- New tab: **Breach Order IDs + Reason** (shows every breached order with CDR ID, delay, rider, Last Failed Remark, etc.)

Works with the 3 store CSVs from Zippee / Pidge:
1. Kalyan Nagar (Hyperlocal)
2. Varthur
3. Indranagar

---

## How to run on your laptop

1. Extract the ZIP
2. Open terminal in the `st_dashboard` folder
3. Install once:
   ```
   pip install -r requirements.txt
   ```
4. Start:
   ```
   streamlit run app.py
   ```

Browser opens at http://localhost:8501

---

## How to use

1. Choose **Daily Mode** (recommended) or All Data
2. Upload the 3 CSVs
3. Enter **Rider Count** for each store in the sidebar
4. Click **Generate Reports**
5. Check the new **Breach Order IDs + Reason** tab for full list of breached orders

---

Logic is identical to the Excel workbook + handwritten daily SOP.
