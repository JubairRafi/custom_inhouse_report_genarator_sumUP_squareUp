# 📊 St George Sales Report Manager

Generate professional **Top 10 Sales Reports** for Kings Road and St Martin outlets with one click.

---

## ✅ Requirements

- Python 3.9 or later
- The following packages (see `requirements.txt`):

| Package | Purpose |
|---|---|
| `streamlit` | Web UI |
| `pandas` | Data processing |
| `openpyxl` | Excel formatting |

---

## 🚀 Quick Start

### 1 — Install dependencies

Open a terminal in this folder and run:

```powershell
pip install -r requirements.txt
```

### 2 — (Optional) Generate sample test data

```powershell
python create_sample_data.py
```

This creates `sample_kings_road.xlsx` and `sample_st_martin.xlsx` for testing.

### 3 — Launch the app

```powershell
streamlit run app.py
```

The app opens automatically at **http://localhost:8501**.

---

## 📥 Input File Format

Both outlet files must be Excel (`.xlsx` / `.xls`) with at least these columns:

| Column | Notes |
|---|---|
| `Date` | Any recognisable date format |
| `Item Name` | Product / menu item name |
| `Qty` | Units sold (numeric) |

Extra columns are ignored. Column names are matched case-insensitively.

---

## 📤 Output

A single `.xlsx` file with three sheets:

| Sheet | Content |
|---|---|
| **Combined** | Top 10 Best Seller St George (both outlets merged) |
| **Kings Road** | Top 10 Kings Road only |
| **St Martin** | Top 10 St Martin only |

Each sheet has:
- Merged, colour-coded title row
- Bold, coloured header row
- Bordered data cells with auto-fitted widths
- Bold grand-total summary row

---

## 📁 File Structure

```
weekly sales reports genarator manager/
│
├── app.py                  ← Streamlit UI (run this)
├── data_loader.py          ← File I/O & date filtering
├── report_engine.py        ← Aggregation & ranking logic
├── excel_exporter.py       ← Excel formatting (openpyxl)
├── create_sample_data.py   ← Test data generator
├── requirements.txt        ← Python dependencies
└── README.md               ← This file
```

---

## 🌐 Online Deployment

The easiest way to use this tool online is via **Streamlit Community Cloud**.

### 1 — Push to GitHub
1.  Create a new repository on [GitHub](https://github.com).
2.  Initialize git in your local folder:
    ```bash
    git init
    git add .
    git commit -m "Initial commit"
    ```
3.  Link to your GitHub repo and push:
    ```bash
    git remote add origin YOUR_GITHUB_REPO_URL
    git branch -M main
    git push -u origin main
    ```

### 2 — Deploy to Streamlit Cloud
1.  Go to [share.streamlit.io](https://share.streamlit.io).
2.  Click **"New app"**.
3.  Select your repository, branch (`main`), and main file path (`app.py`).
4.  Click **"Deploy!"**.

Your tool will be live at a public URL (e.g., `https://st-george-reports.streamlit.app`).
