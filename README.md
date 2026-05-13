# ⚡ MPPCL Scheduling & BD Generator

A desktop application for **Madhya Pradesh Power Company Limited (MPPCL)** power scheduling engineers. It automates the daily process of generating Scheduling and Bid-Dispatch (BD) Excel workbooks from raw input files — replacing hours of manual work with a single click.

---

## 📋 Overview

Every day, load dispatch engineers must compile power entitlements from dozens of generating stations, merge zone-wise demand forecasts, compute merit-order dispatch (MOD), and prepare IEX bid sheets. This tool automates the entire pipeline:

1. Reads **Day-Ahead Entitlement** data from the MPPCL Excel file
2. Reads **zone-wise demand** (Central Zone, Eastern Zone, Western Zone)
3. Reads **MOD rate** data for merit-order ranking
4. Optionally parses an **IEX PDF** for market transaction data
5. Outputs a fully formatted **Excel workbook** with all required sheets

---

## 🗂️ Repository Structure

```
.
├── main.py                              # Main application (UI + all logic)
├── requirements.txt                     # Python dependencies
├── sample_data/
│   ├── Day_Ahead_Entitlement_of_MPPCL.xlsx   # Sample entitlement file
│   ├── Demand-CZ.xlsx                         # Central Zone demand sample
│   ├── Demand-EZ.xlsx                         # Eastern Zone demand sample
│   ├── Demand-WZ.xlsx                         # Western Zone demand sample
│   └── MOD_Rate.xlsx                          # Merit-order dispatch rate sample
└── README.md
```

---

## 🚀 Getting Started

### Prerequisites

- Python **3.10+**
- `pip` package manager
- Windows / macOS / Linux with a display (Tkinter required)

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/<Duvvada-Naveen-Kumar>/mppcl-scheduling.git
cd mppcl-scheduling

# 2. (Recommended) Create a virtual environment
python -m venv venv
source venv/bin/activate        # Linux / macOS
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

### Running the App

```bash
python main.py
```

A GUI window will open. Browse and select your input files, choose an output path, then click **Generate Scheduling + BD**.

---

## 📥 Input Files

| File | Description | Required |
|------|-------------|----------|
| `Day_Ahead_Entitlement_of_MPPCL.xlsx` | Day-ahead entitlement from MPPCL/SLDC; must contain sheets: `ISGS`, `ISGS NTPC RAJ SOLAR`, `REMC & IPP`, `SSGS` | ✅ Yes |
| `Demand-CZ.xlsx` | Central Zone restricted demand (96 time blocks) | At least 2 of the 3 zones |
| `Demand-EZ.xlsx` | Eastern Zone restricted demand (96 time blocks) | At least 2 of the 3 zones |
| `Demand-WZ.xlsx` | Western Zone restricted demand (96 time blocks) | At least 2 of the 3 zones |
| `MOD_Rate.xlsx` | Merit-order dispatch rates by station | ✅ Yes |
| IEX PDF | Daily Trade Report PDF from IEX (optional) | ❌ Optional |

### Demand File Format

Demand files should have a header row containing columns: `Time Block`, `From`, `To`, and `Restricted` demand (in MW). Blocks should be numbered 1–96 (each = 15 minutes).

### Entitlement File Format

Must follow the MPPCL Day-Ahead Entitlement template with sheets `ISGS`, `ISGS NTPC RAJ SOLAR`, `REMC & IPP`, and `SSGS`. The tool auto-detects the `TIME BLOCK` header and reads entitlement values per station per block.

---

## 📤 Output Workbook Sheets

| Sheet | Description |
|-------|-------------|
| `Scheduling` | Final station-wise 96-block schedule |
| `BD` | Bid-Dispatch summary |
| `Hydel ReSchedule` | Hydel station rescheduling data |
| `Bid Sheet` | IEX bid sheet with price-quantity pairs |
| `Name corrector` | Station name normalization & MOD rates |
| `Deamnd-MP` | Merged MP demand across zones |
| `ENTT-2` | Full entitlement matrix (96 blocks × all stations) |
| `RE` | Renewable energy data |
| `Hydel` | Raw hydel station data |
| `Entt` | Processed entitlement data |
| `IEX` | IEX market data (only if PDF provided) |
| `PowerBI` | Summary KPIs and fuel-mix table for dashboards |

---

## 🏭 Supported Generating Stations

The tool maps **100+ generating stations** across categories:

- **NTPC Thermal** – Singrauli, Rihand I/II/III, SIPAT I/II, VSTPS I–V, Gadarwara, Lara, Khargone, etc.
- **NTPC Gas** – Gandhar, KAWAS, Anta, Auraiya, Dadri (APM / NAPM / Liquid / LNG / CLNG)
- **NTPC Hydro** – Tehri, NJPC, Koldam, Rampur, Kishanganga, Dulhasti, URI II, etc.
- **State Thermal (SSGS)** – AMK III, SGT I/II, STP-IV, SSTPP I/II
- **State Hydro** – Gandhi Sagar, Pench, Bargi, Rajghat, Birsinghpur, Madikheda, etc.
- **IPP / Medium Term** – MB Power, Jhabua, JP Bina, JP Nigrie, RKM Power, SKS Raigarh, JIPL, etc.
- **Renewables (REMC)** – Wind, Solar, Small Hydro, Biogas, Biomass, MSW, and RUMS projects

---

## ⚙️ Key Configuration (in `main.py`)

| Constant | Purpose |
|----------|---------|
| `NAME_MAP` | Maps raw station names (from entitlement file) to standard internal codes |
| `RANKED_STATIONS` | Merit-order ranking list for scheduling priority |
| `HYDEL_CODES` | Station codes classified as hydel for special handling |
| `SSGS_CODES` | State Genco station codes |
| `BID_PRICE_POINTS` | Price levels (₹/MWh) used in the IEX Bid Sheet |
| `BID_MARGIN` | Bid volume margin factor (default: 0.15) |

---

## 🔧 Dependencies

```
openpyxl>=3.1
pdfplumber>=0.10
```

See `requirements.txt` for pinned versions.

> **Note:** `pdfplumber` is only needed if you use the optional IEX PDF parsing feature. The app works without it if no PDF is provided.

---

## 🖥️ GUI Screenshot

The application presents a simple file-picker interface:

```
┌──────────────────────────────────────────────┐
│  Input Files                                 │
│  Demand-CZ.xlsx    [path...]      [Browse]   │
│  Demand-EZ.xlsx    [path...]      [Browse]   │
│  Demand-WZ.xlsx    [path...]      [Browse]   │
│  IEX.pdf           [path...]      [Browse]   │
│  Entitlement.xlsx  [path...]      [Browse]   │
│  MoD_Rate.xlsx     [path...]      [Browse]   │
│                                              │
│  Output                                      │
│  Output file       [path...]      [Browse]   │
│                                              │
│       [ Generate Scheduling + BD ]           │
│  Ready                                       │
└──────────────────────────────────────────────┘
```

---

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you'd like to change.

---

## 📄 License

This project is for internal MPPCL operations. Please check with your organization before distributing.
