# ITAM - IT Asset Management Dashboard System

<div align="center">

![Version](https://img.shields.io/badge/version-2.4.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.8+-green.svg)
![Streamlit](https://img.shields.io/badge/streamlit-1.28+-red.svg)
![License](https://img.shields.io/badge/license-MIT-yellow.svg)

**Professional Asset Tracking & Analytics Platform**

[Features](#-features) • [Installation](#-installation) • [Quick Start](#-quick-start) • [Usage](#-usage) • [Contact](#-contact)

</div>

---

## 📋 Overview

**ITAM (IT Asset Management Dashboard System)** adalah platform web-based yang komprehensif untuk pengurusan aset IT organisasi. Sistem ini dibangunkan menggunakan Python Streamlit untuk menyediakan antara muka yang intuitif dan powerful analytics untuk menguruskan inventori aset IT.

### Objektif Sistem

- **Centralized Management** - Satu platform untuk menguruskan semua aset IT (Workstation & Mobile)
- **Real-time Analytics** - Dashboard interaktif dengan visualisasi data yang bermakna
- **Warranty Tracking** - Automated monitoring dan alerts untuk warranty expiry
- **Lifecycle Planning** - Asset age analysis untuk replacement planning
- **Data Validation** - Auto-detect duplicates, missing data, dan anomalies
- **Flexible Reporting** - Multi-format exports untuk procurement dan audit purposes

---

## ✨ Features

### 🎯 Core Capabilities

**Smart Detection**
- Auto-detect asset type (Workstation/Mobile) dari column names
- Intelligent header row detection dalam Excel files
- Flexible column matching walaupun ada typo atau format berbeza

**Comprehensive Dashboard**
- Summary metrics: Total assets, active/expired breakdown, replacement rate
- Asset type statistics dan regional distribution
- Interactive visual analytics (pie charts, bar charts)
- Real-time filtering dan search capabilities

**Warranty Management** (Workstation)
- Three-tier status: Active, Expiring Soon (90 days), Expired
- Automated expiry calculations
- Exportable lists untuk procurement planning

**Asset Lifecycle Analysis**
- Age categorization: New (0-1yr), Active (1-3yr), Aging (3-5yr), Old (5+yr)
- Average age calculations
- Replacement planning tools

**Data Quality Assurance**
- Duplicate detection (Asset tags, Serial numbers)
- Missing data identification
- Email format validation
- Severity-based prioritization (High, Medium, Low)

**Advanced Filtering**
- Multi-level filters: Model, Type, Site, Location, Department, Status
- Global text search across all fields
- Session state management

**Export & Reporting**
- Excel export untuk filtered data
- Segmented exports: All, Expired assets, Expired warranties
- Auto-generated filenames dengan timestamps
- Sample template downloads

**Professional UI/UX**
- Modern responsive design
- Bilingual support (English & Bahasa Malaysia)
- Interactive hover effects dan animations
- Mobile-friendly interface

---

## 💻 System Requirements

- **Python**: 3.8 or higher
- **Operating System**: Windows, macOS, Linux
- **RAM**: 4GB minimum (8GB recommended)
- **Browser**: Chrome, Firefox, Safari, Edge (latest versions)

---

## 🚀 Installation

### 1. Clone Repository
```bash
git clone https://github.com/mdk869/ITAM.git
cd ITAM
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run Application
```bash
streamlit run asset_dashboard.py
```

### 4. Access Dashboard
Open browser dan navigate ke `http://localhost:8501`

---

## ⚡ Quick Start

1. **Launch** aplikasi menggunakan command di atas
2. **Select Language** - Choose English atau Bahasa Malaysia
3. **Upload Excel File** - Klik "Upload Excel File (.xlsx)" button
4. **Explore Dashboard** - View metrics, apply filters, analyze data
5. **Export Reports** - Download filtered data dalam Excel format

### Sample Templates

Download sample Excel templates untuk reference:
- **Workstation Template** - Includes all standard columns untuk PC/Laptop assets
- **Mobile Template** - Template untuk phones dan tablets

Kedua-dua templates tersedia dalam aplikasi.

---

## 📖 Usage

### Supported Asset Types

**1. Workstation Assets (PC/Laptop)**
- Required: `Model`, `Asset Tag`, `Serial Number`, `User`
- Optional: `Workstation Type`, `Warranty Expiry`, `Year Of Purchase`, `Place`, `Location`, `Department`, `State`, `User Email`

**2. Mobile Assets (Phone/Tablet)**
- Required: `Product`, `Asset Tag`, `Serial Number`, `User`
- Optional: `Product Type`, `Programme`, `Site`, `Year Of Purchase`, `Location`, `Department`, `State`, `User Email`

### Excel File Format

- **File Type**: `.xlsx` (Excel 2007 or later)
- **Structure**: Header row followed by data rows
- **No Restrictions**: Remove password protection dan file permissions
- **Column Names**: Flexible - sistem will auto-detect variations

### Key Functions

**Dashboard Views:**
- Summary metrics cards
- Asset type distribution
- Warranty status tracking (Workstation)
- Age analysis dan lifecycle planning
- Regional/departmental breakdowns
- Visual analytics charts

**Filtering Options:**
- Filter by Model/Product
- Filter by Type, Site, Location, Department
- Filter by Status (Workstation) atau Programme (Mobile)
- Global text search

**Export Options:**
- Export all filtered data
- Export marked for replacement
- Export expired warranties (Workstation only)

---

## 🔒 Data Security

- **No Server Storage** - Data tidak disimpan di server
- **In-Memory Processing** - Semua pemprosesan dalam session memory
- **Session-Based** - Data cleared bila browser closed
- **Private & Secure** - Data remains completely confidential

---

## 🛠️ Tech Stack

- **Framework**: Streamlit 1.28+
- **Data Processing**: Pandas 2.0+
- **Excel Handling**: OpenPyXL 3.1+
- **Visualization**: Plotly 5.17+

---

## 📞 Contact & Support

**Developer**: MKAR  
**Email**: khalis.abdrahim@gmail.com  
**GitHub**: [@mdk869](https://github.com/mdk869)

**Response Time:**
- Mon-Fri: Within 24 hours
- Weekend: Within 48 hours

**Issues & Bug Reports**: [Create an Issue](https://github.com/mdk869/ITAM/issues)

---

## 📝 License

MIT License - see [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

Built with:
- [Streamlit](https://streamlit.io/) - Web framework
- [Pandas](https://pandas.pydata.org/) - Data manipulation
- [Plotly](https://plotly.com/) - Interactive visualizations
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Excel processing

---

<div align="center">

**Made with ❤️ by MKAR**

© 2025 MKAR. All rights reserved.

[Report Bug](https://github.com/mdk869/ITAM/issues) · [Request Feature](https://github.com/mdk869/ITAM/issues)

</div>
