# 📊 DATASUS - SIGTAP Unified Table

[![Sync Status](https://github.com/btran12/SIGTAP/actions/workflows/sync-datasus.yml/badge.svg)](https://github.com/btran12/SIGTAP/actions/workflows/sync-datasus.yml)
[![Last Update](https://img.shields.io/badge/last%20update-2026--11--14-blue)](https://github.com/btran12/SIGTAP/commits/main)

Automated repository mirroring the files of the **Unified Table of the Procedures Table Management System (SIGTAP)** from DATASUS.

## 📦 Unified Tables (Last 6 Months)

> **Total:** 6 files | **Size:** 13M

| File | Period | Size | Download |
|------|--------|------|----------|
| `TabelaUnificada_202611_v2611051737.zip` | November/2026 | 2.1M | [⬇️ Download](https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202611_v2611051737.zip) |
| `TabelaUnificada_202610_v2610271043.zip` | October/2026 | 2.1M | [⬇️ Download](https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202610_v2610271043.zip) |
| `TabelaUnificada_202609_v2610261311.zip` | September/2026 | 2.1M | [⬇️ Download](https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202609_v2610261311.zip) |
| `TabelaUnificada_202608_v2610231712.zip` | August/2026 | 2.1M | [⬇️ Download](https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202608_v2610231712.zip) |
| `TabelaUnificada_202607_v2610231237.zip` | July/2026 | 2.1M | [⬇️ Download](https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202607_v2610231237.zip) |
| `TabelaUnificada_202606_v2610231231.zip` | June/2026 | 2.1M | [⬇️ Download](https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202606_v2610231231.zip) |

## 🔍 Monthly Diff Reports

Download the latest tables and browse month-to-month comparison reports from [`index.html`](./index.html). Date-related changes are ignored.

Open `index.html` and use the 🔔 button to enable on-device notifications when a new monthly table is published.


## 📄 Technical Notes

> **Total:** 227 documents

All available technical notes are in the [`notastecnicas/`](./notastecnicas/) folder

### Latest Notes

- [`nota_tecnica_cgsi_sigtap_2026_11.pdf`](https://github.com/btran12/SIGTAP/blob/main/notastecnicas/nota_tecnica_cgsi_sigtap_2026_11.pdf)
- [`nota_tecnica_cgsi_sigtap_2026_10.pdf`](https://github.com/btran12/SIGTAP/blob/main/notastecnicas/nota_tecnica_cgsi_sigtap_2026_10.pdf)
- [`nota_tecnica_cgsi_sigtap_2026_09.pdf`](https://github.com/btran12/SIGTAP/blob/main/notastecnicas/nota_tecnica_cgsi_sigtap_2026_09.pdf)
- [`nota_tecnica_cgsi_sigtap_2026_08.pdf`](https://github.com/btran12/SIGTAP/blob/main/notastecnicas/nota_tecnica_cgsi_sigtap_2026_08.pdf)
- [`nota_tecnica_cgsi_sigtap_2026_07.pdf`](https://github.com/btran12/SIGTAP/blob/main/notastecnicas/nota_tecnica_cgsi_sigtap_2026_07.pdf)

## 📖 About SIGTAP

The SIGTAP Unified Table contains:
- Medical and dental procedures
- Medications and OPM (Orthoses, Prostheses and Special Materials)
- SUS reimbursement values
- ICD-10 and compatibility mappings
- Authorizations and classifications

## 🔄 Automatic Update

This repository is updated **daily at 5 AM (Brasília time)** via GitHub Actions.

**Official source:** [DATASUS FTP](ftp://ftp2.datasus.gov.br/pub/sistemas/tup/downloads/)

## 💻 How to Use

### Download via wget/curl
```bash
# Latest available table
wget https://github.com/btran12/SIGTAP/raw/main/tabelas/TabelaUnificada_202601_v2601221740.zip
```

### Clone the repository
```bash
git clone https://github.com/btran12/SIGTAP.git
cd SIGTAP/tabelas
```

### GitHub API (programmatic)
```bash
curl -s https://api.github.com/repos/btran12/SIGTAP/contents/tabelas | jq -r '.[].download_url'
```

## 📊 File Structure

```
SIGTAP/
├── tabelas/                    # Unified tables (6 most recent)
│   └── TabelaUnificada_*.zip
└── notastecnicas/              # Official technical documentation
    └── *.pdf
```

## ⚖️ License and Public Data

The data is **public domain** and maintained by the Brazilian Ministry of Health.
This repository only facilitates access to the official files.

---

<div align="center">

**🤖 Automatically updated** | Last run: 2026-11-14 11:56:47 BRT

</div>
