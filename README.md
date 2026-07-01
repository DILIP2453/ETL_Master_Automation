# Production Python ETL Automation

A Python-based ETL automation system that consolidates return data from multiple ecommerce marketplaces into a single daily reporting workflow.

This project was built while learning Python and solving a real operational problem in an ecommerce returns process.

---

## Overview

Managing returns across multiple marketplaces often requires different manual workflows.

This project automates return processing for:

- Flipkart (Courier PDF)
- Amazon (Return Sheet)
- Myntra (Daily Export)

Instead of handling each marketplace separately, all data flows through one shared ETL pipeline that validates, matches, reconciles, and generates a single report.

---

## Workflow

```
Flipkart Returns PDF
          │
Amazon Return Sheet
          │
 Myntra Daily Export
          │
──────── Extract ────────
          │
──────── Clean ──────────
          │
────── Data Validation ──
          │
──── Shared MIS Lookup ──
          │
──── Record Matching ────
          │
── Incremental Processing
          │
──── Report Generation ──
          │
Daily Summary Report
```

---

## Features

- PDF Processing
- Excel Automation
- Multi-Portal ETL Workflow
- Shared MIS Lookup
- Data Validation
- Record Matching
- Multi-Portal Reconciliation
- Duplicate Prevention
- Incremental Processing
- Audit Logging
- Daily Report Generation

---

## Supported Marketplaces

| Marketplace | Input |
|-------------|-------|
| Flipkart | Courier PDF |
| Amazon | Return Excel Sheet |
| Myntra | Daily Export Sheet |

---

## Tech Stack

- Python
- Pandas
- OpenPyXL
- PDFPlumber
- Regular Expressions (Regex)
- Excel Automation

---

## Project Structure

```
ETL_Master_Automation/
│
├── Input/
├── Output/
├── Logs/
├── Config/
├── ETL_Master_Automation.py
├── requirements.txt
└── README.md
```

---

## Highlights

✔ Multi-portal automation

✔ Shared ETL pipeline

✔ Reusable validation logic

✔ Duplicate-safe processing

✔ Incremental execution

✔ Automated Excel reporting

✔ Audit logging

---

## Learning Outcomes

Building this project helped me learn:

- Python automation
- ETL workflow design
- Data validation
- Excel automation
- PDF data extraction
- Reusable workflow design
- Production-oriented scripting
- Problem solving for business operations

---

## Future Improvements

- Scheduled execution
- Email notification support
- Better configuration management
- Enhanced validation rules
- Dashboard reporting

---

## Repository

GitHub

https://github.com/DILIP2453/ETL_Master_Automation

---

## LinkedIn

If this project was useful or interesting, feel free to connect with me on LinkedIn.

---

## License

This project is shared for learning and portfolio purposes.
