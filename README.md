# MahakimScraper-DataScraper

Automated scraper for [Mahakim.ma](https://www.mahakim.ma) to extract judicial police report data from Moroccan courts.  

⚖️ **Legal Note:**  
This scraper only collects **information publicly available** on Mahakim.ma. No private or confidential data is accessed. Use responsibly and in compliance with local regulations.

---

## Features

- ✅ Dynamic dropdown selection for court, police unit, and police station.
- ✅ Checkbox handling and precise case number/year input.
- ✅ Robust table detection with fallback strategies.
- ✅ Excel export using `openpyxl`.
- ✅ Detailed progress logs saved in `progress.txt`, including all dropdown selections and step information.
- ✅ Flexible and adaptive to changes in website options.
- ✅ Resumable scraping (picks up from last saved progress).

---

## Requirements

- Python 3.9+
- Google Chrome
- Packages:

```bash
pip install selenium>=4.12.0 webdriver-manager>=4.0.0 pandas>=2.1.0 openpyxl>=3.1.2
