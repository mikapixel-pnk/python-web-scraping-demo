# Python Web Scraping → Excel

[![Tests](https://github.com/mikapixel-pnk/python-web-scraping-demo/actions/workflows/tests.yml/badge.svg)](https://github.com/mikapixel-pnk/python-web-scraping-demo/actions/workflows/tests.yml)

**A reproducible Python scraping, data cleaning, validation, and Excel export demonstration.**

Built by **Novel Hari Keswono** (Pontianak, Indonesia) as a public technical portfolio sample.

> **Demo disclosure:** This repository is a newly prepared demonstration based on a Books to Scrape workflow. The site [Books to Scrape](https://books.toscrape.com/) is designed for scraping practice. The included HTML fixtures and sample workbook contain **synthetic example entries**, not live scraping results or a real client's data.

## What it does

- Reads paginated book listings from `https://books.toscrape.com/`
- Extracts **Product Name, Availability, Product URL, Price GBP, Rating**
- Resolves relative links and handles pagination safely (up to 50 pages)
- Uses HTTP timeouts, retries, a conservative delay and a fixed allowed domain
- Verifies missing fields, duplicate product URLs, out-of-range prices and ratings
- Exports a fully editable **.xlsx** workbook, including filters, frozen header, and currency formatting
- Includes offline unit tests (no network requests needed)

## Quick start

Python 3.10+ recommended.

```bash
git clone https://github.com/mikapixel-pnk/python-web-scraping-demo.git
cd python-web-scraping-demo
python -m pip install -r requirements.txt
python scraper.py --max-pages 2 --output books_clean.xlsx
```

For a full catalogue run:

```bash
python scraper.py --max-pages 50 --delay 0.7 --output books_clean.xlsx
```

The practice site has traditionally contained about 1,000 book listings, but **record counts can change**. To require an exact count, optionally append `--expect-count 1000`.

## Output schema

| Column | Example |
| --- | --- |
| Product Name | Sample Book Alpha |
| Availability | In stock |
| Product URL | https://books.toscrape.com/catalogue/sample-book-alpha_1/index.html |
| Price GBP | 12.50 |
| Rating | 3 |

The example above is synthetic. Real exports contain the listings returned by the website at run time.

**Sample Excel:** Every successful [GitHub Actions run](https://github.com/mikapixel-pnk/python-web-scraping-demo/actions/workflows/tests.yml) generates a downloadable artifact named **`books-demo-synthetic-xlsx`**, with a three-row synthetic workbook built from the included HTML fixtures. This is *not* a live scrape. For a full dataset, run `scraper.py` against the practice website.

## Offline tests

```bash
python -m unittest discover -s tests -v
```

Tests cover parsing, pagination, price/rating validation, duplicate detection, rejecting off-domain pagination, and Excel formatting. **A passing CI run verifies the local fixtures and code path; it does not assert that the live website is reachable or has 1,000 books.**

## Files

| Path | Purpose |
| --- | --- |
| `scraper.py` | Scraper, validation and Excel export |
| `tests/test_scraper.py` | Offline automated tests |
| `examples/catalogue_page_1.html` | Synthetic fixture with pagination |
| `examples/catalogue_page_2.html` | Synthetic fixture (last page) |
| `examples/books_demo_sample.csv` | Previewable synthetic three-row sample | 
| `requirements.txt` | Python dependencies |
| `.github/workflows/tests.yml` | GitHub Actions continuous integration |

## Responsible use

The scraper is limited to the dedicated practice site. Check terms and policies before working with other domains. Respect rate limits, avoid collecting personal information, and do not scrape authenticated or private content without permission.

## About

Novel Hari Keswono · IT Support, PHP/Laravel, Python automation, web maintenance  
[GitHub Profile](https://github.com/mikapixel-pnk) · [LinkedIn](https://www.linkedin.com/in/novelharikeswono/)
