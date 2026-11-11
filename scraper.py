#!/usr/bin/env python3
"""Books to Scrape -> validated Microsoft Excel (.xlsx).

Portfolio example for https://books.toscrape.com — a public practice website.
Do not reuse on third-party sites without reviewing permission and rate limits.
"""
from __future__ import annotations

import argparse
import logging
import re
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
COLUMNS = ["Product Name", "Availability", "Product URL", "Price GBP", "Rating"]
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
LOG = logging.getLogger("books_scraper")


@dataclass(frozen=True)
class Book:
    name: str
    availability: str
    url: str
    price_gbp: float
    rating: int


class ScrapeError(RuntimeError):
    """An unexpected response or document structure was encountered."""


def is_allowed_url(url: str) -> bool:
    parsed = urlparse(url)
    return parsed.scheme == "https" and parsed.hostname == "books.toscrape.com"


def parse_listing(html: str, page_url: str) -> tuple[list[Book], str | None]:
    """Extract product cards and the pagination link from a catalogue page."""
    soup = BeautifulSoup(html, "html.parser")
    cards = soup.select("article.product_pod")
    if not cards:
        raise ScrapeError(f"No product cards found on {page_url}; layout may have changed")

    books: list[Book] = []
    for card in cards:
        anchor = card.select_one("h3 a[href]")
        price = card.select_one("p.price_color")
        stock = card.select_one("p.instock.availability")
        rating_tag = card.select_one("p.star-rating")
        if not all((anchor, price, stock, rating_tag)):
            raise ScrapeError(f"Missing product field on {page_url}")
        assert anchor is not None and price is not None
        assert stock is not None and rating_tag is not None
        name = (anchor.get("title") or anchor.get_text(" ", strip=True)).strip()
        link = urljoin(page_url, anchor["href"])
        raw_price = price.get_text(" ", strip=True)
        match = re.search(r"£\s*(\d+(?:\.\d{1,2})?)", raw_price)
        rating = next((RATINGS[k] for k in rating_tag.get("class", []) if k in RATINGS), None)
        if not name or not is_allowed_url(link) or not match or rating is None:
            raise ScrapeError(f"Invalid product data on {page_url}: {name!r}")
        books.append(Book(name, stock.get_text(" ", strip=True), link, float(match.group(1)), rating))

    next_node = soup.select_one("li.next a[href]")
    next_url = urljoin(page_url, next_node["href"]) if next_node else None
    if next_url and not is_allowed_url(next_url):
        raise ScrapeError(f"Unexpected pagination destination: {next_url}")
    return books, next_url


def build_session() -> requests.Session:
    """Retry transient errors while maintaining a polite request cadence."""
    session = requests.Session()
    session.headers.update({"User-Agent": "NovelHariKeswono-PortfolioScraper/1.0 (educational demo)"})
    retries = Retry(total=3, backoff_factor=0.7, status_forcelist=(429, 500, 502, 503, 504), allowed_methods={"GET"})
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    return session


def scrape_catalogue(session: requests.Session, max_pages: int = 50, delay: float = 0.7) -> list[Book]:
    """Follow only the allowed practice site's pagination; avoid loops and duplicates."""
    if not 1 <= max_pages <= 50:
        raise ValueError("max_pages must be between 1 and 50")
    if delay < 0:
        raise ValueError("delay must not be negative")
    result: list[Book] = []
    seen_pages: set[str] = set()
    next_url: str | None = BASE_URL
    while next_url and len(seen_pages) < max_pages:
        if next_url in seen_pages:
            raise ScrapeError(f"Pagination loop detected: {next_url}")
        if not is_allowed_url(next_url):
            raise ScrapeError(f"Blocked out-of-domain request: {next_url}")
        response = session.get(next_url, timeout=(5, 20))
        response.raise_for_status()
        if not is_allowed_url(response.url):
            raise ScrapeError(f"Unexpected HTTP redirect: {response.url}")
        books, following = parse_listing(response.text, response.url)
        result.extend(books)
        seen_pages.add(next_url)
        LOG.info("Page %d: %d books (total %d)", len(seen_pages), len(books), len(result))
        next_url = following
        if next_url and len(seen_pages) < max_pages:
            time.sleep(delay)
    return result


def to_dataframe(books: list[Book]) -> pd.DataFrame:
    return pd.DataFrame(
        [(b.name, b.availability, b.url, b.price_gbp, b.rating) for b in books],
        columns=COLUMNS,
    )


def validate_dataframe(df: pd.DataFrame, expected_count: int | None = None) -> dict[str, int]:
    """Fail loudly rather than silently exporting an incorrect workbook."""
    if df.empty:
        raise ValueError("No books were scraped")
    if list(df.columns) != COLUMNS:
        raise ValueError(f"Unexpected columns: {list(df.columns)}")
    missing = int(df.isna().sum().sum()) + int(df[COLUMNS[:3]].eq("").sum().sum())
    duplicates = int(df["Product URL"].duplicated().sum())
    invalid_prices = int((~df["Price GBP"].between(0.01, 1000)).sum())
    invalid_ratings = int((~df["Rating"].between(1, 5)).sum())
    invalid_urls = int((~df["Product URL"].map(is_allowed_url)).sum())
    diagnostics = {
        "records": len(df),
        "missing_fields": missing,
        "duplicate_urls": duplicates,
        "invalid_prices": invalid_prices,
        "invalid_ratings": invalid_ratings,
        "invalid_urls": invalid_urls,
    }
    failures = {k: v for k, v in diagnostics.items() if k != "records" and v}
    if expected_count is not None and len(df) != expected_count:
        failures["expected_count_mismatch"] = abs(expected_count - len(df))
    if failures:
        raise ValueError(f"Validation failed: {failures}")
    return diagnostics


def export_excel(df: pd.DataFrame, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Books")
        ws = writer.sheets["Books"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(name="Calibri", bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.alignment = Alignment(wrap_text=True)
        for col, width in {"A": 44, "B": 18, "C": 86, "D": 15, "E": 12}.items():
            ws.column_dimensions[col].width = width
        for row in ws.iter_rows(min_row=2):
            row[3].number_format = '"£"#,##0.00'
            row[4].number_format = "0"


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape Books to Scrape into a validated Excel workbook")
    parser.add_argument("--max-pages", type=int, default=50, help="Max 50 catalogue pages (default: 50)")
    parser.add_argument("--delay", type=float, default=0.7, help="Seconds between pages (minimum 0.5)")
    parser.add_argument("--output", type=Path, default=Path("books_clean.xlsx"))
    parser.add_argument("--expect-count", type=int, help="Optional expected count; use 1000 if checking a full catalogue")
    args = parser.parse_args()
    if args.delay < 0.5:
        parser.error("Use at least a 0.5-second delay between pages")
    if args.expect_count is not None and args.expect_count < 1:
        parser.error("--expect-count must be positive")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    with build_session() as session:
        books = scrape_catalogue(session, max_pages=args.max_pages, delay=args.delay)
    df = to_dataframe(books)
    report = validate_dataframe(df, expected_count=args.expect_count)
    export_excel(df, args.output)
    LOG.info("Saved %s; validation: %s", args.output, report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
