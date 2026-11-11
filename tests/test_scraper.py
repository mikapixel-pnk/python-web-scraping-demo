import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from openpyxl import load_workbook

from scraper import Book, ScrapeError, export_excel, parse_listing, scrape_catalogue, to_dataframe, validate_dataframe

ROOT = Path(__file__).resolve().parents[1]
PAGE1 = (ROOT / "examples/catalogue_page_1.html").read_text(encoding="utf-8")
PAGE2 = (ROOT / "examples/catalogue_page_2.html").read_text(encoding="utf-8")


class FakeResponse:
    def __init__(self, url, text):
        self.url = url
        self.text = text

    def raise_for_status(self):
        pass


class FakeSession:
    def __init__(self):
        self.calls = []

    def get(self, url, timeout):
        self.calls.append(url)
        pages = {"https://books.toscrape.com/": PAGE1,
                 "https://books.toscrape.com/catalogue/page-2.html": PAGE2}
        return FakeResponse(url, pages[url])


class ScraperTests(unittest.TestCase):
    def test_parse_html_and_next_page(self):
        books, next_url = parse_listing(PAGE1, "https://books.toscrape.com/")
        self.assertEqual(len(books), 2)
        self.assertEqual(books[0].name, "Sample Book Alpha")
        self.assertEqual(books[0].price_gbp, 12.5)
        self.assertEqual(books[0].rating, 3)
        self.assertEqual(books[1].rating, 5)
        self.assertEqual(next_url, "https://books.toscrape.com/catalogue/page-2.html")

    def test_scrape_multiple_pages_and_validate(self):
        session = FakeSession()
        books = scrape_catalogue(session, max_pages=5, delay=0)
        self.assertEqual(len(books), 3)
        self.assertEqual(len(session.calls), 2)
        report = validate_dataframe(to_dataframe(books), expected_count=3)
        self.assertEqual(report["duplicate_urls"], 0)

    def test_reject_duplicate_url(self):
        book = Book("Example", "In stock", "https://books.toscrape.com/catalogue/example/index.html", 5.0, 3)
        with self.assertRaisesRegex(ValueError, "duplicate_urls"):
            validate_dataframe(to_dataframe([book, replace(book, name="Other")]))

    def test_reject_price_and_rating_errors(self):
        book = Book("Example", "In stock", "https://books.toscrape.com/catalogue/example/index.html", 0, 6)
        with self.assertRaisesRegex(ValueError, "invalid_prices"):
            validate_dataframe(to_dataframe([book]))

    def test_block_external_pagination(self):
        altered = PAGE1.replace("catalogue/page-2.html", "https://evil.example/test")
        with self.assertRaises(ScrapeError):
            parse_listing(altered, "https://books.toscrape.com/")

    def test_export_xlsx_features(self):
        df = to_dataframe(scrape_catalogue(FakeSession(), max_pages=5, delay=0))
        with tempfile.TemporaryDirectory() as temp:
            p = Path(temp) / "output.xlsx"
            export_excel(df, p)
            wb = load_workbook(p)
            ws = wb["Books"]
            self.assertEqual(ws.max_row, 4)
            self.assertEqual(ws.freeze_panes, "A2")
            self.assertEqual(ws["A1"].value, "Product Name")
            self.assertEqual(ws["D2"].value, 12.5)
            self.assertEqual(ws.auto_filter.ref, "A1:E4")


if __name__ == "__main__":
    unittest.main()
