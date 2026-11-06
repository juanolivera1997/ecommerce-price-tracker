"""E-commerce price tracker for the Books to Scrape demo store.

Crawls the paginated catalogue, extracts each book's title, price, star rating
and stock availability, cleans the data with pandas, and exports a formatted
Excel report.

Usage:
    python scraper.py
    python scraper.py --max-pages 3 --output output/sample.xlsx
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "http://books.toscrape.com/"

# Separate (connect, read) timeouts, in seconds, so a request can never hang.
REQUEST_TIMEOUT = (5, 20)
DEFAULT_DELAY_SECONDS = 1.0

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}

# The site encodes ratings as a CSS class, e.g. <p class="star-rating Three">.
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

COLUMN_NAMES = {
    "title": "Title",
    "price": "Price (GBP)",
    "rating": "Rating",
    "availability": "Availability",
}

logger = logging.getLogger("price_tracker")


def configure_logging() -> None:
    """Send clean, timestamped log lines to the console."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )
    # Hide urllib3's verbose retry messages; final failures are logged by fetch_page.
    logging.getLogger("urllib3").setLevel(logging.ERROR)


def create_session() -> requests.Session:
    """Build a session with realistic headers and automatic retries."""
    retry_policy = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        raise_on_status=False,  # Let raise_for_status() report the final status.
    )
    adapter = HTTPAdapter(max_retries=retry_policy)

    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def fetch_page(session: requests.Session, url: str) -> BeautifulSoup | None:
    """Download a page and parse it. Return None if the request fails."""
    try:
        response = session.get(url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
    except requests.exceptions.Timeout:
        logger.error("Timed out while requesting %s", url)
    except requests.exceptions.ConnectionError:
        logger.error("Could not connect to %s", url)
    except requests.exceptions.HTTPError as exc:
        logger.error("HTTP %s returned by %s", exc.response.status_code, url)
    except requests.exceptions.RequestException as exc:
        logger.error("Request failed for %s: %s", url, exc)
    else:
        # The server sends no charset, so requests would otherwise assume
        # ISO-8859-1 and turn "£" into "Â£". The pages are actually UTF-8.
        response.encoding = "utf-8"
        return BeautifulSoup(response.text, "html.parser")
    return None


def parse_rating(card: Tag) -> int | None:
    """Convert a card's star-rating CSS class into a number from 1 to 5."""
    rating_tag = card.select_one("p.star-rating")
    if rating_tag is None:
        return None
    for css_class in rating_tag.get("class", []):
        if css_class in RATING_WORDS:
            return RATING_WORDS[css_class]
    return None


def parse_book(card: Tag) -> dict[str, str | int | None]:
    """Extract the raw fields from one product card.

    Raises:
        ValueError: If a required field is missing from the card.
    """
    title_tag = card.select_one("h3 a")
    price_tag = card.select_one("p.price_color")
    availability_tag = card.select_one("p.availability")

    if title_tag is None or price_tag is None or availability_tag is None:
        raise ValueError("product card is missing a required element")

    # The visible link text is truncated; the title attribute has the full name.
    title = title_tag.get("title") or title_tag.get_text()

    return {
        "title": title,
        "price": price_tag.get_text(),
        "rating": parse_rating(card),
        "availability": availability_tag.get_text(),
    }


def parse_books_from_page(soup: BeautifulSoup, url: str) -> list[dict]:
    """Extract every valid book on a page, skipping malformed cards."""
    books = []
    for card in soup.select("article.product_pod"):
        try:
            books.append(parse_book(card))
        except ValueError as exc:
            logger.warning("Skipped a product on %s: %s", url, exc)
    return books


def get_next_page_url(soup: BeautifulSoup, current_url: str) -> str | None:
    """Return the absolute URL of the next page, or None on the last page."""
    next_link = soup.select_one("li.next a")
    if next_link is None or not next_link.get("href"):
        return None
    return urljoin(current_url, next_link["href"])


def scrape_catalogue(
    session: requests.Session,
    start_url: str = BASE_URL,
    max_pages: int | None = None,
    delay: float = DEFAULT_DELAY_SECONDS,
) -> list[dict]:
    """Follow the pagination links and collect books from each page.

    If a page fails after all retries, the crawl stops and returns the books
    collected so far, so a late failure does not discard earlier work.
    """
    books: list[dict] = []
    visited: set[str] = set()
    url: str | None = start_url

    while url and url not in visited:
        if max_pages is not None and len(visited) >= max_pages:
            logger.info("Reached the page limit (%d)", max_pages)
            break

        visited.add(url)
        soup = fetch_page(session, url)
        if soup is None:
            logger.warning("Stopping early; keeping the %d books collected", len(books))
            break

        page_books = parse_books_from_page(soup, url)
        books.extend(page_books)
        logger.info(
            "Page %d: %d books (total: %d)", len(visited), len(page_books), len(books)
        )

        url = get_next_page_url(soup, url)
        if url:
            time.sleep(delay)  # Be polite to the server between requests.

    return books


def clean_data(raw_books: list[dict]) -> pd.DataFrame:
    """Normalize raw records into a typed DataFrame ready for export."""
    df = pd.DataFrame(raw_books, columns=list(COLUMN_NAMES))

    for column in ("title", "availability"):
        df[column] = (
            df[column].astype("string").str.replace(r"\s+", " ", regex=True).str.strip()
        )

    # Keep digits and the decimal point only, so "£51.77" becomes 51.77.
    df["price"] = pd.to_numeric(
        df["price"].astype("string").str.replace(r"[^\d.]", "", regex=True),
        errors="coerce",
    ).astype("float64")

    # Nullable integers keep ratings whole even when one is missing.
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce").astype("Int64")

    invalid_prices = df["price"].isna()
    if invalid_prices.any():
        logger.warning("Dropped %d rows with an invalid price", invalid_prices.sum())
        df = df[~invalid_prices]

    return df.rename(columns=COLUMN_NAMES).reset_index(drop=True)


def format_worksheet(worksheet: Worksheet) -> None:
    """Apply report styling: header, number formats, filters and column widths."""
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    row_border = Border(bottom=Side(style="thin", color="D9D9D9"))
    column_widths = {"Title": 60, "Price (GBP)": 14, "Rating": 10, "Availability": 16}

    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        letter = get_column_letter(cell.column)
        worksheet.column_dimensions[letter].width = column_widths[cell.value]

    for title, price, rating, availability in worksheet.iter_rows(min_row=2):
        title.alignment = Alignment(wrap_text=True, vertical="top")
        price.number_format = '"£"#,##0.00'
        rating.alignment = Alignment(horizontal="center", vertical="top")
        price.alignment = Alignment(vertical="top")
        availability.alignment = Alignment(horizontal="center", vertical="top")
        for cell in (title, price, rating, availability):
            cell.border = row_border

    worksheet.row_dimensions[1].height = 22
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def export_to_excel(df: pd.DataFrame, output_path: Path) -> None:
    """Write the DataFrame to a formatted .xlsx file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Books", index=False)
        format_worksheet(writer.sheets["Books"])


def default_output_path() -> Path:
    """Build a timestamped path so each run keeps its own price snapshot."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path("output") / f"books_{timestamp}.xlsx"


def positive_int(value: str) -> int:
    """Argparse type for integers greater than zero."""
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be 1 or greater")
    return number


def non_negative_float(value: str) -> float:
    """Argparse type for numbers that are zero or greater."""
    number = float(value)
    if number < 0:
        raise argparse.ArgumentTypeError("must be 0 or greater")
    return number


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse command-line options."""
    parser = argparse.ArgumentParser(
        description="Scrape book prices from Books to Scrape and export them to Excel."
    )
    parser.add_argument(
        "--url", default=BASE_URL, help="catalogue URL to start from (default: %(default)s)"
    )
    parser.add_argument(
        "--max-pages",
        type=positive_int,
        default=None,
        help="stop after this many pages (default: all pages)",
    )
    parser.add_argument(
        "--delay",
        type=non_negative_float,
        default=DEFAULT_DELAY_SECONDS,
        help="seconds to wait between pages (default: %(default)s)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="output .xlsx path (default: output/books_<timestamp>.xlsx)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Run the pipeline and return a process exit code."""
    configure_logging()
    args = parse_args(argv)
    output_path = args.output or default_output_path()

    logger.info("Starting crawl at %s", args.url)
    try:
        with create_session() as session:
            raw_books = scrape_catalogue(session, args.url, args.max_pages, args.delay)
    except KeyboardInterrupt:
        logger.warning("Interrupted by user")
        return 130

    if not raw_books:
        logger.error("No books were extracted; no file was written")
        return 1

    df = clean_data(raw_books)
    if df.empty:
        logger.error("No valid rows remained after cleaning; no file was written")
        return 1

    try:
        export_to_excel(df, output_path)
    except PermissionError:
        logger.error("Cannot write %s; close it if it is open in Excel", output_path)
        return 1
    except OSError as exc:
        logger.error("Could not save %s: %s", output_path, exc)
        return 1

    logger.info("Saved %d books to %s", len(df), output_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
