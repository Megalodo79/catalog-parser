"""
Парсер каталога интернет-магазина с выгрузкой в Excel и учётом изменений между запусками.

Демо-источник — books.toscrape.com, учебный магазин, сделанный специально для отработки
парсинга. Под другой сайт меняются только функции разбора страниц (parse_*).

Что делает:
  1. обходит каталог по страницам и собирает ссылки на карточки товаров;
  2. параллельно скачивает карточки, повторяя запрос при сбоях сети и ответах 429/5xx;
  3. сохраняет снимок каталога в SQLite и сравнивает с прошлым запуском:
     новые и пропавшие товары, изменение цены и остатка;
  4. выгружает Excel: товары, изменения, сводка по категориям.
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import re
import sqlite3
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict, fields
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/catalogue/"
START_URL = urljoin(BASE_URL, "page-1.html")
DB_PATH = "catalog.db"

log = logging.getLogger("catalog")

RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}


# --------------------------------------------------------------------------- модель


@dataclass
class Product:
    sku: str  # артикул (на демо-сайте — UPC), ключ для сравнения запусков
    title: str
    category: str
    price: float
    stock: int
    rating: int
    reviews: int
    url: str
    image: str


# --------------------------------------------------------------------------- сеть


def make_session(retries: int = 4) -> requests.Session:
    """Сессия с повторами: сеть иногда сбоит, а сайт может ответить 429 при нагрузке."""
    session = requests.Session()
    retry = Retry(
        total=retries,
        backoff_factor=0.5,  # 0.5, 1, 2, 4 с между попытками
        status_forcelist=(429, 500, 502, 503, 504),
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_maxsize=16)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers["User-Agent"] = "Mozilla/5.0 (catalog-parser demo)"
    return session


def fetch(session: requests.Session, url: str) -> str:
    response = session.get(url, timeout=20)
    response.raise_for_status()
    # Сайт не всегда указывает кодировку в заголовке — без этого «£» превращается в «Â£».
    response.encoding = "utf-8"
    return response.text


# --------------------------------------------------------------------------- разбор


def parse_list_page(html: str, page_url: str) -> tuple[list[str], str | None]:
    """Ссылки на карточки со страницы каталога и адрес следующей страницы."""
    soup = BeautifulSoup(html, "lxml")
    links = [urljoin(page_url, a["href"]) for a in soup.select("article.product_pod h3 a")]
    nxt = soup.select_one("li.next a")
    return links, urljoin(page_url, nxt["href"]) if nxt else None


def _money(text: str) -> float:
    return float(re.sub(r"[^\d.]", "", text))


def parse_product_page(html: str, url: str) -> Product:
    soup = BeautifulSoup(html, "lxml")
    table = {row.th.get_text(strip=True): row.td.get_text(strip=True)
             for row in soup.select("table.table tr")}
    stock = re.search(r"\((\d+) available\)", table.get("Availability", ""))
    rating = soup.select_one("div.product_main p.star-rating")
    crumbs = soup.select("ul.breadcrumb li a")
    image = soup.select_one("div#product_gallery img, div.item img")
    return Product(
        sku=table["UPC"],
        title=soup.select_one("div.product_main h1").get_text(strip=True),
        category=crumbs[-1].get_text(strip=True) if len(crumbs) >= 3 else "",
        price=_money(table["Price (incl. tax)"]),
        stock=int(stock.group(1)) if stock else 0,
        rating=next((RATINGS[c] for c in rating["class"] if c in RATINGS), 0) if rating else 0,
        reviews=int(table.get("Number of reviews", "0") or 0),
        url=url,
        image=urljoin(url, image["src"]) if image else "",
    )


# --------------------------------------------------------------------------- обход


def collect_links(session, start_url: str, max_pages: int | None) -> list[str]:
    links, url, page = [], start_url, 0
    while url and (max_pages is None or page < max_pages):
        page_links, url = parse_list_page(fetch(session, url), url)
        links.extend(page_links)
        page += 1
        log.info("страница %d: %d товаров", page, len(page_links))
    return links


def scrape(start_url: str = START_URL, max_pages: int | None = None,
           workers: int = 8, session: requests.Session | None = None):
    """Возвращает (товары, ошибки). Битая карточка не останавливает весь обход."""
    session = session or make_session()
    links = collect_links(session, start_url, max_pages)
    products, errors = [], []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(fetch, session, link): link for link in links}
        for future in as_completed(futures):
            link = futures[future]
            try:
                products.append(parse_product_page(future.result(), link))
            except Exception as exc:  # noqa: BLE001 — ошибку пишем в отчёт и идём дальше
                errors.append((link, f"{type(exc).__name__}: {exc}"))
                log.warning("не разобрана %s: %s", link, exc)
    products.sort(key=lambda p: (p.category, p.title))
    return products, errors


# --------------------------------------------------------------------------- история


def connect(db_path: str = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS runs (
            run_id INTEGER PRIMARY KEY AUTOINCREMENT,
            started_at TEXT NOT NULL,
            items INTEGER NOT NULL
        );
        CREATE TABLE IF NOT EXISTS snapshots (
            run_id INTEGER NOT NULL REFERENCES runs(run_id),
            sku TEXT NOT NULL,
            title TEXT NOT NULL,
            category TEXT NOT NULL,
            price REAL NOT NULL,
            stock INTEGER NOT NULL,
            rating INTEGER NOT NULL,
            reviews INTEGER NOT NULL,
            url TEXT NOT NULL,
            image TEXT NOT NULL,
            PRIMARY KEY (run_id, sku)
        );
        """
    )
    return conn


def save_run(conn: sqlite3.Connection, products: list[Product]) -> int:
    with conn:
        cur = conn.execute("INSERT INTO runs (started_at, items) VALUES (?, ?)",
                           (dt.datetime.now().isoformat(timespec="seconds"), len(products)))
        run_id = cur.lastrowid
        cols = [f.name for f in fields(Product)]
        conn.executemany(
            f"INSERT INTO snapshots (run_id, {', '.join(cols)}) "
            f"VALUES (?, {', '.join('?' * len(cols))})",
            [(run_id, *asdict(p).values()) for p in products],
        )
    return run_id


def load_run(conn: sqlite3.Connection, run_id: int) -> dict[str, Product]:
    cols = [f.name for f in fields(Product)]
    rows = conn.execute(f"SELECT {', '.join(cols)} FROM snapshots WHERE run_id = ?", (run_id,))
    return {row[0]: Product(*row) for row in rows}


def previous_run_id(conn: sqlite3.Connection, run_id: int) -> int | None:
    row = conn.execute("SELECT MAX(run_id) FROM runs WHERE run_id < ?", (run_id,)).fetchone()
    return row[0]


@dataclass
class Change:
    kind: str  # Новый | Пропал | Цена | Остаток
    sku: str
    title: str
    old: float | int | None
    new: float | int | None
    url: str

    @property
    def delta_pct(self) -> float | None:
        if self.kind == "Цена" and self.old:
            return (self.new - self.old) / self.old
        return None


def diff(old: dict[str, Product], new: dict[str, Product]) -> list[Change]:
    changes = []
    for sku, p in new.items():
        before = old.get(sku)
        if before is None:
            changes.append(Change("Новый", sku, p.title, None, p.price, p.url))
            continue
        if round(before.price, 2) != round(p.price, 2):
            changes.append(Change("Цена", sku, p.title, before.price, p.price, p.url))
        if before.stock != p.stock:
            changes.append(Change("Остаток", sku, p.title, before.stock, p.stock, p.url))
    for sku, p in old.items():
        if sku not in new:
            changes.append(Change("Пропал", sku, p.title, p.price, None, p.url))
    order = {"Цена": 0, "Новый": 1, "Пропал": 2, "Остаток": 3}
    return sorted(changes, key=lambda c: (order[c.kind], c.title))


# --------------------------------------------------------------------------- Excel

HEADER_FILL = PatternFill("solid", fgColor="1F4E46")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY = '#,##0.00 "£"'


def _sheet(wb: Workbook, title: str, headers: list[str], widths: list[int], first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(headers)
    for col, width in enumerate(widths, 1):
        cell = ws.cell(row=1, column=col)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    return ws


def _link(cell, url: str, text: str = "открыть") -> None:
    cell.value, cell.hyperlink = text, url
    cell.font = Font(color="0563C1", underline="single")


def export_excel(path: str, products: list[Product], changes: list[Change] | None,
                 errors: list[tuple[str, str]] | None = None) -> str:
    wb = Workbook()

    ws = _sheet(wb, "Товары",
                ["Артикул", "Название", "Категория", "Цена", "Остаток", "Рейтинг",
                 "Отзывы", "Ссылка", "Фото"],
                [18, 55, 20, 12, 10, 10, 9, 11, 11], first=True)
    for p in products:
        ws.append([p.sku, p.title, p.category, p.price, p.stock, p.rating, p.reviews])
        row = ws.max_row
        ws.cell(row=row, column=4).number_format = MONEY
        _link(ws.cell(row=row, column=8), p.url)
        if p.image:
            _link(ws.cell(row=row, column=9), p.image, "фото")
    last = ws.max_row
    ws.auto_filter.ref = f"A1:I{last}"
    # Мало на складе — подсветка, чтобы менеджер видел сразу.
    ws.conditional_formatting.add(
        f"E2:E{last}",
        CellIsRule(operator="lessThan", formula=["5"], fill=PatternFill("solid", fgColor="F8CBAD")),
    )

    ws = _sheet(wb, "Изменения",
                ["Что изменилось", "Артикул", "Название", "Было", "Стало", "Δ, %", "Ссылка"],
                [16, 18, 55, 12, 12, 9, 11])
    if changes is None:
        ws.append(["Первый запуск — сравнивать пока не с чем"])
    elif not changes:
        ws.append(["С прошлого запуска ничего не изменилось"])
    for c in changes or []:
        ws.append([c.kind, c.sku, c.title, c.old, c.new, c.delta_pct])
        row = ws.max_row
        money = c.kind in ("Цена", "Новый", "Пропал")
        for col in (4, 5):
            if money:
                ws.cell(row=row, column=col).number_format = MONEY
        ws.cell(row=row, column=6).number_format = "+0.0%;-0.0%"
        _link(ws.cell(row=row, column=7), c.url)
        if c.delta_pct is not None:
            color = "C6EFCE" if c.delta_pct < 0 else "FFC7CE"  # дешевле — зелёный
            ws.cell(row=row, column=6).fill = PatternFill("solid", fgColor=color)

    ws = _sheet(wb, "Сводка", ["Категория", "Товаров", "Мин. цена", "Средняя цена",
                               "Макс. цена", "Остаток, шт."], [24, 10, 12, 14, 12, 13])
    by_cat: dict[str, list[Product]] = {}
    for p in products:
        by_cat.setdefault(p.category, []).append(p)
    for cat, items in sorted(by_cat.items(), key=lambda kv: -len(kv[1])):
        prices = [p.price for p in items]
        ws.append([cat, len(items), min(prices), sum(prices) / len(prices), max(prices),
                   sum(p.stock for p in items)])
        for col in (3, 4, 5):
            ws.cell(row=ws.max_row, column=col).number_format = MONEY
    total = ws.max_row
    ws.append(["Итого", f"=SUM(B2:B{total})", f"=MIN(C2:C{total})", None,
               f"=MAX(E2:E{total})", f"=SUM(F2:F{total})"])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    for col in (3, 5):
        ws.cell(row=ws.max_row, column=col).number_format = MONEY

    if errors:
        ws = _sheet(wb, "Ошибки", ["Ссылка", "Ошибка"], [70, 60])
        for url, err in errors:
            ws.append([url, err])

    wb.save(path)
    return path


# --------------------------------------------------------------------------- запуск


def run(out: str, max_pages: int | None, workers: int, db_path: str = DB_PATH) -> dict:
    started = time.perf_counter()
    products, errors = scrape(max_pages=max_pages, workers=workers)
    conn = connect(db_path)
    run_id = save_run(conn, products)
    prev = previous_run_id(conn, run_id)
    changes = diff(load_run(conn, prev), load_run(conn, run_id)) if prev else None
    export_excel(out, products, changes, errors)
    return {
        "run_id": run_id,
        "products": len(products),
        "errors": len(errors),
        "changes": None if changes is None else len(changes),
        "seconds": round(time.perf_counter() - started, 1),
        "out": out,
    }


def main(argv: list[str] | None = None) -> None:
    ap = argparse.ArgumentParser(description="Парсер каталога → Excel с историей изменений")
    ap.add_argument("--out", default="catalog.xlsx", help="файл Excel")
    ap.add_argument("--pages", type=int, default=None, help="сколько страниц каталога (по умолчанию все)")
    ap.add_argument("--workers", type=int, default=8, help="параллельных загрузок карточек")
    ap.add_argument("--db", default=DB_PATH, help="база истории запусков")
    args = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")
    result = run(args.out, args.pages, args.workers, args.db)
    changes = "первый запуск" if result["changes"] is None else result["changes"]
    print(f"Товаров: {result['products']}, ошибок: {result['errors']}, изменений: {changes}, "
          f"время: {result['seconds']} с → {result['out']}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
