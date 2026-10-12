"""Тесты работают без интернета: на сохранённых страницах демо-сайта."""

import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import catalog_parser as cp  # noqa: E402

FIX = Path(__file__).parent / "fixtures"
LIST_URL = "https://books.toscrape.com/catalogue/page-1.html"
PRODUCT_URL = "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"


def html(name):
    return (FIX / name).read_text(encoding="utf-8")


def product(sku="A", price=10.0, stock=5, title="Книга"):
    return cp.Product(sku, title, "Poetry", price, stock, 3, 0, "https://x/" + sku, "")


def test_list_page_gives_links_and_next():
    links, nxt = cp.parse_list_page(html("list_page.html"), LIST_URL)
    assert len(links) == 20
    assert links[0] == PRODUCT_URL
    assert nxt == "https://books.toscrape.com/catalogue/page-2.html"


def test_last_page_has_no_next():
    links, nxt = cp.parse_list_page(html("last_page.html"),
                                    "https://books.toscrape.com/catalogue/page-50.html")
    assert len(links) == 20 and nxt is None


def test_product_page_fields():
    p = cp.parse_product_page(html("product_page.html"), PRODUCT_URL)
    assert p.sku == "a897fe39b1053632"
    assert p.title == "A Light in the Attic"
    assert p.category == "Poetry"
    assert p.price == 51.77
    assert p.stock == 22
    assert p.rating == 3
    assert p.image.startswith("https://books.toscrape.com/media/")


class FakeSession:
    """Отдаёт сохранённые страницы; одна карточка «сломана», чтобы проверить устойчивость."""

    def get(self, url, timeout=None):
        class R:
            encoding = None

            def raise_for_status(self):
                pass

        r = R()
        if url == LIST_URL:
            r.text = html("list_page.html")
        elif url == PRODUCT_URL:
            r.text = html("product_page.html")
        else:
            r.text = "<html>карточка без таблицы</html>"
        return r


def test_broken_card_does_not_stop_scrape():
    products, errors = cp.scrape(LIST_URL, max_pages=1, workers=4, session=FakeSession())
    assert len(products) == 1 and products[0].sku == "a897fe39b1053632"
    assert len(errors) == 19


def test_diff_finds_every_kind_of_change():
    old = {p.sku: p for p in [product("A", 10), product("B", 20), product("C", stock=5),
                              product("D")]}
    new = {p.sku: p for p in [product("A", 8.5), product("B", 20), product("C", stock=0),
                              product("E")]}
    kinds = {(c.kind, c.sku) for c in cp.diff(old, new)}
    assert kinds == {("Цена", "A"), ("Остаток", "C"), ("Пропал", "D"), ("Новый", "E")}
    price = next(c for c in cp.diff(old, new) if c.kind == "Цена")
    assert price.delta_pct == pytest.approx(-0.15)


def test_history_roundtrip(tmp_path):
    conn = cp.connect(str(tmp_path / "h.db"))
    first = cp.save_run(conn, [product("A", 10)])
    second = cp.save_run(conn, [product("A", 12)])
    assert cp.previous_run_id(conn, second) == first
    assert cp.previous_run_id(conn, first) is None
    changes = cp.diff(cp.load_run(conn, first), cp.load_run(conn, second))
    assert [(c.kind, c.old, c.new) for c in changes] == [("Цена", 10, 12)]


def test_excel_sheets_and_links(tmp_path):
    path = tmp_path / "out.xlsx"
    items = [product("A", 10), product("B", 30, stock=2)]
    cp.export_excel(str(path), items, cp.diff({}, {p.sku: p for p in items}),
                    [("https://x/bad", "ValueError")])
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Товары", "Изменения", "Сводка", "Ошибки"]
    ws = wb["Товары"]
    assert ws.max_row == 3 and ws["H2"].hyperlink.target == "https://x/A"
    assert wb["Сводка"]["B3"].value == "=SUM(B2:B2)"
