"""Сквозной прогон: два запуска парсера и отчёт об изменениях между ними.

Демо-сайт статичен, цены на нём не меняются. Поэтому второй запуск получает те же
данные, а изменения в них вносятся здесь явно — так, как они происходили бы в живом
магазине: скидка, подорожание, распродан остаток, товар снят, товар добавлен.
"""

import copy
import os
import sys
import time
from dataclasses import replace
from pathlib import Path

HERE = Path(__file__).parent
os.chdir(HERE)
sys.path.insert(0, str(HERE))
sys.stdout.reconfigure(encoding="utf-8")

import catalog_parser as cp  # noqa: E402

DB = "demo.db"
OUT = Path("out")


def step(title):
    print(f"\n{'=' * 70}\n{title}\n{'=' * 70}")


OUT.mkdir(exist_ok=True)
Path(DB).unlink(missing_ok=True)
conn = cp.connect(DB)

step("1. Первый запуск: весь каталог")
started = time.perf_counter()
products, errors = cp.scrape()
elapsed = time.perf_counter() - started
first = cp.save_run(conn, products)
cp.export_excel(str(OUT / "catalog_run1.xlsx"), products, None, errors)
print(f"товаров {len(products)}, ошибок {len(errors)}, за {elapsed:.0f} с")
print(f"категорий {len({p.category for p in products})}, "
      f"артикулов уникальных {len({p.sku for p in products})}")

step("2. Второй запуск: в магазине за это время что-то поменялось")
second = copy.deepcopy(products)
second[10] = replace(second[10], price=round(second[10].price * 0.85, 2))   # скидка 15%
second[11] = replace(second[11], price=round(second[11].price * 1.10, 2))   # подорожание
second[12] = replace(second[12], stock=0)                                  # распродан
gone = [second.pop(13).title, second.pop(13).title]                        # сняты с продажи
second.append(cp.Product("demo-new-0001", "Новинка: книга добавлена после первого запуска",
                         "Poetry", 19.99, 12, 5, 0, cp.START_URL, ""))
run2 = cp.save_run(conn, second)
changes = cp.diff(cp.load_run(conn, first), cp.load_run(conn, run2))
cp.export_excel(str(OUT / "catalog_run2.xlsx"), second, changes)
for c in changes:
    delta = f" ({c.delta_pct:+.0%})" if c.delta_pct is not None else ""
    print(f"  {c.kind:8} {c.title[:45]:45} {c.old} → {c.new}{delta}")

step("3. Итог")
print(f"изменений найдено: {len(changes)} (ожидалось 6)")
print(f"файлы: {OUT / 'catalog_run1.xlsx'}, {OUT / 'catalog_run2.xlsx'}")
