"""Картинка для README: лист «Изменения» из out/catalog_run2.xlsx (после run_demo.py)."""

import matplotlib
import openpyxl

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

ws = openpyxl.load_workbook("out/catalog_run2.xlsx")["Изменения"]
rows = [[c.value for c in r][:6] for r in ws.iter_rows()]
head, body = ["Изменение", "Артикул", "Название", "Было", "Стало", "Δ"], rows[1:]


def fmt(value, col, kind):
    if value is None:
        return "—"
    if col == 5:
        return f"{value:+.1%}"
    if col in (3, 4) and kind != "Остаток":
        return f"£{value:,.2f}"
    if col == 2 and len(value) > 46:
        return value[:44] + "…"
    return str(value)


cells = [[fmt(v, i, r[0]) for i, v in enumerate(r)] for r in body]
fig, ax = plt.subplots(figsize=(12, 3.0), dpi=150)
ax.axis("off")
table = ax.table(cellText=cells, colLabels=head, loc="upper center", cellLoc="left",
                 colWidths=[0.11, 0.17, 0.42, 0.09, 0.09, 0.08])
table.auto_set_font_size(False)
table.set_fontsize(9.5)
table.scale(1, 1.6)
for (r, c), cell in table.get_celld().items():
    cell.PAD = 0.04
    cell.set_edgecolor("#D0D7D5")
    if r == 0:
        cell.set_facecolor("#1F4E46")
        cell.get_text().set_color("white")
        cell.get_text().set_weight("bold")
    elif c == 5 and body[r - 1][5] is not None:
        cell.set_facecolor("#C6EFCE" if body[r - 1][5] < 0 else "#FFC7CE")
ax.set_title("Лист «Изменения» — что поменялось в каталоге с прошлого запуска",
             loc="left", fontsize=11, pad=4)
plt.savefig("preview.png", bbox_inches="tight", facecolor="white")
print("preview.png")
