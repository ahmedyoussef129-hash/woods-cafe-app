from __future__ import annotations
from pathlib import Path
import sqlite3
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
XLSX = Path("/home/Ahmedofsa129/woods cafe.xlsx")
DB = ROOT / "woods_cafe.db"

if not XLSX.exists():
    raise SystemExit(f"ملف Excel غير موجود: {XLSX}")

wb = load_workbook(XLSX, read_only=True, data_only=True)

ws_inv = wb["المخزن"]
ws_recipe = wb["الريسبي"]
ws_menu = wb["هندسة_المنيو"]

# أسعار البيع من هندسة المنيو، مع الاحتفاظ بالسعر الموجود في قاعدة البيانات
prices = {}
for r in ws_menu.iter_rows(min_row=4, values_only=True):
    name, price = r[0], r[1]
    if name and isinstance(price, (int, float)):
        prices[str(name).strip()] = float(price)

c = sqlite3.connect(DB)
c.row_factory = sqlite3.Row
c.execute("PRAGMA foreign_keys=ON")

# تأكد من وجود الجداول الأساسية
c.executescript("""
CREATE TABLE IF NOT EXISTS products(
  id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL,
  base_price REAL NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS inventory(
  id INTEGER PRIMARY KEY, item_name TEXT UNIQUE NOT NULL,
  unit TEXT NOT NULL DEFAULT 'piece', balance REAL NOT NULL DEFAULT 0,
  unit_cost REAL NOT NULL DEFAULT 0, reorder_level REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS recipes(
  id INTEGER PRIMARY KEY, product_id INTEGER NOT NULL,
  component_id INTEGER NOT NULL, qty REAL NOT NULL,
  FOREIGN KEY(product_id) REFERENCES products(id),
  FOREIGN KEY(component_id) REFERENCES inventory(id)
);
""")

# 1) استيراد المخزن
inventory_count = 0
inventory_names = set()

for r in ws_inv.iter_rows(min_row=4, values_only=True):
    name = r[0]
    if not name:
        continue

    name = str(name).strip()
    unit = str(r[1] or "piece").strip()
    current_balance = float(r[10] or 0)       # الرصيد الحالي K
    unit_cost = float(r[5] or 0)              # سعر الوحدة الأساسية F
    reorder_level = float(r[11] or 0)         # حد إعادة الطلب L

    inventory_names.add(name)

    existing = c.execute(
        "SELECT id FROM inventory WHERE item_name=?",
        (name,)
    ).fetchone()

    if existing:
        c.execute("""
            UPDATE inventory
            SET unit=?, balance=?, unit_cost=?, reorder_level=?
            WHERE id=?
        """, (unit, current_balance, unit_cost, reorder_level, existing["id"]))
    else:
        c.execute("""
            INSERT INTO inventory(item_name,unit,balance,unit_cost,reorder_level)
            VALUES(?,?,?,?,?)
        """, (name, unit, current_balance, unit_cost, reorder_level))

    inventory_count += 1

# 2) المنتجات: كل أصناف الريسبي + الأسعار المعروفة
recipe_products = set()
for r in ws_recipe.iter_rows(min_row=4, values_only=True):
    if r[0]:
        recipe_products.add(str(r[0]).strip())

product_names = set(recipe_products) | set(prices.keys())

product_count = 0
for name in sorted(product_names):
    existing = c.execute(
        "SELECT id,base_price FROM products WHERE name=?",
        (name,)
    ).fetchone()

    if existing:
        # لا نستبدل السعر الموجود إلا لو Excel عنده سعر فعلي
        new_price = prices.get(name, float(existing["base_price"] or 0))
        c.execute(
            "UPDATE products SET base_price=?, active=1 WHERE id=?",
            (new_price, existing["id"])
        )
    else:
        c.execute(
            "INSERT INTO products(name,base_price,active) VALUES(?,?,1)",
            (name, prices.get(name, 0))
        )
    product_count += 1

# 3) أضف أي مكون موجود في الريسبي لكنه غير موجود في المخزن
#    برصيد صفر وتكلفة صفر، حتى يظهر كعنصر ناقص بدل تجاهله.
for r in ws_recipe.iter_rows(min_row=4, values_only=True):
    component = r[1]
    unit = r[3]
    if not component:
        continue

    component = str(component).strip()
    if component not in inventory_names:
        c.execute("""
            INSERT OR IGNORE INTO inventory(item_name,unit,balance,unit_cost,reorder_level)
            VALUES(?,?,0,0,0)
        """, (component, str(unit or "piece").strip()))

# 4) إعادة بناء الريسبي من Excel
c.execute("DELETE FROM recipes")

recipe_count = 0
missing_components = []

for r in ws_recipe.iter_rows(min_row=4, values_only=True):
    product, component, qty, unit = r[0], r[1], r[2], r[3]

    if not product or not component or qty is None:
        continue

    product = str(product).strip()
    component = str(component).strip()

    try:
        qty = float(qty)
    except (TypeError, ValueError):
        continue

    p = c.execute(
        "SELECT id FROM products WHERE name=?",
        (product,)
    ).fetchone()

    i = c.execute(
        "SELECT id FROM inventory WHERE item_name=?",
        (component,)
    ).fetchone()

    if not p or not i:
        missing_components.append(f"{product} -> {component}")
        continue

    c.execute("""
        INSERT INTO recipes(product_id,component_id,qty)
        VALUES(?,?,?)
    """, (p["id"], i["id"], qty))

    recipe_count += 1

c.commit()
c.close()
wb.close()

print("======================================")
print("تم استيراد بيانات WOODS بنجاح")
print(f"المخزن: {inventory_count} صنف")
print(f"المنتجات: {product_count} صنف")
print(f"مكونات الريسبي: {recipe_count} سطر")
if missing_components:
    print("أسطر لم يتم استيرادها:")
    for x in missing_components[:20]:
        print(" -", x)
print("======================================")
