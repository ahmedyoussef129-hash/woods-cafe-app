from __future__ import annotations
from pathlib import Path
import sqlite3
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parent
XLSX=Path("/home/Ahmedofsa129/woods cafe.xlsx")
DB=ROOT/"woods_cafe.db"

if not XLSX.exists():
    raise SystemExit(f"ملف Excel غير موجود: {XLSX}")

wb=load_workbook(XLSX,read_only=True,data_only=True)

c=sqlite3.connect(DB)
c.row_factory=sqlite3.Row
c.execute("PRAGMA foreign_keys=ON")
c.executescript("""
CREATE TABLE IF NOT EXISTS products(
 id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,base_price REAL NOT NULL DEFAULT 0,active INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS inventory(
 id INTEGER PRIMARY KEY,item_name TEXT UNIQUE NOT NULL,unit TEXT NOT NULL DEFAULT 'piece',
 balance REAL NOT NULL DEFAULT 0,unit_cost REAL NOT NULL DEFAULT 0,reorder_level REAL NOT NULL DEFAULT 0,
 purchase_unit TEXT DEFAULT '',package_qty REAL NOT NULL DEFAULT 1,package_price REAL NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS recipes(
 id INTEGER PRIMARY KEY,product_id INTEGER NOT NULL,component_id INTEGER NOT NULL,qty REAL NOT NULL,
 FOREIGN KEY(product_id) REFERENCES products(id),FOREIGN KEY(component_id) REFERENCES inventory(id));
""")
for col,typ in [("purchase_unit","TEXT DEFAULT ''"),("package_qty","REAL NOT NULL DEFAULT 1"),("package_price","REAL NOT NULL DEFAULT 0")]:
    try:c.execute(f"ALTER TABLE inventory ADD COLUMN {col} {typ}")
    except sqlite3.OperationalError:pass

inv=wb["المخزن"]
inventory_count=0
for r in inv.iter_rows(min_row=4,values_only=True):
    name=r[0]
    if not name: continue
    name=str(name).strip()
    unit=str(r[1] or "قطعة").strip()
    purchase_unit=str(r[2] or "").strip()
    package_qty=float(r[3] or 1)
    package_price=float(r[4] or 0)
    unit_cost=float(r[5] or 0)
    opening=float(r[6] or 0)
    balance=float(r[10] or 0)
    reorder=float(r[11] or 0)
    # Import the current Excel balance as the opening/current balance for initial setup.
    row=c.execute("SELECT id FROM inventory WHERE item_name=?",(name,)).fetchone()
    if row:
        c.execute("""UPDATE inventory SET unit=?,balance=?,unit_cost=?,reorder_level=?,
                     purchase_unit=?,package_qty=?,package_price=? WHERE id=?""",
                  (unit,balance,unit_cost,reorder,purchase_unit,package_qty,package_price,row["id"]))
    else:
        c.execute("""INSERT INTO inventory(item_name,unit,balance,unit_cost,reorder_level,purchase_unit,package_qty,package_price)
                     VALUES(?,?,?,?,?,?,?,?)""",
                  (name,unit,balance,unit_cost,reorder,purchase_unit,package_qty,package_price))
    inventory_count+=1

menu=wb["هندسة_المنيو"]
prices={}
for r in menu.iter_rows(min_row=4,values_only=True):
    name=r[0]; price=r[1]
    if name and isinstance(price,(int,float)):
        prices[str(name).strip()]=float(price)

recipe=wb["الريسبي"]
product_names=set(prices)
recipe_rows=[]
for r in recipe.iter_rows(min_row=4,values_only=True):
    product,component,qty,unit=r[0],r[1],r[2],r[3]
    if product: product_names.add(str(product).strip())
    if product and component and qty is not None:
        try: q=float(qty)
        except: continue
        recipe_rows.append((str(product).strip(),str(component).strip(),q,str(unit or "قطعة").strip()))

for name in sorted(product_names):
    row=c.execute("SELECT id,base_price FROM products WHERE name=?",(name,)).fetchone()
    price=prices.get(name,float(row["base_price"]) if row else 0)
    if row:c.execute("UPDATE products SET base_price=?,active=1 WHERE id=?",(price,row["id"]))
    else:c.execute("INSERT INTO products(name,base_price) VALUES(?,?)",(name,price))

# Ensure every recipe component exists in inventory.
for product,component,qty,unit in recipe_rows:
    c.execute("""INSERT OR IGNORE INTO inventory(item_name,unit,balance,unit_cost,reorder_level,purchase_unit,package_qty,package_price)
                 VALUES(?,?,0,0,0,?,1,0)""",(component,unit,""))

c.execute("DELETE FROM recipes")
recipe_count=0
for product,component,qty,unit in recipe_rows:
    p=c.execute("SELECT id FROM products WHERE name=?",(product,)).fetchone()
    i=c.execute("SELECT id FROM inventory WHERE item_name=?",(component,)).fetchone()
    if p and i:
        c.execute("INSERT INTO recipes(product_id,component_id,qty) VALUES(?,?,?)",(p["id"],i["id"],qty))
        recipe_count+=1

c.commit();c.close();wb.close()
print("======================================")
print("تم استيراد بيانات WOODS بنجاح")
print(f"المخزن: {inventory_count} صنف")
print(f"المنتجات: {len(product_names)} صنف")
print(f"مكونات الريسبي: {recipe_count} سطر")
print("======================================")
