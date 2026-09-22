from __future__ import annotations
from datetime import datetime
from pathlib import Path
import sqlite3
from typing import Optional
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

DB = Path(__file__).with_name("woods_cafe.db")
app = FastAPI(title="WOODS Cafe API", version="1.0.0")

def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init_db():
    c=conn()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS products(
      id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL,
      base_price REAL NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS suppliers(
      id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT NOT NULL,
      opening_balance REAL NOT NULL DEFAULT 0,
      paid REAL NOT NULL DEFAULT 0,
      notes TEXT DEFAULT ''
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
    CREATE TABLE IF NOT EXISTS sales(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL,
      product_id INTEGER NOT NULL, qty REAL NOT NULL,
      sale_type TEXT NOT NULL, base_price REAL NOT NULL,
      unit_price REAL NOT NULL, tax REAL NOT NULL,
      total REAL NOT NULL, cogs REAL NOT NULL, gross_profit REAL NOT NULL,
      FOREIGN KEY(product_id) REFERENCES products(id)
    );
    CREATE TABLE IF NOT EXISTS purchases(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL,
      supplier_id INTEGER, item_name TEXT NOT NULL, qty REAL NOT NULL,
      unit TEXT NOT NULL, package_price REAL NOT NULL, total REAL NOT NULL
    );
    CREATE TABLE IF NOT EXISTS waste(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL,
      item_name TEXT NOT NULL, qty REAL NOT NULL, unit TEXT NOT NULL,
      reason TEXT DEFAULT '', total_cost REAL NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS expenses(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL,
      item TEXT NOT NULL, category TEXT DEFAULT '', amount REAL NOT NULL,
      payment_method TEXT DEFAULT '', notes TEXT DEFAULT ''
    );
    """)
    # Demo seed only when empty; production should import from Excel.
    if c.execute("SELECT COUNT(*) n FROM products").fetchone()["n"] == 0:
        c.execute("INSERT INTO products(name,base_price) VALUES('كورتو',130)")
        c.execute("INSERT INTO products(name,base_price) VALUES('كابتشينو',140)")
        c.execute("INSERT INTO products(name,base_price) VALUES('شاي',50)")
        c.execute("INSERT INTO products(name,base_price) VALUES('قهوة',60)")
    c.commit(); c.close()

init_db()

class SaleIn(BaseModel):
    product_name: str
    qty: float = Field(gt=0)
    sale_type: str
    base_price: Optional[float] = None

class SupplierIn(BaseModel):
    code: str = ""
    name: str
    opening_balance: float = 0
    paid: float = 0
    notes: str = ""

class WasteIn(BaseModel):
    item_name: str
    qty: float = Field(gt=0)
    unit: str
    reason: str = ""

class PurchaseIn(BaseModel):
    supplier_id: Optional[int] = None
    item_name: str
    qty: float = Field(gt=0)
    unit: str
    package_price: float = Field(ge=0)

@app.get("/health")
def health(): return {"ok": True, "service": "WOODS Cafe API"}

@app.get("/api/products")
def products():
    c=conn(); rows=c.execute("SELECT * FROM products WHERE active=1 ORDER BY name").fetchall(); c.close()
    return [dict(r) for r in rows]

@app.get("/api/inventory")
def inventory():
    c=conn(); rows=c.execute("SELECT * FROM inventory ORDER BY item_name").fetchall(); c.close()
    return [dict(r) for r in rows]

@app.get("/api/suppliers")
def suppliers():
    c=conn(); rows=c.execute("SELECT * FROM suppliers ORDER BY name").fetchall(); c.close()
    return [dict(r) for r in rows]

@app.post("/api/suppliers")
def add_supplier(x: SupplierIn):
    c=conn()
    try:
        cur=c.execute("INSERT INTO suppliers(code,name,opening_balance,paid,notes) VALUES(?,?,?,?,?)",
                      (x.code,x.name,x.opening_balance,x.paid,x.notes))
        c.commit(); return {"id":cur.lastrowid, **x.model_dump()}
    except sqlite3.IntegrityError as e:
        c.rollback(); raise HTTPException(400, "كود المورد مستخدم بالفعل") from e
    finally: c.close()

@app.get("/api/recipes/{product_name}")
def recipe(product_name: str):
    c=conn()
    rows=c.execute("""SELECT i.item_name,i.unit,r.qty,i.unit_cost
                      FROM recipes r JOIN products p ON p.id=r.product_id
                      JOIN inventory i ON i.id=r.component_id
                      WHERE p.name=? ORDER BY i.item_name""",(product_name,)).fetchall()
    c.close(); return [dict(r) for r in rows]

@app.post("/api/sales")
def sale(x: SaleIn):
    c=conn()
    try:
        p=c.execute("SELECT * FROM products WHERE name=? AND active=1",(x.product_name,)).fetchone()
        if not p: raise HTTPException(404,"الصنف غير موجود")
        base=float(p["base_price"] if x.base_price is None else x.base_price)
        if x.sale_type == "أصحاب الكافيه":
            unit=round(base*0.70,2)
        elif x.sale_type == "Manager" and x.product_name in {"شاي","قهوة","قهوة فرنساوي"}:
            unit=0.0
        else:
            unit=base
        tax=round(unit*0.14,2)
        total=round((unit+tax)*x.qty,2)
        recipe_rows=c.execute("""SELECT r.qty,i.id,i.item_name,i.unit_cost,i.balance
            FROM recipes r JOIN products p ON p.id=r.product_id
            JOIN inventory i ON i.id=r.component_id WHERE p.id=?""",(p["id"],)).fetchall()
        cogs=0.0
        for r in recipe_rows:
            used=float(r["qty"])*x.qty
            if float(r["balance"]) < used:
                raise HTTPException(409,f"الرصيد غير كافٍ للصنف: {r['item_name']}")
            cogs += used*float(r["unit_cost"])
        cogs=round(cogs,2)
        gp=round(total-cogs,2)
        now=datetime.now().isoformat(timespec="seconds")
        cur=c.execute("""INSERT INTO sales(created_at,product_id,qty,sale_type,base_price,unit_price,tax,total,cogs,gross_profit)
                         VALUES(?,?,?,?,?,?,?,?,?,?)""",
                      (now,p["id"],x.qty,x.sale_type,base,unit,tax,total,cogs,gp))
        for r in recipe_rows:
            used=float(r["qty"])*x.qty
            c.execute("UPDATE inventory SET balance=balance-? WHERE id=?",(used,r["id"]))
        c.commit()
        return {"id":cur.lastrowid,"product":x.product_name,"qty":x.qty,
                "unit_price":unit,"tax":tax,"total":total,"cogs":cogs,"gross_profit":gp}
    except HTTPException:
        c.rollback(); raise
    finally: c.close()

@app.get("/api/sales")
def sales(limit:int=100):
    c=conn(); rows=c.execute("""SELECT s.*,p.name product_name FROM sales s
        JOIN products p ON p.id=s.product_id ORDER BY s.id DESC LIMIT ?""",(limit,)).fetchall(); c.close()
    return [dict(r) for r in rows]

@app.post("/api/waste")
def add_waste(x: WasteIn):
    c=conn()
    row=c.execute("SELECT id,unit_cost,balance FROM inventory WHERE item_name=?",(x.item_name,)).fetchone()
    if not row: c.close(); raise HTTPException(404,"الصنف غير موجود في المخزن")
    if row["balance"] < x.qty: c.close(); raise HTTPException(409,"الرصيد غير كافٍ")
    total=round(x.qty*row["unit_cost"],2)
    now=datetime.now().isoformat(timespec="seconds")
    c.execute("INSERT INTO waste(created_at,item_name,qty,unit,reason,total_cost) VALUES(?,?,?,?,?,?)",
              (now,x.item_name,x.qty,x.unit,x.reason,total))
    c.execute("UPDATE inventory SET balance=balance-? WHERE id=?",(x.qty,row["id"]))
    c.commit(); c.close()
    return {"ok":True,"total_cost":total}

@app.post("/api/purchases")
def add_purchase(x: PurchaseIn):
    c=conn()
    now=datetime.now().isoformat(timespec="seconds")
    total=round(x.qty*x.package_price,2)
    c.execute("INSERT INTO purchases(created_at,supplier_id,item_name,qty,unit,package_price,total) VALUES(?,?,?,?,?,?,?)",
              (now,x.supplier_id,x.item_name,x.qty,x.unit,x.package_price,total))
    row=c.execute("SELECT id FROM inventory WHERE item_name=?",(x.item_name,)).fetchone()
    if row:
        c.execute("UPDATE inventory SET balance=balance+? WHERE id=?",(x.qty,row["id"]))
    else:
        c.execute("INSERT INTO inventory(item_name,unit,balance,unit_cost) VALUES(?,?,?,?)",
                  (x.item_name,x.unit,x.qty,x.package_price))
    c.commit(); c.close()
    return {"ok":True,"total":total}
