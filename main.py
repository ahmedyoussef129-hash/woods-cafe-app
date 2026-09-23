from __future__ import annotations
from datetime import datetime
from pathlib import Path
import sqlite3, shutil
from typing import Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from openpyxl import load_workbook

DB = Path(__file__).with_name("woods_cafe.db")
XLSX = Path("/home/Ahmedofsa129/woods cafe.xlsx")
BACKUPS = Path("/home/Ahmedofsa129/woods_cafe_backups")
BACKUPS.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="WOODS Cafe API", version="2.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://ahmedyoussef129-hash.github.io"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
      unit_cost REAL NOT NULL DEFAULT 0, reorder_level REAL NOT NULL DEFAULT 0,
      purchase_unit TEXT DEFAULT '', package_qty REAL NOT NULL DEFAULT 1,
      package_price REAL NOT NULL DEFAULT 0
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
      invoice_no TEXT DEFAULT '', supplier_id INTEGER, supplier_name TEXT DEFAULT '',
      item_name TEXT NOT NULL, qty REAL NOT NULL, unit TEXT NOT NULL,
      package_price REAL NOT NULL, basic_qty REAL NOT NULL, total REAL NOT NULL,
      payment_method TEXT DEFAULT ''
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
    CREATE TABLE IF NOT EXISTS employees(
      id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT NOT NULL,
      salary REAL NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS attendance(
      id INTEGER PRIMARY KEY, employee_id INTEGER NOT NULL, work_date TEXT NOT NULL,
      check_in TEXT DEFAULT '', check_out TEXT DEFAULT '', deduction REAL NOT NULL DEFAULT 0,
      UNIQUE(employee_id,work_date)
    );
    CREATE TABLE IF NOT EXISTS inventory_counts(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, item_name TEXT NOT NULL,
      book_qty REAL NOT NULL, actual_qty REAL NOT NULL, difference REAL NOT NULL,
      reason TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS deliveries(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, platform TEXT NOT NULL,
      order_no TEXT DEFAULT '', amount REAL NOT NULL, commission REAL NOT NULL DEFAULT 0,
      notes TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS shortages(
      id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, item_name TEXT NOT NULL,
      qty REAL NOT NULL, cost_value REAL NOT NULL DEFAULT 0,
      sale_value REAL NOT NULL DEFAULT 0, employee TEXT DEFAULT '', notes TEXT DEFAULT ''
    );
    """)
    # Backward-compatible migrations for databases created by older WOODS builds.
    migrations = {
        "inventory": [
            ("purchase_unit","TEXT DEFAULT ''"),
            ("package_qty","REAL NOT NULL DEFAULT 1"),
            ("package_price","REAL NOT NULL DEFAULT 0"),
        ],
        "purchases": [
            ("invoice_no","TEXT DEFAULT ''"),
            ("supplier_id","INTEGER"),
            ("supplier_name","TEXT DEFAULT ''"),
            ("unit","TEXT NOT NULL DEFAULT ''"),
            ("package_price","REAL NOT NULL DEFAULT 0"),
            ("basic_qty","REAL NOT NULL DEFAULT 0"),
            ("total","REAL NOT NULL DEFAULT 0"),
            ("payment_method","TEXT DEFAULT ''"),
        ],
        "suppliers": [
            ("code","TEXT"),
            ("opening_balance","REAL NOT NULL DEFAULT 0"),
            ("paid","REAL NOT NULL DEFAULT 0"),
            ("notes","TEXT DEFAULT ''"),
        ],
    }
    for table, cols in migrations.items():
        existing = {row["name"] for row in c.execute(f"PRAGMA table_info({table})").fetchall()}
        for col, typ in cols:
            if col not in existing:
                try:
                    c.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
                except sqlite3.OperationalError:
                    pass
    if c.execute("SELECT COUNT(*) n FROM products").fetchone()["n"] == 0:
        for n,p in [("كورتو",130),("كابتشينو",140),("شاي",50),("قهوة",60)]:
            c.execute("INSERT INTO products(name,base_price) VALUES(?,?)",(n,p))
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

class PurchaseIn(BaseModel):
    supplier_id: Optional[int] = None
    supplier_name: str = ""
    invoice_no: str = ""
    item_name: str
    qty: float = Field(gt=0)
    purchase_unit: str
    package_price: float = Field(ge=0)
    payment_method: str = "آجل"

class InventoryIn(BaseModel):
    item_name: str
    unit: str
    purchase_unit: str = ""
    package_qty: float = Field(gt=0)
    package_price: float = Field(ge=0)
    opening_balance: float = 0
    reorder_level: float = 0

class RecipeIn(BaseModel):
    product_name: str
    component_name: str
    qty: float = Field(gt=0)
    unit: str

class WasteIn(BaseModel):
    item_name: str
    qty: float = Field(gt=0)
    unit: str
    reason: str = ""

class ExpenseIn(BaseModel):
    item: str
    category: str = ""
    amount: float = Field(gt=0)
    payment_method: str = ""
    notes: str = ""

class EmployeeIn(BaseModel):
    code: str = ""
    name: str
    salary: float = 0

class AttendanceIn(BaseModel):
    employee_id: int
    work_date: str
    check_in: str = ""
    check_out: str = ""
    deduction: float = 0

class CountIn(BaseModel):
    item_name: str
    actual_qty: float = Field(ge=0)
    reason: str = ""

class DeliveryIn(BaseModel):
    platform: str
    order_no: str = ""
    amount: float = Field(ge=0)
    commission: float = 0
    notes: str = ""


class ProductUpdate(BaseModel):
    name: str
    base_price: float = Field(ge=0)

class InventoryUpdate(BaseModel):
    item_name: str
    unit: str
    purchase_unit: str = ""
    package_qty: float = Field(gt=0)
    package_price: float = Field(ge=0)
    opening_balance: Optional[float] = None
    reorder_level: float = 0

class SupplierUpdate(BaseModel):
    code: str = ""
    name: str
    opening_balance: float = 0
    paid: float = 0
    notes: str = ""

class RecipeUpdate(BaseModel):
    component_name: str
    qty: float = Field(gt=0)
    unit: str

class WasteUpdate(BaseModel):
    item_name: str
    qty: float = Field(gt=0)
    unit: str
    reason: str = ""

class ExpenseUpdate(BaseModel):
    item: str
    category: str = ""
    amount: float = Field(gt=0)
    payment_method: str = ""
    notes: str = ""

def excel_backup():
    if not XLSX.exists(): return None
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
    dst=BACKUPS/f"woods cafe_{stamp}.xlsx"
    shutil.copy2(XLSX,dst)
    return dst

def first_blank_row(ws, start=4, cols=10, limit=10000):
    for r in range(start, min(ws.max_row+100,limit)+1):
        if all(ws.cell(r,c).value in (None,"") for c in range(1,cols+1)):
            return r
    return ws.max_row+1


def first_data_row(ws, required_blank_cols, start=4, limit=10000):
    """Find a user-entry row while ignoring formula columns prefilled in the template."""
    for r in range(start, min(ws.max_row + 100, limit) + 1):
        if all(ws.cell(r,c).value in (None,"") for c in required_blank_cols):
            return r
    return ws.max_row + 1

def excel_find_sale_row(ws, sale):
    # Match newest row by product, qty, sale type and date.
    wanted_type={"أصحاب الكافيه":"Owner","Manager":"Manager","Guest":"Guest","عادي":"Guest"}.get(sale["sale_type"], sale["sale_type"])
    for r in range(ws.max_row, 3, -1):
        if ws.cell(r,2).value != sale["product_name"]: continue
        try:
            if abs(float(ws.cell(r,3).value or 0)-float(sale["qty"])) > 1e-9: continue
        except Exception: continue
        if str(ws.cell(r,12).value or "") != wanted_type: continue
        return r
    return None

def excel_find_purchase_row(ws, row):
    for r in range(ws.max_row, 3, -1):
        if row["invoice_no"] and str(ws.cell(r,2).value or "") == str(row["invoice_no"]):
            if str(ws.cell(r,4).value or "") == str(row["item_name"]):
                return r
        if str(ws.cell(r,4).value or "") == str(row["item_name"]):
            try:
                if abs(float(ws.cell(r,5).value or 0)-float(row["qty"])) < 1e-9 and abs(float(ws.cell(r,7).value or 0)-float(row["package_price"])) < 1e-9:
                    return r
            except Exception:
                pass
    return None

def excel_clear_row(ws, r, cols):
    for c in cols:
        ws.cell(r,c).value = None

def excel_purchase(x: PurchaseIn, basic_qty: float, total: float, supplier_name: str):
    if not XLSX.exists(): return
    excel_backup()
    wb=load_workbook(XLSX)
    ws=wb["المشتريات"]
    r=first_data_row(ws,[1,2,3,4,5,6,7,10],4)
    ws.cell(r,1).value=datetime.now()
    ws.cell(r,2).value=x.invoice_no or f"APP-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    ws.cell(r,3).value=supplier_name
    ws.cell(r,4).value=x.item_name
    ws.cell(r,5).value=x.qty
    ws.cell(r,6).value=x.purchase_unit
    ws.cell(r,7).value=x.package_price
    ws.cell(r,8).value=f"=E{r}*G{r}"
    ws.cell(r,9).value=basic_qty
    ws.cell(r,10).value=x.payment_method
    wb.save(XLSX); wb.close()

def excel_supplier(x: SupplierIn):
    if not XLSX.exists(): return
    excel_backup()
    wb=load_workbook(XLSX)
    ws=wb["الموردين"]
    r=first_data_row(ws,[1,2],4)
    ws.cell(r,1).value=x.code
    ws.cell(r,2).value=x.name
    ws.cell(r,3).value=x.opening_balance
    ws.cell(r,4).value=f'=SUMIF(المشتريات!C:C,B{r},المشتريات!H:H)'
    ws.cell(r,5).value=x.paid
    ws.cell(r,6).value=f"=C{r}+D{r}-E{r}"
    ws.cell(r,7).value=f'=IF(F{r}=0,"✅ مسدد بالكامل",IF(F{r}>0,"⚠️ متبقي للمورد","🔴 مدفوع بالزيادة"))'
    wb.save(XLSX); wb.close()

def excel_inventory(x: InventoryIn):
    if not XLSX.exists(): return
    excel_backup()
    wb=load_workbook(XLSX)
    ws=wb["المخزن"]
    r=first_data_row(ws,[1],4)
    ws.cell(r,1).value=x.item_name
    ws.cell(r,2).value=x.unit
    ws.cell(r,3).value=x.purchase_unit
    ws.cell(r,4).value=x.package_qty
    ws.cell(r,5).value=x.package_price
    ws.cell(r,6).value=f'=IFERROR(((G{r}*(E{r}/D{r}))+SUMIF(المشتريات!D:D,A{r},المشتريات!H:H))/(G{r}+SUMIF(المشتريات!D:D,A{r},المشتريات!I:I)),E{r}/D{r})'
    ws.cell(r,7).value=x.opening_balance
    ws.cell(r,8).value=f'=SUMIFS(المشتريات!I:I,المشتريات!D:D,A{r})'
    ws.cell(r,9).value=f'=SUMIFS(استهلاك_المكونات!F:F,استهلاك_المكونات!C:C,A{r})'
    ws.cell(r,10).value=f'=SUMIFS(الهالك!C:C,الهالك!B:B,A{r})'
    ws.cell(r,11).value=f'=G{r}+H{r}-I{r}-J{r}'
    ws.cell(r,12).value=x.reorder_level
    ws.cell(r,13).value=f'=K{r}*F{r}'
    ws.cell(r,14).value=f'=IF(K{r}<=L{r},"⚠ طلب شراء","✅ متوفر")'
    wb.save(XLSX); wb.close()

def excel_sale(x: SaleIn, base_price: float):
    if not XLSX.exists(): return
    excel_backup()
    wb=load_workbook(XLSX)
    ws=wb["المبيعات"]
    r=first_data_row(ws,[1,2,3,12,13],4)
    sale_type={"أصحاب الكافيه":"Owner","Manager":"Manager","Guest":"Guest","عادي":"Guest"}.get(x.sale_type,x.sale_type)
    ws.cell(r,1).value=datetime.now()
    ws.cell(r,2).value=x.product_name
    ws.cell(r,3).value=x.qty
    ws.cell(r,12).value=sale_type
    ws.cell(r,13).value=base_price
    # D:K are the workbook's existing formulas; keep them intact.
    wb.save(XLSX); wb.close()

@app.get("/health")
def health():
    return {"ok":True,"service":"WOODS Cafe API","version":"2.1.0"}

@app.get("/api/ping")
def ping():
    c = conn()
    try:
        products_n = c.execute("SELECT COUNT(*) n FROM products").fetchone()["n"]
        inventory_n = c.execute("SELECT COUNT(*) n FROM inventory").fetchone()["n"]
        return {"ok": True, "products": products_n, "inventory": inventory_n}
    finally:
        c.close()

@app.get("/api/products")
def products():
    c=conn(); rows=c.execute("SELECT * FROM products WHERE active=1 ORDER BY name").fetchall(); c.close()
    return [dict(r) for r in rows]

@app.post("/api/products")
def add_product(name:str, base_price:float=0):
    c=conn()
    try:
        cur=c.execute("INSERT INTO products(name,base_price) VALUES(?,?)",(name.strip(),base_price))
        c.commit(); return {"id":cur.lastrowid,"name":name,"base_price":base_price}
    except sqlite3.IntegrityError: raise HTTPException(400,"الصنف موجود بالفعل")
    finally: c.close()

@app.get("/api/inventory")
def inventory():
    c=conn(); rows=c.execute("SELECT * FROM inventory ORDER BY item_name").fetchall(); c.close()
    return [dict(r) for r in rows]

@app.post("/api/inventory")
def add_inventory(x:InventoryIn):
    c=conn()
    try:
        cur=c.execute("""INSERT INTO inventory(item_name,unit,purchase_unit,package_qty,package_price,balance,unit_cost,reorder_level)
                         VALUES(?,?,?,?,?,?,?,?)""",
                      (x.item_name,x.unit,x.purchase_unit,x.package_qty,x.package_price,x.opening_balance,
                       x.package_price/x.package_qty if x.package_qty else 0,x.reorder_level))
        c.commit()
        try: excel_inventory(x)
        except Exception as e: print("Excel inventory warning:",e)
        return {"id":cur.lastrowid,"ok":True}
    except sqlite3.IntegrityError: raise HTTPException(400,"الصنف موجود بالفعل")
    finally: c.close()

@app.get("/api/suppliers")
def suppliers():
    c=conn()
    rows=c.execute("""SELECT s.*, COALESCE(SUM(p.total),0) purchases_total,
                     s.opening_balance+COALESCE(SUM(p.total),0)-s.paid balance_due
                     FROM suppliers s LEFT JOIN purchases p ON p.supplier_id=s.id
                     GROUP BY s.id ORDER BY s.name""").fetchall()
    c.close(); return [dict(r) for r in rows]

@app.post("/api/suppliers")
def add_supplier(x:SupplierIn):
    c=conn()
    try:
        cur=c.execute("INSERT INTO suppliers(code,name,opening_balance,paid,notes) VALUES(?,?,?,?,?)",
                      (x.code,x.name,x.opening_balance,x.paid,x.notes))
        c.commit()
        try: excel_supplier(x)
        except Exception as e: print("Excel supplier warning:",e)
        return {"id":cur.lastrowid,**x.model_dump()}
    except sqlite3.IntegrityError as e: raise HTTPException(400,"كود المورد مستخدم بالفعل") from e
    finally: c.close()

@app.get("/api/recipes/{product_name}")
def recipe(product_name:str):
    c=conn(); rows=c.execute("""SELECT r.id,i.item_name,i.unit,r.qty,i.unit_cost
        FROM recipes r JOIN products p ON p.id=r.product_id
        JOIN inventory i ON i.id=r.component_id WHERE p.name=? ORDER BY i.item_name""",(product_name,)).fetchall()
    c.close(); return [dict(r) for r in rows]

@app.post("/api/recipes")
def add_recipe(x:RecipeIn):
    c=conn()
    p=c.execute("SELECT id FROM products WHERE name=?",(x.product_name,)).fetchone()
    i=c.execute("SELECT id FROM inventory WHERE item_name=?",(x.component_name,)).fetchone()
    if not p or not i: c.close(); raise HTTPException(404,"المنتج أو المكون غير موجود")
    c.execute("INSERT INTO recipes(product_id,component_id,qty) VALUES(?,?,?)",(p["id"],i["id"],x.qty))
    c.commit(); c.close(); return {"ok":True}

@app.post("/api/sales")
def sale(x:SaleIn):
    c=conn()
    try:
        p=c.execute("SELECT * FROM products WHERE name=? AND active=1",(x.product_name,)).fetchone()
        if not p: raise HTTPException(404,"الصنف غير موجود")
        base=float(p["base_price"] if x.base_price is None else x.base_price)
        if x.sale_type=="أصحاب الكافيه": unit=round(base*.70,2)
        elif x.sale_type=="Manager" and x.product_name in {"شاي","قهوة","قهوة فرنساوي"}: unit=0.0
        else: unit=base
        tax=round(unit*.14,2); total=round((unit+tax)*x.qty,2)
        rows=c.execute("""SELECT r.qty,i.id,i.item_name,i.unit_cost,i.balance FROM recipes r
          JOIN products p ON p.id=r.product_id JOIN inventory i ON i.id=r.component_id WHERE p.id=?""",(p["id"],)).fetchall()
        cogs=0.0
        for r in rows:
            used=float(r["qty"])*x.qty
            if float(r["balance"])<used: raise HTTPException(409,f"الرصيد غير كافٍ للصنف: {r['item_name']}")
            cogs+=used*float(r["unit_cost"])
        cogs=round(cogs,2); gp=round(total-cogs,2); now=datetime.now().isoformat(timespec="seconds")
        cur=c.execute("""INSERT INTO sales(created_at,product_id,qty,sale_type,base_price,unit_price,tax,total,cogs,gross_profit)
                         VALUES(?,?,?,?,?,?,?,?,?,?)""",(now,p["id"],x.qty,x.sale_type,base,unit,tax,total,cogs,gp))
        for r in rows: c.execute("UPDATE inventory SET balance=balance-? WHERE id=?",(float(r["qty"])*x.qty,r["id"]))
        c.commit()
        try: excel_sale(x, base)
        except Exception as e: print("Excel sale warning:",e)
        return {"id":cur.lastrowid,"product":x.product_name,"qty":x.qty,"unit_price":unit,"tax":tax,"total":total,"cogs":cogs,"gross_profit":gp}
    except HTTPException: c.rollback(); raise
    finally: c.close()

@app.get("/api/sales")
def sales(limit:int=100):
    c=conn(); rows=c.execute("""SELECT s.*,p.name product_name FROM sales s JOIN products p ON p.id=s.product_id ORDER BY s.id DESC LIMIT ?""",(limit,)).fetchall(); c.close()
    return [dict(r) for r in rows]

@app.get("/api/purchases")
def purchases(limit:int=100):
    c=conn(); rows=c.execute("""SELECT p.*,s.name supplier_name FROM purchases p LEFT JOIN suppliers s ON s.id=p.supplier_id ORDER BY p.id DESC LIMIT ?""",(limit,)).fetchall(); c.close()
    return [dict(r) for r in rows]

@app.post("/api/purchases")
def add_purchase(x:PurchaseIn):
    c=conn()
    try:
        inv=c.execute("SELECT * FROM inventory WHERE item_name=?",(x.item_name,)).fetchone()
        if not inv: raise HTTPException(404,"الصنف غير موجود في المخزن. أضفه أولًا من شاشة المخزن.")
        basic_qty=x.qty*float(inv["package_qty"] or 1)
        total=round(x.qty*x.package_price,2)
        old_balance=float(inv["balance"] or 0); old_cost=float(inv["unit_cost"] or 0)
        new_cost=((old_balance*old_cost)+(basic_qty*(x.package_price/float(inv["package_qty"] or 1))))/(old_balance+basic_qty) if old_balance+basic_qty else 0
        supplier_id=x.supplier_id
        supplier_name=x.supplier_name
        if supplier_id:
            sr=c.execute("SELECT name FROM suppliers WHERE id=?",(supplier_id,)).fetchone()
            if sr: supplier_name=sr["name"]
        cur=c.execute("""INSERT INTO purchases(created_at,invoice_no,supplier_id,supplier_name,item_name,qty,unit,package_price,basic_qty,total,payment_method)
                         VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                      (datetime.now().isoformat(timespec="seconds"),x.invoice_no,supplier_id,supplier_name,x.item_name,x.qty,
                       inv["purchase_unit"] or x.purchase_unit,x.package_price,basic_qty,total,x.payment_method))
        c.execute("UPDATE inventory SET balance=balance+?,unit_cost=?,package_price=?,purchase_unit=? WHERE id=?",
                  (basic_qty,new_cost,x.package_price,x.purchase_unit,inv["id"]))
        c.commit()
        try: excel_purchase(x,basic_qty,total,supplier_name)
        except Exception as e: print("Excel purchase warning:",e)
        return {"ok":True,"id":cur.lastrowid,"basic_qty":basic_qty,"total":total,"new_unit_cost":round(new_cost,6)}
    except HTTPException: c.rollback(); raise
    finally: c.close()

@app.post("/api/waste")
def add_waste(x:WasteIn):
    c=conn()
    row=c.execute("SELECT id,unit_cost,balance FROM inventory WHERE item_name=?",(x.item_name,)).fetchone()
    if not row: c.close(); raise HTTPException(404,"الصنف غير موجود في المخزن")
    if row["balance"]<x.qty: c.close(); raise HTTPException(409,"الرصيد غير كافٍ")
    total=round(x.qty*row["unit_cost"],2)
    c.execute("INSERT INTO waste(created_at,item_name,qty,unit,reason,total_cost) VALUES(?,?,?,?,?,?)",(datetime.now().isoformat(timespec="seconds"),x.item_name,x.qty,x.unit,x.reason,total))
    c.execute("UPDATE inventory SET balance=balance-? WHERE id=?",(x.qty,row["id"])); c.commit(); c.close()
    return {"ok":True,"total_cost":total}

@app.post("/api/expenses")
def add_expense(x:ExpenseIn):
    c=conn(); cur=c.execute("INSERT INTO expenses(created_at,item,category,amount,payment_method,notes) VALUES(?,?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"),x.item,x.category,x.amount,x.payment_method,x.notes)); c.commit(); c.close()
    return {"ok":True,"id":cur.lastrowid}

@app.get("/api/expenses")
def expenses(limit:int=100):
    c=conn(); rows=c.execute("SELECT * FROM expenses ORDER BY id DESC LIMIT ?",(limit,)).fetchall(); c.close(); return [dict(r) for r in rows]

@app.get("/api/employees")
def employees():
    c=conn(); rows=c.execute("SELECT * FROM employees WHERE active=1 ORDER BY name").fetchall(); c.close(); return [dict(r) for r in rows]

@app.post("/api/employees")
def add_employee(x:EmployeeIn):
    c=conn()
    try:
        cur=c.execute("INSERT INTO employees(code,name,salary) VALUES(?,?,?)",(x.code,x.name,x.salary)); c.commit(); return {"ok":True,"id":cur.lastrowid}
    except sqlite3.IntegrityError: raise HTTPException(400,"كود الموظف مستخدم بالفعل")
    finally: c.close()

@app.post("/api/attendance")
def add_attendance(x:AttendanceIn):
    c=conn(); c.execute("""INSERT INTO attendance(employee_id,work_date,check_in,check_out,deduction)
      VALUES(?,?,?,?,?) ON CONFLICT(employee_id,work_date) DO UPDATE SET check_in=excluded.check_in,check_out=excluded.check_out,deduction=excluded.deduction""",
      (x.employee_id,x.work_date,x.check_in,x.check_out,x.deduction)); c.commit(); c.close(); return {"ok":True}

@app.post("/api/counts")
def add_count(x:CountIn):
    c=conn(); row=c.execute("SELECT balance,unit_cost FROM inventory WHERE item_name=?",(x.item_name,)).fetchone()
    if not row: c.close(); raise HTTPException(404,"الصنف غير موجود")
    diff=x.actual_qty-float(row["balance"]); value=round(diff*float(row["unit_cost"]),2)
    c.execute("INSERT INTO inventory_counts(created_at,item_name,book_qty,actual_qty,difference,reason) VALUES(?,?,?,?,?,?)",
              (datetime.now().isoformat(timespec="seconds"),x.item_name,row["balance"],x.actual_qty,diff,x.reason))
    c.execute("UPDATE inventory SET balance=? WHERE item_name=?",(x.actual_qty,x.item_name)); c.commit(); c.close()
    return {"ok":True,"difference":diff,"value":value}

@app.post("/api/deliveries")
def add_delivery(x:DeliveryIn):
    c=conn(); cur=c.execute("INSERT INTO deliveries(created_at,platform,order_no,amount,commission,notes) VALUES(?,?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"),x.platform,x.order_no,x.amount,x.commission,x.notes)); c.commit(); c.close(); return {"ok":True,"id":cur.lastrowid}


# -----------------------------
# CRUD: تعديل / حذف
# -----------------------------

def _recalc_inventory_cost_after_purchase_delete(c, inv_id, removed_basic_qty, removed_total):
    row=c.execute("SELECT balance,unit_cost FROM inventory WHERE id=?",(inv_id,)).fetchone()
    if not row: return
    new_balance=float(row["balance"])-float(removed_basic_qty)
    if new_balance < -1e-9:
        raise HTTPException(409,"لا يمكن حذف الشراء لأن الكمية تم استهلاكها بالفعل من المخزن")
    old_value=float(row["balance"])*float(row["unit_cost"])
    new_value=max(0.0, old_value-float(removed_total)/(1.0))  # temporary; replaced below
    # removed_total is package total; convert to basic-unit cost contribution
    # Use purchase's basic quantity in caller; this helper only receives total here for compatibility.
    c.execute("UPDATE inventory SET balance=? WHERE id=?",(max(0.0,new_balance),inv_id))

def _clear_excel_sale(sale_dict):
    if not XLSX.exists(): return
    wb=load_workbook(XLSX)
    try:
        ws=wb["المبيعات"]
        row=excel_find_sale_row(ws,sale_dict)
        if row:
            excel_clear_row(ws,row,[1,2,3,12,13])
            wb.save(XLSX)
    finally:
        wb.close()

def _clear_excel_purchase(purchase_dict):
    if not XLSX.exists(): return
    wb=load_workbook(XLSX)
    try:
        ws=wb["المشتريات"]
        row=excel_find_purchase_row(ws,purchase_dict)
        if row:
            excel_clear_row(ws,row,[1,2,3,4,5,6,7,9,10])
            wb.save(XLSX)
    finally:
        wb.close()

def _clear_excel_inventory(item_name):
    if not XLSX.exists(): return
    wb=load_workbook(XLSX)
    try:
        ws=wb["المخزن"]
        for r in range(4,ws.max_row+1):
            if str(ws.cell(r,1).value or "") == item_name:
                excel_clear_row(ws,r,list(range(1,15)))
                break
        wb.save(XLSX)
    finally:
        wb.close()

def _clear_excel_supplier(supplier_name):
    if not XLSX.exists(): return
    wb=load_workbook(XLSX)
    try:
        ws=wb["الموردين"]
        for r in range(4,ws.max_row+1):
            if str(ws.cell(r,2).value or "") == supplier_name:
                excel_clear_row(ws,r,list(range(1,8)))
                break
        wb.save(XLSX)
    finally:
        wb.close()

@app.put("/api/products/{product_id}")
def update_product(product_id:int, x:ProductUpdate):
    c=conn()
    try:
        old=c.execute("SELECT * FROM products WHERE id=?",(product_id,)).fetchone()
        if not old: raise HTTPException(404,"المنتج غير موجود")
        dup=c.execute("SELECT id FROM products WHERE name=? AND id<>?",(x.name,product_id)).fetchone()
        if dup: raise HTTPException(400,"اسم المنتج مستخدم بالفعل")
        c.execute("UPDATE products SET name=?,base_price=? WHERE id=?",(x.name,x.base_price,product_id))
        c.commit()
        # Keep historical Excel sales names aligned with the renamed product.
        if XLSX.exists() and old["name"] != x.name:
            excel_backup()
            wb=load_workbook(XLSX)
            try:
                ws=wb["المبيعات"]
                for r in range(4,ws.max_row+1):
                    if ws.cell(r,2).value == old["name"]:
                        ws.cell(r,2).value=x.name
                wb.save(XLSX)
            finally: wb.close()
        return {"ok":True,"id":product_id,"name":x.name,"base_price":x.base_price}
    except sqlite3.IntegrityError as e:
        c.rollback(); raise HTTPException(400,"تعذر تعديل المنتج") from e
    finally: c.close()

@app.delete("/api/products/{product_id}")
def delete_product(product_id:int):
    c=conn()
    try:
        p=c.execute("SELECT * FROM products WHERE id=?",(product_id,)).fetchone()
        if not p: raise HTTPException(404,"المنتج غير موجود")
        used=c.execute("SELECT COUNT(*) n FROM sales WHERE product_id=?",(product_id,)).fetchone()["n"]
        if used:
            # Historical transactions must remain intact: deactivate instead of destroying them.
            c.execute("UPDATE products SET active=0 WHERE id=?",(product_id,))
        else:
            c.execute("DELETE FROM recipes WHERE product_id=?",(product_id,))
            c.execute("DELETE FROM products WHERE id=?",(product_id,))
        c.commit()
        return {"ok":True,"mode":"deactivated" if used else "deleted"}
    finally: c.close()

@app.put("/api/inventory/{item_id}")
def update_inventory(item_id:int, x:InventoryUpdate):
    c=conn()
    try:
        old=c.execute("SELECT * FROM inventory WHERE id=?",(item_id,)).fetchone()
        if not old: raise HTTPException(404,"الصنف غير موجود")
        dup=c.execute("SELECT id FROM inventory WHERE item_name=? AND id<>?",(x.item_name,item_id)).fetchone()
        if dup: raise HTTPException(400,"اسم الصنف مستخدم بالفعل")
        opening=float(old["balance"]) if x.opening_balance is None else float(x.opening_balance)
        c.execute("""UPDATE inventory SET item_name=?,unit=?,purchase_unit=?,package_qty=?,package_price=?,reorder_level=? WHERE id=?""",
                  (x.item_name,x.unit,x.purchase_unit,x.package_qty,x.package_price,x.reorder_level,item_id))
        # If name changes, preserve recipe/purchase references.
        if old["item_name"] != x.item_name:
            c.execute("UPDATE purchases SET item_name=?,unit=? WHERE item_name=?",(x.item_name,x.purchase_unit,old["item_name"]))
            c.execute("UPDATE waste SET item_name=?,unit=? WHERE item_name=?",(x.item_name,x.unit,old["item_name"]))
        if x.opening_balance is not None:
            c.execute("UPDATE inventory SET balance=? WHERE id=?",(opening,item_id))
        c.commit()
        if XLSX.exists():
            excel_backup()
            wb=load_workbook(XLSX)
            try:
                ws=wb["المخزن"]
                for r in range(4,ws.max_row+1):
                    if ws.cell(r,1).value == old["item_name"]:
                        ws.cell(r,1).value=x.item_name; ws.cell(r,2).value=x.unit
                        ws.cell(r,3).value=x.purchase_unit; ws.cell(r,4).value=x.package_qty
                        ws.cell(r,5).value=x.package_price; ws.cell(r,12).value=x.reorder_level
                        if x.opening_balance is not None: ws.cell(r,7).value=opening
                        break
                wb.save(XLSX)
            finally: wb.close()
        return {"ok":True}
    except sqlite3.IntegrityError as e:
        c.rollback(); raise HTTPException(400,"تعذر تعديل الصنف") from e
    finally: c.close()

@app.delete("/api/inventory/{item_id}")
def delete_inventory(item_id:int):
    c=conn()
    try:
        row=c.execute("SELECT * FROM inventory WHERE id=?",(item_id,)).fetchone()
        if not row: raise HTTPException(404,"الصنف غير موجود")
        refs=c.execute("SELECT COUNT(*) n FROM recipes WHERE component_id=?",(item_id,)).fetchone()["n"]
        if refs: raise HTTPException(409,"لا يمكن حذف الصنف لأنه مستخدم في ريسبي. احذف مكوناته من الريسبي أولًا.")
        if abs(float(row["balance"])) > 1e-9:
            raise HTTPException(409,"لا يمكن حذف صنف له رصيد مخزن. صفّر/صحّح الرصيد أولًا.")
        c.execute("DELETE FROM inventory WHERE id=?",(item_id,)); c.commit()
        try:
            excel_backup(); _clear_excel_inventory(row["item_name"])
        except Exception as e: print("Excel inventory delete warning:",e)
        return {"ok":True}
    finally: c.close()

@app.put("/api/suppliers/{supplier_id}")
def update_supplier(supplier_id:int, x:SupplierUpdate):
    c=conn()
    try:
        old=c.execute("SELECT * FROM suppliers WHERE id=?",(supplier_id,)).fetchone()
        if not old: raise HTTPException(404,"المورد غير موجود")
        c.execute("UPDATE suppliers SET code=?,name=?,opening_balance=?,paid=?,notes=? WHERE id=?",
                  (x.code,x.name,x.opening_balance,x.paid,x.notes,supplier_id))
        c.execute("UPDATE purchases SET supplier_name=? WHERE supplier_id=?",(x.name,supplier_id))
        c.commit()
        return {"ok":True}
    except sqlite3.IntegrityError as e:
        c.rollback(); raise HTTPException(400,"كود المورد مستخدم بالفعل") from e
    finally: c.close()

@app.delete("/api/suppliers/{supplier_id}")
def delete_supplier(supplier_id:int):
    c=conn()
    try:
        row=c.execute("SELECT * FROM suppliers WHERE id=?",(supplier_id,)).fetchone()
        if not row: raise HTTPException(404,"المورد غير موجود")
        refs=c.execute("SELECT COUNT(*) n FROM purchases WHERE supplier_id=?",(supplier_id,)).fetchone()["n"]
        if refs: raise HTTPException(409,"لا يمكن حذف المورد لأن له فواتير شراء مسجلة. احذف/عدّل الفواتير أولًا.")
        c.execute("DELETE FROM suppliers WHERE id=?",(supplier_id,)); c.commit()
        try: excel_backup(); _clear_excel_supplier(row["name"])
        except Exception as e: print("Excel supplier delete warning:",e)
        return {"ok":True}
    finally: c.close()

@app.put("/api/recipes/{recipe_id}")
def update_recipe(recipe_id:int, x:RecipeUpdate):
    c=conn()
    try:
        row=c.execute("""SELECT r.*,p.name product_name,i.item_name component_name
                        FROM recipes r JOIN products p ON p.id=r.product_id
                        JOIN inventory i ON i.id=r.component_id WHERE r.id=?""",(recipe_id,)).fetchone()
        if not row: raise HTTPException(404,"مكون الريسبي غير موجود")
        comp=c.execute("SELECT id FROM inventory WHERE item_name=?",(x.component_name,)).fetchone()
        if not comp: raise HTTPException(404,"المكون الجديد غير موجود في المخزن")
        c.execute("UPDATE recipes SET component_id=?,qty=? WHERE id=?",(comp["id"],x.qty,recipe_id)); c.commit()
        return {"ok":True}
    finally: c.close()

@app.delete("/api/recipes/{recipe_id}")
def delete_recipe(recipe_id:int):
    c=conn()
    try:
        row=c.execute("SELECT id FROM recipes WHERE id=?",(recipe_id,)).fetchone()
        if not row: raise HTTPException(404,"مكون الريسبي غير موجود")
        c.execute("DELETE FROM recipes WHERE id=?",(recipe_id,)); c.commit(); return {"ok":True}
    finally: c.close()

@app.put("/api/sales/{sale_id}")
def update_sale(sale_id:int, x:SaleIn):
    # Reverse old sale, then apply the new sale in one transaction.
    c=conn()
    try:
        old=c.execute("""SELECT s.*,p.name product_name FROM sales s JOIN products p ON p.id=s.product_id WHERE s.id=?""",(sale_id,)).fetchone()
        if not old: raise HTTPException(404,"عملية البيع غير موجودة")
        old_recipe=c.execute("SELECT r.qty,i.id FROM recipes r JOIN inventory i ON i.id=r.component_id WHERE r.product_id=?",(old["product_id"],)).fetchall()
        for r in old_recipe:
            c.execute("UPDATE inventory SET balance=balance+? WHERE id=?",(float(r["qty"])*float(old["qty"]),r["id"]))
        p=c.execute("SELECT * FROM products WHERE name=? AND active=1",(x.product_name,)).fetchone()
        if not p: raise HTTPException(404,"المنتج الجديد غير موجود")
        base=float(p["base_price"] if x.base_price is None else x.base_price)
        if x.sale_type=="أصحاب الكافيه": unit=round(base*.70,2)
        elif x.sale_type=="Manager" and x.product_name in {"شاي","قهوة","قهوة فرنساوي"}: unit=0.0
        else: unit=base
        tax=round(unit*.14,2); total=round((unit+tax)*x.qty,2)
        rows=c.execute("SELECT r.qty,i.id,i.item_name,i.unit_cost,i.balance FROM recipes r JOIN inventory i ON i.id=r.component_id WHERE r.product_id=?",(p["id"],)).fetchall()
        cogs=0
        for r in rows:
            used=float(r["qty"])*x.qty
            if float(r["balance"])<used: raise HTTPException(409,f"الرصيد غير كافٍ للصنف: {r['item_name']}")
            cogs+=used*float(r["unit_cost"])
        gp=round(total-round(cogs,2),2)
        c.execute("""UPDATE sales SET product_id=?,qty=?,sale_type=?,base_price=?,unit_price=?,tax=?,total=?,cogs=?,gross_profit=? WHERE id=?""",
                  (p["id"],x.qty,x.sale_type,base,unit,tax,total,round(cogs,2),gp,sale_id))
        for r in rows: c.execute("UPDATE inventory SET balance=balance-? WHERE id=?",(float(r["qty"])*x.qty,r["id"]))
        c.commit()
        try:
            if XLSX.exists():
                excel_backup()
                wb=load_workbook(XLSX)
                try:
                    ws=wb["المبيعات"]
                    row=excel_find_sale_row(ws, {**dict(old), "product_name": old["product_name"]})
                    if row:
                        sale_type={"أصحاب الكافيه":"Owner","Manager":"Manager","Guest":"Guest","عادي":"Guest"}.get(x.sale_type,x.sale_type)
                        ws.cell(row,1).value=datetime.now()
                        ws.cell(row,2).value=x.product_name
                        ws.cell(row,3).value=x.qty
                        ws.cell(row,12).value=sale_type
                        ws.cell(row,13).value=base
                    wb.save(XLSX)
                finally: wb.close()
        except Exception as e: print("Excel sale update warning:",e)
        return {"ok":True,"total":total,"cogs":round(cogs,2),"gross_profit":gp}
    except HTTPException:
        c.rollback(); raise
    finally: c.close()

@app.delete("/api/sales/{sale_id}")
def delete_sale(sale_id:int):
    c=conn()
    try:
        old=c.execute("""SELECT s.*,p.name product_name FROM sales s JOIN products p ON p.id=s.product_id WHERE s.id=?""",(sale_id,)).fetchone()
        if not old: raise HTTPException(404,"عملية البيع غير موجودة")
        recipe=c.execute("SELECT r.qty,i.id FROM recipes r JOIN inventory i ON i.id=r.component_id WHERE r.product_id=?",(old["product_id"],)).fetchall()
        for r in recipe:
            c.execute("UPDATE inventory SET balance=balance+? WHERE id=?",(float(r["qty"])*float(old["qty"]),r["id"]))
        c.execute("DELETE FROM sales WHERE id=?",(sale_id,)); c.commit()
        try:
            excel_backup(); _clear_excel_sale({**dict(old)})
        except Exception as e: print("Excel sale delete warning:",e)
        return {"ok":True}
    finally: c.close()

@app.put("/api/purchases/{purchase_id}")
def update_purchase(purchase_id:int, x:PurchaseIn):
    c=conn()
    try:
        old=c.execute("SELECT * FROM purchases WHERE id=?",(purchase_id,)).fetchone()
        if not old: raise HTTPException(404,"فاتورة الشراء غير موجودة")
        inv_old=c.execute("SELECT * FROM inventory WHERE item_name=?",(old["item_name"],)).fetchone()
        if not inv_old: raise HTTPException(404,"صنف الشراء غير موجود")
        old_basic=float(old["basic_qty"]); old_cost=float(old["package_price"])/float(inv_old["package_qty"] or 1)
        if float(inv_old["balance"]) < old_basic-1e-9: raise HTTPException(409,"لا يمكن تعديل الشراء لأن جزءًا منه تم استهلاكه بالفعل")
        # remove old purchase
        new_balance=float(inv_old["balance"])-old_basic
        old_value=float(inv_old["balance"])*float(inv_old["unit_cost"])
        new_value=max(0.0,old_value-old_basic*old_cost)
        # resolve new item
        inv=c.execute("SELECT * FROM inventory WHERE item_name=?",(x.item_name,)).fetchone()
        if not inv: raise HTTPException(404,"الصنف الجديد غير موجود")
        new_basic=x.qty*float(inv["package_qty"] or 1)
        new_total=round(x.qty*x.package_price,2)
        if inv["id"]==inv_old["id"]:
            final_balance=new_balance+new_basic
            final_cost=(new_value+new_basic*(x.package_price/float(inv["package_qty"] or 1)))/final_balance if final_balance else 0
            c.execute("UPDATE inventory SET balance=?,unit_cost=?,package_price=?,purchase_unit=? WHERE id=?",
                      (final_balance,final_cost,x.package_price,x.purchase_unit,inv["id"]))
        else:
            c.execute("UPDATE inventory SET balance=? WHERE id=?",(new_balance,inv_old["id"]))
            final_balance=float(inv["balance"])+new_basic
            final_cost=(float(inv["balance"])*float(inv["unit_cost"])+new_basic*(x.package_price/float(inv["package_qty"] or 1)))/final_balance if final_balance else 0
            c.execute("UPDATE inventory SET balance=?,unit_cost=?,package_price=?,purchase_unit=? WHERE id=?",
                      (final_balance,final_cost,x.package_price,x.purchase_unit,inv["id"]))
        supplier_name=x.supplier_name
        if x.supplier_id:
            sr=c.execute("SELECT name FROM suppliers WHERE id=?",(x.supplier_id,)).fetchone()
            if sr: supplier_name=sr["name"]
        c.execute("""UPDATE purchases SET invoice_no=?,supplier_id=?,supplier_name=?,item_name=?,qty=?,unit=?,package_price=?,basic_qty=?,total=?,payment_method=? WHERE id=?""",
                  (x.invoice_no,x.supplier_id,supplier_name,x.item_name,x.qty,inv["purchase_unit"] or x.purchase_unit,x.package_price,new_basic,new_total,x.payment_method,purchase_id))
        c.commit()
        try:
            if XLSX.exists():
                excel_backup()
                wb=load_workbook(XLSX)
                try:
                    ws=wb["المشتريات"]
                    row=excel_find_purchase_row(ws, dict(old))
                    if row:
                        ws.cell(row,1).value=datetime.now()
                        ws.cell(row,2).value=x.invoice_no
                        ws.cell(row,3).value=supplier_name
                        ws.cell(row,4).value=x.item_name
                        ws.cell(row,5).value=x.qty
                        ws.cell(row,6).value=x.purchase_unit
                        ws.cell(row,7).value=x.package_price
                        ws.cell(row,9).value=new_basic
                        ws.cell(row,10).value=x.payment_method
                    wb.save(XLSX)
                finally: wb.close()
        except Exception as e: print("Excel purchase update warning:",e)
        return {"ok":True}
    except HTTPException:
        c.rollback(); raise
    finally: c.close()

@app.delete("/api/purchases/{purchase_id}")
def delete_purchase(purchase_id:int):
    c=conn()
    try:
        old=c.execute("SELECT * FROM purchases WHERE id=?",(purchase_id,)).fetchone()
        if not old: raise HTTPException(404,"فاتورة الشراء غير موجودة")
        inv=c.execute("SELECT * FROM inventory WHERE item_name=?",(old["item_name"],)).fetchone()
        if not inv: raise HTTPException(404,"الصنف غير موجود")
        basic=float(old["basic_qty"])
        if float(inv["balance"]) < basic-1e-9: raise HTTPException(409,"لا يمكن حذف الشراء لأن الكمية تم استهلاكها بالفعل")
        unit_cost_purchase=float(old["package_price"])/float(inv["package_qty"] or 1)
        new_balance=float(inv["balance"])-basic
        old_value=float(inv["balance"])*float(inv["unit_cost"])
        new_value=max(0.0,old_value-basic*unit_cost_purchase)
        new_cost=new_value/new_balance if new_balance else 0
        c.execute("UPDATE inventory SET balance=?,unit_cost=? WHERE id=?",(new_balance,new_cost,inv["id"]))
        c.execute("DELETE FROM purchases WHERE id=?",(purchase_id,)); c.commit()
        try:
            excel_backup(); _clear_excel_purchase(dict(old))
        except Exception as e: print("Excel purchase delete warning:",e)
        return {"ok":True}
    finally: c.close()

@app.get("/api/waste")
def waste(limit:int=100):
    c=conn(); rows=c.execute("SELECT * FROM waste ORDER BY id DESC LIMIT ?",(limit,)).fetchall(); c.close(); return [dict(r) for r in rows]

@app.delete("/api/waste/{waste_id}")
def delete_waste(waste_id:int):
    c=conn()
    try:
        row=c.execute("SELECT * FROM waste WHERE id=?",(waste_id,)).fetchone()
        if not row: raise HTTPException(404,"سجل الهالك غير موجود")
        inv=c.execute("SELECT id FROM inventory WHERE item_name=?",(row["item_name"],)).fetchone()
        if inv: c.execute("UPDATE inventory SET balance=balance+? WHERE id=?",(row["qty"],inv["id"]))
        c.execute("DELETE FROM waste WHERE id=?",(waste_id,)); c.commit(); return {"ok":True}
    finally: c.close()

@app.put("/api/expenses/{expense_id}")
def update_expense(expense_id:int, x:ExpenseUpdate):
    c=conn()
    try:
        if not c.execute("SELECT id FROM expenses WHERE id=?",(expense_id,)).fetchone(): raise HTTPException(404,"المصروف غير موجود")
        c.execute("UPDATE expenses SET item=?,category=?,amount=?,payment_method=?,notes=? WHERE id=?",
                  (x.item,x.category,x.amount,x.payment_method,x.notes,expense_id)); c.commit(); return {"ok":True}
    finally: c.close()

@app.delete("/api/expenses/{expense_id}")
def delete_expense(expense_id:int):
    c=conn()
    try:
        if not c.execute("SELECT id FROM expenses WHERE id=?",(expense_id,)).fetchone(): raise HTTPException(404,"المصروف غير موجود")
        c.execute("DELETE FROM expenses WHERE id=?",(expense_id,)); c.commit(); return {"ok":True}
    finally: c.close()

@app.get("/api/dashboard")
def dashboard():
    c=conn()
    s=c.execute("SELECT COALESCE(SUM(total),0) v,COALESCE(SUM(cogs),0)c FROM sales").fetchone()
    e=c.execute("SELECT COALESCE(SUM(amount),0) v FROM expenses").fetchone()
    p=c.execute("SELECT COALESCE(SUM(total),0) v FROM purchases").fetchone()
    w=c.execute("SELECT COALESCE(SUM(total_cost),0) v FROM waste").fetchone()
    inv=c.execute("SELECT COALESCE(SUM(balance*unit_cost),0) v FROM inventory").fetchone()
    c.close()
    return {"sales":s["v"],"cogs":s["c"],"gross_profit":s["v"]-s["c"],"expenses":e["v"],"purchases":p["v"],"waste":w["v"],"inventory_value":inv["v"],"net_profit":s["v"]-s["c"]-e["v"]-w["v"]}
