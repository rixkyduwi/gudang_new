# admin_sales.py
from . import render_pjax
from flask_jwt_extended import jwt_required
from flask import Blueprint, request, render_template, jsonify, g

admin_sales_bp = Blueprint("admin_sales", __name__, url_prefix="/admin")

def fetch(sql, params=()):
    g.con.execute(sql, params)
    cols = [c[0] for c in g.con.description]
    return [dict(zip(cols, r)) for r in g.con.fetchall()]

# ---------- ASSIGN PELANGGAN PER SALESPERSON ----------
@admin_sales_bp.route("/salespersons/assign")
@jwt_required()
def page_assign_customers():
    # dropdown sales, dropdown customers
    sales = fetch("SELECT id, name FROM salespersons WHERE is_active=1 ORDER BY name")
    customers = fetch("SELECT id, name, address FROM customers ORDER BY name")
    return render_pjax("admin/salespersons_assign.html", sales=sales, customers=customers)
                           
@admin_sales_bp.route("/salespersons/<int:sp_id>/customers")
@jwt_required()
def list_sp_customers(sp_id):
    rows = fetch("""
      SELECT sc.customer_id, c.name, c.address
      FROM salesperson_customers sc
      JOIN customers c ON c.id = sc.customer_id
      WHERE sc.salesperson_id=%s
      ORDER BY c.name
    """, (sp_id,))
    return jsonify(rows)

@admin_sales_bp.route("/salespersons/<int:sp_id>/customers", methods=["POST"])
@jwt_required()
def assign_sp_customers(sp_id):
    payload = request.get_json() or {}
    customer_ids = payload.get("customer_ids", [])
    if not customer_ids:
        return jsonify({"error":"customer_ids wajib"}), 400
    try:
        for cid in customer_ids:
            g.con.execute("""
              INSERT IGNORE INTO salesperson_customers (salesperson_id, customer_id)
              VALUES (%s,%s)
            """, (sp_id, cid))
        g.con.connection.commit()
        return jsonify({"msg":"SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@admin_sales_bp.route("/salespersons/<int:sp_id>/customers/<int:cid>", methods=["DELETE"])
@jwt_required()
def unassign_sp_customer(sp_id, cid):
    try:
        g.con.execute("""
          DELETE FROM salesperson_customers
          WHERE salesperson_id=%s AND customer_id=%s
        """, (sp_id, cid))
        g.con.connection.commit()
        return jsonify({"msg":"SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

# ---------- SALES TARGETS (CRUD + BULK) ----------

@admin_sales_bp.route("/sales_targets")
@jwt_required()
def page_sales_targets():
    sales = fetch("SELECT id, name FROM salespersons WHERE is_active=1 ORDER BY name")
    return render_pjax("admin/sales_targets.html", sales=sales)

@admin_sales_bp.route("/sales_targets/data")
@jwt_required()
def list_targets():
    sp = request.args.get('salesperson_id', type=int)
    y  = request.args.get('year', type=int)
    where, params = [], []
    if sp: where.append("t.salesperson_id=%s"); params.append(sp)
    if y:  where.append("t.year=%s"); params.append(y)
    w = "WHERE " + " AND ".join(where) if where else ""
    rows = fetch(f"""
      SELECT t.id, t.salesperson_id, s.name AS salesperson_name,
             t.month, t.year, t.target_amount
      FROM sales_targets t
      JOIN salespersons s ON s.id = t.salesperson_id
      {w}
      ORDER BY t.year DESC, t.month ASC, s.name ASC
    """, params)
    return jsonify(rows)

@admin_sales_bp.route("/sales_targets", methods=["POST"])
@jwt_required()
def upsert_target():
    d = request.get_json() or {}
    sp, month, year, amount = d.get('salesperson_id'), d.get('month'), d.get('year'), d.get('target_amount')
    if not all([sp, month, year, amount]): return jsonify({"error":"salesperson_id, month, year, target_amount wajib"}), 400
    m = int(month); y = int(year)
    if m < 1 or m > 12: return jsonify({"error":"month harus 1-12"}), 400
    try:
        g.con.execute("""
          INSERT INTO sales_targets (salesperson_id, month, year, target_amount)
          VALUES (%s,%s,%s,%s)
          ON DUPLICATE KEY UPDATE target_amount=VALUES(target_amount)
        """, (sp, m, y, amount))
        g.con.connection.commit()
        return jsonify({"msg":"SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@admin_sales_bp.route("/sales_targets/<int:id>", methods=["DELETE"])
@jwt_required()
def del_target(id):
    try:
        g.con.execute("DELETE FROM sales_targets WHERE id=%s", (id,))
        g.con.connection.commit()
        return jsonify({"msg":"SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@admin_sales_bp.route("/sales_targets/bulk", methods=["POST"])
@jwt_required()
def bulk_targets():
    d = request.get_json() or {}
    sp, year, base = d.get('salesperson_id'), d.get('year'), d.get('base_amount', 0)
    months = d.get('months')  # optional dict {1: amt, 2: amt, ...}
    if not sp or not year:
        return jsonify({"error":"salesperson_id & year wajib"}), 400
    try:
        g.con.execute("START TRANSACTION")
        for m in range(1,13):
            amt = (months or {}).get(str(m)) or (months or {}).get(m) or base or 0
            g.con.execute("""
              INSERT INTO sales_targets (salesperson_id, month, year, target_amount)
              VALUES (%s,%s,%s,%s)
              ON DUPLICATE KEY UPDATE target_amount=VALUES(target_amount)
            """, (sp, m, int(year), amt))
        g.con.connection.commit()
        return jsonify({"msg":"SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500
