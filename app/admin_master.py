# admin_master.py
from . import render_pjax
from flask_jwt_extended import jwt_required
from flask import Blueprint, request, render_template, jsonify, g
from math import ceil

admin_master_bp = Blueprint("admin_master", __name__, url_prefix="/admin")

# ---------- UTIL ----------
def fetch(sql, params=()):
    g.con.execute(sql, params)
    cols = [c[0] for c in g.con.description]
    return [dict(zip(cols, r)) for r in g.con.fetchall()]

def one(sql, params=()):
    g.con.execute(sql, params)
    return g.con.fetchone()

def paginate(total, page, per_page):
    total_pages = max(1, ceil(total / per_page))
    return {
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": total_pages,
        "has_prev": page > 1,
        "has_next": page < total_pages,
        "page_range": range(max(1, page-3), min(total_pages+1, page+3))
    }

# ---------- GENERIC CONFIG ----------
# definisi entity master: table, columns, search_fields, order_by
ENTITIES = {
    "products": {
        "table": "products",
        "header_name": "barang",
        "columns": ["code", "name", "unit", "stock_min"],   # editable cols
        "columns_name": ["kode", "nama barang", "unit", "stock_min"],   # editable cols
        "id": "id",
        "search": ["code", "name"],
        "order_by": "id ASC, name ASC, code ASC",
        "uniques": {"code": "Kode barang sudah ada"},
        "soft_delete": True,  # arsipkan, bukan delete fisik
        "extra_select": "is_active"  # untuk toggle
    },
    "suppliers": {
        "table": "suppliers",
        "header_name": "suppliers",
        "columns": ["name", "address", "phone"],
        "columns_name": ["nama", "alamat", "telepon"],
        "id": "id",
        "search": ["name", "phone"],
        "order_by": "id ASC"
    },
    "customers": {
        "table": "customers",
        "header_name": "outlet",
        "columns": ["name", "address", "npwp"],
        "columns_name": ["nama", "alamat", "npwp"],
        "id": "id",
        "search": ["name", "npwp"],
        "order_by": "id ASC"
    },
    "salespersons": {
        "table": "salespersons",
        "header_name": "sales",
        "columns": ["name", "is_active"],
        "columns_name": ["name", "status"],
        "id": "id",
        "search": ["name"],
        "order_by": "id ASC"
    },
    "senders": {
        "table": "senders",
        "header_name": "pengirim",
        "columns": ["name", "is_active"],
        "columns_name": ["name", "status"],
        "id": "id",
        "search": ["name"],
        "order_by": "id ASC"
    }
}

# ---------- ROUTES ----------
@admin_master_bp.route("/<entity>")
@jwt_required()
def list_entity(entity):
    cfg = ENTITIES.get(entity)
    if not cfg:
        return "Unknown entity", 404

    page     = request.args.get("page", 1, type=int)
    per_page = request.args.get("per_page", 10, type=int)
    q        = request.args.get("q", "", type=str)

    where, params = [], []
    if q and cfg["search"]:
        likes = []
        for f in cfg["search"]:
            likes.append(f"{f} LIKE %s")
            params.append(f"%{q}%")
        where.append("(" + " OR ".join(likes) + ")")

    # hide archived products
    if cfg.get("soft_delete") and entity == "products":
        where.append("is_active = 1")

    where_clause = " WHERE " + " AND ".join(where) if where else ""
    count_sql = f"SELECT COUNT(*) FROM {cfg['table']}{where_clause}"
    g.con.execute(count_sql, params)
    total = g.con.fetchone()[0]

    offset = (page-1)*per_page
    extra = f", {cfg['extra_select']}" if cfg.get("extra_select") else ""
    list_sql = f"""
        SELECT {cfg['id']} AS id, {", ".join(cfg["columns"])}{extra}
        FROM {cfg['table']}
        {where_clause}
        ORDER BY {cfg['order_by']}
        LIMIT %s OFFSET %s
    """
    rows = fetch(list_sql, params + [per_page, offset])

    pager = paginate(total, page, per_page)
    return render_pjax("admin/master_list.html",
                           entity=entity,
                           header_name=cfg["header_name"],
                           rows=rows,
                           columns=cfg["columns"],
                           columns_name=cfg["columns_name"],
                           q=q, **pager)

@admin_master_bp.route("/<entity>", methods=["POST"])
@jwt_required()
def create_entity(entity):
    cfg = ENTITIES.get(entity)
    if not cfg:
        return jsonify({"error": "Unknown entity"}), 404
    data = request.get_json() or {}
    print(data)
    # unique checks
    for col, errmsg in cfg.get("uniques", {}).items():
        if data.get(col):
            if one(f"SELECT 1 FROM {cfg['table']} WHERE {col}=%s", (data[col],)):
                return jsonify({"error": errmsg}), 409

    cols = []
    vals = []
    ph   = []
    for c in cfg["columns"]:
        if c in data:
            cols.append(c)
            vals.append(data[c])
            ph.append("%s")
    if not cols:
        return jsonify({"error": "No data"}), 400

    sql = f"INSERT INTO {cfg['table']} ({', '.join(cols)}) VALUES ({', '.join(ph)})"
    try:
        g.con.execute(sql, tuple(vals))
        new_id = g.con.lastrowid
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES", "id": new_id})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@admin_master_bp.route("/<entity>", methods=["PUT"])
@jwt_required()
def update_entity(entity):
    cfg = ENTITIES.get(entity)
    if not cfg:
        return jsonify({"error": "Unknown entity"}), 404
    data = request.get_json() or {}
    rec_id = data.get("id")
    if not rec_id:
        return jsonify({"error": "id wajib"}), 400

    # unique checks (exclude self)
    for col, errmsg in cfg.get("uniques", {}).items():
        if data.get(col):
            if one(f"SELECT 1 FROM {cfg['table']} WHERE {col}=%s AND {cfg['id']}<>%s",
                   (data[col], rec_id)):
                return jsonify({"error": errmsg}), 409

    sets, vals = [], []
    for c in cfg["columns"]:
        if c in data:
            sets.append(f"{c}=%s")
            vals.append(data[c])
    if not sets:
        return jsonify({"error": "No changes"}), 400

    sql = f"UPDATE {cfg['table']} SET {', '.join(sets)} WHERE {cfg['id']}=%s"
    vals.append(rec_id)
    try:
        g.con.execute(sql, tuple(vals))
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@admin_master_bp.route("/<entity>", methods=["DELETE"])
@jwt_required()
def delete_entity(entity):
    cfg = ENTITIES.get(entity)
    if not cfg:
        return jsonify({"error": "Unknown entity"}), 404
    data = request.get_json() or {}
    rec_id = data.get("id")
    if not rec_id:
        return jsonify({"error": "id wajib"}), 400

    try:
        if cfg.get("soft_delete") and entity == "products":
            g.con.execute("UPDATE products SET is_active=0, archived_at=NOW() WHERE id=%s", (rec_id,))
        else:
            g.con.execute(f"DELETE FROM {cfg['table']} WHERE {cfg['id']}=%s", (rec_id,))
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500


