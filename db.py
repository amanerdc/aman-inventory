from __future__ import annotations

import hashlib
from typing import Iterable, Sequence

import psycopg2
from psycopg2.extras import RealDictCursor

from config import load_db_config
from constants import DEFAULT_LOW_STOCK_LEVEL, DEFAULT_PRODUCTS


def _hash_password(password: str) -> str:
    salt = "aman_inventory_salt"
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def connect():
    cfg = load_db_config()
    if cfg.port == 0 and "://" in cfg.host:
        return psycopg2.connect(cfg.host, cursor_factory=RealDictCursor)
    return psycopg2.connect(
        host=cfg.host,
        port=cfg.port,
        dbname=cfg.dbname,
        user=cfg.user,
        password=cfg.password,
        cursor_factory=RealDictCursor,
    )


def init_db() -> None:
    conn = connect()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            business TEXT NOT NULL,
            is_admin BOOLEAN NOT NULL DEFAULT FALSE
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS user_businesses (
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            business TEXT NOT NULL,
            PRIMARY KEY (user_id, business)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS products (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            category TEXT NOT NULL,
            unit TEXT NOT NULL,
            photo_path TEXT,
            opening_stock NUMERIC NOT NULL DEFAULT 0,
            low_stock_level NUMERIC NOT NULL DEFAULT %s,
            business TEXT NOT NULL
        )
        """,
        (DEFAULT_LOW_STOCK_LEVEL,),
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS perishable_in (
            id SERIAL PRIMARY KEY,
            product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
            delivery_date DATE NOT NULL,
            expiry_date DATE,
            quantity NUMERIC NOT NULL
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS perishable_in_breakdown (
            id SERIAL PRIMARY KEY,
            in_id INTEGER NOT NULL REFERENCES perishable_in(id) ON DELETE CASCADE,
            expiry_date DATE,
            quantity NUMERIC NOT NULL
        )
        """
    )
    try:
        cur.execute("ALTER TABLE perishable_in ALTER COLUMN expiry_date DROP NOT NULL")
    except Exception:
        pass

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS perishable_out (
            id SERIAL PRIMARY KEY,
            product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
            out_date DATE NOT NULL,
            out_time TEXT NOT NULL,
            quantity NUMERIC NOT NULL
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS assets (
            id SERIAL PRIMARY KEY,
            picture_path TEXT,
            name TEXT NOT NULL,
            brand TEXT,
            model TEXT,
            specifications TEXT,
            series_number TEXT,
            acquisition_date DATE NOT NULL,
            acquisition_cost NUMERIC NOT NULL,
            delivery_cost NUMERIC,
            quantity NUMERIC NOT NULL,
            location TEXT,
            status TEXT,
            business TEXT NOT NULL,
            shop_link TEXT,
            type TEXT NOT NULL,
            inventory_type TEXT NOT NULL
        )
        """
    )
    try:
        cur.execute("ALTER TABLE assets ALTER COLUMN acquisition_date DROP NOT NULL")
    except Exception:
        pass
    try:
        cur.execute("ALTER TABLE assets ALTER COLUMN acquisition_cost DROP NOT NULL")
    except Exception:
        pass
    try:
        cur.execute("ALTER TABLE assets ADD COLUMN IF NOT EXISTS specifications TEXT")
    except Exception:
        pass
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS asset_statuses (
            id SERIAL PRIMARY KEY,
            asset_id INTEGER NOT NULL REFERENCES assets(id) ON DELETE CASCADE,
            status TEXT NOT NULL,
            quantity NUMERIC NOT NULL
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS asset_acquisitions (
            id SERIAL PRIMARY KEY,
            asset_id INTEGER NOT NULL REFERENCES assets(id) ON DELETE CASCADE,
            acquisition_date DATE NOT NULL,
            acquisition_cost NUMERIC NOT NULL,
            delivery_cost NUMERIC,
            quantity NUMERIC NOT NULL,
            shop_link TEXT
        )
        """
    )
    try:
        cur.execute("ALTER TABLE asset_acquisitions ADD COLUMN IF NOT EXISTS shop_link TEXT")
    except Exception:
        pass

    conn.commit()

    # Migrate existing users into user_businesses if missing
    cur.execute("SELECT id, business FROM users")
    users = cur.fetchall()
    for user in users:
        cur.execute(
            "SELECT 1 FROM user_businesses WHERE user_id = %s LIMIT 1",
            (user["id"],),
        )
        if cur.fetchone():
            continue
        business = (user.get("business") or "").strip()
        if business == "Both":
            for biz in ("Unica", "HDN Integrated Farm"):
                cur.execute(
                    "INSERT INTO user_businesses (user_id, business) VALUES (%s, %s)",
                    (user["id"], biz),
                )
        elif business:
            cur.execute(
                "INSERT INTO user_businesses (user_id, business) VALUES (%s, %s)",
                (user["id"], business),
            )

    cur.execute("SELECT COUNT(*) as cnt FROM users")
    if cur.fetchone()["cnt"] == 0:
        cur.execute(
            "INSERT INTO users (username, password_hash, business, is_admin) VALUES (%s, %s, %s, %s)",
            ("admin", _hash_password("admin123"), "Both", True),
        )
        cur.execute("SELECT id FROM users WHERE username = %s", ("admin",))
        admin_id = cur.fetchone()["id"]
        for biz in ("Unica", "HDN Integrated Farm"):
            cur.execute(
                "INSERT INTO user_businesses (user_id, business) VALUES (%s, %s)",
                (admin_id, biz),
            )

    cur.execute("SELECT COUNT(*) as cnt FROM products")
    if cur.fetchone()["cnt"] == 0:
        cur.executemany(
            """
            INSERT INTO products (name, category, unit, photo_path, opening_stock, low_stock_level, business)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """,
            [(name, category, "unit", None, 0, DEFAULT_LOW_STOCK_LEVEL, "Unica") for name, category in DEFAULT_PRODUCTS],
        )

    conn.commit()

    # Seed acquisition records from legacy asset fields if none exist
    cur.execute("SELECT COUNT(*) as cnt FROM asset_acquisitions")
    if cur.fetchone()["cnt"] == 0:
        cur.execute(
            """
            INSERT INTO asset_acquisitions (asset_id, acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link)
            SELECT id, acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link
            FROM assets
            WHERE acquisition_date IS NOT NULL
            """
        )
        conn.commit()
    else:
        # Backfill shop links from assets where missing
        cur.execute(
            """
            UPDATE asset_acquisitions aa
            SET shop_link = a.shop_link
            FROM assets a
            WHERE a.id = aa.asset_id AND aa.shop_link IS NULL AND a.shop_link IS NOT NULL
            """
        )
        conn.commit()
    conn.close()


def verify_user(username: str, password: str) -> tuple[bool, dict | None]:
    conn = connect()
    cur = conn.cursor()
    cur.execute("SELECT id, username, password_hash, business, is_admin FROM users WHERE username = %s", (username,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return False, None
    if row["password_hash"] == _hash_password(password):
        conn = connect()
        cur = conn.cursor()
        cur.execute(
            "SELECT business FROM user_businesses WHERE user_id = %s ORDER BY business ASC",
            (row["id"],),
        )
        biz_rows = cur.fetchall()
        conn.close()
        row["businesses"] = [b["business"] for b in biz_rows] if biz_rows else [row["business"]]
        return True, row
    return False, None


def list_users() -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute("SELECT id, username, business, is_admin FROM users ORDER BY username ASC")
    rows = cur.fetchall()
    for row in rows:
        cur.execute(
            "SELECT business FROM user_businesses WHERE user_id = %s ORDER BY business ASC",
            (row["id"],),
        )
        biz_rows = cur.fetchall()
        if biz_rows:
            row["business"] = ", ".join([b["business"] for b in biz_rows])
    conn.close()
    return rows


def add_user(username: str, password: str, businesses: Sequence[str], is_admin: bool) -> None:
    conn = connect()
    cur = conn.cursor()
    business_label = ", ".join(businesses) if businesses else ""
    cur.execute(
        "INSERT INTO users (username, password_hash, business, is_admin) VALUES (%s, %s, %s, %s)",
        (username, _hash_password(password), business_label or "Unica", is_admin),
    )
    cur.execute("SELECT id FROM users WHERE username = %s", (username,))
    user_id = cur.fetchone()["id"]
    for biz in businesses:
        cur.execute(
            "INSERT INTO user_businesses (user_id, business) VALUES (%s, %s)",
            (user_id, biz),
        )
    conn.commit()
    conn.close()


def update_user(user_id: int, password: str | None, businesses: Sequence[str], is_admin: bool) -> None:
    conn = connect()
    cur = conn.cursor()
    business_label = ", ".join(businesses) if businesses else ""
    if password:
        cur.execute(
            "UPDATE users SET password_hash=%s, business=%s, is_admin=%s WHERE id=%s",
            (_hash_password(password), business_label or "Unica", is_admin, user_id),
        )
    else:
        cur.execute(
            "UPDATE users SET business=%s, is_admin=%s WHERE id=%s",
            (business_label or "Unica", is_admin, user_id),
        )
    cur.execute("DELETE FROM user_businesses WHERE user_id = %s", (user_id,))
    for biz in businesses:
        cur.execute(
            "INSERT INTO user_businesses (user_id, business) VALUES (%s, %s)",
            (user_id, biz),
        )
    conn.commit()
    conn.close()


def delete_user(user_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id=%s", (user_id,))
    conn.commit()
    conn.close()


def list_products(business: str, search: str | None = None) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business]
    query = "SELECT * FROM products WHERE business = %s"
    if search:
        query += " AND (CAST(id AS TEXT) ILIKE %s OR name ILIKE %s OR category ILIKE %s OR unit ILIKE %s)"
        like = f"%{search}%"
        params.extend([like, like, like, like])
    cur.execute(query + " ORDER BY name ASC", params)
    rows = cur.fetchall()
    conn.close()
    return rows


def add_product(
    name: str,
    category: str,
    unit: str,
    opening_stock: float,
    photo_path: str | None,
    low_stock_level: float,
    business: str,
) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO products (name, category, unit, photo_path, opening_stock, low_stock_level, business)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        """,
        (name, category, unit, photo_path, opening_stock, low_stock_level, business),
    )
    conn.commit()
    conn.close()


def update_product(
    product_id: int,
    name: str,
    category: str,
    unit: str,
    opening_stock: float,
    photo_path: str | None,
    low_stock_level: float,
) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE products
        SET name=%s, category=%s, unit=%s, opening_stock=%s, photo_path=%s, low_stock_level=%s
        WHERE id=%s
        """,
        (name, category, unit, opening_stock, photo_path, low_stock_level, product_id),
    )
    conn.commit()
    conn.close()


def delete_product(product_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM products WHERE id=%s", (product_id,))
    conn.commit()
    conn.close()


def record_in(product_id: int, delivery_date: str, quantity: float) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO perishable_in (product_id, delivery_date, expiry_date, quantity) VALUES (%s, %s, %s, %s)",
        (product_id, delivery_date, None, quantity),
    )
    conn.commit()
    conn.close()


def update_in_log(log_id: int, delivery_date: str, quantity: float) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "UPDATE perishable_in SET delivery_date=%s, quantity=%s WHERE id=%s",
        (delivery_date, quantity, log_id),
    )
    conn.commit()
    conn.close()


def delete_in_log(log_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM perishable_in WHERE id=%s", (log_id,))
    conn.commit()
    conn.close()


def record_out(product_id: int, out_date: str, out_time: str, quantity: float) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO perishable_out (product_id, out_date, out_time, quantity) VALUES (%s, %s, %s, %s)",
        (product_id, out_date, out_time, quantity),
    )
    conn.commit()
    conn.close()


def update_out_log(log_id: int, out_date: str, out_time: str, quantity: float) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "UPDATE perishable_out SET out_date=%s, out_time=%s, quantity=%s WHERE id=%s",
        (out_date, out_time, quantity, log_id),
    )
    conn.commit()
    conn.close()


def delete_out_log(log_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM perishable_out WHERE id=%s", (log_id,))
    conn.commit()
    conn.close()


def get_perishable_stock(business: str, search: str | None = None, category: str | None = None) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business]
    query = """
        SELECT
            p.id,
            p.name,
            p.category,
            p.unit,
            p.opening_stock,
            p.low_stock_level,
            p.photo_path,
            COALESCE((SELECT SUM(quantity) FROM perishable_in i WHERE i.product_id = p.id), 0) as in_qty,
            COALESCE((SELECT SUM(quantity) FROM perishable_out o WHERE o.product_id = p.id), 0) as out_qty,
            (
                SELECT MIN(b.expiry_date)
                FROM perishable_in_breakdown b
                JOIN perishable_in i ON i.id = b.in_id
                WHERE i.product_id = p.id AND b.expiry_date IS NOT NULL
            ) as next_expiry,
            COALESCE(
                (
                    SELECT SUM(b.quantity)
                    FROM perishable_in_breakdown b
                    JOIN perishable_in i ON i.id = b.in_id
                    WHERE i.product_id = p.id AND b.expiry_date IS NOT NULL AND b.expiry_date <= CURRENT_DATE + 3
                ),
                0
            ) as expiring_3_qty,
            COALESCE(
                (
                    SELECT SUM(b.quantity)
                    FROM perishable_in_breakdown b
                    JOIN perishable_in i ON i.id = b.in_id
                    WHERE i.product_id = p.id AND b.expiry_date IS NOT NULL AND b.expiry_date <= CURRENT_DATE + 7
                ),
                0
            ) as expiring_7_qty
        FROM products p
        WHERE p.business = %s
    """
    if search:
        query += " AND (CAST(p.id AS TEXT) ILIKE %s OR p.name ILIKE %s OR p.category ILIKE %s OR p.unit ILIKE %s)"
        like = f"%{search}%"
        params.extend([like, like, like, like])
    if category:
        query += " AND p.category ILIKE %s"
        params.append(f"%{category}%")
    query += " ORDER BY p.name ASC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def get_perishable_report(business: str, start_date: str, end_date: str) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            p.id as product_id,
            p.name,
            p.category,
            p.unit,
            COALESCE((SELECT SUM(quantity) FROM perishable_in i WHERE i.product_id = p.id AND i.delivery_date BETWEEN %s AND %s), 0) as in_qty,
            COALESCE((SELECT SUM(quantity) FROM perishable_out o WHERE o.product_id = p.id AND o.out_date BETWEEN %s AND %s), 0) as out_qty
        FROM products p
        WHERE p.business = %s
        ORDER BY p.category ASC, p.name ASC
        """,
        (start_date, end_date, start_date, end_date, business),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def list_in_out_logs(kind: str, product_id: int) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    if kind == "in":
        cur.execute(
            "SELECT id, delivery_date, quantity FROM perishable_in WHERE product_id = %s ORDER BY delivery_date DESC",
            (product_id,),
        )
    else:
        cur.execute(
            "SELECT id, out_date, out_time, quantity FROM perishable_out WHERE product_id = %s ORDER BY out_date DESC",
            (product_id,),
        )
    rows = cur.fetchall()
    conn.close()
    return rows


def list_expiry_dates(product_id: int) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT b.id, i.delivery_date, b.expiry_date, b.quantity
        FROM perishable_in_breakdown b
        JOIN perishable_in i ON i.id = b.in_id
        WHERE i.product_id = %s
        ORDER BY b.expiry_date ASC NULLS LAST, i.delivery_date DESC
        """,
        (product_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def list_in_breakdown(in_id: int) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, expiry_date, quantity
        FROM perishable_in_breakdown
        WHERE in_id = %s
        ORDER BY expiry_date ASC NULLS LAST
        """,
        (in_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def add_in_breakdown(in_id: int, expiry_date: str | None, quantity: float) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO perishable_in_breakdown (in_id, expiry_date, quantity) VALUES (%s, %s, %s)",
        (in_id, expiry_date, quantity),
    )
    conn.commit()
    conn.close()


def delete_in_breakdown(breakdown_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM perishable_in_breakdown WHERE id=%s", (breakdown_id,))
    conn.commit()
    conn.close()


def list_assets(
    business: str,
    inventory_type: str,
    search: str | None = None,
    type_filter: str | None = None,
) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business, inventory_type]
    query = """
        SELECT
            a.*,
            (
                SELECT MAX(acquisition_date)
                FROM asset_acquisitions aa
                WHERE aa.asset_id = a.id
            ) as latest_acquisition_date,
            COALESCE(
                (
                    SELECT SUM(quantity)
                    FROM asset_acquisitions aa
                    WHERE aa.asset_id = a.id
                ),
                0
            ) as total_acquired_qty,
            COALESCE(
                (
                    SELECT SUM(acquisition_cost * quantity)
                    FROM asset_acquisitions aa
                    WHERE aa.asset_id = a.id
                ),
                0
            ) as total_spent
        FROM assets a
        WHERE a.business = %s AND a.inventory_type = %s
    """
    if search:
        query += (
            " AND (CAST(id AS TEXT) ILIKE %s OR name ILIKE %s OR brand ILIKE %s OR model ILIKE %s "
            "OR specifications ILIKE %s OR series_number ILIKE %s OR location ILIKE %s OR shop_link ILIKE %s)"
        )
        like = f"%{search}%"
        params.extend([like, like, like, like, like, like, like, like])
    if type_filter:
        query += " AND type = %s"
        params.append(type_filter)
    query += " ORDER BY name ASC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def list_asset_statuses(asset_id: int) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, status, quantity
        FROM asset_statuses
        WHERE asset_id = %s
        ORDER BY status ASC
        """,
        (asset_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def list_asset_statuses_report(business: str, inventory_type: str) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT a.id as asset_id, a.name, a.type, s.status, s.quantity
        FROM assets a
        JOIN asset_statuses s ON s.asset_id = a.id
        WHERE a.business = %s AND a.inventory_type = %s
        ORDER BY a.name ASC, s.status ASC
        """,
        (business, inventory_type),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def add_asset_status(asset_id: int, status: str, quantity: float) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO asset_statuses (asset_id, status, quantity) VALUES (%s, %s, %s)",
        (asset_id, status, quantity),
    )
    conn.commit()
    conn.close()


def delete_asset_status(status_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM asset_statuses WHERE id=%s", (status_id,))
    conn.commit()
    conn.close()


def list_asset_acquisitions(asset_id: int) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link
        FROM asset_acquisitions
        WHERE asset_id = %s
        ORDER BY acquisition_date DESC, id DESC
        """,
        (asset_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def add_asset_acquisition(
    asset_id: int,
    acquisition_date: str,
    acquisition_cost: float,
    delivery_cost: float | None,
    quantity: float,
    shop_link: str | None,
) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO asset_acquisitions (asset_id, acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link)
        VALUES (%s, %s, %s, %s, %s, %s)
        """,
        (asset_id, acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link),
    )
    conn.commit()
    conn.close()


def update_asset_acquisition(
    acquisition_id: int,
    acquisition_date: str,
    acquisition_cost: float,
    delivery_cost: float | None,
    quantity: float,
    shop_link: str | None,
) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE asset_acquisitions
        SET acquisition_date=%s, acquisition_cost=%s, delivery_cost=%s, quantity=%s, shop_link=%s
        WHERE id=%s
        """,
        (acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link, acquisition_id),
    )
    conn.commit()
    conn.close()


def delete_asset_acquisition(acquisition_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM asset_acquisitions WHERE id=%s", (acquisition_id,))
    conn.commit()
    conn.close()


def list_asset_acquisitions_report(
    business: str,
    inventory_type: str,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business, inventory_type]
    query = """
        SELECT
            a.id as asset_id,
            a.name,
            a.type,
            aa.acquisition_date,
            aa.acquisition_cost,
            aa.delivery_cost,
            aa.quantity,
            aa.shop_link
        FROM assets a
        JOIN asset_acquisitions aa ON aa.asset_id = a.id
        WHERE a.business = %s AND a.inventory_type = %s
    """
    if start_date and end_date:
        query += " AND aa.acquisition_date BETWEEN %s AND %s"
        params.extend([start_date, end_date])
    query += " ORDER BY a.name ASC, aa.acquisition_date DESC, aa.id DESC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def add_asset(
    picture_path: str | None,
    name: str,
    brand: str | None,
    model: str | None,
    specifications: str | None,
    series_number: str | None,
    quantity: float,
    location: str | None,
    status: str | None,
    business: str,
    type_: str,
    inventory_type: str,
) -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO assets (
            picture_path, name, brand, model, specifications, series_number, acquisition_date,
            acquisition_cost, delivery_cost, quantity, location, status,
            business, shop_link, type, inventory_type
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id
        """,
        (
            picture_path,
            name,
            brand,
            model,
            specifications,
            series_number,
            None,
            None,
            None,
            quantity,
            location,
            status,
            business,
            None,
            type_,
            inventory_type,
        ),
    )
    asset_id = cur.fetchone()["id"]
    conn.commit()
    conn.close()
    return asset_id


def update_asset(
    asset_id: int,
    picture_path: str | None,
    name: str,
    brand: str | None,
    model: str | None,
    specifications: str | None,
    series_number: str | None,
    quantity: float,
    location: str | None,
    status: str | None,
    type_: str,
) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE assets
        SET picture_path=%s, name=%s, brand=%s, model=%s, specifications=%s, series_number=%s,
            quantity=%s, location=%s, status=%s, type=%s
        WHERE id=%s
        """,
        (
            picture_path,
            name,
            brand,
            model,
            specifications,
            series_number,
            quantity,
            location,
            status,
            type_,
            asset_id,
        ),
    )
    conn.commit()
    conn.close()


def duplicate_asset(asset_id: int) -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute("SELECT * FROM assets WHERE id = %s", (asset_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return 0
    new_name = f"{row.get('name') or ''} (copy)".strip()
    cur.execute(
        """
        INSERT INTO assets (
            picture_path, name, brand, model, specifications, series_number, acquisition_date,
            acquisition_cost, delivery_cost, quantity, location, status,
            business, shop_link, type, inventory_type
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id
        """,
        (
            row.get("picture_path"),
            new_name,
            row.get("brand"),
            row.get("model"),
            row.get("specifications"),
            row.get("series_number"),
            None,
            None,
            None,
            row.get("quantity"),
            row.get("location"),
            row.get("status"),
            row.get("business"),
            None,
            row.get("type"),
            row.get("inventory_type"),
        ),
    )
    new_id = cur.fetchone()["id"]

    cur.execute("SELECT status, quantity FROM asset_statuses WHERE asset_id = %s", (asset_id,))
    status_rows = cur.fetchall()
    if status_rows:
        cur.executemany(
            "INSERT INTO asset_statuses (asset_id, status, quantity) VALUES (%s, %s, %s)",
            [(new_id, r["status"], r["quantity"]) for r in status_rows],
        )

    cur.execute(
        "SELECT acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link FROM asset_acquisitions WHERE asset_id = %s",
        (asset_id,),
    )
    acq_rows = cur.fetchall()
    if acq_rows:
        cur.executemany(
            """
            INSERT INTO asset_acquisitions (asset_id, acquisition_date, acquisition_cost, delivery_cost, quantity, shop_link)
            VALUES (%s, %s, %s, %s, %s, %s)
            """,
            [
                (
                    new_id,
                    r["acquisition_date"],
                    r["acquisition_cost"],
                    r.get("delivery_cost"),
                    r["quantity"],
                    r.get("shop_link"),
                )
                for r in acq_rows
            ],
        )

    conn.commit()
    conn.close()
    return new_id


def delete_asset(asset_id: int) -> None:
    conn = connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM assets WHERE id=%s", (asset_id,))
    conn.commit()
    conn.close()


def get_assets_summary(business: str, inventory_type: str) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT MIN(picture_path) as picture_path, name, type, status, SUM(quantity) as total_quantity
        FROM assets
        WHERE business = %s AND inventory_type = %s
        GROUP BY name, type, status
        ORDER BY name ASC
        """,
        (business, inventory_type),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def get_assets_summary_range(business: str, inventory_type: str, start_date: str, end_date: str) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT MIN(picture_path) as picture_path, name, type, status, SUM(quantity) as total_quantity
        FROM assets
        WHERE business = %s AND inventory_type = %s AND acquisition_date BETWEEN %s AND %s
        GROUP BY name, type, status
        ORDER BY name ASC
        """,
        (business, inventory_type, start_date, end_date),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def list_assets_for_export(
    business: str,
    inventory_type: str,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            a.*,
            COALESCE(
                (
                    SELECT SUM(acquisition_cost * quantity)
                    FROM asset_acquisitions aa
                    WHERE aa.asset_id = a.id
                ),
                0
            ) as total_spent
        FROM assets a
        WHERE a.business = %s AND a.inventory_type = %s
        ORDER BY a.type ASC, a.name ASC, a.id ASC
        """,
        (business, inventory_type),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def list_expiry_dates_report(
    business: str,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business]
    query = """
        SELECT p.id as product_id, p.name, i.delivery_date, b.expiry_date, b.quantity
        FROM perishable_in_breakdown b
        JOIN perishable_in i ON i.id = b.in_id
        JOIN products p ON p.id = i.product_id
        WHERE p.business = %s
    """
    if start_date and end_date:
        query += " AND b.expiry_date BETWEEN %s AND %s"
        params.extend([start_date, end_date])
    query += " ORDER BY p.name ASC, b.expiry_date ASC NULLS LAST, i.delivery_date DESC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def list_in_logs_report(
    business: str,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business]
    query = """
        SELECT p.id as product_id, p.name, i.delivery_date, i.quantity
        FROM perishable_in i
        JOIN products p ON p.id = i.product_id
        WHERE p.business = %s
    """
    if start_date and end_date:
        query += " AND i.delivery_date BETWEEN %s AND %s"
        params.extend([start_date, end_date])
    query += " ORDER BY p.name ASC, i.delivery_date DESC, i.id DESC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def list_out_logs_report(
    business: str,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict]:
    conn = connect()
    cur = conn.cursor()
    params: list[object] = [business]
    query = """
        SELECT p.id as product_id, p.name, o.out_date, o.out_time, o.quantity
        FROM perishable_out o
        JOIN products p ON p.id = o.product_id
        WHERE p.business = %s
    """
    if start_date and end_date:
        query += " AND o.out_date BETWEEN %s AND %s"
        params.extend([start_date, end_date])
    query += " ORDER BY p.name ASC, o.out_date DESC, o.id DESC"
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows
