from __future__ import annotations

import base64
import csv
import json
import os
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib import error as urllib_error
from urllib import request as urllib_request

from flask import Flask, jsonify, redirect, render_template, request, session, url_for

try:
    from pymongo import MongoClient
    from pymongo.errors import PyMongoError

    PYMONGO_AVAILABLE = True
except ImportError:
    MongoClient = None
    PyMongoError = Exception
    PYMONGO_AVAILABLE = False

try:
    from openpyxl import Workbook, load_workbook

    OPENPYXL_AVAILABLE = True
except ImportError:
    Workbook = None
    load_workbook = None
    OPENPYXL_AVAILABLE = False


BASE_DIR = Path(__file__).resolve().parent


def _env_int(name: str, default: int) -> int:
    try:
        return int(str(os.getenv(name, default)).strip())
    except (TypeError, ValueError):
        return default

# Canonical columns for each persisted table.
TABLE_COLUMNS: dict[str, list[str]] = {
    "users": ["Name", "Phone Number"],
    "user_cart": ["cart_id", "product_id", "product_name", "quantity", "price", "total_cost"],
    "products_database": [
        "product_id",
        "product_name",
        "stock_quantity",
        "price",
        "last_restock_date",
        "reorder_level",
        "critical_level",
        "last_low_stock_alert_at",
        "last_low_stock_alert_level",
    ],
    "store_sales": [
        "transaction_id",
        "customer_name",
        "items_bought",
        "total_amount",
        "currency",
        "date",
        "payment_order_id",
        "payment_id",
        "status",
    ],
    "inventory_alerts": [
        "alert_id",
        "product_id",
        "product_name",
        "stock_quantity",
        "reorder_level",
        "critical_level",
        "severity",
        "status",
        "notification_channel",
        "notification_status",
        "notification_error",
        "transaction_id",
        "created_at",
        "resolved_at",
    ],
    "inventory_movements": [
        "movement_id",
        "product_id",
        "product_name",
        "movement_type",
        "quantity_change",
        "stock_before",
        "stock_after",
        "related_transaction_id",
        "created_at",
    ],
}

PRODUCT_IMPORT_COLUMNS = {
    "product_id": ("product_id", "barcode", "Barcode", "Product ID", "id"),
    "product_name": ("product_name", "Product Name", "Name", "item_name"),
    "stock_quantity": ("stock_quantity", "Stock", "stock", "quantity"),
    "price": ("price", "Price", "amount", "mrp"),
}

PAYPAL_CLIENT_ID = "AT4YjVQk1fNnJTW1sdd7KRMlB5OVYTBAuKh4dFp76BUiAmLiqbPP8VlJmrZhhZb5-w_0fRzNQG3BLTrw"
PAYPAL_CLIENT_SECRET = "EDumKbVywv9AnZGKWHwCtl9o9UkEpMO_6hVbvuxk64_bbxEu303Se5pTEJAiy0KpOy4dX2G6R7rldyPa"
PAYPAL_BASE_URL = os.getenv("PAYPAL_BASE_URL", "https://api-m.sandbox.paypal.com").strip()
PAYMENT_CURRENCY = os.getenv("PAYMENT_CURRENCY", "USD").strip().upper() or "USD"
MONGODB_URI = os.getenv("MONGODB_URI", "").strip()
MONGODB_DB_NAME = os.getenv("MONGODB_DB_NAME", "quickbill").strip() or "quickbill"
LOW_STOCK_DEFAULT_THRESHOLD = _env_int("LOW_STOCK_DEFAULT_THRESHOLD", 10)
CRITICAL_STOCK_DEFAULT_THRESHOLD = _env_int("CRITICAL_STOCK_DEFAULT_THRESHOLD", 3)
LOW_STOCK_ALERT_COOLDOWN_MINUTES = _env_int("LOW_STOCK_ALERT_COOLDOWN_MINUTES", 180)
TELEGRAM_BOT_TOKEN = "8541759344:AAFhR2fS1s8OQiOeNmgRDKj91Nc5c0jYthU"
TELEGRAM_ADMIN_CHAT_ID = "2092635206"

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "quickbill-dev-secret")

mongo_client: MongoClient | None = None
mongo_db = None


def _init_mongodb() -> None:
    global mongo_client, mongo_db
    if not PYMONGO_AVAILABLE:
        print("[db] pymongo is not installed, using file storage fallback.")
        return
    if not MONGODB_URI:
        return
    try:
        mongo_client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=2500)
        mongo_client.admin.command("ping")
        mongo_db = mongo_client[MONGODB_DB_NAME]
        print(f"[db] MongoDB connected: {MONGODB_DB_NAME}")
    except PyMongoError as exc:
        mongo_client = None
        mongo_db = None
        print(f"[db] MongoDB unavailable, using file storage fallback. Reason: {exc}")


def _using_mongodb() -> bool:
    return mongo_db is not None


# ---------- Generic parsing + table I/O helpers ----------
def _table_paths(table_name: str) -> tuple[Path, Path]:
    return BASE_DIR / f"{table_name}.xlsx", BASE_DIR / f"{table_name}.csv"


def _normalize_field_name(name: str) -> str:
    return "".join(ch for ch in str(name).lower() if ch.isalnum())


def _find_value(row: dict[str, Any], candidates: tuple[str, ...], default: Any = "") -> Any:
    key_map = {_normalize_field_name(key): value for key, value in row.items()}
    for candidate in candidates:
        value = key_map.get(_normalize_field_name(candidate))
        if value not in ("", None):
            return value
    return default


def _to_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _to_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def _read_rows_csv(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8", newline="") as file:
        reader = csv.DictReader(file)
        return [dict(row) for row in reader]


def _write_rows_csv(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in columns})


def _read_rows_xlsx(path: Path) -> list[dict[str, Any]]:
    if not OPENPYXL_AVAILABLE or load_workbook is None or not path.exists():
        return []

    workbook = load_workbook(path)
    sheet = workbook.active
    iterator = sheet.iter_rows(values_only=True)
    header_row = next(iterator, None)

    if not header_row:
        return []

    headers = [str(cell).strip() if cell is not None else "" for cell in header_row]
    rows: list[dict[str, Any]] = []
    for values in iterator:
        if not values or not any(cell not in (None, "") for cell in values):
            continue
        row: dict[str, Any] = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            row[header] = values[index] if index < len(values) else ""
        rows.append(row)
    return rows


def _write_rows_xlsx(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    if not OPENPYXL_AVAILABLE or Workbook is None:
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column, "") for column in columns])
    workbook.save(path)


def _read_table(table_name: str) -> list[dict[str, Any]]:
    if _using_mongodb():
        try:
            rows = list(mongo_db[table_name].find({}, {"_id": 0}))
            return [dict(row) for row in rows]
        except PyMongoError as exc:
            print(f"[db] read failed for {table_name}, falling back to files. Reason: {exc}")

    xlsx_path, csv_path = _table_paths(table_name)

    if xlsx_path.exists() and OPENPYXL_AVAILABLE:
        return _read_rows_xlsx(xlsx_path)

    if csv_path.exists():
        return _read_rows_csv(csv_path)

    if xlsx_path.exists():
        # xlsx exists but openpyxl is missing
        return []

    return []


def _write_table(table_name: str, rows: list[dict[str, Any]]) -> None:
    columns = TABLE_COLUMNS[table_name]

    if _using_mongodb():
        try:
            collection = mongo_db[table_name]
            collection.delete_many({})
            if rows:
                sanitized_rows = [{column: row.get(column, "") for column in columns} for row in rows]
                collection.insert_many(sanitized_rows)
            return
        except PyMongoError as exc:
            print(f"[db] write failed for {table_name}, falling back to files. Reason: {exc}")

    xlsx_path, csv_path = _table_paths(table_name)

    if OPENPYXL_AVAILABLE:
        _write_rows_xlsx(xlsx_path, rows, columns)
        return

    _write_rows_csv(csv_path, rows, columns)


def _normalize_phone(phone: str) -> str:
    return "".join(ch for ch in phone if ch.isdigit())


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _utc_now_iso() -> str:
    return _utc_now().isoformat(timespec="seconds")


def _parse_iso_datetime(value: Any) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


# ---------- Domain helpers (products/cart) ----------
def _canonical_product_row(row: dict[str, Any]) -> dict[str, Any] | None:
    product_id = str(_find_value(row, PRODUCT_IMPORT_COLUMNS["product_id"], "")).strip()
    if not product_id:
        return None

    product_name = str(_find_value(row, PRODUCT_IMPORT_COLUMNS["product_name"], "Unknown Item")).strip()
    stock_quantity = _to_int(_find_value(row, PRODUCT_IMPORT_COLUMNS["stock_quantity"], 0), 0)
    price = round(_to_float(_find_value(row, PRODUCT_IMPORT_COLUMNS["price"], 0.0), 0.0), 2)
    restock_date = str(row.get("last_restock_date", "")).strip()
    reorder_level = _to_int(row.get("reorder_level", LOW_STOCK_DEFAULT_THRESHOLD), LOW_STOCK_DEFAULT_THRESHOLD)
    critical_level = _to_int(
        row.get("critical_level", min(CRITICAL_STOCK_DEFAULT_THRESHOLD, reorder_level)),
        min(CRITICAL_STOCK_DEFAULT_THRESHOLD, reorder_level),
    )
    if reorder_level < 0:
        reorder_level = LOW_STOCK_DEFAULT_THRESHOLD
    if critical_level < 0:
        critical_level = min(CRITICAL_STOCK_DEFAULT_THRESHOLD, reorder_level)
    if critical_level > reorder_level:
        critical_level = reorder_level
    last_low_stock_alert_at = str(row.get("last_low_stock_alert_at", "")).strip()
    last_low_stock_alert_level = str(row.get("last_low_stock_alert_level", "")).strip().lower()

    return {
        "product_id": product_id,
        "product_name": product_name,
        "stock_quantity": stock_quantity,
        "price": price,
        "last_restock_date": restock_date,
        "reorder_level": reorder_level,
        "critical_level": critical_level,
        "last_low_stock_alert_at": last_low_stock_alert_at,
        "last_low_stock_alert_level": last_low_stock_alert_level,
    }


def _bootstrap_products_table() -> None:
    products_xlsx, products_csv = _table_paths("products_database")
    if OPENPYXL_AVAILABLE:
        if products_xlsx.exists() or products_csv.exists():
            return
    else:
        if products_csv.exists():
            return

    source_csv = BASE_DIR / "products.csv"
    rows: list[dict[str, Any]] = []
    if source_csv.exists():
        imported_rows = _read_rows_csv(source_csv)
        for imported_row in imported_rows:
            canonical = _canonical_product_row(imported_row)
            if canonical is not None:
                rows.append(canonical)

    _write_table("products_database", rows)


def _ensure_tables() -> None:
    for table_name in ("users", "user_cart", "store_sales", "inventory_alerts", "inventory_movements"):
        table_xlsx, table_csv = _table_paths(table_name)
        if OPENPYXL_AVAILABLE:
            if not table_xlsx.exists() and not table_csv.exists():
                _write_table(table_name, [])
        else:
            if not table_csv.exists():
                _write_table(table_name, [])
    _bootstrap_products_table()


def _load_products() -> list[dict[str, Any]]:
    raw_rows = _read_table("products_database")
    products: list[dict[str, Any]] = []
    for row in raw_rows:
        canonical = _canonical_product_row(row)
        if canonical is not None:
            products.append(canonical)

    if products:
        return products

    source_csv = BASE_DIR / "products.csv"
    if source_csv.exists():
        source_rows = _read_rows_csv(source_csv)
        for row in source_rows:
            canonical = _canonical_product_row(row)
            if canonical is not None:
                products.append(canonical)

    return products


def _save_products(products: list[dict[str, Any]]) -> None:
    rows: list[dict[str, Any]] = []
    for product in products:
        canonical = _canonical_product_row(product)
        if canonical is not None:
            rows.append(canonical)
    _write_table("products_database", rows)


def _product_lookup_key(value: Any) -> str:
    return str(value or "").strip().lower()


def _build_product_index(products: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    index: dict[str, dict[str, Any]] = {}
    for product in products:
        product_id = _product_lookup_key(product.get("product_id"))
        product_name = _product_lookup_key(product.get("product_name"))
        if product_id:
            index[product_id] = product
        if product_name:
            index[product_name] = product
    return index


def _stock_alert_severity(product: dict[str, Any]) -> str:
    stock_quantity = _to_int(product.get("stock_quantity", 0), 0)
    reorder_level = _to_int(product.get("reorder_level", LOW_STOCK_DEFAULT_THRESHOLD), LOW_STOCK_DEFAULT_THRESHOLD)
    critical_level = _to_int(
        product.get("critical_level", min(CRITICAL_STOCK_DEFAULT_THRESHOLD, reorder_level)),
        min(CRITICAL_STOCK_DEFAULT_THRESHOLD, reorder_level),
    )
    if stock_quantity <= critical_level:
        return "critical"
    if stock_quantity <= reorder_level:
        return "low"
    return ""


def _send_admin_telegram_notification(message: str) -> tuple[str, str]:
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_ADMIN_CHAT_ID:
        return "skipped", "Telegram is not configured."

    payload = json.dumps({"chat_id": TELEGRAM_ADMIN_CHAT_ID, "text": message}).encode("utf-8")
    notification_request = urllib_request.Request(
        f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib_request.urlopen(notification_request, timeout=10) as response:
            response_body = response.read().decode("utf-8")
    except urllib_error.HTTPError as exc:
        error_text = exc.read().decode("utf-8", errors="ignore")
        return "failed", error_text or str(exc.reason)
    except urllib_error.URLError as exc:
        return "failed", str(exc.reason)

    try:
        response_payload = json.loads(response_body)
    except json.JSONDecodeError:
        response_payload = {}

    if response_payload.get("ok") is False:
        description = str(response_payload.get("description", "Unknown Telegram error")).strip()
        return "failed", description

    return "sent", ""


def _resolve_open_inventory_alerts(
    alert_rows: list[dict[str, Any]],
    product_id: str,
    resolved_at: str,
) -> bool:
    changed = False
    product_key = _product_lookup_key(product_id)
    for alert in alert_rows:
        if _product_lookup_key(alert.get("product_id")) != product_key:
            continue
        if str(alert.get("status", "")).strip().lower() != "open":
            continue
        alert["status"] = "resolved"
        alert["resolved_at"] = resolved_at
        changed = True
    return changed


def _record_inventory_movement(
    movement_rows: list[dict[str, Any]],
    product: dict[str, Any],
    quantity_change: int,
    stock_before: int,
    stock_after: int,
    movement_type: str,
    transaction_id: str,
) -> None:
    movement_rows.append(
        {
            "movement_id": f"MOV-{uuid.uuid4().hex[:10].upper()}",
            "product_id": product.get("product_id", ""),
            "product_name": product.get("product_name", "Unknown Item"),
            "movement_type": movement_type,
            "quantity_change": quantity_change,
            "stock_before": stock_before,
            "stock_after": stock_after,
            "related_transaction_id": transaction_id,
            "created_at": _utc_now_iso(),
        }
    )


def _should_send_low_stock_alert(product: dict[str, Any], severity: str) -> bool:
    if not severity:
        return False

    last_level = str(product.get("last_low_stock_alert_level", "")).strip().lower()
    if severity != last_level:
        return True

    last_sent_at = _parse_iso_datetime(product.get("last_low_stock_alert_at", ""))
    if last_sent_at is None:
        return True

    cooldown = timedelta(minutes=max(LOW_STOCK_ALERT_COOLDOWN_MINUTES, 0))
    return (_utc_now() - last_sent_at) >= cooldown


def _append_low_stock_alert(
    alert_rows: list[dict[str, Any]],
    product: dict[str, Any],
    severity: str,
    transaction_id: str,
) -> dict[str, Any]:
    message = (
        f"{severity.upper()} stock alert\n"
        f"Product: {product.get('product_name', 'Unknown Item')}\n"
        f"ID: {product.get('product_id', '')}\n"
        f"Stock left: {_to_int(product.get('stock_quantity', 0), 0)}\n"
        f"Reorder level: {_to_int(product.get('reorder_level', LOW_STOCK_DEFAULT_THRESHOLD), LOW_STOCK_DEFAULT_THRESHOLD)}\n"
        f"Critical level: {_to_int(product.get('critical_level', CRITICAL_STOCK_DEFAULT_THRESHOLD), CRITICAL_STOCK_DEFAULT_THRESHOLD)}\n"
        f"Transaction: {transaction_id}"
    )
    notification_status, notification_error = _send_admin_telegram_notification(message)
    alert_row = {
        "alert_id": f"ALT-{uuid.uuid4().hex[:10].upper()}",
        "product_id": product.get("product_id", ""),
        "product_name": product.get("product_name", "Unknown Item"),
        "stock_quantity": _to_int(product.get("stock_quantity", 0), 0),
        "reorder_level": _to_int(product.get("reorder_level", LOW_STOCK_DEFAULT_THRESHOLD), LOW_STOCK_DEFAULT_THRESHOLD),
        "critical_level": _to_int(
            product.get("critical_level", CRITICAL_STOCK_DEFAULT_THRESHOLD),
            CRITICAL_STOCK_DEFAULT_THRESHOLD,
        ),
        "severity": severity,
        "status": "open",
        "notification_channel": "telegram",
        "notification_status": notification_status,
        "notification_error": notification_error,
        "transaction_id": transaction_id,
        "created_at": _utc_now_iso(),
        "resolved_at": "",
    }
    alert_rows.append(alert_row)
    return alert_row


def _validate_cart_inventory(items: list[dict[str, Any]]) -> str | None:
    products = _load_products()
    product_index = _build_product_index(products)
    requested_by_product: dict[str, int] = {}

    for item in items:
        lookup_key = _product_lookup_key(item.get("product_id")) or _product_lookup_key(item.get("name"))
        if not lookup_key:
            continue
        requested_by_product[lookup_key] = requested_by_product.get(lookup_key, 0) + _to_int(item.get("quantity", 0), 0)

    errors: list[str] = []
    for lookup_key, requested_quantity in requested_by_product.items():
        product = product_index.get(lookup_key)
        if product is None:
            errors.append(f"Product {lookup_key} is no longer available.")
            continue

        stock_quantity = _to_int(product.get("stock_quantity", 0), 0)
        if requested_quantity > stock_quantity:
            errors.append(
                f"{product.get('product_name', lookup_key)} has only {stock_quantity} unit(s) left."
            )

    if errors:
        return " ".join(errors)
    return None


def _update_inventory_after_sale(items: list[dict[str, Any]], transaction_id: str) -> list[dict[str, Any]]:
    products = _load_products()
    product_index = _build_product_index(products)
    alert_rows = _read_table("inventory_alerts")
    movement_rows = _read_table("inventory_movements")
    triggered_alerts: list[dict[str, Any]] = []
    alerts_changed = False

    for item in items:
        lookup_key = _product_lookup_key(item.get("product_id")) or _product_lookup_key(item.get("name"))
        product = product_index.get(lookup_key)
        if product is None:
            print(f"[inventory] Product missing during stock deduction: {item.get('product_id') or item.get('name')}")
            continue

        quantity = max(_to_int(item.get("quantity", 0), 0), 0)
        stock_before = _to_int(product.get("stock_quantity", 0), 0)
        stock_after = max(stock_before - quantity, 0)
        product["stock_quantity"] = stock_after

        _record_inventory_movement(
            movement_rows=movement_rows,
            product=product,
            quantity_change=-quantity,
            stock_before=stock_before,
            stock_after=stock_after,
            movement_type="sale",
            transaction_id=transaction_id,
        )

        severity = _stock_alert_severity(product)
        if severity and _should_send_low_stock_alert(product, severity):
            triggered_alerts.append(_append_low_stock_alert(alert_rows, product, severity, transaction_id))
            alerts_changed = True
            product["last_low_stock_alert_at"] = _utc_now_iso()
            product["last_low_stock_alert_level"] = severity
        elif not severity:
            if _resolve_open_inventory_alerts(alert_rows, str(product.get("product_id", "")), _utc_now_iso()):
                alerts_changed = True
            product["last_low_stock_alert_at"] = ""
            product["last_low_stock_alert_level"] = ""
        elif not str(product.get("last_low_stock_alert_level", "")).strip():
            product["last_low_stock_alert_at"] = ""
            product["last_low_stock_alert_level"] = ""

    _save_products(products)
    _write_table("inventory_movements", movement_rows)
    if alerts_changed:
        _write_table("inventory_alerts", alert_rows)

    return triggered_alerts


def _build_inventory_insights(window_days: int = 30) -> dict[str, Any]:
    products = _load_products()
    alerts = _read_table("inventory_alerts")
    movements = _read_table("inventory_movements")
    cutoff = _utc_now() - timedelta(days=max(window_days, 1))
    sales_by_product: dict[str, int] = {}

    for movement in movements:
        if str(movement.get("movement_type", "")).strip().lower() != "sale":
            continue
        created_at = _parse_iso_datetime(movement.get("created_at", ""))
        if created_at is None or created_at < cutoff:
            continue
        product_id = _product_lookup_key(movement.get("product_id"))
        quantity_sold = abs(_to_int(movement.get("quantity_change", 0), 0))
        sales_by_product[product_id] = sales_by_product.get(product_id, 0) + quantity_sold

    product_metrics: list[dict[str, Any]] = []
    low_stock_items: list[dict[str, Any]] = []
    for product in products:
        product_id = _product_lookup_key(product.get("product_id"))
        sold_quantity = sales_by_product.get(product_id, 0)
        current_stock = _to_int(product.get("stock_quantity", 0), 0)
        reorder_level = _to_int(product.get("reorder_level", LOW_STOCK_DEFAULT_THRESHOLD), LOW_STOCK_DEFAULT_THRESHOLD)
        severity = _stock_alert_severity(product)
        avg_daily_sales = round(sold_quantity / max(window_days, 1), 2)
        stock_cover_days = round(current_stock / avg_daily_sales, 1) if avg_daily_sales > 0 else None

        metric = {
            "product_id": product.get("product_id", ""),
            "product_name": product.get("product_name", "Unknown Item"),
            "sold_quantity": sold_quantity,
            "current_stock": current_stock,
            "reorder_level": reorder_level,
            "stock_cover_days": stock_cover_days,
        }
        product_metrics.append(metric)

        if severity:
            low_stock_items.append({**metric, "severity": severity})

    fast_moving = [
        item
        for item in sorted(product_metrics, key=lambda item: item["sold_quantity"], reverse=True)
        if item["sold_quantity"] > 0
    ][:5]
    slow_moving = [
        item
        for item in sorted(product_metrics, key=lambda item: (item["sold_quantity"], -item["current_stock"]))
        if item["sold_quantity"] > 0 and item["current_stock"] > item["reorder_level"]
    ][:5]
    dead_stock = [
        item
        for item in sorted(product_metrics, key=lambda item: item["current_stock"], reverse=True)
        if item["sold_quantity"] == 0 and item["current_stock"] > 0
    ][:5]
    open_alerts = [alert for alert in alerts if str(alert.get("status", "")).strip().lower() == "open"]
    critical_count = sum(1 for item in low_stock_items if item["severity"] == "critical")

    return {
        "summary": {
            "window_days": window_days,
            "total_products": len(products),
            "open_alerts": len(open_alerts),
            "low_stock_count": len(low_stock_items),
            "critical_stock_count": critical_count,
        },
        "fast_moving": fast_moving,
        "slow_moving": slow_moving,
        "dead_stock": dead_stock,
        "low_stock": sorted(low_stock_items, key=lambda item: (item["severity"] != "critical", item["current_stock"])),
    }


def _load_cart_items() -> list[dict[str, Any]]:
    rows = _read_table("user_cart")
    items: list[dict[str, Any]] = []

    for row in rows:
        product_name = str(
            _find_value(row, ("product_name", "Product Name", "name", "Item Name"), "Unknown Item")
        ).strip()
        if not product_name:
            continue

        quantity = _to_int(_find_value(row, ("quantity", "qty", "Quantity"), 1), 1)
        quantity = quantity if quantity > 0 else 1
        price = round(_to_float(_find_value(row, ("price", "Price"), 0.0), 0.0), 2)
        total_cost = round(
            _to_float(_find_value(row, ("total_cost", "line_total", "total"), price * quantity), 0.0),
            2,
        )

        items.append(
            {
                "product_id": str(_find_value(row, ("product_id", "Product ID", "barcode"), "")).strip(),
                "name": product_name,
                "quantity": quantity,
                "price": price,
                "line_total": total_cost,
            }
        )

    return items


def _save_cart_items(items: list[dict[str, Any]]) -> None:
    rows: list[dict[str, Any]] = []
    for item in items:
        rows.append(
            {
                "cart_id": item.get("cart_id", "default"),
                "product_id": item.get("product_id", ""),
                "product_name": item.get("name", "Unknown Item"),
                "quantity": item.get("quantity", 1),
                "price": item.get("price", 0.0),
                "total_cost": item.get("line_total", 0.0),
            }
        )
    _write_table("user_cart", rows)


def _cart_total(items: list[dict[str, Any]]) -> float:
    return round(sum(_to_float(item.get("line_total"), 0.0) for item in items), 2)


def _snapshot_cart_for_undo(items: list[dict[str, Any]]) -> None:
    session["undo_cart_snapshot"] = {
        "items": [dict(item) for item in items],
        "created_at": _utc_now_iso(),
    }
    session.modified = True


# ---------- Payment gateway helpers ----------
def _paypal_access_token() -> str:
    if not PAYPAL_CLIENT_ID or not PAYPAL_CLIENT_SECRET:
        raise ValueError("Set PAYPAL_CLIENT_ID and PAYPAL_CLIENT_SECRET before checkout.")

    credentials = f"{PAYPAL_CLIENT_ID}:{PAYPAL_CLIENT_SECRET}".encode("utf-8")
    basic_auth = base64.b64encode(credentials).decode("utf-8")
    token_request = urllib_request.Request(
        f"{PAYPAL_BASE_URL}/v1/oauth2/token",
        data=b"grant_type=client_credentials",
        headers={
            "Authorization": f"Basic {basic_auth}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urllib_request.urlopen(token_request, timeout=20) as response:
            token_body = response.read().decode("utf-8")
    except urllib_error.HTTPError as exc:
        error_text = exc.read().decode("utf-8", errors="ignore")
        raise ValueError(f"PayPal auth failed: {error_text or exc.reason}") from exc
    except urllib_error.URLError as exc:
        raise ValueError(f"PayPal is unreachable: {exc.reason}") from exc

    token_payload = json.loads(token_body)
    access_token = str(token_payload.get("access_token", "")).strip()
    if not access_token:
        raise ValueError("PayPal auth response missing access token.")
    return access_token


def _create_paypal_order(
    total_amount: float,
    customer_name: str,
    return_url: str,
    cancel_url: str,
) -> dict[str, Any]:
    access_token = _paypal_access_token()
    payload = json.dumps(
        {
            "intent": "CAPTURE",
            "purchase_units": [
                {
                    "reference_id": f"quickbill-{uuid.uuid4().hex[:10]}",
                    "description": f"QuickBill checkout for {customer_name}",
                    "amount": {
                        "currency_code": PAYMENT_CURRENCY,
                        "value": f"{total_amount:.2f}",
                    },
                }
            ],
            "payment_source": {
                "paypal": {
                    "experience_context": {
                        "brand_name": "QuickBill",
                        "user_action": "PAY_NOW",
                        "return_url": return_url,
                        "cancel_url": cancel_url,
                    }
                }
            },
        }
    ).encode("utf-8")

    api_request = urllib_request.Request(
        f"{PAYPAL_BASE_URL}/v2/checkout/orders",
        data=payload,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "PayPal-Request-Id": f"qb-{uuid.uuid4().hex}",
            "Prefer": "return=representation",
        },
        method="POST",
    )

    try:
        with urllib_request.urlopen(api_request, timeout=20) as response:
            response_body = response.read().decode("utf-8")
    except urllib_error.HTTPError as exc:
        error_text = exc.read().decode("utf-8", errors="ignore")
        raise ValueError(f"PayPal order creation failed: {error_text or exc.reason}") from exc
    except urllib_error.URLError as exc:
        raise ValueError(f"PayPal is unreachable: {exc.reason}") from exc

    order = json.loads(response_body)
    if not order.get("id"):
        raise ValueError("PayPal response missing order id.")
    return order


def _capture_paypal_order(order_id: str) -> dict[str, Any]:
    access_token = _paypal_access_token()
    api_request = urllib_request.Request(
        f"{PAYPAL_BASE_URL}/v2/checkout/orders/{order_id}/capture",
        data=b"{}",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib_request.urlopen(api_request, timeout=20) as response:
            response_body = response.read().decode("utf-8")
    except urllib_error.HTTPError as exc:
        error_text = exc.read().decode("utf-8", errors="ignore")
        raise ValueError(f"PayPal capture failed: {error_text or exc.reason}") from exc
    except urllib_error.URLError as exc:
        raise ValueError(f"PayPal is unreachable: {exc.reason}") from exc

    capture = json.loads(response_body)
    if capture.get("status") != "COMPLETED":
        raise ValueError("PayPal capture is not completed.")
    return capture


def _paypal_approve_url(order: dict[str, Any]) -> str:
    links = order.get("links") or []
    fallback_href = ""
    for link in links:
        if not isinstance(link, dict):
            continue
        rel = str(link.get("rel", "")).strip().lower()
        href = str(link.get("href", "")).strip()
        if not href:
            continue
        if rel == "payer-action":
            return href
        if rel == "approve":
            fallback_href = href
    if fallback_href:
        return fallback_href
    available_rels = ", ".join(
        str(link.get("rel", "")).strip()
        for link in links
        if isinstance(link, dict) and str(link.get("rel", "")).strip()
    )
    if available_rels:
        raise ValueError(f"PayPal response missing approval link. Available links: {available_rels}")
    raise ValueError("PayPal response missing approval link. No HATEOAS links returned.")


def _extract_paypal_capture_id(capture: dict[str, Any]) -> str:
    capture_units = capture.get("purchase_units") or []
    if not capture_units or not isinstance(capture_units[0], dict):
        return ""
    capture_records = capture_units[0].get("payments", {}).get("captures", [])
    if not capture_records or not isinstance(capture_records[0], dict):
        return ""
    return str(capture_records[0].get("id", "")).strip()


def _finalize_paypal_payment(order_id: str, capture: dict[str, Any]) -> None:
    pending_payment = session.get("pending_payment")
    if not pending_payment or not isinstance(pending_payment, dict):
        raise ValueError("No pending payment found.")

    if order_id != pending_payment.get("order_id"):
        raise ValueError("Order id mismatch for pending payment.")

    payment_id = _extract_paypal_capture_id(capture)
    if not payment_id:
        payment_id = f"PP-{uuid.uuid4().hex[:12].upper()}"

    customer_name = str(pending_payment.get("customer", {}).get("name", "Guest Customer"))
    items = pending_payment.get("items", [])
    items_bought = ", ".join(f"{item['name']} x{item['quantity']}" for item in items)
    transaction_id = f"TXN-{uuid.uuid4().hex[:10].upper()}"
    _update_inventory_after_sale(items, transaction_id)

    sales = _read_table("store_sales")
    sales.append(
        {
            "transaction_id": transaction_id,
            "customer_name": customer_name,
            "items_bought": items_bought,
            "total_amount": pending_payment.get("total_amount", 0.0),
            "currency": pending_payment.get("currency", PAYMENT_CURRENCY),
            "date": datetime.now().isoformat(timespec="seconds"),
            "payment_order_id": order_id,
            "payment_id": payment_id,
            "status": "paid",
        }
    )
    _write_table("store_sales", sales)

    _save_cart_items([])
    session.pop("undo_cart_snapshot", None)
    session.pop("pending_payment", None)
    session.modified = True


# ---------- Page routes ----------
@app.route("/")
def index():
    return render_template("app.html")


@app.route("/form")
def form():
    return render_template("form.html")


@app.route("/customer")
def customer():
    return render_template(
        "customer.html",
        paypal_client_id=PAYPAL_CLIENT_ID,
        payment_currency=PAYMENT_CURRENCY,
        payment_message=request.args.get("payment_message", "").strip(),
    )


@app.route("/checkout-success")
def checkout_success():
    return render_template("checkout_success.html")


@app.route("/admin/inventory-alerts")
def admin_inventory_alerts():
    status_filter = str(request.args.get("status", "open")).strip().lower()
    alerts = _read_table("inventory_alerts")
    filtered_alerts = alerts
    if status_filter and status_filter != "all":
        filtered_alerts = [
            alert
            for alert in alerts
            if str(alert.get("status", "")).strip().lower() == status_filter
        ]

    filtered_alerts.sort(
        key=lambda alert: _parse_iso_datetime(alert.get("created_at", "")) or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True,
    )
    return jsonify({"alerts": filtered_alerts, "count": len(filtered_alerts), "status_filter": status_filter})
  

@app.route("/admin/inventory-insights")
def admin_inventory_insights():
    days = _to_int(request.args.get("days", 30), 30)
    return jsonify(_build_inventory_insights(days))


# ---------- Signup/session route ----------
@app.route("/submit_form", methods=["POST"])
def submit_form():
    name = request.form.get("name", "").strip()
    phone = request.form.get("phone", "").strip() or request.form.get("mobile_number", "").strip()

    if not name or not phone:
        return "Name and phone number are required.", 400

    users = _read_table("users")
    input_phone = _normalize_phone(phone)

    duplicate_found = False
    for row in users:
        existing_phone = str(_find_value(row, ("Phone Number", "Mobile Number", "phone"), "")).strip()
        if input_phone and _normalize_phone(existing_phone) == input_phone:
            duplicate_found = True
            break

    if not duplicate_found:
        users.append({"Name": name, "Phone Number": phone})
        _write_table("users", users)

    session["customer"] = {"name": name, "phone": phone}
    session.modified = True
    return redirect(url_for("customer"))


# ---------- Cart routes ----------
@app.route("/get_items")
def get_items():
    items = _load_cart_items()
    payload: list[dict[str, Any]] = []
    for index, item in enumerate(items):
        payload.append(
            {
                "line_index": index,
                "Item Name": item["name"],
                "Price": item["price"],
                "Quantity": item["quantity"],
                "Line Total": item["line_total"],
            }
        )
    return jsonify(
        {
            "items": payload,
            "total_amount": _cart_total(items),
            "currency": PAYMENT_CURRENCY,
        }
    )


@app.route("/recommendations")
def recommendations():
    products = _load_products()
    items = [
        {
            "product_id": product["product_id"],
            "name": product["product_name"],
            "price": product["price"],
        }
        for product in products[:8]
    ]
    return jsonify({"items": items})


@app.route("/add_item", methods=["POST"])
def add_item():
    payload = request.get_json(silent=True) or request.form.to_dict()
    product_id = str(payload.get("product_id") or payload.get("barcode") or "").strip()
    normalized_product_id = product_id.lower()
    quantity = _to_int(payload.get("quantity", 1), 1)
    quantity = quantity if quantity > 0 else 1

    if not product_id:
        return jsonify({"error": "product_id or barcode is required."}), 400

    products = _load_products()
    product_match = next(
        (
            product
            for product in products
            if str(product["product_id"]).strip().lower() == normalized_product_id
            or str(product["product_name"]).strip().lower() == normalized_product_id
        ),
        None,
    )
    if product_match is None:
        return jsonify({"error": f"Product {product_id} not found."}), 404

    available_stock = _to_int(product_match.get("stock_quantity", 0), 0)
    if available_stock <= 0:
        return jsonify({"error": f"{product_match['product_name']} is out of stock."}), 400

    cart_items = _load_cart_items()
    reserved_quantity = sum(
        _to_int(item.get("quantity", 0), 0)
        for item in cart_items
        if _product_lookup_key(item.get("product_id")) == _product_lookup_key(product_match["product_id"])
    )
    if reserved_quantity + quantity > available_stock:
        return jsonify({"error": f"Only {available_stock - reserved_quantity} unit(s) left for {product_match['product_name']}."}), 400

    line_total = round(product_match["price"] * quantity, 2)
    cart_items.append(
        {
            "cart_id": str(payload.get("cart_id") or "default"),
            "product_id": product_match["product_id"],
            "name": product_match["product_name"],
            "quantity": quantity,
            "price": product_match["price"],
            "line_total": line_total,
        }
    )
    _save_cart_items(cart_items)

    return jsonify(
        {
            "Item Name": product_match["product_name"],
            "Price": product_match["price"],
            "Quantity": quantity,
            "Line Total": line_total,
        }
    )


@app.route("/remove_item", methods=["POST"])
def remove_item():
    payload = request.get_json(silent=True) or request.form.to_dict()
    line_index = _to_int(payload.get("line_index", -1), -1)

    cart_items = _load_cart_items()
    if line_index < 0 or line_index >= len(cart_items):
        return jsonify({"error": "Invalid cart item index."}), 400

    _snapshot_cart_for_undo(cart_items)
    removed_item = cart_items.pop(line_index)
    _save_cart_items(cart_items)

    return jsonify(
        {
            "message": "Item removed from cart.",
            "removed_item": removed_item.get("name", "Item"),
            "items_remaining": len(cart_items),
            "total_amount": _cart_total(cart_items),
        }
    )


@app.route("/clear_cart", methods=["POST"])
def clear_cart():
    cart_items = _load_cart_items()
    if not cart_items:
        return jsonify({"message": "Cart is already empty.", "items_remaining": 0, "total_amount": 0.0})

    _snapshot_cart_for_undo(cart_items)
    _save_cart_items([])

    return jsonify(
        {
            "message": "All items removed from cart.",
            "removed_count": len(cart_items),
            "items_remaining": 0,
            "total_amount": 0.0,
        }
    )


@app.route("/undo_cart_action", methods=["POST"])
def undo_cart_action():
    snapshot = session.get("undo_cart_snapshot")
    if not snapshot or not isinstance(snapshot, dict):
        return jsonify({"error": "No cart action available to undo."}), 400

    items = snapshot.get("items", [])
    if not isinstance(items, list):
        return jsonify({"error": "Undo snapshot is invalid."}), 400

    _save_cart_items(items)
    session.pop("undo_cart_snapshot", None)
    session.modified = True

    return jsonify(
        {
            "message": "Last cart action undone.",
            "restored_count": len(items),
            "total_amount": _cart_total(items),
        }
    )


@app.route("/checkout", methods=["POST"])
def checkout():
    items = _load_cart_items()
    if not items:
        return jsonify({"message": "Cart is empty. Add items before checkout."}), 400

    inventory_error = _validate_cart_inventory(items)
    if inventory_error:
        return jsonify({"message": inventory_error}), 400

    total_amount = _cart_total(items)
    if total_amount <= 0:
        return jsonify({"message": "Cart total is zero. Cannot create payment order."}), 400

    customer = session.get("customer", {})
    customer_name = str(customer.get("name", "Guest Customer"))

    try:
        order = _create_paypal_order(
            total_amount=total_amount,
            customer_name=customer_name,
            return_url=url_for("paypal_return", _external=True),
            cancel_url=url_for("paypal_cancel", _external=True),
        )
        approve_url = _paypal_approve_url(order)
    except ValueError as exc:
        return jsonify({"message": str(exc)}), 400

    # Save immutable checkout snapshot for post-payment capture.
    session["pending_payment"] = {
        "order_id": order["id"],
        "total_amount": total_amount,
        "currency": PAYMENT_CURRENCY,
        "customer": customer,
        "items": items,
        "created_at": _utc_now_iso(),
    }
    session.modified = True

    return jsonify(
        {
            "message": "Order created.",
            "order_id": order["id"],
            "approve_url": approve_url,
        }
    )


@app.route("/paypal/return")
def paypal_return():
    order_id = str(request.args.get("token", "")).strip()
    if not order_id:
        return redirect(url_for("customer", payment_message="Missing PayPal order id."))

    pending_payment = session.get("pending_payment")
    if not pending_payment or not isinstance(pending_payment, dict):
        return redirect(url_for("customer", payment_message="No pending payment found."))

    inventory_error = _validate_cart_inventory(pending_payment.get("items", []))
    if inventory_error:
        return redirect(url_for("customer", payment_message=inventory_error))

    try:
        capture = _capture_paypal_order(order_id=order_id)
        _finalize_paypal_payment(order_id=order_id, capture=capture)
    except ValueError as exc:
        return redirect(url_for("customer", payment_message=str(exc)))

    return redirect(url_for("checkout_success"))


@app.route("/paypal/cancel")
def paypal_cancel():
    session.pop("pending_payment", None)
    session.modified = True
    return redirect(url_for("customer", payment_message="PayPal checkout was cancelled."))


@app.route("/verify-payment", methods=["POST"])
def verify_payment():
    pending_payment = session.get("pending_payment")
    if not pending_payment:
        return jsonify({"message": "No pending payment found."}), 400

    payload = request.get_json(silent=True) or {}
    order_id = str(payload.get("paypal_order_id") or payload.get("orderID") or "").strip()

    if not order_id:
        return jsonify({"message": "Missing PayPal order id."}), 400

    if order_id != pending_payment.get("order_id"):
        return jsonify({"message": "Order id mismatch for pending payment."}), 400

    inventory_error = _validate_cart_inventory(pending_payment.get("items", []))
    if inventory_error:
        return jsonify({"message": inventory_error}), 400

    try:
        capture = _capture_paypal_order(order_id=order_id)
        _finalize_paypal_payment(order_id=order_id, capture=capture)
    except ValueError as exc:
        return jsonify({"message": str(exc)}), 400

    return jsonify(
        {
            "message": "Payment successful. Bill generated.",
            "redirect_url": url_for("checkout_success"),
        }
    )


_init_mongodb()
_ensure_tables()

if __name__ == "__main__":
    app.run(debug=True, port=5000)
