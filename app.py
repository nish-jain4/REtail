"""
REtail - Main Flask Application
Backend API routes for frontend connectivity.
"""
from flask import Flask, render_template, request, redirect, url_for, jsonify
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "retail-dev-key")

# File paths
EXCEL_FILE = "users.xlsx"
USER_CART_FILE = "user_cart.xlsx"
PRODUCTS_FILE = "products_database.xlsx"
STORE_SALES_FILE = "store_sales.xlsx"

# Ensure users.xlsx exists
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=["Name", "Phone Number"])
    df.to_excel(EXCEL_FILE, index=False)

# Ensure user cart exists
if not os.path.exists(USER_CART_FILE):
    pd.DataFrame(columns=["cart_id", "product_id", "product_name", "quantity", "price", "total_cost"]).to_excel(USER_CART_FILE, index=False)

# Ensure products database exists (from products.csv or new)
if not os.path.exists(PRODUCTS_FILE) and os.path.exists("products.csv"):
    products_df = pd.read_csv("products.csv")
    if "Barcode" in products_df.columns:
        products_df = products_df.rename(columns={"Barcode": "product_id", "Name": "product_name", "Stock": "stock_quantity"})
        if "last_restock_date" not in products_df.columns:
            products_df["last_restock_date"] = datetime.today().strftime("%Y-%m-%d")
    products_df.to_excel(PRODUCTS_FILE, index=False)


def load_cart():
    if os.path.exists(USER_CART_FILE):
        return pd.read_excel(USER_CART_FILE, engine="openpyxl")
    return pd.DataFrame(columns=["cart_id", "product_id", "product_name", "quantity", "price", "total_cost"])


def save_cart(df):
    df.to_excel(USER_CART_FILE, index=False, engine="openpyxl")


# ---- Pages ----
@app.route("/")
def index():
    return render_template("customer.html")


@app.route("/customer")
def customer():
    return render_template("customer.html")


@app.route("/checkout-success")
def checkout_form():
    return render_template("checkout_success.html")


# ---- API: Signup ----
@app.route("/submit_form", methods=["POST"])
def submit_form():
    """Expects: name, phone (or mobile_number for compatibility)"""
    name = request.form.get("name", "").strip()
    phone = request.form.get("phone") or request.form.get("mobile_number", "").strip()

    if not name or not phone:
        return "Name and phone number are required.", 400

    df = pd.read_excel(EXCEL_FILE)
    col_phone = "Phone Number" if "Phone Number" in df.columns else "Mobile Number"
    if col_phone not in df.columns:
        df[col_phone] = ""
    if phone in df[col_phone].astype(str).values:
        return "Phone number already registered.", 409
    new_user = pd.DataFrame([[name, phone]], columns=df.columns)
    df = pd.concat([df, new_user], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    return redirect(url_for("index"))


# ---- API: Cart & Checkout ----
@app.route("/get_items")
def get_items():
    """Returns JSON: [{ "Item Name": str, "Price": float, "quantity": int, ... }]"""
    cart = load_cart()
    items = []
    for _, row in cart.iterrows():
        items.append({
            "Item Name": str(row.get("product_name", row.get("Product Name", "Unknown"))),
            "Price": float(row.get("price", row.get("Price", 0))),
            "quantity": int(row.get("quantity", 1)),
        })
    return jsonify(items)


@app.route("/checkout", methods=["POST"])
def checkout():
    """Process checkout, clear cart. Returns JSON: { "message": str }"""
    cart = load_cart()
    if cart.empty:
        return jsonify({"message": "Cart is empty.", "success": False}), 400
    total = float(cart["total_cost"].sum()) if "total_cost" in cart.columns else 0
    items_str = ", ".join([
        f"{row.get('product_name', 'Item')} x{row.get('quantity', 1)}"
        for _, row in cart.iterrows()
    ])
    # Save to store_sales if file exists
    if os.path.exists(STORE_SALES_FILE):
        sales = pd.read_excel(STORE_SALES_FILE, engine="openpyxl")
        new_row = pd.DataFrame([{
            "transaction_id": f"TXN{len(sales) + 1:04d}",
            "customer_name": request.form.get("customer_name", "Guest"),
            "items_bought": items_str,
            "total_amount": total,
            "date": datetime.today().strftime("%Y-%m-%d"),
        }])
        sales = pd.concat([sales, new_row], ignore_index=True)
        sales.to_excel(STORE_SALES_FILE, index=False, engine="openpyxl")
    # Clear cart
    save_cart(pd.DataFrame(columns=load_cart().columns))
    return jsonify({"message": f"Payment of ${total:.2f} received. Thank you!", "success": True})


# ---- Add item (for barcode scan / manual add) ----
@app.route("/add_item", methods=["POST"])
def add_item():
    """Expects JSON: { product_id, quantity } or { barcode, quantity }"""
    data = request.get_json() or {}
    product_id = str(data.get("product_id") or data.get("barcode", "")).strip()
    quantity = int(data.get("quantity", 1))
    if not product_id:
        return jsonify({"error": "product_id or barcode required"}), 400
    if not os.path.exists(PRODUCTS_FILE):
        return jsonify({"error": "Products database not found"}), 503
    products = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
    prod_col = "product_id" if "product_id" in products.columns else "Barcode"
    match = products[products[prod_col].astype(str) == product_id]
    if match.empty:
        return jsonify({"error": f"Product {product_id} not found"}), 404
    row = match.iloc[0]
    name = row.get("product_name", row.get("Product Name", row.get("Name", "Unknown")))
    price = float(row.get("price", row.get("Price", 0)))
    total_cost = price * quantity
    cart = load_cart()
    cart_id = str(data.get("cart_id") or request.form.get("cart_id") or request.args.get("cart_id", "default"))
    new_row = pd.DataFrame([{
        "cart_id": cart_id,
        "product_id": product_id,
        "product_name": name,
        "quantity": quantity,
        "price": price,
        "total_cost": total_cost,
    }])
    cart = pd.concat([cart, new_row], ignore_index=True)
    save_cart(cart)
    return jsonify({"Item Name": name, "Price": price, "quantity": quantity})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
