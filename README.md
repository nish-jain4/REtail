# QuickBill

QuickBill is a retail self-checkout web app where users can:
- sign up,
- add/scan items into a cart,
- see live totals,
- pay using Razorpay,
- and store paid transactions.

## Single Entry Point

Use only:
- `Quick-Bill/app.py`

Other Python files in this folder are old prototypes and are not part of the main web flow.

## Routes

- `GET /` -> home
- `GET /form` -> signup page
- `GET /customer` -> cart + add item + checkout page
- `POST /submit_form` -> save user and start shopping session
- `GET /get_items` -> cart items + total
- `POST /add_item` -> add item by product id/barcode
- `POST /checkout` -> create Razorpay order
- `POST /verify-payment` -> verify signature and finalize bill
- `GET /checkout-success` -> success page

## Required Environment Variables

- `RAZORPAY_KEY_ID`
- `RAZORPAY_KEY_SECRET`

Recommended:
- `FLASK_SECRET_KEY`
- `PAYMENT_CURRENCY` (default: `INR`)

## Run Locally (Windows PowerShell)

```powershell
cd Quick-Bill
python -m pip install -r requirements.txt
set RAZORPAY_KEY_ID=rzp_test_xxxxxxxxxxxx
set RAZORPAY_KEY_SECRET=xxxxxxxxxxxxxxxx
set FLASK_SECRET_KEY=change_me
python app.py
```

Open: `http://127.0.0.1:5000`

## Storage

The app uses these tables:
- `users`
- `user_cart`
- `products_database`
- `store_sales`

Storage is now MongoDB-only.

- Default cluster URI matches the working Admin setup.
- Default database is `retail`.
- Default collection prefix is `retail_` (for example `retail_users`, `retail_products_database`) to avoid collisions with Admin collections.

Optional overrides:
- `MONGODB_URI`
- `MONGODB_DB_NAME`
- `MONGODB_COLLECTION_PREFIX`

On first startup, existing local `.xlsx`/`.csv` tables are auto-migrated into MongoDB collections.

## Admin Login

Admin can now log in from the normal `/login` page using:
- admin email
- admin password
- secret key

Configurable variables:
- `ADMIN_LOGIN_EMAIL` (default: `admin@retail.local`)
- `ADMIN_LOGIN_PASSWORD` (default: `Admin@Retail2026`)

Secret key uses `ADMIN_DASHBOARD_SECRET_KEY` in code (default: `retail-admin-2026`).

## Admin Stocking Flow (Barcode + Category)

Admin dashboard now supports scanner-first stocking:

- Scan/enter unique barcode
- Choose category name (same category can contain many barcodes)
- Set category price (required only for first barcode in a new category)
- Submit adds exactly 1 stock unit for that barcode

Rules:
- Barcode must be unique
- Same category must keep one common price
