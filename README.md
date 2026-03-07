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
python -m pip install -r require.txt
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

When `openpyxl` is installed, files are written as `.xlsx`.  
If `openpyxl` is missing, the app falls back to `.csv` files.
