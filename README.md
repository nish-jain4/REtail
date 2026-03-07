<<<<<<< Updated upstream
# QuickBill
QuickBill is a self-checkout web app for retail stores. Shoppers can scan items, see the running bill, and pay in-app before leaving the store.

## Payment Gateway
This app now supports billing through Razorpay checkout.

### Required environment variables
- `RAZORPAY_KEY_ID`
- `RAZORPAY_KEY_SECRET`
- `FLASK_SECRET_KEY` (recommended for production)
- `PAYMENT_CURRENCY` (optional, defaults to `INR`)

### Run locally
```powershell
cd Quick-Bill
set RAZORPAY_KEY_ID=rzp_test_xxxxxxxx
set RAZORPAY_KEY_SECRET=xxxxxxxxxxxxxxxx
set FLASK_SECRET_KEY=change_me
python app.py
```

Open `http://127.0.0.1:5000`.

## Core checkout endpoints
- `GET /get_items` - returns scanned cart items and total amount.
- `POST /checkout` - creates a Razorpay order for the current cart.
- `POST /verify-payment` - verifies Razorpay signature and stores the sale in `store_sales`.
=======
# REtail

REtail is your personal cashier in your pocket. It eliminates long queues by allowing users to scan items, track their bills in real time, and pay directly through the app before leaving the store.
>>>>>>> Stashed changes
