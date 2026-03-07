# REtail – Backend API Reference

This document describes the API contracts the frontend expects. Backend developers can use this to integrate or replace endpoints.

## Run the app

```bash
python app.py
```

Runs on `http://127.0.0.1:5000`

---

## Routes & Endpoints

### Pages (HTML)

| Route | Method | Response |
|-------|--------|----------|
| `/` | GET | Home page (index.html) |
| `/customer` | GET | Shopping / scan page (customer.html) |
| `/form` | GET | Sign-up form (form.html) |
| `/checkout-success` | GET | Thank-you page (checkout_success.html) |

---

### API: Sign-up

**`POST /submit_form`**

- **Content-Type:** `application/x-www-form-urlencoded`
- **Form fields:**
  - `name` (required): Full name
  - `phone` or `mobile_number` (required): Phone number
- **Success:** 302 redirect to `/`
- **Errors:**
  - 400: Missing name or phone
  - 409: Phone number already registered (body: plain text message)

---

### API: Cart & Checkout

**`GET /get_items`**

- **Response:** JSON array of cart items
- **Shape:**
  ```json
  [
    {
      "Item Name": "Product name",
      "Price": 2.99,
      "quantity": 2
    }
  ]
  ```

**`POST /checkout`**

- **Response:** JSON object
- **Success:**
  ```json
  {
    "message": "Payment of $X.XX received. Thank you!",
    "success": true
  }
  ```
- **Error (400):**
  ```json
  {
    "message": "Cart is empty.",
    "success": false
  }
  ```

**`POST /add_item`**

- **Content-Type:** `application/json`
- **Body:**
  ```json
  {
    "product_id": "12345",
    "quantity": 1
  }
  ```
  - `barcode` can be used instead of `product_id`
- **Success:** JSON with `Item Name`, `Price`, `quantity`
- **Errors:**
  - 400: Missing `product_id` or `barcode`
  - 404: Product not found
  - 503: Products database unavailable

---

## Data files

- `users.xlsx` – Name, Phone Number
- `user_cart.xlsx` – cart_id, product_id, product_name, quantity, price, total_cost
- `products_database.xlsx` – product_id, product_name, stock_quantity, price, last_restock_date
- `store_sales.xlsx` – transaction_id, customer_name, items_bought, total_amount, date

---

## Barcode / camera integration

The customer page has a placeholder for a barcode scanner. Backend can:

1. Stream video from the camera and decode barcodes (e.g. via pyzbar).
2. When a barcode is detected, call `POST /add_item` with the decoded value and quantity.
3. Alternatively, the frontend supports manual entry via the “Add by Barcode” input.
