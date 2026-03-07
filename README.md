# REtail

REtail is your personal cashier in your pocket. It eliminates long queues by allowing users to scan items, track their bills in real time, and pay directly through the app before leaving the store.

## Run locally

```bash
python app.py
```

Open `http://127.0.0.1:5000`.

## Core endpoints

- `GET /get_items` – returns cart items and total
- `POST /checkout` – processes checkout, clears cart
- `POST /add_item` – add product by barcode/product ID
