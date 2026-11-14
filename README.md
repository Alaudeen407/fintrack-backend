# fintrack-backend

FinTrack & FinSet Expense & Income Tracker REST API built with Flask, MongoDB Atlas Cloud Database, and OpenPyXL Excel reporting engine.

## Overview
A lightweight financial management backend providing:
- Full Credit (Income) & Debit (Expense) transaction tracking
- MongoDB Atlas Cloud integration with automatic synchronization
- Local SQLite3 persistent fallback cache
- Formatted Multi-Sheet Microsoft Excel (`.xlsx`) generator with dedicated `Transaction Date` and `Time` columns
- Recurring Monthly Bill Management & Due Date Calendar tracking
- User Authentication (Login / Register with password hashing)

## Tech Stack
- **Framework:** Python 3.12 + Flask 3.x
- **Cloud Database:** MongoDB Atlas (PyMongo 4.x, dnspython)
- **Local Fallback:** SQLite 3 (expenses.db)
- **Excel Generator:** OpenPyXL 3.x
- **Cross-Origin:** Flask-CORS
- **Configuration:** Python-Dotenv

## Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/Alaudeen407/fintrack-backend.git
cd fintrack-backend
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Configure Environment Variables
Copy `.env.example` to `.env` and fill in your MongoDB Atlas connection details:
```bash
cp .env.example .env
```
Update your credentials inside `.env`:
```env
MONGODB_USERNAME=your_username
MONGODB_PASSWORD=your_password
MONGODB_URI=mongodb+srv://<username>:<password>@cluster0.kh3a2oo.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0
MONGODB_DB_NAME=finset_db
```

### 4. Run the Backend API
```bash
python app.py
```
The server will start on `http://127.0.0.1:5000`.

## API Endpoints

### Authentication
- `POST /api/auth/register` — Create a new user account
- `POST /api/auth/login` — Sign in with email and password
- `GET /api/auth/me` — Retrieve active user session profile

### Transactions
- `GET /api/transactions` — Query all transactions (supports `type`, `category`, `search`, `startDate`, `endDate`, `sort`)
- `POST /api/transactions` — Create a new credit or debit transaction
- `GET /api/transactions/<id>` — Retrieve single transaction by ID
- `PUT /api/transactions/<id>` — Update transaction details
- `DELETE /api/transactions/<id>` — Delete a transaction

### Summary & Analytics
- `GET /api/summary` — Overview metrics (total balance, monthly credit/debit, category breakdown, monthly trends)

### Monthly Bills Management
- `GET /api/bills` — List recurring bills (filterable by status: `paid`, `unpaid`, `overdue`)
- `POST /api/bills` — Add a new bill with due day and amount
- `PUT /api/bills/<id>` — Update bill details
- `DELETE /api/bills/<id>` — Delete a bill
- `POST /api/bills/<id>/pay` — Mark bill as paid and record corresponding debit expense

### Excel & Reports
- `GET /api/report/preview?month=YYYY-MM` — Live preview of statement metrics and ledger
- `GET /api/report/excel?month=YYYY-MM` — Download formatted 3-sheet Excel workbook (`.xlsx`)
  - **Sheet 1:** Transactions Ledger with explicit `Transaction Date` and `Time`
  - **Sheet 2:** Financial Summary with KPI blocks & category totals
  - **Sheet 3:** Monthly Bills schedule & auto-pay status
- `GET /api/export` — Export full ledger to CSV

### System & Database Status
- `GET /api/mongodb/status` — Live MongoDB Atlas connection and collection statistics
- `POST /api/reset` — Reset database to demo dataset
