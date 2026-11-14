import os
import sqlite3
import csv
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, request, jsonify, Response, send_from_directory
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
from pymongo import MongoClient

# Load environment variables (.env)
load_dotenv()

FRONTEND_DIST = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'frontend', 'dist'))

app = Flask(__name__, static_folder=os.path.join(FRONTEND_DIST, 'assets'))
CORS(app, resources={r"/api/*": {"origins": "*"}})

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'expenses.db')

# ==================== MONGODB ATLAS CONNECTION ====================
MONGODB_URI = os.getenv(
    'MONGODB_URI',
    'mongodb+srv://aladdina407_db_user:jwGTTA0FEV7ydhEi@cluster0.kh3a2oo.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0'
)
MONGODB_DB_NAME = os.getenv('MONGODB_DB_NAME', 'finset_db')

_mongo_client = None
_mongo_db = None

def get_mongo_db():
    global _mongo_client, _mongo_db
    if _mongo_db is not None:
        return _mongo_db
    try:
        _mongo_client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=4000)
        _mongo_client.admin.command('ping')
        _mongo_db = _mongo_client[MONGODB_DB_NAME]
        print(f"[*] MongoDB Atlas Connected successfully: {MONGODB_DB_NAME}")
        return _mongo_db
    except Exception as e:
        print(f"[!] MongoDB Atlas connection notice: {e}")
        return None

def clean_doc(doc):
    if not doc:
        return None
    d = dict(doc)
    if '_id' in d:
        d['_id'] = str(d['_id'])
    return d

def clean_docs(docs):
    return [clean_doc(d) for d in docs]


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(force_reseed=False):
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Transactions table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL CHECK(type IN ('credit', 'debit')),
            amount REAL NOT NULL CHECK(amount > 0),
            category TEXT NOT NULL,
            description TEXT NOT NULL,
            date TEXT NOT NULL,
            payment_method TEXT NOT NULL DEFAULT 'Cash',
            notes TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_trans_date ON transactions(date)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_trans_type ON transactions(type)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_trans_category ON transactions(category)')

    # 2. Users table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            avatar TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 3. Monthly Bills table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            amount REAL NOT NULL CHECK(amount > 0),
            due_day INTEGER NOT NULL,
            due_date TEXT NOT NULL,
            category TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'unpaid',
            auto_pay INTEGER NOT NULL DEFAULT 0,
            paid_date TEXT,
            notes TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_bills_due ON bills(due_date)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_bills_status ON bills(status)')

    if force_reseed:
        cursor.execute('DELETE FROM transactions')
        cursor.execute('DELETE FROM users')
        cursor.execute('DELETE FROM bills')

    # Seed users if empty
    cursor.execute('SELECT COUNT(*) FROM users')
    if cursor.fetchone()[0] == 0:
        cursor.execute('''
            INSERT INTO users (name, email, password_hash, avatar)
            VALUES (?, ?, ?, ?)
        ''', (
            'Adaline Lively',
            'adalinect@gmail.com',
            generate_password_hash('password123'),
            'https://images.unsplash.com/photo-1494790108377-be9c29b29330?auto=format&fit=crop&w=120&q=80'
        ))

    # Seed transactions if empty
    cursor.execute('SELECT COUNT(*) FROM transactions')
    if cursor.fetchone()[0] == 0:
        seed_sample_data(cursor)
    else:
        # Ensure realistic timestamps matching transaction date if out of sync
        cursor.execute('''
            UPDATE transactions 
            SET created_at = date || ' ' || printf('%02d:%02d:%02d', 9 + (id % 10), (id * 7) % 60, (id * 13) % 60)
            WHERE SUBSTR(created_at, 1, 10) != date
        ''')

    # Seed bills if empty
    cursor.execute('SELECT COUNT(*) FROM bills')
    if cursor.fetchone()[0] == 0:
        seed_sample_bills(cursor)

    conn.commit()
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            if force_reseed or mdb.transactions.count_documents({}) == 0:
                conn_sync = get_db_connection()
                c_sync = conn_sync.cursor()
                c_sync.execute('SELECT * FROM transactions')
                all_tx = [dict(r) for r in c_sync.fetchall()]
                c_sync.execute('SELECT * FROM users')
                all_u = [dict(r) for r in c_sync.fetchall()]
                c_sync.execute('SELECT * FROM bills')
                all_b = [dict(r) for r in c_sync.fetchall()]
                conn_sync.close()

                if all_tx:
                    mdb.transactions.delete_many({})
                    mdb.transactions.insert_many(all_tx)
                if all_u:
                    mdb.users.delete_many({})
                    mdb.users.insert_many(all_u)
                if all_b:
                    mdb.bills.delete_many({})
                    mdb.bills.insert_many(all_b)
                print(f"[*] MongoDB Atlas collections synchronized successfully ({len(all_tx)} tx, {len(all_u)} users, {len(all_b)} bills)")
        except Exception as e:
            print(f"[!] MongoDB Atlas sync notice: {e}")


def seed_sample_data(cursor):
    today = datetime.now()
    year = today.year
    month = today.month

    def d(day):
        return f"{year:04d}-{month:02d}-{min(max(day, 1), 28):02d}"

    prev_month = 12 if month == 1 else month - 1
    prev_year = year - 1 if month == 1 else year
    def pd(day):
        return f"{prev_year:04d}-{prev_month:02d}-{min(max(day, 1), 28):02d}"

    sample_items = [
        # Previous balance: Total Credit = 18,422.00, Total Debit = 5,000.00 -> Net = 13,422.00
        ('credit', 15000.00, 'Salary', 'Opening Account Reserve', pd(1), 'Bank Transfer', 'Primary financial balance'),
        ('credit', 3422.00, 'Investments', 'Portfolio Liquidation', pd(10), 'Bank Transfer', 'Index returns'),
        ('debit', 3500.00, 'Rent', 'Quarterly Rent Deposit', pd(5), 'Bank Transfer', 'Leasehold'),
        ('debit', 1500.00, 'Money transfer', 'Family Support Transfer', pd(15), 'UPI / Online', 'Direct wire'),

        # Current Month Credits (Income) -> Total: $8,500.00
        ('credit', 5000.00, 'Salary', 'Tech Lead Monthly Salary', d(1), 'Bank Transfer', 'Adaline Lively base pay'),
        ('credit', 1800.00, 'Freelance', 'FinTech UI/UX Design System', d(4), 'Bank Transfer', 'Client project delivery'),
        ('credit', 700.00, 'Investments', 'Dividend Growth Portfolio', d(8), 'Bank Transfer', 'Quarterly payout'),
        ('credit', 500.00, 'Consulting', 'Web Architecture Audit', d(12), 'UPI / Online', 'Advisory session'),
        ('credit', 300.00, 'Bonus', 'Performance Quarterly Incentive', d(14), 'Bank Transfer', 'Bonus award'),
        ('credit', 200.00, 'Refunds', 'Airline Travel Reimbursement', d(15), 'Credit Card', 'Corporate expensing'),

        # Current Month Debits (Expenses) -> Total: $6,222.00
        ('debit', 2500.00, 'Money transfer', 'Wire Transfer & Savings Vault', d(2), 'Bank Transfer', 'High-yield allocation'),
        ('debit', 1400.00, 'Rent', 'Luxury Studio Apartment Rent', d(3), 'Bank Transfer', 'Monthly residential lease'),
        ('debit', 320.00, 'Cafe & Restaurants', 'Bistro Fine Dining & Wine', d(6), 'Credit Card', 'Weekend celebration'),
        ('debit', 215.00, 'Cafe & Restaurants', 'Artisan Roastery & Lunch Meetings', d(9), 'Credit Card', 'Workday coffees'),
        ('debit', 187.00, 'Cafe & Restaurants', 'Japanese Omakase Dinner', d(13), 'Credit Card', 'Team dinner'),
        ('debit', 650.00, 'Education', 'Executive Leadership Masterclass', d(10), 'Debit Card', 'Professional certification'),
        ('debit', 310.00, 'Food & Groceries', 'Whole Foods Organic Groceries', d(7), 'Debit Card', 'Bi-weekly provisions'),
        ('debit', 240.00, 'Food & Groceries', 'Farmers Market Fresh Pantry', d(11), 'Cash', 'Produce & gourmet goods'),
        ('debit', 250.00, 'Others', 'Workspace Dual Monitor Arm', d(12), 'Credit Card', 'Desk setup'),
        ('debit', 150.00, 'Others', 'Streaming & Cloud Subscriptions', d(14), 'Credit Card', 'SaaS subscriptions')
    ]

    cursor.executemany('''
        INSERT INTO transactions (type, amount, category, description, date, payment_method, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', sample_items)


def seed_sample_bills(cursor):
    today = datetime.now()
    year = today.year
    month = today.month

    def bd(day):
        return f"{year:04d}-{month:02d}-{min(max(day, 1), 28):02d}"

    sample_bills = [
        ('Apartment Rent', 1400.00, 3, bd(3), 'Housing', 'paid', 1, bd(3), 'Automatic direct debit'),
        ('Netflix 4K & Cloud Storage', 28.00, 9, bd(9), 'Subscriptions', 'paid', 1, bd(9), 'Streaming entertainment'),
        ('Electricity & Water Utility', 145.00, 15, bd(15), 'Utilities', 'unpaid', 0, None, 'City power & water grid'),
        ('High-Speed Fiber Internet', 75.00, 18, bd(18), 'Utilities', 'unpaid', 1, None, '1Gbps fiber connection'),
        ('Equinox Fitness & Gym', 120.00, 22, bd(22), 'Fitness', 'unpaid', 0, None, 'Monthly membership'),
        ('Health & Dental Coverage', 280.00, 25, bd(25), 'Insurance', 'unpaid', 1, None, 'Premium tier health plan'),
        ('Spotify Family Audio', 16.99, 28, bd(28), 'Subscriptions', 'unpaid', 1, None, 'Family music subscription')
    ]

    cursor.executemany('''
        INSERT INTO bills (name, amount, due_day, due_date, category, status, auto_pay, paid_date, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', sample_bills)


# ==================== HEALTH ====================

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({
        "status": "healthy",
        "service": "FinSet Expense Tracker API",
        "timestamp": datetime.now().isoformat()
    })


# ==================== AUTHENTICATION ====================

@app.route('/api/auth/register', methods=['POST'])
def register():
    data = request.get_json() or {}
    name = data.get('name', '').strip()
    email = data.get('email', '').strip().lower()
    password = data.get('password', '')

    if not name or not email or not password:
        return jsonify({'error': 'Name, email, and password are required'}), 400

    if len(password) < 6:
        return jsonify({'error': 'Password must be at least 6 characters long'}), 400

    password_hash = generate_password_hash(password)
    avatar = f"https://api.dicebear.com/7.x/initials/svg?seed={name}"

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT INTO users (name, email, password_hash, avatar)
            VALUES (?, ?, ?, ?)
        ''', (name, email, password_hash, avatar))
        conn.commit()
        user_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'An account with this email already exists'}), 400

    cursor.execute('SELECT id, name, email, avatar FROM users WHERE id = ?', (user_id,))
    user = dict(cursor.fetchone())
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mongo_user = dict(user)
            mongo_user['password_hash'] = password_hash
            mongo_user['created_at'] = datetime.now().isoformat()
            mdb.users.update_one({'email': email}, {'$set': mongo_user}, upsert=True)
        except Exception as e:
            print(f"[MongoDB] register sync error: {e}")

    return jsonify({
        'user': user,
        'token': f"session_{user_id}_{int(datetime.now().timestamp())}",
        'message': 'Account created successfully'
    }), 201


@app.route('/api/auth/login', methods=['POST'])
def login():
    data = request.get_json() or {}
    email = data.get('email', '').strip().lower()
    password = data.get('password', '')

    if not email or not password:
        return jsonify({'error': 'Email and password are required'}), 400

    mdb = get_mongo_db()
    user_row = None
    if mdb is not None:
        try:
            u_doc = mdb.users.find_one({'email': email})
            if u_doc:
                user_row = clean_doc(u_doc)
        except Exception:
            user_row = None

    if not user_row:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE email = ?', (email,))
        row = cursor.fetchone()
        conn.close()
        if row:
            user_row = dict(row)

    if user_row is None or not check_password_hash(user_row['password_hash'], password):
        return jsonify({'error': 'Invalid email or password'}), 401

    user_dict = {
        'id': user_row['id'],
        'name': user_row['name'],
        'email': user_row['email'],
        'avatar': user_row['avatar']
    }

    return jsonify({
        'user': user_dict,
        'token': f"session_{user_row['id']}_{int(datetime.now().timestamp())}",
        'message': 'Logged in successfully'
    })


@app.route('/api/auth/me', methods=['GET'])
def get_current_user():
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            u_doc = mdb.users.find_one({}, {'_id': 0, 'password_hash': 0}, sort=[('id', 1)])
            if u_doc:
                return jsonify({'user': u_doc})
        except Exception:
            pass

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id, name, email, avatar FROM users ORDER BY id ASC LIMIT 1')
    user_row = cursor.fetchone()
    conn.close()

    if user_row:
        return jsonify({'user': dict(user_row)})
    return jsonify({'user': None})


# ==================== TRANSACTIONS ====================

@app.route('/api/transactions', methods=['GET'])
def get_transactions():
    trans_type = request.args.get('type')
    category = request.args.get('category')
    search = request.args.get('search', '').strip()
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    sort = request.args.get('sort', 'date_desc')

    # Try MongoDB Atlas first
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            filter_q = {}
            if trans_type and trans_type.lower() in ['credit', 'debit']:
                filter_q['type'] = trans_type.lower()
            if category and category.lower() != 'all':
                filter_q['category'] = category
            if start_date or end_date:
                filter_q['date'] = {}
                if start_date:
                    filter_q['date']['$gte'] = start_date
                if end_date:
                    filter_q['date']['$lte'] = end_date
            if search:
                regex = {'$regex': search, '$options': 'i'}
                filter_q['$or'] = [
                    {'description': regex},
                    {'category': regex},
                    {'payment_method': regex},
                    {'notes': regex}
                ]
            
            if sort == 'date_asc':
                sort_spec = [('date', 1), ('id', 1)]
            elif sort == 'amount_desc':
                sort_spec = [('amount', -1), ('date', -1)]
            elif sort == 'amount_asc':
                sort_spec = [('amount', 1), ('date', -1)]
            else:
                sort_spec = [('date', -1), ('id', -1)]
                
            docs = list(mdb.transactions.find(filter_q, {'_id': 0}).sort(sort_spec))
            return jsonify(docs)
        except Exception as e:
            print(f"[MongoDB] Query fallback: {e}")

    # Local SQLite fallback
    query = 'SELECT * FROM transactions WHERE 1=1'
    params = []

    if trans_type and trans_type.lower() in ['credit', 'debit']:
        query += ' AND type = ?'
        params.append(trans_type.lower())

    if category and category.lower() != 'all':
        query += ' AND category = ?'
        params.append(category)

    if start_date:
        query += ' AND date >= ?'
        params.append(start_date)

    if end_date:
        query += ' AND date <= ?'
        params.append(end_date)

    if search:
        query += ' AND (description LIKE ? OR category LIKE ? OR payment_method LIKE ? OR notes LIKE ?)'
        like_search = f"%{search}%"
        params.extend([like_search, like_search, like_search, like_search])

    if sort == 'date_asc':
        query += ' ORDER BY date ASC, id ASC'
    elif sort == 'amount_desc':
        query += ' ORDER BY amount DESC, date DESC'
    elif sort == 'amount_asc':
        query += ' ORDER BY amount ASC, date DESC'
    else:
        query += ' ORDER BY date DESC, id DESC'

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    transactions = [dict(row) for row in rows]
    return jsonify(transactions)


@app.route('/api/transactions/<int:tx_id>', methods=['GET'])
def get_transaction(tx_id):
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            doc = mdb.transactions.find_one({'id': tx_id}, {'_id': 0})
            if doc:
                return jsonify(doc)
        except Exception:
            pass

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM transactions WHERE id = ?', (tx_id,))
    row = cursor.fetchone()
    conn.close()

    if row is None:
        return jsonify({'error': 'Transaction not found'}), 404
    return jsonify(dict(row))


@app.route('/api/transactions', methods=['POST'])
def create_transaction():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'Request body must be valid JSON'}), 400

    tx_type = data.get('type', '').strip().lower()
    amount = data.get('amount')
    category = data.get('category', '').strip()
    description = data.get('description', '').strip()
    date_val = data.get('date', '').strip()
    payment_method = data.get('payment_method', 'Cash').strip()
    notes = data.get('notes', '').strip()

    if tx_type not in ['credit', 'debit']:
        return jsonify({'error': "Transaction type must be 'credit' (income) or 'debit' (expense)"}), 400

    try:
        amount = float(amount)
        if amount <= 0:
            raise ValueError()
    except (TypeError, ValueError):
        return jsonify({'error': 'Amount must be a positive number greater than 0'}), 400

    if not description:
        return jsonify({'error': 'Description is required'}), 400

    if not category:
        category = 'Others'

    if not date_val:
        date_val = datetime.now().strftime('%Y-%m-%d')
    else:
        try:
            datetime.strptime(date_val, '%Y-%m-%d')
        except ValueError:
            return jsonify({'error': 'Date must be formatted as YYYY-MM-DD'}), 400

    conn = get_db_connection()
    cursor = conn.cursor()
    now_time = datetime.now().strftime('%H:%M:%S')
    created_at_val = f"{date_val} {now_time}"
    cursor.execute('''
        INSERT INTO transactions (type, amount, category, description, date, payment_method, notes, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (tx_type, amount, category, description, date_val, payment_method, notes, created_at_val))
    conn.commit()
    new_id = cursor.lastrowid

    cursor.execute('SELECT * FROM transactions WHERE id = ?', (new_id,))
    new_tx = dict(cursor.fetchone())
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.transactions.insert_one(dict(new_tx))
        except Exception as e:
            print(f"[MongoDB] insert transaction error: {e}")

    return jsonify(new_tx), 201


@app.route('/api/transactions/<int:tx_id>', methods=['PUT'])
def update_transaction(tx_id):
    data = request.get_json()
    if not data:
        return jsonify({'error': 'Request body must be valid JSON'}), 400

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM transactions WHERE id = ?', (tx_id,))
    existing = cursor.fetchone()
    if existing is None:
        conn.close()
        return jsonify({'error': 'Transaction not found'}), 404

    existing_dict = dict(existing)

    tx_type = data.get('type', existing_dict['type']).strip().lower()
    category = data.get('category', existing_dict['category']).strip()
    description = data.get('description', existing_dict['description']).strip()
    date_val = data.get('date', existing_dict['date']).strip()
    payment_method = data.get('payment_method', existing_dict['payment_method']).strip()
    notes = data.get('notes', existing_dict.get('notes', '')).strip()

    if tx_type not in ['credit', 'debit']:
        conn.close()
        return jsonify({'error': "Type must be 'credit' or 'debit'"}), 400

    raw_amount = data.get('amount', existing_dict['amount'])
    try:
        amount = float(raw_amount)
        if amount <= 0:
            raise ValueError()
    except (TypeError, ValueError):
        conn.close()
        return jsonify({'error': 'Amount must be greater than 0'}), 400

    if not description:
        conn.close()
        return jsonify({'error': 'Description cannot be empty'}), 400

    try:
        datetime.strptime(date_val, '%Y-%m-%d')
    except ValueError:
        conn.close()
        return jsonify({'error': 'Date must be formatted as YYYY-MM-DD'}), 400

    cursor.execute('''
        UPDATE transactions
        SET type = ?, amount = ?, category = ?, description = ?, date = ?, payment_method = ?, notes = ?
        WHERE id = ?
    ''', (tx_type, amount, category, description, date_val, payment_method, notes, tx_id))
    conn.commit()

    cursor.execute('SELECT * FROM transactions WHERE id = ?', (tx_id,))
    updated_tx = dict(cursor.fetchone())
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.transactions.update_one({'id': tx_id}, {'$set': {
                'type': tx_type, 'amount': amount, 'category': category,
                'description': description, 'date': date_val,
                'payment_method': payment_method, 'notes': notes
            }})
        except Exception as e:
            print(f"[MongoDB] update transaction error: {e}")

    return jsonify(updated_tx)


@app.route('/api/transactions/<int:tx_id>', methods=['DELETE'])
def delete_transaction(tx_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id FROM transactions WHERE id = ?', (tx_id,))
    if cursor.fetchone() is None:
        conn.close()
        return jsonify({'error': 'Transaction not found'}), 404

    cursor.execute('DELETE FROM transactions WHERE id = ?', (tx_id,))
    conn.commit()
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.transactions.delete_one({'id': tx_id})
        except Exception as e:
            print(f"[MongoDB] delete transaction error: {e}")

    return jsonify({'success': True, 'message': f'Transaction {tx_id} deleted successfully'})


# ==================== SUMMARY ====================

@app.route('/api/summary', methods=['GET'])
def get_summary():
    today = datetime.now()
    cur_month_prefix = f"{today.year:04d}-{today.month:02d}"

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT 
            COALESCE(SUM(CASE WHEN type = 'credit' THEN amount ELSE 0 END), 0) AS total_credit,
            COALESCE(SUM(CASE WHEN type = 'debit' THEN amount ELSE 0 END), 0) AS total_debit,
            COUNT(*) AS total_count
        FROM transactions
    ''')
    overall = cursor.fetchone()
    all_time_balance = float(overall['total_credit']) - float(overall['total_debit'])

    cursor.execute('''
        SELECT 
            COALESCE(SUM(CASE WHEN type = 'credit' THEN amount ELSE 0 END), 0) AS month_credit,
            COALESCE(SUM(CASE WHEN type = 'debit' THEN amount ELSE 0 END), 0) AS month_debit,
            COUNT(CASE WHEN type = 'credit' THEN 1 END) AS month_credit_count,
            COUNT(CASE WHEN type = 'debit' THEN 1 END) AS month_debit_count
        FROM transactions
        WHERE date LIKE ?
    ''', (f"{cur_month_prefix}%",))
    cur_month = cursor.fetchone()

    month_credit = float(cur_month['month_credit'])
    month_debit = float(cur_month['month_debit'])

    cursor.execute('''
        SELECT category, SUM(amount) as total, COUNT(*) as count
        FROM transactions
        WHERE type = 'debit' AND date LIKE ?
        GROUP BY category
        ORDER BY total DESC
    ''', (f"{cur_month_prefix}%",))
    debit_rows = cursor.fetchall()

    category_breakdown = []
    for r in debit_rows:
        cat_amount = float(r['total'])
        pct = (cat_amount / month_debit * 100) if month_debit > 0 else 0
        category_breakdown.append({
            'category': r['category'],
            'total': cat_amount,
            'count': r['count'],
            'percentage': round(pct, 1)
        })

    cursor.execute('''
        SELECT category, SUM(amount) as total, COUNT(*) as count
        FROM transactions
        WHERE type = 'credit' AND date LIKE ?
        GROUP BY category
        ORDER BY total DESC
    ''', (f"{cur_month_prefix}%",))
    credit_rows = cursor.fetchall()

    income_breakdown = []
    for r in credit_rows:
        cat_amount = float(r['total'])
        pct = (cat_amount / month_credit * 100) if month_credit > 0 else 0
        income_breakdown.append({
            'category': r['category'],
            'total': cat_amount,
            'count': r['count'],
            'percentage': round(pct, 1)
        })

    cursor.execute('''
        SELECT 
            SUBSTR(date, 1, 7) AS month,
            COALESCE(SUM(CASE WHEN type = 'credit' THEN amount ELSE 0 END), 0) AS credit,
            COALESCE(SUM(CASE WHEN type = 'debit' THEN amount ELSE 0 END), 0) AS debit
        FROM transactions
        GROUP BY SUBSTR(date, 1, 7)
        ORDER BY month ASC
    ''')
    trend_rows = cursor.fetchall()
    monthly_trend = [
        {
            'month': r['month'],
            'credit': float(r['credit']),
            'debit': float(r['debit']),
            'net': float(r['credit']) - float(r['debit'])
        }
        for r in trend_rows
    ]

    cursor.execute('''
        SELECT payment_method, COUNT(*) as count, SUM(amount) as total
        FROM transactions
        GROUP BY payment_method
        ORDER BY total DESC
    ''')
    pm_rows = cursor.fetchall()
    payment_methods = [{'method': r['payment_method'], 'count': r['count'], 'total': float(r['total'])} for r in pm_rows]

    conn.close()

    return jsonify({
        'total_balance': all_time_balance,
        'balance': all_time_balance,
        'total_credit': float(overall['total_credit']),
        'total_debit': float(overall['total_debit']),
        'month_credit': month_credit,
        'month_debit': month_debit,
        'credit_count': cur_month['month_credit_count'],
        'debit_count': cur_month['month_debit_count'],
        'total_count': overall['total_count'],
        'category_breakdown': category_breakdown,
        'income_breakdown': income_breakdown,
        'monthly_trend': monthly_trend,
        'payment_methods': payment_methods
    })


# ==================== MONTHLY BILL MANAGEMENT & CALENDAR ====================

@app.route('/api/bills', methods=['GET'])
def get_bills():
    status = request.args.get('status')
    
    # Try MongoDB Atlas first
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            filter_q = {}
            if status and status != 'all':
                filter_q['status'] = status
            bills = list(mdb.bills.find(filter_q, {'_id': 0}).sort('due_day', 1))
            return jsonify(bills)
        except Exception as e:
            print(f"[MongoDB] bills query fallback: {e}")

    conn = get_db_connection()
    cursor = conn.cursor()

    query = 'SELECT * FROM bills WHERE 1=1'
    params = []
    if status and status != 'all':
        query += ' AND status = ?'
        params.append(status)

    query += ' ORDER BY due_day ASC'
    cursor.execute(query, params)
    bills = [dict(r) for r in cursor.fetchall()]
    conn.close()

    return jsonify(bills)


@app.route('/api/bills', methods=['POST'])
def create_bill():
    data = request.get_json() or {}
    name = data.get('name', '').strip()
    amount = data.get('amount')
    due_day = data.get('due_day')
    category = data.get('category', 'Utilities').strip()
    auto_pay = 1 if data.get('auto_pay') else 0
    notes = data.get('notes', '').strip()

    if not name:
        return jsonify({'error': 'Bill name is required'}), 400

    try:
        amount = float(amount)
        if amount <= 0:
            raise ValueError()
    except (TypeError, ValueError):
        return jsonify({'error': 'Amount must be greater than 0'}), 400

    try:
        due_day = int(due_day)
        if not (1 <= due_day <= 31):
            raise ValueError()
    except (TypeError, ValueError):
        return jsonify({'error': 'Due day must be between 1 and 31'}), 400

    today = datetime.now()
    due_date = f"{today.year:04d}-{today.month:02d}-{min(due_day, 28):02d}"

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO bills (name, amount, due_day, due_date, category, status, auto_pay, notes)
        VALUES (?, ?, ?, ?, ?, 'unpaid', ?, ?)
    ''', (name, amount, due_day, due_date, category, auto_pay, notes))
    conn.commit()
    new_id = cursor.lastrowid
    cursor.execute('SELECT * FROM bills WHERE id = ?', (new_id,))
    new_bill = dict(cursor.fetchone())
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.bills.insert_one(dict(new_bill))
        except Exception as e:
            print(f"[MongoDB] insert bill error: {e}")

    return jsonify(new_bill), 201


@app.route('/api/bills/<int:bill_id>', methods=['PUT'])
def update_bill(bill_id):
    data = request.get_json() or {}
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM bills WHERE id = ?', (bill_id,))
    existing = cursor.fetchone()
    if not existing:
        conn.close()
        return jsonify({'error': 'Bill not found'}), 404

    existing_dict = dict(existing)
    name = data.get('name', existing_dict['name']).strip()
    category = data.get('category', existing_dict['category']).strip()
    status = data.get('status', existing_dict['status']).strip()
    notes = data.get('notes', existing_dict['notes']).strip()
    auto_pay = 1 if data.get('auto_pay', existing_dict['auto_pay']) else 0

    try:
        amount = float(data.get('amount', existing_dict['amount']))
        due_day = int(data.get('due_day', existing_dict['due_day']))
    except (ValueError, TypeError):
        conn.close()
        return jsonify({'error': 'Invalid amount or due day'}), 400

    today = datetime.now()
    due_date = f"{today.year:04d}-{today.month:02d}-{min(due_day, 28):02d}"
    paid_date = datetime.now().strftime('%Y-%m-%d') if status == 'paid' else None

    cursor.execute('''
        UPDATE bills
        SET name = ?, amount = ?, due_day = ?, due_date = ?, category = ?, status = ?, auto_pay = ?, paid_date = ?, notes = ?
        WHERE id = ?
    ''', (name, amount, due_day, due_date, category, status, auto_pay, paid_date, notes, bill_id))
    conn.commit()
    cursor.execute('SELECT * FROM bills WHERE id = ?', (bill_id,))
    updated = dict(cursor.fetchone())
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.bills.update_one({'id': bill_id}, {'$set': {
                'name': name, 'amount': amount, 'due_day': due_day,
                'due_date': due_date, 'category': category, 'status': status,
                'auto_pay': auto_pay, 'paid_date': paid_date, 'notes': notes
            }})
        except Exception as e:
            print(f"[MongoDB] update bill error: {e}")

    return jsonify(updated)


@app.route('/api/bills/<int:bill_id>', methods=['DELETE'])
def delete_bill(bill_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id FROM bills WHERE id = ?', (bill_id,))
    if not cursor.fetchone():
        conn.close()
        return jsonify({'error': 'Bill not found'}), 404

    cursor.execute('DELETE FROM bills WHERE id = ?', (bill_id,))
    conn.commit()
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.bills.delete_one({'id': bill_id})
        except Exception as e:
            print(f"[MongoDB] delete bill error: {e}")

    return jsonify({'success': True, 'message': 'Bill deleted successfully'})


@app.route('/api/bills/<int:bill_id>/pay', methods=['POST'])
def pay_bill(bill_id):
    data = request.get_json() or {}
    record_expense = data.get('record_expense', True)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM bills WHERE id = ?', (bill_id,))
    bill = cursor.fetchone()
    if not bill:
        conn.close()
        return jsonify({'error': 'Bill not found'}), 404

    bill_dict = dict(bill)
    today_str = datetime.now().strftime('%Y-%m-%d')
    now_time = datetime.now().strftime('%H:%M:%S')

    cursor.execute('''
        UPDATE bills
        SET status = 'paid', paid_date = ?
        WHERE id = ?
    ''', (today_str, bill_id))

    new_tx_id = None
    if record_expense:
        cursor.execute('''
            INSERT INTO transactions (type, amount, category, description, date, payment_method, notes, created_at)
            VALUES ('debit', ?, ?, ?, ?, 'Bank Transfer', ?, ?)
        ''', (
            bill_dict['amount'],
            bill_dict['category'],
            f"Bill Payment: {bill_dict['name']}",
            today_str,
            f"Marked paid on {today_str}",
            f"{today_str} {now_time}"
        ))
        new_tx_id = cursor.lastrowid

    conn.commit()
    cursor.execute('SELECT * FROM bills WHERE id = ?', (bill_id,))
    updated = dict(cursor.fetchone())
    conn.close()

    # Sync to MongoDB Atlas
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            mdb.bills.update_one({'id': bill_id}, {'$set': {'status': 'paid', 'paid_date': today_str}})
            if record_expense and new_tx_id:
                tx_doc = {
                    'id': new_tx_id,
                    'type': 'debit',
                    'amount': bill_dict['amount'],
                    'category': bill_dict['category'],
                    'description': f"Bill Payment: {bill_dict['name']}",
                    'date': today_str,
                    'payment_method': 'Bank Transfer',
                    'notes': f"Marked paid on {today_str}",
                    'created_at': f"{today_str} {now_time}"
                }
                mdb.transactions.insert_one(tx_doc)
        except Exception as e:
            print(f"[MongoDB] pay bill error: {e}")

    return jsonify({
        'bill': updated,
        'message': f"Bill '{bill_dict['name']}' marked as paid!"
    })


@app.route('/api/bills/summary', methods=['GET'])
def get_bills_summary():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT 
            COALESCE(SUM(amount), 0) AS total_bills,
            COALESCE(SUM(CASE WHEN status = 'paid' THEN amount ELSE 0 END), 0) AS total_paid,
            COALESCE(SUM(CASE WHEN status != 'paid' THEN amount ELSE 0 END), 0) AS total_pending,
            COUNT(*) AS total_count,
            COUNT(CASE WHEN status = 'paid' THEN 1 END) AS paid_count,
            COUNT(CASE WHEN status != 'paid' THEN 1 END) AS pending_count
        FROM bills
    ''')
    summary_row = cursor.fetchone()

    today = datetime.now()
    cur_day = today.day

    cursor.execute('''
        SELECT * FROM bills
        WHERE status != 'paid' AND due_day >= ?
        ORDER BY due_day ASC LIMIT 1
    ''', (cur_day,))
    next_bill_row = cursor.fetchone()

    if not next_bill_row:
        cursor.execute('''
            SELECT * FROM bills
            WHERE status != 'paid'
            ORDER BY due_day ASC LIMIT 1
        ''')
        next_bill_row = cursor.fetchone()

    next_bill = dict(next_bill_row) if next_bill_row else None
    days_remaining = None
    if next_bill:
        days_remaining = next_bill['due_day'] - cur_day

    conn.close()

    return jsonify({
        'total_bills': float(summary_row['total_bills']),
        'total_paid': float(summary_row['total_paid']),
        'total_pending': float(summary_row['total_pending']),
        'total_count': summary_row['total_count'],
        'paid_count': summary_row['paid_count'],
        'pending_count': summary_row['pending_count'],
        'next_bill': next_bill,
        'days_remaining': days_remaining
    })


# ==================== MONTHLY EXCEL SHEET REPORT ====================

@app.route('/api/report/preview', methods=['GET'])
def get_report_preview():
    today = datetime.now()
    month = request.args.get('month', f"{today.year:04d}-{today.month:02d}")

    conn = get_db_connection()
    cursor = conn.cursor()

    # Available months
    cursor.execute('SELECT DISTINCT SUBSTR(date, 1, 7) as m FROM transactions ORDER BY m DESC')
    available_months = [r['m'] for r in cursor.fetchall() if r['m']]

    filter_clause = " WHERE 1=1"
    params = []
    if month and month != 'all':
        filter_clause += " AND date LIKE ?"
        params.append(f"{month}%")

    # Metrics
    cursor.execute(f'''
        SELECT 
            COALESCE(SUM(CASE WHEN type = 'credit' THEN amount ELSE 0 END), 0) AS total_income,
            COALESCE(SUM(CASE WHEN type = 'debit' THEN amount ELSE 0 END), 0) AS total_expense,
            COUNT(CASE WHEN type = 'credit' THEN 1 END) AS income_count,
            COUNT(CASE WHEN type = 'debit' THEN 1 END) AS expense_count,
            COUNT(*) AS total_count
        FROM transactions
        {filter_clause}
    ''', params)
    kpis = cursor.fetchone()

    total_income = float(kpis['total_income'])
    total_expense = float(kpis['total_expense'])
    net_savings = total_income - total_expense

    # Top categories
    cursor.execute(f'''
        SELECT category, type, SUM(amount) as total, COUNT(*) as count
        FROM transactions
        {filter_clause}
        GROUP BY category, type
        ORDER BY total DESC
    ''', params)
    category_summary = [dict(r) for r in cursor.fetchall()]

    # Transactions list
    cursor.execute(f'''
        SELECT * FROM transactions
        {filter_clause}
        ORDER BY date DESC, id DESC
    ''', params)
    transactions = [dict(r) for r in cursor.fetchall()]

    conn.close()

    return jsonify({
        'month': month,
        'available_months': available_months or [f"{today.year:04d}-{today.month:02d}"],
        'total_income': total_income,
        'total_expense': total_expense,
        'net_savings': net_savings,
        'income_count': kpis['income_count'],
        'expense_count': kpis['expense_count'],
        'total_count': kpis['total_count'],
        'category_summary': category_summary,
        'transactions': transactions[:25]  # preview first 25
    })


@app.route('/api/report/excel', methods=['GET'])
def export_monthly_excel():
    today = datetime.now()
    month = request.args.get('month', f"{today.year:04d}-{today.month:02d}")

    conn = get_db_connection()
    cursor = conn.cursor()

    filter_clause = " WHERE 1=1"
    params = []
    if month and month != 'all':
        filter_clause += " AND date LIKE ?"
        params.append(f"{month}%")

    # Fetch Transactions
    cursor.execute(f'''
        SELECT id, date, type, category, description, amount, payment_method, notes, created_at
        FROM transactions
        {filter_clause}
        ORDER BY date ASC, id ASC
    ''', params)
    transactions = cursor.fetchall()

    # Fetch Category Summary
    cursor.execute(f'''
        SELECT category, type, SUM(amount) as total, COUNT(*) as count
        FROM transactions
        {filter_clause}
        GROUP BY category, type
        ORDER BY type ASC, total DESC
    ''', params)
    cat_rows = cursor.fetchall()

    # Fetch Bills
    cursor.execute('SELECT id, name, due_date, category, amount, status, auto_pay FROM bills ORDER BY due_day ASC')
    bills = cursor.fetchall()

    conn.close()

    # Total income & expense
    total_income = sum(r['amount'] for r in transactions if r['type'] == 'credit')
    total_expense = sum(r['amount'] for r in transactions if r['type'] == 'debit')
    net_flow = total_income - total_expense

    # Create Excel Workbook
    wb = openpyxl.Workbook()

    # Styles
    header_fill = PatternFill(start_color="7B61FF", end_color="7B61FF", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="14171A")
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="5E6470")
    kpi_title_font = Font(name="Calibri", size=10, bold=True, color="5E6470")
    kpi_num_font = Font(name="Calibri", size=14, bold=True, color="14171A")
    kpi_fill = PatternFill(start_color="F8F9FD", end_color="F8F9FD", fill_type="solid")
    green_font = Font(name="Calibri", size=11, bold=True, color="10B981")
    red_font = Font(name="Calibri", size=11, bold=True, color="F43F5E")
    thin_border = Border(
        left=Side(style='thin', color='EAEDF3'),
        right=Side(style='thin', color='EAEDF3'),
        top=Side(style='thin', color='EAEDF3'),
        bottom=Side(style='thin', color='EAEDF3')
    )

    # ---------------- SHEET 1: Transactions Ledger (Default Active Sheet) ----------------
    ws1 = wb.active
    ws1.title = "Transactions Ledger"
    ws1.views.sheetView[0].showGridLines = True

    tx_headers = ["ID", "Transaction Date", "Time", "Type", "Category", "Description", "Amount ($)", "Payment Method", "Notes"]
    for col_idx, h in enumerate(tx_headers, start=1):
        cell = ws1.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, tx in enumerate(transactions, start=2):
        # Format Transaction Date and Time
        tx_date = tx['date'] or ''
        created_at_val = tx['created_at'] or ''
        time_display = "12:00:00 PM"
        if created_at_val:
            raw_str = str(created_at_val).strip()
            if ' ' in raw_str:
                t_part = raw_str.split(' ')[1]
                try:
                    parsed_t = datetime.strptime(t_part.split('.')[0], '%H:%M:%S')
                    time_display = parsed_t.strftime('%I:%M:%S %p')
                except Exception:
                    time_display = t_part
            elif 'T' in raw_str:
                t_part = raw_str.split('T')[1].split('.')[0]
                try:
                    parsed_t = datetime.strptime(t_part, '%H:%M:%S')
                    time_display = parsed_t.strftime('%I:%M:%S %p')
                except Exception:
                    time_display = t_part

        ws1.cell(row=row_idx, column=1, value=tx['id']).alignment = Alignment(horizontal="center")
        ws1.cell(row=row_idx, column=2, value=tx_date).alignment = Alignment(horizontal="center")
        ws1.cell(row=row_idx, column=3, value=time_display).alignment = Alignment(horizontal="center")
        
        type_cell = ws1.cell(row=row_idx, column=4, value=tx['type'].upper())
        type_cell.font = green_font if tx['type'] == 'credit' else red_font
        type_cell.alignment = Alignment(horizontal="center")

        ws1.cell(row=row_idx, column=5, value=tx['category'])
        ws1.cell(row=row_idx, column=6, value=tx['description'])
        
        amt_cell = ws1.cell(row=row_idx, column=7, value=tx['amount'])
        amt_cell.number_format = '"$"#,##0.00'
        amt_cell.font = green_font if tx['type'] == 'credit' else red_font

        ws1.cell(row=row_idx, column=8, value=tx['payment_method'])
        ws1.cell(row=row_idx, column=9, value=tx['notes'])

        for c_i in range(1, 10):
            ws1.cell(row=row_idx, column=c_i).border = thin_border

    # ---------------- SHEET 2: Financial Summary ----------------
    ws2 = wb.create_sheet(title="Financial Summary")
    ws2.views.sheetView[0].showGridLines = True

    # Title
    ws2['A1'] = "FinSet Pro - Monthly Financial Report"
    ws2['A1'].font = title_font
    ws2['A2'] = f"Reporting Period: {month.upper()} | Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws2['A2'].font = subtitle_font

    # KPI Block
    kpis = [
        ("TOTAL INCOME (CREDIT)", total_income, "B4", "C4", green_font),
        ("TOTAL EXPENSES (DEBIT)", total_expense, "D4", "E4", red_font),
        ("NET CASH FLOW", net_flow, "F4", "G4", green_font if net_flow >= 0 else red_font)
    ]

    for label, val, c1, c2, font_color in kpis:
        cell_lbl = ws2[c1]
        cell_lbl.value = label
        cell_lbl.font = kpi_title_font
        cell_lbl.fill = kpi_fill
        cell_lbl.alignment = Alignment(horizontal="center")

        cell_val = ws2[c2]
        cell_val.value = val
        cell_val.number_format = '"$"#,##0.00'
        cell_val.font = font_color
        cell_val.fill = kpi_fill
        cell_val.alignment = Alignment(horizontal="center")

    # Category Breakdown Header
    ws2['A6'] = "CATEGORY BREAKDOWN"
    ws2['A6'].font = Font(name="Calibri", size=12, bold=True, color="14171A")

    cat_headers = ["Category", "Transaction Type", "Total Amount ($)", "Entries Count"]
    for col_idx, h in enumerate(cat_headers, start=1):
        cell = ws2.cell(row=7, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_num = 8
    for c in cat_rows:
        ws2.cell(row=row_num, column=1, value=c['category']).border = thin_border
        ws2.cell(row=row_num, column=2, value=c['type'].upper()).border = thin_border
        amt_cell = ws2.cell(row=row_num, column=3, value=c['total'])
        amt_cell.number_format = '"$"#,##0.00'
        amt_cell.border = thin_border
        ws2.cell(row=row_num, column=4, value=c['count']).border = thin_border
        row_num += 1

    # ---------------- SHEET 3: Monthly Bills Schedule ----------------
    ws3 = wb.create_sheet(title="Monthly Bills")
    ws3.views.sheetView[0].showGridLines = True

    bill_headers = ["ID", "Bill Name", "Due Date", "Category", "Amount ($)", "Status", "Auto-Pay"]
    for col_idx, h in enumerate(bill_headers, start=1):
        cell = ws3.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, b in enumerate(bills, start=2):
        ws3.cell(row=row_idx, column=1, value=b['id']).alignment = Alignment(horizontal="center")
        ws3.cell(row=row_idx, column=2, value=b['name'])
        ws3.cell(row=row_idx, column=3, value=b['due_date']).alignment = Alignment(horizontal="center")
        ws3.cell(row=row_idx, column=4, value=b['category'])
        
        amt_cell = ws3.cell(row=row_idx, column=5, value=b['amount'])
        amt_cell.number_format = '"$"#,##0.00'

        st_cell = ws3.cell(row=row_idx, column=6, value=b['status'].upper())
        st_cell.font = green_font if b['status'] == 'paid' else red_font
        st_cell.alignment = Alignment(horizontal="center")

        ws3.cell(row=row_idx, column=7, value="YES" if b['auto_pay'] else "NO").alignment = Alignment(horizontal="center")

        for c_i in range(1, 8):
            ws3.cell(row=row_idx, column=c_i).border = thin_border

    # Auto-adjust column widths across all sheets
    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Ensure Sheet 1 (Transactions Ledger with Transaction Date) is the default active sheet
    wb.active = ws1

    # Save to in-memory bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"FinSet_Financial_Report_{month}.xlsx"
    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-disposition": f"attachment; filename={filename}"}
    )


# ==================== CATEGORIES & CSV EXPORT ====================

@app.route('/api/categories', methods=['GET'])
def get_categories():
    predefined = {
        'credit': [
            'Salary', 'Freelance', 'Investments', 'Consulting', 'Bonus', 'Refunds', 'Rental Income'
        ],
        'debit': [
            'Money transfer', 'Cafe & Restaurants', 'Rent', 'Education', 'Food & Groceries', 'Utilities', 'Subscriptions', 'Fitness', 'Insurance', 'Others'
        ],
        'payment_methods': [
            'Bank Transfer', 'Credit Card', 'Debit Card', 'Cash', 'UPI / Online'
        ],
        'bill_categories': [
            'Housing', 'Utilities', 'Subscriptions', 'Fitness', 'Insurance', 'Credit & Loans', 'Education', 'Other'
        ]
    }
    return jsonify(predefined)


@app.route('/api/export', methods=['GET'])
def export_csv():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id, date, type, category, description, amount, payment_method, notes, created_at FROM transactions ORDER BY date DESC, id DESC')
    rows = cursor.fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['ID', 'Transaction Date', 'Time', 'Type (Credit/Debit)', 'Category', 'Description', 'Amount', 'Payment Method', 'Notes'])

    for r in rows:
        created_at_val = r['created_at'] or ''
        time_display = "12:00:00 PM"
        if created_at_val:
            raw_str = str(created_at_val).strip()
            if ' ' in raw_str:
                t_part = raw_str.split(' ')[1]
                try:
                    parsed_t = datetime.strptime(t_part.split('.')[0], '%H:%M:%S')
                    time_display = parsed_t.strftime('%I:%M:%S %p')
                except Exception:
                    time_display = t_part
            elif 'T' in raw_str:
                t_part = raw_str.split('T')[1].split('.')[0]
                try:
                    parsed_t = datetime.strptime(t_part, '%H:%M:%S')
                    time_display = parsed_t.strftime('%I:%M:%S %p')
                except Exception:
                    time_display = t_part

        writer.writerow([r['id'], r['date'], time_display, r['type'].upper(), r['category'], r['description'], f"{r['amount']:.2f}", r['payment_method'], r['notes']])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-disposition": f"attachment; filename=fintrack_transactions_{datetime.now().strftime('%Y%m%d')}.csv"}
    )


@app.route('/api/mongodb/status', methods=['GET'])
def mongodb_status():
    mdb = get_mongo_db()
    if mdb is not None:
        try:
            tx_count = mdb.transactions.count_documents({})
            bills_count = mdb.bills.count_documents({})
            users_count = mdb.users.count_documents({})
            return jsonify({
                'connected': True,
                'provider': 'MongoDB Atlas Cloud',
                'cluster': 'cluster0.kh3a2oo.mongodb.net',
                'database': MONGODB_DB_NAME,
                'collections': {
                    'transactions': tx_count,
                    'bills': bills_count,
                    'users': users_count
                },
                'message': 'Connected to MongoDB Atlas'
            })
        except Exception as e:
            return jsonify({
                'connected': False,
                'provider': 'MongoDB Atlas Cloud',
                'error': str(e)
            }), 200
    return jsonify({
        'connected': False,
        'provider': 'Local SQLite',
        'message': 'Operating in offline SQLite mode'
    }), 200


@app.route('/api/reset', methods=['POST'])
def reset_seed():
    init_db(force_reseed=True)
    return jsonify({"status": "reseeded", "message": "Database reset to FinSet demo data in MongoDB Atlas & SQLite"})


# Serve built React frontend if available
@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_frontend(path):
    if path.startswith('api/'):
        return jsonify({'error': 'API endpoint not found'}), 404

    target_file = os.path.join(FRONTEND_DIST, path)
    if path != "" and os.path.exists(target_file) and not os.path.isdir(target_file):
        return send_from_directory(FRONTEND_DIST, path)

    index_file = os.path.join(FRONTEND_DIST, 'index.html')
    if os.path.exists(index_file):
        return send_from_directory(FRONTEND_DIST, 'index.html')

    return jsonify({
        "service": "FinSet Expense & Income Tracker API",
        "status": "ready"
    })


# Initialize database on startup (for both Gunicorn and direct execution)
try:
    init_db()
except Exception as _e:
    print(f"[!] Database startup notice: {_e}")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    debug_mode = os.environ.get("FLASK_DEBUG", "False").lower() in ("true", "1")
    print(f"Flask Expense Tracker Backend running on port {port}")
    app.run(host='0.0.0.0', port=port, debug=debug_mode)
