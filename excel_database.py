import sqlite3
import json
import hashlib

# =========================================================
# PASSWORD HASHING
# =========================================================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


# =========================================================
# DB CONNECTION (SQLite)
# =========================================================
def get_connection():
    conn = sqlite3.connect("app.db", check_same_thread=False)
    return conn


# =========================================================
# CREATE TABLES
# =========================================================
def create_users_table():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        phone TEXT,
        otp TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    conn.close()


def create_reports_table():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        file_name TEXT,
        file_path TEXT,
        user_name TEXT,

        total_value REAL,
        avg_value REAL,
        max_value REAL,
        min_value REAL,
        records INTEGER,

        insights TEXT,
        chart_count INTEGER,

        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    conn.close()


# =========================================================
# USER FUNCTIONS
# =========================================================
def register_user(username, password, phone):
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
        INSERT INTO users (username, password_hash, phone)
        VALUES (?, ?, ?)
        """, (username, hash_password(password), phone))

        conn.commit()
        return True, None

    except sqlite3.IntegrityError:
        return False, "duplicate"

    finally:
        conn.close()


def verify_user(username, password):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT password_hash FROM users WHERE username = ?", (username,))
    user = cur.fetchone()

    conn.close()

    if user and user[0] == hash_password(password):
        return True
    return False


def save_otp(username, otp):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("UPDATE users SET otp = ? WHERE username = ?", (otp, username))
    conn.commit()

    success = cur.rowcount > 0
    conn.close()
    return success


def verify_otp(username, otp):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT otp FROM users WHERE username = ?", (username,))
    user = cur.fetchone()

    conn.close()

    if user and user[0] == otp:
        return True
    return False


def update_password(username, new_password):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    UPDATE users
    SET password_hash = ?, otp = NULL
    WHERE username = ?
    """, (hash_password(new_password), username))

    conn.commit()

    success = cur.rowcount > 0
    conn.close()
    return success


def get_user_phone(username):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT phone FROM users WHERE username = ?", (username,))
    user = cur.fetchone()

    conn.close()

    return user[0] if user else None


# =========================================================
# REPORT FUNCTIONS
# =========================================================
def save_report(file_name, file_path, user_name, kpis, insights, charts):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    INSERT INTO reports (
        file_name, file_path, user_name,
        total_value, avg_value, max_value, min_value, records,
        insights, chart_count
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        file_name,
        file_path,
        user_name,
        kpis.get("Total Value", 0.0),
        kpis.get("Average", 0.0),
        kpis.get("Max Value", 0.0),
        kpis.get("Min Value", 0.0),
        kpis.get("Records", 0),
        json.dumps(insights),
        len(charts)
    ))

    conn.commit()
    conn.close()

    return True


def get_reports():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM reports ORDER BY created_at DESC")
    rows = cur.fetchall()

    conn.close()
    return rows


# =========================================================
# INIT
# =========================================================
if __name__ == "__main__":
    create_users_table()
    create_reports_table()
    print("SQLite DB ready (app.db)")
