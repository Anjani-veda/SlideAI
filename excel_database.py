import mysql.connector
import json
import hashlib

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


# =========================================================
# DB CONNECTION
# =========================================================
def get_connection():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="123456789",
        database="ExcelDB"
    )


# =========================================================
# CREATE TABLE (UPGRADED STRUCTURE)
# =========================================================
def create_users_table():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INT AUTO_INCREMENT PRIMARY KEY,
        username VARCHAR(255) UNIQUE NOT NULL,
        password_hash VARCHAR(255) NOT NULL,
        phone VARCHAR(20) DEFAULT NULL,
        otp VARCHAR(6) DEFAULT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    cur.close()
    conn.close()

def register_user(username, password, phone):
    try:
        conn = get_connection()
        cur = conn.cursor()
        query = "INSERT INTO users (username, password_hash, phone) VALUES (%s, %s, %s)"
        values = (username, hash_password(password), phone)
        cur.execute(query, values)
        conn.commit()
        return True, None
    except mysql.connector.IntegrityError as err:
        # Duplicate entry error (username already exists)
        return False, "duplicate"
    except mysql.connector.Error as err:
        # General DB error
        return False, str(err)
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

def verify_user(username, password):
    try:
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        
        query = "SELECT password_hash FROM users WHERE username = %s"
        cur.execute(query, (username,))
        user = cur.fetchone()
        
        if user and user['password_hash'] == hash_password(password):
            return True
        return False
    except mysql.connector.Error as err:
        return False
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

def save_otp(username, otp):
    try:
        conn = get_connection()
        cur = conn.cursor()
        query = "UPDATE users SET otp = %s WHERE username = %s"
        cur.execute(query, (otp, username))
        conn.commit()
        return cur.rowcount > 0
    except mysql.connector.Error as err:
        return False
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

def verify_otp(username, otp):
    try:
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        query = "SELECT otp FROM users WHERE username = %s"
        cur.execute(query, (username,))
        user = cur.fetchone()
        if user and user['otp'] == otp:
            return True
        return False
    except mysql.connector.Error as err:
        return False
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

def update_password(username, new_password):
    try:
        conn = get_connection()
        cur = conn.cursor()
        query = "UPDATE users SET password_hash = %s, otp = NULL WHERE username = %s"
        cur.execute(query, (hash_password(new_password), username))
        conn.commit()
        return cur.rowcount > 0
    except mysql.connector.Error as err:
        return False
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

def get_user_phone(username):
    try:
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        query = "SELECT phone FROM users WHERE username = %s"
        cur.execute(query, (username,))
        user = cur.fetchone()
        if user:
            return user['phone']
        return None
    except mysql.connector.Error as err:
        return None
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

def create_reports_table():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS reports (
        id INT AUTO_INCREMENT PRIMARY KEY,
        file_name VARCHAR(255),
        file_path TEXT,
        user_name VARCHAR(255),

        total_value FLOAT,
        avg_value FLOAT,
        max_value FLOAT,
        min_value FLOAT,
        records INT,

        insights LONGTEXT,
        chart_count INT,

        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    cur.close()
    conn.close()


# =========================================================
# SAVE REPORT (FULL ANALYTICS)
# =========================================================
def save_report(
    file_name,
    file_path,
    user_name,
    kpis,
    insights,
    charts
):

    conn = get_connection()
    cur = conn.cursor()

    query = """
    INSERT INTO reports (
        file_name,
        file_path,
        user_name,
        total_value,
        avg_value,
        max_value,
        min_value,
        records,
        insights,
        chart_count
    )
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """

    values = (
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
    )

    cur.execute(query, values)
    conn.commit()

    cur.close()
    conn.close()

    return True


# =========================================================
# FETCH REPORT HISTORY
# =========================================================
def get_reports():
    conn = get_connection()
    cur = conn.cursor(dictionary=True)

    cur.execute("SELECT * FROM reports ORDER BY created_at DESC")
    data = cur.fetchall()

    cur.close()
    conn.close()

    return data


# =========================================================
# INIT
# =========================================================
if __name__ == "__main__":
    create_users_table()
    create_reports_table()
    print("Users and Reports tables ready")