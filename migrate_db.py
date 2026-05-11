import mysql.connector

def update_schema():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="123456789",
            database="ExcelDB"
        )
        cur = conn.cursor()
        
        # Check if phone column exists
        cur.execute("SHOW COLUMNS FROM users LIKE 'phone'")
        if not cur.fetchone():
            print("Adding phone column...")
            cur.execute("ALTER TABLE users ADD COLUMN phone VARCHAR(20) DEFAULT NULL AFTER password_hash")
        
        # Check if otp column exists
        cur.execute("SHOW COLUMNS FROM users LIKE 'otp'")
        if not cur.fetchone():
            print("Adding otp column...")
            cur.execute("ALTER TABLE users ADD COLUMN otp VARCHAR(6) DEFAULT NULL AFTER phone")
            
        conn.commit()
        print("✅ Database schema updated successfully!")
    except Exception as e:
        print(f"❌ Error updating schema: {e}")
    finally:
        if 'cur' in locals():
            cur.close()
        if 'conn' in locals():
            conn.close()

if __name__ == "__main__":
    update_schema()
