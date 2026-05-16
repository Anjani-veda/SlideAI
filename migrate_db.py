import sqlite3

def update_schema():
    try:
        conn = sqlite3.connect("app.db")
        cur = conn.cursor()

        # Check if phone column exists
        cur.execute("PRAGMA table_info(users)")
        columns = [col[1] for col in cur.fetchall()]

        if "phone" not in columns:
            print("Adding phone column...")
            cur.execute("ALTER TABLE users ADD COLUMN phone TEXT")

        if "otp" not in columns:
            print("Adding otp column...")
            cur.execute("ALTER TABLE users ADD COLUMN otp TEXT")

        conn.commit()
        print("✅ SQLite schema updated successfully!")

    except Exception as e:
        print(f"❌ Error updating schema: {e}")

    finally:
        conn.close()


if __name__ == "__main__":
    update_schema()
