import psycopg
from models import get_db

def check_database():
    conn = get_db()
    with conn.cursor() as cursor:
        # Check users table
        print("\nUsers in database:")
        cursor.execute("SELECT * FROM users")
        users = cursor.fetchall()
        for user in users:
            print(f"ID: {user['id']}, Username: {user['username']}, Email: {user['email']}, Role: {user['role']}")
        
        # Check uploaded_files table
        print("\nFiles in database:")
        cursor.execute("SELECT * FROM uploaded_files")
        files = cursor.fetchall()
        for file in files:
            print(f"ID: {file['id']}, User ID: {file['user_id']}, Filename: {file['original_filename']}")

if __name__ == "__main__":
    check_database() 