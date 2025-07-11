import os
import sys
import argparse
from werkzeug.security import generate_password_hash
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Add the current directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

# Import database functions
try:
    from models import get_db
except ImportError:
    logging.error("Could not import get_db from models. Make sure you're running this script from the project root.")
    sys.exit(1)

def create_admin_user(username, email, password):
    """Create an admin user directly in the database"""
    
    # Generate password hash
    password_hash = generate_password_hash(password)
    
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                # Check if user already exists
                cur.execute("SELECT id FROM users WHERE username = %s", (username,))
                existing_user = cur.fetchone()
                
                if existing_user:
                    logging.info(f"User '{username}' already exists")
                    
                    # Update to admin role
                    cur.execute("UPDATE users SET role = 'admin' WHERE username = %s RETURNING id", (username,))
                    user_id = cur.fetchone()['id']
                    conn.commit()
                    logging.info(f"Updated user '{username}' (ID: {user_id}) to admin role")
                    return user_id
                else:
                    # Create new admin user
                    cur.execute("""
                        INSERT INTO users (username, email, password_hash, role) 
                        VALUES (%s, %s, %s, 'admin')
                        RETURNING id
                    """, (username, email, password_hash))
                    user_id = cur.fetchone()['id']
                    conn.commit()
                    logging.info(f"Created new admin user '{username}' with ID: {user_id}")
                    return user_id
        except Exception as e:
            conn.rollback()
            logging.error(f"Database error: {str(e)}")
            return None
        finally:
            conn.close()
    except Exception as e:
        logging.error(f"Connection error: {str(e)}")
        return None

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Create an admin user')
    parser.add_argument('--username', '-u', required=True, help='Admin username')
    parser.add_argument('--email', '-e', required=True, help='Admin email')
    parser.add_argument('--password', '-p', required=True, help='Admin password')
    
    args = parser.parse_args()
    
    logging.info(f"Creating admin user '{args.username}'...")
    user_id = create_admin_user(args.username, args.email, args.password)
    
    if user_id:
        logging.info(f"Successfully created/updated admin user with ID: {user_id}")
        logging.info(f"You can now log in at http://localhost:5000/login with:")
        logging.info(f"Username: {args.username}")
        logging.info(f"Password: (the password you provided)")
    else:
        logging.error("Failed to create admin user")