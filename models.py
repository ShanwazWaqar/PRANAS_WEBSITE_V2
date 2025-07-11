# models.py - Complete fixed version

import os
import psycopg
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin
import base64
import pandas as pd
import io
from dotenv import load_dotenv
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load environment variables
load_dotenv()

def get_db():
    """Get database connection"""
    try:
        conn = psycopg.connect(
            os.getenv('DATABASE_URL'),
            row_factory=psycopg.rows.dict_row
        )
        return conn
    except psycopg.Error as e:
        logging.error(f"Database connection error: {str(e)}")
        raise e

def init_db():
    """Initialize the database with updated schema for file and image storage"""
    conn = get_db()
    cur = conn.cursor()
    
    # Create users table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username VARCHAR(80) UNIQUE NOT NULL,
            email VARCHAR(120) UNIQUE NOT NULL,
            password_hash VARCHAR(255) NOT NULL,
            role VARCHAR(20) NOT NULL DEFAULT 'user',
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Create uploaded_files table with file content and image data columns
    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploaded_files (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id),
            original_filename VARCHAR(255) NOT NULL,
            file_type VARCHAR(100),
            file_size BIGINT,
            file_content BYTEA,
            uploaded_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            processed BOOLEAN DEFAULT FALSE,
            processed_at TIMESTAMP,
            processing_type VARCHAR(50),
            ml_plot_data BYTEA,
            pca_plot_data BYTEA,
            cluster_plot_data BYTEA,
            raw_plot_data BYTEA,
            processing_parameters JSONB,
            results_metadata JSONB,
            source VARCHAR(20) DEFAULT 'local',
            file_path VARCHAR(500)
        )
    """)

    # Create system configuration table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS system_config (
            key VARCHAR(100) PRIMARY KEY,
            value TEXT,
            updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    conn.commit()
    conn.close()

class User(UserMixin):
    def __init__(self, id, username, email, password_hash, role='user', created_at=None):
        self.id = id
        self.username = username
        self.email = email
        self.password_hash = password_hash
        self.role = role
        self.created_at = created_at or datetime.now()

    @classmethod
    def get(cls, user_id):
        """Get a user by ID"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT * FROM users WHERE id = %s", (user_id,))
                    user = cur.fetchone()
                    
                if not user:
                    logging.warning(f"User not found: ID {user_id}")
                    return None
                    
                return User(
                    id=user['id'],
                    username=user['username'],
                    email=user['email'],
                    password_hash=user['password_hash'],
                    role=user['role'],
                    created_at=user['created_at']
                )
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving user: {str(e)}")
            return None

    @staticmethod
    def get_all_users():
        """Get all users from the database"""
        try:
            conn = get_db()
            users = []
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT * FROM users ORDER BY id")
                    rows = cur.fetchall()
                    
                for row in rows:
                    users.append(User(
                        id=row['id'],
                        username=row['username'],
                        email=row['email'],
                        password_hash=row['password_hash'],
                        role=row['role'],
                        created_at=row['created_at']
                    ))
                return users
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving all users: {str(e)}")
            return []

    @staticmethod
    def get_by_username(username):
        """Get a user by username"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT * FROM users WHERE username = %s", (username,))
                    user = cur.fetchone()
                    
                if not user:
                    return None
                    
                return User(
                    id=user['id'],
                    username=user['username'],
                    email=user['email'],
                    password_hash=user['password_hash'],
                    role=user['role'],
                    created_at=user['created_at']
                )
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving user by username: {str(e)}")
            return None
    
    @staticmethod
    def get_by_email(email):
        """Get a user by email"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT * FROM users WHERE email = %s", (email,))
                    user = cur.fetchone()
                    
                if not user:
                    return None
                    
                return User(
                    id=user['id'],
                    username=user['username'],
                    email=user['email'],
                    password_hash=user['password_hash'],
                    role=user['role'],
                    created_at=user['created_at']
                )
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving user by email: {str(e)}")
            return None

    @staticmethod
    def create(username, email, password, role='user'):
        """Create a new user"""
        try:
            conn = get_db()
            password_hash = generate_password_hash(password)
            
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        INSERT INTO users (username, email, password_hash, role)
                        VALUES (%s, %s, %s, %s)
                        RETURNING id, created_at
                    """, (username, email, password_hash, role))
                    result = cur.fetchone()
                    conn.commit()
                    
                return User(
                    id=result['id'],
                    username=username,
                    email=email,
                    password_hash=password_hash,
                    role=role,
                    created_at=result['created_at']
                )
            except Exception as e:
                conn.rollback()
                logging.error(f"Error creating user: {str(e)}")
                raise e
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Database connection error: {str(e)}")
            raise e

    def check_password(self, password):
        """Check if the provided password matches the stored hash"""
        return check_password_hash(self.password_hash, password)

    def is_admin(self):
        """Check if the user is an admin"""
        return self.role == 'admin'
        
    def get_id(self):
        """Get the user ID as a string (required by Flask-Login)"""
        return str(self.id)
        
    @property
    def is_authenticated(self):
        """User is authenticated (required by Flask-Login)"""
        return True
        
    @property
    def is_active(self):
        """User is active (required by Flask-Login)"""
        return True
        
    @property
    def is_anonymous(self):
        """User is anonymous (required by Flask-Login)"""
        return False

class UploadedFile:
    """Class for storing and retrieving uploaded files"""
    def __init__(self, id, user_id, original_filename, file_type=None, file_size=None, 
                 file_content=None, uploaded_at=None, processed=False, processed_at=None, 
                 processing_type=None, ml_plot_data=None, pca_plot_data=None, 
                 cluster_plot_data=None, raw_plot_data=None, processing_parameters=None,
                 results_metadata=None, source='local', file_path=None):
        self.id = id
        self.user_id = user_id
        self.original_filename = original_filename
        self.file_type = file_type
        self.file_size = file_size
        self.file_content = file_content
        self.uploaded_at = uploaded_at or datetime.now()
        self.processed = processed
        self.processed_at = processed_at
        self.processing_type = processing_type
        self.ml_plot_data = ml_plot_data
        self.pca_plot_data = pca_plot_data
        self.cluster_plot_data = cluster_plot_data
        self.raw_plot_data = raw_plot_data
        self.processing_parameters = processing_parameters or {}
        self.results_metadata = results_metadata or {}
        self.source = source
        self.file_path = file_path
        
        # Add plot paths for convenience
        self.ml_plot_path = None
        self.pca_plot_path = None
        self.cluster_plot_path = None
        self.raw_plot_path = None

    @staticmethod
    def create(user_id, original_filename, file_content=None, file_type=None, file_size=None, source='local', file_path=None):
        """Create a new file record with file content stored in the database"""
        try:
            conn = get_db()
            try:
                # Check which columns exist in the table
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT column_name 
                        FROM information_schema.columns 
                        WHERE table_name = 'uploaded_files'
                    """)
                    existing_columns = [row['column_name'] for row in cur.fetchall()]
                
                # Build the SQL dynamically based on existing columns
                columns = ['user_id', 'original_filename']
                values = [user_id, original_filename]
                placeholders = ['%s', '%s']
                
                if 'file_content' in existing_columns and file_content is not None:
                    columns.append('file_content')
                    values.append(file_content)
                    placeholders.append('%s')
                    
                if 'file_type' in existing_columns and file_type is not None:
                    columns.append('file_type')
                    values.append(file_type)
                    placeholders.append('%s')
                    
                if 'file_size' in existing_columns and file_size is not None:
                    columns.append('file_size')
                    values.append(file_size)
                    placeholders.append('%s')
                    
                if 'source' in existing_columns:
                    columns.append('source')
                    values.append(source)
                    placeholders.append('%s')
                    
                if 'file_path' in existing_columns and file_path is not None:
                    columns.append('file_path')
                    values.append(file_path)
                    placeholders.append('%s')
                
                # Create SQL statement
                columns_str = ', '.join(columns)
                placeholders_str = ', '.join(placeholders)
                
                with conn.cursor() as cur:
                    sql = f"""
                        INSERT INTO uploaded_files 
                        ({columns_str}) 
                        VALUES ({placeholders_str})
                        RETURNING id, uploaded_at
                    """
                    cur.execute(sql, values)
                    result = cur.fetchone()
                    conn.commit()
                    
                # Create and return the file object
                return UploadedFile(
                    id=result['id'],
                    user_id=user_id,
                    original_filename=original_filename,
                    file_content=file_content,
                    file_type=file_type,
                    file_size=file_size,
                    uploaded_at=result['uploaded_at'],
                    source=source,
                    file_path=file_path
                )
            except Exception as e:
                conn.rollback()
                logging.error(f"Error creating file record: {str(e)}")
                raise e
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Database connection error: {str(e)}")
            raise e

    @staticmethod
    def get(file_id):
        """Get a file record by ID"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT * FROM uploaded_files WHERE id = %s", (file_id,))
                    file = cur.fetchone()
                    
                if not file:
                    logging.warning(f"File record not found: ID {file_id}")
                    return None
                    
                # Create file record
                try:
                    return UploadedFile(
                        id=file['id'],
                        user_id=file['user_id'],
                        original_filename=file['original_filename'],
                        file_type=file.get('file_type'),
                        file_size=file.get('file_size'),
                        file_content=file.get('file_content'),
                        uploaded_at=file.get('uploaded_at'),
                        processed=file.get('processed', False),
                        processed_at=file.get('processed_at'),
                        processing_type=file.get('processing_type'),
                        ml_plot_data=file.get('ml_plot_data'),
                        pca_plot_data=file.get('pca_plot_data'),
                        cluster_plot_data=file.get('cluster_plot_data'),
                        raw_plot_data=file.get('raw_plot_data'),
                        processing_parameters=file.get('processing_parameters'),
                        results_metadata=file.get('results_metadata'),
                        source=file.get('source', 'local'),
                        file_path=file.get('file_path')
                    )
                except Exception as e:
                    logging.error(f"Error creating file object: {str(e)}")
                    # Create a minimal file record with required fields
                    return UploadedFile(
                        id=file['id'],
                        user_id=file['user_id'],
                        original_filename=file['original_filename'],
                        file_type=file.get('file_type', 'unknown'),
                        file_size=file.get('file_size', 0),
                        source=file.get('source', 'local')
                    )
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving file record: {str(e)}")
            return None
        
    def update(self):
        """Update an existing file record in the database"""
        try:
            conn = get_db()
            try:
                # Check which columns exist in the table
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT column_name 
                        FROM information_schema.columns 
                        WHERE table_name = 'uploaded_files'
                    """)
                    existing_columns = [row['column_name'] for row in cur.fetchall()]
                
                # Build update SQL dynamically
                update_parts = []
                params = []
                
                # Standard fields
                if 'processed' in existing_columns:
                    update_parts.append("processed = %s")
                    params.append(self.processed)
                    
                if 'processed_at' in existing_columns:
                    update_parts.append("processed_at = %s")
                    params.append(self.processed_at)
                    
                if 'processing_type' in existing_columns:
                    update_parts.append("processing_type = %s")
                    params.append(self.processing_type)
                
                # Plot data fields
                if 'ml_plot_data' in existing_columns:
                    update_parts.append("ml_plot_data = %s")
                    params.append(self.ml_plot_data)
                    
                if 'pca_plot_data' in existing_columns:
                    update_parts.append("pca_plot_data = %s")
                    params.append(self.pca_plot_data)
                    
                if 'cluster_plot_data' in existing_columns:
                    update_parts.append("cluster_plot_data = %s")
                    params.append(self.cluster_plot_data)
                    
                if 'raw_plot_data' in existing_columns:
                    update_parts.append("raw_plot_data = %s")
                    params.append(self.raw_plot_data)
                
                # Metadata fields
                if 'processing_parameters' in existing_columns:
                    update_parts.append("processing_parameters = %s")
                    params.append(psycopg.Json(self.processing_parameters) if self.processing_parameters else None)
                    
                if 'results_metadata' in existing_columns:
                    update_parts.append("results_metadata = %s")
                    params.append(psycopg.Json(self.results_metadata) if self.results_metadata else None)
                
                # Path fields
                if hasattr(self, 'ml_plot_path') and self.ml_plot_path and 'ml_plot_path' in existing_columns:
                    update_parts.append("ml_plot_path = %s")
                    params.append(self.ml_plot_path)
                    
                if hasattr(self, 'pca_plot_path') and self.pca_plot_path and 'pca_plot_path' in existing_columns:
                    update_parts.append("pca_plot_path = %s")
                    params.append(self.pca_plot_path)
                    
                if hasattr(self, 'cluster_plot_path') and self.cluster_plot_path and 'cluster_plot_path' in existing_columns:
                    update_parts.append("cluster_plot_path = %s")
                    params.append(self.cluster_plot_path)
                    
                if hasattr(self, 'raw_plot_path') and self.raw_plot_path and 'raw_plot_path' in existing_columns:
                    update_parts.append("raw_plot_path = %s")
                    params.append(self.raw_plot_path)
                
                # If no fields to update, return
                if not update_parts:
                    logging.warning("No fields to update")
                    return False
                
                # Build the SQL
                update_sql = f"UPDATE uploaded_files SET {', '.join(update_parts)} WHERE id = %s"
                params.append(self.id)
                
                # Execute update
                with conn.cursor() as cur:
                    cur.execute(update_sql, params)
                    conn.commit()
                    
                return True
            except Exception as e:
                conn.rollback()
                logging.error(f"Error updating file record: {str(e)}")
                return False
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Database connection error: {str(e)}")
            return False
    
    def get_dataframe(self):
        """Convert stored CSV data to pandas DataFrame"""
        if self.file_content:
            try:
                # Convert binary content to string and parse as CSV
                csv_data = self.file_content.decode('utf-8')
                return pd.read_csv(io.StringIO(csv_data))
            except Exception as e:
                logging.error(f"Error converting file to DataFrame: {str(e)}")
        elif self.file_path and os.path.exists(self.file_path):
            try:
                # Read from filesystem if file_content is not available
                return pd.read_csv(self.file_path)
            except Exception as e:
                logging.error(f"Error reading CSV from filesystem: {str(e)}")
        
        return None
            
    @staticmethod
    def get_user_files(user_id):
        """Get all files for a user"""
        try:
            conn = get_db()
            files = []
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT id, user_id, original_filename, file_type, file_size, 
                               uploaded_at, processed, processed_at, processing_type,
                               processing_parameters, results_metadata, source, file_path
                        FROM uploaded_files 
                        WHERE user_id = %s
                        ORDER BY uploaded_at DESC
                    """, (user_id,))
                    rows = cur.fetchall()
                    
                for row in rows:
                    file = UploadedFile(
                        id=row['id'],
                        user_id=row['user_id'],
                        original_filename=row['original_filename'],
                        file_type=row.get('file_type'),
                        file_size=row.get('file_size'),
                        uploaded_at=row.get('uploaded_at'),
                        processed=row.get('processed', False),
                        processed_at=row.get('processed_at'),
                        processing_type=row.get('processing_type'),
                        processing_parameters=row.get('processing_parameters'),
                        results_metadata=row.get('results_metadata'),
                        source=row.get('source', 'local'),
                        file_path=row.get('file_path')
                    )
                    files.append(file)
            finally:
                conn.close()
            
            return files
        except Exception as e:
            logging.error(f"Error retrieving user files: {str(e)}")
            return []
        
    @staticmethod
    def get_all_files():
        """Get all files in the database"""
        try:
            conn = get_db()
            files = []
            try:
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT id, user_id, original_filename, file_type, file_size, 
                               uploaded_at, processed, processed_at, processing_type,
                               processing_parameters, results_metadata, source, file_path
                        FROM uploaded_files 
                        ORDER BY uploaded_at DESC
                    """)
                    rows = cur.fetchall()
                    
                for row in rows:
                    file = UploadedFile(
                        id=row['id'],
                        user_id=row['user_id'],
                        original_filename=row['original_filename'],
                        file_type=row.get('file_type'),
                        file_size=row.get('file_size'),
                        uploaded_at=row.get('uploaded_at'),
                        processed=row.get('processed', False),
                        processed_at=row.get('processed_at'),
                        processing_type=row.get('processing_type'),
                        processing_parameters=row.get('processing_parameters'),
                        results_metadata=row.get('results_metadata'),
                        source=row.get('source', 'local'),
                        file_path=row.get('file_path')
                    )
                    files.append(file)
            finally:
                conn.close()
            
            return files
        except Exception as e:
            logging.error(f"Error retrieving all files: {str(e)}")
            return []
        
    @staticmethod
    def delete(file_id):
        """Delete a file record from the database"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("DELETE FROM uploaded_files WHERE id = %s", (file_id,))
                conn.commit()
                return True
            except Exception as e:
                conn.rollback()
                logging.error(f"Error deleting file: {str(e)}")
                return False
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Database connection error: {str(e)}")
            return False


class SystemConfig:
    """Store system-wide configuration values like the master OneDrive token"""
    
    @staticmethod
    def get_value(key):
        """Get a configuration value from database"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT value FROM system_config WHERE key = %s", (key,))
                    result = cur.fetchone()
                    return result['value'] if result else None
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving system config value for {key}: {str(e)}")
            return None
    
    @staticmethod
    def set_value(key, value):
        """Set a configuration value in database"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    # Use UPSERT pattern (PostgreSQL 9.5+)
                    cur.execute("""
                        INSERT INTO system_config (key, value, updated_at) 
                        VALUES (%s, %s, CURRENT_TIMESTAMP)
                        ON CONFLICT (key) DO UPDATE SET 
                        value = EXCLUDED.value,
                        updated_at = CURRENT_TIMESTAMP
                    """, (key, value))
                conn.commit()
                return True
            except Exception as e:
                conn.rollback()
                logging.error(f"Error setting system config for {key}: {str(e)}")
                return False
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Database connection error in set_value: {str(e)}")
            return False
    
    @staticmethod
    def get_all_values():
        """Get all configuration values"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT key, value, updated_at FROM system_config")
                    results = cur.fetchall()
                    return results
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Error retrieving all system config values: {str(e)}")
            return []
    
    @staticmethod
    def delete_value(key):
        """Delete a configuration value"""
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    cur.execute("DELETE FROM system_config WHERE key = %s", (key,))
                conn.commit()
                return True
            except Exception as e:
                conn.rollback()
                logging.error(f"Error deleting system config for {key}: {str(e)}")
                return False
            finally:
                conn.close()
        except Exception as e:
            logging.error(f"Database connection error in delete_value: {str(e)}")
            return False