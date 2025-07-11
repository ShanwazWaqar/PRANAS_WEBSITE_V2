from flask import Flask, render_template, request, redirect, url_for, flash, session, send_from_directory, jsonify, send_file
from flask_login import LoginManager, current_user, login_user, logout_user, login_required
from werkzeug.utils import secure_filename
import logging
import requests
import pickle
import numpy as np
import sklearn
from sklearn.preprocessing import StandardScaler
import pandas as pd
from io import StringIO, BytesIO
import os
import uuid
import base64
from datetime import datetime, timedelta  
from werkzeug.serving import WSGIRequestHandler
from dotenv import load_dotenv
import base64
from io import BytesIO, StringIO
import json
from msal import ConfidentialClientApplication
from flask_session import Session
import urllib.parse
from models import SystemConfig
from werkzeug.serving import WSGIRequestHandler
import shutil
import glob
from pathlib import Path
import threading
import time
import atexit
import gc

import matplotlib
matplotlib.use('Agg')

# Load environment variables
load_dotenv()

# Import database models
from models import User, UploadedFile, init_db, get_db
from auth import auth, admin_required

# Import your existing analysis modules
import Scripts.Data_Processor_1 as dp1
import Scripts.Script_Dimensionality_clustering as clustering
import Scripts.ML as ML
import Scripts.visual as VSUL
from Scripts.Spirometry_results import svc_calc, mvv_calc, fvc_calc, sm_calc, raw_data_calc
from Scripts.current_analysis_Rdata_visual import main_function_visual
from Scripts.current_analysis_PCA import main_function_pca
from Scripts.current_analysis_ML import plotter_svm

# IMPORTANT: Increase request sizes BEFORE creating Flask app
WSGIRequestHandler.max_request_line = 16384
WSGIRequestHandler.max_http_header_size = 65536

app = Flask(__name__)

# App configuration - MUST be immediately after app creation
app.secret_key = os.getenv('SECRET_KEY', os.urandom(24))
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max upload
app.config['MAX_CONTENT_PATH'] = 200 * 1024 * 1024

# Session configuration - IMPORTANT: Use filesystem for large data
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_FILE_DIR'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'flask_session')
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=2)

# Create session directory if it doesn't exist
os.makedirs(app.config['SESSION_FILE_DIR'], exist_ok=True)

# Initialize Flask-Session AFTER configuration
Session(app)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)

# Log the configuration to verify
logging.info("=== Flask App Configuration ===")
logging.info(f"MAX_CONTENT_LENGTH: {app.config['MAX_CONTENT_LENGTH']}")
logging.info(f"SESSION_TYPE: {app.config['SESSION_TYPE']}")
logging.info(f"SESSION_FILE_DIR: {app.config['SESSION_FILE_DIR']}")
logging.info(f"WSGIRequestHandler.max_request_line: {WSGIRequestHandler.max_request_line}")
logging.info(f"WSGIRequestHandler.max_http_header_size: {WSGIRequestHandler.max_http_header_size}")




# 1. First, update your app configuration:
CLIENT_ID = os.getenv('MICROSOFT_APP_ID')
CLIENT_SECRET = os.getenv('MICROSOFT_APP_PASSWORD')
REDIRECT_URI = os.getenv('REDIRECT_URI', 'http://localhost:5000/onedrive-callback')

# For personal accounts, we need to use the consumer OAuth endpoint
AUTHORITY = "https://login.microsoftonline.com/consumers"
AUTHORITY_FALLBACK = "https://login.live.com/oauth20_authorize.srf"  # Alternative for personal accounts
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
TOKEN_URL_FALLBACK = "https://login.live.com/oauth20_token.srf"  # Alternative token endpoint

GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
ONEDRIVE_API_ENDPOINT = 'https://api.onedrive.com/v1.0'
SCOPE = ["Files.Read.All", "Files.ReadWrite.All", "User.Read", "onedrive.readonly", "wl.signin"]

# Create necessary directories if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('instance', exist_ok=True)
os.makedirs('Plots', exist_ok=True)
os.makedirs('temp_processing', exist_ok=True)

# Initialize database
init_db()

# Initialize login manager
login_manager = LoginManager()
login_manager.login_view = 'auth.login'
login_manager.init_app(app)

app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_PERMANENT'] = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=1)
Session(app)

@login_manager.user_loader
def load_user(user_id):
    return User.get(int(user_id))

# Add this new function to your app.py (it's an improved version of upload_masterdata_to_onedrive)

def upload_masterdata_to_onedrive_improved(local_masterdata_path, folder_name, access_token):
    """Upload the generated master data file back to OneDrive with improved naming"""
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        username = current_user.username
        
        # Read the master data file
        with open(local_masterdata_path, 'rb') as f:
            file_content = f.read()
        
        # Find user folder in OneDrive
        user_folder_id = None
        response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
        
        if response.status_code != 200:
            response = requests.get('https://api.onedrive.com/v1.0/drive/root/children', headers=headers)
        
        if response.status_code == 200:
            root_items = response.json().get('value', [])
            for item in root_items:
                if item.get('name') == username and item.get('folder'):
                    user_folder_id = item.get('id')
                    break
        
        if not user_folder_id:
            logging.error("User folder not found in OneDrive")
            return False
        
        # Create or find processed_data folder
        results_folder_id = None
        folder_response = requests.get(f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children', headers=headers)
        
        if folder_response.status_code != 200:
            folder_response = requests.get(f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children', headers=headers)
        
        if folder_response.status_code == 200:
            folder_items = folder_response.json().get('value', [])
            for item in folder_items:
                if item.get('name') == 'processed_data' and item.get('folder'):
                    results_folder_id = item.get('id')
                    break
        
        # Create processed_data folder if it doesn't exist
        if not results_folder_id:
            create_folder_data = {
                'name': 'processed_data',
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        results_folder_id = create_response.json().get('id')
                        break
                except Exception as e:
                    logging.error(f"Error creating folder with {endpoint}: {str(e)}")
                    continue
        
        if not results_folder_id:
            logging.error("Could not create or find processed_data folder")
            return False
        
        # Create filename with IMPROVED timestamp format (matches the local file)
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"{folder_name}_Masterdata_{timestamp}.csv"
        
        # Upload the master data file
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'text/csv'
        }
        
        upload_success = False
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{results_folder_id}:/{filename}:/content',
            f'https://api.onedrive.com/v1.0/drive/items/{results_folder_id}:/{filename}:/content'
        ]:
            try:
                upload_response = requests.put(
                    endpoint,
                    headers=upload_headers,
                    data=file_content
                )
                
                if upload_response.status_code in [200, 201]:
                    logging.info(f"Successfully uploaded master data to OneDrive: {filename}")
                    upload_success = True
                    break
                else:
                    logging.warning(f"Upload failed with {endpoint}: {upload_response.status_code}")
            except Exception as e:
                logging.error(f"Error uploading with {endpoint}: {str(e)}")
                continue
        
        return upload_success
            
    except Exception as e:
        logging.error(f"Error uploading master data to OneDrive: {str(e)}")
        return False
    

def ensure_temp_directories():
    """Ensure all required temporary directories exist"""
    temp_dirs = [
        'temp_downloads',
        'temp_processing',
        'uploads',
        'Plots'
    ]
    
    for temp_dir in temp_dirs:
        dir_path = os.path.join(os.getcwd(), temp_dir)
        try:
            os.makedirs(dir_path, exist_ok=True)
        except Exception as e:
            logging.error(f"Could not create directory {dir_path}: {str(e)}")

def cleanup_old_temp_files():
    """Background task to clean up old temporary files"""
    try:
        current_time = time.time()
        
        # Clean up temp_downloads folder
        temp_downloads_dir = os.path.join(os.getcwd(), 'temp_downloads')
        if os.path.exists(temp_downloads_dir):
            for file_path in Path(temp_downloads_dir).glob('*'):
                try:
                    if file_path.is_file():
                        file_time = os.path.getmtime(file_path)
                        # Remove files older than 2 hours
                        if current_time - file_time > 7200:
                            os.remove(file_path)
                            logging.info(f"Cleaned up old temp file: {file_path}")
                except Exception as e:
                    logging.warning(f"Could not clean up temp file {file_path}: {str(e)}")
        
        # Clean up temp_processing folders
        temp_processing_dir = os.path.join(os.getcwd(), 'temp_processing')
        if os.path.exists(temp_processing_dir):
            for folder_path in Path(temp_processing_dir).glob('*'):
                try:
                    if folder_path.is_dir():
                        folder_time = os.path.getmtime(folder_path)
                        # Remove folders older than 1 hour
                        if current_time - folder_time > 3600:
                            shutil.rmtree(folder_path, ignore_errors=True)
                            logging.info(f"Cleaned up old temp folder: {folder_path}")
                except Exception as e:
                    logging.warning(f"Could not clean up temp folder {folder_path}: {str(e)}")
                    
    except Exception as e:
        logging.error(f"Error in cleanup_old_temp_files: {str(e)}")

def initialize_app_directories_and_cleanup():
    """Initialize app directories and start cleanup scheduler"""
    ensure_temp_directories()
    cleanup_old_temp_files()  # Clean up any existing old files
    logging.info("Initialized app directories and cleaned up old temp files")

# Add cleanup on app shutdown
def cleanup_on_shutdown():
    """Clean up temporary files when the app shuts down"""
    try:
        logging.info("App shutting down, cleaning up temporary files...")
        cleanup_old_temp_files()
    except Exception as e:
        logging.error(f"Error during shutdown cleanup: {str(e)}")

# Register shutdown cleanup
atexit.register(cleanup_on_shutdown)

def ensure_source_column_exists():
    """Make sure the source column exists in the uploaded_files table"""
    try:
        conn = get_db()
        try:
            with conn.cursor() as cur:
                # Check if column exists
                cur.execute("""
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name = 'uploaded_files' AND column_name = 'source'
                """)
                if not cur.fetchone():
                    # Add the column if it doesn't exist
                    cur.execute("""
                        ALTER TABLE uploaded_files 
                        ADD COLUMN source VARCHAR(20) DEFAULT 'local'
                    """)
                    conn.commit()
                    logging.info("Added 'source' column to uploaded_files table")
        except Exception as e:
            conn.rollback()
            logging.error(f"Error checking/adding source column: {str(e)}")
        finally:
            conn.close()
    except Exception as e:
        logging.error(f"Database connection error in ensure_source_column_exists: {str(e)}")

# Call this function at app startup, add it before or after init_db():
ensure_source_column_exists()

# Add current datetime to all templates
@app.context_processor
def inject_now():
    return {'now': datetime.now()}

def upload_to_onedrive(file_content, filename, user_id=None, username=None):
    """
    Upload a file to the user's folder in OneDrive
    
    Args:
        file_content: Binary content of the file
        filename: Name to give the file in OneDrive
        user_id: User ID (optional, defaults to current_user.id)
        username: Username (optional, defaults to current_user.username)
        
    Returns:
        dict: OneDrive response or None on failure
    """
    if user_id is None and current_user.is_authenticated:
        user_id = current_user.id
        
    if username is None and current_user.is_authenticated:
        username = current_user.username
        
    if not username:
        logging.error("Cannot upload to OneDrive without a username")
        return None
    
    # Get the master OneDrive token
    access_token = get_master_onedrive_token()
    if not access_token:
        logging.error("No valid OneDrive token available")
        return None
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    # Step 1: Find or create the user folder
    try:
        # Get all items in root
        root_url = "https://api.onedrive.com/v1.0/drive/root/children"
        response = requests.get(root_url, headers=headers)
        
        # Try MS Graph API if that fails
        if response.status_code != 200:
            root_url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
            response = requests.get(root_url, headers=headers)
            
        user_folder_id = None
        
        if response.status_code == 200:
            # Look for the user folder
            root_items = response.json().get("value", [])
            for item in root_items:
                if item.get("name") == username and item.get("folder") is not None:
                    user_folder_id = item.get("id")
                    break
                    
            # Create the folder if it doesn't exist
            if not user_folder_id:
                create_url = "https://api.onedrive.com/v1.0/drive/root/children"
                create_data = {
                    "name": username,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "rename"
                }
                
                create_response = requests.post(
                    create_url,
                    headers=headers,
                    json=create_data
                )
                
                # Try MS Graph API if that fails
                if create_response.status_code not in [200, 201]:
                    create_url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
                    create_response = requests.post(
                        create_url,
                        headers=headers,
                        json=create_data
                    )
                    
                if create_response.status_code in [200, 201]:
                    user_folder_id = create_response.json().get("id")
                else:
                    logging.error(f"Failed to create user folder in OneDrive: {create_response.status_code}")
                    return None
            
            # Step 2: Upload file to the user folder
            if user_folder_id:
                # For small files (< 4MB), we can use a simple upload
                if len(file_content) < 4 * 1024 * 1024:
                    # Simple upload
                    upload_url = f"https://api.onedrive.com/v1.0/drive/items/{user_folder_id}:/{filename}:/content"
                    
                    upload_headers = {
                        'Authorization': f'Bearer {access_token}',
                        'Content-Type': 'application/octet-stream'
                    }
                    
                    upload_response = requests.put(
                        upload_url,
                        headers=upload_headers,
                        data=file_content
                    )
                    
                    # Try MS Graph API if that fails
                    if upload_response.status_code not in [200, 201]:
                        upload_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}:/{filename}:/content"
                        upload_response = requests.put(
                            upload_url,
                            headers=upload_headers,
                            data=file_content
                        )
                        
                    if upload_response.status_code in [200, 201]:
                        logging.info(f"Successfully uploaded file to OneDrive: {filename}")
                        return upload_response.json()
                    else:
                        logging.error(f"Failed to upload file to OneDrive: {upload_response.status_code}")
                        return None
                else:
                    # For larger files, we need to use an upload session
                    logging.warning("Large file upload to OneDrive not implemented yet")
                    return None
        else:
            logging.error(f"Failed to access OneDrive: {response.status_code}")
            return None
    except Exception as e:
        logging.error(f"Error uploading to OneDrive: {str(e)}")
        return None
    
def find_masterdata_file(base_path, max_depth=5):
    """
    Recursively search for Masterdata files in the base path and subdirectories
    
    Based on actual file structure: Data_files/raw_data/masterData_YYYY_MM_DD
    Also maintains backward compatibility with legacy patterns
    
    Args:
        base_path: Root directory to search in
        max_depth: Maximum depth to search (prevents infinite recursion)
    
    Returns:
        str: Path to masterdata file if found, None otherwise
    """
    try:
        base_path = Path(base_path)
        
        # Get today's date and yesterday's date for pattern matching
        today = datetime.now()
        yesterday = today - timedelta(days=1)
        
        date_patterns = [
            today.strftime('%Y_%m_%d'),           # 2025_05_26
            today.strftime('%Y-%m-%d'),           # 2025-05-26
            today.strftime('%Y%m%d'),             # 20250526
            yesterday.strftime('%Y_%m_%d'),       # Yesterday
            yesterday.strftime('%Y-%m-%d'),
            yesterday.strftime('%Y%m%d'),
        ]
        
        # Comprehensive list of possible filenames
        possible_names = []
        
        # NEW: Date-based patterns (current issue)
        for pattern in date_patterns:
            possible_names.extend([
                f'masterData_{pattern}',           # masterData_2025_05_26
                f'masterData_{pattern}.csv',       # masterData_2025_05_26.csv
                f'masterdata_{pattern}',           # lowercase
                f'masterdata_{pattern}.csv',
                f'MasterData_{pattern}',           # Title case
                f'MasterData_{pattern}.csv',
            ])
        
        # LEGACY: Original patterns (maintain compatibility)
        possible_names.extend([
            'Masterdata.csv',
            'masterdata.csv', 
            'MasterData.csv',
            'MASTERDATA.csv',
            'master_data.csv',
            'Master_Data.csv',
            'MasterDataSet.csv',
            'masterdataset.csv',
            'masterData.csv',
            'masterData',                          # Without extension
            'masterdata',
            'MasterData',
            'Masterdata_processed.csv',
            'processed_masterdata.csv'
        ])
        
        logging.info(f"🔍 Searching for masterdata file in: {base_path}")
        logging.info(f"📁 Base path exists: {base_path.exists()}")
        logging.info(f"📅 Looking for date patterns: {date_patterns[:3]}...")
        
        if not base_path.exists():
            logging.error(f"❌ Base path does not exist: {base_path}")
            return None
        
        # PRIORITY SEARCH LOCATIONS - Updated order
        priority_search_patterns = [
            # HIGHEST PRIORITY: Current file structure (Data_files/raw_data)
            base_path / "Data_files" / "raw_data",
            base_path / "data_files" / "raw_data",
            base_path / "Data_Files" / "Raw_Data",
            base_path / "DATA_FILES" / "RAW_DATA",
            
            # MEDIUM PRIORITY: Alternative structures
            base_path / "raw_data",
            base_path / "Raw_Data",
            base_path / "RAW_DATA",
            
            # LEGACY PRIORITY: Original patterns (for existing setups)
            base_path / "NewFolder" / "masterData",
            base_path / "NewFolder" / "MasterData", 
            base_path / "NewFolder" / "masterdata",
            base_path / "newfolder" / "masterdata",
            base_path / "NewFolder" / "master_data",
            
            # COMMON OUTPUT DIRECTORIES
            base_path / "output",
            base_path / "results",
            base_path / "processed",
            base_path / "data",
            base_path / "Output",
            base_path / "Results",
            base_path / "Processed",
            base_path / "Data",
            base_path / "csv_output",
            
            # BASE PATH (direct)
            base_path,
        ]
        
        # Search priority locations first
        for search_dir in priority_search_patterns:
            if search_dir.exists() and search_dir.is_dir():
                logging.info(f"🔍 Checking priority directory: {search_dir}")
                
                # List contents for debugging
                try:
                    contents = list(search_dir.iterdir())
                    relevant_files = [item for item in contents if 
                                    item.is_file() and (
                                        'master' in item.name.lower() or 
                                        item.suffix.lower() == '.csv'
                                    )]
                    logging.info(f"📂 Directory contents ({len(contents)} items, {len(relevant_files)} relevant files)")
                    
                    if relevant_files:
                        logging.info(f"📄 Relevant files: {[f.name for f in relevant_files[:5]]}")
                        
                except Exception as e:
                    logging.warning(f"Could not list directory contents: {str(e)}")
                
                # Check for exact filename matches
                for filename in possible_names:
                    candidate_path = search_dir / filename
                    if candidate_path.exists() and candidate_path.is_file():
                        file_size = candidate_path.stat().st_size
                        logging.info(f"✅ FOUND masterdata file: {candidate_path} ({file_size} bytes)")
                        
                        # Validate file has content
                        if file_size > 0:
                            try:
                                # Quick validation - try to read first line
                                with open(candidate_path, 'r', encoding='utf-8') as f:
                                    first_line = f.readline().strip()
                                    if first_line:
                                        logging.info(f"✅ File validation passed. First line: {first_line[:100]}...")
                                        return str(candidate_path)
                            except UnicodeDecodeError:
                                # Might be binary, but if size > 0 and name matches, likely correct
                                if 'master' in filename.lower():
                                    logging.info(f"✅ Found binary masterdata file: {candidate_path}")
                                    return str(candidate_path)
                            except Exception as e:
                                logging.warning(f"⚠️ Could not validate file content: {str(e)}")
                                # Still return if it matches our pattern
                                if 'master' in filename.lower():
                                    return str(candidate_path)
                        else:
                            logging.warning(f"⚠️ File exists but is empty: {candidate_path}")
        
        # COMPREHENSIVE PATTERN-BASED SEARCH
        logging.info(f"🔍 Exact matches not found, searching with patterns...")
        
        found_candidates = []
        
        try:
            for root, dirs, files in os.walk(str(base_path)):
                try:
                    root_path = Path(root)
                    depth = len(root_path.relative_to(base_path).parts)
                    
                    if depth <= max_depth:
                        for file in files:
                            file_path = os.path.join(root, file)
                            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                            file_lower = file.lower()
                            
                            # PRIORITY SCORING SYSTEM
                            priority = 0
                            reason = ""
                            
                            # HIGHEST: Exact date pattern match (masterData_YYYY_MM_DD)
                            if file.startswith('masterData_') and any(pattern in file for pattern in date_patterns):
                                priority = 5
                                reason = "Exact date pattern match"
                            # HIGH: Master with date pattern
                            elif 'master' in file_lower and any(pattern in file for pattern in date_patterns):
                                priority = 4
                                reason = "Master with date pattern"
                            # MEDIUM-HIGH: Exact legacy filename match
                            elif file in ['Masterdata.csv', 'masterdata.csv', 'MasterData.csv']:
                                priority = 3
                                reason = "Exact legacy filename"
                            # MEDIUM: Master CSV or no extension
                            elif 'master' in file_lower and (file_lower.endswith('.csv') or '.' not in file):
                                priority = 2
                                reason = "Master CSV or no extension"
                            # LOW: Any file with "master"
                            elif 'master' in file_lower:
                                priority = 1
                                reason = "Contains master"
                            
                            if priority > 0:
                                found_candidates.append((file_path, file_size, priority, reason))
                                
                except Exception as e:
                    logging.warning(f"Error processing directory {root}: {str(e)}")
        
        except Exception as e:
            logging.error(f"Error in pattern search: {str(e)}")
        
        # Sort and select best candidate
        if found_candidates:
            # Sort by priority (highest first), then by file size (largest first)
            found_candidates.sort(key=lambda x: (x[2], x[1]), reverse=True)
            
            logging.info(f"🎯 Found {len(found_candidates)} potential masterdata files:")
            for path, size, priority, reason in found_candidates[:10]:
                rel_path = os.path.relpath(path, str(base_path)) if len(path) > 50 else path
                logging.info(f"   📄 {rel_path} ({size} bytes) - P{priority}: {reason}")
            
            # Return the best candidate
            best_candidate = found_candidates[0]
            masterdata_path = best_candidate[0]
            file_size = best_candidate[1]
            priority = best_candidate[2]
            reason = best_candidate[3]
            
            if file_size > 0:
                logging.info(f"✅ SELECTED: {masterdata_path} ({file_size} bytes)")
                logging.info(f"✅ Reason: {reason} (Priority: {priority})")
                return masterdata_path
            else:
                logging.warning(f"⚠️ Best candidate has 0 bytes: {masterdata_path}")
        
        # FINAL FALLBACK: Largest CSV file
        logging.info("🔍 No masterdata found, looking for substantial CSV files...")
        
        try:
            fallback_candidates = []
            for root, dirs, files in os.walk(str(base_path)):
                for file in files:
                    if file.lower().endswith('.csv'):
                        file_path = os.path.join(root, file)
                        try:
                            file_size = os.path.getsize(file_path)
                            if file_size > 10000:  # At least 10KB
                                fallback_candidates.append((file_path, file_size))
                        except Exception:
                            continue
            
            if fallback_candidates:
                fallback_candidates.sort(key=lambda x: x[1], reverse=True)
                largest_csv = fallback_candidates[0][0]
                largest_size = fallback_candidates[0][1]
                
                logging.info(f"🎯 Using largest CSV as fallback: {largest_csv} ({largest_size} bytes)")
                return largest_csv
                
        except Exception as e:
            logging.error(f"Error in fallback search: {str(e)}")
        
        # COMPREHENSIVE DEBUG INFO
        logging.warning(f"❌ No masterdata file found in {base_path}")
        
        try:
            logging.info("🔍 DEBUGGING: Directory analysis...")
            all_files = []
            for root, dirs, files in os.walk(str(base_path)):
                try:
                    depth = len(Path(root).relative_to(base_path).parts) if Path(root) != base_path else 0
                    if depth <= max_depth:
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                size = os.path.getsize(file_path)
                                file_ext = Path(file).suffix.lower()
                                all_files.append(f"📄 {file} ({size} bytes) [{file_ext or 'no ext'}] in {root}")
                            except OSError:
                                all_files.append(f"📄 {file} (unknown size) in {root}")
                except Exception:
                    continue
            
            # Show master-related files first
            master_files = [f for f in all_files if 'master' in f.lower()]
            if master_files:
                logging.info(f"🎯 Files with 'master': {len(master_files)}")
                for mf in master_files[:10]:
                    logging.info(f"   {mf}")
            
            # Show some CSV files
            csv_files = [f for f in all_files if '.csv' in f]
            if csv_files:
                logging.info(f"📄 CSV files found: {len(csv_files)}")
                for cf in csv_files[:10]:
                    logging.info(f"   {cf}")
                    
        except Exception as e:
            logging.error(f"Debug analysis error: {str(e)}")
        
        return None
        
    except Exception as e:
        logging.error(f"❌ Error in find_masterdata_file: {str(e)}")
        return None
    
@app.route('/process-onedrive-pca', methods=['POST'])
@login_required
def process_onedrive_pca_route():
    """Process a file selected from OneDrive for PCA analysis"""
    # Get OneDrive file ID from form
    file_id = request.form.get('onedrive_file_id')
    if not file_id:
        flash("No OneDrive file selected", "error")
        return redirect(url_for('NewPca'))
    
    # Get access token
    access_token = get_master_onedrive_token()
    if not access_token:
        flash("OneDrive is not properly configured", "error")
        return redirect(url_for('NewPca'))
    
    try:
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details and download URL
        file_name = None
        file_content = None
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        # Download the file content
                        download_response = requests.get(download_url)
                        if download_response.status_code == 200:
                            file_content = download_response.content
                            break
                    else:
                        logging.warning("No download URL found in file info")
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        # If we couldn't get the file via the download URL, try direct content access
        if not file_content:
            for endpoint in [
                f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
                f"https://api.onedrive.com/v1.0/drive/items/{file_id}/content"
            ]:
                try:
                    response = requests.get(endpoint, headers=headers)
                    
                    if response.status_code == 200:
                        file_content = response.content
                        
                        # Try to get filename from Content-Disposition if we don't have it
                        if not file_name:
                            content_disposition = response.headers.get('Content-Disposition', '')
                            if 'filename=' in content_disposition:
                                file_name = content_disposition.split('filename=')[1].strip('"')
                            else:
                                file_name = f"file_{file_id}.csv"
                        break
                except Exception as e:
                    logging.error(f"Error with {endpoint}: {str(e)}")
                    continue
        
        if not file_content:
            flash("Could not download file from OneDrive", "error")
            return redirect(url_for('NewPca'))
        
        if not file_name:
            file_name = f"unknown_file_{file_id}.csv"
            
        # Process the file with the PCA function
        try:
            # Convert to a file-like object for immediate processing
            from io import BytesIO
            temp_file = BytesIO(file_content)
            
            # Read CSV directly from the BytesIO object
            df = pd.read_csv(temp_file)
            
            # Extract unique values for the form
            strings = df['bacteria'].unique().tolist() if 'bacteria' in df.columns else []
            unq_concs = df['concentration'].unique().tolist() if 'concentration' in df.columns else []
            unq_vols = df['volume'].unique().tolist() if 'volume' in df.columns else []
            unq_sli = df['slide'].unique().tolist() if 'slide' in df.columns else []
            unq_tri = df['trial'].unique().tolist() if 'trial' in df.columns else []
            
            # Store the DataFrame in session for processing
            # Note: For large files, you might want to store the file path instead
            session['pca_dataframe'] = df.to_json()
            session['pca_filename'] = file_name
            session.modified = True
            
            # Return the template with options
            return render_template('current_analysis_PCA.HTML', 
                                 item=strings, 
                                 item2=unq_concs, 
                                 item3=unq_vols, 
                                 item4=unq_sli, 
                                 item5=unq_tri,
                                 file_name=file_name)
                               
        except Exception as e:
            logging.error(f"Error processing OneDrive file for PCA: {str(e)}")
            flash(f"Error processing file: {str(e)}", "error")
            return redirect(url_for('NewPca'))
            
    except Exception as e:
        logging.error(f"Error accessing OneDrive file for PCA: {str(e)}")
        flash(f"Error accessing OneDrive file: {str(e)}", "error")
        return redirect(url_for('NewPca'))
    
@app.route('/api/save-pca-to-onedrive', methods=['POST'])
@login_required
def save_pca_to_onedrive():
    """API endpoint to save PCA visualization result to OneDrive"""
    try:
        # Get the image data from the request
        image_data = request.form.get('image_data')
        if not image_data:
            return jsonify({
                'success': False,
                'message': 'No image data provided'
            }), 400
        
        # Get filename with timestamp
        filename = request.form.get('filename', 'pca_visualization.png')
        if not filename.lower().endswith('.png'):
            filename += '.png'
        
        # Add timestamp to filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"pca_{timestamp}_{filename}"
        
        # Convert base64 image data to binary
        try:
            # Remove data URL prefix if present
            if ',' in image_data:
                image_data = image_data.split(',', 1)[1]
            
            file_content = base64.b64decode(image_data)
        except Exception as e:
            return jsonify({
                'success': False,
                'message': f'Invalid image data: {str(e)}'
            }), 400
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            return jsonify({
                'success': False,
                'message': 'OneDrive is not properly configured'
            }), 500
        
        # Create headers for API calls
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Get username for the folder
        username = current_user.username
        
        # Step 1: Find or create user folder
        user_folder_id = None
        
        # Look for user folder in root
        for endpoint in [
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://api.onedrive.com/v1.0/drive/root/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    root_items = response.json().get('value', [])
                    for item in root_items:
                        if item.get('name') == username and item.get('folder') is not None:
                            user_folder_id = item.get('id')
                            logging.info(f"Found user folder: {username}, id: {user_folder_id}")
                            break
                    
                    if user_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create user folder if it doesn't exist
        if not user_folder_id:
            logging.info(f"Creating user folder: {username}")
            create_folder_data = {
                'name': username,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                'https://graph.microsoft.com/v1.0/me/drive/root/children',
                'https://api.onedrive.com/v1.0/drive/root/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        user_folder_id = create_response.json().get('id')
                        logging.info(f"Created user folder: {username}, id: {user_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating user folder with {endpoint}: {str(e)}")
        
        if not user_folder_id:
            return jsonify({
                'success': False,
                'message': 'Could not find or create user folder in OneDrive'
            }), 500
        
        # Step 2: Find or create PCA results subfolder
        pca_folder_name = 'pca_results'
        pca_folder_id = None
        
        # Look for PCA folder in user folder
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
            f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    folder_items = response.json().get('value', [])
                    for item in folder_items:
                        if item.get('name') == pca_folder_name and item.get('folder') is not None:
                            pca_folder_id = item.get('id')
                            logging.info(f"Found PCA folder: {pca_folder_name}, id: {pca_folder_id}")
                            break
                    
                    if pca_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create PCA folder if it doesn't exist
        if not pca_folder_id:
            logging.info(f"Creating PCA folder: {pca_folder_name}")
            create_folder_data = {
                'name': pca_folder_name,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        pca_folder_id = create_response.json().get('id')
                        logging.info(f"Created PCA folder: {pca_folder_name}, id: {pca_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating PCA folder with {endpoint}: {str(e)}")
        
        if not pca_folder_id:
            return jsonify({
                'success': False,
                'message': f'Could not find or create {pca_folder_name} folder in OneDrive'
            }), 500
        
        # Step 3: Upload PCA visualization to PCA folder
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'image/png'
        }
        
        file_uploaded = False
        
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{pca_folder_id}:/{filename}:/content',
            f'https://api.onedrive.com/v1.0/drive/items/{pca_folder_id}:/{filename}:/content'
        ]:
            try:
                upload_response = requests.put(
                    endpoint,
                    headers=upload_headers,
                    data=file_content
                )
                
                if upload_response.status_code in [200, 201]:
                    logging.info(f"Successfully uploaded PCA visualization to OneDrive: {filename}")
                    file_uploaded = True
                    break
            except Exception as e:
                logging.error(f"Error uploading PCA visualization with {endpoint}: {str(e)}")
        
        if not file_uploaded:
            return jsonify({
                'success': False,
                'message': 'Failed to upload PCA visualization to OneDrive'
            }), 500
        
        # Return success response
        return jsonify({
            'success': True,
            'message': f'PCA visualization successfully saved to OneDrive',
            'filename': filename,
            'folder': pca_folder_name
        })
        
    except Exception as e:
        logging.error(f"Error saving PCA visualization to OneDrive: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        }), 500
    
def process_onedrive_file_for_ml():
    """Process a file selected from OneDrive for ML analysis - Enhanced version"""
    try:
        # Get OneDrive file ID from form
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            logging.error("No OneDrive file ID provided for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details first
        file_name = None
        file_size = 0
        download_url = None
        
        logging.info(f"🔍 Getting ML file info for OneDrive file ID: {file_id}")
        
        # Try to get file info and download URL
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    
                    logging.info(f"📁 ML File: {file_name}, Size: {file_size} bytes ({file_size/1024/1024:.1f} MB)")
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        logging.info(f"✅ Got download URL for ML processing")
                        break
                    else:
                        logging.warning("⚠️ No download URL found in file info")
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            logging.error("❌ Could not get file info or download URL for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # Download and process file (with large file support)
        try:
            # Create temporary file
            import tempfile
            temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
            temp_file_path = temp_file.name
            
            # Download file content
            if file_size > 50 * 1024 * 1024:  # 50MB threshold
                logging.info(f"📥 Large ML file detected ({file_size/1024/1024:.1f} MB), using streaming download")
                
                with requests.get(download_url, stream=True) as download_response:
                    download_response.raise_for_status()
                    
                    for chunk in download_response.iter_content(chunk_size=8192):
                        if chunk:
                            temp_file.write(chunk)
                            
                temp_file.close()
                is_large_file = True
            else:
                logging.info(f"📥 Standard download for ML file ({file_size/1024/1024:.1f} MB)")
                
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    temp_file.write(download_response.content)
                    temp_file.close()
                    is_large_file = False
                else:
                    temp_file.close()
                    os.unlink(temp_file_path)
                    logging.error(f"❌ Download failed: {download_response.status_code}")
                    return render_template('current_analysis_ML.HTML', 
                                         sample_placeholder="Error: Could not download file from OneDrive")
            
            # Store the file path in session for step 2
            session['csv_file_path'] = temp_file_path
            session['data_type'] = request.form.get("data_type", "raw_data")
            session['is_large_file'] = is_large_file
            session['file_size_mb'] = file_size / 1024 / 1024
            session.modified = True
            
            # Read and extract unique values
            if is_large_file and file_size > 100 * 1024 * 1024:
                logging.info(f"🔄 Processing very large ML file with chunking...")
                
                # Use chunked processing for very large files
                chunk_size_rows = 10000
                unique_bacteria = set()
                unique_concs = set()
                unique_vols = set()
                unique_slides = set()
                unique_trials = set()
                
                chunk_count = 0
                for chunk in pd.read_csv(temp_file_path, chunksize=chunk_size_rows):
                    chunk_count += 1
                    
                    # Extract unique values from this chunk
                    if 'bacteria' in chunk.columns:
                        unique_bacteria.update(chunk['bacteria'].dropna().unique())
                    if 'concentration' in chunk.columns:
                        unique_concs.update(chunk['concentration'].dropna().unique())
                    if 'volume' in chunk.columns:
                        unique_vols.update(chunk['volume'].dropna().unique())
                    if 'slide' in chunk.columns:
                        unique_slides.update(chunk['slide'].dropna().unique())
                    if 'trial' in chunk.columns:
                        unique_trials.update(chunk['trial'].dropna().unique())
                    
                    if chunk_count % 10 == 0:
                        logging.info(f"📊 ML: Processed {chunk_count} chunks, found {len(unique_bacteria)} bacteria types")
                
                # Convert sets to sorted lists
                strings = sorted(list(unique_bacteria))
                unq_concs = sorted(list(unique_concs))
                unq_vols = sorted(list(unique_vols))
                unq_sli = sorted(list(unique_slides))
                unq_tri = sorted(list(unique_trials))
                
                logging.info(f"✅ ML chunked processing complete: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                
            else:
                # Regular processing for smaller files
                logging.info(f"🔄 Processing ML file normally...")
                dfX = pd.read_csv(temp_file_path)
                logging.info(f"📊 Successfully read ML CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                
                # Extract unique values
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                
                logging.info(f"✅ ML: Extracted unique values: {len(strings)} bacteria types")
            
            logging.info(f"💾 Stored ML file path in session: {temp_file_path}")
            
            return render_template('current_analysis_ML.HTML', 
                                  item=strings, 
                                  item2=unq_concs, 
                                  item3=unq_vols, 
                                  item4=unq_sli, 
                                  item5=unq_tri,
                                  file_size_info=f"File: {file_size/1024/1024:.1f} MB" if is_large_file else None)
                                  
        except Exception as e:
            logging.error(f"❌ Error processing OneDrive ML file: {str(e)}")
            # Clean up temp file on error
            try:
                if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except:
                pass
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive ML file processing: {str(e)}")
        return render_template('current_analysis_ML.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")

def process_large_file_ml(file_path, bacts, conc, vol, slide, trails, chunk_size=10000):
    """Process large files for ML analysis using chunked reading"""
    try:
        logging.info(f"🔄 Processing large file for ML: {file_path}")
        
        # First pass: get data statistics and validate
        total_rows = 0
        matching_rows = 0
        
        # Read first chunk to understand data structure
        first_chunk = pd.read_csv(file_path, nrows=1000)
        logging.info(f"📊 ML Data structure - Columns: {list(first_chunk.columns)}")
        
        # Validate required columns for ML
        required_cols = ['bacteria']
        missing_cols = [col for col in required_cols if col not in first_chunk.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns for ML: {missing_cols}")
        
        # Second pass: filter and collect matching data
        filtered_chunks = []
        
        for chunk in pd.read_csv(file_path, chunksize=chunk_size):
            total_rows += len(chunk)
            
            # Apply filters
            original_chunk_size = len(chunk)
            
            # Filter by bacteria (required)
            if bacts and 'bacteria' in chunk.columns:
                chunk = chunk[chunk['bacteria'].isin(bacts)]
            
            # Filter by concentration
            if conc and 'concentration' in chunk.columns:
                try:
                    conc_numeric = float(conc)
                    chunk = chunk[chunk['concentration'] == conc_numeric]
                except (ValueError, TypeError):
                    chunk = chunk[chunk['concentration'].astype(str) == str(conc)]
            
            # Filter by volume
            if vol and 'volume' in chunk.columns:
                try:
                    vol_numeric = float(vol)
                    chunk = chunk[chunk['volume'] == vol_numeric]
                except (ValueError, TypeError):
                    chunk = chunk[chunk['volume'].astype(str) == str(vol)]
            
            # Filter by slide
            if slide and 'slide' in chunk.columns:
                slide_str = [str(s) for s in slide]
                chunk = chunk[chunk['slide'].astype(str).isin(slide_str)]
            
            # Filter by trial
            if trails and 'trial' in chunk.columns:
                trails_str = [str(t) for t in trails]
                chunk = chunk[chunk['trial'].astype(str).isin(trails_str)]
            
            if not chunk.empty:
                filtered_chunks.append(chunk)
                matching_rows += len(chunk)
        
        if not filtered_chunks:
            raise ValueError("No data matches the selected criteria")
        
        # Combine all filtered chunks
        logging.info(f"📊 Combining {len(filtered_chunks)} filtered chunks for ML...")
        combined_df = pd.concat(filtered_chunks, ignore_index=True)
        
        # Optimize memory usage
        combined_df = optimize_pca_dataframe_memory(combined_df)  # Reuse the function
        
        logging.info(f"✅ Large file ML processing complete:")
        logging.info(f"   - Total rows processed: {total_rows}")
        logging.info(f"   - Matching rows: {matching_rows}")
        logging.info(f"   - Final dataset size: {len(combined_df)} rows")
        
        return combined_df
        
    except Exception as e:
        logging.error(f"❌ Error in large file ML processing: {str(e)}")
        raise e
    
def optimize_pca_dataframe_memory(df):
    """Optimize DataFrame memory usage specifically for PCA analysis"""
    try:
        initial_memory = df.memory_usage(deep=True).sum()
        logging.info(f"Initial DataFrame memory usage: {initial_memory/1024/1024:.1f} MB")
        
        # For PCA, we need to preserve numeric columns for analysis
        # Only optimize categorical/string columns
        for col in df.select_dtypes(include=['object']):
            if df[col].nunique() / len(df) < 0.5:  # If less than 50% unique values, convert to category
                df[col] = df[col].astype('category')
        
        # Optimize integer columns
        for col in df.select_dtypes(include=['int64']):
            if df[col].min() >= 0:
                if df[col].max() < 255:
                    df[col] = df[col].astype('uint8')
                elif df[col].max() < 65535:
                    df[col] = df[col].astype('uint16')
                elif df[col].max() < 4294967295:
                    df[col] = df[col].astype('uint32')
            else:
                if df[col].min() > -128 and df[col].max() < 127:
                    df[col] = df[col].astype('int8')
                elif df[col].min() > -32768 and df[col].max() < 32767:
                    df[col] = df[col].astype('int16')
                elif df[col].min() > -2147483648 and df[col].max() < 2147483647:
                    df[col] = df[col].astype('int32')
        
        # Optimize float columns - be careful here for PCA as precision matters
        for col in df.select_dtypes(include=['float64']):
            # Only downcast if the range allows it without significant precision loss
            if df[col].min() >= np.finfo(np.float32).min and df[col].max() <= np.finfo(np.float32).max:
                # Check if conversion loses significant precision
                test_conversion = df[col].astype('float32')
                if np.allclose(df[col].values, test_conversion.values, rtol=1e-6):
                    df[col] = test_conversion
        
        final_memory = df.memory_usage(deep=True).sum()
        reduction = (initial_memory - final_memory) / initial_memory * 100
        
        logging.info(f"📊 PCA DataFrame memory optimization: {reduction:.1f}% reduction")
        logging.info(f"📊 Final memory usage: {final_memory/1024/1024:.1f} MB")
        
        return df
    except Exception as e:
        logging.warning(f"⚠️ PCA memory optimization failed: {str(e)}")
        return df

# Add this function to handle large file PCA processing
def process_large_file_pca(file_path, bacts, conc, vol, slide, trails, chunk_size=10000):
    """Process large files for PCA analysis using chunked reading"""
    try:
        logging.info(f"🔄 Processing large file for PCA: {file_path}")
        
        # First pass: get data statistics and validate
        total_rows = 0
        matching_rows = 0
        sample_data = None
        
        # Read first chunk to understand data structure
        first_chunk = pd.read_csv(file_path, nrows=1000)
        logging.info(f"📊 Data structure - Columns: {list(first_chunk.columns)}")
        
        # Validate required columns for PCA
        required_cols = ['bacteria']
        missing_cols = [col for col in required_cols if col not in first_chunk.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns for PCA: {missing_cols}")
        
        # Second pass: filter and collect matching data
        filtered_chunks = []
        
        for chunk in pd.read_csv(file_path, chunksize=chunk_size):
            total_rows += len(chunk)
            
            # Apply filters
            original_chunk_size = len(chunk)
            
            # Filter by bacteria (required)
            if bacts and 'bacteria' in chunk.columns:
                chunk = chunk[chunk['bacteria'].isin(bacts)]
            
            # Filter by concentration
            if conc and 'concentration' in chunk.columns:
                try:
                    conc_numeric = float(conc)
                    chunk = chunk[chunk['concentration'] == conc_numeric]
                except (ValueError, TypeError):
                    chunk = chunk[chunk['concentration'].astype(str) == str(conc)]
            
            # Filter by volume
            if vol and 'volume' in chunk.columns:
                try:
                    vol_numeric = float(vol)
                    chunk = chunk[chunk['volume'] == vol_numeric]
                except (ValueError, TypeError):
                    chunk = chunk[chunk['volume'].astype(str) == str(vol)]
            
            # Filter by slide
            if slide and 'slide' in chunk.columns:
                slide_str = [str(s) for s in slide]
                chunk = chunk[chunk['slide'].astype(str).isin(slide_str)]
            
            # Filter by trial
            if trails and 'trial' in chunk.columns:
                trails_str = [str(t) for t in trails]
                chunk = chunk[chunk['trial'].astype(str).isin(trails_str)]
            
            if not chunk.empty:
                filtered_chunks.append(chunk)
                matching_rows += len(chunk)
                
                # Keep a sample for validation
                if sample_data is None:
                    sample_data = chunk.head(100).copy()
        
        if not filtered_chunks:
            raise ValueError("No data matches the selected criteria")
        
        # Combine all filtered chunks
        logging.info(f"📊 Combining {len(filtered_chunks)} filtered chunks...")
        combined_df = pd.concat(filtered_chunks, ignore_index=True)
        
        # Optimize memory usage
        combined_df = optimize_pca_dataframe_memory(combined_df)
        
        logging.info(f"✅ Large file PCA processing complete:")
        logging.info(f"   - Total rows processed: {total_rows}")
        logging.info(f"   - Matching rows: {matching_rows}")
        logging.info(f"   - Final dataset size: {len(combined_df)} rows")
        
        return combined_df
        
    except Exception as e:
        logging.error(f"❌ Error in large file PCA processing: {str(e)}")
        raise e


@app.context_processor
def inject_session_data():
    """Make session data available to all templates"""
    # Extract only the primitive data from session that templates need
    session_data = {}
    
    # Extract OneDrive user info if available
    if 'onedrive_user' in session:
        session_data['onedrive_user'] = session.get('onedrive_user')
    else:
        session_data['onedrive_user'] = None
    
    # Add a flag for whether the user is connected to OneDrive
    session_data['is_connected_to_onedrive'] = 'onedrive_token' in session
    
    # Add any other session data needed in templates
    return session_data

# Register blueprints
app.register_blueprint(auth)


# === Utility Functions ===

def build_msal_app(cache=None):
    """Build the MSAL app for OneDrive authentication"""
    return ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET, token_cache=cache)


def get_user_file_path(filename, user_id=None):
    """Get path for user-specific file storage"""
    if user_id is None:
        user_id = current_user.id if current_user.is_authenticated else 'anonymous'
    
    user_folder = os.path.join(app.config['UPLOAD_FOLDER'], str(user_id))
    os.makedirs(user_folder, exist_ok=True)
    return os.path.join(user_folder, filename)

def build_msal_app(cache=None):
    """Build the MSAL app for OneDrive authentication"""
    return ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET, token_cache=cache)

def get_plot_path(plot_type, filename, user_id=None):
    """Get path for user-specific plot storage"""
    if user_id is None:
        user_id = current_user.id if current_user.is_authenticated else 'anonymous'
    
    plot_folder = os.path.join('Plots', str(user_id), plot_type)
    os.makedirs(plot_folder, exist_ok=True)
    return os.path.join(plot_folder, filename)

def save_plot_image(img_base64, plot_path):
    """Save a base64 encoded plot image to the file system"""
    try:
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(plot_path), exist_ok=True)
        
        # Check if the image data includes the base64 prefix
        if ',' in img_base64:
            # Extract only the base64 data part
            img_base64 = img_base64.split(',', 1)[1]
            
        # Convert base64 to image data and save
        img_data = base64.b64decode(img_base64)
        with open(plot_path, 'wb') as f:
            f.write(img_data)
        return True
    except Exception as e:
        logging.error(f"Error saving plot image: {str(e)}")
        return False

def add_file_record(file_path, original_filename, file_type=None, file_size=None):
    """Add a file record to the database"""
    if not current_user.is_authenticated:
        return None
        
    # Generate a unique stored filename
    stored_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
    
    # Determine file type if not provided
    if not file_type:
        if original_filename.lower().endswith('.csv'):
            file_type = 'text/csv'
        elif original_filename.lower().endswith('.txt'):
            file_type = 'text/plain'
        else:
            file_type = 'application/octet-stream'
    
    # Get file size if not provided
    if not file_size and os.path.exists(file_path):
        file_size = os.path.getsize(file_path)
    
    # Create file record using the UploadedFile.create method
    try:
        file_record = UploadedFile.create(
            user_id=current_user.id,
            original_filename=original_filename,
            stored_filename=stored_filename,
            file_path=file_path,
            file_type=file_type or 'unknown',
            file_size=file_size or os.path.getsize(file_path) if os.path.exists(file_path) else 0
        )
        logging.info(f"Created file record ID {file_record.id} for user {current_user.id}")
        return file_record
    except Exception as e:
        logging.error(f"Error creating file record: {str(e)}")
        return None

def user_can_access_file(file_id):
    """Check if current user can access the specified file"""
    if not current_user.is_authenticated:
        return False
    
    try:
        # Get the file record
        file_record = UploadedFile.get(int(file_id))
        if not file_record:
            return False
        
        # User has access if they own the file or are an admin
        return file_record.user_id == current_user.id or current_user.is_admin()
    except Exception as e:
        logging.error(f"Error checking file access: {str(e)}")
        return False

def get_master_onedrive_token():
    """
    Get the OneDrive access token, refreshing if necessary.
    Returns a valid access token or None if not configured.
    """
    # First check if token is in the database
    access_token = SystemConfig.get_value('onedrive_token')
    refresh_token = SystemConfig.get_value('onedrive_refresh_token')
    token_expires_str = SystemConfig.get_value('onedrive_token_expires')
    
    if not access_token or not refresh_token:
        logging.warning("No master OneDrive token found in system configuration")
        return None
    
    # Check if token is expired or about to expire (within 5 minutes)
    now = int(datetime.now().timestamp())
    
    try:
        # Safely convert to int, with fallback
        token_expires = int(token_expires_str) if token_expires_str and token_expires_str.isdigit() else 0
        
        if token_expires and token_expires - now < 300:
            logging.info("OneDrive token expired or about to expire, refreshing...")
            # Refresh the token
            token_data = {
                'client_id': CLIENT_ID,
                'client_secret': CLIENT_SECRET,
                'refresh_token': refresh_token,
                'grant_type': 'refresh_token',
                'redirect_uri': REDIRECT_URI
            }
            
            try:
                # Try both token endpoints
                for token_url in [TOKEN_URL, TOKEN_URL_FALLBACK]:
                    try:
                        logging.info(f"Attempting to refresh token using: {token_url}")
                        token_response = requests.post(token_url, data=token_data)
                        
                        if token_response.status_code == 200:
                            token_info = token_response.json()
                            access_token = token_info.get("access_token")
                            new_refresh_token = token_info.get("refresh_token")
                            
                            if new_refresh_token:  # Only update if we got a new one
                                refresh_token = new_refresh_token
                                
                            token_expires = int(now + token_info.get("expires_in", 3600))
                            
                            # Update database
                            SystemConfig.set_value('onedrive_token', access_token)
                            SystemConfig.set_value('onedrive_refresh_token', refresh_token)
                            SystemConfig.set_value('onedrive_token_expires', str(token_expires))
                            
                            logging.info("OneDrive token refreshed successfully")
                            break  # Exit the loop if successful
                        else:
                            logging.warning(f"Failed to refresh token with {token_url}: {token_response.status_code}")
                            if token_response.status_code != 400:  # Don't log potentially sensitive error data
                                logging.warning(f"Response: {token_response.text[:100]}...")
                    except Exception as e:
                        logging.error(f"Error with {token_url}: {str(e)}")
                
                # If we still don't have a valid token after trying both endpoints
                if token_expires and token_expires - now < 0:
                    logging.error("Failed to refresh OneDrive token with all endpoints")
                    return None
                    
            except Exception as e:
                logging.error(f"Error refreshing token: {str(e)}")
                return None
    except (ValueError, TypeError) as e:
        logging.error(f"Error processing token expiry time: {str(e)}")
        # Continue with existing token if we can't parse the expiry time
    
    return access_token

# === Main Application Routes ===

# Root route - redirect to dashboard if logged in, or homepage if not
@app.route('/')
def index():
    return render_template('breath1.html')

# Then add a separate home route that redirects to index:
@app.route('/home')
def home():
    return redirect(url_for('index'))

# Dashboard route - Central hub for the application
@app.route('/dashboard')
@login_required
def dashboard():
    return redirect(url_for('home'))

@app.route('/dashboard2')
def dashboard2():
    return render_template('dashboard.html')

# === Bacteria Analysis Routes ===

@app.route('/page2breathe')
def page2breathe():
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('page2breathe.html', sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_page2')
@login_required
def redirect_to_page2():
    """Redirect to the bacteria analysis page"""
    return redirect(url_for('page2breathe'))

@app.route('/page2newAnalysis')
def page2newAnalysis():
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('new_analysis_landing.html', sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_new_analysis')
def redirect_to_new_analysis():
    return redirect(url_for('page2newAnalysis'))

# === Data Management Routes ===

@app.route('/DMwebpage')
def DMwebpage():
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('DMwebpage.html', sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_DMwebpage')
def redirect_to_DMwebpage():
    return redirect(url_for('DMwebpage'))

# === New Analysis Routes ===

# Fix for the Method Not Allowed error in app.py
# Add this to your existing app.py file

@app.route('/NewrawVisual', methods=['GET', 'POST'])
def NewrawVisual():
    if request.method == 'POST':
        # This is the same code as in your "NewVisualization" route, but specifically for the first step
        try:
            # Check file source
            file_source = request.form.get('fileSource', 'local')
            logging.info(f"🔍 Raw data visualization with file source: {file_source}")
            
            if file_source == 'cloud' or file_source == 'onedrive':
                # Handle OneDrive file processing
                logging.info(f"☁️ Processing OneDrive file...")
                return process_onedrive_file_for_visualization()
            else:
                # Handle local file upload
                if 'csv_file' not in request.files:
                    logging.warning("⚠️ No file part in request")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: No file uploaded")
                    
                folder_path = request.files['csv_file']
                if folder_path.filename == '':
                    logging.warning("⚠️ No selected file")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: No file selected")
                
                # Check file extension
                if not folder_path.filename.lower().endswith('.csv'):
                    logging.warning(f"⚠️ Invalid file type: {folder_path.filename}")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: Only CSV files are allowed")
                
                # Get data type selection
                data_type = request.form.get("data_type")
                if not data_type:
                    logging.warning("⚠️ No data type selected")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: Please select a data type")
                
                # Save the uploaded file
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                safe_filename = f"{timestamp}_{secure_filename(folder_path.filename)}"
                csv_filename = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
                folder_path.save(csv_filename)
                logging.info(f"💾 File saved to: {csv_filename}")
                
                # Read CSV file to extract unique values
                try:
                    dfX = pd.read_csv(csv_filename)
                    logging.info(f"📊 Successfully read CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                except Exception as e:
                    logging.error(f"❌ Error reading CSV: {str(e)}")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder=f"Error reading CSV file: {str(e)}")
                
                # Store the file path and data type in session
                session['csv_file_path'] = csv_filename
                session['data_type'] = data_type
                session['is_large_file'] = False
                session.modified = True
                
                # Extract unique values from different columns
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                
                logging.info(f"✅ Extracted unique values - bacteria: {len(strings)}, concs: {len(unq_concs)}")
                
                return render_template('current_analysis_rawDatavisual.HTML', 
                                      item=strings, 
                                      item2=unq_concs, 
                                      item3=unq_vols, 
                                      item4=unq_sli, 
                                      item5=unq_tri)
        except Exception as e:
            logging.error(f"❌ Error processing uploaded file: {str(e)}")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
    else:
        # GET request - original functionality
        if current_user.is_authenticated:
            # Get user's recent visualization files if logged in
            user_files = UploadedFile.get_user_files(current_user.id)
            visual_files = [f for f in user_files if f.processing_type == 'raw_data_visualization'][:5]
        else:
            user_files = []
            visual_files = []
            
        prediction_text_placeholder = "Please wait for processor response"
        return render_template('current_analysis_rawDatavisual.HTML', 
                             sample_placeholder=prediction_text_placeholder,
                             user_files=visual_files)

@app.route('/redirect_to_NewrawVisual')
def redirect_to_NewrawVisual():
    return redirect(url_for('NewrawVisual'))

@app.route('/NewPca')
def NewPca():
    if current_user.is_authenticated:
        # Get user's recent PCA files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        pca_files = [f for f in user_files if f.processing_type == 'pca_analysis'][:5]
    else:
        user_files = []
        pca_files = []
        
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('current_analysis_PCA.HTML', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=pca_files)

@app.route('/redirect_to_NewPca')
def redirect_to_NewPca():
    return redirect(url_for('NewPca'))


@app.route('/admin/setup-onedrive')
@admin_required
def admin_setup_onedrive():
    """Set up master OneDrive account"""
    # Generate a state for CSRF protection
    state = str(uuid.uuid4())
    session['onedrive_state'] = state
    
    # For personal Microsoft accounts, use the Live.com endpoint with simple scopes
    redirect_uri = url_for('admin_onedrive_callback', _external=True)
    scope = "Files.ReadWrite.All offline_access"
    
    auth_url = (
        "https://login.live.com/oauth20_authorize.srf"
        f"?client_id={CLIENT_ID}"
        f"&scope={urllib.parse.quote(scope)}"
        f"&redirect_uri={urllib.parse.quote(redirect_uri)}"
        "&response_type=code"
        f"&state={state}"
    )
    
    logging.info(f"Redirecting admin to OneDrive auth URL: {auth_url}")
    return redirect(auth_url)

@app.route('/admin/onedrive-callback')
@admin_required
def admin_onedrive_callback():
    """Handle the OAuth callback from Microsoft"""
    # Verify state for security
    if request.args.get('state') != session.get('onedrive_state'):
        flash("State verification failed", "error")
        return redirect(url_for('admin_dashboard'))
    
    # Check for error
    if 'error' in request.args:
        flash(f"OAuth error: {request.args.get('error_description', request.args.get('error'))}", "error")
        return redirect(url_for('admin_dashboard'))
    
    # Get the authorization code
    code = request.args.get('code')
    if not code:
        flash("No authorization code received", "error")
        return redirect(url_for('admin_dashboard'))
    
    # Exchange code for token - try Live.com endpoint for personal Microsoft accounts
    redirect_uri = url_for('admin_onedrive_callback', _external=True)
    token_url = "https://login.live.com/oauth20_token.srf"
    
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'code': code,
        'redirect_uri': redirect_uri,
        'grant_type': 'authorization_code'
    }
    
    try:
        response = requests.post(token_url, data=data)
        logging.info(f"Token response status: {response.status_code}")
        
        if response.status_code == 200:
            token_data = response.json()
            
            # Save the tokens to the database
            SystemConfig.set_value('onedrive_token', token_data.get('access_token'))
            SystemConfig.set_value('onedrive_refresh_token', token_data.get('refresh_token'))
            SystemConfig.set_value('onedrive_token_expires', 
                                 str(int(datetime.now().timestamp() + token_data.get('expires_in', 3600))))
            
            # Get user info
            headers = {'Authorization': f"Bearer {token_data.get('access_token')}"}
            
            # Try both APIs for user info
            user_info = None
            for endpoint in ['https://graph.microsoft.com/v1.0/me', 'https://apis.live.net/v5.0/me']:
                try:
                    user_response = requests.get(endpoint, headers=headers)
                    if user_response.status_code == 200:
                        user_info = user_response.json()
                        break
                except Exception as e:
                    logging.error(f"Error with {endpoint}: {str(e)}")
            
            if user_info:
                # Extract user info based on which API succeeded
                if 'displayName' in user_info:  # Graph API
                    SystemConfig.set_value('onedrive_user_name', user_info.get('displayName', 'Unknown'))
                    SystemConfig.set_value('onedrive_user_email', user_info.get('userPrincipalName', 'Unknown'))
                else:  # Live.com API
                    SystemConfig.set_value('onedrive_user_name', user_info.get('name', 'Unknown'))
                    email = user_info.get('emails', {}).get('account', 'Unknown')
                    SystemConfig.set_value('onedrive_user_email', email)
            
            flash("Successfully connected to OneDrive", "success")
        else:
            error_text = "Unknown error"
            try:
                error_data = response.json()
                error_text = error_data.get('error_description', response.text[:200])
            except:
                error_text = response.text[:200]
                
            flash(f"Error getting token: {error_text}", "error")
            logging.error(f"Token error: {error_text}")
            
    except Exception as e:
        flash(f"Error: {str(e)}", "error")
        logging.error(f"Exception in callback: {str(e)}")
    
    return redirect(url_for('admin_dashboard'))

@app.route('/NewML')
def NewML():
    if current_user.is_authenticated:
        # Get user's recent ML files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        ml_files = [f for f in user_files if f.processing_type == 'machine_learning_analysis'][:5]
    else:
        user_files = []
        ml_files = []
        
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('current_analysis_ML.HTML', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=ml_files)

@app.route('/onedrive-callback')
@login_required
def onedrive_callback():
    """Handle OAuth callback from Live.com"""
    # Verify state
    received_state = request.args.get('state')
    expected_state = session.get("onedrive_auth_state")
    
    if received_state != expected_state:
        flash("State validation failed", "error")
        return redirect(url_for('dashboard'))
    
    # Clear state from session
    if 'onedrive_auth_state' in session:
        del session['onedrive_auth_state']
    
    # Check for code
    if 'code' not in request.args:
        flash("No authorization code received", "error")
        return redirect(url_for('dashboard'))
    
    # Exchange code for token using Live.com token endpoint
    code = request.args.get('code')
    
    # Prepare token request data
    token_url = "https://login.live.com/oauth20_token.srf"
    token_data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'code': code,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code'
    }
    
    try:
        # Make the token request
        token_response = requests.post(token_url, data=token_data)
        
        logging.info(f"Token response status: {token_response.status_code}")
        
        if token_response.status_code == 200:
            token_info = token_response.json()
            
            # Store token in session
            session["onedrive_token"] = token_info.get("access_token")
            session["onedrive_refresh_token"] = token_info.get("refresh_token")
            session["onedrive_token_expires"] = int(datetime.now().timestamp()) + token_info.get("expires_in", 3600)
            
            # Get user info - with Live.com we might need to use a different endpoint
            user_info = {"display_name": "OneDrive User", "email": "Connected Account"}
            
            try:
                # Try the Live.com user endpoint
                user_endpoint = "https://apis.live.net/v5.0/me"
                user_response = requests.get(
                    user_endpoint,
                    headers={'Authorization': f'Bearer {session["onedrive_token"]}'}
                )
                
                if user_response.status_code == 200:
                    user_data = user_response.json()
                    user_info = {
                        "display_name": user_data.get("name", "OneDrive User"),
                        "email": user_data.get("emails", {}).get("account", "Connected Account")
                    }
                else:
                    # Fallback to Microsoft Graph
                    graph_user_endpoint = "https://graph.microsoft.com/v1.0/me"
                    graph_response = requests.get(
                        graph_user_endpoint,
                        headers={'Authorization': f'Bearer {session["onedrive_token"]}'}
                    )
                    
                    if graph_response.status_code == 200:
                        graph_data = graph_response.json()
                        user_info = {
                            "display_name": graph_data.get("displayName", "OneDrive User"),
                            "email": graph_data.get("userPrincipalName", "Connected Account")
                        }
            except Exception as e:
                logging.error(f"Error getting user info: {str(e)}")
            
            # Store user info
            session["onedrive_user"] = user_info
            session.modified = True
            
            flash("Successfully connected to OneDrive", "success")
            return redirect(url_for('dashboard'))
        else:
            error_text = "Unknown error"
            try:
                error_data = token_response.json()
                error_text = error_data.get('error_description', 'Unknown error')
            except:
                error_text = token_response.text[:200]
            
            logging.error(f"Token error: {error_text}")
            flash(f"Error connecting to OneDrive: {error_text}", "error")
            return redirect(url_for('dashboard'))
    except Exception as e:
        logging.error(f"Exception in callback: {str(e)}")
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for('dashboard'))


@app.route('/redirect_to_NewML')
def redirect_to_NewML():
    return redirect(url_for('NewML'))

@app.route('/login-onedrive')
@login_required
def login_onedrive():
    """Direct OAuth flow for personal Microsoft accounts using Live.com"""
    # Generate state for CSRF protection
    state = str(uuid.uuid4())
    session["onedrive_auth_state"] = state
    session.modified = True
    
    # For personal Microsoft accounts, use the Live.com endpoint
    # with simple, basic scopes
    auth_url = (
        "https://login.live.com/oauth20_authorize.srf"
        f"?client_id={CLIENT_ID}"
        "&scope=onedrive.readonly"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
        "&response_type=code"
        f"&state={state}"
    )
    
    logging.info(f"Using Live.com OAuth endpoint: {auth_url}")
    return redirect(auth_url)
# === Machine Learning Routes ===

@app.route('/ML_webpage')
def ML_webpage():
    if current_user.is_authenticated:
        # Get user's recent ML files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        ml_files = [f for f in user_files if 'ml' in f.processing_type][:5]
    else:
        user_files = []
        ml_files = []
        
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('ML_webpage.html', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=ml_files)

@app.route('/redirect_to_ML_webpage')
def redirect_to_ML_webpage():
    return redirect(url_for('ML_webpage'))

def is_folder_within_user_scope(folder_id, user_folder_id, headers, username):
    """
    Check if a folder is within the user's permitted scope
    This prevents users from accessing folders outside their designated area
    """
    try:
        # If the requested folder is the user folder itself, it's allowed
        if folder_id == user_folder_id:
            return True
        
        # Trace the folder hierarchy up to see if it leads to the user folder
        current_id = folder_id
        max_depth = 20  # Prevent infinite loops
        depth = 0
        
        while current_id and depth < max_depth:
            # Get folder info
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{current_id}',
                f'https://api.onedrive.com/v1.0/drive/items/{current_id}'
            ]:
                try:
                    response = requests.get(endpoint, headers=headers)
                    
                    if response.status_code == 200:
                        item_info = response.json()
                        
                        # Get parent reference
                        parent_ref = item_info.get('parentReference', {})
                        parent_id = parent_ref.get('id')
                        
                        # If parent is the user folder, we're good
                        if parent_id == user_folder_id:
                            return True
                        
                        # If parent is root, check if current folder is the user folder
                        if parent_ref.get('path') == '/drive/root':
                            # We've reached root level
                            folder_name = item_info.get('name', '')
                            return folder_name == username
                        
                        # Move up to parent
                        current_id = parent_id
                        break
                except Exception as e:
                    logging.error(f"Error checking folder scope with {endpoint}: {str(e)}")
                    continue
            else:
                # If we couldn't get info from any endpoint, deny access
                return False
            
            depth += 1
        
        # If we've gone too deep or couldn't trace back to user folder, deny access
        return False
        
    except Exception as e:
        logging.error(f"Error checking folder scope: {str(e)}")
        return False

# === Clustering Routes ===

@app.route('/clusterwebpage')
def clusterwebpage():
    if current_user.is_authenticated:
        # Get user's recent clustering files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        cluster_files = [f for f in user_files if 'cluster' in f.processing_type][:5]
    else:
        user_files = []
        cluster_files = []
        
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('clusterwebpage.html', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=cluster_files)

@app.route('/redirect_to_clusterwebpage')
def redirect_to_clusterwebpage():
    return redirect(url_for('clusterwebpage'))

# === Visualization Routes ===

@app.route('/visual')
def visual():
    if current_user.is_authenticated:
        # Get user's recent visualization files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        visual_files = [f for f in user_files if 'visual' in f.processing_type][:5]
    else:
        user_files = []
        visual_files = []
        
    prediction_text_placeholder = "Please select data type and upload a CSV file to visualize."
    return render_template('visual.html', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=visual_files)

@app.route('/redirect_to_visual')
def redirect_to_visual():
    return redirect(url_for('visual'))

# === Breathing Analysis Routes ===

@app.route('/breathepage')
def breathepage():
    if current_user.is_authenticated:
        # Get user's breath analysis files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        breathing_types = ['raw_data', 'svc', 'fvc', 'mvv', 'sm']
        breath_files = [f for f in user_files if f.processing_type in breathing_types][:5]
    else:
        breath_files = []
        
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('BreatheAnalysisHome.html', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=breath_files)

@app.route('/redirect_to_breathepage')
@login_required
def redirect_to_breathepage():
    return redirect(url_for('breathepage'))

@app.route('/api/onedrive-status')
@login_required
def onedrive_status():
    """API endpoint to check if OneDrive is connected"""
    # Check if master OneDrive token exists and is valid
    access_token = get_master_onedrive_token()
    
    return jsonify({
        'success': True,
        'connected': access_token is not None
    })

def build_folder_path(folder_id, user_folder_id, headers, username):
    """
    Build the folder path from current folder back to user root folder
    """
    try:
        path = []
        current_id = folder_id
        max_depth = 20  # Prevent infinite loops
        depth = 0
        
        while current_id and current_id != user_folder_id and depth < max_depth:
            # Get folder info
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{current_id}',
                f'https://api.onedrive.com/v1.0/drive/items/{current_id}'
            ]:
                try:
                    response = requests.get(endpoint, headers=headers)
                    
                    if response.status_code == 200:
                        item_info = response.json()
                        
                        # Add current folder to path
                        path.insert(0, {
                            'id': current_id,
                            'name': item_info.get('name', 'Unknown')
                        })
                        
                        # Get parent reference
                        parent_ref = item_info.get('parentReference', {})
                        parent_id = parent_ref.get('id')
                        
                        # If parent is the user folder, add it and we're done
                        if parent_id == user_folder_id:
                            path.insert(0, {
                                'id': user_folder_id,
                                'name': username
                            })
                            return path
                        
                        # Move up to parent
                        current_id = parent_id
                        break
                except Exception as e:
                    print(f"Error building folder path with {endpoint}: {str(e)}")
                    continue
            else:
                # If we couldn't get info from any endpoint, return what we have
                break
            
            depth += 1
        
        # If we're here and current_id is user_folder_id, add it
        if current_id == user_folder_id:
            path.insert(0, {
                'id': user_folder_id,
                'name': username
            })
        
        return path
        
    except Exception as e:
        print(f"Error building folder path: {str(e)}")
        return []


@app.route('/api/onedrive-files')
def get_onedrive_files():
    """API endpoint to get REAL files and folders from user's OneDrive folder with navigation support"""
    # Get the access token
    access_token = get_master_onedrive_token()
    if not access_token:
        return jsonify({
            'success': False,
            'message': 'OneDrive is not properly configured'
        })
    
    # Get requested folder ID from query parameters
    requested_folder_id = request.args.get('folder_id')
    
    # Use 'shanwaz' as the username based on your OneDrive structure
    username = "shanwaz"
    files = []
    user_folder_id = None
    current_folder_id = None
    current_path = []
    
    try:
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Step 1: Find user's root folder ('shanwaz') in OneDrive
        print(f"Looking for user folder: {username}")
        
        # Try both Graph API endpoints
        for root_endpoint in [
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://api.onedrive.com/v1.0/drive/root/children'
        ]:
            try:
                response = requests.get(root_endpoint, headers=headers)
                
                if response.status_code == 200:
                    # Look for user's folder in root
                    root_items = response.json().get('value', [])
                    print(f"Found {len(root_items)} items in root")
                    
                    for item in root_items:
                        print(f"Root item: {item.get('name')} (folder: {item.get('folder') is not None})")
                        if item.get('name') == username and item.get('folder'):
                            user_folder_id = item.get('id')
                            print(f"Found user folder: {username}, id: {user_folder_id}")
                            break
                    
                    if user_folder_id:
                        break  # Exit the loop if we found the user folder
                        
            except Exception as e:
                print(f"Error with {root_endpoint}: {str(e)}")
                continue
        
        if not user_folder_id:
            return jsonify({
                'success': False,
                'message': f'User folder "{username}" not found in OneDrive root'
            })
        
        # Step 2: Determine which folder to list
        if requested_folder_id:
            # User requested a specific folder - validate it's within user's scope
            current_folder_id = requested_folder_id
            # Build path for this folder
            current_path = build_folder_path(requested_folder_id, user_folder_id, headers, username)
        else:
            # List user's root folder ('shanwaz')
            current_folder_id = user_folder_id
            current_path = [{'id': user_folder_id, 'name': username}]
        
        # Step 3: Get contents of the current folder
        if current_folder_id:
            print(f"Getting contents of folder: {current_folder_id}")
            
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{current_folder_id}/children',
                f'https://api.onedrive.com/v1.0/drive/items/{current_folder_id}/children'
            ]:
                try:
                    files_response = requests.get(endpoint, headers=headers)
                    print(f"Files response status: {files_response.status_code}")
                    
                    if files_response.status_code == 200:
                        all_items = files_response.json().get('value', [])
                        print(f"Found {len(all_items)} items in folder")
                        
                        # Log what we found for debugging
                        for item in all_items:
                            item_type = "folder" if item.get('folder') else "file"
                            print(f"  - {item.get('name')} ({item_type})")
                        
                        # Return all items (both files and folders)
                        files = all_items
                        break
                except Exception as e:
                    print(f"Error with {endpoint}: {str(e)}")
                    continue
        
        if not files:
            print("No files found or error accessing folder")
        
        return jsonify({
            'success': True,
            'files': files,
            'currentPath': current_path,
            'folderPath': '/'.join([folder['name'] for folder in current_path]),
            'userFolderId': user_folder_id,
            'currentFolderId': current_folder_id
        })
        
    except Exception as e:
        print(f"Error accessing OneDrive: {str(e)}")
        return jsonify({
            'success': False,
            'message': f"Error connecting to OneDrive: {str(e)}"
        })


@app.route('/process-onedrive-metrics', methods=['POST'])
@login_required
def process_onedrive_metrics():
    """Process a file selected from OneDrive for breathing metrics analysis"""
    # Get OneDrive file ID from form
    file_id = request.form.get('onedrive_file_id')
    if not file_id:
        flash("No OneDrive file selected", "error")
        return redirect(url_for('sm'))
    
    # Get analysis type (should be 'sm' for standard metrics)
    analysis_type = request.form.get('radioOption')
    if analysis_type != 'sm':
        flash("Invalid analysis type", "error")
        return redirect(url_for('sm'))
    
    # Get access token
    access_token = get_master_onedrive_token()
    if not access_token:
        flash("OneDrive is not properly configured", "error")
        return redirect(url_for('sm'))
    
    try:
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details and download URL
        file_name = None
        file_content = None
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        # Download the file content
                        download_response = requests.get(download_url)
                        if download_response.status_code == 200:
                            file_content = download_response.content
                            break
                    else:
                        logging.warning("No download URL found in file info")
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        # If we couldn't get the file via the download URL, try direct content access
        if not file_content:
            for endpoint in [
                f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
                f"https://api.onedrive.com/v1.0/drive/items/{file_id}/content"
            ]:
                try:
                    response = requests.get(endpoint, headers=headers)
                    
                    if response.status_code == 200:
                        file_content = response.content
                        
                        # Try to get filename from Content-Disposition if we don't have it
                        if not file_name:
                            content_disposition = response.headers.get('Content-Disposition', '')
                            if 'filename=' in content_disposition:
                                file_name = content_disposition.split('filename=')[1].strip('"')
                            else:
                                file_name = f"file_{file_id}.csv"
                        break
                except Exception as e:
                    logging.error(f"Error with {endpoint}: {str(e)}")
                    continue
        
        if not file_content:
            flash("Could not download file from OneDrive", "error")
            return redirect(url_for('sm'))
        
        if not file_name:
            file_name = f"unknown_file_{file_id}.csv"
            
        # Process the file with the sm_calc function
        try:
            # Convert to a file-like object for immediate processing
            from io import BytesIO
            temp_file = BytesIO(file_content)
            
            # Read CSV directly from the BytesIO object
            df = pd.read_csv(temp_file)
            
            # Process the data
            resp_rate, LungCapacity, Quality = sm_calc(df)
            
            # Generate a unique stored filename - IMPORTANT: This must be assigned
            stored_filename = f"{uuid.uuid4().hex}_{secure_filename(file_name)}"
            
            # Save file to disk in user's folder
            user_folder = os.path.join(app.config['UPLOAD_FOLDER'], str(current_user.id))
            os.makedirs(user_folder, exist_ok=True)
            file_path = os.path.join(user_folder, stored_filename)
            
            with open(file_path, 'wb') as f:
                f.write(file_content)
            
            # Get file size
            file_size = len(file_content)
            
            # Create file record in database
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    # Explicitly include the stored_filename in the SQL query
                    cur.execute("""
                        INSERT INTO uploaded_files 
                        (user_id, original_filename, stored_filename, file_path, file_type, file_size, source)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        RETURNING id
                    """, (current_user.id, file_name, stored_filename, file_path, 'text/csv', file_size, 'onedrive'))
                    file_id = cur.fetchone()['id']
                conn.commit()
                
                # Update file record with processing information
                with conn.cursor() as cur:
                    cur.execute("""
                        UPDATE uploaded_files 
                        SET processed = TRUE, processed_at = CURRENT_TIMESTAMP, processing_type = %s
                        WHERE id = %s
                    """, ('sm', file_id))
                conn.commit()
            except Exception as e:
                conn.rollback()
                logging.error(f"Error creating file record: {str(e)}")
                flash(f"Error creating file record: {str(e)}", "error")
                return redirect(url_for('sm'))
            finally:
                conn.close()
            
            # Return the template with results
            flash("Standard metrics calculated successfully!", "success")
            return render_template('Standard_Metrics_Breathe.HTML', 
                               resp_rate=resp_rate,
                               LungCapacity=LungCapacity,
                               Quality=Quality,
                               sample_placeholder="Analysis completed successfully!")
                               
        except Exception as e:
            logging.error(f"Error processing OneDrive file: {str(e)}")
            flash(f"Error processing file: {str(e)}", "error")
            return redirect(url_for('sm'))
            
    except Exception as e:
        logging.error(f"Error accessing OneDrive file: {str(e)}")
        flash(f"Error accessing OneDrive file: {str(e)}", "error")
        return redirect(url_for('sm'))

@app.route('/api/save-metrics-to-onedrive', methods=['POST'])
@login_required
def save_metrics_to_onedrive():
    """API endpoint to save breathing metrics as a text file to OneDrive"""
    try:
        # Get metrics data from form
        metrics_data = request.form.get('metrics_data')
        if not metrics_data:
            return jsonify({
                'success': False,
                'message': 'No metrics data provided'
            }), 400
        
        # Get filename with .txt extension
        filename = request.form.get('filename', 'breathing_metrics.txt')
        if not filename.lower().endswith('.txt'):
            filename += '.txt'
        
        # Format metrics as plain text
        try:
            metrics_json = json.loads(metrics_data)
            
            metrics_content = "PRANAS DATA ANALYZER - BREATHING METRICS\n"
            metrics_content += "=====================================\n\n"
            metrics_content += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
            
            for key, value in metrics_json.items():
                if key == 'timestamp':
                    metrics_content += f"\nTimestamp: {value}\n"
                elif key == 'resp_rate':
                    metrics_content += f"Respiratory Rate: {value}\n"
                elif key == 'lung_capacity':
                    metrics_content += f"Lung Capacity: {value}\n"
                elif key == 'lung_quality':
                    metrics_content += f"Lung Quality: {value}\n"
                else:
                    metrics_content += f"{key}: {value}\n"
        except Exception as e:
            # If there's an issue with JSON, use plain text
            metrics_content = f"Raw Metrics Data: {metrics_data}\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            return jsonify({
                'success': False,
                'message': 'OneDrive is not properly configured'
            }), 500
        
        # Create headers for API calls
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Get username for the folder
        username = current_user.username
        
        # Step 1: Find or create user folder
        user_folder_id = None
        
        # Look for user folder in root
        for endpoint in [
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://api.onedrive.com/v1.0/drive/root/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    root_items = response.json().get('value', [])
                    for item in root_items:
                        if item.get('name') == username and item.get('folder') is not None:
                            user_folder_id = item.get('id')
                            logging.info(f"Found user folder: {username}, id: {user_folder_id}")
                            break
                    
                    if user_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create user folder if it doesn't exist
        if not user_folder_id:
            logging.info(f"Creating user folder: {username}")
            create_folder_data = {
                'name': username,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                'https://graph.microsoft.com/v1.0/me/drive/root/children',
                'https://api.onedrive.com/v1.0/drive/root/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        user_folder_id = create_response.json().get('id')
                        logging.info(f"Created user folder: {username}, id: {user_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating user folder with {endpoint}: {str(e)}")
        
        if not user_folder_id:
            return jsonify({
                'success': False,
                'message': 'Could not find or create user folder in OneDrive'
            }), 500
        
        # Step 2: Find or create metrics subfolder
        metrics_folder_name = 'metrics'
        metrics_folder_id = None
        
        # Look for metrics folder in user folder
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
            f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    folder_items = response.json().get('value', [])
                    for item in folder_items:
                        if item.get('name') == metrics_folder_name and item.get('folder') is not None:
                            metrics_folder_id = item.get('id')
                            logging.info(f"Found metrics folder: {metrics_folder_name}, id: {metrics_folder_id}")
                            break
                    
                    if metrics_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create metrics folder if it doesn't exist
        if not metrics_folder_id:
            logging.info(f"Creating metrics folder: {metrics_folder_name}")
            create_folder_data = {
                'name': metrics_folder_name,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        metrics_folder_id = create_response.json().get('id')
                        logging.info(f"Created metrics folder: {metrics_folder_name}, id: {metrics_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating metrics folder with {endpoint}: {str(e)}")
        
        if not metrics_folder_id:
            return jsonify({
                'success': False,
                'message': f'Could not find or create {metrics_folder_name} folder in OneDrive'
            }), 500
        
        # Step 3: Upload metrics file to metrics folder
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'text/plain'  # Use text/plain for text files
        }
        
        file_uploaded = False
        text_bytes = metrics_content.encode('utf-8')  # Convert text to bytes
        
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{metrics_folder_id}:/{filename}:/content',
            f'https://api.onedrive.com/v1.0/drive/items/{metrics_folder_id}:/{filename}:/content'
        ]:
            try:
                upload_response = requests.put(
                    endpoint,
                    headers=upload_headers,
                    data=text_bytes  # Send text content as bytes
                )
                
                if upload_response.status_code in [200, 201]:
                    logging.info(f"Successfully uploaded metrics file to OneDrive: {filename}")
                    file_uploaded = True
                    break
            except Exception as e:
                logging.error(f"Error uploading metrics file with {endpoint}: {str(e)}")
        
        if not file_uploaded:
            return jsonify({
                'success': False,
                'message': 'Failed to upload metrics file to OneDrive'
            }), 500
        
        # Return success response
        return jsonify({
            'success': True,
            'message': f'Metrics successfully saved to OneDrive',
            'filename': filename,
            'folder': metrics_folder_name
        })
        
    except Exception as e:
        logging.error(f"Error saving metrics to OneDrive: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        }), 500
    
def is_file_within_user_scope(file_id, user_folder_id, headers):
    """
    Check if a file is within the user's permitted scope
    """
    try:
        # Get file info
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}',
            f'https://api.onedrive.com/v1.0/drive/items/{file_id}'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    
                    # Get parent reference
                    parent_ref = file_info.get('parentReference', {})
                    parent_id = parent_ref.get('id')
                    
                    # Check if the file's parent folder is within user scope
                    return is_folder_within_user_scope(parent_id, user_folder_id, headers, current_user.username)
            except Exception as e:
                logging.error(f"Error checking file scope with {endpoint}: {str(e)}")
                continue
        
        return False
        
    except Exception as e:
        logging.error(f"Error checking file scope: {str(e)}")
        return False

@app.route('/process-onedrive-file', methods=['POST'])
@login_required
def process_onedrive_file():
    """Process a file selected from OneDrive for raw data analysis"""
    # Get OneDrive file ID from form
    file_id = request.form.get('onedrive_file_id')
    if not file_id:
        flash("No OneDrive file selected", "error")
        return redirect(url_for('rd'))
    
    # Get access token
    access_token = get_master_onedrive_token()
    if not access_token:
        flash("OneDrive is not properly configured", "error")
        return redirect(url_for('rd'))
    
    # Validate that the file is within user's scope before processing
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # Get user's folder ID
    user_folder_id = None
    username = current_user.username
    
    try:
        # Find user folder
        response = requests.get(
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            headers=headers
        )
        
        if response.status_code == 200:
            root_items = response.json().get('value', [])
            for item in root_items:
                if item.get('name') == username and item.get('folder'):
                    user_folder_id = item.get('id')
                    break
        
        # Validate file is within user's scope
        if user_folder_id and not is_file_within_user_scope(file_id, user_folder_id, headers):
            flash("Access denied: file is outside your permitted scope", "error")
            return redirect(url_for('rd'))
        
    except Exception as e:
        logging.error(f"Error validating file scope: {str(e)}")
        flash("Error validating file access", "error")
        return redirect(url_for('rd'))
    
    # Continue with existing file processing logic...
    # [Rest of the existing process_onedrive_file function remains the same]
    
    try:
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details and download URL
        file_name = None
        file_content = None
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        # Download the file content
                        download_response = requests.get(download_url)
                        if download_response.status_code == 200:
                            file_content = download_response.content
                            break
                    else:
                        logging.warning("No download URL found in file info")
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        # If we couldn't get the file via the download URL, try direct content access
        if not file_content:
            for endpoint in [
                f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
                f"https://api.onedrive.com/v1.0/drive/items/{file_id}/content"
            ]:
                try:
                    response = requests.get(endpoint, headers=headers)
                    
                    if response.status_code == 200:
                        file_content = response.content
                        
                        # Try to get filename from Content-Disposition if we don't have it
                        if not file_name:
                            content_disposition = response.headers.get('Content-Disposition', '')
                            if 'filename=' in content_disposition:
                                file_name = content_disposition.split('filename=')[1].strip('"')
                            else:
                                file_name = f"file_{file_id}.csv"
                        break
                except Exception as e:
                    logging.error(f"Error with {endpoint}: {str(e)}")
                    continue
        
        if not file_content:
            flash("Could not download file from OneDrive", "error")
            return redirect(url_for('rd'))
        
        if not file_name:
            file_name = f"unknown_file_{file_id}.csv"
            
        # Process the file with the raw_data_calc function
        try:
            # Convert to a file-like object for immediate processing
            from io import BytesIO
            temp_file = BytesIO(file_content)
            
            # Read CSV directly from the BytesIO object
            df = pd.read_csv(temp_file)
            
            # Process the data
            result = raw_data_calc(df)
            
            # Generate a unique stored filename
            stored_filename = f"{uuid.uuid4().hex}_{secure_filename(file_name)}"
            
            # Save file to disk in user's folder
            user_folder = os.path.join(app.config['UPLOAD_FOLDER'], str(current_user.id))
            os.makedirs(user_folder, exist_ok=True)
            file_path = os.path.join(user_folder, stored_filename)
            
            with open(file_path, 'wb') as f:
                f.write(file_content)
            
            # Get file size
            file_size = len(file_content)
            
            # Create file record in database
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    # First, check which columns exist
                    cur.execute("""
                        SELECT column_name 
                        FROM information_schema.columns 
                        WHERE table_name = 'uploaded_files'
                    """)
                    columns = [row['column_name'] for row in cur.fetchall()]
                    
                    # Build SQL dynamically based on existing columns
                    sql_columns = ['user_id', 'original_filename', 'stored_filename']
                    sql_values = [current_user.id, file_name, stored_filename]
                    sql_placeholders = ['%s', '%s', '%s']
                    
                    if 'file_path' in columns:
                        sql_columns.append('file_path')
                        sql_values.append(file_path)
                        sql_placeholders.append('%s')
                        
                    if 'file_type' in columns:
                        sql_columns.append('file_type')
                        sql_values.append('text/csv')
                        sql_placeholders.append('%s')
                        
                    if 'file_size' in columns:
                        sql_columns.append('file_size')
                        sql_values.append(file_size)
                        sql_placeholders.append('%s')
                        
                    if 'source' in columns:
                        sql_columns.append('source')
                        sql_values.append('onedrive')
                        sql_placeholders.append('%s')
                    
                    # Insert the record
                    sql = f"""
                        INSERT INTO uploaded_files 
                        ({', '.join(sql_columns)})
                        VALUES ({', '.join(sql_placeholders)})
                        RETURNING id, uploaded_at
                    """
                    cur.execute(sql, sql_values)
                    result_row = cur.fetchone()
                    file_id = result_row['id']
                    uploaded_at = result_row['uploaded_at']
                    conn.commit()
                    
                    # Create file record object
                    file_record = UploadedFile(
                        id=file_id,
                        user_id=current_user.id,
                        original_filename=file_name,
                        file_type='text/csv',
                        file_size=file_size,
                        uploaded_at=uploaded_at,
                        source='onedrive',
                        file_path=file_path
                    )
                    
                # Update file record with processing information
                with conn.cursor() as cur:
                    cur.execute("""
                        UPDATE uploaded_files 
                        SET processed = TRUE, processed_at = CURRENT_TIMESTAMP, processing_type = 'raw_data'
                        WHERE id = %s
                    """, (file_id,))
                    conn.commit()
                    
                    # Update the file_record object
                    file_record.processed = True
                    file_record.processed_at = datetime.utcnow()
                    file_record.processing_type = 'raw_data'
                
            except Exception as e:
                conn.rollback()
                logging.error(f"Database error: {str(e)}")
                raise e
            finally:
                conn.close()
            
            # Return the template with results
            return render_template('Raw_signalBreathe.HTML', 
                               img_base64=result,
                               file_record=file_record,
                               sample_placeholder="Analysis completed successfully!")
                               
        except Exception as e:
            logging.error(f"Error processing OneDrive file: {str(e)}")
            flash(f"Error processing file: {str(e)}", "error")
            return redirect(url_for('rd'))
            
    except Exception as e:
        logging.error(f"Error accessing OneDrive file: {str(e)}")
        flash(f"Error accessing OneDrive file: {str(e)}", "error")
        return redirect(url_for('rd'))
    
@app.route('/api/save-masterdata-to-onedrive', methods=['POST'])
@login_required
def save_masterdata_to_onedrive():
    """API endpoint to save masterdata file to OneDrive - FIXED"""
    try:
        # Get the masterdata file path from session
        masterdata_path = session.get('temp_masterdata_path')
        
        if not masterdata_path or not os.path.exists(masterdata_path):
            return jsonify({
                'success': False,
                'message': 'No masterdata file available to save'
            }), 400
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            return jsonify({
                'success': False,
                'message': 'OneDrive is not properly configured'
            }), 500
        
        # Extract folder name from the session filename
        masterdata_name = session.get('temp_masterdata_name', 'Masterdata.csv')
        # Try to extract folder name from filename pattern
        folder_name = 'processed_data'
        if '_Masterdata_' in masterdata_name:
            folder_name = masterdata_name.split('_Masterdata_')[0]
        
        # Upload to OneDrive using the improved function
        upload_success = upload_masterdata_to_onedrive_improved(
            masterdata_path, 
            folder_name, 
            access_token
        )
        
        if upload_success:
            # Create the expected filename that was uploaded
            timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            onedrive_filename = f"{folder_name}_Masterdata_{timestamp}.csv"
            
            return jsonify({
                'success': True,
                'message': 'File successfully saved to OneDrive',
                'filename': onedrive_filename,
                'location': f'{current_user.username}/processed_data/'
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Failed to upload file to OneDrive'
            }), 500
        
    except Exception as e:
        logging.error(f"Error saving masterdata to OneDrive: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        }), 500
    
@app.route('/api/save-to-onedrive', methods=['POST'])
@login_required
def save_to_onedrive_api():
    """API endpoint to save a file to the user's OneDrive folder"""
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({
                'success': False,
                'message': 'No file part in the request'
            }), 400
            
        file = request.files['file']
        if file.filename == '':
            return jsonify({
                'success': False,
                'message': 'No selected file'
            }), 400
        
        # Get folder path (defaulting to 'results')
        folder_path = request.form.get('folder_path', 'results')
        
        # Read file content
        file_content = file.read()
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            return jsonify({
                'success': False,
                'message': 'OneDrive is not properly configured'
            }), 500
        
        # Create headers for API calls
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Get username for the folder
        username = current_user.username
        
        # Step 1: Find or create user folder
        user_folder_id = None
        
        # Look for user folder in root
        for endpoint in [
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://api.onedrive.com/v1.0/drive/root/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    root_items = response.json().get('value', [])
                    for item in root_items:
                        if item.get('name') == username and item.get('folder') is not None:
                            user_folder_id = item.get('id')
                            logging.info(f"Found user folder: {username}, id: {user_folder_id}")
                            break
                    
                    if user_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create user folder if it doesn't exist
        if not user_folder_id:
            logging.info(f"Creating user folder: {username}")
            create_folder_data = {
                'name': username,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                'https://graph.microsoft.com/v1.0/me/drive/root/children',
                'https://api.onedrive.com/v1.0/drive/root/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        user_folder_id = create_response.json().get('id')
                        logging.info(f"Created user folder: {username}, id: {user_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating user folder with {endpoint}: {str(e)}")
        
        if not user_folder_id:
            return jsonify({
                'success': False,
                'message': 'Could not find or create user folder in OneDrive'
            }), 500
        
        # Step 2: Find or create results subfolder
        results_folder_id = None
        
        # Look for results folder in user folder
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
            f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    folder_items = response.json().get('value', [])
                    for item in folder_items:
                        if item.get('name') == folder_path and item.get('folder') is not None:
                            results_folder_id = item.get('id')
                            logging.info(f"Found results folder: {folder_path}, id: {results_folder_id}")
                            break
                    
                    if results_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create results folder if it doesn't exist
        if not results_folder_id:
            logging.info(f"Creating results folder: {folder_path}")
            create_folder_data = {
                'name': folder_path,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        results_folder_id = create_response.json().get('id')
                        logging.info(f"Created results folder: {folder_path}, id: {results_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating results folder with {endpoint}: {str(e)}")
        
        if not results_folder_id:
            return jsonify({
                'success': False,
                'message': f'Could not find or create {folder_path} folder in OneDrive'
            }), 500
        
        # Step 3: Upload file to results folder
        filename = secure_filename(file.filename)
        
        # For simple upload (< 4MB)
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/octet-stream'
        }
        
        file_uploaded = False
        
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{results_folder_id}:/{filename}:/content',
            f'https://api.onedrive.com/v1.0/drive/items/{results_folder_id}:/{filename}:/content'
        ]:
            try:
                upload_response = requests.put(
                    endpoint,
                    headers=upload_headers,
                    data=file_content
                )
                
                if upload_response.status_code in [200, 201]:
                    logging.info(f"Successfully uploaded file to OneDrive: {filename}")
                    file_uploaded = True
                    break
            except Exception as e:
                logging.error(f"Error uploading file with {endpoint}: {str(e)}")
        
        if not file_uploaded:
            return jsonify({
                'success': False,
                'message': 'Failed to upload file to OneDrive'
            }), 500
        
        # Return success response
        return jsonify({
            'success': True,
            'message': f'File successfully saved to OneDrive',
            'filename': filename,
            'folder': folder_path
        })
        
    except Exception as e:
        logging.error(f"Error saving to OneDrive: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        }), 500

@app.route('/rd')
def rd():
    if current_user.is_authenticated:
        # Get user's recent raw data analyses if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        raw_files = [f for f in user_files if f.processing_type == 'raw_data'][:5]
    else:
        raw_files = []
        
    return render_template('Raw_signalBreathe.HTML', 
                         sample_placeholder="Upload a CSV file to see raw signal analysis results here.",
                         user_files=raw_files)

@app.route('/redirect_to_rd')
def redirect_to_rd():
    return redirect(url_for('rd'))

@app.route('/fvc')
def fvc():
    if current_user.is_authenticated:
        # Get user's FVC files if logged in
        user_files = UploadedFile.get_user_files(current_user.id)
        fvc_files = [f for f in user_files if f.processing_type == 'fvc'][:5]
    else:
        fvc_files = []
        
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('FVC.HTML', 
                         sample_placeholder=prediction_text_placeholder,
                         user_files=fvc_files)

@app.route('/onedrive-files')
@login_required
def onedrive_files():
    """List files from the user's folder in OneDrive"""
    # Get the access token
    access_token = get_master_onedrive_token()
    
    if not access_token:
        flash("OneDrive is not configured yet. Please ask an administrator to set it up.", "error")
        return redirect(url_for('dashboard'))
    
    # Get user's folder name (username)
    username = current_user.username
    files = []
    user_folder_id = None
    
    try:
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Step 1: Look for user's folder in root
        logging.info("Attempting to list root items from OneDrive")
        response = requests.get(
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            headers=headers
        )
        
        if response.status_code != 200:
            # Try alternative endpoint for OneDrive personal
            logging.info("First attempt failed, trying alternative endpoint")
            response = requests.get(
                'https://api.onedrive.com/v1.0/drive/root/children',
                headers=headers
            )
            
        if response.status_code == 200:
            # Look for user's folder
            root_items = response.json().get('value', [])
            for item in root_items:
                if item.get('name') == username and item.get('folder'):
                    user_folder_id = item.get('id')
                    logging.info(f"Found user folder: {username}, id: {user_folder_id}")
                    break
            
            # If user folder doesn't exist, create it
            if not user_folder_id:
                logging.info(f"User folder not found, creating new folder: {username}")
                # Try both APIs for creating folder
                endpoints = [
                    'https://graph.microsoft.com/v1.0/me/drive/root/children',
                    'https://api.onedrive.com/v1.0/drive/root/children'
                ]
                
                folder_data = {
                    'name': username,
                    'folder': {},
                    '@microsoft.graph.conflictBehavior': 'rename'
                }
                
                for endpoint in endpoints:
                    try:
                        create_response = requests.post(
                            endpoint,
                            headers={**headers, 'Content-Type': 'application/json'},
                            json=folder_data
                        )
                        
                        if create_response.status_code in [200, 201]:
                            user_folder_id = create_response.json().get('id')
                            logging.info(f"Created user folder with {endpoint}, id: {user_folder_id}")
                            break
                        else:
                            logging.warning(f"Failed to create folder with {endpoint}: {create_response.status_code}")
                    except Exception as e:
                        logging.error(f"Error creating folder with {endpoint}: {str(e)}")
                
                if not user_folder_id:
                    flash("Could not create your folder in OneDrive", "error")
            
            # If we have a user folder ID, get its contents
            if user_folder_id:
                logging.info(f"Getting files from user folder: {user_folder_id}")
                # Try both APIs for listing files
                endpoints = [
                    f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                    f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
                ]
                
                for endpoint in endpoints:
                    try:
                        files_response = requests.get(endpoint, headers=headers)
                        
                        if files_response.status_code == 200:
                            files = files_response.json().get('value', [])
                            logging.info(f"Successfully retrieved {len(files)} files with {endpoint}")
                            break
                        else:
                            logging.warning(f"Failed to get files with {endpoint}: {files_response.status_code}")
                    except Exception as e:
                        logging.error(f"Error getting files with {endpoint}: {str(e)}")
                
                if not files:
                    logging.warning("No files found or could not access files in user folder")
            else:
                logging.error("Could not find or create user folder")
        else:
            error_text = "Unknown error"
            try:
                error_data = response.json()
                error_text = f"{error_data.get('error', {}).get('code')}: {error_data.get('error', {}).get('message')}"
            except:
                error_text = response.text[:200]
                
            logging.error(f"OneDrive API error: {error_text}")
            flash("Could not access OneDrive. Please try again later.", "error")
            
    except Exception as e:
        logging.error(f"OneDrive error: {str(e)}")
        flash(f"Error accessing OneDrive: {str(e)}", "error")
    
    # Get OneDrive user info
    onedrive_user = {
        'display_name': SystemConfig.get_value('onedrive_user_name') or 'OneDrive User',
        'email': SystemConfig.get_value('onedrive_user_email') or 'Connected Account'
    }
    
    folder_info = {
        'name': username,
        'id': user_folder_id,
        'exists': user_folder_id is not None
    }
    
    return render_template('onedrive_files.html',
                         files=files,
                         folder_info=folder_info,
                         onedrive_user=onedrive_user)

@app.route('/svc')
def svc():
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('SVC.HTML', sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_svc')
def redirect_to_svc():
    return redirect(url_for('svc'))

@app.route('/mvv')
def mvv():
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('MVV.HTML',sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_mvv')
def redirect_to_mvv():
    return redirect(url_for('mvv'))


@app.route('/redirect_to_fvc')
def redirect_to_fvc():
    return redirect(url_for('fvc'))

@app.route('/logout-onedrive')
@login_required
def logout_onedrive():
    """Disconnect from OneDrive by removing session data"""
    # Remove OneDrive-related session data
    keys_to_remove = ['onedrive_token', 'onedrive_refresh_token', 
                      'onedrive_token_expires', 'onedrive_user']
    
    for key in keys_to_remove:
        if key in session:
            del session[key]
    
    session.modified = True
    flash("Successfully disconnected from OneDrive", "success")
    return redirect(url_for('dashboard'))

@app.route('/select-onedrive-file/<file_id>', methods=['POST'])
@login_required
def select_onedrive_file(file_id):
    """Download and process a file from OneDrive"""
    # Get the access token
    access_token = get_master_onedrive_token()
    if not access_token:
        flash("OneDrive is not properly configured. Please contact the administrator.", "error")
        return redirect(url_for('dashboard'))
    
    processing_type = request.form.get('processing_type', 'raw_visual')
    
    try:
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details and download URL
        # Try both Graph API and OneDrive API
        file_name = None
        file_content = None
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    
                    # Get download URL - different keys for different APIs
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        logging.info(f"Got download URL for {file_name}")
                        # Download the file content
                        download_response = requests.get(download_url)
                        if download_response.status_code == 200:
                            file_content = download_response.content
                            break
                        else:
                            logging.error(f"Download failed: {download_response.status_code}")
                    else:
                        logging.warning("No download URL found in file info")
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # If we couldn't get the file via the download URL, try direct content access
        if not file_content:
            logging.info("Trying direct content access")
            for endpoint in [
                f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
                f"https://api.onedrive.com/v1.0/drive/items/{file_id}/content"
            ]:
                try:
                    response = requests.get(endpoint, headers=headers)
                    
                    if response.status_code == 200:
                        file_content = response.content
                        
                        # Try to get filename from Content-Disposition if we don't have it
                        if not file_name:
                            content_disposition = response.headers.get('Content-Disposition', '')
                            if 'filename=' in content_disposition:
                                file_name = content_disposition.split('filename=')[1].strip('"')
                            else:
                                file_name = f"file_{file_id}"
                                
                        logging.info(f"Got file content directly: {file_name}")
                        break
                except Exception as e:
                    logging.error(f"Error with {endpoint}: {str(e)}")
        
        if not file_content:
            flash("Could not download file from OneDrive", "error")
            return redirect(url_for('onedrive_files'))
        
        if not file_name:
            file_name = f"unknown_file_{file_id}.csv"
            
        # Save the file for the current user
        try:
            # Determine file type based on extension
            file_type = 'text/csv'  # Default
            if file_name.lower().endswith('.csv'):
                file_type = 'text/csv'
            elif file_name.lower().endswith(('.xlsx', '.xls')):
                file_type = 'application/vnd.ms-excel'
            elif file_name.lower().endswith('.txt'):
                file_type = 'text/plain'
            
            # Create file record in database
            file_record = UploadedFile.create(
                user_id=current_user.id,
                original_filename=file_name,
                file_content=file_content,
                file_type=file_type,
                file_size=len(file_content)
            )
            
            # Update the record to show it came from OneDrive
            file_record.source = 'onedrive'
            file_record.update()
            
            flash(f"Successfully imported file '{file_name}' from OneDrive", "success")
            
            # Redirect based on processing type
            if processing_type == 'raw_visual':
                return redirect(url_for('NewrawVisual'))
            elif processing_type == 'pca':
                return redirect(url_for('NewPca'))
            elif processing_type == 'clustering_KMS':
                return redirect(url_for('clusterwebpage'))
            elif processing_type == 'ml':
                return redirect(url_for('NewML'))
            else:
                return redirect(url_for('dashboard'))
                
        except Exception as e:
            logging.error(f"Error creating file record: {str(e)}")
            flash(f"Error saving file: {str(e)}", "error")
            return redirect(url_for('onedrive_files'))
            
    except Exception as e:
        logging.error(f"Error processing OneDrive file: {str(e)}")
        flash(f"Error processing file: {str(e)}", "error")
        return redirect(url_for('onedrive_files'))

@app.route('/sm')
def sm():
    """Standard Metrics page"""
    # DEBUGGING: Log that we're entering the sm route
    logging.info("Entered sm route")
    
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('Standard_Metrics_Breathe.HTML', sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_sm')  # This MUST be different from '/sm'
def redirect_to_sm():
    """Redirect to the Standard Metrics page"""
    # DEBUGGING: Log that we're redirecting
    logging.info("Redirecting to sm")
    
    return redirect(url_for('sm'))

@app.route('/doghome')
def doghome():
    prediction_text_placeholder = "Please wait for processor response"
    return render_template('dogBreatheHome.HTML', sample_placeholder=prediction_text_placeholder)

@app.route('/redirect_to_doghome')
@login_required
def redirect_to_doghome():
    return redirect(url_for('doghome'))

# === API Processing Routes ===

# Replace your existing datamaker route with this updated version

@app.route("/datamaker", methods=['POST'])
@login_required
def datamaker():
    """Handle both local and OneDrive data processing with improved error handling"""
    if request.method == 'POST':
        data_source = request.form.get('dataSource', 'local')
        
        if data_source == 'onedrive':
            # Process OneDrive folder
            folder_id = request.form.get('onedrive_folder_id')
            folder_name = request.form.get('onedrive_folder_name')
            
            if not folder_id or not folder_name:
                return render_template('DMwebpage.html', 
                                     sample_placeholder="Error: Please select a valid OneDrive folder")
            
            access_token = get_master_onedrive_token()
            if not access_token:
                return render_template('DMwebpage.html', 
                                     sample_placeholder="Error: OneDrive not configured properly")
            
            temp_base_path = None
            try:
                logging.info(f"Processing OneDrive folder: {folder_name} (ID: {folder_id})")
                
                # Create temporary local folder structure
                temp_base_path = create_temp_folder_structure(folder_name)
                logging.info(f"Created temporary folder: {temp_base_path}")
                
                # Download all files from OneDrive folder (including subfolders)
                download_success = download_onedrive_folder_contents(folder_id, temp_base_path, access_token)
                
                if not download_success:
                    cleanup_temp_folder(temp_base_path)
                    return render_template('DMwebpage.html', 
                                         sample_placeholder="Error: No compatible files found in the selected OneDrive folder")
                
                logging.info(f"Downloaded files from OneDrive successfully")
                
                # Process using existing data processor logic
                result = dp1.create_and_move_csv(temp_base_path)
                logging.info(f"Data processing result: {result}")
                
                if result == "Successfull":
                    # Search for masterdata file in multiple possible locations
                    project_root = os.getcwd()
                    masterdata_path = find_masterdata_file(project_root)
                    
                    # If not found in project root, search in temp folder
                    if not masterdata_path:
                        logging.info(f"Searching in temp folder: {temp_base_path}")
                        masterdata_path = find_masterdata_file(temp_base_path)
                    
                    if masterdata_path:
                        logging.info(f"Found masterdata file at: {masterdata_path}")
                        
                        # Create a properly named copy with timestamp
                        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                        new_filename = f"{folder_name}_Masterdata_{timestamp}.csv"
                        
                        # Create a new path for the renamed file in temp directory
                        temp_dir = os.path.join(os.getcwd(), 'temp_downloads')
                        os.makedirs(temp_dir, exist_ok=True)
                        new_masterdata_path = os.path.join(temp_dir, new_filename)
                        
                        # Copy the original file to the new location with new name
                        import shutil
                        shutil.copy2(masterdata_path, new_masterdata_path)
                        
                        # Get file size for display
                        file_size = os.path.getsize(new_masterdata_path)
                        file_size_mb = round(file_size / (1024 * 1024), 2)
                        
                        # Store the new file path in session for download
                        session['temp_masterdata_path'] = new_masterdata_path
                        session['temp_masterdata_name'] = new_filename
                        session.modified = True
                        
                        # Try to upload results back to OneDrive
                        upload_success = False
                        try:
                            upload_success = upload_masterdata_to_onedrive_improved(
                                new_masterdata_path, 
                                folder_name, 
                                access_token
                            )
                        except Exception as e:
                            logging.error(f"Error uploading to OneDrive: {str(e)}")
                        
                        # Prepare simple success message
                        if upload_success:
                            success_message = f"""Operation Successful! Master data has been created and saved.

File: {new_filename}
Size: {file_size_mb} MB
OneDrive: Automatically saved
Status: Ready for download and analysis

The master dataset is ready for use in Current Analysis v2."""
                        else:
                            success_message = f"""Operation Successful! Master data has been created.

File: {new_filename}
Size: {file_size_mb} MB
Status: Ready for download and analysis

The master dataset is ready for use in Current Analysis v2."""
                        
                        # Don't cleanup yet - keep file for download
                        return render_template('DMwebpage.html', 
                                             sample_placeholder=success_message,
                                             show_download=True,
                                             file_size_mb=file_size_mb,
                                             onedrive_saved=upload_success)
                    else:
                        # Simple error message
                        error_message = f"""Error: Master data file was not created

Processing Status: {result}
Folder: {folder_name}

The data processing completed but no master data file was generated. 
Please check if the folder contains valid data files and try again."""
                        
                        cleanup_temp_folder(temp_base_path)
                        return render_template('DMwebpage.html', 
                                             sample_placeholder=error_message)
                else:
                    cleanup_temp_folder(temp_base_path)
                    return render_template('DMwebpage.html', 
                                         sample_placeholder=f"Error: Processing failed - {result}")
                    
            except Exception as e:
                logging.error(f"Error processing OneDrive folder: {str(e)}")
                if temp_base_path:
                    cleanup_temp_folder(temp_base_path)
                return render_template('DMwebpage.html', 
                                     sample_placeholder=f"Error: {str(e)}")
        else:
            # Local processing logic
            folder_path = request.form.get('folder')
            if not folder_path:
                return render_template('DMwebpage.html', 
                                     sample_placeholder="Error: Please enter a valid folder path")
            
            # Build the full folder path
            full_folder_path = 'Scripts/Bacteria_Data/' + folder_path
            result = dp1.create_and_move_csv(full_folder_path)
            logging.info(f"Local processing result: {result} for path: {full_folder_path}")
            
            if result == "Successfull":
                # Search for masterdata file
                project_root = os.getcwd()
                masterdata_path = find_masterdata_file(project_root)
                
                # If not found in project root, search in the processed folder
                if not masterdata_path:
                    logging.info(f"Searching in processed folder: {full_folder_path}")
                    masterdata_path = find_masterdata_file(full_folder_path)
                
                if masterdata_path:
                    logging.info(f"Found masterdata file at: {masterdata_path}")
                    
                    # Create a properly named copy with timestamp
                    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                    new_filename = f"{folder_path}_Masterdata_{timestamp}.csv"
                    
                    # Create a new path for the renamed file
                    temp_dir = os.path.join(os.getcwd(), 'temp_downloads')
                    os.makedirs(temp_dir, exist_ok=True)
                    new_masterdata_path = os.path.join(temp_dir, new_filename)
                    
                    # Copy the original file to the new location with new name
                    import shutil
                    shutil.copy2(masterdata_path, new_masterdata_path)
                    
                    file_size = os.path.getsize(new_masterdata_path)
                    file_size_mb = round(file_size / (1024 * 1024), 2)
                    
                    # Store the new file path in session for download
                    session['temp_masterdata_path'] = new_masterdata_path
                    session['temp_masterdata_name'] = new_filename
                    session.modified = True
                    
                    # Simple success message
                    success_message = f"""Operation Successful! Master data has been created.

File: {new_filename}
Size: {file_size_mb} MB
Status: Ready for download and analysis

The master dataset is ready for use in Current Analysis v2."""
                    
                    return render_template('DMwebpage.html', 
                                         sample_placeholder=success_message,
                                         show_download=True,
                                         file_size_mb=file_size_mb,
                                         onedrive_saved=False)
                else:
                    # Simple error message for local processing
                    error_message = f"""Error: Master data file was not created

Processing Status: {result}
Folder: Scripts/Bacteria_Data/{folder_path}

The data processing completed but no master data file was generated.
Please check if the folder exists and contains valid data files."""
                        
                    return render_template('DMwebpage.html', 
                                         sample_placeholder=error_message)
            else:
                return render_template('DMwebpage.html', 
                                     sample_placeholder=f"Error: Processing failed - {result}")
    else:
        return render_template('page2breathe.html')
    
# Replace your existing download-masterdata route with this fixed version

@app.route('/download-masterdata')
@login_required
def download_masterdata():
    """Download masterdata file directly from OneDrive - AGGRESSIVE CLEANUP VERSION"""
    try:
        logging.info("Download request - Using OneDrive-only approach")
        
        # Get OneDrive access token
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available")
            flash("OneDrive not configured. Please contact administrator.", "error")
            return redirect(url_for('DMwebpage'))
        
        # Find and download the most recent masterdata file from OneDrive
        username = current_user.username
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Find user's folder in OneDrive
        user_folder_id = None
        for endpoint in [
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://api.onedrive.com/v1.0/drive/root/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                if response.status_code == 200:
                    root_items = response.json().get('value', [])
                    for item in root_items:
                        if item.get('name') == username and item.get('folder'):
                            user_folder_id = item.get('id')
                            break
                    if user_folder_id:
                        break
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        if not user_folder_id:
            logging.error("User folder not found in OneDrive")
            flash("Your folder not found in OneDrive", "error")
            return redirect(url_for('DMwebpage'))
        
        # Look for processed_data folder
        processed_folder_id = None
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
            f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
        ]:
            try:
                folder_response = requests.get(endpoint, headers=headers)
                if folder_response.status_code == 200:
                    folder_items = folder_response.json().get('value', [])
                    for item in folder_items:
                        if item.get('name') == 'processed_data' and item.get('folder'):
                            processed_folder_id = item.get('id')
                            break
                    if processed_folder_id:
                        break
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        if not processed_folder_id:
            logging.error("Processed_data folder not found in OneDrive")
            flash("Processed data folder not found in OneDrive", "error")
            return redirect(url_for('DMwebpage'))
        
        # Get files from processed_data folder and find the most recent masterdata file
        files = []
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{processed_folder_id}/children',
            f'https://api.onedrive.com/v1.0/drive/items/{processed_folder_id}/children'
        ]:
            try:
                files_response = requests.get(endpoint, headers=headers)
                if files_response.status_code == 200:
                    files = files_response.json().get('value', [])
                    break
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        if not files:
            logging.error("Could not access files in processed_data folder")
            flash("Could not access files in OneDrive", "error")
            return redirect(url_for('DMwebpage'))
        
        # Filter for masterdata files and find the most recent one
        masterdata_files = []
        for file_item in files:
            file_name = file_item.get('name', '').lower()
            if file_name.endswith('.csv') and 'masterdata' in file_name:
                masterdata_files.append(file_item)
        
        if not masterdata_files:
            logging.error("No masterdata files found in OneDrive processed_data folder")
            flash("No master data files found in OneDrive", "error")
            return redirect(url_for('DMwebpage'))
        
        # Sort by creation date (most recent first)
        masterdata_files.sort(key=lambda x: x.get('createdDateTime', ''), reverse=True)
        latest_file = masterdata_files[0]
        
        logging.info(f"Downloading latest masterdata file from OneDrive: {latest_file.get('name')}")
        
        # Download the file content from OneDrive
        file_id = latest_file.get('id')
        file_content = None
        
        # Try different download methods
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}/content"
        ]:
            try:
                download_response = requests.get(endpoint, headers=headers)
                if download_response.status_code == 200:
                    file_content = download_response.content
                    break
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        # If direct download fails, try getting download URL
        if not file_content:
            for info_endpoint in [
                f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
                f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
            ]:
                try:
                    info_response = requests.get(info_endpoint, headers=headers)
                    if info_response.status_code == 200:
                        file_info = info_response.json()
                        download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                        
                        if download_url:
                            download_response = requests.get(download_url)
                            if download_response.status_code == 200:
                                file_content = download_response.content
                                break
                except Exception as e:
                    logging.error(f"Error with {info_endpoint}: {str(e)}")
                    continue
        
        if not file_content:
            logging.error("Failed to download file from OneDrive")
            flash("Failed to download master data file from OneDrive", "error")
            return redirect(url_for('DMwebpage'))
        
        # Create a BytesIO object from the downloaded content
        file_data = BytesIO(file_content)
        
        # Use the original filename from OneDrive
        download_filename = latest_file.get('name')
        if not download_filename.lower().endswith('.csv'):
            download_filename += '.csv'
        
        logging.info(f"Successfully downloading from OneDrive: {download_filename} ({len(file_content)} bytes)")
        
        # Send the file for download
        return send_file(
            file_data,
            as_attachment=True,
            download_name=download_filename,
            mimetype='text/csv'
        )
        
    except Exception as e:
        logging.error(f"Critical error in download_masterdata: {str(e)}")
        flash(f"Error downloading file: {str(e)}", "error")
        return redirect(url_for('DMwebpage'))



@app.route('/cleanup-temp-files', methods=['POST'])
@login_required
def cleanup_temp_files():
    """Aggressive cleanup - remove all temporary files immediately"""
    try:
        masterdata_path = session.get('temp_masterdata_path')
        
        # Immediately remove the file if it exists
        if masterdata_path and os.path.exists(masterdata_path):
            try:
                os.remove(masterdata_path)
                logging.info(f"Aggressively cleaned up masterdata file: {masterdata_path}")
            except Exception as e:
                logging.warning(f"Could not delete masterdata file: {str(e)}")
        
        # Clean up the temp_downloads directory
        temp_downloads_dir = os.path.join(os.getcwd(), 'temp_downloads')
        if os.path.exists(temp_downloads_dir):
            try:
                for file_path in os.listdir(temp_downloads_dir):
                    full_path = os.path.join(temp_downloads_dir, file_path)
                    if os.path.isfile(full_path):
                        os.remove(full_path)
                        logging.info(f"Aggressively cleaned up temp file: {full_path}")
                
                # Remove directory if empty
                if not os.listdir(temp_downloads_dir):
                    os.rmdir(temp_downloads_dir)
                    logging.info(f"Removed empty temp_downloads directory")
            except Exception as e:
                logging.warning(f"Could not clean up temp_downloads directory: {str(e)}")
        
        # Clean up temp processing folders immediately
        temp_processing_base = 'temp_processing'
        if os.path.exists(temp_processing_base):
            try:
                for folder_name in os.listdir(temp_processing_base):
                    folder_path = os.path.join(temp_processing_base, folder_name)
                    if os.path.isdir(folder_path):
                        shutil.rmtree(folder_path, ignore_errors=True)
                        logging.info(f"Aggressively cleaned up temp folder: {folder_path}")
            except Exception as e:
                logging.warning(f"Error cleaning up temp processing folders: {str(e)}")
        
        # Clear all session data related to temp files
        session.pop('temp_masterdata_path', None)
        session.pop('temp_masterdata_name', None)
        session.modified = True
        
        return '', 204  # No content response
        
    except Exception as e:
        logging.error(f"Error in aggressive cleanup: {str(e)}")
        return '', 500
    
@app.route('/api/onedrive-folders')
@login_required
def get_onedrive_folders():
    """API endpoint to get folders from user's OneDrive folder"""
    access_token = get_master_onedrive_token()
    if not access_token:
        return jsonify({
            'success': False,
            'message': 'OneDrive is not properly configured'
        })
    
    username = current_user.username
    folders = []
    user_folder_id = None
    
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Find user's folder in root
        response = requests.get(
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            headers=headers
        )
        
        if response.status_code != 200:
            response = requests.get(
                'https://api.onedrive.com/v1.0/drive/root/children',
                headers=headers
            )
            
        if response.status_code == 200:
            root_items = response.json().get('value', [])
            
            # First, look for user folder
            for item in root_items:
                if item.get('name') == username and item.get('folder'):
                    user_folder_id = item.get('id')
                    logging.info(f"Found user folder: {username}, id: {user_folder_id}")
                    break
            
            # Get folders within user folder
            if user_folder_id:
                folders_response = requests.get(
                    f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                    headers=headers
                )
                
                if folders_response.status_code != 200:
                    folders_response = requests.get(
                        f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children',
                        headers=headers
                    )
                
                if folders_response.status_code == 200:
                    items = folders_response.json().get('value', [])
                    # Filter only folders
                    folders = [item for item in items if item.get('folder') is not None]
                    logging.info(f"Found {len(folders)} folders in user directory")
            
            # If no user folder found, get all folders from root
            if not folders:
                logging.info("No user folder found, getting all folders from root")
                folders = [item for item in root_items if item.get('folder') is not None]
        
        return jsonify({
            'success': True,
            'folders': folders,
            'user_folder_id': user_folder_id,
            'debug_info': f"Found {len(folders)} folders"
        })
    except Exception as e:
        logging.error(f"Error accessing OneDrive folders: {str(e)}")
        return jsonify({
            'success': False,
            'message': f"Error connecting to OneDrive: {str(e)}"
        })

@app.route("/dataclustering", methods=['POST'])
@login_required
def dataclustering():
    if request.method == 'POST':
        visit = request.form.get("visit")
        
        if visit == "one":
            # Get the clustering option selected by the user
            clustering_option = request.form.get('clusteringOption')
            if not clustering_option:
                flash("Please select a clustering algorithm option.", "error")
                return redirect(url_for('clusterwebpage'))
            
            # Get the uploaded file
            if 'csv_file' not in request.files:
                flash("No file part", "error")
                return redirect(url_for('clusterwebpage'))
                
            folder_path = request.files['csv_file']
            if folder_path.filename == '':
                flash("No selected file", "error")
                return redirect(url_for('clusterwebpage'))
            
            # Check file extension
            if not folder_path.filename.lower().endswith('.csv'):
                flash("Only CSV files are allowed", "error")
                return redirect(url_for('clusterwebpage'))
            
            # Create user-specific filename with UUID to avoid collisions
            original_filename = folder_path.filename
            unique_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
            file_path = get_user_file_path(unique_filename)
            
            # Save the uploaded file
            folder_path.save(file_path)
            
            # Create file record in database
            try:
                file_record = add_file_record(file_path, original_filename, folder_path.content_type)
                if not file_record:
                    flash("Error creating file record", "error")
                    return redirect(url_for('clusterwebpage'))
                
                # Store the file ID and clustering option in the session
                session['file_id'] = file_record.id
                session['clusteringOption'] = clustering_option
                
                # Log the successful upload
                logging.info(f"User {current_user.id} uploaded file {file_record.id} for clustering analysis")
                
                # Read the CSV data into a DataFrame and extract unique values
                dfX = pd.read_csv(file_path)
                
                # Extract unique values from various columns
                unq_bacts = dfX['bacteria'].unique() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique() if 'slide' in dfX.columns else []
                
                strings = unq_bacts.tolist()
                unq_concs = unq_concs.tolist()
                unq_vols = unq_vols.tolist()
                unq_sli = unq_sli.tolist()
                
                return render_template('clusterwebpage.html', 
                                      item=strings, 
                                      item2=unq_concs, 
                                      item3=unq_vols, 
                                      item4=unq_sli)
            except Exception as e:
                # If there's an error, log it and notify the user
                logging.error(f"Error processing uploaded file: {str(e)}")
                flash(f"Error processing file: {str(e)}", "error")
                return redirect(url_for('clusterwebpage'))
        else:
            # Step 2: Process the selected options
            try:
                # Verify we have a file ID in the session
                file_id = request.form.get('file_id') or session.get('file_id')
                if not file_id:
                    flash("No file selected. Please upload a file first.", "error")
                    return redirect(url_for('clusterwebpage'))
                    
                # Get file record from database
                file_record = UploadedFile.get(int(file_id))
                if not file_record:
                    flash("File not found. Please upload a file again.", "error")
                    return redirect(url_for('clusterwebpage'))
                    
                # Security check: ensure the current user owns this file
                if file_record.user_id != current_user.id and not current_user.is_admin():
                    flash("You don't have permission to access this file.", "error")
                    return redirect(url_for('dashboard'))
                
                # Get the selected options
                bacts = request.form.get('selected_values', '')
                conc = request.form.get('conc_typeX')
                vol = request.form.get('vol_typeX')
                slide = request.form.get('slide_typeX')
                
                # Validate required options
                if not bacts or not conc or not vol or not slide:
                    flash("Please select all required options.", "error")
                    return redirect(url_for('clusterwebpage'))
                
                # Get clustering algorithm option
                clusteringOption = request.form.get('clusteringOption') or session.get('clusteringOption')
                if not clusteringOption:
                    flash("Clustering algorithm not specified.", "error")
                    return redirect(url_for('clusterwebpage'))
                
                # Read the file and process it
                df = pd.read_csv(file_record.file_path)
                
                # Convert string of bacteria names to list if needed
                if isinstance(bacts, str) and ',' in bacts:
                    bact_list = bacts.split(',')
                else:
                    bact_list = bacts
                
                # Process the file and generate cluster plots
                result = clustering.Clustered_data(df, clusteringOption, bact_list, conc, vol, slide)
                
                # Prepare to save the cluster plot
                base_filename = os.path.splitext(os.path.basename(file_record.original_filename))[0]
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                
                # Create plot directory if it doesn't exist
                cluster_plot_filename = f"{base_filename}_{timestamp}_cluster.png"
                cluster_plot_path = get_plot_path('cluster_plots', cluster_plot_filename, current_user.id)
                
                # Save the plot image
                try:
                    save_plot_image(result, cluster_plot_path)
                except Exception as e:
                    logging.error(f"Error saving cluster plot: {str(e)}")
                
                # Update file record with processing information
                file_record.processed = True
                file_record.processed_at = datetime.utcnow()
                file_record.processing_type = f'clustering_{clusteringOption}'
                file_record.cluster_plot_path = cluster_plot_path
                file_record.update()
                
                # Log the successful processing
                logging.info(f"Clustering analysis completed for file {file_record.id} by user {current_user.id}")
                
                # Get user's recent clustering files
                user_files = UploadedFile.get_user_files(current_user.id)
                cluster_files = [f for f in user_files if 'cluster' in f.processing_type][:5]
                
                flash("Clustering analysis completed successfully!", "success")
                return render_template('clusterwebpage.html', 
                                      sample_placeholder="Clustering analysis completed successfully!", 
                                      img_base64=result,
                                      user_files=cluster_files)
            except Exception as e:
                # Log the error and notify the user
                logging.error(f"Error during clustering analysis: {str(e)}")
                flash(f"Error during analysis: {str(e)}", "error")
                return redirect(url_for('clusterwebpage'))
    else:
        return redirect(url_for('index'))

@app.route("/datavisualization", methods=['POST'])
@login_required
def datavisualization():
    if request.method == 'POST':
        visit = request.form.get("visit")
        
        if visit == "one":
            try:
                # Check if file was uploaded
                if 'csv_file' not in request.files:
                    flash("No file part", "error")
                    return redirect(url_for('visual'))
                    
                folder_path = request.files['csv_file']
                if folder_path.filename == '':
                    flash("No selected file", "error")
                    return redirect(url_for('visual'))
                
                # Check file extension
                if not folder_path.filename.lower().endswith('.csv'):
                    flash("Only CSV files are allowed", "error")
                    return redirect(url_for('visual'))
                
                # Get data type selection
                data_type = request.form.get("data_type")
                if not data_type:
                    flash("Please select a data type", "error")
                    return redirect(url_for('visual'))
                
                # Create user-specific filename with UUID
                original_filename = folder_path.filename
                unique_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
                file_path = get_user_file_path(unique_filename)
                
                # Save the uploaded file
                folder_path.save(file_path)
                
                # Create file record in database
                file_record = add_file_record(file_path, original_filename, folder_path.content_type)
                if not file_record:
                    flash("Error creating file record", "error")
                    return redirect(url_for('visual'))
                
                # Store the file ID and data type in the session
                session['file_id'] = file_record.id
                session['data_type'] = data_type
                
                # Log the successful upload
                logging.info(f"User {current_user.id} uploaded file {file_record.id} for visualization")
                
                # Read CSV file to extract unique values
                dfX = pd.read_csv(file_path)
                
                if data_type == 'cluster_data':
                    unq_bacts = dfX['bacteria'].unique() if 'bacteria' in dfX.columns else []
                    strings = unq_bacts.tolist()
                    return render_template('visual.html', item=strings)
                else:
                    # For raw data
                    unq_bacts = dfX['bacteria'].unique() if 'bacteria' in dfX.columns else []
                    unq_concs = dfX['concentration'].unique() if 'concentration' in dfX.columns else []
                    unq_vols = dfX['volume'].unique() if 'volume' in dfX.columns else []
                    unq_tri = dfX['trail'].unique() if 'trail' in dfX.columns else []
                    
                    strings = unq_bacts.tolist()
                    unq_concs = unq_concs.tolist()
                    unq_vols = unq_vols.tolist()
                    unq_tri = unq_tri.tolist()
                    
                    return render_template('visual.html', 
                                          item=strings, 
                                          item2=unq_concs, 
                                          item3=unq_vols, 
                                          item4=["abc"], 
                                          item5=unq_tri)
            
            except Exception as e:
                # If there's an error, log it and notify the user
                logging.error(f"Error processing uploaded file: {str(e)}")
                flash(f"Error processing file: {str(e)}", "error")
                return redirect(url_for('visual'))
        else:
            # Step 2: Process the selected options
            try:
                # Get file ID from form or session
                file_id = request.form.get('file_id') or session.get('file_id')
                if not file_id:
                    flash("No file selected. Please upload a file first.", "error")
                    return redirect(url_for('visual'))
                    
                # Get file record from database
                file_record = UploadedFile.get(int(file_id))
                if not file_record:
                    flash("File not found. Please upload a file again.", "error")
                    return redirect(url_for('visual'))
                    
                # Security check: ensure the current user owns this file
                if file_record.user_id != current_user.id and not current_user.is_admin():
                    flash("You don't have permission to access this file.", "error")
                    return redirect(url_for('dashboard'))
                
                d_type = session.get('data_type')
                
                # Read the file and process it
                df = pd.read_csv(file_record.file_path)
                
                if d_type == 'cluster_data':
                    bacts = request.form.getlist('options[]')
                    
                    # Validate selections
                    if not bacts:
                        flash("Please select at least one bacteria type.", "error")
                        return redirect(url_for('visual'))
                    
                    img_base64 = VSUL.clustered_plotter(df, bacts)
                    
                    # Get plot paths for database record
                    base_filename = os.path.splitext(os.path.basename(file_record.original_filename))[0]
                    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                    
                    # Save plot to file system
                    cluster_plot_filename = f"{base_filename}_{timestamp}_cluster_visual.png"
                    cluster_plot_path = get_plot_path('cluster_plots', cluster_plot_filename, current_user.id)
                    
                    # Save the plot image
                    try:
                        save_plot_image(img_base64, cluster_plot_path)
                    except Exception as e:
                        logging.error(f"Error saving cluster visualization plot: {str(e)}")
                    
                    # Update file record with processing information
                    file_record.processed = True
                    file_record.processed_at = datetime.utcnow()
                    file_record.processing_type = 'cluster_visual'
                    file_record.cluster_plot_path = cluster_plot_path
                    file_record.update()
                    
                    # Get user's recent visualization files
                    user_files = UploadedFile.get_user_files(current_user.id)
                    visual_files = [f for f in user_files if 'visual' in f.processing_type][:5]
                    
                    # Log the successful processing
                    logging.info(f"Cluster visualization completed for file {file_record.id} by user {current_user.id}")
                    
                    flash("Cluster visualization completed successfully!", "success")
                    return render_template('visual.html', 
                                          img_base64=img_base64,
                                          user_files=visual_files)
                else:
                    # For raw data
                    bacts = request.form.getlist('options[]')
                    conc = request.form.get('conc_typeX')
                    vol = request.form.get('vol_typeX')
                    slide = request.form.get('slide_typeX')
                    trails = request.form.getlist('options2[]')
                    plot_type = request.form.get('plot_type')
                    
                    # Validate selections
                    if not bacts:
                        flash("Please select at least one bacteria type.", "error")
                        return redirect(url_for('visual'))
                    
                    if not conc:
                        flash("Please select a concentration option.", "error")
                        return redirect(url_for('visual'))
                    
                    if not vol:
                        flash("Please select a volume option.", "error")
                        return redirect(url_for('visual'))
                    
                    if not slide:
                        flash("Please select a slide option.", "error")
                        return redirect(url_for('visual'))
                    
                    if not plot_type:
                        flash("Please select a plot type (Individual or Merged).", "error")
                        return redirect(url_for('visual'))
                    
                    # Process the data and generate visualization
                    img_base64 = VSUL.visual(df, bacts, d_type, plot_type, conc, vol, slide, trails)
                    
                    # Get plot paths for database record
                    base_filename = os.path.splitext(os.path.basename(file_record.original_filename))[0]
                    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                    
                    # Save plot to file system
                    raw_plot_filename = f"{base_filename}_{timestamp}_raw_visual.png"
                    raw_plot_path = get_plot_path('raw_plots', raw_plot_filename, current_user.id)
                    
                    # Save the plot image
                    try:
                        save_plot_image(img_base64, raw_plot_path)
                    except Exception as e:
                        logging.error(f"Error saving raw visualization plot: {str(e)}")
                    
                    # Update file record with processing information
                    file_record.processed = True
                    file_record.processed_at = datetime.utcnow()
                    file_record.processing_type = 'raw_visual'
                    file_record.raw_plot_path = raw_plot_path
                    file_record.update()
                    
                    # Get user's recent visualization files
                    user_files = UploadedFile.get_user_files(current_user.id)
                    visual_files = [f for f in user_files if 'visual' in f.processing_type][:5]
                    
                    # Log the successful processing
                    logging.info(f"Raw data visualization completed for file {file_record.id} by user {current_user.id}")
                    
                    flash("Raw data visualization completed successfully!", "success")
                    return render_template('visual.html', 
                                          img_base64=img_base64,
                                          user_files=visual_files)
            
            except Exception as e:
                # Log the error and notify the user
                logging.error(f"Error during visualization: {str(e)}")
                flash(f"Error during visualization: {str(e)}", "error")
                return redirect(url_for('visual'))
    else:
        return redirect(url_for('index'))

@app.route('/dataML', methods=['POST'])
@login_required
def dataML():
    if request.method == 'POST':
        ML_type = request.form.get('radioOption')
        
        if ML_type == "ML_train":
            visit = request.form.get("visit")
            
            if visit == "one":
                try:
                    # Check if file was uploaded
                    if 'csv_file' not in request.files:
                        flash("No file part", "error")
                        return redirect(url_for('ML_webpage'))
                        
                    folder_path = request.files['csv_file']
                    if folder_path.filename == '':
                        flash("No selected file", "error")
                        return redirect(url_for('ML_webpage'))
                    
                    # Check file extension
                    if not folder_path.filename.lower().endswith('.csv'):
                        flash("Only CSV files are allowed", "error")
                        return redirect(url_for('ML_webpage'))
                    
                    # Get ML training type
                    data_type = request.form.get("Train_ML")
                    if not data_type:
                        flash("Please select a training data type", "error")
                        return redirect(url_for('ML_webpage'))
                    
                    # Create user-specific filename with UUID
                    original_filename = folder_path.filename
                    unique_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
                    file_path = get_user_file_path(unique_filename)
                    
                    # Save the uploaded file
                    folder_path.save(file_path)
                    
                    # Create file record in database
                    file_record = add_file_record(file_path, original_filename, folder_path.content_type)
                    if not file_record:
                        flash("Error creating file record", "error")
                        return redirect(url_for('ML_webpage'))
                    
                    # Store file ID and data type in session
                    session['file_id'] = file_record.id
                    session['data_type'] = data_type
                    
                    # Log the successful upload
                    logging.info(f"User {current_user.id} uploaded file {file_record.id} for ML training")
                    
                    # Read CSV file to extract unique values
                    dfX = pd.read_csv(file_path)
                    unq_bacts = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                    
                    # Get user's recent ML files
                    user_files = UploadedFile.get_user_files(current_user.id)
                    ml_files = [f for f in user_files if 'ml' in f.processing_type][:5]
                    
                    return render_template('ML_webpage.html', 
                                          item=unq_bacts,
                                          user_files=ml_files)
                except Exception as e:
                    # If there's an error, log it and notify the user
                    logging.error(f"Error processing uploaded file: {str(e)}")
                    flash(f"Error processing file: {str(e)}", "error")
                    return redirect(url_for('ML_webpage'))
            else:
                try:
                    # Get file ID from form or session
                    file_id = request.form.get('file_id') or session.get('file_id')
                    if not file_id:
                        flash("No file selected. Please upload a file first.", "error")
                        return redirect(url_for('ML_webpage'))
                        
                    # Get file record from database
                    file_record = UploadedFile.get(int(file_id))
                    if not file_record:
                        flash("File not found. Please upload a file again.", "error")
                        return redirect(url_for('ML_webpage'))
                        
                    # Security check: ensure the current user owns this file
                    if file_record.user_id != current_user.id and not current_user.is_admin():
                        flash("You don't have permission to access this file.", "error")
                        return redirect(url_for('dashboard'))
                    
                    d_type = session.get('data_type')
                    
                    # Get selected bacteria types
                    bacts = request.form.getlist('options')
                    
                    # Validate selections
                    if not bacts:
                        flash("Please select at least one bacteria type.", "error")
                        return redirect(url_for('ML_webpage'))
                    
                    # Read the file and process it
                    df = pd.read_csv(file_record.file_path)
                    img_base64 = ML.Training_ML(df, d_type, bacts)
                    
                    # Get plot paths for database record
                    base_filename = os.path.splitext(os.path.basename(file_record.original_filename))[0]
                    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                    
                    # Save plot to file system
                    ml_plot_filename = f"{base_filename}_{timestamp}_ml_training.png"
                    ml_plot_path = get_plot_path('ML_Plots', ml_plot_filename, current_user.id)
                    
                    # Save the plot image
                    try:
                        save_plot_image(img_base64, ml_plot_path)
                    except Exception as e:
                        logging.error(f"Error saving ML training plot: {str(e)}")
                    
                    # Update file record with processing information
                    file_record.processed = True
                    file_record.processed_at = datetime.utcnow()
                    file_record.processing_type = 'ml_training'
                    file_record.ml_plot_path = ml_plot_path
                    file_record.update()
                    
                    # Get user's recent ML files
                    user_files = UploadedFile.get_user_files(current_user.id)
                    ml_files = [f for f in user_files if 'ml' in f.processing_type][:5]
                    
                    # Log the successful processing
                    logging.info(f"ML training completed for file {file_record.id} by user {current_user.id}")
                    
                    flash("ML training completed successfully!", "success")
                    return render_template('ML_webpage.html', 
                                          img_base64=img_base64,
                                          user_files=ml_files)
                except Exception as e:
                    # Log the error and notify the user
                    logging.error(f"Error during ML training: {str(e)}")
                    flash(f"Error during training: {str(e)}", "error")
                    return redirect(url_for('ML_webpage'))
        else:
            # ML testing
            try:
                # Check if a file was uploaded
                if 'csv_file' not in request.files:
                    flash("No file part", "error")
                    return redirect(url_for('ML_webpage'))
                    
                data = request.files['csv_file']
                if data.filename == '':
                    flash("No selected file", "error")
                    return redirect(url_for('ML_webpage'))
                
                # Check file extension
                if not data.filename.lower().endswith('.csv'):
                    flash("Only CSV files are allowed", "error")
                    return redirect(url_for('ML_webpage'))
                
                # Create user-specific filename with UUID
                original_filename = data.filename
                unique_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
                file_path = get_user_file_path(unique_filename)
                
                # Save the uploaded file
                data.save(file_path)
                
                # Create file record in database
                file_record = add_file_record(file_path, original_filename, data.content_type)
                if not file_record:
                    flash("Error creating file record", "error")
                    return redirect(url_for('ML_webpage'))
                
                # Process the file
                df = pd.read_csv(file_path)
                img_base64 = ML.testing(df)
                
                # Get plot paths for database record
                base_filename = os.path.splitext(os.path.basename(file_record.original_filename))[0]
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                
                # Save plot to file system
                ml_plot_filename = f"{base_filename}_{timestamp}_ml_testing.png"
                ml_plot_path = get_plot_path('ML_Plots', ml_plot_filename, current_user.id)
                
                # Save the plot image
                try:
                    save_plot_image(img_base64, ml_plot_path)
                except Exception as e:
                    logging.error(f"Error saving ML testing plot: {str(e)}")
                
                # Update file record with processing information
                file_record.processed = True
                file_record.processed_at = datetime.utcnow()
                file_record.processing_type = 'ml_testing'
                file_record.ml_plot_path = ml_plot_path
                file_record.update()
                
                # Get user's recent ML files
                user_files = UploadedFile.get_user_files(current_user.id)
                ml_files = [f for f in user_files if 'ml' in f.processing_type][:5]
                
                # Log the successful processing
                logging.info(f"ML testing completed for file {file_record.id} by user {current_user.id}")
                
                flash("ML testing completed successfully!", "success")
                return render_template('ML_webpage.html', 
                                      img_base64=img_base64,
                                      user_files=ml_files)
            except Exception as e:
                # Log the error and notify the user
                logging.error(f"Error during ML testing: {str(e)}")
                flash(f"Error during testing: {str(e)}", "error")
                return redirect(url_for('ML_webpage'))
    else:
        return redirect(url_for('ML_webpage'))

@app.route('/rawData', methods=['POST'])
@login_required
def rawData():
    if request.method == 'POST':
        try:
            # Check if file was uploaded
            if 'csv_file' not in request.files:
                flash("No file part", "error")
                return redirect(url_for('rd'))
                
            csv_file = request.files['csv_file']
            if csv_file.filename == '':
                flash("No selected file", "error")
                return redirect(url_for('rd'))
            
            # Check file extension
            if not csv_file.filename.lower().endswith('.csv'):
                flash("Only CSV files are allowed", "error")
                return redirect(url_for('rd'))
            
            # Process the file directly
            try:
                # Read CSV into dataframe
                df = pd.read_csv(csv_file)
                # Process the dataframe
                result = raw_data_calc(df)
                
                # Return the result
                return render_template('Raw_signalBreathe.HTML', 
                                      img_base64=result,
                                      sample_placeholder="Analysis completed successfully!")
            except Exception as e:
                logging.error(f"Error processing CSV data: {str(e)}")
                flash(f"Error processing CSV data: {str(e)}", "error")
                return redirect(url_for('rd'))
        except Exception as e:
            logging.error(f"Error in raw data analysis: {str(e)}")
            flash(f"Error processing file: {str(e)}", "error")
            return redirect(url_for('rd'))
    else:
        return redirect(url_for('breathepage'))

@app.route('/spiro', methods=['POST'])
def spiro():
    """Process uploaded file for breathing analysis based on selected type without saving to database"""
    if request.method == 'POST':
        try:
            # Check if file was uploaded
            if 'csv_file' not in request.files:
                flash("No file part", "error")
                return redirect(url_for('breathepage'))
                
            csv_file = request.files['csv_file']
            if csv_file.filename == '':
                flash("No selected file", "error")
                return redirect(url_for('breathepage'))
            
            # Check file extension
            if not csv_file.filename.lower().endswith('.csv'):
                flash("Only CSV files are allowed", "error")
                return redirect(url_for('breathepage'))
            
            # Get the analysis type
            analysis_type = request.form.get('radioOption')
            if not analysis_type:
                flash("Please select an analysis type", "error")
                return redirect(url_for('breathepage'))
                
            logging.info(f"Processing file: {csv_file.filename} for analysis type: {analysis_type}")
            
            # Read CSV directly without saving to disk or database
            df = pd.read_csv(csv_file)
            
            # Process based on analysis type
            if analysis_type == "svc":
                try:
                    # Process the dataframe using svc_calc function
                    result = svc_calc(df)
                    
                    # Return the result directly without storing in database
                    flash("SVC analysis completed successfully!", "success")
                    return render_template('SVC.HTML', img_base64=result)
                    
                except Exception as e:
                    logging.error(f"Error during SVC analysis: {str(e)}")
                    flash(f"Error during SVC analysis: {str(e)}", "error")
                    return redirect(url_for('svc'))
                    
            elif analysis_type == "fvc":
                try:
                    # Process the dataframe using fvc_calc function
                    plot1, plot2 = fvc_calc(df)
                    
                    # Return the result directly without storing in database
                    flash("FVC analysis completed successfully!", "success")
                    return render_template('FVC.HTML', img_base64_1=plot1, img_base64_2=plot2)
                    
                except Exception as e:
                    logging.error(f"Error during FVC analysis: {str(e)}")
                    flash(f"Error during FVC analysis: {str(e)}", "error")
                    return redirect(url_for('fvc'))
                    
            elif analysis_type == "mvv":
                try:
                    # Process the dataframe using mvv_calc function
                    result = mvv_calc(df)
                    
                    # Return the result directly without storing in database
                    flash("MVV analysis completed successfully!", "success")
                    return render_template('MVV.HTML', img_base64=result)
                    
                except Exception as e:
                    logging.error(f"Error during MVV analysis: {str(e)}")
                    flash(f"Error during MVV analysis: {str(e)}", "error")
                    return redirect(url_for('mvv'))
                    
            elif analysis_type == "sm":
                try:
                    # Calculate standard metrics
                    resp_rate, LungCapacity, Quality = sm_calc(df)
                    
                    flash("Standard Metrics analysis completed successfully!", "success")
                    return render_template('Standard_Metrics_Breathe.HTML', 
                                         resp_rate=resp_rate, 
                                         LungCapacity=LungCapacity, 
                                         Quality=Quality,
                                         sample_placeholder="Analysis completed successfully!")
                                         
                except Exception as e:
                    logging.error(f"Error in SM calculation: {str(e)}")
                    flash(f"Error calculating metrics: {str(e)}", "error")
                    return redirect(url_for('sm'))
                    
            else:
                flash("Unknown analysis type", "error")
                return redirect(url_for('breathepage'))
                
        except Exception as e:
            # Log and display the error
            import traceback
            logging.error(f"Error in spiro route: {str(e)}")
            logging.error(traceback.format_exc())
            flash(f"Error processing file: {str(e)}", "error")
            return redirect(url_for('breathepage'))
    else:
        return redirect(url_for('breathepage'))

# === New Analysis Processing Routes ===

def process_onedrive_file_for_visualization():
    """Process a file selected from OneDrive for raw data visualization - FIXED for large files"""
    try:
        # Get OneDrive file ID from form
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            logging.error("No OneDrive file ID provided")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details first (without downloading content)
        file_name = None
        file_size = 0
        download_url = None
        
        logging.info(f"🔍 Getting file info for OneDrive file ID: {file_id}")
        
        # Try to get file info and download URL
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    
                    logging.info(f"📁 File: {file_name}, Size: {file_size} bytes ({file_size/1024/1024:.1f} MB)")
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        logging.info(f"✅ Got download URL successfully")
                        break
                    else:
                        logging.warning("⚠️ No download URL found in file info")
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            logging.error("❌ Could not get file info or download URL")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # For large files (>50MB), use streaming download to temporary file
        if file_size > 50 * 1024 * 1024:  # 50MB threshold
            logging.info(f"📥 Large file detected ({file_size/1024/1024:.1f} MB), using streaming download")
            
            # Create a temporary file with proper cleanup
            import tempfile
            import os
            
            try:
                # Create temporary file
                temp_fd, temp_file_path = tempfile.mkstemp(suffix='.csv', prefix='onedrive_large_')
                
                # Download file in chunks using streaming
                logging.info(f"🔽 Starting streaming download...")
                
                with requests.get(download_url, stream=True) as download_response:
                    download_response.raise_for_status()
                    
                    total_size = int(download_response.headers.get('content-length', 0))
                    downloaded_size = 0
                    chunk_size = 8192  # 8KB chunks
                    
                    with os.fdopen(temp_fd, 'wb') as temp_file:
                        for chunk in download_response.iter_content(chunk_size=chunk_size):
                            if chunk:
                                temp_file.write(chunk)
                                downloaded_size += len(chunk)
                                
                                # Log progress for very large files
                                if downloaded_size % (10 * 1024 * 1024) == 0:  # Every 10MB
                                    progress = (downloaded_size / total_size * 100) if total_size > 0 else 0
                                    logging.info(f"📊 Download progress: {progress:.1f}% ({downloaded_size/1024/1024:.1f} MB)")
                
                logging.info(f"✅ Download completed: {downloaded_size/1024/1024:.1f} MB")
                
                # Now process the file using pandas with chunking for very large files
                try:
                    if file_size > 100 * 1024 * 1024:  # 100MB+ files
                        logging.info(f"🔄 Processing very large file with chunking...")
                        
                        # Read file in chunks to get unique values
                        chunk_size_rows = 10000
                        unique_bacteria = set()
                        unique_concs = set()
                        unique_vols = set()
                        unique_slides = set()
                        unique_trials = set()
                        
                        chunk_count = 0
                        for chunk in pd.read_csv(temp_file_path, chunksize=chunk_size_rows):
                            chunk_count += 1
                            
                            # Extract unique values from this chunk
                            if 'bacteria' in chunk.columns:
                                unique_bacteria.update(chunk['bacteria'].dropna().unique())
                            if 'concentration' in chunk.columns:
                                unique_concs.update(chunk['concentration'].dropna().unique())
                            if 'volume' in chunk.columns:
                                unique_vols.update(chunk['volume'].dropna().unique())
                            if 'slide' in chunk.columns:
                                unique_slides.update(chunk['slide'].dropna().unique())
                            if 'trial' in chunk.columns:
                                unique_trials.update(chunk['trial'].dropna().unique())
                            
                            # Log progress
                            if chunk_count % 10 == 0:
                                logging.info(f"📊 Processed {chunk_count} chunks, found {len(unique_bacteria)} bacteria types")
                        
                        # Convert sets to sorted lists
                        strings = sorted(list(unique_bacteria))
                        unq_concs = sorted(list(unique_concs))
                        unq_vols = sorted(list(unique_vols))
                        unq_sli = sorted(list(unique_slides))
                        unq_tri = sorted(list(unique_trials))
                        
                        logging.info(f"✅ Chunked processing complete: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                        
                    else:
                        # Regular processing for smaller large files
                        logging.info(f"🔄 Processing large file normally...")
                        dfX = pd.read_csv(temp_file_path)
                        logging.info(f"📊 Successfully read CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                        
                        # Extract unique values
                        strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                        unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                        unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                        unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                        unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                        
                        logging.info(f"✅ Extracted unique values: {len(strings)} bacteria types")
                    
                    # Store the file path in session for step 2
                    session['csv_file_path'] = temp_file_path
                    session['data_type'] = request.form.get("data_type", "raw_data")
                    session['is_large_file'] = True
                    session['file_size_mb'] = file_size / 1024 / 1024
                    session.modified = True
                    
                    logging.info(f"💾 Stored file path in session: {temp_file_path}")
                    
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                          item=strings, 
                                          item2=unq_concs, 
                                          item3=unq_vols, 
                                          item4=unq_sli, 
                                          item5=unq_tri,
                                          file_size_info=f"Large file: {file_size/1024/1024:.1f} MB")
                                          
                except Exception as e:
                    logging.error(f"❌ Error processing large file: {str(e)}")
                    # Clean up temp file on error
                    try:
                        os.unlink(temp_file_path)
                    except:
                        pass
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder=f"Error processing large file: {str(e)}")
                        
            except Exception as e:
                logging.error(f"❌ Error downloading large file: {str(e)}")
                return render_template('current_analysis_rawDatavisual.HTML', 
                                     sample_placeholder=f"Error downloading large file: {str(e)}")
        
        else:
            # Standard processing for smaller files (<50MB)
            logging.info(f"📥 Standard download for file ({file_size/1024/1024:.1f} MB)")
            
            try:
                # Download the file content
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    file_content = download_response.content
                    logging.info(f"✅ Downloaded {len(file_content)} bytes")
                else:
                    logging.error(f"❌ Download failed: {download_response.status_code}")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: Could not download file from OneDrive")
                
                # Create temporary file
                import tempfile
                temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
                temp_file.write(file_content)
                temp_file.close()
                
                # Read CSV file to extract unique values
                dfX = pd.read_csv(temp_file.name)
                logging.info(f"📊 Successfully read CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                
                # Store the file path in session for step 2
                session['csv_file_path'] = temp_file.name
                session['data_type'] = request.form.get("data_type", "raw_data")
                session['is_large_file'] = False
                session.modified = True
                
                # Extract unique values
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                
                logging.info(f"✅ Extracted unique values: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                
                return render_template('current_analysis_rawDatavisual.HTML', 
                                      item=strings, 
                                      item2=unq_concs, 
                                      item3=unq_vols, 
                                      item4=unq_sli, 
                                      item5=unq_tri)
                                      
            except Exception as e:
                logging.error(f"❌ Error processing standard file: {str(e)}")
                return render_template('current_analysis_rawDatavisual.HTML', 
                                     sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive file processing: {str(e)}")
        return render_template('current_analysis_rawDatavisual.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")

@app.route("/NewVisualization", methods=['GET', 'POST'])
def NewVisualization():
    """COMPLETE FIXED NewVisualization route with proper OneDrive large file support"""
    if request.method == 'POST':
        visit = request.form.get("visit")
        
        # Enhanced logging for debugging
        logging.info(f"🔍 NewVisualization called with visit={visit}")
        logging.info(f"📝 Form data: {dict(request.form)}")
        logging.info(f"📁 Files: {list(request.files.keys())}")
        
        if visit == "one":
            try:
                # Check if it's a OneDrive file or local upload
                file_source = request.form.get('fileSource', 'local')
                logging.info(f"📂 File source: {file_source}")
                
                if file_source == 'cloud' or file_source == 'onedrive':
                    # Handle OneDrive file processing
                    logging.info(f"☁️ Processing OneDrive file...")
                    return process_onedrive_file_for_visualization()
                else:
                    # Handle local file upload
                    if 'csv_file' not in request.files:
                        logging.warning("⚠️ No file part in request")
                        return render_template('current_analysis_rawDatavisual.HTML', 
                                             sample_placeholder="Error: No file uploaded")
                        
                    folder_path = request.files['csv_file']
                    if folder_path.filename == '':
                        logging.warning("⚠️ No selected file")
                        return render_template('current_analysis_rawDatavisual.HTML', 
                                             sample_placeholder="Error: No file selected")
                    
                    # Check file extension
                    if not folder_path.filename.lower().endswith('.csv'):
                        logging.warning(f"⚠️ Invalid file type: {folder_path.filename}")
                        return render_template('current_analysis_rawDatavisual.HTML', 
                                             sample_placeholder="Error: Only CSV files are allowed")
                    
                    # Get data type selection
                    data_type = request.form.get("data_type")
                    if not data_type:
                        logging.warning("⚠️ No data type selected")
                        return render_template('current_analysis_rawDatavisual.HTML', 
                                             sample_placeholder="Error: Please select a data type")
                    
                    # Save the uploaded file
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    safe_filename = f"{timestamp}_{secure_filename(folder_path.filename)}"
                    csv_filename = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
                    folder_path.save(csv_filename)
                    logging.info(f"💾 File saved to: {csv_filename}")
                    
                    # Read CSV file to extract unique values
                    try:
                        dfX = pd.read_csv(csv_filename)
                        logging.info(f"📊 Successfully read CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                    except Exception as e:
                        logging.error(f"❌ Error reading CSV: {str(e)}")
                        return render_template('current_analysis_rawDatavisual.HTML', 
                                             sample_placeholder=f"Error reading CSV file: {str(e)}")
                    
                    # Store the file path and data type in session
                    session['csv_file_path'] = csv_filename
                    session['data_type'] = data_type
                    session['is_large_file'] = False
                    session.modified = True
                    
                    # Extract unique values from different columns
                    strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                    unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                    unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                    unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                    unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                    
                    logging.info(f"✅ Extracted unique values - bacteria: {len(strings)}, concs: {len(unq_concs)}")
                    
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                          item=strings, 
                                          item2=unq_concs, 
                                          item3=unq_vols, 
                                          item4=unq_sli, 
                                          item5=unq_tri)
            except Exception as e:
                logging.error(f"❌ Error processing uploaded file: {str(e)}")
                return render_template('current_analysis_rawDatavisual.HTML', 
                                     sample_placeholder=f"Error processing file: {str(e)}")
        else:
            # Step 2: Process the selected options with large file support
            try:
                # Get the CSV path from session
                csv_file_path = session.get('csv_file_path')
                d_type = session.get('data_type')
                is_large_file = session.get('is_large_file', False)
                file_size_mb = session.get('file_size_mb', 0)
                
                logging.info(f"🔄 Step 2 - Processing file: {csv_file_path}")
                logging.info(f"📊 File info - Type: {d_type}, Large: {is_large_file}, Size: {file_size_mb:.1f}MB")
                
                if not csv_file_path:
                    logging.error("❌ No file path in session")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: No file found. Please upload a file first.")
                
                if not os.path.exists(csv_file_path):
                    logging.error(f"❌ File not found on disk: {csv_file_path}")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: File not found. Please upload the file again.")
                
                # Get selected options from form
                bacts = request.form.getlist('options[]')
                conc = request.form.get('conc_typeX')
                vol = request.form.getlist('vol_typeX[]')
                slide = request.form.getlist('options3[]')
                trails = request.form.getlist('options2[]')
                
                logging.info(f"📋 Selected options - bacteria: {len(bacts)} ({bacts}), conc: {conc}, vol: {len(vol)} ({vol})")
                logging.info(f"📋 Selected slide: {len(slide)} ({slide}), trials: {len(trails)} ({trails})")
                
                # Validate selections
                if not bacts:
                    logging.warning("⚠️ No bacteria selected")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: Please select at least one bacteria type.")
                
                # Read and process the file with memory optimization for large files
                try:
                    if is_large_file and file_size_mb > 100:
                        logging.info(f"🔄 Processing very large file with memory optimization...")
                        
                        # For very large files, use chunked processing
                        # Read file in chunks and filter only the data we need
                        chunk_size_rows = 5000
                        filtered_chunks = []
                        total_rows_processed = 0
                        chunk_count = 0
                        
                        for chunk in pd.read_csv(csv_file_path, chunksize=chunk_size_rows):
                            chunk_count += 1
                            original_chunk_size = len(chunk)
                            total_rows_processed += original_chunk_size
                            
                            # Debug: Show sample values from the first chunk
                            if chunk_count == 1:
                                logging.info(f"🔍 DEBUG - Sample data from first chunk:")
                                if 'bacteria' in chunk.columns:
                                    sample_bacteria = chunk['bacteria'].unique()[:5]
                                    logging.info(f"🔍 Sample bacteria values: {sample_bacteria}")
                                if 'concentration' in chunk.columns:
                                    sample_conc = chunk['concentration'].unique()[:5]
                                    logging.info(f"🔍 Sample concentration values: {sample_conc}")
                                if 'volume' in chunk.columns:
                                    sample_vol = chunk['volume'].unique()[:5]
                                    logging.info(f"🔍 Sample volume values: {sample_vol}")
                                if 'slide' in chunk.columns:
                                    sample_slide = chunk['slide'].unique()[:5]
                                    logging.info(f"🔍 Sample slide values: {sample_slide}")
                                if 'trial' in chunk.columns:
                                    sample_trial = chunk['trial'].unique()[:5]
                                    logging.info(f"🔍 Sample trial values: {sample_trial}")
                                
                                logging.info(f"🔍 User selected - bacteria: {bacts}, conc: {conc}, vol: {vol}, slide: {slide}, trials: {trails}")
                            
                            # Filter chunk based on selected criteria - FIXED VERSION
                            # Filter by bacteria (required)
                            if 'bacteria' in chunk.columns and bacts:
                                before_bacteria = len(chunk)
                                chunk = chunk[chunk['bacteria'].isin(bacts)]
                                logging.info(f"📋 After bacteria filter ({bacts}): {len(chunk)} rows from {before_bacteria}")

                            # Filter by concentration (if specified)
                            if conc and 'concentration' in chunk.columns:
                                before_conc = len(chunk)
                                # Handle both string and numeric concentrations
                                try:
                                    # Try numeric comparison first
                                    conc_numeric = float(conc)
                                    chunk = chunk[chunk['concentration'] == conc_numeric]
                                except (ValueError, TypeError):
                                    # Fall back to string comparison
                                    chunk_conc_str = chunk['concentration'].astype(str)
                                    conc_str = str(conc)
                                    chunk = chunk[chunk_conc_str == conc_str]
                                logging.info(f"📋 After concentration filter ({conc}): {len(chunk)} rows from {before_conc}")

                            # Filter by volume (if specified)
                            if vol and 'volume' in chunk.columns:
                                before_vol = len(chunk)
                                # Handle both string and numeric volumes
                                try:
                                    vol_numeric = [float(v) for v in vol]
                                    chunk = chunk[chunk['volume'].isin(vol_numeric)]
                                except (ValueError, TypeError):
                                    # Fall back to string comparison
                                    chunk_vol_str = chunk['volume'].astype(str)
                                    vol_str = [str(v) for v in vol]
                                    chunk = chunk[chunk_vol_str.isin(vol_str)]
                                logging.info(f"📋 After volume filter ({vol}): {len(chunk)} rows from {before_vol}")

                            # Filter by slide (if specified)
                            if slide and 'slide' in chunk.columns:
                                before_slide = len(chunk)
                                chunk_slide_str = chunk['slide'].astype(str)
                                slide_str = [str(s) for s in slide]
                                chunk = chunk[chunk_slide_str.isin(slide_str)]
                                logging.info(f"📋 After slide filter ({slide}): {len(chunk)} rows from {before_slide}")

                            # Filter by trial (if specified)
                            if trails and 'trial' in chunk.columns:
                                before_trial = len(chunk)
                                chunk_trial_str = chunk['trial'].astype(str)
                                trails_str = [str(t) for t in trails]
                                chunk = chunk[chunk_trial_str.isin(trails_str)]
                                logging.info(f"📋 After trial filter ({trails}): {len(chunk)} rows from {before_trial}")
                            
                            # Add non-empty chunks to filtered list
                            if not chunk.empty:
                                filtered_chunks.append(chunk)
                                logging.info(f"✅ Chunk {chunk_count} retained: {len(chunk)} rows")
                            else:
                                if chunk_count <= 5:  # Only log for first few chunks to avoid spam
                                    logging.warning(f"⚠️ Chunk {chunk_count} empty after filtering")
                            
                            # Log progress every 50k rows
                            if total_rows_processed % 50000 == 0:
                                logging.info(f"📊 Processed {total_rows_processed} rows, {len(filtered_chunks)} chunks retained")
                            
                            # Memory management
                            if chunk_count % 50 == 0:
                                gc.collect()
                        
                        # Combine filtered chunks
                        if filtered_chunks:
                            df = pd.concat(filtered_chunks, ignore_index=True)
                            logging.info(f"✅ Combined filtered data: {len(df)} rows from {len(filtered_chunks)} chunks")
                        else:
                            logging.warning("⚠️ No data matches the selected criteria")
                            return render_template('current_analysis_rawDatavisual.HTML', 
                                                 sample_placeholder="No data matches your selection criteria. Please try different options or check your data values.")
                    else:
                        # Standard processing for smaller files
                        logging.info(f"🔄 Standard processing for file...")
                        df = pd.read_csv(csv_file_path)
                        logging.info(f"📊 Successfully read CSV for processing: {len(df)} rows")
                        
                        # Apply the same filtering logic for consistency
                        original_size = len(df)
                        
                        # Debug: Show sample values
                        logging.info(f"🔍 DEBUG - Sample data:")
                        if 'bacteria' in df.columns:
                            sample_bacteria = df['bacteria'].unique()[:5]
                            logging.info(f"🔍 Sample bacteria values: {sample_bacteria}")
                        if 'concentration' in df.columns:
                            sample_conc = df['concentration'].unique()[:5]
                            logging.info(f"🔍 Sample concentration values: {sample_conc}")
                        
                        logging.info(f"🔍 User selected - bacteria: {bacts}, conc: {conc}")
                        
                        # Apply filters
                        if 'bacteria' in df.columns and bacts:
                            df = df[df['bacteria'].isin(bacts)]
                            logging.info(f"📋 After bacteria filter: {len(df)} rows from {original_size}")
                        
                        if conc and 'concentration' in df.columns:
                            before_conc = len(df)
                            try:
                                conc_numeric = float(conc)
                                df = df[df['concentration'] == conc_numeric]
                            except (ValueError, TypeError):
                                df_conc_str = df['concentration'].astype(str)
                                conc_str = str(conc)
                                df = df[df_conc_str == conc_str]
                            logging.info(f"📋 After concentration filter: {len(df)} rows from {before_conc}")
                        
                        if vol and 'volume' in df.columns:
                            before_vol = len(df)
                            try:
                                vol_numeric = [float(v) for v in vol]
                                df = df[df['volume'].isin(vol_numeric)]
                            except (ValueError, TypeError):
                                df_vol_str = df['volume'].astype(str)
                                vol_str = [str(v) for v in vol]
                                df = df[df_vol_str.isin(vol_str)]
                            logging.info(f"📋 After volume filter: {len(df)} rows from {before_vol}")
                        
                        if slide and 'slide' in df.columns:
                            before_slide = len(df)
                            df_slide_str = df['slide'].astype(str)
                            slide_str = [str(s) for s in slide]
                            df = df[df_slide_str.isin(slide_str)]
                            logging.info(f"📋 After slide filter: {len(df)} rows from {before_slide}")
                        
                        if trails and 'trial' in df.columns:
                            before_trial = len(df)
                            df_trial_str = df['trial'].astype(str)
                            trails_str = [str(t) for t in trails]
                            df = df[df_trial_str.isin(trails_str)]
                            logging.info(f"📋 After trial filter: {len(df)} rows from {before_trial}")
                        
                        if df.empty:
                            logging.warning("⚠️ No data matches the selected criteria")
                            return render_template('current_analysis_rawDatavisual.HTML', 
                                                 sample_placeholder="No data matches your selection criteria. Please try different options or check your data values.")
                
                    # Generate the visualization
                    logging.info(f"🎨 Generating visualization with {len(df)} rows...")
                    img_base64 = main_function_visual(df, bacts, conc, vol, slide, trails)
                    logging.info("✅ Visualization generated successfully")
                    
                    # Clean up the session after processing
                    session.pop('csv_file_path', None)
                    session.pop('data_type', None)
                    session.pop('is_large_file', None)
                    session.pop('file_size_mb', None)
                    session.modified = True
                    
                    # Clean up the temporary file with retry logic
                    cleanup_attempts = 3
                    for attempt in range(cleanup_attempts):
                        try:
                            if os.path.exists(csv_file_path):
                                os.remove(csv_file_path)
                                logging.info(f"✅ Cleaned up temporary file: {csv_file_path}")
                                break
                        except Exception as e:
                            if attempt < cleanup_attempts - 1:
                                logging.warning(f"⚠️ Cleanup attempt {attempt + 1} failed: {str(e)}, retrying...")
                                time.sleep(0.5)
                            else:
                                logging.warning(f"⚠️ Could not delete temporary file after {cleanup_attempts} attempts: {str(e)}")
                    
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         img_base64=img_base64,
                                         sample_placeholder="Raw data visualization completed successfully!")
                                         
                except Exception as e:
                    logging.error(f"❌ Error generating visualization: {str(e)}")
                    import traceback
                    logging.error(f"❌ Full traceback: {traceback.format_exc()}")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder=f"Error creating visualization: {str(e)}")
                                         
            except Exception as e:
                logging.error(f"❌ Error during step 2 processing: {str(e)}")
                import traceback
                logging.error(f"❌ Full traceback: {traceback.format_exc()}")
                return render_template('current_analysis_rawDatavisual.HTML', 
                                     sample_placeholder=f"Error during visualization: {str(e)}")
    else:
        # GET request - show the initial form
        return render_template('current_analysis_rawDatavisual.HTML',
                             sample_placeholder="Please wait for processor response")


def process_onedrive_file_for_visualization():
    """Process a file selected from OneDrive for raw data visualization - FIXED for large files"""
    try:
        # Get OneDrive file ID from form
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            logging.error("No OneDrive file ID provided")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details first (without downloading content)
        file_name = None
        file_size = 0
        download_url = None
        
        logging.info(f"🔍 Getting file info for OneDrive file ID: {file_id}")
        
        # Try to get file info and download URL
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    
                    logging.info(f"📁 File: {file_name}, Size: {file_size} bytes ({file_size/1024/1024:.1f} MB)")
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        logging.info(f"✅ Got download URL successfully")
                        break
                    else:
                        logging.warning("⚠️ No download URL found in file info")
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            logging.error("❌ Could not get file info or download URL")
            return render_template('current_analysis_rawDatavisual.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # For large files (>50MB), use streaming download to temporary file
        if file_size > 50 * 1024 * 1024:  # 50MB threshold
            logging.info(f"📥 Large file detected ({file_size/1024/1024:.1f} MB), using streaming download")
            
            # Create a temporary file with proper cleanup
            import tempfile
            import os
            
            try:
                # Create temporary file
                temp_fd, temp_file_path = tempfile.mkstemp(suffix='.csv', prefix='onedrive_large_')
                
                # Download file in chunks using streaming
                logging.info(f"🔽 Starting streaming download...")
                
                with requests.get(download_url, stream=True) as download_response:
                    download_response.raise_for_status()
                    
                    total_size = int(download_response.headers.get('content-length', 0))
                    downloaded_size = 0
                    chunk_size = 8192  # 8KB chunks
                    
                    with os.fdopen(temp_fd, 'wb') as temp_file:
                        for chunk in download_response.iter_content(chunk_size=chunk_size):
                            if chunk:
                                temp_file.write(chunk)
                                downloaded_size += len(chunk)
                                
                                # Log progress for very large files
                                if downloaded_size % (10 * 1024 * 1024) == 0:  # Every 10MB
                                    progress = (downloaded_size / total_size * 100) if total_size > 0 else 0
                                    logging.info(f"📊 Download progress: {progress:.1f}% ({downloaded_size/1024/1024:.1f} MB)")
                
                logging.info(f"✅ Download completed: {downloaded_size/1024/1024:.1f} MB")
                
                # Now process the file using pandas with chunking for very large files
                try:
                    if file_size > 100 * 1024 * 1024:  # 100MB+ files
                        logging.info(f"🔄 Processing very large file with chunking...")
                        
                        # Read file in chunks to get unique values
                        chunk_size_rows = 10000
                        unique_bacteria = set()
                        unique_concs = set()
                        unique_vols = set()
                        unique_slides = set()
                        unique_trials = set()
                        
                        chunk_count = 0
                        for chunk in pd.read_csv(temp_file_path, chunksize=chunk_size_rows):
                            chunk_count += 1
                            
                            # Extract unique values from this chunk
                            if 'bacteria' in chunk.columns:
                                unique_bacteria.update(chunk['bacteria'].dropna().unique())
                            if 'concentration' in chunk.columns:
                                unique_concs.update(chunk['concentration'].dropna().unique())
                            if 'volume' in chunk.columns:
                                unique_vols.update(chunk['volume'].dropna().unique())
                            if 'slide' in chunk.columns:
                                unique_slides.update(chunk['slide'].dropna().unique())
                            if 'trial' in chunk.columns:
                                unique_trials.update(chunk['trial'].dropna().unique())
                            
                            # Log progress
                            if chunk_count % 10 == 0:
                                logging.info(f"📊 Processed {chunk_count} chunks, found {len(unique_bacteria)} bacteria types")
                        
                        # Convert sets to sorted lists
                        strings = sorted(list(unique_bacteria))
                        unq_concs = sorted(list(unique_concs))
                        unq_vols = sorted(list(unique_vols))
                        unq_sli = sorted(list(unique_slides))
                        unq_tri = sorted(list(unique_trials))
                        
                        logging.info(f"✅ Chunked processing complete: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                        
                    else:
                        # Regular processing for smaller large files
                        logging.info(f"🔄 Processing large file normally...")
                        dfX = pd.read_csv(temp_file_path)
                        logging.info(f"📊 Successfully read CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                        
                        # Extract unique values
                        strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                        unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                        unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                        unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                        unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                        
                        logging.info(f"✅ Extracted unique values: {len(strings)} bacteria types")
                    
                    # Store the file path in session for step 2
                    session['csv_file_path'] = temp_file_path
                    session['data_type'] = request.form.get("data_type", "raw_data")
                    session['is_large_file'] = True
                    session['file_size_mb'] = file_size / 1024 / 1024
                    session.modified = True
                    
                    logging.info(f"💾 Stored file path in session: {temp_file_path}")
                    
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                          item=strings, 
                                          item2=unq_concs, 
                                          item3=unq_vols, 
                                          item4=unq_sli, 
                                          item5=unq_tri,
                                          file_size_info=f"Large file: {file_size/1024/1024:.1f} MB")
                                          
                except Exception as e:
                    logging.error(f"❌ Error processing large file: {str(e)}")
                    # Clean up temp file on error
                    try:
                        os.unlink(temp_file_path)
                    except:
                        pass
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder=f"Error processing large file: {str(e)}")
                        
            except Exception as e:
                logging.error(f"❌ Error downloading large file: {str(e)}")
                return render_template('current_analysis_rawDatavisual.HTML', 
                                     sample_placeholder=f"Error downloading large file: {str(e)}")
        
        else:
            # Standard processing for smaller files (<50MB)
            logging.info(f"📥 Standard download for file ({file_size/1024/1024:.1f} MB)")
            
            try:
                # Download the file content
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    file_content = download_response.content
                    logging.info(f"✅ Downloaded {len(file_content)} bytes")
                else:
                    logging.error(f"❌ Download failed: {download_response.status_code}")
                    return render_template('current_analysis_rawDatavisual.HTML', 
                                         sample_placeholder="Error: Could not download file from OneDrive")
                
                # Create temporary file
                import tempfile
                temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
                temp_file.write(file_content)
                temp_file.close()
                
                # Read CSV file to extract unique values
                dfX = pd.read_csv(temp_file.name)
                logging.info(f"📊 Successfully read CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                
                # Store the file path in session for step 2
                session['csv_file_path'] = temp_file.name
                session['data_type'] = request.form.get("data_type", "raw_data")
                session['is_large_file'] = False
                session.modified = True
                
                # Extract unique values
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                
                logging.info(f"✅ Extracted unique values: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                
                return render_template('current_analysis_rawDatavisual.HTML', 
                                      item=strings, 
                                      item2=unq_concs, 
                                      item3=unq_vols, 
                                      item4=unq_sli, 
                                      item5=unq_tri)
                                      
            except Exception as e:
                logging.error(f"❌ Error processing standard file: {str(e)}")
                return render_template('current_analysis_rawDatavisual.HTML', 
                                     sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive file processing: {str(e)}")
        return render_template('current_analysis_rawDatavisual.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")
    
def optimize_dataframe_memory(df):
    """Optimize DataFrame memory usage for large files"""
    try:
        initial_memory = df.memory_usage(deep=True).sum()
        
        # Optimize string columns
        for col in df.select_dtypes(include=['object']):
            df[col] = df[col].astype('category')
        
        # Optimize numeric columns
        for col in df.select_dtypes(include=['int64']):
            if df[col].min() >= 0:
                if df[col].max() < 255:
                    df[col] = df[col].astype('uint8')
                elif df[col].max() < 65535:
                    df[col] = df[col].astype('uint16')
                elif df[col].max() < 4294967295:
                    df[col] = df[col].astype('uint32')
            else:
                if df[col].min() > -128 and df[col].max() < 127:
                    df[col] = df[col].astype('int8')
                elif df[col].min() > -32768 and df[col].max() < 32767:
                    df[col] = df[col].astype('int16')
                elif df[col].min() > -2147483648 and df[col].max() < 2147483647:
                    df[col] = df[col].astype('int32')
        
        final_memory = df.memory_usage(deep=True).sum()
        reduction = (initial_memory - final_memory) / initial_memory * 100
        
        logging.info(f"📊 Memory optimization: {reduction:.1f}% reduction ({initial_memory/1024/1024:.1f}MB → {final_memory/1024/1024:.1f}MB)")
        
        return df
    except Exception as e:
        logging.warning(f"⚠️ Memory optimization failed: {str(e)}")
        return df
    

@app.route('/debug/session-info')
@login_required
def debug_session_info():
    """Debug endpoint to check session status"""
    info = {
        'session_keys': list(session.keys()),
        'session_file_dir': app.config.get('SESSION_FILE_DIR'),
        'session_type': app.config.get('SESSION_TYPE'),
        'max_content_length': app.config.get('MAX_CONTENT_LENGTH'),
        'upload_folder': app.config.get('UPLOAD_FOLDER')
    }
    return jsonify(info)
    
def ensure_proper_base64_format(img_data):
    """
    Ensures that image data is properly formatted as a base64 string
    with the correct data URI prefix.
    """
    if img_data is None:
        return None
        
    # If it's not a string, assume it's bytes and encode it
    if not isinstance(img_data, str):
        img_data = base64.b64encode(img_data).decode('utf-8')
    
    # If it doesn't have the data URL prefix, add it
    if not img_data.startswith('data:image'):
        img_data = f"data:image/png;base64,{img_data}"
        
    return img_data

def get_plot_data_as_uri(self, plot_type):
    """
    Returns the plot data as a properly formatted data URI
    
    Args:
        plot_type: Type of plot to retrieve ('ml', 'pca', 'cluster', or 'raw')
        
    Returns:
        str: Data URI formatted image data, or None if not found
    """
    try:
        if plot_type == 'ml' and self.ml_plot_data:
            img_data = self.ml_plot_data
        elif plot_type == 'pca' and self.pca_plot_data:
            img_data = self.pca_plot_data
        elif plot_type == 'cluster' and self.cluster_plot_data:
            img_data = self.cluster_plot_data
        elif plot_type == 'raw' and self.raw_plot_data:
            img_data = self.raw_plot_data
        else:
            return None
            
        # Convert binary data to base64 string with data URI prefix
        base64_str = base64.b64encode(img_data).decode('utf-8')
        return f"data:image/png;base64,{base64_str}"
    except Exception as e:
        print(f"Error retrieving image data: {str(e)}")
        return None

@app.route("/pcaMethod", methods=['GET', 'POST'])
def pcaMethod():
    """Handle PCA analysis file uploads and processing - COMPLETE FIXED VERSION"""
    if request.method == 'GET':
        # GET request - show the initial PCA form
        return render_template('current_analysis_PCA.HTML', 
                             sample_placeholder="Please wait for processor response")
    
    if request.method == 'POST':
        visit = request.form.get("visit")
        
        if visit == "one":
            # First stage - handle file upload (both local and OneDrive)
            try:
                # Check file source
                file_source = request.form.get('fileSource', 'local')
                logging.info(f"🔍 PCA processing with file source: {file_source}")
                
                if file_source == 'cloud' or file_source == 'onedrive':
                    # Handle OneDrive file processing
                    logging.info("☁️ Processing OneDrive file for PCA...")
                    return process_onedrive_file_for_pca()
                else:
                    # Handle local file upload
                    if 'csv_file' not in request.files:
                        logging.warning("⚠️ No file part in request")
                        return render_template('current_analysis_PCA.HTML', 
                                             sample_placeholder="Error: No file uploaded")
                        
                    folder_path = request.files['csv_file']
                    if folder_path.filename == '':
                        logging.warning("⚠️ No selected file")
                        return render_template('current_analysis_PCA.HTML', 
                                             sample_placeholder="Error: No file selected")
                    
                    # Check file extension
                    if not folder_path.filename.lower().endswith('.csv'):
                        logging.warning(f"⚠️ Invalid file type: {folder_path.filename}")
                        return render_template('current_analysis_PCA.HTML', 
                                             sample_placeholder="Error: Only CSV files are allowed")
                    
                    # Get data type selection
                    data_type = request.form.get("data_type")
                    
                    # Save uploaded file to uploads folder
                    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                    
                    # Use a timestamp to avoid filename collisions
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    safe_filename = f"{timestamp}_{secure_filename(folder_path.filename)}"
                    csv_filename = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
                    folder_path.save(csv_filename)
                    logging.info(f"📁 File saved to: {csv_filename}")
                    
                    # Store file path and data type in session
                    session['csv_file_path'] = csv_filename
                    session['data_type'] = data_type
                    session['is_large_file'] = False
                    session.modified = True
                    
                    # Read CSV file to extract unique values
                    try:
                        dfX = pd.read_csv(csv_filename)
                        logging.info(f"📊 Successfully read CSV with {len(dfX)} rows")
                    except Exception as e:
                        logging.error(f"❌ Error reading CSV: {str(e)}")
                        return render_template('current_analysis_PCA.HTML', 
                                             sample_placeholder=f"Error reading CSV file: {str(e)}")
                    
                    # Extract unique values from different columns
                    strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                    unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                    unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                    unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                    unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                    
                    logging.info(f"✅ Extracted unique values - bacteria: {len(strings)}, concs: {len(unq_concs)}")
                    
                    return render_template('current_analysis_PCA.HTML', 
                                         item=strings,
                                         item2=unq_concs,
                                         item3=unq_vols,
                                         item4=unq_sli,
                                         item5=unq_tri)
            except Exception as e:
                logging.error(f"❌ Error processing uploaded file: {str(e)}")
                return render_template('current_analysis_PCA.HTML', 
                                     sample_placeholder=f"Error: {str(e)}")
        else:
            # Second stage - process the data with selected options
            try:
                # Retrieve file path from session
                csv_file_path = session.get('csv_file_path')
                is_large_file = session.get('is_large_file', False)
                
                logging.info(f"🔄 PCA Step 2 - Processing file: {csv_file_path}")
                logging.info(f"📊 Large file processing: {is_large_file}")
                
                if not csv_file_path or not os.path.exists(csv_file_path):
                    return render_template('current_analysis_PCA.HTML', 
                                         sample_placeholder="Error: File not found. Please upload again.")
                
                # Get selected options from form
                bacts = request.form.getlist('options[]')
                conc = request.form.get('conc_typeX')
                vol = request.form.get('vol_typeX')
                slide = request.form.getlist('options3[]')
                trails = request.form.getlist('options2[]')
                
                logging.info(f"📋 Selected options - bacteria: {len(bacts)} ({bacts}), conc: {conc}")
                logging.info(f"📋 Selected vol: {vol}, slide: {len(slide)} ({slide}), trials: {len(trails)} ({trails})")
                
                # Validate selections
                if not bacts:
                    return render_template('current_analysis_PCA.HTML', 
                                         sample_placeholder="Error: Please select at least one bacteria type.")
                
                # Process the data with large file support
                if is_large_file:
                    logging.info(f"🔄 Processing large file for PCA...")
                    df = process_large_file_pca(csv_file_path, bacts, conc, vol, slide, trails)
                else:
                    # Standard processing
                    df = pd.read_csv(csv_file_path)
                    logging.info(f"📊 Read CSV for PCA processing: {len(df)} rows")
                    
                    # Apply filtering for standard processing (same as large file logic)
                    original_size = len(df)
                    
                    # Filter by bacteria (required)
                    if 'bacteria' in df.columns and bacts:
                        df = df[df['bacteria'].isin(bacts)]
                        logging.info(f"📋 After bacteria filter: {len(df)} rows from {original_size}")
                    
                    # Filter by concentration (if specified)
                    if conc and 'concentration' in df.columns:
                        before_conc = len(df)
                        try:
                            conc_numeric = float(conc)
                            df = df[df['concentration'] == conc_numeric]
                        except (ValueError, TypeError):
                            df_conc_str = df['concentration'].astype(str)
                            conc_str = str(conc)
                            df = df[df_conc_str == conc_str]
                        logging.info(f"📋 After concentration filter: {len(df)} rows from {before_conc}")
                    
                    # Filter by volume (if specified)
                    if vol and 'volume' in df.columns:
                        before_vol = len(df)
                        try:
                            vol_numeric = float(vol)
                            df = df[df['volume'] == vol_numeric]
                        except (ValueError, TypeError):
                            df_vol_str = df['volume'].astype(str)
                            vol_str = str(vol)
                            df = df[df_vol_str == vol_str]
                        logging.info(f"📋 After volume filter: {len(df)} rows from {before_vol}")
                    
                    # Filter by slide (if specified)
                    if slide and 'slide' in df.columns:
                        before_slide = len(df)
                        df_slide_str = df['slide'].astype(str)
                        slide_str = [str(s) for s in slide]
                        df = df[df_slide_str.isin(slide_str)]
                        logging.info(f"📋 After slide filter: {len(df)} rows from {before_slide}")
                    
                    # Filter by trial (if specified)
                    if trails and 'trial' in df.columns:
                        before_trial = len(df)
                        df_trial_str = df['trial'].astype(str)
                        trails_str = [str(t) for t in trails]
                        df = df[df_trial_str.isin(trails_str)]
                        logging.info(f"📋 After trial filter: {len(df)} rows from {before_trial}")
                    
                    if df.empty:
                        logging.warning("⚠️ No data matches the selected criteria")
                        return render_template('current_analysis_PCA.HTML', 
                                             sample_placeholder="No data matches your selection criteria. Please try different options.")
                
                # Generate PCA visualization
                logging.info(f"🎨 Generating PCA visualization with {len(df)} rows...")
                img_base64 = main_function_pca(df, bacts, conc, vol, slide, trails)
                logging.info("✅ PCA visualization generated successfully")
                
                # Clean up the session after processing
                session.pop('csv_file_path', None)
                session.pop('data_type', None)
                session.pop('is_large_file', None)
                session.modified = True
                
                # Clean up the temporary file with retry logic
                cleanup_attempts = 3
                for attempt in range(cleanup_attempts):
                    try:
                        if os.path.exists(csv_file_path):
                            os.remove(csv_file_path)
                            logging.info(f"✅ Cleaned up temporary file: {csv_file_path}")
                            break
                    except Exception as e:
                        if attempt < cleanup_attempts - 1:
                            logging.warning(f"⚠️ Cleanup attempt {attempt + 1} failed: {str(e)}, retrying...")
                            time.sleep(0.5)
                        else:
                            logging.warning(f"⚠️ Could not delete temporary file after {cleanup_attempts} attempts: {str(e)}")
                
                return render_template('current_analysis_PCA.HTML', 
                                     img_base64=img_base64,
                                     sample_placeholder="PCA analysis completed successfully!")
            except Exception as e:
                logging.error(f"❌ Error during PCA analysis: {str(e)}")
                import traceback
                logging.error(f"❌ Full traceback: {traceback.format_exc()}")
                return render_template('current_analysis_PCA.HTML', 
                                     sample_placeholder=f"Error during analysis: {str(e)}")
    else:
        return redirect(url_for('NewPca'))
    
def process_onedrive_file_for_pca():
    """Process a file selected from OneDrive for PCA analysis - ENHANCED VERSION"""
    try:
        # Get OneDrive file ID from form
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            logging.error("No OneDrive file ID provided for PCA")
            return render_template('current_analysis_PCA.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available for PCA")
            return render_template('current_analysis_PCA.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        # Create API headers
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file details first
        file_name = None
        file_size = 0
        download_url = None
        
        logging.info(f"🔍 Getting PCA file info for OneDrive file ID: {file_id}")
        
        # Try to get file info and download URL
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    
                    logging.info(f"📁 PCA File: {file_name}, Size: {file_size} bytes ({file_size/1024/1024:.1f} MB)")
                    
                    # Get download URL
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        logging.info(f"✅ Got download URL for PCA processing")
                        break
                    else:
                        logging.warning("⚠️ No download URL found in file info")
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            logging.error("❌ Could not get file info or download URL for PCA")
            return render_template('current_analysis_PCA.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # Download and process file (with large file support)
        try:
            # Create temporary file
            import tempfile
            temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
            temp_file_path = temp_file.name
            
            # Download file content
            if file_size > 50 * 1024 * 1024:  # 50MB threshold
                logging.info(f"📥 Large PCA file detected ({file_size/1024/1024:.1f} MB), using streaming download")
                
                with requests.get(download_url, stream=True) as download_response:
                    download_response.raise_for_status()
                    
                    for chunk in download_response.iter_content(chunk_size=8192):
                        if chunk:
                            temp_file.write(chunk)
                            
                temp_file.close()
                is_large_file = True
            else:
                logging.info(f"📥 Standard download for PCA file ({file_size/1024/1024:.1f} MB)")
                
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    temp_file.write(download_response.content)
                    temp_file.close()
                    is_large_file = False
                else:
                    temp_file.close()
                    os.unlink(temp_file_path)
                    logging.error(f"❌ Download failed: {download_response.status_code}")
                    return render_template('current_analysis_PCA.HTML', 
                                         sample_placeholder="Error: Could not download file from OneDrive")
            
            # Store the file path in session for step 2
            session['csv_file_path'] = temp_file_path
            session['data_type'] = request.form.get("data_type", "raw_data")
            session['is_large_file'] = is_large_file
            session['file_size_mb'] = file_size / 1024 / 1024
            session.modified = True
            
            # Read and extract unique values
            if is_large_file and file_size > 100 * 1024 * 1024:
                logging.info(f"🔄 Processing very large PCA file with chunking...")
                
                # Use chunked processing for very large files
                chunk_size_rows = 10000
                unique_bacteria = set()
                unique_concs = set()
                unique_vols = set()
                unique_slides = set()
                unique_trials = set()
                
                chunk_count = 0
                for chunk in pd.read_csv(temp_file_path, chunksize=chunk_size_rows):
                    chunk_count += 1
                    
                    # Extract unique values from this chunk
                    if 'bacteria' in chunk.columns:
                        unique_bacteria.update(chunk['bacteria'].dropna().unique())
                    if 'concentration' in chunk.columns:
                        unique_concs.update(chunk['concentration'].dropna().unique())
                    if 'volume' in chunk.columns:
                        unique_vols.update(chunk['volume'].dropna().unique())
                    if 'slide' in chunk.columns:
                        unique_slides.update(chunk['slide'].dropna().unique())
                    if 'trial' in chunk.columns:
                        unique_trials.update(chunk['trial'].dropna().unique())
                    
                    if chunk_count % 10 == 0:
                        logging.info(f"📊 PCA: Processed {chunk_count} chunks, found {len(unique_bacteria)} bacteria types")
                
                # Convert sets to sorted lists
                strings = sorted(list(unique_bacteria))
                unq_concs = sorted(list(unique_concs))
                unq_vols = sorted(list(unique_vols))
                unq_sli = sorted(list(unique_slides))
                unq_tri = sorted(list(unique_trials))
                
                logging.info(f"✅ PCA chunked processing complete: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                
            else:
                # Regular processing for smaller files
                logging.info(f"🔄 Processing PCA file normally...")
                dfX = pd.read_csv(temp_file_path)
                logging.info(f"📊 Successfully read PCA CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                
                # Extract unique values
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                
                logging.info(f"✅ PCA: Extracted unique values: {len(strings)} bacteria types")
            
            logging.info(f"💾 Stored PCA file path in session: {temp_file_path}")
            
            return render_template('current_analysis_PCA.HTML', 
                                  item=strings, 
                                  item2=unq_concs, 
                                  item3=unq_vols, 
                                  item4=unq_sli, 
                                  item5=unq_tri,
                                  file_size_info=f"File: {file_size/1024/1024:.1f} MB" if is_large_file else None)
                                  
        except Exception as e:
            logging.error(f"❌ Error processing OneDrive PCA file: {str(e)}")
            # Clean up temp file on error
            try:
                if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except:
                pass
            return render_template('current_analysis_PCA.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive PCA file processing: {str(e)}")
        return render_template('current_analysis_PCA.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")

def process_onedrive_file_for_ml_simple():
    """Simple OneDrive file processing for ML"""
    try:
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        access_token = get_master_onedrive_token()
        if not access_token:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file info and download
        file_name = None
        file_size = 0
        download_url = None
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    if download_url:
                        break
            except Exception as e:
                continue
        
        if not download_url:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # Download file to temp location
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
        temp_file_path = temp_file.name
        
        try:
            # Simple download approach
            download_response = requests.get(download_url, stream=True)
            download_response.raise_for_status()
            
            for chunk in download_response.iter_content(chunk_size=8192):
                if chunk:
                    temp_file.write(chunk)
            temp_file.close()
            
            # Store file path in session
            session['csv_file_path'] = temp_file_path
            session['data_type'] = request.form.get("data_type", "raw_data")
            session.modified = True
            
            # Extract unique values efficiently
            if file_size > 100 * 1024 * 1024:  # For files > 100MB
                # Sample first 50k rows to get unique values
                sample_df = pd.read_csv(temp_file_path, nrows=50000)
                strings = sample_df['bacteria'].unique().tolist() if 'bacteria' in sample_df.columns else []
                unq_concs = sample_df['concentration'].unique().tolist() if 'concentration' in sample_df.columns else []
                unq_vols = sample_df['volume'].unique().tolist() if 'volume' in sample_df.columns else []
                unq_sli = sample_df['slide'].unique().tolist() if 'slide' in sample_df.columns else []
                unq_tri = sample_df['trial'].unique().tolist() if 'trial' in sample_df.columns else []
            else:
                # Read full file for smaller files
                dfX = pd.read_csv(temp_file_path)
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
            
            return render_template('current_analysis_ML.HTML', 
                                  item=strings, item2=unq_concs, item3=unq_vols, 
                                  item4=unq_sli, item5=unq_tri)
                                  
        except Exception as e:
            logging.error(f"❌ Error processing OneDrive ML file: {str(e)}")
            try:
                if os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except:
                pass
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive ML file processing: {str(e)}")
        return render_template('current_analysis_ML.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")
            
@app.route("/MLMethod", methods=['GET', 'POST'])
def MLMethod():
    """Handle ML analysis - ULTRA FIXED VERSION"""
    if request.method == 'GET':
        return render_template('current_analysis_ML.HTML', 
                             sample_placeholder="Please wait for processor response")
    
    if request.method == 'POST':
        visit = request.form.get("visit")
        
        if visit == "one":
            # First stage - handle file upload (both local and OneDrive)
            try:
                file_source = request.form.get('fileSource', 'local')
                logging.info(f"🔍 ML processing with file source: {file_source}")
                
                if file_source == 'cloud' or file_source == 'onedrive':
                    # Handle OneDrive file processing - ULTRA FIXED VERSION
                    return process_onedrive_file_for_ml_ultra_fix()
                else:
                    # Handle local file upload - keep original working logic
                    if 'csv_file' not in request.files:
                        return render_template('current_analysis_ML.HTML', 
                                             sample_placeholder="Error: No file uploaded")
                        
                    folder_path = request.files['csv_file']
                    if folder_path.filename == '':
                        return render_template('current_analysis_ML.HTML', 
                                             sample_placeholder="Error: No file selected")
                    
                    if not folder_path.filename.lower().endswith('.csv'):
                        return render_template('current_analysis_ML.HTML', 
                                             sample_placeholder="Error: Only CSV files are allowed")
                    
                    data_type = request.form.get("data_type")
                    
                    # Save file
                    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    safe_filename = f"{timestamp}_{secure_filename(folder_path.filename)}"
                    csv_filename = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
                    folder_path.save(csv_filename)
                    
                    # Store in session
                    session['csv_file_path'] = csv_filename
                    session['data_type'] = data_type
                    session['is_large_file'] = False
                    session.modified = True
                    
                    # Read and extract unique values using the same efficient method
                    strings, unq_concs, unq_vols, unq_sli, unq_tri = extract_unique_values_efficiently(csv_filename)
                    
                    logging.info(f"✅ Local ML: Extracted {len(strings)} bacteria, {len(unq_concs)} concentrations")
                    
                    return render_template('current_analysis_ML.HTML', 
                                         item=strings, item2=unq_concs, item3=unq_vols, 
                                         item4=unq_sli, item5=unq_tri)
            except Exception as e:
                logging.error(f"❌ Error processing uploaded ML file: {str(e)}")
                return render_template('current_analysis_ML.HTML', 
                                     sample_placeholder=f"Error: {str(e)}")
        else:
            # Second stage - process with selected options (keep existing logic)
            try:
                csv_file_path = session.get('csv_file_path')
                is_large_file = session.get('is_large_file', False)
                
                if not csv_file_path or not os.path.exists(csv_file_path):
                    return render_template('current_analysis_ML.HTML', 
                                         sample_placeholder="Error: File not found. Please upload again.")
                
                # Get selected options
                bacts = request.form.getlist('options[]')
                conc = request.form.get('conc_typeX')
                vol = request.form.get('vol_typeX')
                slide = request.form.getlist('options3[]')
                trails = request.form.getlist('options2[]')
                
                if not bacts:
                    return render_template('current_analysis_ML.HTML', 
                                         sample_placeholder="Error: Please select at least one bacteria type.")
                
                # Process the data
                df = pd.read_csv(csv_file_path)
                logging.info(f"🎨 Calling plotter_svm with dataset: {len(df)} rows")
                
                # Call ML function
                img_base64 = plotter_svm(df, bacts, conc, vol, slide, trails)
                
                # Cleanup
                session.pop('csv_file_path', None)
                session.pop('data_type', None)
                session.pop('is_large_file', None)
                session.modified = True
                
                try:
                    if os.path.exists(csv_file_path):
                        os.remove(csv_file_path)
                except:
                    pass
                
                return render_template('current_analysis_ML.HTML', 
                                     img_base64=img_base64,
                                     sample_placeholder="ML analysis completed successfully!")
                                     
            except Exception as e:
                logging.error(f"❌ Error during ML analysis: {str(e)}")
                return render_template('current_analysis_ML.HTML', 
                                     sample_placeholder=f"Error during analysis: {str(e)}")
    else:
        return redirect(url_for('NewML'))
    
def process_onedrive_file_for_ml_proper():
    """Process OneDrive file for ML - NO PRE-FILTERING VERSION"""
    try:
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        access_token = get_master_onedrive_token()
        if not access_token:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file info
        file_name = None
        file_size = 0
        download_url = None
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    if download_url:
                        break
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # Download file
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
        temp_file_path = temp_file.name
        
        try:
            if file_size > 50 * 1024 * 1024:  # 50MB threshold
                # Stream download for large files
                with requests.get(download_url, stream=True) as download_response:
                    download_response.raise_for_status()
                    for chunk in download_response.iter_content(chunk_size=8192):
                        if chunk:
                            temp_file.write(chunk)
                temp_file.close()
            else:
                # Direct download for smaller files
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    temp_file.write(download_response.content)
                    temp_file.close()
                else:
                    temp_file.close()
                    os.unlink(temp_file_path)
                    return render_template('current_analysis_ML.HTML', 
                                         sample_placeholder="Error: Could not download file from OneDrive")
            
            # Store file path in session
            session['csv_file_path'] = temp_file_path
            session['data_type'] = request.form.get("data_type", "raw_data")
            session.modified = True
            
            # Extract unique values for the form - READ EFFICIENTLY FOR LARGE FILES
            if file_size > 100 * 1024 * 1024:
                # For very large files, sample data to get unique values
                logging.info(f"🔄 Sampling large file for unique values...")
                sample_df = pd.read_csv(temp_file_path, nrows=50000)  # Sample first 50k rows
                strings = sample_df['bacteria'].unique().tolist() if 'bacteria' in sample_df.columns else []
                unq_concs = sample_df['concentration'].unique().tolist() if 'concentration' in sample_df.columns else []
                unq_vols = sample_df['volume'].unique().tolist() if 'volume' in sample_df.columns else []
                unq_sli = sample_df['slide'].unique().tolist() if 'slide' in sample_df.columns else []
                unq_tri = sample_df['trial'].unique().tolist() if 'trial' in sample_df.columns else []
            else:
                # For smaller files, read normally
                dfX = pd.read_csv(temp_file_path)
                strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
            
            return render_template('current_analysis_ML.HTML', 
                                  item=strings, item2=unq_concs, item3=unq_vols, 
                                  item4=unq_sli, item5=unq_tri)
                                  
        except Exception as e:
            logging.error(f"❌ Error processing OneDrive ML file: {str(e)}")
            try:
                if os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except:
                pass
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive ML file processing: {str(e)}")
        return render_template('current_analysis_ML.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")
    
def extract_unique_values_efficiently(file_path):
    """Extract unique values efficiently by reading only the columns we need"""
    try:
        logging.info("🔄 Using efficient column-specific extraction...")
        
        # First, get column names
        sample_df = pd.read_csv(file_path, nrows=1)
        available_columns = list(sample_df.columns)
        logging.info(f"📊 Available columns: {available_columns}")
        
        # Define the columns we need for unique values
        target_columns = ['bacteria', 'concentration', 'volume', 'slide', 'trial']
        columns_to_read = [col for col in target_columns if col in available_columns]
        
        if not columns_to_read:
            logging.warning("⚠️ No target columns found in file")
            return [], [], [], [], []
        
        logging.info(f"📊 Reading columns: {columns_to_read}")
        
        # Read only the columns we need for unique value extraction
        df_subset = pd.read_csv(file_path, usecols=columns_to_read)
        logging.info(f"📊 Read {len(df_subset)} rows with {len(df_subset.columns)} columns")
        
        # Extract unique values
        strings = df_subset['bacteria'].unique().tolist() if 'bacteria' in df_subset.columns else []
        unq_concs = df_subset['concentration'].unique().tolist() if 'concentration' in df_subset.columns else []
        unq_vols = df_subset['volume'].unique().tolist() if 'volume' in df_subset.columns else []
        unq_sli = df_subset['slide'].unique().tolist() if 'slide' in df_subset.columns else []
        unq_tri = df_subset['trial'].unique().tolist() if 'trial' in df_subset.columns else []
        
        # Sort for consistent output
        strings = sorted(strings) if strings else []
        unq_concs = sorted(unq_concs) if unq_concs else []
        unq_vols = sorted(unq_vols) if unq_vols else []
        unq_sli = sorted(unq_sli) if unq_sli else []
        unq_tri = sorted(unq_tri) if unq_tri else []
        
        logging.info(f"✅ Efficient extraction results:")
        logging.info(f"   🦠 Bacteria: {len(strings)} types - {strings}")
        logging.info(f"   🧪 Concentrations: {len(unq_concs)} values - {unq_concs}")
        logging.info(f"   📏 Volumes: {len(unq_vols)} values - {unq_vols}")
        logging.info(f"   🔬 Slides: {len(unq_sli)} values - {unq_sli}")
        logging.info(f"   🔄 Trials: {len(unq_tri)} values - {unq_tri}")
        
        return strings, unq_concs, unq_vols, unq_sli, unq_tri
        
    except Exception as e:
        logging.error(f"❌ Error in efficient extraction: {str(e)}")
        raise e


def process_onedrive_file_for_ml_ultra_fix():
    """Process OneDrive file for ML - ULTRA FIX with efficient column reading"""
    try:
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            logging.error("No OneDrive file ID provided for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file info and download (same as before)
        file_name = None
        file_size = 0
        download_url = None
        
        logging.info(f"🔍 Getting ML file info for OneDrive file ID: {file_id}")
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        break
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # Download file
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
        temp_file_path = temp_file.name
        
        try:
            # Download the file
            download_response = requests.get(download_url, stream=True)
            download_response.raise_for_status()
            
            for chunk in download_response.iter_content(chunk_size=8192):
                if chunk:
                    temp_file.write(chunk)
            temp_file.close()
            
            logging.info(f"✅ Downloaded OneDrive file: {file_name} ({file_size/1024/1024:.1f} MB)")
            
            # Store file path in session
            session['csv_file_path'] = temp_file_path
            session['data_type'] = request.form.get("data_type", "raw_data")
            session['is_large_file'] = file_size > 50 * 1024 * 1024
            session['file_size_mb'] = file_size / 1024 / 1024
            session.modified = True
            
            # ULTRA FIX: Use the most efficient extraction method
            strings, unq_concs, unq_vols, unq_sli, unq_tri = extract_unique_values_efficiently(temp_file_path)
            
            return render_template('current_analysis_ML.HTML', 
                                  item=strings, 
                                  item2=unq_concs, 
                                  item3=unq_vols, 
                                  item4=unq_sli, 
                                  item5=unq_tri,
                                  file_size_info=f"OneDrive: {len(strings)} bacteria types extracted")
                                  
        except Exception as e:
            logging.error(f"❌ Error processing OneDrive ML file: {str(e)}")
            try:
                if os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except:
                pass
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive ML file processing: {str(e)}")
        return render_template('current_analysis_ML.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")
    

def process_large_file_ml_fixed(file_path, bacts, conc, vol, slide, trails, chunk_size=10000):
    """Process large files for ML analysis using chunked reading - IMPROVED VERSION"""
    try:
        logging.info(f"🔄 Processing large file for ML: {file_path}")
        
        # Read first chunk to understand data structure
        first_chunk = pd.read_csv(file_path, nrows=1000)
        logging.info(f"📊 ML Data structure - Columns: {list(first_chunk.columns)}")
        
        # Validate required columns for ML
        required_cols = ['bacteria']
        missing_cols = [col for col in required_cols if col not in first_chunk.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns for ML: {missing_cols}")
        
        # Process in chunks and filter
        filtered_chunks = []
        total_rows = 0
        matching_rows = 0
        
        for chunk in pd.read_csv(file_path, chunksize=chunk_size):
            total_rows += len(chunk)
            
            # Apply basic filtering - only filter by bacteria to avoid empty results
            original_chunk_size = len(chunk)
            
            # Filter by bacteria (the most important filter)
            if bacts and 'bacteria' in chunk.columns:
                chunk = chunk[chunk['bacteria'].isin(bacts)]
            
            # Apply other filters more carefully
            if conc and 'concentration' in chunk.columns and not chunk.empty:
                try:
                    conc_numeric = float(conc)
                    chunk = chunk[chunk['concentration'] == conc_numeric]
                except (ValueError, TypeError):
                    chunk = chunk[chunk['concentration'].astype(str) == str(conc)]
            
            if vol and 'volume' in chunk.columns and not chunk.empty:
                try:
                    vol_numeric = float(vol)
                    chunk = chunk[chunk['volume'] == vol_numeric]
                except (ValueError, TypeError):
                    chunk = chunk[chunk['volume'].astype(str) == str(vol)]
            
            if slide and 'slide' in chunk.columns and not chunk.empty:
                slide_str = [str(s) for s in slide]
                chunk = chunk[chunk['slide'].astype(str).isin(slide_str)]
            
            if trails and 'trial' in chunk.columns and not chunk.empty:
                trails_str = [str(t) for t in trails]
                chunk = chunk[chunk['trial'].astype(str).isin(trails_str)]
            
            if not chunk.empty:
                filtered_chunks.append(chunk)
                matching_rows += len(chunk)
        
        if not filtered_chunks:
            raise ValueError("No data matches the selected bacteria types")
        
        # Combine all filtered chunks
        logging.info(f"📊 Combining {len(filtered_chunks)} filtered chunks for ML...")
        combined_df = pd.concat(filtered_chunks, ignore_index=True)
        
        # Optimize memory usage
        combined_df = optimize_pca_dataframe_memory(combined_df)
        
        logging.info(f"✅ Large file ML processing complete:")
        logging.info(f"   - Total rows processed: {total_rows}")
        logging.info(f"   - Matching rows: {matching_rows}")
        logging.info(f"   - Final dataset size: {len(combined_df)} rows")
        
        return combined_df
        
    except Exception as e:
        logging.error(f"❌ Error in large file ML processing: {str(e)}")
        raise e
    
@app.route('/api/save-ml-to-onedrive', methods=['POST'])
@login_required
def save_ml_to_onedrive():
    """API endpoint to save ML analysis result to OneDrive"""
    try:
        # Get the image data from the request
        image_data = request.form.get('image_data')
        if not image_data:
            return jsonify({
                'success': False,
                'message': 'No image data provided'
            }), 400
        
        # Get filename with timestamp
        filename = request.form.get('filename', 'ml_analysis.png')
        if not filename.lower().endswith('.png'):
            filename += '.png'
        
        # Add timestamp to filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"ml_{timestamp}_{filename}"
        
        # Convert base64 image data to binary
        try:
            # Remove data URL prefix if present
            if ',' in image_data:
                image_data = image_data.split(',', 1)[1]
            
            file_content = base64.b64decode(image_data)
        except Exception as e:
            return jsonify({
                'success': False,
                'message': f'Invalid image data: {str(e)}'
            }), 400
        
        # Get access token
        access_token = get_master_onedrive_token()
        if not access_token:
            return jsonify({
                'success': False,
                'message': 'OneDrive is not properly configured'
            }), 500
        
        # Create headers for API calls
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Get username for the folder
        username = current_user.username
        
        # Step 1: Find or create user folder
        user_folder_id = None
        
        # Look for user folder in root
        for endpoint in [
            'https://graph.microsoft.com/v1.0/me/drive/root/children',
            'https://api.onedrive.com/v1.0/drive/root/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    root_items = response.json().get('value', [])
                    for item in root_items:
                        if item.get('name') == username and item.get('folder') is not None:
                            user_folder_id = item.get('id')
                            logging.info(f"Found user folder: {username}, id: {user_folder_id}")
                            break
                    
                    if user_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create user folder if it doesn't exist
        if not user_folder_id:
            logging.info(f"Creating user folder: {username}")
            create_folder_data = {
                'name': username,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                'https://graph.microsoft.com/v1.0/me/drive/root/children',
                'https://api.onedrive.com/v1.0/drive/root/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        user_folder_id = create_response.json().get('id')
                        logging.info(f"Created user folder: {username}, id: {user_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating user folder with {endpoint}: {str(e)}")
        
        if not user_folder_id:
            return jsonify({
                'success': False,
                'message': 'Could not find or create user folder in OneDrive'
            }), 500
        
        # Step 2: Find or create ML results subfolder
        ml_folder_name = 'ml_results'
        ml_folder_id = None
        
        # Look for ML folder in user folder
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
            f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    folder_items = response.json().get('value', [])
                    for item in folder_items:
                        if item.get('name') == ml_folder_name and item.get('folder') is not None:
                            ml_folder_id = item.get('id')
                            logging.info(f"Found ML folder: {ml_folder_name}, id: {ml_folder_id}")
                            break
                    
                    if ml_folder_id:
                        break  # Exit the loop if we found the folder
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
        
        # Create ML folder if it doesn't exist
        if not ml_folder_id:
            logging.info(f"Creating ML folder: {ml_folder_name}")
            create_folder_data = {
                'name': ml_folder_name,
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            for endpoint in [
                f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                f'https://api.onedrive.com/v1.0/drive/items/{user_folder_id}/children'
            ]:
                try:
                    create_response = requests.post(
                        endpoint,
                        headers={**headers, 'Content-Type': 'application/json'},
                        json=create_folder_data
                    )
                    
                    if create_response.status_code in [200, 201]:
                        ml_folder_id = create_response.json().get('id')
                        logging.info(f"Created ML folder: {ml_folder_name}, id: {ml_folder_id}")
                        break
                except Exception as e:
                    logging.error(f"Error creating ML folder with {endpoint}: {str(e)}")
        
        if not ml_folder_id:
            return jsonify({
                'success': False,
                'message': f'Could not find or create {ml_folder_name} folder in OneDrive'
            }), 500
        
        # Step 3: Upload ML visualization to ML folder
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'image/png'
        }
        
        file_uploaded = False
        
        for endpoint in [
            f'https://graph.microsoft.com/v1.0/me/drive/items/{ml_folder_id}:/{filename}:/content',
            f'https://api.onedrive.com/v1.0/drive/items/{ml_folder_id}:/{filename}:/content'
        ]:
            try:
                upload_response = requests.put(
                    endpoint,
                    headers=upload_headers,
                    data=file_content
                )
                
                if upload_response.status_code in [200, 201]:
                    logging.info(f"Successfully uploaded ML visualization to OneDrive: {filename}")
                    file_uploaded = True
                    break
            except Exception as e:
                logging.error(f"Error uploading ML visualization with {endpoint}: {str(e)}")
        
        if not file_uploaded:
            return jsonify({
                'success': False,
                'message': 'Failed to upload ML visualization to OneDrive'
            }), 500
        
        # Return success response
        return jsonify({
            'success': True,
            'message': f'ML analysis result successfully saved to OneDrive',
            'filename': filename,
            'folder': ml_folder_name
        })
        
    except Exception as e:
        logging.error(f"Error saving ML visualization to OneDrive: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        }), 500


def process_onedrive_file_for_ml_fixed():
    """Process OneDrive file for ML - FIXED VERSION that properly extracts unique values"""
    try:
        file_id = request.form.get('onedrive_file_id')
        if not file_id:
            logging.error("No OneDrive file ID provided for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: No OneDrive file selected")
        
        access_token = get_master_onedrive_token()
        if not access_token:
            logging.error("No OneDrive access token available for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: OneDrive is not properly configured")
        
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get file info first
        file_name = None
        file_size = 0
        download_url = None
        
        logging.info(f"🔍 Getting ML file info for OneDrive file ID: {file_id}")
        
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                
                if response.status_code == 200:
                    file_info = response.json()
                    file_name = file_info.get('name')
                    file_size = file_info.get('size', 0)
                    
                    logging.info(f"📁 ML File: {file_name}, Size: {file_size} bytes ({file_size/1024/1024:.1f} MB)")
                    
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        logging.info(f"✅ Got download URL for ML processing")
                        break
                    else:
                        logging.warning("⚠️ No download URL found in file info")
            except Exception as e:
                logging.error(f"❌ Error with {endpoint}: {str(e)}")
                continue
        
        if not download_url or not file_name:
            logging.error("❌ Could not get file info or download URL for ML")
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder="Error: Could not access file from OneDrive")
        
        # Download file to temporary location
        try:
            import tempfile
            temp_file = tempfile.NamedTemporaryFile(mode='wb', suffix='.csv', delete=False)
            temp_file_path = temp_file.name
            
            # Download file content based on size
            if file_size > 50 * 1024 * 1024:  # 50MB threshold
                logging.info(f"📥 Large ML file detected ({file_size/1024/1024:.1f} MB), using streaming download")
                
                with requests.get(download_url, stream=True) as download_response:
                    download_response.raise_for_status()
                    for chunk in download_response.iter_content(chunk_size=8192):
                        if chunk:
                            temp_file.write(chunk)
                temp_file.close()
                is_large_file = True
            else:
                logging.info(f"📥 Standard download for ML file ({file_size/1024/1024:.1f} MB)")
                
                download_response = requests.get(download_url)
                if download_response.status_code == 200:
                    temp_file.write(download_response.content)
                    temp_file.close()
                    is_large_file = False
                else:
                    temp_file.close()
                    os.unlink(temp_file_path)
                    logging.error(f"❌ Download failed: {download_response.status_code}")
                    return render_template('current_analysis_ML.HTML', 
                                         sample_placeholder="Error: Could not download file from OneDrive")
            
            # Store file path in session
            session['csv_file_path'] = temp_file_path
            session['data_type'] = request.form.get("data_type", "raw_data")
            session['is_large_file'] = is_large_file
            session['file_size_mb'] = file_size / 1024 / 1024
            session.modified = True
            
            # CRITICAL FIX: Extract unique values properly for ALL file sizes
            strings = []
            unq_concs = []
            unq_vols = []
            unq_sli = []
            unq_tri = []
            
            try:
                if is_large_file and file_size > 100 * 1024 * 1024:
                    logging.info(f"🔄 Processing very large ML file with smart sampling...")
                    
                    # For very large files, use smart sampling to get comprehensive unique values
                    unique_bacteria = set()
                    unique_concs = set()
                    unique_vols = set()
                    unique_slides = set()
                    unique_trials = set()
                    
                    # Sample from beginning, middle, and end of file
                    total_lines = sum(1 for line in open(temp_file_path, 'r'))
                    logging.info(f"📊 Total lines in large ML file: {total_lines}")
                    
                    # Sample from different parts of the file
                    sample_sizes = [10000, 5000, 5000]  # Beginning, middle, end
                    skip_values = [0, max(0, total_lines//2 - 2500), max(0, total_lines - 5000)]
                    
                    for i, (sample_size, skip_rows) in enumerate(zip(sample_sizes, skip_values)):
                        try:
                            if skip_rows > 0:
                                chunk = pd.read_csv(temp_file_path, skiprows=range(1, skip_rows), nrows=sample_size)
                            else:
                                chunk = pd.read_csv(temp_file_path, nrows=sample_size)
                            
                            logging.info(f"📊 ML Sample {i+1}: {len(chunk)} rows from line {skip_rows}")
                            
                            # Extract unique values from this chunk
                            if 'bacteria' in chunk.columns:
                                unique_bacteria.update(chunk['bacteria'].dropna().unique())
                            if 'concentration' in chunk.columns:
                                unique_concs.update(chunk['concentration'].dropna().unique())
                            if 'volume' in chunk.columns:
                                unique_vols.update(chunk['volume'].dropna().unique())
                            if 'slide' in chunk.columns:
                                unique_slides.update(chunk['slide'].dropna().unique())
                            if 'trial' in chunk.columns:
                                unique_trials.update(chunk['trial'].dropna().unique())
                                
                        except Exception as e:
                            logging.warning(f"⚠️ Error processing sample {i+1}: {str(e)}")
                            continue
                    
                    # Convert sets to sorted lists
                    strings = sorted(list(unique_bacteria))
                    unq_concs = sorted(list(unique_concs))
                    unq_vols = sorted(list(unique_vols))
                    unq_sli = sorted(list(unique_slides))
                    unq_tri = sorted(list(unique_trials))
                    
                    logging.info(f"✅ ML smart sampling complete: {len(strings)} bacteria, {len(unq_concs)} concentrations")
                    
                else:
                    # Regular processing for smaller files
                    logging.info(f"🔄 Processing ML file normally...")
                    dfX = pd.read_csv(temp_file_path)
                    logging.info(f"📊 Successfully read ML CSV with {len(dfX)} rows and {len(dfX.columns)} columns")
                    
                    # Extract unique values - FIXED to handle missing columns gracefully
                    strings = dfX['bacteria'].unique().tolist() if 'bacteria' in dfX.columns else []
                    unq_concs = dfX['concentration'].unique().tolist() if 'concentration' in dfX.columns else []
                    unq_vols = dfX['volume'].unique().tolist() if 'volume' in dfX.columns else []
                    unq_sli = dfX['slide'].unique().tolist() if 'slide' in dfX.columns else []
                    unq_tri = dfX['trial'].unique().tolist() if 'trial' in dfX.columns else []
                    
                    logging.info(f"✅ ML OneDrive: Extracted {len(strings)} bacteria types, {len(unq_concs)} concentrations")
                    
                    # DEBUG: Log the actual values found
                    logging.info(f"🔍 ML DEBUG - Bacteria found: {strings[:5] if strings else 'None'}")
                    logging.info(f"🔍 ML DEBUG - Concentrations found: {unq_concs[:5] if unq_concs else 'None'}")
                    logging.info(f"🔍 ML DEBUG - Volumes found: {unq_vols[:5] if unq_vols else 'None'}")
                    logging.info(f"🔍 ML DEBUG - Slides found: {unq_sli[:5] if unq_sli else 'None'}")
                    logging.info(f"🔍 ML DEBUG - Trials found: {unq_tri[:5] if unq_tri else 'None'}")
            
            except Exception as e:
                logging.error(f"❌ Error extracting unique values from ML file: {str(e)}")
                # Try to provide some fallback values if possible
                try:
                    # Read just the first 1000 rows as absolute fallback
                    fallback_df = pd.read_csv(temp_file_path, nrows=1000)
                    strings = fallback_df['bacteria'].unique().tolist() if 'bacteria' in fallback_df.columns else []
                    unq_concs = fallback_df['concentration'].unique().tolist() if 'concentration' in fallback_df.columns else []
                    unq_vols = fallback_df['volume'].unique().tolist() if 'volume' in fallback_df.columns else []
                    unq_sli = fallback_df['slide'].unique().tolist() if 'slide' in fallback_df.columns else []
                    unq_tri = fallback_df['trial'].unique().tolist() if 'trial' in fallback_df.columns else []
                    logging.info(f"🔄 ML fallback extraction: {len(strings)} bacteria found")
                except Exception as fallback_error:
                    logging.error(f"❌ Even fallback extraction failed: {str(fallback_error)}")
                    return render_template('current_analysis_ML.HTML', 
                                         sample_placeholder=f"Error reading file structure: {str(e)}")
            
            logging.info(f"💾 Stored ML file path in session: {temp_file_path}")
            
            # CRITICAL: Return template with properly extracted values
            return render_template('current_analysis_ML.HTML', 
                                  item=strings, 
                                  item2=unq_concs, 
                                  item3=unq_vols, 
                                  item4=unq_sli, 
                                  item5=unq_tri,
                                  file_size_info=f"OneDrive file: {file_size/1024/1024:.1f} MB" if is_large_file else None)
                                  
        except Exception as e:
            logging.error(f"❌ Error processing OneDrive ML file: {str(e)}")
            # Clean up temp file on error
            try:
                if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except:
                pass
            return render_template('current_analysis_ML.HTML', 
                                 sample_placeholder=f"Error processing file: {str(e)}")
            
    except Exception as e:
        logging.error(f"❌ Critical error in OneDrive ML file processing: {str(e)}")
        return render_template('current_analysis_ML.HTML', 
                             sample_placeholder=f"Error accessing OneDrive file: {str(e)}")
    

# === User File Management Routes ===

# Route to serve plot files with better authentication check and security
@app.route('/plot/<path:filename>')
@login_required
def serve_plot(filename):
   """Serve a plot file with proper user access control"""
   # Extract user_id from the path to ensure proper access control
   try:
       path_parts = filename.split(os.sep)
       if len(path_parts) >= 2:
           plot_user_id = path_parts[1]  # User ID is the second part of the path
           
           # Check if current user has access to this plot
           if str(current_user.id) != plot_user_id and not current_user.is_admin():
               flash('You do not have permission to view this plot', 'error')
               return redirect(url_for('dashboard'))
           
           # Validate the file path to prevent directory traversal attacks
           if '..' in filename or filename.startswith('/'):
               flash('Invalid file path', 'error')
               return redirect(url_for('dashboard'))
               
           # Ensure the plot exists
           plot_path = os.path.join('Plots', filename)
           if not os.path.exists(plot_path) or not os.path.isfile(plot_path):
               flash('Plot file not found', 'error')
               return redirect(url_for('dashboard'))
               
           # Serve the file from Plots directory
           return send_from_directory('Plots', filename)
       else:
           flash('Invalid plot path', 'error')
           return redirect(url_for('dashboard'))
   except Exception as e:
       logging.error(f"Error serving plot file: {str(e)}")
       flash('Error accessing plot file', 'error')
       return redirect(url_for('dashboard'))

# Route to view file details
@app.route('/view_file/<int:file_id>')
@login_required
def view_file(file_id):
    """View details of a specific file"""
    try:
        # Get the file record
        file_record = UploadedFile.get(file_id)
        if not file_record:
            flash("File not found", "error")
            return redirect(url_for('dashboard'))
        
        # Check if user has permission to access this file
        if file_record.user_id != current_user.id and not current_user.is_admin():
            flash("You don't have permission to view this file", "error")
            return redirect(url_for('dashboard'))
        
        # Get file preview if possible
        preview_data = None
        try:
            # Try to get the DataFrame from the stored file content
            df = file_record.get_dataframe()
            if df is not None:
                # Limit to first 10 rows and 10 columns for preview
                if len(df) > 10:
                    preview_data = df.head(10)
                else:
                    preview_data = df
                
                if len(df.columns) > 10:
                    preview_data = preview_data.iloc[:, :10]
        except Exception as e:
            logging.error(f"Error reading file for preview: {str(e)}")
            flash(f"Could not generate data preview: {str(e)}", "warning")
        
        # Get related files (files processed with the same original filename)
        related_files = []
        try:
            all_files = UploadedFile.get_user_files(current_user.id)
            # Find files with similar original filename
            base_name = os.path.splitext(file_record.original_filename)[0]
            related_files = [f for f in all_files if f.id != file_record.id and 
                           base_name in f.original_filename][:5]  # Limit to 5
        except Exception as e:
            logging.error(f"Error fetching related files: {str(e)}")
        
        # Get image data based on processing type
        img_base64 = None
        if file_record.processed and file_record.processing_type:
            if 'raw' in file_record.processing_type:
                img_base64 = get_image_data_as_base64(file_record, 'raw')
            elif 'pca' in file_record.processing_type:
                img_base64 = get_image_data_as_base64(file_record, 'pca')
            elif 'ml' in file_record.processing_type or 'machine_learning' in file_record.processing_type:
                img_base64 = get_image_data_as_base64(file_record, 'ml')
            elif 'cluster' in file_record.processing_type:
                img_base64 = get_image_data_as_base64(file_record, 'cluster')
        
        return render_template('view_file.html', 
                             file=file_record, 
                             preview_data=preview_data,
                             related_files=related_files,
                             img_base64=img_base64,
                             processing_params=file_record.processing_parameters,
                             results_metadata=file_record.results_metadata)
    except Exception as e:
        logging.error(f"Error viewing file {file_id}: {str(e)}")
        flash(f"Error viewing file: {str(e)}", "error")
        return redirect(url_for('dashboard'))


# Route to download a file
@app.route('/download-file/<int:file_id>')
@login_required
def download_file(file_id):
    """Download the CSV file stored in the database"""
    try:
        # Get the file record
        file_record = UploadedFile.get(file_id)
        if not file_record:
            flash("File not found", "error")
            return redirect(url_for('dashboard'))
        
        # Check if user has permission to access this file
        if file_record.user_id != current_user.id and not current_user.is_admin():
            flash("You don't have permission to download this file", "error")
            return redirect(url_for('dashboard'))
        
        # Check if file content exists
        if not file_record.file_content:
            flash("File content not found in database", "error")
            return redirect(url_for('dashboard'))
        
        # Create a BytesIO object from the file content
        file_data = BytesIO(file_record.file_content)
        
        # Send the file for download
        return send_file(
            file_data,
            as_attachment=True,
            download_name=file_record.original_filename,
            mimetype=file_record.file_type or 'text/csv'
        )
    except Exception as e:
        logging.error(f"Error downloading file {file_id}: {str(e)}")
        flash(f"Error downloading file: {str(e)}", "error")
        return redirect(url_for('dashboard'))

# Route to delete a file
@app.route('/delete-file/<int:file_id>', methods=['POST'])
@login_required
def delete_file(file_id):
   """Delete a file and its associated plots with proper user access control"""
   try:
       # Verify user has permission to delete this file
       if not user_can_access_file(file_id):
           return jsonify({'success': False, 'error': 'You don\'t have permission to delete this file'}), 403
           
       # Get file record
       file_record = UploadedFile.get(file_id)
       if not file_record:
           return jsonify({'success': False, 'error': 'File not found'}), 404
       
       # Delete the file from disk
       if os.path.exists(file_record.file_path):
           try:
               os.remove(file_record.file_path)
               logging.info(f"Deleted file from disk: {file_record.file_path}")
           except Exception as e:
               logging.error(f"Error deleting file from disk: {str(e)}")

       # Delete associated plots if they exist
       for plot_path in [file_record.raw_plot_path, file_record.ml_plot_path, 
                        file_record.pca_plot_path, file_record.cluster_plot_path]:
           if plot_path and os.path.exists(plot_path):
               try:
                   os.remove(plot_path)
                   logging.info(f"Deleted plot from disk: {plot_path}")
               except Exception as e:
                   logging.error(f"Error deleting plot from disk: {str(e)}")

       # Delete from database
       if UploadedFile.delete(file_id):
           logging.info(f"User {current_user.id} deleted file {file_id}")
           return jsonify({'success': True})
       else:
           return jsonify({'success': False, 'error': 'Error deleting file record from database'}), 500
   except Exception as e:
       logging.error(f"Error deleting file {file_id}: {str(e)}")
       return jsonify({'success': False, 'error': str(e)}), 500

# New route to handle file uploads
@app.route('/upload-file', methods=['POST'])
@login_required
def upload_file():
   if 'file' not in request.files:
       return jsonify({'success': False, 'error': 'No file part'}), 400
   
   file = request.files['file']
   if file.filename == '':
       return jsonify({'success': False, 'error': 'No selected file'}), 400
   
   if file:
       # Create user-specific filename
       original_filename = file.filename
       unique_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
       file_path = get_user_file_path(unique_filename)
       
       # Save the uploaded file
       file.save(file_path)
       
       # Create file record
       file_record = add_file_record(file_path, original_filename, file.content_type)
       if not file_record:
           return jsonify({'success': False, 'error': 'Failed to create file record'}), 500
       
       return jsonify({
           'success': True,
           'file_id': file_record.id,
           'filename': file_record.original_filename
       })
   
   return jsonify({'success': False, 'error': 'File upload failed'}), 400

# Route for breathing analysis dashboard
@app.route('/breathe/dashboard')
@login_required
def breathe_dashboard():
   return redirect(url_for('home'))

# API route to get file processing status
@app.route('/api/file-status/<int:file_id>')
@login_required
def file_status_api(file_id):
   """API endpoint to get current status of a file"""
   if not user_can_access_file(file_id):
       return jsonify({'success': False, 'error': 'Access denied'}), 403
   
   file = UploadedFile.get(file_id)
   if not file:
       return jsonify({'success': False, 'error': 'File not found'}), 404
   
   result = {
       'success': True,
       'file_id': file.id,
       'filename': file.original_filename,
       'processed': file.processed,
       'processing_type': file.processing_type,
       'processed_at': file.processed_at.isoformat() if file.processed_at else None
   }
   
   # Add plot URLs if available
   plots = {}
   if file.raw_plot_path:
       plots['raw'] = url_for('serve_plot', filename=file.raw_plot_path)
   if file.ml_plot_path:
       plots['ml'] = url_for('serve_plot', filename=file.ml_plot_path)
   if file.pca_plot_path:
       plots['pca'] = url_for('serve_plot', filename=file.pca_plot_path)
   if file.cluster_plot_path:
       plots['cluster'] = url_for('serve_plot', filename=file.cluster_plot_path)
   
   result['plots'] = plots
   
   return jsonify(result)

# === Admin Routes ===

# Admin dashboard
@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    # Get users directly from the database
    conn = get_db()
    users = []
    files = []
    
    try:
        with conn.cursor() as cur:
            # Get users
            cur.execute("SELECT * FROM users ORDER BY id")
            user_rows = cur.fetchall()
            
            for row in user_rows:
                users.append(User(
                    id=row['id'],
                    username=row['username'],
                    email=row['email'],
                    password_hash=row['password_hash'],
                    role=row.get('role', 'user'),
                    created_at=row.get('created_at')
                ))
            
            # Get files
            cur.execute("SELECT * FROM uploaded_files ORDER BY uploaded_at DESC")
            file_rows = cur.fetchall()
            
            for row in file_rows:
                files.append(UploadedFile(
                    id=row['id'],
                    user_id=row['user_id'],
                    original_filename=row['original_filename'],
                    file_type=row.get('file_type'),
                    file_size=row.get('file_size'),
                    uploaded_at=row.get('uploaded_at'),
                    processed=row.get('processed', False),
                    processed_at=row.get('processed_at'),
                    processing_type=row.get('processing_type')
                ))
    finally:
        conn.close()
    
    # Check OneDrive configuration
    master_onedrive_token = SystemConfig.get_value('onedrive_token')
    master_onedrive_user = SystemConfig.get_value('onedrive_user_name') or 'Unknown'
    master_onedrive_updated = None
    
    # Get token expiration time if available
    token_expires = SystemConfig.get_value('onedrive_token_expires')
    if token_expires and token_expires.isdigit():
        try:
            timestamp = int(token_expires)
            master_onedrive_updated = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
        except:
            master_onedrive_updated = 'Unknown'
    
    return render_template(
        'admin/dashboard.html',
        users=users,
        files=files,
        master_onedrive_configured=(master_onedrive_token is not None),
        master_onedrive_user=master_onedrive_user,
        master_onedrive_updated=master_onedrive_updated or 'Never'
    )

# Admin user management
@app.route('/admin/user/<int:user_id>')
@admin_required
def admin_user_view(user_id):
   user = User.get(user_id)
   if not user:
       flash("User not found.")
       return redirect(url_for('admin_dashboard'))
       
   user_files = UploadedFile.get_user_files(user_id)
   return render_template('admin/user_view.html', user=user, files=user_files)

# Admin delete user
@app.route('/admin/user/<int:user_id>/delete')
@admin_required
def admin_delete_user(user_id):
   user = User.get(user_id)
   if not user:
       flash("User not found.")
       return redirect(url_for('admin_dashboard'))
   
   # Don't allow deleting yourself
   if user.id == current_user.id:
       flash("You cannot delete your own account.")
       return redirect(url_for('admin_dashboard'))
   
   # Get all files for the user
   user_files = UploadedFile.get_user_files(user_id)
   
   # Delete all files and their associated plots
   for file_record in user_files:
       # Delete the actual file
       if os.path.exists(file_record.file_path):
           os.remove(file_record.file_path)
       
       # Delete associated plots if they exist
       for plot_path in [file_record.ml_plot_path, file_record.pca_plot_path, 
                        file_record.cluster_plot_path, file_record.raw_plot_path]:
           if plot_path and os.path.exists(plot_path):
               os.remove(plot_path)
       
       # Delete the database record
       UploadedFile.delete(file_record.id)
   
   # Delete the user
   conn = get_db()
   try:
       with conn.cursor() as cur:
           cur.execute("DELETE FROM users WHERE id = %s", (user_id,))
       conn.commit()
       flash(f'User {user.username} and all their files have been deleted.')
   except Exception as e:
       conn.rollback()
       flash(f'Error deleting user: {str(e)}')
   finally:
       conn.close()
   
   return redirect(url_for('admin_dashboard'))

# Admin change user role
@app.route('/admin/user/<int:user_id>/change_role')
@admin_required
def admin_change_user_role(user_id):
   user = User.get(user_id)
   if not user:
       flash("User not found.")
       return redirect(url_for('admin_dashboard'))
   
   # Don't allow changing your own role
   if user.id == current_user.id:
       flash("You cannot change your own role.")
       return redirect(url_for('admin_dashboard'))
   
   # Toggle role between admin and user
   new_role = 'admin' if user.role == 'user' else 'user'
   
   conn = get_db()
   try:
       with conn.cursor() as cur:
           cur.execute("UPDATE users SET role = %s WHERE id = %s", (new_role, user_id))
       conn.commit()
       flash(f'User {user.username} role changed to {new_role}.')
   except Exception as e:
       conn.rollback()
       flash(f'Error changing user role: {str(e)}')
   finally:
       conn.close()
   
   return redirect(url_for('admin_user_view', user_id=user_id))

def process_uploaded_file(file_obj, user_id):
    """
    Process an uploaded file and store it on the filesystem instead of in the database
    
    Args:
        file_obj: The file object from request.files
        user_id: The ID of the current user
        
    Returns:
        UploadedFile: The created file record, or None if there was an error
    """
    try:
        # Get original filename and content type
        original_filename = file_obj.filename
        file_type = file_obj.content_type or 'text/csv'
        
        # Create user-specific filename with UUID
        unique_filename = f"{uuid.uuid4().hex}_{secure_filename(original_filename)}"
        file_path = get_user_file_path(unique_filename, user_id)
        
        # Save the file to disk instead of storing content in database
        file_obj.save(file_path)
        file_size = os.path.getsize(file_path)
        
        # Create file record in database without storing file content
        conn = get_db()
        try:
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO uploaded_files 
                    (user_id, original_filename, stored_filename, file_path, file_type, file_size)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    RETURNING id
                """, (user_id, original_filename, unique_filename, file_path, file_type, file_size))
                file_id = cur.fetchone()[0]
            conn.commit()
            
            # Create a lightweight file record object to return
            class FileRecord:
                def __init__(self, id, path):
                    self.id = id
                    self.file_path = path
                
                def get_dataframe(self):
                    """Read the file from disk and return as DataFrame"""
                    try:
                        return pd.read_csv(self.file_path)
                    except Exception as e:
                        logging.error(f"Error reading CSV from disk: {str(e)}")
                        return None
            
            file_record = FileRecord(file_id, file_path)
            logging.info(f"Created file record ID {file_id} for user {user_id}")
            return file_record
        except Exception as e:
            conn.rollback()
            logging.error(f"Database error: {str(e)}")
            # If we can't create a database record, we'll still return a temporary object
            # that can be used for processing but won't be permanently stored
            
            class TempFileRecord:
                def __init__(self, path):
                    self.id = None
                    self.file_path = path
                
                def get_dataframe(self):
                    """Read the file from disk and return as DataFrame"""
                    try:
                        return pd.read_csv(self.file_path)
                    except Exception as e:
                        logging.error(f"Error reading CSV from disk: {str(e)}")
                        return None
            
            temp_record = TempFileRecord(file_path)
            return temp_record
            
    except Exception as e:
        logging.error(f"Error processing uploaded file: {str(e)}")
        return None

def save_plot_to_db(img_base64, file_record, plot_type):
    """
    Save a base64 encoded plot image to the database
    
    Args:
        img_base64: Base64 encoded image data
        file_record: UploadedFile object to attach the image to
        plot_type: Type of plot ('ml', 'pca', 'cluster', or 'raw')
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Check if the image data includes the base64 prefix
        if isinstance(img_base64, str) and ',' in img_base64:
            # Extract only the base64 data part
            img_base64 = img_base64.split(',', 1)[1]
            
        # Convert base64 to binary data
        img_data = base64.b64decode(img_base64)
        
        # Update the appropriate field on the file record
        if plot_type == 'ml':
            file_record.ml_plot_data = img_data
        elif plot_type == 'pca':
            file_record.pca_plot_data = img_data
        elif plot_type == 'cluster':
            file_record.cluster_plot_data = img_data
        elif plot_type == 'raw':
            file_record.raw_plot_data = img_data
        
        # Save the updated record to the database
        file_record.update()
        return True
    except Exception as e:
        logging.error(f"Error saving plot image to database: {str(e)}")
        return False

def get_image_data_as_base64(file_record, plot_type):
    """
    Get image data from the database and return as base64 string
    
    Args:
        file_record: UploadedFile object containing the image data
        plot_type: Type of plot to retrieve ('ml', 'pca', 'cluster', or 'raw')
        
    Returns:
        str: Base64 encoded image data, or None if not found or on error
    """
    try:
        if plot_type == 'ml' and file_record.ml_plot_data:
            img_data = file_record.ml_plot_data
        elif plot_type == 'pca' and file_record.pca_plot_data:
            img_data = file_record.pca_plot_data
        elif plot_type == 'cluster' and file_record.cluster_plot_data:
            img_data = file_record.cluster_plot_data
        elif plot_type == 'raw' and file_record.raw_plot_data:
            img_data = file_record.raw_plot_data
        else:
            return None
            
        # Convert binary data to base64 string
        return base64.b64encode(img_data).decode('utf-8')
    except Exception as e:
        logging.error(f"Error retrieving image data: {str(e)}")
        return None

def save_processing_parameters(file_record, parameters):
    """
    Save processing parameters to the file record
    
    Args:
        file_record: UploadedFile object to update
        parameters: Dictionary of parameters to save
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        file_record.processing_parameters = parameters
        file_record.update()
        return True
    except Exception as e:
        logging.error(f"Error saving processing parameters: {str(e)}")
        return False

def save_results_metadata(file_record, metadata):
    """
    Save results metadata to the file record
    
    Args:
        file_record: UploadedFile object to update
        metadata: Dictionary of metadata to save
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        file_record.results_metadata = metadata
        file_record.update()
        return True
    except Exception as e:
        logging.error(f"Error saving results metadata: {str(e)}")
        return False

def user_can_access_file(file_id):
    """
    Check if current user can access the specified file
    
    Args:
        file_id: ID of the file to check
        
    Returns:
        bool: True if the user can access the file, False otherwise
    """
    if not current_user.is_authenticated:
        return False
    
    try:
        # Get the file record
        file_record = UploadedFile.get(int(file_id))
        if not file_record:
            return False
        
        # User has access if they own the file or are an admin
        return file_record.user_id == current_user.id or current_user.is_admin()
    except Exception as e:
        logging.error(f"Error checking file access: {str(e)}")
        return False
    


# Add this context processor for templates
@app.context_processor
def inject_session_data():
    """Make session data available to all templates"""
    # Extract only the primitive data that templates need
    session_data = {}
    
    # Check if OneDrive is configured
    onedrive_configured = get_master_onedrive_token() is not None
    
    # Get OneDrive user info for display
    if onedrive_configured:
        session_data['onedrive_user'] = {
            'display_name': SystemConfig.get_value('onedrive_user_name') or 'OneDrive User',
            'email': SystemConfig.get_value('onedrive_user_email') or 'Connected Account'
        }
    else:
        session_data['onedrive_user'] = None
    
    # Set a flag for OneDrive connection
    session_data['is_connected_to_onedrive'] = onedrive_configured
    
    # Add any other session data needed in templates
    return session_data

# Admin database view
@app.route('/admin/database')
@admin_required
def view_database():
   conn = get_db()
   users = []
   files = []
   
   try:
       with conn.cursor() as cur:
           # Get users
           cur.execute("SELECT * FROM users")
           users = cur.fetchall()
           
           # Get files
           cur.execute("SELECT * FROM uploaded_files")
           files = cur.fetchall()
   finally:
       conn.close()
   
   return render_template('admin/database.html', users=users, files=files)

# === Error Handlers ===

@app.errorhandler(404)
def page_not_found(e):
   return render_template('error.html', error_code=404, error_message="Page not found"), 404

@app.errorhandler(500)
def server_error(e):
   return render_template('error.html', error_code=500, error_message="Internal server error"), 500

@app.errorhandler(403)
def forbidden(e):
   return render_template('error.html', error_code=403, error_message="Access forbidden"), 403

def create_temp_folder_structure(folder_name):
    """Create temporary folder structure for processing OneDrive data with better error handling"""
    try:
        # Create a more unique temp ID to avoid conflicts
        temp_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
        temp_base_path = f"temp_processing/{folder_name}_{temp_id}"
        
        # Create the main temp folder
        os.makedirs(temp_base_path, exist_ok=True)
        
        # Verify the folder was created
        if not os.path.exists(temp_base_path):
            raise OSError(f"Failed to create temporary folder: {temp_base_path}")
        
        logging.info(f"✅ Created temporary folder structure: {temp_base_path}")
        return temp_base_path
    except Exception as e:
        logging.error(f"❌ Error creating temporary folder structure: {str(e)}")
        raise e

def download_onedrive_folder_contents(folder_id, local_path, access_token):
    """Download all files from OneDrive folder and ALL subfolders recursively"""
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        
        # Get all items in the OneDrive folder
        items_response = requests.get(
            f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children',
            headers=headers
        )
        
        if items_response.status_code != 200:
            items_response = requests.get(
                f'https://api.onedrive.com/v1.0/drive/items/{folder_id}/children',
                headers=headers
            )
        
        if items_response.status_code == 200:
            items = items_response.json().get('value', [])
            logging.info(f"Found {len(items)} items in OneDrive folder")
            
            downloaded_count = 0
            
            # Process each item
            for item in items:
                try:
                    if item.get('folder') is not None:
                        # It's a subfolder - create local subfolder and download recursively
                        subfolder_name = item['name']
                        local_subfolder_path = os.path.join(local_path, subfolder_name)
                        os.makedirs(local_subfolder_path, exist_ok=True)
                        
                        logging.info(f"Processing subfolder: {subfolder_name}")
                        
                        # Recursively download from subfolder
                        subfolder_count = download_onedrive_folder_contents(
                            item['id'], local_subfolder_path, access_token
                        )
                        downloaded_count += subfolder_count
                        
                    elif item.get('file') is not None:
                        # It's a file - check if it's compatible
                        file_name = item.get('name', '')
                        if (file_name.lower().endswith('.csv') or 
                            file_name.lower().endswith('.txt') or
                            file_name.lower().endswith('.xlsx') or
                            file_name.lower().endswith('.xls')):
                            
                            # Download the file
                            file_content = download_single_onedrive_file(item['id'], headers)
                            
                            if file_content:
                                file_path = os.path.join(local_path, file_name)
                                
                                # Create directory if it doesn't exist
                                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                                
                                with open(file_path, 'wb') as f:
                                    f.write(file_content)
                                
                                downloaded_count += 1
                                logging.info(f"Downloaded: {file_name}")
                            else:
                                logging.warning(f"Failed to download: {file_name}")
                        else:
                            logging.info(f"Skipping non-compatible file: {file_name}")
                            
                except Exception as e:
                    logging.error(f"Error processing item {item.get('name', 'unknown')}: {str(e)}")
                    continue
            
            logging.info(f"Downloaded {downloaded_count} files from folder and subfolders")
            return downloaded_count > 0
        else:
            logging.error(f"Failed to get folder contents: {items_response.status_code}")
            return False
            
    except Exception as e:
        logging.error(f"Error downloading OneDrive folder contents: {str(e)}")
        return False

def download_single_onedrive_file(file_id, headers):
    """Download a single file from OneDrive and return its content"""
    try:
        # Try different endpoints to get the file content
        for endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}/content"
        ]:
            try:
                response = requests.get(endpoint, headers=headers)
                if response.status_code == 200:
                    return response.content
                else:
                    logging.warning(f"Download failed with {endpoint}: {response.status_code}")
            except Exception as e:
                logging.error(f"Error with {endpoint}: {str(e)}")
                continue
        
        # If direct download fails, try getting download URL first
        for info_endpoint in [
            f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}",
            f"https://api.onedrive.com/v1.0/drive/items/{file_id}"
        ]:
            try:
                info_response = requests.get(info_endpoint, headers=headers)
                if info_response.status_code == 200:
                    file_info = info_response.json()
                    download_url = file_info.get('@microsoft.graph.downloadUrl') or file_info.get('@content.downloadUrl')
                    
                    if download_url:
                        download_response = requests.get(download_url)
                        if download_response.status_code == 200:
                            return download_response.content
            except Exception as e:
                logging.error(f"Error with {info_endpoint}: {str(e)}")
                continue
        
        return None
    except Exception as e:
        logging.error(f"Error downloading single file: {str(e)}")
        return None

def upload_masterdata_to_onedrive(local_masterdata_path, folder_name, access_token):
    """Upload the generated master data file back to OneDrive"""
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        username = current_user.username
        
        # Read the master data file
        with open(local_masterdata_path, 'rb') as f:
            file_content = f.read()
        
        # Find user folder in OneDrive
        user_folder_id = None
        response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root/children', headers=headers)
        
        if response.status_code == 200:
            root_items = response.json().get('value', [])
            for item in root_items:
                if item.get('name') == username and item.get('folder'):
                    user_folder_id = item.get('id')
                    break
        
        if not user_folder_id:
            logging.error("User folder not found in OneDrive")
            return False
        
        # Create or find processed_data folder
        results_folder_id = None
        folder_response = requests.get(f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children', headers=headers)
        
        if folder_response.status_code == 200:
            folder_items = folder_response.json().get('value', [])
            for item in folder_items:
                if item.get('name') == 'processed_data' and item.get('folder'):
                    results_folder_id = item.get('id')
                    break
        
        # Create processed_data folder if it doesn't exist
        if not results_folder_id:
            create_folder_data = {
                'name': 'processed_data',
                'folder': {},
                '@microsoft.graph.conflictBehavior': 'rename'
            }
            
            create_response = requests.post(
                f'https://graph.microsoft.com/v1.0/me/drive/items/{user_folder_id}/children',
                headers={**headers, 'Content-Type': 'application/json'},
                json=create_folder_data
            )
            
            if create_response.status_code in [200, 201]:
                results_folder_id = create_response.json().get('id')
        
        if not results_folder_id:
            logging.error("Could not create or find processed_data folder")
            return False
        
        # Create filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{folder_name}_Masterdata_{timestamp}.csv"
        
        # Upload the master data file
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'text/csv'
        }
        
        upload_response = requests.put(
            f'https://graph.microsoft.com/v1.0/me/drive/items/{results_folder_id}:/{filename}:/content',
            headers=upload_headers,
            data=file_content
        )
        
        if upload_response.status_code in [200, 201]:
            logging.info(f"Successfully uploaded master data to OneDrive: {filename}")
            return True
        else:
            logging.error(f"Failed to upload master data: {upload_response.status_code}")
            return False
            
    except Exception as e:
        logging.error(f"Error uploading master data to OneDrive: {str(e)}")
        return False

def cleanup_temp_folder(temp_path):
    """Clean up temporary folder and all its contents with better error handling"""
    try:
        if temp_path and os.path.exists(temp_path):
            # Wait a moment to ensure all file handles are closed
            import time
            time.sleep(0.5)
            
            # Try to remove the folder
            shutil.rmtree(temp_path, ignore_errors=True)
            
            # Verify cleanup
            if os.path.exists(temp_path):
                logging.warning(f"⚠️ Temporary folder still exists after cleanup attempt: {temp_path}")
                # Try alternative cleanup method
                try:
                    import subprocess
                    if os.name == 'nt':  # Windows
                        subprocess.run(['rmdir', '/s', '/q', temp_path], shell=True, check=False)
                    else:  # Unix-like
                        subprocess.run(['rm', '-rf', temp_path], check=False)
                except Exception as e:
                    logging.warning(f"Alternative cleanup method also failed: {str(e)}")
            else:
                logging.info(f"✅ Successfully cleaned up temporary folder: {temp_path}")
        else:
            logging.info(f"ℹ️ No cleanup needed - folder doesn't exist: {temp_path}")
    except Exception as e:
        logging.error(f"❌ Error cleaning up temporary folder: {str(e)}")

# Add this route for health check
@app.route('/health')
def health_check():
    """Health check endpoint"""
    try:
        # Check if temp directories exist
        temp_dirs_status = {}
        for temp_dir in ['temp_downloads', 'temp_processing', 'uploads']:
            dir_path = os.path.join(os.getcwd(), temp_dir)
            temp_dirs_status[temp_dir] = os.path.exists(dir_path)
        
        # Check OneDrive connection
        onedrive_status = get_master_onedrive_token() is not None
        
        return jsonify({
            'status': 'healthy',
            'timestamp': datetime.now().isoformat(),
            'temp_directories': temp_dirs_status,
            'onedrive_configured': onedrive_status
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'error': str(e)
        }), 500

if __name__ == "__main__":
    # init_db()
    # ensure_source_column_exists()
    initialize_app_directories_and_cleanup()
    app.run(debug=True)