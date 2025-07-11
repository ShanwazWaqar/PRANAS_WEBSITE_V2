import msal
import requests
import os
import json
from urllib.parse import quote
from flask import session, redirect, url_for, request, render_template, current_app, flash
import logging

class OneDriveHandler:
    def __init__(self, app=None):
        self.app = None
        self.client_id = None
        self.client_secret = None
        self.authority = None
        self.token = None
        self.scopes = ['.default']  # Use .default for application permissions
        self.logger = logging.getLogger(__name__)  # Use a regular logger instead of app.logger
        
        if app:
            self.init_app(app)
    
    def init_app(self, app):
        self.app = app
        # Load configuration from app with fallback values
        self.client_id = app.config.get('MICROSOFT_APP_ID')
        self.client_secret = app.config.get('MICROSOFT_APP_PASSWORD')
        self.authority = f"https://login.microsoftonline.com/{app.config.get('MICROSOFT_TENANT_ID', 'common')}"
        
        # Add debugging info without using current_app
        self.logger.info(f"OneDrive initialized with client_id: {self.client_id is not None}, "
                         f"client_secret: {self.client_secret is not None}")
        
        # IMPORTANT: Don't call get_app_token() here at startup
        # Instead, register a function to run after the app is fully initialized
        @app.before_first_request
        def initialize_onedrive():
            with app.app_context():
                self.get_app_token()
    
    def get_app_token(self):
        """Get an access token for application permissions (no user login required)"""
        try:
            app = msal.ConfidentialClientApplication(
                self.client_id, authority=self.authority,
                client_credential=self.client_secret
            )
            
            # Get token using client credentials (no user required)
            result = app.acquire_token_for_client(scopes=self.scopes)
            
            if "access_token" in result:
                self.logger.info("Successfully acquired application token")
                self.token = result["access_token"]
                return True
            
            self.logger.error(f"Failed to get app token. Error: {result.get('error')}, Description: {result.get('error_description')}")
            return False
        except Exception as e:
            self.logger.error(f"Exception getting app token: {str(e)}")
            return False
    
    def get_files_from_folder(self, folder_path=None):
        """Get files from a specific folder in OneDrive"""
        # Make sure we have a token
        if not self.token:
            if not self.get_app_token():
                return None
            
        # If folder path is provided, encode it properly for the URL
        if folder_path:
            # Remove leading slash if present
            if folder_path.startswith('/'):
                folder_path = folder_path[1:]
            path = f"/drives/{self.get_main_drive_id()}/root:/{quote(folder_path)}:/children"
        else:
            path = f"/drives/{self.get_main_drive_id()}/root/children"
            
        endpoint = f"https://graph.microsoft.com/v1.0{path}"
        headers = {'Authorization': f'Bearer {self.token}'}
        
        try:
            response = requests.get(endpoint, headers=headers)
            if response.status_code == 401:
                # Token expired, try to refresh
                if self.get_app_token():
                    # Try again with the new token
                    headers = {'Authorization': f'Bearer {self.token}'}
                    response = requests.get(endpoint, headers=headers)
                    
            response.raise_for_status()  # Raise an exception for 4XX/5XX status codes
            return response.json().get('value', [])
        except requests.RequestException as e:
            self.logger.error(f"Error fetching OneDrive files: {str(e)}")
            return None
    
    def download_file(self, file_id):
        """Download a file from OneDrive by ID"""
        # Make sure we have a token
        if not self.token:
            if not self.get_app_token():
                return None
            
        endpoint = f"https://graph.microsoft.com/v1.0/drives/{self.get_main_drive_id()}/items/{file_id}/content"
        headers = {'Authorization': f'Bearer {self.token}'}
        
        try:
            response = requests.get(endpoint, headers=headers)
            if response.status_code == 401:
                # Token expired, try to refresh
                if self.get_app_token():
                    # Try again with the new token
                    headers = {'Authorization': f'Bearer {self.token}'}
                    response = requests.get(endpoint, headers=headers)
                    
            response.raise_for_status()
            return response.content
        except requests.RequestException as e:
            self.logger.error(f"Error downloading file: {str(e)}")
            return None
    
    def get_file_metadata(self, file_id):
        """Get metadata for a specific file"""
        # Make sure we have a token
        if not self.token:
            if not self.get_app_token():
                return None
            
        endpoint = f"https://graph.microsoft.com/v1.0/drives/{self.get_main_drive_id()}/items/{file_id}"
        headers = {'Authorization': f'Bearer {self.token}'}
        
        try:
            response = requests.get(endpoint, headers=headers)
            if response.status_code == 401:
                # Token expired, try to refresh
                if self.get_app_token():
                    # Try again with the new token
                    headers = {'Authorization': f'Bearer {self.token}'}
                    response = requests.get(endpoint, headers=headers)
                    
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            self.logger.error(f"Error fetching file metadata: {str(e)}")
            return None
    
    def get_main_drive_id(self):
        """Get the ID of the primary drive"""
        if not self.token:
            if not self.get_app_token():
                return None
                
        endpoint = "https://graph.microsoft.com/v1.0/drives"
        headers = {'Authorization': f'Bearer {self.token}'}
        
        try:
            response = requests.get(endpoint, headers=headers)
            if response.status_code == 401:
                # Token expired, try to refresh
                if self.get_app_token():
                    # Try again with the new token
                    headers = {'Authorization': f'Bearer {self.token}'}
                    response = requests.get(endpoint, headers=headers)
                    
            response.raise_for_status()
            drives = response.json().get('value', [])
            if drives:
                # Return the first drive's ID (usually the main one)
                return drives[0]['id']
            return None
        except requests.RequestException as e:
            self.logger.error(f"Error fetching drives: {str(e)}")
            return None
        
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