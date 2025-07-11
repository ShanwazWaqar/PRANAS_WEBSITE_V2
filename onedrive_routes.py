from flask import Blueprint, redirect, url_for, render_template, render_template_string, session, request, current_app, flash
from flask_login import login_required, current_user
import logging

# Create blueprint with NO prefix - this is crucial
onedrive_bp = Blueprint('onedrive_integration', __name__, url_prefix='')

@onedrive_bp.route('/onedrive/list_files')
@login_required
def list_files():
    """List files from OneDrive without requiring user login"""
    try:
        # Get OneDrive handler from app extensions
        onedrive_handler = current_app.extensions.get('onedrive_handler')
        
        if not onedrive_handler:
            flash("OneDrive integration is not properly configured", "error")
            return redirect(url_for('dashboard'))
        
        # Get folder path from query parameter, default to root
        folder_path = request.args.get('folder', '')
        
        # Get files from the folder
        files = onedrive_handler.get_files_from_folder(folder_path)
        
        if files is None:
            flash("Error fetching files from OneDrive", "error")
            return redirect(url_for('dashboard'))
        
        # Create a simple template string for fallback
        template_string = """
        {% extends "base.html" %}
        {% block content %}
        <div class="container mt-4">
            <h1 class="mb-4">OneDrive Files</h1>
            <nav aria-label="breadcrumb">
                <ol class="breadcrumb">
                    <li class="breadcrumb-item"><a href="{{ url_for('dashboard') }}">Dashboard</a></li>
                    <li class="breadcrumb-item active" aria-current="page">OneDrive Files</li>
                </ol>
            </nav>

            <div class="card">
                <div class="card-header">
                    <h5 class="mb-0">Folder: {{ folder_path or 'Root' }}</h5>
                </div>
                <div class="card-body">
                    {% if files %}
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead class="thead-light">
                                    <tr>
                                        <th>Name</th>
                                        <th>Type</th>
                                        <th>Actions</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {% for file in files %}
                                        <tr>
                                            <td>{{ file.name }}</td>
                                            <td>
                                                {% if 'folder' in file %}
                                                    Folder
                                                {% else %}
                                                    File
                                                {% endif %}
                                            </td>
                                            <td>
                                                {% if 'folder' in file %}
                                                    <a href="{{ url_for('onedrive_files', folder=folder_path + '/' + file.name if folder_path else file.name) }}" class="btn btn-sm btn-outline-primary">
                                                        Open Folder
                                                    </a>
                                                {% elif 'file' in file %}
                                                    <a href="{{ url_for('onedrive_integration.download_file', file_id=file.id) }}" class="btn btn-sm btn-outline-success">
                                                        Download
                                                    </a>
                                                {% endif %}
                                            </td>
                                        </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    {% else %}
                        <div class="alert alert-info">
                            No files found in this folder.
                        </div>
                    {% endif %}
                </div>
                <div class="card-footer">
                    <a href="{{ url_for('dashboard') }}" class="btn btn-secondary">Back to Dashboard</a>
                </div>
            </div>
        </div>
        {% endblock %}
        """
        
        # Try to use the template file, with fallback to template string
        try:
            return render_template('onedrive/list.html', files=files, folder_path=folder_path)
        except Exception as e:
            current_app.logger.warning(f"Template not found, using template string: {str(e)}")
            return render_template_string(template_string, files=files, folder_path=folder_path)
            
    except Exception as e:
        current_app.logger.error(f"Error listing OneDrive files: {str(e)}")
        flash(f"Error listing OneDrive files: {str(e)}", "error")
        return redirect(url_for('dashboard'))

@onedrive_bp.route('/onedrive/download_file/<file_id>')
@login_required
def download_file(file_id):
    """Download a file from OneDrive"""
    try:
        # Get OneDrive handler
        onedrive_handler = current_app.extensions.get('onedrive_handler')
        
        if not onedrive_handler:
            flash("OneDrive integration is not properly configured", "error")
            return redirect(url_for('dashboard'))
        
        # Get file info first to get the filename
        file_info = onedrive_handler.get_file_metadata(file_id)
        if not file_info:
            flash("Error getting file information", "error")
            return redirect(url_for('onedrive_files'))
        
        # Download the file content
        file_content = onedrive_handler.download_file(file_id)
        if not file_content:
            flash("Error downloading file", "error")
            return redirect(url_for('onedrive_files'))
        
        # Send the file to the user
        from io import BytesIO
        from flask import send_file
        
        filename = file_info.get('name', f"file_{file_id}")
        mimetype = file_info.get('file', {}).get('mimeType', 'application/octet-stream')
        
        return send_file(
            BytesIO(file_content),
            mimetype=mimetype,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        current_app.logger.error(f"Error downloading file: {str(e)}")
        flash(f"Error downloading file: {str(e)}", "error")
        return redirect(url_for('onedrive_files'))