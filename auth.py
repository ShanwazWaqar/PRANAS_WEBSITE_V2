from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash
from models import User
import os
from functools import wraps
import psycopg

auth = Blueprint('auth', __name__)

# Admin required decorator
def admin_required(f):
    @wraps(f)
    @login_required
    def decorated_function(*args, **kwargs):
        if not current_user.is_admin():
            flash('You need admin privileges to access this page')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

@auth.route('/login', methods=['GET', 'POST'])
def login():
    # Get the next parameter if available (where to redirect after login)
    next_page = request.args.get('next')
    
    if current_user.is_authenticated:
        # If user is already logged in, redirect to the next page or home page
        return redirect(next_page or url_for('index'))
    
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        remember = True if request.form.get('remember') else False
        
        user = User.get_by_email(email)
        
        if not user or not user.check_password(password):
            flash('Please check your login details and try again.', 'error')
            return redirect(url_for('auth.login', next=next_page))

        login_user(user, remember=remember)
        
        # Redirect to the requested page or home page (index)
        return redirect(next_page or url_for('index'))
    
    return render_template('login.html', next=next_page)

@auth.route('/signup', methods=['GET', 'POST'])
def signup():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
        
    if request.method == 'POST':
        email = request.form.get('email')
        username = request.form.get('username')
        password = request.form.get('password')
        
        if not email or not username or not password:
            flash('All fields are required.')
            return redirect(url_for('auth.signup'))
        
        # Check if user already exists
        try:
            user_email = User.get_by_email(email)
            if user_email:
                flash('Email already exists.')
                return redirect(url_for('auth.signup'))
            
            # Create new user with the create method from the User class
            new_user = User.create(username, email, password)
            
            # Create user directory for uploads
            user_upload_dir = os.path.join('uploads', str(new_user.id))
            os.makedirs(user_upload_dir, exist_ok=True)
            
            # Create user directories for plots
            for plot_type in ['ML_Plots', 'pca_plots', 'cluster_plots', 'raw_plots']:
                plot_dir = os.path.join('Plots', str(new_user.id), plot_type)
                os.makedirs(plot_dir, exist_ok=True)
            
            flash('Account created successfully!')
            return redirect(url_for('auth.login'))
        except psycopg.Error as e:
            print(f"Database error during signup: {str(e)}")
            flash(f'Database error: {str(e)}')
            return redirect(url_for('auth.signup'))
        except Exception as e:
            print(f"Unexpected error during signup: {str(e)}")
            flash('An unexpected error occurred. Please try again.')
            return redirect(url_for('auth.signup'))
        
    return render_template('signup.html')

@auth.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out successfully.', 'success')
    return redirect(url_for('index'))

@auth.route('/register', methods=['GET', 'POST'])
def register():
    """Alias for signup route"""
    return redirect(url_for('auth.signup'))