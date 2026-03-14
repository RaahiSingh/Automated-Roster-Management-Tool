from flask import Blueprint, render_template, request, redirect, flash, url_for, session, current_app
from .models import User
from . import db, mail, serializer
from flask_mail import Message

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/')
def index():
    return redirect('/login')

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        input_password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and user.password == input_password:
            session['username'] = user.username
            session['role'] = user.role
            flash("Login successful.", "success")
            return redirect('/dash')
        flash("Invalid credentials.", "danger")
        return "Invalid credentials", 401
    return render_template('login.html')

@auth_bp.route('/logout')
def logout():
    session.clear()
    flash("You have been logged out successfully.")
    return redirect(url_for('auth.login'))

@auth_bp.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        user = User.query.filter_by(username=username, email=email).first()
        if user:
            token = serializer.dumps(user.email, salt='reset-salt')
            reset_link = url_for('auth.reset_password', token=token, _external=True)
            send_email(user.email, reset_link)
            flash("Password reset link sent.", "info")
        else:
            flash("User not found.", "danger")
        return redirect(url_for('auth.forgot_password'))
    return render_template('forgot_password.html')

@auth_bp.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    try:
        email = serializer.loads(token, salt='reset-salt', max_age=3600)
        user = User.query.filter_by(email=email).first()
        if not user:
            flash("Invalid user.", "danger")
            return redirect(url_for('auth.forgot_password'))
    except:
        flash("Invalid or expired link.", "danger")
        return redirect(url_for('auth.forgot_password'))

    if request.method == 'POST':
        new_password = request.form.get('password')
        user.password = new_password
        db.session.commit()
        flash("Password updated.", "success")
        return redirect(url_for('auth.login'))
    return render_template('reset_password.html')

def send_email(to, link):
    try:
        msg = Message("Password Reset", recipients=[to])
        msg.body = f"Click to reset password: {link}"
        mail.send(msg)
    except Exception as e:
        print("Email error:", e)


