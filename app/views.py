from flask import Blueprint, render_template, redirect, url_for, flash, session

views_bp = Blueprint('views', __name__)

@views_bp.route('/dash')
def dash():
    if 'username' not in session:
        flash("Please log in first.", "warning")
        return redirect(url_for('auth.login'))
    return render_template('dash.html')


@views_bp.route('/role_redirect')
def role_redirect():
    if 'username' not in session or 'role' not in session:
        flash("Please log in first.", "warning")
        return redirect(url_for('auth.login'))

    role = session['role']

    if role == "manager":
        return redirect(url_for('manager.manager'))
    elif role == "employee":
        return redirect(url_for('employee.employee'))
    elif role == "lead":
        return redirect(url_for('lead.edit_shift'))  # Assuming lead is also in views
    else:
        flash("Unknown role.", "danger")
        return redirect(url_for('auth.login'))


@views_bp.route('/logout')
def logout():
    session.clear()
    flash("You have been logged out successfully.")
    return redirect(url_for('auth.login'))
