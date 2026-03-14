from flask import Blueprint, request, render_template, flash, redirect, url_for, session, send_file
import pandas as pd
import os

employee_bp = Blueprint('employee', __name__)

@employee_bp.route('/employee', methods=['GET', 'POST'])
def employee():
    shifts = {"OC": [], "G": [], "E1": [], "E2": [], "N": [], "L": []}
    name = signum = month = year = None

    if request.method == 'POST':
        name = request.form.get('name')
        signum = request.form.get('signum')
        month = request.form.get('month')
        year = request.form.get('year')

        # Set session values to use later
        session['name'] = name
        session['signum'] = signum
        session['month'] = month
        session['year'] = year

        if all([name, signum, month, year]):
            try:
                base_dir = os.path.abspath(os.path.dirname(__file__))
                filename = f'Shift_Plan_Sample_{month}_{year}_1.1.xlsx'
                file_path = os.path.join(base_dir, '..', filename)

                df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
                df = pd.read_excel(file_path, skiprows=3, header=None, engine='openpyxl')

                for index, row in df.iterrows():
                    if str(row[1]).strip().upper() == signum.strip().upper() and str(row[4]).strip().upper() == name.strip().upper():
                        for col in range(6, len(row)):
                            shift = str(row[col]).strip().upper()
                            date = df_raw.iloc[2, col]
                            day = df_raw.iloc[1, col]
                            if shift in shifts:
                                shifts[shift].append(f"{date} {month.title()}, ({day})")
                        break
                else:
                    flash(f"No matching record found for {name} ({signum}).", "warning")

            except FileNotFoundError:
                flash(f"Shift file for {month} {year} not found.", "danger")
            except Exception as e:
                flash(f"Error reading shift file: {e}", "danger")
                

    return render_template(
        'employee.html',
        name=name,
        signum=signum,
        month=month,
        year=year,
        month_year=f"{month.title()} {year}" if month and year else "",
        shifts=shifts
    )

@employee_bp.route('/employee/download-excel', methods=['GET'])
def download_original_excel():
    month = session.get('month')
    year = session.get('year')

    if not month or not year:
        flash("Session expired or missing data. Please fill the form again.", "warning")
        return redirect(url_for('employee.employee'))

    try:
        base_dir = os.path.abspath(os.path.dirname(__file__))
        filename = f'Shift_Plan_Sample_{month}_{year}_1.1.xlsx'
        file_path = os.path.join(base_dir, '..', filename)

        if not os.path.exists(file_path):
            flash(f"Shift file for {month} {year} not found.", "danger")
            return redirect(url_for('employee.employee'))

        return send_file(file_path, download_name=filename, as_attachment=True)

    except Exception as e:
        flash(f"Error downloading file: {e}", "danger")
        return redirect(url_for('employee.employee'))
   
