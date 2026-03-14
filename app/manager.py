from datetime import datetime
import glob
from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app,send_file
from .utils import fill_excel_template
import os
import pandas as pd
import calendar
import tempfile
from calendar import day_abbr, weekday, monthrange
manager_bp = Blueprint('manager', __name__)

@manager_bp.route('/manager', methods=['GET', 'POST'])
def manager():
    if request.method == 'POST':
        if 'team' in request.form and 'name1' not in request.form:
            return render_template('manager.html', team_name=request.form.get('team'))

        elif 'name1' in request.form:
            team_name = request.form.get('team_name')

            members = [{
                'name': request.form.get(f'name{i}'),
                'signum': request.form.get(f'signum{i}'),
                'location': request.form.get(f'location{i}'),
                'function': request.form.get(f'function{i}')
            } for i in range(1, 13)]

            template_path = os.path.join(current_app.root_path, 'static', 'template.xlsx')
            output_filename = f'template_{team_name}_June_2025_1.1.xlsx'
            output_path = os.path.join(current_app.root_path, 'static', output_filename)

            fill_excel_template(template_path, output_path, members)

            return redirect(url_for('manager.manager_overview', filename=output_filename))

    return render_template('manager.html')

@manager_bp.route('/manager/overview')
def manager_overview():
    filename = request.args.get('filename')
    if not filename:
        flash("No file to display.", "warning")
        return redirect(url_for('manager.manager'))
    return render_template('manager_overview.html', filename=filename)

@manager_bp.route('/manager/analytics', methods=['GET', 'POST'])
def manager_analytics():
  

    filters = {
        'employee': request.form.get('employee') if request.method == 'POST' else '',
        'group': request.form.get('group') if request.method == 'POST' else '',
        'shift_type': request.form.get('shift_type') if request.method == 'POST' else '',
        'month': request.form.get('month') if request.method == 'POST' else '',
        'year': request.form.get('year') if request.method == 'POST' else ''
    }

    if filters['month'] and filters['year']:
        pattern = f"Shift_Plan_Sample_{filters['month']}_{filters['year']}*.xlsx"
    else:
        pattern = "Shift_Plan_*.xlsx"

    # Find latest file
    excel_files = glob.glob(os.path.join(current_app.root_path, 'static', 'Shift_Plan_*.xlsx'))
    if not excel_files:
        flash("Excel data not found.", "danger")
        return redirect(url_for('manager.manager'))

    latest_file = max(excel_files, key=os.path.getmtime)

    try:
        # 1. Get Month from Cell A1
        raw_title = pd.read_excel(latest_file, header=None, nrows=1).iloc[0, 0]
        month = raw_title.split("-")[-1].strip() if isinstance(raw_title, str) else "Month Not Found"

        # 2. Read actual data from row 3 onward
        df = pd.read_excel(latest_file, skiprows=2)

        # 3. Assign column names
        fixed_columns = ["function", "signum", "lc", "oc", "name", "location", "working_days"]
        dynamic_columns = [f"day_{i}" for i in range(df.shape[1] - len(fixed_columns))]
        # Drop rows where function or name is null or contains shift headers like 'M', 'E1', etc.
         # remove empty or nan-like names

        df.columns = fixed_columns + dynamic_columns
        
        df = df[~df['function'].isin(['M', 'E1', 'E2', 'OC', 'L', 'G', 'N', 'H', ''])]  # remove extra rows
        df = df[~df['name'].isin(['nan', ''])] 

        # 4. Rename day columns to "1 (Sun)", "2 (Mon)", etc.
        if filters['month'] and filters['year']:
            capitalized_month = filters['month'].capitalize()
            if capitalized_month not in calendar.month_name:
                flash(f"Invalid month: {capitalized_month}", "danger")
                return redirect(url_for('manager.manager'))
        
            month_num = list(calendar.month_name).index(capitalized_month)
            year = int(filters['year'])
        else:
            now = datetime.now()
            month_num = now.month
            year = now.year
        

        num_days = calendar.monthrange(year, month_num)[1]
        pretty_day_columns = [
            f"{day} ({day_abbr[weekday(year, month_num, day)]})" for day in range(1, num_days + 1)
        ]

        # Rename columns for day-wise headers
        rename_map = dict(zip(dynamic_columns, pretty_day_columns))
        df.rename(columns=rename_map, inplace=True)

        # Replace all NaN with empty string
        df.fillna("", inplace=True)  # ✅ This line ensures "nan" won't appear in HTML

        # 5. Filter data
        df['name'] = df['name'].astype(str).str.strip()
        df = df[df['name'] != ""]


        if filters['employee']:
            df = df[df['name'].str.contains(filters['employee'], case=False, na=False)]
        if filters['group']:
            df = df[df['function'].str.contains(filters['group'], case=False, na=False)]
        if filters['shift_type']:
            df = df[df[list(rename_map.values())].apply(
                lambda row: row.astype(str).str.contains(filters['shift_type'], case=False, na=False).any(), axis=1
            )]

        # 6. Count shifts
        df[list(rename_map.values())] = df[list(rename_map.values())].astype(str)
        shift_counts = pd.Series(df[list(rename_map.values())].values.ravel()).value_counts().to_dict()

        analytics_data = {
            'summary': {
                'month': filters['month'].capitalize() if filters['month'] else "Unknown",
                'year': filters['year'] or datetime.now().year,
                'total_employees': df['name'].nunique(),
                'total_leaves': shift_counts.get('L', 0),
                'general_shifts': shift_counts.get('G', 0),
                'night_shifts': shift_counts.get('N', 0),
                'e1_shifts': shift_counts.get('E1', 0),
                'e2_shifts': shift_counts.get('E2', 0)
            }
        }

        # 7. Prepare table for display
        table_columns = ["function", "signum", "name", "location"] + list(rename_map.values())
        


        table_rows = df[table_columns].to_dict(orient='records')

        return render_template(
            'manager_analytics.html',
            filters=filters,
            data=analytics_data,
            month=month,
            table_rows=table_rows,
            columns=table_columns
        )

    except Exception as e:
        flash(f"Error processing Excel: {e}", "danger")
        return redirect(url_for('manager.manager'))
    
@manager_bp.route('/manager/download-summary-excel', methods=['POST'])
def download_summary_excel():
    filters = {
        'employee': request.form.get('employee', ''),
        'group': request.form.get('group', ''),
        'shift_type': request.form.get('shift_type', ''),
        'month': request.form.get('month', ''),
        'year': request.form.get('year', '')
    }

    try:
        pattern = f"Shift_Plan_Sample_{filters['month']}_{filters['year']}*.xlsx" if filters['month'] and filters['year'] else "Shift_Plan_*.xlsx"
        excel_files = glob.glob(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', pattern)))
        if not excel_files:
            flash("No file found for this month/year.", "danger")
            return redirect(url_for('manager.manager_analytics'))

        latest_file = max(excel_files, key=os.path.getmtime)

        df = pd.read_excel(latest_file, skiprows=2)
        fixed_columns = ["function", "signum", "lc", "oc", "name", "location", "working_days"]
        dynamic_columns = [f"day_{i}" for i in range(df.shape[1] - len(fixed_columns))]
        df.columns = fixed_columns + dynamic_columns

        # Filtering
        df = df[~df['function'].isin(['M', 'E1', 'E2', 'OC', 'L', 'G', 'N', 'H', ''])]
        df['name'] = df['name'].astype(str).str.strip()
        df = df[df['name'] != ""]

        if filters['employee']:
            df = df[df['name'].str.contains(filters['employee'], case=False, na=False)]
        if filters['group']:
            df = df[df['function'].str.contains(filters['group'], case=False, na=False)]
        if filters['shift_type']:
            df[dynamic_columns] = df[dynamic_columns].astype(str)
            df = df[df[dynamic_columns].apply(
                lambda row: row.str.contains(filters['shift_type'], case=False, na=False).any(), axis=1
            )]

        # Create day column names like "1 (Sun)"
        if filters['month'] and filters['year']:
            month_num = list(calendar.month_name).index(filters['month'].capitalize())
            year = int(filters['year'])
        else:
            now = datetime.now()
            month_num = now.month
            year = now.year

        num_days = calendar.monthrange(year, month_num)[1]
        pretty_day_columns = [
            f"{day} ({calendar.day_abbr[calendar.weekday(year, month_num, day)]})"
            for day in range(1, num_days + 1)
        ]
        rename_map = dict(zip(dynamic_columns[:len(pretty_day_columns)], pretty_day_columns))

        df.rename(columns=rename_map, inplace=True)
        day_columns = list(rename_map.values())

        # Summary
        shift_counts = pd.Series(df[day_columns].values.ravel()).value_counts().to_dict()
        summary_df = pd.DataFrame([{
            'Month': filters['month'].capitalize() if filters['month'] else "Unknown",
            'Year': filters['year'] or datetime.now().year,
            'Total Employees': df['name'].nunique(),
            'Leaves (L)': shift_counts.get('L', 0),
            'General (G)': shift_counts.get('G', 0),
            'Night (N)': shift_counts.get('N', 0),
            'E1': shift_counts.get('E1', 0),
            'E2': shift_counts.get('E2', 0)
        }])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
                summary_df.to_excel(writer, index=False, sheet_name='Summary')
                df_to_save = df[["function", "signum", "name", "location"] + day_columns]
                df_to_save.to_excel(writer, index=False, sheet_name='Filtered Roster')

            tmp.seek(0)
            return send_file(tmp.name, as_attachment=True, download_name="Shift_Summary_With_Roster.xlsx")

    except Exception as e:
        flash(f"Error generating summary: {e}", "danger")
        return redirect(url_for('manager.manager_analytics'))


