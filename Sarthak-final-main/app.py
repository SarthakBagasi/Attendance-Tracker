from flask import Flask, render_template, redirect, url_for, send_file, make_response, request, flash, session, jsonify
from models import db, Employee, ShiftType, ShiftRota, Attendance, ExceptionReport
from datetime import date, timedelta, time, datetime
import pandas as pd
import io
from jinja2 import Template
from xhtml2pdf import pisa
import csv
import os

def calculate_duration(time_in, time_out):
    """
    Calculate duration between two time objects.
    Returns duration in HH:MM format.
    """
    if not time_in or not time_out:
        return 'N/A'
    
    # Convert time objects to datetime objects for the same date
    today = date.today()
    datetime_in = datetime.combine(today, time_in)
    datetime_out = datetime.combine(today, time_out)
    
    # If time_out is earlier than time_in, assume it's the next day
    if datetime_out < datetime_in:
        datetime_out = datetime.combine(today + timedelta(days=1), time_out)
    
    # Calculate duration
    duration = datetime_out - datetime_in
    
    # Convert to hours and minutes
    total_minutes = int(duration.total_seconds() / 60)
    hours = total_minutes // 60
    minutes = total_minutes % 60
    
    return f"{hours:02d}:{minutes:02d}"

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'replace-this-with-a-strong-random-secret-key')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///attendance.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)  # Session expires after 8 hours

db.init_app(app)

SHIFT_CODES = [
    ('M', 'Morning'),
    ('E', 'Evening'),
    ('N', 'Night'),
    ('G', 'General'),
    ('Off', 'Off'),
    ('Leave', 'Leave')
]

SHIFT_START = {
    'M': time(7, 0),
    'E': time(15, 0),
    'N': time(23, 0),
    'G': time(9, 0)
}

LATE_THRESHOLD = timedelta(minutes=15)

ADMIN_PASSWORD = 'admin123'  # Change this in production!

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    # If already logged in, check for redirect parameter
    if session.get('admin'):
        redirect_to = request.args.get('redirect', 'admin_employees')
        if redirect_to == 'generate_rota':
            return redirect(url_for('generate_rota'))
        else:
            flash('Already logged in as admin.', 'info')
            return redirect(url_for('admin_employees'))
    
    if request.method == 'POST':
        password = request.form.get('password')
        if password == ADMIN_PASSWORD:
            session.permanent = True  # Make session permanent
            session['admin'] = True
            session['login_time'] = datetime.now().isoformat()
            session['user_agent'] = request.headers.get('User-Agent', 'Unknown')
            flash('Successfully logged in as admin.', 'success')
            # Check if there's a redirect parameter for rota generation
            redirect_to = request.args.get('redirect', 'admin_employees')
            if redirect_to == 'generate_rota':
                return redirect(url_for('generate_rota'))
            else:
                return redirect(url_for('admin_employees'))
        else:
            flash('Incorrect password. Please try again.', 'danger')
    return render_template('admin_login.html')

@app.route('/admin/logout')
def admin_logout():
    if session.get('admin'):
        session.clear()  # Clear all session data
        flash('Successfully logged out.', 'success')
    else:
        flash('You were not logged in.', 'info')
    return redirect(url_for('admin_login'))

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('admin'):
            flash('Admin login required.', 'danger')
            return redirect(url_for('admin_login'))
        
        # Check session timeout (optional additional security)
        login_time = session.get('login_time')
        if login_time:
            try:
                login_datetime = datetime.fromisoformat(login_time)
                if datetime.now() - login_datetime > timedelta(hours=8):
                    session.clear()
                    flash('Session expired. Please login again.', 'warning')
                    return redirect(url_for('admin_login'))
            except ValueError:
                pass  # If login_time is invalid, continue
        
        return f(*args, **kwargs)
    return decorated_function

def is_admin_logged_in():
    """Helper function to check if admin is logged in"""
    return session.get('admin', False)

def get_session_info():
    """Helper function to get session information"""
    if session.get('admin'):
        login_time = session.get('login_time')
        if login_time:
            try:
                login_datetime = datetime.fromisoformat(login_time)
                return {
                    'logged_in': True,
                    'login_time': login_datetime.strftime('%Y-%m-%d %H:%M:%S'),
                    'session_duration': str(datetime.now() - login_datetime).split('.')[0]
                }
            except ValueError:
                pass
        return {'logged_in': True, 'login_time': 'Unknown', 'session_duration': 'Unknown'}
    return {'logged_in': False}

# Make these functions available in templates
app.jinja_env.globals.update(is_admin_logged_in=is_admin_logged_in)
app.jinja_env.globals.update(get_session_info=get_session_info)

@app.route('/admin/employees')
@admin_required
def admin_employees():
    employees = Employee.query.all()
    return render_template('admin_employees.html', employees=employees)

@app.route('/admin/employee/add', methods=['GET', 'POST'])
@admin_required
def admin_employee_add():
    if request.method == 'POST':
        emp_id = request.form.get('emp_id')
        name = request.form.get('name')
        designation = request.form.get('designation')
        location = request.form.get('location')
        department = request.form.get('department')
        grade = request.form.get('grade')
        status = request.form.get('status')
        if emp_id and name:
            emp = Employee(emp_id=emp_id, name=name, designation=designation, location=location, department=department, grade=grade, status=status)
            db.session.add(emp)
            db.session.commit()
            flash('Employee added.', 'success')
            return redirect(url_for('admin_employees'))
        else:
            flash('EmpID and Name are required.', 'danger')
    return render_template('admin_employee_form.html', action='Add', employee=None)

@app.route('/admin/employee/edit/<int:emp_id>', methods=['GET', 'POST'])
@admin_required
def admin_employee_edit(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    if request.method == 'POST':
        emp.emp_id = request.form.get('emp_id')
        emp.name = request.form.get('name')
        emp.designation = request.form.get('designation')
        emp.location = request.form.get('location')
        emp.department = request.form.get('department')
        emp.grade = request.form.get('grade')
        emp.status = request.form.get('status')
        db.session.commit()
        flash('Employee updated.', 'success')
        return redirect(url_for('admin_employees'))
    return render_template('admin_employee_form.html', action='Edit', employee=emp)

@app.route('/admin/employee/delete/<int:emp_id>', methods=['POST'])
@admin_required
def admin_employee_delete(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    db.session.delete(emp)
    db.session.commit()
    flash('Employee deleted.', 'success')
    return redirect(url_for('admin_employees'))

def init_db():
    with app.app_context():
        db.create_all()
        
        # Add shift types if not present
        if ShiftType.query.count() == 0:
            for code, desc in SHIFT_CODES:
                db.session.add(ShiftType(code=code, description=desc))
            db.session.commit()

def generate_monthly_rota(year, month):
    with app.app_context():
        employees = Employee.query.filter_by(status='active').all()
        shift_types = ShiftType.query.filter(ShiftType.code.in_(['M','E','N','G','Off'])).all()
        
        # Get date range for the month
        first_day = date(year, month, 1)
        if month == 12:
            next_month = date(year+1, 1, 1)
        else:
            next_month = date(year, month+1, 1)
        last_day = next_month - timedelta(days=1)
        
        # Remove existing rota for the month
        ShiftRota.query.filter(
            ShiftRota.date >= first_day,
            ShiftRota.date <= last_day
        ).delete()
        db.session.commit()
        
        # Generate rota for each day - create a basic pattern that admin can modify
        days = (next_month - first_day).days
        for emp in employees:
            for d in range(days):
                day = first_day + timedelta(days=d)
                # Create a simple pattern: General shift for weekdays, Off for weekends
                if day.weekday() < 5:  # Monday to Friday
                    shift = ShiftType.query.filter_by(code='G').first()
                else:  # Saturday and Sunday
                    shift = ShiftType.query.filter_by(code='Off').first()
                
                if shift:
                    db.session.add(ShiftRota(employee_id=emp.id, date=day, shift_type_id=shift.id))
        db.session.commit()

def process_attendance_and_exceptions(year, month):
    with app.app_context():
        # Get date range for the month
        first_day = date(year, month, 1)
        if month == 12:
            next_month = date(year+1, 1, 1)
        else:
            next_month = date(year, month+1, 1)
        last_day = next_month - timedelta(days=1)
        
        ExceptionReport.query.filter(
            ExceptionReport.date >= first_day,
            ExceptionReport.date <= last_day
        ).delete()
        db.session.commit()
        
        rotas = ShiftRota.query.filter(
            ShiftRota.date >= first_day,
            ShiftRota.date <= last_day
        ).all()
        
        for rota in rotas:
            emp = Employee.query.get(rota.employee_id)
            shift = ShiftType.query.get(rota.shift_type_id)
            att = Attendance.query.filter_by(employee_id=emp.id, date=rota.date).first()
            # Absenteeism
            if not att:
                if shift.code not in ['Off', 'Leave']:
                    db.session.add(ExceptionReport(
                        employee_id=emp.id, 
                        date=rota.date, 
                        issue='Absent without info (Leave not marked)',
                        status='pending'
                    ))
                continue
            # Shift mismatch
            if att.status == 'P' and shift.code not in ['Off', 'Leave']:
                if att.status == 'P' and att.time_in and shift.code in SHIFT_START:
                    # Late arrival
                    if (datetime.combine(rota.date, att.time_in) - datetime.combine(rota.date, SHIFT_START[shift.code])) > LATE_THRESHOLD:
                        db.session.add(ExceptionReport(
                            employee_id=emp.id, 
                            date=rota.date, 
                            issue='Late Arrival',
                            status='pending'
                        ))
                if att.status == 'P' and att.time_in and shift.code not in SHIFT_START:
                    db.session.add(ExceptionReport(
                        employee_id=emp.id, 
                        date=rota.date, 
                        issue='Shift mismatch',
                        status='pending'
                    ))
        db.session.commit()

@app.route('/generate_rota')
def generate_rota():
    today = date.today()
    generate_monthly_rota(today.year, today.month)
    return redirect(url_for('view_rota'))

@app.route('/rota')
def view_rota():
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    rotas = db.session.query(ShiftRota, Employee, ShiftType).join(
        Employee, ShiftRota.employee_id==Employee.id
    ).join(
        ShiftType, ShiftRota.shift_type_id==ShiftType.id
    ).filter(
        ShiftRota.date >= start_date,
        ShiftRota.date <= end_date
    ).order_by(ShiftRota.date, Employee.name).all()
    
    return render_template('rota.html', rotas=rotas)

@app.route('/')
def index():
    employees = Employee.query.all()
    return render_template('index.html', employees=employees)

@app.route('/process_exceptions')
def process_exceptions():
    # Check if admin is logged in, but don't require it
    is_admin = session.get('admin', False)
    
    try:
        today = date.today()
        process_attendance_and_exceptions(today.year, today.month)
        
        if is_admin:
            flash('Discrepancies processed successfully for current month.', 'success')
        else:
            flash('Discrepancies processed successfully for current month. (Note: Admin login recommended for full access)', 'warning')
        
        return redirect(url_for('view_exceptions'))
    except Exception as e:
        flash(f'Error processing discrepancies: {str(e)}', 'danger')
        return redirect(url_for('view_exceptions'))

@app.route('/exceptions')
def view_exceptions():
    # Get filter parameters
    status_filter = request.args.get('status', 'all')
    employee_filter = request.args.get('employee', 'all')
    issue_filter = request.args.get('issue', 'all')
    
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    # Build query with filters
    query = db.session.query(ExceptionReport, Employee).join(
        Employee, ExceptionReport.employee_id==Employee.id
    ).filter(
        ExceptionReport.date >= start_date,
        ExceptionReport.date <= end_date
    )
    
    # Apply filters
    if status_filter != 'all':
        query = query.filter(ExceptionReport.status == status_filter)
    
    if employee_filter != 'all':
        query = query.filter(Employee.id == int(employee_filter))
    
    if issue_filter != 'all':
        query = query.filter(ExceptionReport.issue.contains(issue_filter))
    
    exceptions = query.order_by(ExceptionReport.date.desc(), Employee.name).all()
    
    # Get filter options
    employees = Employee.query.order_by(Employee.name).all()
    statuses = ['pending', 'processed', 'resolved']
    issues = ['Late Arrival', 'Absent', 'Shift mismatch']
    
    return render_template('exceptions.html', 
                         exceptions=exceptions, 
                         employees=employees,
                         statuses=statuses,
                         issues=issues,
                         current_filters={'status': status_filter, 'employee': employee_filter, 'issue': issue_filter})

@app.route('/exception/<int:exception_id>/update', methods=['POST'])
@admin_required
def update_exception(exception_id):
    exception = ExceptionReport.query.get_or_404(exception_id)
    action = request.form.get('action')
    notes = request.form.get('notes', '').strip()
    
    if action == 'process':
        exception.status = 'processed'
        flash(f'Discrepancy marked as processed.', 'success')
    elif action == 'resolve':
        exception.status = 'resolved'
        flash(f'Discrepancy marked as resolved.', 'success')
    elif action == 'reopen':
        exception.status = 'pending'
        flash(f'Discrepancy reopened for review.', 'warning')
    
    if notes:
        exception.notes = notes
    
    db.session.commit()
    return redirect(url_for('view_exceptions'))

@app.route('/exception/<int:exception_id>/details')
@admin_required
def exception_details(exception_id):
    exception = db.session.query(ExceptionReport, Employee).join(
        Employee, ExceptionReport.employee_id==Employee.id
    ).filter(ExceptionReport.id == exception_id).first_or_404()
    
    return render_template('exception_details.html', exception=exception)



@app.route('/export_exceptions_excel')
def export_exceptions_excel():
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    exceptions = db.session.query(ExceptionReport, Employee).join(
        Employee, ExceptionReport.employee_id==Employee.id
    ).filter(
        ExceptionReport.date >= start_date,
        ExceptionReport.date <= end_date
    ).order_by(ExceptionReport.date, Employee.name).all()
    
    data = [{
        'Date': exception.date.strftime('%Y-%m-%d'),
        'Employee': emp.name,
        'Issue': exception.issue,
        'Status': exception.status.title(),
        'Notes': exception.notes or 'N/A'
    } for exception, emp in exceptions]
    
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Discrepancies')
    output.seek(0)
    return send_file(output, download_name='discrepancy_report.xlsx', as_attachment=True)

@app.route('/export_exceptions_pdf')
def export_exceptions_pdf():
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    exceptions = db.session.query(ExceptionReport, Employee).join(
        Employee, ExceptionReport.employee_id==Employee.id
    ).filter(
        ExceptionReport.date >= start_date,
        ExceptionReport.date <= end_date
    ).order_by(ExceptionReport.date, Employee.name).all()
    
    html = render_template('exceptions_pdf.html', exceptions=exceptions)
    result = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html), dest=result)
    result.seek(0)
    return send_file(result, download_name='discrepancy_report.pdf', as_attachment=True)

@app.route('/export_reports_excel')
def export_reports_excel():
    today = date.today()
    employees = Employee.query.all()
    summary_data = []
    
    for emp in employees:
        # Get attendance records for current month
        start_date = date(today.year, today.month, 1)
        if today.month == 12:
            end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
        
        records = Attendance.query.filter(
            Attendance.employee_id == emp.id,
            Attendance.date >= start_date,
            Attendance.date <= end_date
        ).all()
        
        # Initialize summary with proper status mapping
        summary = {
            'Employee': emp.name, 
            'Present': 0, 
            'Off': 0, 
            'Leave': 0, 
            'Absent': 0, 
            'Late': 0, 
            'Early Leave': 0, 
            'On Duty': 0
        }
        
        # Map status codes to summary keys
        status_mapping = {
            'P': 'Present',
            'A': 'Absent', 
            'L': 'Late',
            'E': 'Early Leave',
            'OD': 'On Duty'
        }
        
        for att in records:
            if att.status in status_mapping:
                summary[status_mapping[att.status]] += 1
        
        # Count offs and leaves from rota
        rotas = ShiftRota.query.filter(
            ShiftRota.employee_id == emp.id,
            ShiftRota.date >= start_date,
            ShiftRota.date <= end_date
        ).all()
        
        for rota in rotas:
            shift = ShiftType.query.get(rota.shift_type_id)
            if shift and shift.code == 'Off':
                summary['Off'] += 1
            if shift and shift.code == 'Leave':
                summary['Leave'] += 1
        
        # Calculate attendance percentage
        total_days = sum([summary['Present'], summary['Off'], summary['Leave'], summary['Absent'], summary['Late'], summary['Early Leave'], summary['On Duty']])
        if total_days > 0:
            attendance_percentage = round((summary['Present'] / total_days) * 100, 2)
        else:
            attendance_percentage = 0
        summary['Attendance %'] = attendance_percentage
        summary_data.append(summary)
    
    df = pd.DataFrame(summary_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Monthly Summary')
    output.seek(0)
    return send_file(output, download_name='monthly_attendance_report.xlsx', as_attachment=True)

@app.route('/export_reports_pdf')
def export_reports_pdf():
    today = date.today()
    employees = Employee.query.all()
    summary_data = []
    
    for emp in employees:
        # Get attendance records for current month
        start_date = date(today.year, today.month, 1)
        if today.month == 12:
            end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
        
        records = Attendance.query.filter(
            Attendance.employee_id == emp.id,
            Attendance.date >= start_date,
            Attendance.date <= end_date
        ).all()
        
        # Initialize summary with proper status mapping
        summary = {
            'name': emp.name, 
            'P': 0, 
            'Off': 0, 
            'Leave': 0, 
            'A': 0, 
            'L': 0, 
            'E': 0, 
            'OD': 0
        }
        
        # Map status codes to summary keys
        status_mapping = {
            'P': 'P',
            'A': 'A', 
            'L': 'L',
            'E': 'E',
            'OD': 'OD'
        }
        
        for att in records:
            if att.status in status_mapping:
                summary[status_mapping[att.status]] += 1
        
        # Count offs and leaves from rota
        rotas = ShiftRota.query.filter(
            ShiftRota.employee_id == emp.id,
            ShiftRota.date >= start_date,
            ShiftRota.date <= end_date
        ).all()
        
        for rota in rotas:
            shift = ShiftType.query.get(rota.shift_type_id)
            if shift and shift.code == 'Off':
                summary['Off'] += 1
            if shift and shift.code == 'Leave':
                summary['Leave'] += 1
        
        summary_data.append(summary)
    
    # Calculate attendance percentage for each employee
    for summary in summary_data:
        total_days = sum([summary['P'], summary['Off'], summary['Leave'], summary['A'], summary['L'], summary['E'], summary['OD']])
        if total_days > 0:
            attendance_percentage = round((summary['P'] / total_days) * 100, 2)
        else:
            attendance_percentage = 0
        summary['attendance_percentage'] = attendance_percentage
    
    # Create simple professional PDF template
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ 
                font-family: Arial, sans-serif;
                margin: 2cm;
                padding: 0;
                background-color: white;
                color: #333;
                line-height: 1.3;
            }}
            
            .header-section {{
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 1px solid #333;
                padding-bottom: 10px;
            }}
            
            .report-title {{
                font-size: 16pt;
                font-weight: bold;
                color: #333;
                margin-bottom: 5px;
            }}
            
            .report-meta {{
                font-size: 10pt;
                color: #666;
                margin-bottom: 10px;
            }}
            
            table {{ 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 15px;
                font-size: 10pt;
            }}
            
            th {{ 
                background-color: #f0f0f0;
                color: #333; 
                font-weight: bold;
                font-size: 10pt;
                padding: 8px 6px;
                text-align: center;
                border: 1px solid #333;
            }}
            
            td {{ 
                padding: 6px 4px;
                text-align: center;
                border: 1px solid #ccc;
                vertical-align: middle;
            }}
            
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            
            .employee-name {{
                font-weight: bold;
                text-align: left;
            }}
            
            .footer-section {{
                margin-top: 20px;
                text-align: center;
                border-top: 1px solid #ccc;
                padding-top: 10px;
            }}
            
            .footer-text {{
                font-size: 9pt;
                color: #666;
            }}
        </style>
    </head>
    <body>
        <div class="header-section">
            <div class="report-title">Monthly Attendance Summary Report</div>
            <div class="report-meta">
                Generated on: {today.strftime('%d %B %Y')} | 
                Period: {today.strftime('%B %Y')} | 
                Total Employees: {len(employees)}
            </div>
        </div>
        
        <table>
            <thead>
                <tr>
                    <th>Employee Name</th>
                    <th>Present</th>
                    <th>Off</th>
                    <th>Leave</th>
                    <th>Absent</th>
                    <th>Late</th>
                    <th>Early Leave</th>
                    <th>On Duty</th>
                    <th>Attendance %</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for summary in summary_data:
        html_content += f"""
                    <tr>
                        <td class="employee-name">{summary['name']}</td>
                        <td>{summary['P']}</td>
                        <td>{summary['Off']}</td>
                        <td>{summary['Leave']}</td>
                        <td>{summary['A']}</td>
                        <td>{summary['L']}</td>
                        <td>{summary['E']}</td>
                        <td>{summary['OD']}</td>
                        <td>{summary['attendance_percentage']}%</td>
                    </tr>
        """
    
    html_content += f"""
                </tbody>
            </table>
        
        <div class="footer-section">
            <div class="footer-text">
                Report Generated on: {today.strftime('%d %B %Y')}
            </div>
        </div>
    </body>
    </html>
    """
    
    result = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html_content), dest=result)
    result.seek(0)
    return send_file(result, download_name='monthly_attendance_report.pdf', as_attachment=True)

@app.route('/export_attendance_excel')
def export_attendance_excel():
    attendance = db.session.query(Attendance, Employee).join(Employee, Attendance.employee_id==Employee.id).order_by(Attendance.date.desc()).all()
    data = [{
        'Date': att.date.strftime('%Y-%m-%d'),
        'Employee': emp.name,
        'Employee ID': emp.emp_id,
        'Status': att.status,
        'Time In': att.time_in.strftime('%H:%M') if att.time_in else 'N/A',
        'Time Out': att.time_out.strftime('%H:%M') if att.time_out else 'N/A',
        'Duration': calculate_duration(att.time_in, att.time_out) if att.time_in and att.time_out else 'N/A'
    } for att, emp in attendance]
    
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Attendance Data')
    output.seek(0)
    return send_file(output, download_name='attendance_data.xlsx', as_attachment=True)

@app.route('/export_attendance_pdf')
def export_attendance_pdf():
    attendance = db.session.query(Attendance, Employee).join(Employee, Attendance.employee_id==Employee.id).order_by(Attendance.date.desc()).all()
    
    # Create simple professional PDF template
    html_content = """
    <html>
    <head>
        <style>
            @page {
                size: A4 landscape;
                margin: 2cm;
            }
            
            body { 
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: white;
                color: #333;
                line-height: 1.3;
            }
            
            .header-section {
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 1px solid #333;
                padding-bottom: 10px;
            }
            
            .report-title {
                font-size: 16pt;
                font-weight: bold;
                color: #333;
                margin-bottom: 5px;
            }
            
            .report-meta {
                font-size: 10pt;
                color: #666;
                margin-bottom: 10px;
            }
            
            table { 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 15px;
                font-size: 10pt;
            }
            
            th { 
                background-color: #f0f0f0;
                color: #333; 
                font-weight: bold;
                font-size: 10pt;
                padding: 8px 6px;
                text-align: center;
                border: 1px solid #333;
            }
            
            td { 
                padding: 6px 4px;
                text-align: center;
                border: 1px solid #ccc;
                vertical-align: middle;
            }
            
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            
            .employee-name {
                font-weight: bold;
                text-align: left;
            }
            
            .footer-section {
                margin-top: 20px;
                text-align: center;
                border-top: 1px solid #ccc;
                padding-top: 10px;
            }
            
            .footer-text {
                font-size: 9pt;
                color: #666;
            }
        </style>
    </head>
    <body>
        <div class="header-section">
            <div class="report-title">Attendance Report</div>
            <div class="report-meta">
                Generated on: """ + date.today().strftime('%d %B %Y') + """ | 
                Total Records: """ + str(len(attendance)) + """
            </div>
        </div>
        
        <table>
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Employee Name</th>
                    <th>Employee ID</th>
                    <th>Status</th>
                    <th>Time In</th>
                    <th>Time Out</th>
                    <th>Duration</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for att, emp in attendance:
        # Calculate duration properly
        if att.time_in and att.time_out:
            # Convert times to datetime objects for calculation
            today = datetime.now().date()
            time_in_dt = datetime.combine(today, att.time_in)
            time_out_dt = datetime.combine(today, att.time_out)
            
            # Handle case where time_out is on the next day
            if time_out_dt < time_in_dt:
                time_out_dt += timedelta(days=1)
            
            duration_td = time_out_dt - time_in_dt
            # Format duration as HH:MM
            total_seconds = int(duration_td.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            duration = f"{hours:02d}:{minutes:02d}"
        else:
            duration = 'N/A'
            
        html_content += f"""
                <tr>
                    <td>{att.date.strftime('%Y-%m-%d')}</td>
                    <td class="employee-name">{emp.name}</td>
                    <td>{emp.emp_id}</td>
                    <td>{att.status}</td>
                    <td>{att.time_in.strftime('%H:%M') if att.time_in else 'N/A'}</td>
                    <td>{att.time_out.strftime('%H:%M') if att.time_out else 'N/A'}</td>
                    <td>{duration}</td>
                </tr>
        """
    
    html_content += f"""
            </tbody>
        </table>
        
        <div class="footer-section">
            <div class="footer-text">
                Report Generated on: {date.today().strftime('%d %B %Y')}
            </div>
        </div>
    </body>
    </html>
    """
    
    result = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html_content), dest=result)
    result.seek(0)
    return send_file(result, download_name='attendance_report.pdf', as_attachment=True)

@app.route('/export_employees_excel')
def export_employees_excel():
    employees = Employee.query.all()
    data = [{
        'Employee ID': emp.emp_id,
        'Name': emp.name,
        'Designation': emp.designation or 'N/A',
        'Department': emp.department or 'N/A',
        'Location': emp.location or 'N/A',
        'Grade': emp.grade or 'N/A',
        'Status': emp.status or 'N/A'
    } for emp in employees]
    
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Employee Data')
    output.seek(0)
    return send_file(output, download_name='employee_data.xlsx', as_attachment=True)

@app.route('/export_employees_pdf')
def export_employees_pdf():
    employees = Employee.query.all()
    
    # Create simple professional PDF template
    html_content = f"""
    <html>
    <head>
        <style>
            @page {{
                size: A4;
                margin: 1.5cm;
            }}
            
            body {{ 
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: white;
                color: #333;
                line-height: 1.3;
            }}
            
            .header-section {{
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 1px solid #333;
                padding-bottom: 10px;
            }}
            
            .report-title {{
                font-size: 16pt;
                font-weight: bold;
                color: #333;
                margin-bottom: 5px;
            }}
            
            .report-meta {{
                font-size: 10pt;
                color: #666;
                margin-bottom: 10px;
            }}
            
            table {{ 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 15px;
                font-size: 10pt;
            }}
            
            th {{ 
                background-color: #f0f0f0;
                color: #333; 
                font-weight: bold;
                font-size: 10pt;
                padding: 8px 6px;
                text-align: center;
                border: 1px solid #333;
            }}
            
            td {{ 
                padding: 6px 4px;
                text-align: center;
                border: 1px solid #ccc;
                vertical-align: middle;
            }}
            
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            
            .employee-name {{
                font-weight: bold;
                text-align: left;
            }}
            
            .footer-section {{
                margin-top: 20px;
                text-align: center;
                border-top: 1px solid #ccc;
                padding-top: 10px;
            }}
            
            .footer-text {{
                font-size: 9pt;
                color: #666;
            }}
        </style>
    </head>
    <body>
        <div class="header-section">
            <div class="report-title">Employee Directory Report</div>
            <div class="report-meta">
                Total Employees: {len(employees)} | 
                Generated on: {date.today().strftime('%d %B %Y')}
            </div>
        </div>
        
        <table>
            <thead>
                <tr>
                    <th>Employee Name</th>
                    <th>Employee ID</th>
                    <th>Department</th>
                    <th>Designation</th>
                    <th>Location</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for emp in employees:
        html_content += f"""
                    <tr>
                        <td class="employee-name">{emp.name}</td>
                        <td>{emp.emp_id}</td>
                        <td>{emp.department or 'N/A'}</td>
                        <td>{emp.designation or 'N/A'}</td>
                        <td>{emp.location or 'N/A'}</td>
                    </tr>
        """
    
    html_content += f"""
                </tbody>
            </table>
        
        <div class="footer-section">
            <div class="footer-text">
                Report Generated on: {date.today().strftime('%d %B %Y')}
            </div>
        </div>
    </body>
    </html>
    """
    
    result = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html_content), dest=result)
    result.seek(0)
    return send_file(result, download_name='employee_directory.pdf', as_attachment=True)

@app.route('/export_rota_excel')
def export_rota_excel():
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    rotas = db.session.query(ShiftRota, Employee, ShiftType).join(
        Employee, ShiftRota.employee_id==Employee.id
    ).join(
        ShiftType, ShiftRota.shift_type_id==ShiftType.id
    ).filter(
        ShiftRota.date >= start_date,
        ShiftRota.date <= end_date
    ).order_by(ShiftRota.date, Employee.name).all()
    
    data = [{
        'Date': rota.date.strftime('%Y-%m-%d'),
        'Employee': emp.name,
        'Employee ID': emp.emp_id,
        'Shift': shift.description,
        'Shift Code': shift.code,
        'Start Time': SHIFT_START.get(shift.code, 'N/A').strftime('%H:%M') if isinstance(SHIFT_START.get(shift.code, 'N/A'), time) else 'N/A',
        'Department': emp.department or 'N/A'
    } for rota, emp, shift in rotas]
    
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Shift Rota')
    output.seek(0)
    return send_file(output, download_name='shift_rota.xlsx', as_attachment=True)

@app.route('/export_rota_pdf')
def export_rota_pdf():
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    rotas = db.session.query(ShiftRota, Employee, ShiftType).join(
        Employee, ShiftRota.employee_id==Employee.id
    ).join(
        ShiftType, ShiftRota.shift_type_id==ShiftType.id
    ).filter(
        ShiftRota.date >= start_date,
        ShiftRota.date <= end_date
    ).order_by(ShiftRota.date, Employee.name).all()
    
    # Create simple professional PDF template
    html_content = f"""
    <html>
    <head>
        <style>
            @page {{
                size: A4 landscape;
                margin: 2cm;
            }}
            
            body {{ 
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: white;
                color: #333;
                line-height: 1.3;
            }}
            
            .header-section {{
                text-align: center;
                margin-bottom: 20px;
                border-bottom: 1px solid #333;
                padding-bottom: 10px;
            }}
            
            .report-title {{
                font-size: 16pt;
                font-weight: bold;
                color: #333;
                margin-bottom: 5px;
            }}
            
            .report-meta {{
                font-size: 10pt;
                color: #666;
                margin-bottom: 10px;
            }}
            
            table {{ 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 15px;
                font-size: 10pt;
            }}
            
            th {{ 
                background-color: #f0f0f0;
                color: #333; 
                font-weight: bold;
                font-size: 10pt;
                padding: 8px 6px;
                text-align: center;
                border: 1px solid #333;
            }}
            
            td {{ 
                padding: 6px 4px;
                text-align: center;
                border: 1px solid #ccc;
                vertical-align: middle;
            }}
            
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            
            .employee-name {{
                font-weight: bold;
                text-align: left;
            }}
            
            .footer-section {{
                margin-top: 20px;
                text-align: center;
                border-top: 1px solid #ccc;
                padding-top: 10px;
            }}
            
            .footer-text {{
                font-size: 9pt;
                color: #666;
            }}
        </style>
    </head>
    <body>
        <div class="header-section">
            <div class="report-title">Shift Rota Report</div>
            <div class="report-meta">
                Period: {today.strftime('%B %Y')} | 
                Total Entries: {len(rotas)}
            </div>
        </div>
        
                    <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Employee Name</th>
                        <th>Employee ID</th>
                        <th>Shift Type</th>
                        <th>Start Time</th>
                        <th>Department</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    if rotas:
        for rota, emp, shift in rotas:
            start_time = SHIFT_START.get(shift.code, 'N/A')
            if start_time != 'N/A':
                start_time = start_time.strftime('%H:%M')
            
            html_content += f"""
                    <tr>
                        <td>{rota.date.strftime('%Y-%m-%d')}</td>
                        <td class="employee-name">{emp.name}</td>
                        <td>{emp.emp_id}</td>
                        <td>{shift.description}</td>
                        <td>{start_time}</td>
                        <td>{emp.department or 'N/A'}</td>
                    </tr>
            """
    else:
        html_content += """
                    <tr>
                        <td colspan="6" style="text-align: center; padding: 20px; color: #666;">
                            No rota data available for the current month.
                        </td>
                    </tr>
        """
    
    html_content += f"""
            </tbody>
        </table>
        
        <div class="footer-section">
            <div class="footer-text">
                Report Generated on: {today.strftime('%d %B %Y')}
            </div>
        </div>
    </body>
    </html>
    """
    
    result = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html_content), dest=result)
    result.seek(0)
    return send_file(result, download_name='shift_rota.pdf', as_attachment=True)

@app.route('/employee')
def employee_page():
    employees = Employee.query.all()
    return render_template('employee.html', employees=employees)

@app.route('/attendance')
def attendance_page():
    attendance = db.session.query(Attendance, Employee).join(Employee, Attendance.employee_id==Employee.id).order_by(Attendance.date.desc()).all()
    return render_template('attendance.html', attendance=attendance)

@app.route('/reports')
def reports_page():
    today = date.today()
    
    # Get date range for current month
    start_date = date(today.year, today.month, 1)
    if today.month == 12:
        end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
    
    exceptions = db.session.query(ExceptionReport, Employee).join(
        Employee, ExceptionReport.employee_id==Employee.id
    ).filter(
        ExceptionReport.date >= start_date,
        ExceptionReport.date <= end_date
    ).order_by(ExceptionReport.date.desc()).all()
    
    # Monthly summary
    employees = Employee.query.all()
    summary_data = []
    for emp in employees:
        records = Attendance.query.filter(
            Attendance.employee_id == emp.id,
            Attendance.date >= start_date,
            Attendance.date <= end_date
        ).all()
        
        summary = {'name': emp.name, 'P': 0, 'Off': 0, 'Leave': 0, 'A': 0, 'L': 0, 'E': 0, 'OD': 0}
        
        # Map status codes to summary keys
        status_mapping = {
            'P': 'P',
            'A': 'A', 
            'L': 'L',
            'E': 'E',
            'OD': 'OD'
        }
        
        for att in records:
            if att.status in status_mapping:
                summary[status_mapping[att.status]] += 1
        
        # Count offs and leaves from rota
        rotas = ShiftRota.query.filter(
            ShiftRota.employee_id == emp.id,
            ShiftRota.date >= start_date,
            ShiftRota.date <= end_date
        ).all()
        
        for rota in rotas:
            shift = ShiftType.query.get(rota.shift_type_id)
            if shift and shift.code == 'Off':
                summary['Off'] += 1
            if shift and shift.code == 'Leave':
                summary['Leave'] += 1
        
        summary_data.append(summary)
    
    return render_template('reports.html', exceptions=exceptions, summary_data=summary_data)

@app.route('/attendance_upload', methods=['GET', 'POST'])
def attendance_upload():
    if request.method == 'POST':
        file = request.files.get('file')
        if file and file.filename.endswith('.csv'):
            stream = io.StringIO(file.stream.read().decode('UTF8'), newline=None)
            reader = csv.DictReader(stream)
            count = 0
            for row in reader:
                emp = Employee.query.filter_by(emp_id=row.get('EmpID')).first()
                if emp:
                    att = Attendance(
                        employee_id=emp.id,
                        date=datetime.strptime(row.get('Date'), '%Y-%m-%d').date(),
                        status=row.get('Status'),
                        time_in=datetime.strptime(row.get('TimeIn'), '%H:%M').time() if row.get('TimeIn') else None,
                        time_out=datetime.strptime(row.get('TimeOut'), '%H:%M').time() if row.get('TimeOut') else None
                    )
                    db.session.add(att)
                    count += 1
            db.session.commit()
            flash(f'Successfully uploaded {count} attendance records.', 'success')
            return redirect(url_for('attendance_page'))
        else:
            flash('Please upload a valid CSV file.', 'danger')
    return render_template('attendance_upload.html')

@app.route('/attendance_entry', methods=['GET', 'POST'])
def attendance_entry():
    employees = Employee.query.filter_by(status='active').all()
    if request.method == 'POST':
        emp_id = request.form.get('employee_id')
        date_str = request.form.get('date')
        status = request.form.get('status')
        time_in = request.form.get('time_in')
        time_out = request.form.get('time_out')
        emp = Employee.query.get(emp_id)
        if emp and date_str and status:
            att = Attendance(
                employee_id=emp.id,
                date=datetime.strptime(date_str, '%Y-%m-%d').date(),
                status=status,
                time_in=datetime.strptime(time_in, '%H:%M').time() if time_in else None,
                time_out=datetime.strptime(time_out, '%H:%M').time() if time_out else None
            )
            db.session.add(att)
            db.session.commit()
            flash('Attendance record added.', 'success')
            return redirect(url_for('attendance_page'))
        else:
            flash('Please fill all required fields.', 'danger')
    return render_template('attendance_entry.html', employees=employees)

@app.route('/init-db')
def initialize_database():
    init_db()
    flash('Database initialized with sample data.', 'success')
    return redirect(url_for('index'))



@app.route('/clear-sample-data')
def clear_sample_data():
    """Clear all sample data from the database"""
    try:
        with app.app_context():
            # Clear all data from all tables
            ExceptionReport.query.delete()
            ShiftRota.query.delete()
            Attendance.query.delete()
            Employee.query.delete()
            ShiftType.query.delete()
            
            # Commit the changes
            db.session.commit()
            
            flash('All sample data has been cleared from the database.', 'success')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error clearing sample data: {str(e)}', 'danger')
        return redirect(url_for('index'))

@app.route('/debug-data')
def debug_data():
    employees = Employee.query.all()
    attendance = Attendance.query.all()
    rotas = ShiftRota.query.all()
    exceptions = ExceptionReport.query.all()
    
    debug_info = {
        'employees_count': len(employees),
        'attendance_count': len(attendance),
        'rotas_count': len(rotas),
        'exceptions_count': len(exceptions),
        'employees': [{'id': e.id, 'name': e.name, 'emp_id': e.emp_id} for e in employees[:3]],
        'attendance': [{'date': str(a.date), 'status': a.status, 'employee_id': a.employee_id} for a in attendance[:3]]
    }
    
    return f"""
    <h1>Debug Information</h1>
    <p>Employees: {debug_info['employees_count']}</p>
    <p>Attendance Records: {debug_info['attendance_count']}</p>
    <p>Rota Records: {debug_info['rotas_count']}</p>
    <p>Exception Records: {debug_info['exceptions_count']}</p>
    <h2>Sample Employees:</h2>
    <pre>{debug_info['employees']}</pre>
    <h2>Sample Attendance:</h2>
    <pre>{debug_info['attendance']}</pre>
    """

@app.route('/test-connection')
def test_connection():
    """Simple test route to check if requests are working"""
    return jsonify({
        'status': 'success',
        'message': 'Connection test successful',
        'timestamp': datetime.now().isoformat(),
        'session_info': {
            'admin_logged_in': session.get('admin', False),
            'session_id': session.get('_id', 'Not set')
        }
    })

@app.route('/debug-session')
def debug_session():
    """Debug route to check session state"""
    session_info = {
        'admin_logged_in': session.get('admin', False),
        'login_time': session.get('login_time', 'Not set'),
        'user_agent': session.get('user_agent', 'Not set'),
        'session_id': session.get('_id', 'Not set'),
        'all_session_keys': list(session.keys())
    }
    
    # Check if session is expired
    if session.get('login_time'):
        try:
            login_datetime = datetime.fromisoformat(session.get('login_time'))
            time_diff = datetime.now() - login_datetime
            session_info['session_age'] = str(time_diff)
            session_info['session_expired'] = time_diff > timedelta(hours=8)
        except ValueError:
            session_info['session_age'] = 'Invalid login time format'
            session_info['session_expired'] = False
    
    return f"""
    <h1>Session Debug Information</h1>
    <pre>{session_info}</pre>
    <p><a href="/exceptions">Back to Exceptions</a></p>
    <p><a href="/admin/login">Admin Login</a></p>
    """

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=8080) 