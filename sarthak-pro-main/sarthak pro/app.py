from flask import Flask, render_template, redirect, url_for, send_file, make_response
from models import db, Employee, ShiftType, ShiftRota, Attendance, ExceptionReport
from datetime import date, timedelta, time, datetime
import random
import pandas as pd
import io
from jinja2 import Template
from xhtml2pdf import pisa

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///attendance.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

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

def init_db():
    with app.app_context():
        db.create_all()
        # Add shift types if not present
        if ShiftType.query.count() == 0:
            for code, desc in SHIFT_CODES:
                db.session.add(ShiftType(code=code, description=desc))
            db.session.commit()
        # Add employees if not present
        if Employee.query.count() == 0:
            employees = [
                Employee(emp_id='E001', name='Alice Smith', designation='Manager', location='Delhi', department='HR', grade='A', status='active'),
                Employee(emp_id='E002', name='Bob Johnson', designation='Engineer', location='Mumbai', department='IT', grade='B', status='active'),
                Employee(emp_id='E003', name='Charlie Lee', designation='Technician', location='Bangalore', department='Maintenance', grade='C', status='inactive'),
                Employee(emp_id='E004', name='Diana King', designation='Analyst', location='Chennai', department='Finance', grade='B', status='active'),
                Employee(emp_id='E005', name='Ethan Brown', designation='Supervisor', location='Pune', department='Production', grade='A', status='active'),
            ]
            db.session.bulk_save_objects(employees)
            db.session.commit()

def generate_monthly_rota(year, month):
    with app.app_context():
        employees = Employee.query.filter_by(status='active').all()
        shift_types = ShiftType.query.filter(ShiftType.code.in_(['M','E','N','G','Off'])).all()
        # Remove existing rota for the month
        ShiftRota.query.filter(db.extract('year', ShiftRota.date)==year, db.extract('month', ShiftRota.date)==month).delete()
        db.session.commit()
        # Generate rota for each day
        first_day = date(year, month, 1)
        if month == 12:
            next_month = date(year+1, 1, 1)
        else:
            next_month = date(year, month+1, 1)
        days = (next_month - first_day).days
        for emp in employees:
            for d in range(days):
                day = first_day + timedelta(days=d)
                shift = random.choice(shift_types)
                db.session.add(ShiftRota(employee_id=emp.id, date=day, shift_type_id=shift.id))
        db.session.commit()

def process_attendance_and_exceptions(year, month):
    with app.app_context():
        ExceptionReport.query.filter(db.extract('year', ExceptionReport.date)==year, db.extract('month', ExceptionReport.date)==month).delete()
        db.session.commit()
        rotas = ShiftRota.query.filter(db.extract('year', ShiftRota.date)==year, db.extract('month', ShiftRota.date)==month).all()
        for rota in rotas:
            emp = Employee.query.get(rota.employee_id)
            shift = ShiftType.query.get(rota.shift_type_id)
            att = Attendance.query.filter_by(employee_id=emp.id, date=rota.date).first()
            # Absenteeism
            if not att:
                if shift.code not in ['Off', 'Leave']:
                    db.session.add(ExceptionReport(employee_id=emp.id, date=rota.date, issue='Absent'))
                continue
            # Shift mismatch
            if att.status == 'P' and shift.code not in ['Off', 'Leave']:
                if att.status == 'P' and att.time_in and shift.code in SHIFT_START:
                    # Late arrival
                    if (datetime.combine(rota.date, att.time_in) - datetime.combine(rota.date, SHIFT_START[shift.code])) > LATE_THRESHOLD:
                        db.session.add(ExceptionReport(employee_id=emp.id, date=rota.date, issue='Late Arrival'))
                if att.status == 'P' and att.time_in and shift.code not in SHIFT_START:
                    db.session.add(ExceptionReport(employee_id=emp.id, date=rota.date, issue='Shift mismatch'))
        db.session.commit()

@app.route('/generate_rota')
def generate_rota():
    today = date.today()
    generate_monthly_rota(today.year, today.month)
    return redirect(url_for('view_rota'))

@app.route('/rota')
def view_rota():
    today = date.today()
    rotas = db.session.query(ShiftRota, Employee, ShiftType).join(Employee, ShiftRota.employee_id==Employee.id).join(ShiftType, ShiftRota.shift_type_id==ShiftType.id).filter(db.extract('year', ShiftRota.date)==today.year, db.extract('month', ShiftRota.date)==today.month).order_by(ShiftRota.date, Employee.name).all()
    return render_template('rota.html', rotas=rotas)

@app.route('/')
def index():
    employees = Employee.query.all()
    return render_template('index.html', employees=employees)

@app.route('/process_exceptions')
def process_exceptions():
    today = date.today()
    process_attendance_and_exceptions(today.year, today.month)
    return redirect(url_for('view_exceptions'))

@app.route('/exceptions')
def view_exceptions():
    today = date.today()
    exceptions = db.session.query(ExceptionReport, Employee).join(Employee, ExceptionReport.employee_id==Employee.id).filter(db.extract('year', ExceptionReport.date)==today.year, db.extract('month', ExceptionReport.date)==today.month).order_by(ExceptionReport.date, Employee.name).all()
    return render_template('exceptions.html', exceptions=exceptions)

@app.route('/export_exceptions_excel')
def export_exceptions_excel():
    today = date.today()
    exceptions = db.session.query(ExceptionReport, Employee).join(Employee, ExceptionReport.employee_id==Employee.id).filter(db.extract('year', ExceptionReport.date)==today.year, db.extract('month', ExceptionReport.date)==today.month).order_by(ExceptionReport.date, Employee.name).all()
    data = [{
        'Date': exception.date,
        'Employee': emp.name,
        'Issue': exception.issue
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
    exceptions = db.session.query(ExceptionReport, Employee).join(Employee, ExceptionReport.employee_id==Employee.id).filter(db.extract('year', ExceptionReport.date)==today.year, db.extract('month', ExceptionReport.date)==today.month).order_by(ExceptionReport.date, Employee.name).all()
    html = render_template('exceptions_pdf.html', exceptions=exceptions)
    result = io.BytesIO()
    pisa.CreatePDF(io.StringIO(html), dest=result)
    result.seek(0)
    return send_file(result, download_name='discrepancy_report.pdf', as_attachment=True)

@app.route('/employee')
def employee_page():
    return render_template('employee.html')

@app.route('/attendance')
def attendance_page():
    return render_template('attendance.html')

@app.route('/reports')
def reports_page():
    return render_template('reports.html')

if __name__ == '__main__':
    init_db()
    app.run(debug=True) 