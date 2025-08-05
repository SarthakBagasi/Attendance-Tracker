from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(20), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    designation = db.Column(db.String(100))
    location = db.Column(db.String(100))
    department = db.Column(db.String(100))
    grade = db.Column(db.String(50))
    status = db.Column(db.String(20), default='active')

class ShiftType(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(10), unique=True, nullable=False)  # M, E, N, G, Off, Leave
    description = db.Column(db.String(100))

class ShiftRota(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'))
    date = db.Column(db.Date, nullable=False)
    shift_type_id = db.Column(db.Integer, db.ForeignKey('shift_type.id'))

class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'))
    date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(10))  # P, A, L, E, OD
    time_in = db.Column(db.Time)
    time_out = db.Column(db.Time)

class ExceptionReport(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'))
    date = db.Column(db.Date, nullable=False)
    issue = db.Column(db.String(200))
    status = db.Column(db.String(20), default='pending')  # pending, processed, resolved
    notes = db.Column(db.Text)  # For admin comments/notes
    created_at = db.Column(db.DateTime, default=db.func.current_timestamp())
    updated_at = db.Column(db.DateTime, default=db.func.current_timestamp(), onupdate=db.func.current_timestamp()) 