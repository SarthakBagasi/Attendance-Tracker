# Attendance & Shift Rota Management System

## Overview
A web-based system to manage employee database, shift rota, attendance, and exception reports.

## Tech Stack
- Backend: Python (Flask)
- Frontend: HTML + Bootstrap
- Database: SQLite (easy to migrate to PostgreSQL/Oracle)

## Features
- Employee database management
- Shift rota generator and calendar view
- Attendance input (CSV upload & web entry)
- Attendance processing and summary
- Exception reporting

## Setup Instructions
1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Run the app:
   ```
   python app.py
   ```
3. Open your browser at [http://127.0.0.1:5000](http://127.0.0.1:5000)

## Folder Structure
- `app.py`: Main Flask app
- `models.py`: Database models
- `templates/`: HTML templates
- `static/`: Static files (CSS, JS) 