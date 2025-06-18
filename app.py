from flask import Flask, render_template, request, redirect, flash
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
from email.mime.text import MIMEText

# --- CONFIGURATION ---
SHEET_NAME = 'TTAM'
WORKSHEET_NAME = 'Sheet1'
CREDENTIALS_FILE = 'service_account.json'
FROM_EMAIL = 'business.gpardeshi@gmail.com'
APP_PASSWORD = 'fvhi lljb nqau bqtq'
ASSIGNEE_EMAILS = {
    "Alice": "alice@example.com",
    "Bob": "bob@example.com",
    "Carol": "carol@example.com"
}

# --- GOOGLE SHEETS AUTH ---
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
gc = gspread.authorize(creds)
sheet = gc.open('TTAM').worksheet('Sheet1')

app = Flask(__name__)
app.secret_key = 'supersecretkey'

def send_email(subject, body, to_email):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = FROM_EMAIL
    msg['To'] = to_email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(FROM_EMAIL, APP_PASSWORD)
        server.sendmail(FROM_EMAIL, [to_email], msg.as_string())

def check_and_notify():
    tasks = sheet.get_all_records()
    today = datetime.now().date()
    for idx, task in enumerate(tasks, start=2):
        status = task['Status']
        assignee = task['Assignee']
        task_name = task['Task Name']
        due_date_str = task['Due Date']
        to_email = ASSIGNEE_EMAILS.get(assignee, FROM_EMAIL)
        try:
            due_date = datetime.strptime(due_date_str, '%Y-%m-%d').date()
        except Exception:
            due_date = None

        # Notify on completion
        if status.lower() == 'completed' and not task.get('Notified', False):
            subject = f"Task Completed: {task_name}"
            body = f"Hi {assignee},\n\nThe task '{task_name}' has been marked as completed."
            send_email(subject, body, to_email)
            sheet.update_cell(idx, 8, 'Notified')  # Mark as notified in Notes

        # Check task status and set appropriate flag
        if status.lower() != 'completed':
            if due_date:
                days_remaining = (due_date - today).days
                if days_remaining < 0:
                    # Task is overdue
                    sheet.update_cell(idx, 7, 'Overdue')
                    subject = f"Task Overdue: {task_name}"
                    body = f"Hi {assignee},\n\nThe task '{task_name}' is overdue! Please take action."
                    send_email(subject, body, to_email)
                elif days_remaining <= 2:
                    # Task is due soon
                    sheet.update_cell(idx, 7, 'Due Soon')
                    if days_remaining == 0:
                        subject = f"Task Due Today: {task_name}"
                        body = f"Hi {assignee},\n\nThe task '{task_name}' is due today!"
                    else:
                        subject = f"Task Due Soon: {task_name}"
                        body = f"Hi {assignee},\n\nThe task '{task_name}' is due in {days_remaining} days!"
                    send_email(subject, body, to_email)
                else:
                    # Task is on track
                    sheet.update_cell(idx, 7, 'On Track')
            else:
                # No due date set
                sheet.update_cell(idx, 7, 'No Due Date')
        else:
            # Task is completed
            sheet.update_cell(idx, 7, 'Completed')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Add new task
        task_name = request.form['task_name']
        status = request.form['status']
        assignee = request.form['assignee']
        due_date = request.form['due_date']
        dependencies = request.form['dependencies']
        notes = request.form['notes']
        task_id = len(sheet.get_all_records()) + 1
        sheet.append_row([task_id, task_name, status, assignee, due_date, dependencies, 'On Track', notes])
        flash('Task added successfully!')
        return redirect('/')
    check_and_notify()
    tasks = sheet.get_all_records()
    return render_template('index.html', tasks=tasks)

@app.route('/update_task_status/<int:task_id>', methods=['POST'])
def update_task_status(task_id):
    try:
        # Find the task row
        tasks = sheet.get_all_records()
        for idx, task in enumerate(tasks, start=2):
            if task['Task ID'] == task_id:
                # Update the status to Completed
                sheet.update_cell(idx, 3, 'Completed')  # Column 3 is Status
                flash('Task marked as completed!')
                return redirect('/')
        flash('Task not found!')
    except Exception as e:
        flash(f'Error updating task: {str(e)}')
    return redirect('/')

@app.route('/undo_task_status/<int:task_id>', methods=['POST'])
def undo_task_status(task_id):
    try:
        # Find the task row
        tasks = sheet.get_all_records()
        for idx, task in enumerate(tasks, start=2):
            if task['Task ID'] == task_id:
                # Update the status back to To Do
                sheet.update_cell(idx, 3, 'To Do')  # Column 3 is Status
                flash('Task status reverted to To Do!')
                return redirect('/')
        flash('Task not found!')
    except Exception as e:
        flash(f'Error updating task: {str(e)}')
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)