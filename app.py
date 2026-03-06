from flask import Flask, render_template, request, jsonify, session, redirect, url_for, flash
from services.expense_handler import ExpenseHandler
from services.invoice_parser import InvoiceParser
from services.report_handler import ReportHandler
from services.speech_to_text import SpeechHandler
from services.audit_logger import AuditLogger
import os
import uuid
import functools

from config import Config, USERS
from translations import TRANSLATIONS

app = Flask(__name__)
app.config.from_object(Config)

# Configuration
UPLOAD_FOLDER = app.config['UPLOAD_FOLDER']
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Services
expense_handler = ExpenseHandler()
invoice_parser  = InvoiceParser()
report_handler  = ReportHandler()
speech_handler  = SpeechHandler()
audit_logger    = AuditLogger()

# Context Processor for Translations
@app.context_processor
def inject_translations():
    lang = session.get('lang', 'ar')
    def get_text(key):
        return TRANSLATIONS.get(lang, TRANSLATIONS['ar']).get(key, key)
    return dict(t=get_text, lang=lang, dir='rtl' if lang == 'ar' else 'ltr')

# Authentication Decorator
def login_required(f):
    def wrap(*args, **kwargs):
        if 'logged_in' in session:
            return f(*args, **kwargs)
        else:
            return redirect(url_for('login'))
    wrap.__name__ = f.__name__
    return wrap

def verify_user(username, password):
    if username in USERS and USERS[username]['password'] == password:
        return USERS[username]
    return None

def role_required(required_role):
    def decorator(f):
        @functools.wraps(f)
        def wrapped(*args, **kwargs):
            if 'user_role' not in session:
                return redirect(url_for('login'))
            if session['user_role'] != required_role and session['user_role'] != 'admin':
                flash('Access denied')
                return redirect(url_for('login'))
            return f(*args, **kwargs)
        return wrapped
    return decorator

@app.route('/set_lang/<lang_code>')
def set_lang(lang_code):
    if lang_code in ['ar', 'en']:
        session['lang'] = lang_code
    return redirect(request.referrer or url_for('index'))

@app.route('/')
@role_required('record')
def index():
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'View Home')
    return render_template('index.html')

@app.route('/upload_audio', methods=['POST'])
def upload_audio():
    if 'audio_data' not in request.files:
        return jsonify({'error': 'No audio file provided'}), 400
    
    file = request.files['audio_data']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    filename = f"{uuid.uuid4()}.webm"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    text = speech_handler.transcribe(filepath)
    
    if os.path.exists(filepath):
        try: os.remove(filepath)
        except: pass
            
    if text is None:
        return jsonify({'error': 'Error processing audio'}), 500
    
    amount, description, year, explicit_type = speech_handler.parse_text(text)
    
    inferred_type = expense_handler.infer_expense_type(description) if description else 'Essential'
    final_type = explicit_type if explicit_type else inferred_type

    audit_logger.log_action(session.get('username', 'Anonymous'), session.get('user_role', 'None'), 'Transcribe Audio', f"Parsed: {amount} - {description} (Type: {final_type}, Year: {year})")

    return jsonify({
        'text': text,
        'amount': amount,
        'description': description,
        'expense_type': final_type,
        'year': year
    })

@app.route('/save_expense', methods=['POST'])
def save_expense():
    data = request.json
    amount = data.get('amount')
    description = data.get('description')
    expense_type = data.get('expense_type', 'Essential')
    year = data.get('year')

    if not amount or not description:
        return jsonify({'success': False, 'message': 'Missing data'}), 400
        
    record = expense_handler.add_expense(amount, description, expense_type, expense_year=year)
    audit_logger.log_action(session.get('username', 'Anonymous'), session.get('user_role', 'None'), 'Add Expense (Voice)', f"Amount: {amount}, Desc: {description}, Year: {year}")
    return jsonify({'success': True, 'record': record})


# --- Admin Routes ---

@app.route('/admin/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = verify_user(username, password)
        if user:
            session['logged_in'] = True
            session['username'] = username
            session['user_role'] = user['role']
            
            audit_logger.log_action(username, user['role'], 'Login', 'Success')

            # Redirect based on role
            if user['role'] == 'uploader':
                return redirect(url_for('upload_report_page'))
            elif user['role'] == 'record':
                return redirect(url_for('index'))
            return redirect(url_for('dashboard'))
        else:
            audit_logger.log_action(username, 'Unknown', 'Login', 'Failed - Invalid Credentials')
            flash('Invalid credentials')
    return render_template('login.html')

@app.route('/admin/logout')
def logout():
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Logout')
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/admin/dashboard', methods=['GET'])
@role_required('admin')
def dashboard():
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'View Dashboard')
    selected_year = request.args.get('year')
    stats = expense_handler.get_stats(selected_year)
    expenses = expense_handler.get_all_expenses()[:5]
    return render_template('admin/dashboard.html', stats=stats, expenses=expenses)

@app.route('/admin/table')
@role_required('admin')
def expenses_table():
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'View Table')
    selected_year = request.args.get('year')
    expenses = expense_handler.get_all_expenses(selected_year)
    stats = expense_handler.get_stats(selected_year)
    available_years = expense_handler.get_available_years()
    return render_template('admin/table.html', expenses=expenses, stats=stats, available_years=available_years, selected_year=selected_year)

@app.route('/admin/upload_report', methods=['GET', 'POST'])
@login_required
def upload_report_page():
    report_data = []
    columns = []
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and file.filename.endswith(('.xls', '.xlsx')):
            filename = f"report_{uuid.uuid4()}.xlsx"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            audit_logger.log_action(session.get('username'), session.get('user_role'), 'Upload Report', filename)
            report_data = report_handler.parse_uploaded_report(filepath)
            if report_data:
                columns = list(report_data[0].keys())
                report_handler.save_uploaded_data(report_data)
                flash('تم رفع البيانات وحفظها بنجاح!', 'success')

            if os.path.exists(filepath):
                os.remove(filepath)
        else:
            flash('Invalid file type')
            
    return render_template('admin/upload_report.html', data=report_data, columns=columns)

@app.route('/admin/reports_dashboard')
@role_required('admin')
def reports_dashboard():
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'View Reports Dashboard')
    stats = report_handler.get_reports_analytics()
    return render_template('admin/reports_dashboard.html', stats=stats)

@app.route('/admin/view_reports')
@login_required
def view_reports():
    all_data = report_handler.get_uploaded_data()
    selected_period = request.args.get('period')
    
    # Get unique periods for filter dropdown
    available_periods = sorted(list(set(row.get('ReportPeriod') for row in all_data if row.get('ReportPeriod') and row.get('ReportPeriod') != 'Unknown')))
    
    # Filter Data
    if selected_period:
        data = [row for row in all_data if row.get('ReportPeriod') == selected_period]
    else:
        data = all_data

    columns = []
    stats = {
        'total_records': len(data),
        'total_commission': 0,
        'total_converted': 0,
        'total_allam': 0,
        'unique_periods': len(available_periods) if not selected_period else 1,
        'total_quantity': 0
    }
    
    if data:
        columns = list(data[0].keys())
        
        # Rename ReportPeriod to التاريخ
        if 'ReportPeriod' in columns:
            idx = columns.index('ReportPeriod')
            columns[idx] = 'التاريخ'
            
        # Move ID to the front
        if 'ID' in columns:
            columns.remove('ID')
            columns.insert(0, 'ID')
            
        # Add 'الاجمالي بعد العمولة' before 'التاريخ'
        if 'التاريخ' in columns:
            if 'الاجمالي بعد العمولة' not in columns:
                idx = columns.index('التاريخ')
                columns.insert(idx, 'الاجمالي بعد العمولة')
        else:
            if 'الاجمالي بعد العمولة' not in columns:
                columns.append('الاجمالي بعد العمولة')
        
        for row in data:
            if 'ReportPeriod' in row:
                row['التاريخ'] = row.pop('ReportPeriod')
                
            # Row-level calculations
            r_commission = 0
            try: r_commission = float(str(row.get('العمولة قبل إجمالي', 0)))
            except: pass
            
            r_converted = 0
            try: r_converted = float(str(row.get('العمولة قيمة', row.get('العمولة المحولة', 0))))
            except: pass
            
            row['الاجمالي بعد العمولة'] = round(r_commission - r_converted, 2)
            
            # Global Stats
            # Sum Commission (Total)
            try:
                val = str(row.get('العمولة قبل إجمالي', 0))
                stats['total_commission'] += float(val)
            except: pass
            
            # Sum Commission (Converted / Value)
            try:
                # Try 'العمولة قيمة' first (from image), then 'العمولة المحولة'
                val = str(row.get('العمولة قيمة', row.get('العمولة المحولة', 0)))
                stats['total_converted'] += float(val)
            except: pass

            # Sum Quantity (عدد)
            try:
                # Try 'عدد' first, if likely column name from image
                val = str(row.get('عدد', 0))
                stats['total_quantity'] += float(val)
            except: pass

        stats['total_commission'] = round(stats['total_commission'], 2)
        stats['total_converted'] = round(stats['total_converted'], 2)
        stats['total_allam'] = round(stats['total_commission'] - stats['total_converted'], 2)
        stats['total_quantity'] = int(stats['total_quantity'])
        
    return render_template('admin/view_reports.html', data=data, columns=columns, stats=stats, available_periods=available_periods, selected_period=selected_period)

@app.route('/admin/report/update', methods=['POST'])
@login_required
def update_report():
    data = dict(request.form)
    id = data.pop('id', None)
    if not id:
        flash('No ID provided')
        return redirect(request.referrer or url_for('view_reports'))
        
    success = report_handler.update_report(id, data)
    
    if success:
        audit_logger.log_action(session.get('username'), session.get('user_role'), 'Update Report', f"ID: {id}")
        flash('Updated successfully')
    else:
        flash('Error updating')
    return redirect(request.referrer or url_for('view_reports'))

@app.route('/admin/report/delete/<int:id>', methods=['POST'])
@login_required
def delete_report(id):
    report_handler.delete_report(id)
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Delete Report', f"ID: {id}")
    flash('Deleted successfully')
    return redirect(request.referrer or url_for('view_reports'))

@app.route('/admin/report/delete_all', methods=['POST'])
@login_required
def delete_all_reports():
    report_handler.delete_all_reports()
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Delete All Reports')
    flash('All reports deleted successfully')
    return redirect(request.referrer or url_for('view_reports'))

@app.route('/admin/upload_invoice', methods=['GET', 'POST'])
@login_required
def upload_invoice_page():
    invoice_data = []
    columns = []
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and file.filename.endswith(('.xls', '.xlsx')):
            ext = os.path.splitext(file.filename)[1]
            filename = f"invoice_{uuid.uuid4()}{ext}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            audit_logger.log_action(session.get('username'), session.get('user_role'), 'Upload Invoice', filename)
            invoice_data = invoice_parser.parse_uploaded_invoice(filepath)
            if invoice_data:
                columns = list(invoice_data[0].keys())
                invoice_parser.save_uploaded_invoice_data(invoice_data)
                flash('تم رفع الفاتورة بنجاح!', 'success')
            else:
                flash('الملف لا يحتوي على بيانات فاتورة صالحة أو فشل في القراءة.', 'error')

            if os.path.exists(filepath):
                os.remove(filepath)
        else:
            flash('Invalid file type')
            
    return render_template('admin/upload_invoice.html', data=invoice_data, columns=columns)

@app.route('/admin/view_invoices')
@login_required
def view_invoices():
    all_data = invoice_parser.get_uploaded_invoices()
    selected_date = request.args.get('date')
    
    available_dates = sorted(list(set(row.get('التاريخ') for row in all_data if row.get('التاريخ') and row.get('التاريخ') != 'Unknown')))
    
    if selected_date:
        data = [row for row in all_data if row.get('التاريخ') == selected_date]
    else:
        data = all_data

    columns = []
    stats = {
        'total_records': len(data),
        'total_quantity': 0,
        'total_amount': 0,
        'total_expenses': 0,
        'total_net': 0
    }
    
    if data:
        columns = list(data[0].keys())
        
        if 'ID' in columns:
            columns.remove('ID')
            columns.insert(0, 'ID')
            
        for row in data:
            try: stats['total_quantity'] += float(str(row.get('العدد', 0)).replace(',',''))
            except: pass
            
            try: stats['total_amount'] += float(str(row.get('الاجمالي', 0)).replace(',',''))
            except: pass

        stats['total_quantity'] = int(stats['total_quantity'])
        stats['total_amount'] = round(stats['total_amount'], 2)
        
        unique_dates = {}
        for row in data:
            date_val = row.get('التاريخ')
            if date_val not in unique_dates:
                try: exp = float(str(row.get('تنزيل المنصرف', 0)).replace(',',''))
                except: exp = 0
                try: net = float(str(row.get('صافي الفاتورة', 0)).replace(',',''))
                except: net = 0
                unique_dates[date_val] = {'expenses': exp, 'net': net}
                
        stats['total_expenses'] = round(sum(d['expenses'] for d in unique_dates.values()), 2)
        stats['total_net'] = round(sum(d['net'] for d in unique_dates.values()), 2)
        
    return render_template('admin/view_invoices.html', data=data, columns=columns, stats=stats, available_dates=available_dates, selected_date=selected_date)

@app.route('/admin/invoice/update', methods=['POST'])
@login_required
def update_invoice():
    data = dict(request.form)
    id = data.pop('id', None)
    if not id:
        flash('No ID provided')
        return redirect(request.referrer or url_for('view_invoices'))
        
    success = invoice_parser.update_invoice(id, data)
    
    if success:
        audit_logger.log_action(session.get('username'), session.get('user_role'), 'Update Invoice', f"ID: {id}")
        flash('Updated successfully')
    else:
        flash('Error updating')
    return redirect(request.referrer or url_for('view_invoices'))

@app.route('/admin/invoice/delete/<int:id>', methods=['POST'])
@login_required
def delete_invoice(id):
    invoice_parser.delete_invoice(id)
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Delete Invoice', f"ID: {id}")
    flash('Deleted successfully')
    return redirect(request.referrer or url_for('view_invoices'))

@app.route('/admin/invoice/delete_all', methods=['POST'])
@login_required
def delete_all_invoices():
    invoice_parser.delete_all_invoices()
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Delete All Invoices')
    flash('All invoices deleted successfully')
    return redirect(request.referrer or url_for('view_invoices'))

@app.route('/admin/audit_log')
@login_required
def audit_log():
    # Only 'System' user can view this
    if session.get('username') != 'System':
        flash('Access Denied: System User Only')
        return redirect(url_for('dashboard'))
    
    logs = audit_logger.get_logs()
    columns = ['Timestamp', 'Username', 'Role', 'Action', 'Details']
    return render_template('admin/audit_log.html', logs=logs, columns=columns)

@app.route('/admin/expense/add', methods=['POST'])
@role_required('admin')
def manual_add():
    amount = request.form['amount']
    description = request.form['description']
    expense_type = request.form['expense_type']
    
    # User selects a Year (Accounting Year)
    # The Date will be automatically set to NOW() by the handler
    expense_year = request.form.get('year')
    
    expense_handler.add_expense(amount, description, expense_type, expense_year=expense_year)
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Add Expense (Manual)', f"Amount: {amount}, Desc: {description}, Year: {expense_year}")
    flash('Expense added successfully')
    # If added from dashboard, redirect to dashboard. If from table, table.
    # Default to referrer or dashboard
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/admin/expense/update', methods=['POST'])
@role_required('admin')
def update_expense():
    id = request.form['id']
    amount = request.form['amount']
    description = request.form['description']
    expense_type = request.form['expense_type']
    custom_date = request.form.get('date') # "yyyy-mm-dd"
    expense_year = request.form.get('year')
    
    data = {
        'Amount': amount,
        'Description': description,
        'ExpenseType': expense_type
    }
    if custom_date:
        data['Date'] = custom_date
    if expense_year:
        data['ExpenseYear'] = expense_year

    success = expense_handler.update_expense(id, data)
    
    if success:
        audit_logger.log_action(session.get('username'), session.get('user_role'), 'Update Expense', f"ID: {id}, Data: {data}")
        flash('Updated successfully')
    else:
        flash('Error updating')
    return redirect(request.referrer or url_for('table_page'))

@app.route('/admin/expense/delete/<int:id>', methods=['POST'])
@login_required
def delete_expense(id):
    expense_handler.delete_expense(id)
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Delete Expense', f"ID: {id}")
    flash('Deleted successfully')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/admin/expense/delete_all', methods=['POST'])
@login_required
def delete_all_expenses():
    expense_handler.delete_all_expenses()
    audit_logger.log_action(session.get('username'), session.get('user_role'), 'Delete All Expenses')
    flash('All expenses deleted successfully')
    return redirect(request.referrer or url_for('dashboard'))

if __name__ == '__main__':
    # Run slightly insecurely on 0.0.0.0 to allow mobile testing
    # Note: Microphone access on mobile usually requires HTTPS.
    # To test, users often need to configure browser flags (chrome://flags/#unsafely-treat-insecure-origin-as-secure)
    app.run(debug=True, host='0.0.0.0', port=5000)
