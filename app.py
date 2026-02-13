import imaplib
import email
from email.header import decode_header
from flask import Flask, render_template, request, session, redirect, url_for, send_file, flash, jsonify
import io
import zipfile

app = Flask(__name__)
app.secret_key = 'gmail_vault_2026' # Change this to any random string

def get_imap_conn():
    if 'email_user' not in session or 'email_pass' not in session:
        return None
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(session['email_user'], session['email_pass'])
        return mail
    except:
        return None

def decode_str(s):
    if not s: return ""
    decoded_list = decode_header(s)
    header_value, charset = decoded_list[0]
    if isinstance(header_value, bytes):
        return header_value.decode(charset or 'utf-8', errors='ignore')
    return str(header_value)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session['email_user'] = request.form['email']
        session['email_pass'] = request.form['password']
        mail = get_imap_conn()
        if mail:
            mail.logout()
            return redirect(url_for('dashboard'))
        else:
            flash("Login failed. Verify your Gmail address and App Password.")
    return render_template('index.html', page='login')

@app.route('/dashboard')
def dashboard():
    mail = get_imap_conn()
    if not mail: return redirect(url_for('index'))
    _, folders_data = mail.list()
    folders = [f.decode().split(' "/" ')[-1].strip('"') for f in folders_data]
    mail.logout()
    return render_template('index.html', page='dashboard', folders=folders)

@app.route('/get_emails', methods=['POST'])
def get_emails():
    folder = request.form.get('folder', 'INBOX')
    page = int(request.form.get('page', 1))
    per_page = 100
    
    mail = get_imap_conn()
    if not mail: return jsonify({'error': 'Session expired'})
    
    try:
        mail.select(f'"{folder}"', readonly=True)
        _, messages = mail.search(None, 'ALL')
        all_ids = messages[0].split()
        
        # Sort Newest First
        all_ids.reverse()
        
        start = (page - 1) * per_page
        end = start + per_page
        page_ids = all_ids[start:end]
        
        email_list = []
        for num in page_ids:
            # Fetch Subject, From, and Date specifically for performance
            _, data = mail.fetch(num, '(BODY[HEADER.FIELDS (SUBJECT FROM DATE)])')
            msg = email.message_from_bytes(data[0][1])
            email_list.append({
                'id': num.decode(),
                'subject': decode_str(msg["Subject"]),
                'sender': decode_str(msg["From"]),
                'date': decode_str(msg["Date"])
            })
            
        return jsonify({
            'emails': email_list,
            'has_more': end < len(all_ids)
        })
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/download_raw/<folder>/<msg_id>')
def download_raw(folder, msg_id):
    mail = get_imap_conn()
    mail.select(f'"{folder}"', readonly=True)
    _, data = mail.fetch(msg_id, '(RFC822)')
    return send_file(
        io.BytesIO(data[0][1]), 
        mimetype='text/plain', 
        as_attachment=True, 
        download_name=f"email_{msg_id}.txt"
    )

@app.route('/bulk_download', methods=['POST'])
def bulk_download():
    folder = request.form.get('folder')
    msg_ids = request.form.getlist('msg_ids[]')
    mail = get_imap_conn()
    mail.select(f'"{folder}"', readonly=True)
    
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w') as zf:
        for mid in msg_ids:
            _, data = mail.fetch(mid, '(RFC822)')
            zf.writestr(f"email_{mid}.txt", data[0][1])
    
    memory_file.seek(0)
    return send_file(
        memory_file, 
        mimetype='application/zip', 
        as_attachment=True, 
        download_name=f"{folder}_export.zip"
    )

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
