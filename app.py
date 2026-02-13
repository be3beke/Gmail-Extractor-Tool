import imaplib
import email
from email.header import decode_header
from flask import Flask, render_template, request, session, redirect, url_for, send_file, flash, jsonify
import io
import zipfile

app = Flask(__name__)
app.secret_key = 'super_secret_key_change_this'  # Needed for session storage

# Helper: Connect to Gmail
def get_imap_conn():
    if 'email_user' not in session or 'email_pass' not in session:
        return None
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(session['email_user'], session['email_pass'])
        return mail
    except Exception as e:
        return None

# Helper: Decode email subjects/senders
def decode_str(s):
    if not s: return ""
    decoded_list = decode_header(s)
    header_value, charset = decoded_list[0]
    if isinstance(header_value, bytes):
        try:
            return header_value.decode(charset or 'utf-8')
        except:
            return header_value.decode('utf-8', errors='ignore')
    return str(header_value)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session['email_user'] = request.form['email']
        session['email_pass'] = request.form['password']
        
        # Test connection
        mail = get_imap_conn()
        if mail:
            mail.logout()
            return redirect(url_for('dashboard'))
        else:
            flash("Login failed. Check email or App Password.")
    
    return render_template('index.html', page='login')

@app.route('/dashboard')
def dashboard():
    mail = get_imap_conn()
    if not mail: return redirect(url_for('index'))

    # Get Folders (Labels)
    status, folders_data = mail.list()
    folders = []
    for f in folders_data:
        # Parse folder name (simple parsing)
        name = f.decode().split(' "/" ')[-1].strip('"')
        folders.append(name)
    
    mail.logout()
    return render_template('index.html', page='dashboard', folders=folders)

@app.route('/get_emails', methods=['POST'])
def get_emails():
    folder = request.form.get('folder', 'INBOX')
    mail = get_imap_conn()
    if not mail: return jsonify({'error': 'Not logged in'})

    try:
        mail.select(f'"{folder}"', readonly=True)
        # Search for all emails (Limit to last 20 for speed in this demo)
        status, messages = mail.search(None, 'ALL')
        mail_ids = messages[0].split()[-20:] # Get last 20 emails (Remove [-20:] to get all)
        
        email_list = []
        # Fetch headers only
        for num in reversed(mail_ids):
            _, msg_data = mail.fetch(num, '(RFC822.HEADER)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject = decode_str(msg["Subject"])
                    sender = decode_str(msg["From"])
                    email_list.append({
                        'id': num.decode(),
                        'subject': subject,
                        'sender': sender
                    })
        return jsonify({'emails': email_list})
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/download_raw/<folder>/<msg_id>')
def download_raw(folder, msg_id):
    mail = get_imap_conn()
    if not mail: return redirect(url_for('index'))
    
    mail.select(f'"{folder}"', readonly=True)
    _, data = mail.fetch(msg_id, '(RFC822)')
    raw_email = data[0][1]
    
    return send_file(
        io.BytesIO(raw_email),
        mimetype='message/rfc822',
        as_attachment=True,
        download_name=f"email_{msg_id}.eml"
    )

@app.route('/bulk_download', methods=['POST'])
def bulk_download():
    folder = request.form.get('folder')
    msg_ids = request.form.getlist('msg_ids[]') # Get list of selected IDs
    
    if not msg_ids:
        return "No emails selected", 400

    mail = get_imap_conn()
    mail.select(f'"{folder}"', readonly=True)

    # Create an in-memory zip file
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for mid in msg_ids:
            _, data = mail.fetch(mid, '(RFC822)')
            raw_email = data[0][1]
            zf.writestr(f"email_{mid}.eml", raw_email)
    
    memory_file.seek(0)
    return send_file(
        memory_file,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f"bulk_emails_{folder}.zip"
    )

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
