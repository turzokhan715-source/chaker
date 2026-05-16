from flask import Flask, render_template, request, jsonify
import requests
import imaplib
import re

app = Flask(__name__)

def get_access_token(client_id, refresh_token):
    """রিফ্রেশ টোকেন দিয়ে মাইক্রোসফট থেকে নতুন এক্সেস টোকেন নেওয়া"""
    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    payload = {
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access'
    }
    try:
        response = requests.post(url, data=payload, timeout=10).json()
        return response.get('access_token')
    except Exception:
        return None

def fetch_facebook_code(email, access_token):
    """IMAP এর মাধ্যমে ইনবক্স থেকে ফেসবুকের কোড রিড করা"""
    try:
        mail = imaplib.IMAP4_SSL("outlook.office365.com")
        auth_string = f"user={email}\x01auth=Bearer {access_token}\x01\x01"
        mail.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))
        
        mail.select("inbox")
        
        # ফেসবুকের পাঠানো ভেরিফিকেশন মেইল ফিল্টার করা
        status, messages = mail.search(None, '(FROM "facebookmail.com")')
        
        if status == 'OK' and messages[0]:
            latest_email_id = messages[0].split()[-1]
            status, data = mail.fetch(latest_email_id, '(RFC822)')
            email_content = data[0][1].decode('utf-8', errors='ignore')
            
            # ৫ বা ৬ ডিজিটের ফেসবুক ওটিপি কোড খুঁজে বের করা
            code_match = re.search(r'\b\d{5,6}\b', email_content)
            if code_match:
                return code_match.group(0)
                
        return "No code found"
    except Exception as e:
        return f"Error: {str(e)}"

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/get-code', methods=['POST'])
def get_code():
    data = request.json
    raw_input = data.get('raw_input', '')
    
    if not raw_input or '|' not in raw_input:
        return jsonify({'status': 'error', 'message': 'Invalid format'})
    
    parts = raw_input.strip().split('|')
    if len(parts) < 4:
        return jsonify({'status': 'error', 'message': 'Format: email|pass|token|client_id'})
        
    email, password, refresh_token, client_id = parts[0], parts[1], parts[2], parts[3]
    
    access_token = get_access_token(client_id, refresh_token)
    if not access_token:
        return jsonify({'status': 'error', 'message': 'Failed to get access token.'})
    
    fb_code = fetch_facebook_code(email, access_token)
    
    return jsonify({
        'status': 'success',
        'email': email,
        'code': fb_code
    })

if __name__ == '__main__':
    app.run(debug=True, port=5000)
