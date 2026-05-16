from flask import Flask, render_template, request, jsonify
import requests
import re

app = Flask(__name__)

def extract_fb_code_via_api(email, refresh_token, client_id):
    try:
        # -------------------------------------------------------------------
        # ১. সেশন ক্লোনিং ও টোকেন এক্সচেঞ্জ (Session Cloning & OAuth Exchange)
        # -------------------------------------------------------------------
        token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        
        # মোবাইল অ্যাপের মতো হুবহু হেডার স্পুফিং (Fingerprint Spoofing)
        token_headers = {
            "User-Agent": "Outlook-Android/4.2415.1 (com.microsoft.office.outlook; build:42415818; Android 12)",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        
        token_data = {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": "https://outlook.office.com/Mail.Read https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
        }
        
        token_res = requests.post(token_url, headers=token_headers, data=token_data, timeout=10)
        if token_res.status_code != 200:
            return "Failed to refresh session token"
            
        access_token = token_res.json().get("access_token")
        if not access_token:
            return "Access token missing in response"

        # -------------------------------------------------------------------
        # ২. হেডার স্পুফিং ও ইন্টারনাল এপিআই রিকোয়েস্ট (Header Spoofing & API GET)
        # -------------------------------------------------------------------
        # আপনার উল্লিখিত ইন্টারনাল OWA/REST API এন্ডপয়েন্ট (সর্বশেষ ৩টি মেইল রিড করার জন্য)
        owa_api_url = "https://outlook.office.com/api/v2.0/me/mailfolders/inbox/messages?$top=3&$select=Subject,Body,From"
        
        api_headers = {
            "User-Agent": "Outlook-Android/4.2415.1 (com.microsoft.office.outlook; build:42415818; Android 12)",
            "Authorization": f"Bearer {access_token}",
            "X-AnchorMailbox": email,  # সরাসরি ইন্টারনাল রাউটিং ইনবক্সে নিয়ে যাওয়ার জন্য
            "Accept": "application/json"
        }
        
        api_res = requests.get(owa_api_url, headers=api_headers, timeout=10)
        if api_res.status_code != 200:
            return f"API Access Denied (Status: {api_res.status_code})"
            
        # মাইক্রোসফট সার্ভার থেকে আসা JSON ডাটা পার্স করা
        json_data = api_res.json()
        messages = json_data.get("value", [])
        
        if not messages:
            return "Inbox is empty"

        # -------------------------------------------------------------------
        # ৩. JSON ডাটা স্ক্র্যাপিং ও রেগুলার এক্সপ্রেশন (Regex Filtering)
        # -------------------------------------------------------------------
        for msg in messages:
            sender_email = msg.get("From", {}).get("EmailAddress", {}).get("Address", "").lower()
            subject = msg.get("Subject", "")
            body_content = msg.get("Body", {}).get("Content", "")
            
            # চেক করা হচ্ছে মেইলটি ফেসবুক থেকে এসেছে কিনা
            if "facebook" in sender_email or "facebook" in subject.lower():
                # মেইলের সাবজেক্ট এবং বডি থেকে ৫ বা ৬ ডিজিটের ওটিপি কোড ম্যাচ করা
                # আপনার লজিক: match(/\d{5,6}/)
                code_match = re.search(r'\b\d{5,6}\b', subject + " " + body_content)
                if code_match:
                    return code_match.group(0) # ১ সেকেন্ডে ফিল্টার করা কোড রিটার্ন করবে
                    
        return "No recent Facebook OTP code found"

    except Exception as e:
        return f"System Error: {str(e)}"

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/get-code', methods=['POST'])
def get_code():
    data = request.json
    raw_input = data.get('raw_input', '')
    
    if not raw_input or '|' not in raw_input:
        return jsonify({'status': 'error', 'message': 'Invalid Input Format'})
    
    parts = raw_input.strip().split('|')
    if len(parts) < 4:
        return jsonify({'status': 'error', 'message': 'Format: email|pass|token|client_id'})
        
    email, password, refresh_token, client_id = parts[0], parts[1], parts[2], parts[3]
    
    # অ্যাডভান্সড এপিআই মেথডে কোড খোঁজা
    fb_code = extract_fb_code_via_api(email, refresh_token, client_id)
    
    if "Error" in fb_code or "Failed" in fb_code or "No recent" in fb_code:
        return jsonify({'status': 'error', 'message': fb_code})
        
    return jsonify({
        'status': 'success',
        'email': email,
        'code': fb_code
    })

if __name__ == '__main__':
    app.run(debug=True, port=5000)
