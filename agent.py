import os
import requests
import msal
import google.generativeai as genai

# Preluăm parolele din variabilele de mediu
CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
TENANT_ID = os.environ["MS_TENANT_ID"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]

# 1. Configurare Gemini Pro
genai.configure(api_key=GEMINI_KEY)
model = genai.GenerativeModel('gemini-pro')

# 2. Conectarea la Microsoft Graph API
authority = f"https://login.microsoftonline.com/{TENANT_ID}"
app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
)
token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

if "access_token" in token_response:
    print("✅ Autentificare reușită! S-a obținut token-ul.")
    access_token = token_response["access_token"]
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # 3. Interogarea Microsoft Graph
    endpoint = "https://graph.microsoft.com/v1.0/users/george.datcu@hotmail.com/messages?$top=5"
    response = requests.get(endpoint, headers=headers)
    
    # === AICI AM ADĂUGAT DIAGNOSTICAREA ===
    if response.status_code != 200:
        print(f"❌ Microsoft a refuzat cererea! Cod eroare: {response.status_code}")
        print("Detalii eroare:", response.text)
    else:
        emailuri = response.json().get('value', [])
        
        if not emailuri:
            print("⚠️ Conexiunea a reușit, dar Microsoft spune că nu ai niciun email în Inbox (sau nu le vede).")
        
        # 4. Procesarea cu Gemini
        for email in emailuri:
            subiect = email.get('subject', 'Fără subiect')
            corp = email.get('bodyPreview', 'Fără conținut')
            
            prompt = f"Ești un asistent personal. Citește acest email și rezumă-l într-o singură frază, specificând dacă necesită o acțiune urgentă.\nSubiect: {subiect}\nConținut: {corp}"
            
            try:
                raspuns_gemini = model.generate_content(prompt)
                print(f"📧 Subiect Original: {subiect}")
                print(f"🤖 Analiza Gemini: {raspuns_gemini.text}")
                print("-" * 40)
            except Exception as e:
                print(f"❌ Eroare la Gemini: {e}")

else:
    print("❌ Eroare la obținerea tokenului Microsoft:", token_response.get("error_description"))