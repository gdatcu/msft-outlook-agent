import os
import requests
import msal
import google.generativeai as genai

# Preluăm parolele din variabilele de mediu (pentru securitate)
CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
TENANT_ID = os.environ["MS_TENANT_ID"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]

# 1. Configurare Gemini Pro
genai.configure(api_key=GEMINI_KEY)
model = genai.GenerativeModel('gemini-pro') # Poți folosi și variantele flash

# 2. Conectarea la Microsoft Graph API (Obținerea token-ului de acces)
authority = f"https://login.microsoftonline.com/{TENANT_ID}"
app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
)
token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

if "access_token" in token_response:
    access_token = token_response["access_token"]
    headers = {'Authorization': f'Bearer {access_token}'}
    
    # 3. Citirea ultimelor 5 emailuri din Inbox
    endpoint = "https://graph.microsoft.com/v1.0/users/george.datcu@hotmail.com/messages?$top=5"
    response = requests.get(endpoint, headers=headers)
    emailuri = response.json().get('value', [])
    
    # 4. Procesarea cu Gemini
    for email in emailuri:
        subiect = email['subject']
        corp = email['bodyPreview']
        
        prompt = f"Ești un asistent personal. Citește acest email și rezumă-l într-o singură frază, specificând dacă necesită o acțiune urgentă.\nSubiect: {subiect}\nConținut: {corp}"
        
        raspuns_gemini = model.generate_content(prompt)
        
        print(f"📧 Subiect Original: {subiect}")
        print(f"🤖 Analiza Gemini: {raspuns_gemini.text}")
        print("-" * 40)
else:
    print("Eroare la autentificarea Microsoft:", token_response.get("error_description"))