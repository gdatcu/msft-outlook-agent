import os
import requests
import msal
import google.generativeai as genai

CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]

genai.configure(api_key=GEMINI_KEY)
model = genai.GenerativeModel('gemini-pro')

# Conectarea folosind Refresh Token-ul pentru conturi personale
app = msal.ConfidentialClientApplication(
    CLIENT_ID, 
    authority="https://login.microsoftonline.com/consumers",
    client_credential=CLIENT_SECRET
)

token_response = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=["Mail.Read"])

if "access_token" in token_response:
    print("✅ Autentificare reușită ca utilizator!")
    headers = {'Authorization': f'Bearer {token_response["access_token"]}'}
    
    endpoint = "https://graph.microsoft.com/v1.0/me/messages?$top=5"
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code != 200:
        print(f"❌ Eroare Microsoft: {response.status_code}")
        print(response.text)
    else:
        emailuri = response.json().get('value', [])
        
        if not emailuri:
            print("Inbox-ul este gol (nu sunt emailuri).")
        
        for email in emailuri:
            subiect = email.get('subject', 'Fără subiect')
            corp = email.get('bodyPreview', 'Fără conținut')
            
            prompt = f"Asistent: Rezumă acest email într-o frază scurtă. Urgent sau nu?\nSubiect: {subiect}\nText: {corp}"
            
            try:
                print(f"📧 {subiect}")
                raspuns = model.generate_content(prompt)
                print(f"🤖 {raspuns.text}\n{'-'*40}")
            except Exception as e:
                print(f"❌ Eroare Gemini: {e}")

else:
    print("❌ Eroare la generarea token-ului:", token_response.get("error_description"))