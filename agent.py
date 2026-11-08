import os
import requests
import msal
from google import genai

CLIENT_ID = os.environ["MS_CLIENT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]

# Inițializăm noul client Google
client = genai.Client(api_key=GEMINI_KEY)

# ==========================================
# 🧠 LOGICA INTELIGENTĂ (ADAPTATĂ PENTRU NOUL SDK)
# ==========================================
model_ales = None
toate_modelele = []

# Cerem lista live de la Google
for m in client.models.list():
    nume = m.name.lower()
    toate_modelele.append(m.name)
    
    # Căutăm cel mai bun model din seria "flash"
    if 'gemini' in nume and 'flash' in nume:
        model_ales = m.name
        break # Am găsit varianta optimă, oprim căutarea

# Dacă nu există "flash", ia primul model "gemini" disponibil
if not model_ales:
    for nume in toate_modelele:
        if 'gemini' in nume.lower():
            model_ales = nume
            break

# Dacă lista e complet goală sau nu recunoaște nimic, printăm diagnosticul
if not model_ales:
    print(f"❌ Eroare: Nu am găsit modele Gemini. Lista completă primită de la Google este: {toate_modelele}")
    exit(1)

print(f"🧠 Agentul a scanat serverele Google și a ales automat modelul: {model_ales}\n" + "="*40)


# ==========================================
# 🔐 CONECTAREA LA MICROSOFT ȘI CITIREA EMAILURILOR
# ==========================================
app = msal.PublicClientApplication(
    CLIENT_ID, 
    authority="https://login.microsoftonline.com/consumers"
)

token_response = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=["Mail.Read"])

if "access_token" in token_response:
    headers = {'Authorization': f'Bearer {token_response["access_token"]}'}
    
    endpoint = "https://graph.microsoft.com/v1.0/me/messages?$top=5"
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code != 200:
        print(f"❌ Eroare Microsoft: {response.status_code}")
        print(response.text)
    else:
        emailuri = response.json().get('value', [])
        
        if not emailuri:
            print("Inbox-ul este gol (nu ai emailuri noi).")
        
        for email in emailuri:
            subiect = email.get('subject', 'Fără subiect')
            corp = email.get('bodyPreview', 'Fără conținut')
            
            prompt = f"Asistent: Rezumă acest email într-o frază scurtă. Urgent sau nu?\nSubiect: {subiect}\nText: {corp}"
            
            try:
                print(f"📧 Subiect Original: {subiect}")
                
                # Trimitem către modelul ales dinamic
                raspuns = client.models.generate_content(
                    model=model_ales,
                    contents=prompt
                )
                print(f"🤖 Analiza Gemini: {raspuns.text}\n{'-'*40}")
                
            except Exception as e:
                print(f"❌ Eroare Gemini la procesarea acestui email: {e}")

else:
    print("❌ Eroare la reînnoirea token-ului Microsoft:", token_response.get("error_description"))