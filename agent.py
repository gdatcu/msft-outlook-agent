import os
import requests
import msal
from google import genai

CLIENT_ID = os.environ["MS_CLIENT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]

client = genai.Client(api_key=GEMINI_KEY)

# ==========================================
# 🧠 LOGICA SUPER-INTELIGENTĂ DE ALEGERE ȘI TESTARE
# ==========================================
print("🔍 Căutăm și testăm modelele disponibile...")

# Extragem doar modelele care conțin "gemini"
toate_modelele = [m.name for m in client.models.list() if 'gemini' in m.name.lower()]

# Sortăm descrescător ca să luăm versiunile noi primele (ex: 3.8 înaintea lui 2.5)
modele_flash = sorted([m for m in toate_modelele if 'flash' in m.lower()], reverse=True)
modele_restul = sorted([m for m in toate_modelele if m not in modele_flash], reverse=True)
modele_de_incercat = modele_flash + modele_restul

model_ales = None

for nume_model in modele_de_incercat:
    try:
        # Facem un test rapid (ping) ca să vedem dacă Google ne dă voie să-l folosim
        client.models.generate_content(model=nume_model, contents="salut")
        model_ales = nume_model
        break # Dacă a mers fără eroare, ne oprim. Acesta e modelul bun!
    except Exception:
        # Dacă Google ne dă 404 sau altă eroare, trecem silențios la următorul
        continue

if not model_ales:
    print(f"❌ Eroare critică: Google a respins accesul la absolut toate modelele. Lista: {toate_modelele}")
    exit(1)

print(f"🧠 Agentul a testat cu succes și a validat modelul: {model_ales}\n" + "="*40)


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
                
                # Trimitem către modelul validat
                raspuns = client.models.generate_content(
                    model=model_ales,
                    contents=prompt
                )
                print(f"🤖 Analiza Gemini: {raspuns.text}\n{'-'*40}")
                
            except Exception as e:
                print(f"❌ Eroare Gemini la procesarea acestui email: {e}")

else:
    print("❌ Eroare la reînnoirea token-ului Microsoft:", token_response.get("error_description"))