import os
import time
import requests
import msal
from google import genai
import concurrent.futures
from datetime import datetime, timedelta, timezone # Adăugat pentru calculul timpului

CLIENT_ID = os.environ["MS_CLIENT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]
DISCORD_WEBHOOK = os.environ.get("DISCORD_WEBHOOK_URL")

client = genai.Client(api_key=GEMINI_KEY)

# ==========================================
# 🧠 CĂUTARE MODELE CU RETRY ȘI TIMEOUT
# ==========================================
def preia_lista_modele():
    return list(client.models.list())

lista_raw = None
for incercare in range(1, 4):
    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            lista_raw = executor.submit(preia_lista_modele).result(timeout=15)
        break
    except concurrent.futures.TimeoutError:
        time.sleep(2)

if not lista_raw:
    print("❌ Eroare: Serverul Google a picat. Renunțăm.")
    exit(1)

toate_modelele = [m.name for m in lista_raw if 'gemini' in m.name.lower()]
modele_flash = sorted([m for m in toate_modelele if 'flash' in m.lower()], reverse=True)
modele_de_incercat = modele_flash + sorted([m for m in toate_modelele if m not in modele_flash], reverse=True)

model_ales = None
for nume_model in modele_de_incercat:
    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            executor.submit(lambda: client.models.generate_content(model=nume_model, contents="salut")).result(timeout=5)
            model_ales = nume_model
            break
    except Exception:
        pass

if not model_ales:
    exit(1)

# ==========================================
# 🔐 CONECTAREA LA MICROSOFT ȘI CITIREA EMAILURILOR
# ==========================================
app = msal.PublicClientApplication(CLIENT_ID, authority="https://login.microsoftonline.com/consumers")
token_response = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=["Mail.Read"])

if "access_token" in token_response:
    headers = {'Authorization': f'Bearer {token_response["access_token"]}'}
    
    # Calculăm ora de acum 1 oră în format UTC (formatul cerut de Microsoft)
    ora_limita = (datetime.now(timezone.utc) - timedelta(hours=1)).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    # Interogăm DOAR emailurile care au data primirii (receivedDateTime) mai mare sau egală (ge) cu ora_limita
    endpoint = f"https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime ge {ora_limita}"
    
    response = None
    for incercare in range(1, 4):
        try:
            response = requests.get(endpoint, headers=headers, timeout=15)
            break
        except requests.exceptions.Timeout:
            time.sleep(2)
            
    if not response or response.status_code != 200:
        exit(1)
        
    emailuri = response.json().get('value', [])
    if not emailuri:
        print(f"Inbox gol în ultima oră (de la {ora_limita}).")
        exit(0)
    
    # Construim raportul pentru Discord
    mesaj_discord = f"🔔 **Raportul de emailuri (ultimele 60 de minute):**\n\n"
    rezumate_valide = 0
    
    for email in emailuri:
        subiect = email.get('subject', 'Fără subiect')
        corp = email.get('bodyPreview', 'Fără conținut')
        prompt = f"Asistent: Rezumă acest email într-o frază scurtă. Urgent sau nu?\nSubiect: {subiect}\nText: {corp}"
        
        for incercare in range(1, 4):
            try:
                with concurrent.futures.ThreadPoolExecutor() as executor:
                    raspuns = executor.submit(lambda: client.models.generate_content(model=model_ales, contents=prompt)).result(timeout=10)
                    
                mesaj_discord += f"📧 **{subiect}**\n🤖 {raspuns.text}\n\n"
                rezumate_valide += 1
                break
                
            except concurrent.futures.TimeoutError:
                time.sleep(2)
            except Exception:
                break 
                
    # ==========================================
    # 🚀 TRIMITEREA CĂTRE DISCORD
    # ==========================================
    if rezumate_valide > 0 and DISCORD_WEBHOOK:
        # Tăiem mesajul dacă depășește limita de 2000 caractere a Discord-ului
        payload = {"content": mesaj_discord[:1990]}
        
        try:
            req_discord = requests.post(DISCORD_WEBHOOK, json=payload, timeout=10)
            if req_discord.status_code in [200, 204]:
                print("✅ Raport trimis cu succes pe Discord!")
            else:
                print(f"❌ Eroare la trimiterea pe Discord: {req_discord.status_code}")
        except Exception as e:
            print(f"❌ Eroare de rețea către Discord: {e}")