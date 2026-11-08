import os
import time
import requests
import msal
from google import genai
import concurrent.futures
from datetime import datetime, timedelta, timezone

CLIENT_ID = os.environ["MS_CLIENT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]
DISCORD_WEBHOOK = os.environ.get("DISCORD_WEBHOOK_URL")

print("🚀 START AGENT...")
client = genai.Client(api_key=GEMINI_KEY)

# ==========================================
# 🧠 SETĂM UN EXECUTOR GLOBAL (Fără blocaj "with")
# ==========================================
# Acest executor va rula comenzile. Dacă dau timeout, le abandonăm în fundal!
executor = concurrent.futures.ThreadPoolExecutor(max_workers=10)

def preia_lista_modele():
    return list(client.models.list())

print("🔍 Interogăm serverele Google pentru modele...")
lista_raw = None

for incercare in range(1, 4):
    try:
        viitor = executor.submit(preia_lista_modele)
        lista_raw = viitor.result(timeout=15)
        print("✅ Am primit lista de modele!")
        break
    except concurrent.futures.TimeoutError:
        print(f"⏳ Timeout Google (Încercarea {incercare}/3). Reîncercăm...")
        time.sleep(2)

if not lista_raw:
    print("❌ Eroare: Serverul Google a picat complet. Renunțăm.")
    os._exit(1) # Oprește brutal scriptul, ignorând orice proces blocat în fundal

toate_modelele = [m.name for m in lista_raw if 'gemini' in m.name.lower()]
modele_flash = sorted([m for m in toate_modelele if 'flash' in m.lower()], reverse=True)
modele_de_incercat = modele_flash + sorted([m for m in toate_modelele if m not in modele_flash], reverse=True)

model_ales = None
print("🧪 Testăm modelele pentru a găsi unul rapid și funcțional...")
for nume_model in modele_de_incercat:
    try:
        viitor = executor.submit(lambda: client.models.generate_content(model=nume_model, contents="salut"))
        viitor.result(timeout=5)
        model_ales = nume_model
        print(f"🧠 Model validat cu succes: {model_ales}")
        break
    except concurrent.futures.TimeoutError:
        print(f"⚠️ {nume_model} a dat timeout. Sărim peste el.")
    except Exception:
        pass

if not model_ales:
    print("❌ Niciun model nu funcționează.")
    os._exit(1)


# ==========================================
# 🔐 CONECTAREA LA MICROSOFT ȘI CITIREA EMAILURILOR
# ==========================================
print("🔐 Ne conectăm la Microsoft Graph...")
app = msal.PublicClientApplication(CLIENT_ID, authority="https://login.microsoftonline.com/consumers")
token_response = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=["Mail.Read"])

if "access_token" in token_response:
    headers = {'Authorization': f'Bearer {token_response["access_token"]}'}
    
    ora_limita = (datetime.now(timezone.utc) - timedelta(hours=1)).strftime('%Y-%m-%dT%H:%M:%SZ')
    print(f"📅 Căutăm doar emailuri noi primite de la ora: {ora_limita}")
    
    endpoint = f"https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime ge {ora_limita}"
    
    response = None
    for incercare in range(1, 4):
        try:
            response = requests.get(endpoint, headers=headers, timeout=15)
            break
        except requests.exceptions.Timeout:
            print(f"⏳ Timeout Microsoft (Încercarea {incercare}/3)...")
            time.sleep(2)
            
    if not response or response.status_code != 200:
        print("❌ Eroare la citirea emailurilor.")
        os._exit(1)
        
    emailuri = response.json().get('value', [])
    if not emailuri:
        print("📭 Inbox gol în ultima oră.")
        os._exit(0) # Gata cu succes! Oprim brutal pentru a preveni blocajele
    
    print(f"📥 Am găsit {len(emailuri)} email(uri) nou(i)! Începem procesarea...")
    
    mesaj_discord = f"🔔 **Raportul de emailuri (ultimele 60 de minute):**\n\n"
    rezumate_valide = 0
    
    for email in emailuri:
        subiect = email.get('subject', 'Fără subiect')
        corp = email.get('bodyPreview', 'Fără conținut')
        prompt = f"Asistent: Rezumă acest email într-o frază scurtă. Urgent sau nu?\nSubiect: {subiect}\nText: {corp}"
        
        print(f"📧 Procesăm: {subiect}")
        
        for incercare in range(1, 4):
            try:
                viitor = executor.submit(lambda: client.models.generate_content(model=model_ales, contents=prompt))
                raspuns = viitor.result(timeout=10)
                mesaj_discord += f"📧 **{subiect}**\n🤖 {raspuns.text}\n\n"
                rezumate_valide += 1
                break
            except concurrent.futures.TimeoutError:
                print(f"   ⏳ Timeout Gemini (Încercarea {incercare}/3)")
                time.sleep(2)
            except Exception as e:
                print(f"❌ Eroare Gemini la acest email: {e}")
                break 
                
    # ==========================================
    # 🚀 TRIMITEREA CĂTRE DISCORD
    # ==========================================
    if rezumate_valide > 0 and DISCORD_WEBHOOK:
        print("🚀 Trimitem raportul pe Discord...")
        payload = {"content": mesaj_discord[:1990]}
        
        try:
            req_discord = requests.post(DISCORD_WEBHOOK, json=payload, timeout=10)
            if req_discord.status_code in [200, 204]:
                print("✅ Raport trimis cu succes pe Discord!")
            else:
                print(f"❌ Eroare Discord: {req_discord.status_code}")
        except Exception as e:
            print(f"❌ Eroare rețea Discord: {e}")
else:
    print("❌ Eroare conectare Microsoft.")

print("🏁 Rularea s-a terminat.")
os._exit(0) # Închidem procesul instant, ucigând orice fire rătăcite în fundal