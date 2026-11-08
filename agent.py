import os
import time
import requests
import msal
from google import genai
import concurrent.futures

CLIENT_ID = os.environ["MS_CLIENT_ID"]
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]

client = genai.Client(api_key=GEMINI_KEY)

# ==========================================
# 🧠 CĂUTARE MODELE CU RETRY ȘI TIMEOUT
# ==========================================
print("🔍 Căutăm și testăm modelele disponibile...")

def preia_lista_modele():
    return list(client.models.list())

lista_raw = None
for incercare in range(1, 4): # Încearcă de 3 ori
    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            lista_raw = executor.submit(preia_lista_modele).result(timeout=15)
        break # Succes, ieșim din bucla de reîncercare
    except concurrent.futures.TimeoutError:
        print(f"⏳ Google API timeout la listare. Încercarea {incercare}/3...")
        time.sleep(2)

if not lista_raw:
    print("❌ Eroare: Serverul Google nu a răspuns la listare după 3 încercări. Renunțăm.")
    exit(1)

toate_modelele = [m.name for m in lista_raw if 'gemini' in m.name.lower()]
modele_flash = sorted([m for m in toate_modelele if 'flash' in m.lower()], reverse=True)
modele_restul = sorted([m for m in toate_modelele if m not in modele_flash], reverse=True)
modele_de_incercat = modele_flash + modele_restul

model_ales = None

# La testare nu facem retry per model, ci trecem rapid la următorul din listă dacă dă timeout
for nume_model in modele_de_incercat:
    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Testăm trimițând un salut simplu
            viitor = executor.submit(lambda: client.models.generate_content(model=nume_model, contents="salut"))
            viitor.result(timeout=5)
            model_ales = nume_model
            break # Am validat modelul!
    except concurrent.futures.TimeoutError:
        print(f"⚠️ {nume_model} a dat timeout. Trecem la următorul...")
    except Exception:
        # A dat eroare 404 sau altele (blocat de Google), sărim silențios
        pass

if not model_ales:
    print("❌ Eroare critică: Niciun model nu a trecut testul de funcționare.")
    exit(1)

print(f"🧠 Agentul a validat modelul: {model_ales}\n" + "="*40)


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
    
    # Retry pentru Microsoft Graph API
    response = None
    for incercare in range(1, 4):
        try:
            response = requests.get(endpoint, headers=headers, timeout=15)
            break
        except requests.exceptions.Timeout:
            print(f"⏳ Microsoft timeout. Încercarea {incercare}/3...")
            time.sleep(2)
            
    if not response:
        print("❌ Eroare: Microsoft nu a răspuns după 3 încercări.")
        exit(1)
        
    if response.status_code != 200:
        print(f"❌ Eroare Microsoft: {response.status_code}\n{response.text}")
    else:
        emailuri = response.json().get('value', [])
        
        if not emailuri:
            print("Inbox-ul este gol (nu ai emailuri noi).")
        
        for email in emailuri:
            subiect = email.get('subject', 'Fără subiect')
            corp = email.get('bodyPreview', 'Fără conținut')
            prompt = f"Asistent: Rezumă acest email într-o frază scurtă. Urgent sau nu?\nSubiect: {subiect}\nText: {corp}"
            
            print(f"📧 Subiect Original: {subiect}")
            
            # Retry pentru generarea rezumatului cu Gemini
            rezumat_generat = False
            for incercare in range(1, 4):
                try:
                    with concurrent.futures.ThreadPoolExecutor() as executor:
                        viitor = executor.submit(lambda: client.models.generate_content(model=model_ales, contents=prompt))
                        raspuns = viitor.result(timeout=10)
                        
                    print(f"🤖 Analiza Gemini: {raspuns.text}\n{'-'*40}")
                    rezumat_generat = True
                    break # Succes, ieșim din loop-ul de reîncercări
                    
                except concurrent.futures.TimeoutError:
                    print(f"   [⏳ Așteptăm Gemini... încercarea {incercare}/3]")
                    time.sleep(2)
                except Exception as e:
                    print(f"❌ Eroare Gemini definitivă la acest email: {e}\n{'-'*40}")
                    rezumat_generat = True # E o eroare internă a modelului (ex. text prea lung), nu are sens retry-ul
                    break 
                    
            if not rezumat_generat:
                print("⚠️ Am renunțat la acest email după 3 încercări (timeout).\n" + "-"*40)

else:
    print("❌ Eroare la reînnoirea token-ului Microsoft:", token_response.get("error_description"))