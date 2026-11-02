import os
import sys
import subprocess
from http.server import HTTPServer, BaseHTTPRequestHandler
from google_auth_oauthlib.flow import InstalledAppFlow
from auth.credential_store import get_credential_store
from auth.scopes import BASE_SCOPES, GMAIL_SCOPES, DRIVE_SCOPES, DOCS_SCOPES, SHEETS_SCOPES
from googleapiclient.discovery import build

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
TARGET_CONFIG_PATH = os.path.expanduser(r"~/.gemini/config/google_client_secret.json")

class OAuthCallbackHandler(BaseHTTPRequestHandler):
    auth_response = None

    def log_message(self, format, *args):
        pass  # Quiet logging

    def do_GET(self):
        OAuthCallbackHandler.auth_response = f"https://localhost:8080{self.path}"
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Yetkilendirme Başarılı</title>
            <style>
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; text-align: center; padding: 60px; background-color: #f8f9fa; }
                .card { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); max-width: 480px; margin: 0 auto; }
                h1 { color: #2e7d32; font-size: 24px; margin-bottom: 16px; }
                p { color: #555; font-size: 16px; line-height: 1.5; }
            </style>
        </head>
        <body>
            <div class="card">
                <h1>✅ Yetkilendirme Başarılı!</h1>
                <p>Google Workspace hesabınız Antigravity'ye başarıyla bağlandı.</p>
                <p>Bu sekmeyi kapatıp sohbete geri dönebilirsiniz.</p>
            </div>
        </body>
        </html>
        """
        self.wfile.write(html_content.encode('utf-8'))

def main():
    sys.stdout.reconfigure(line_buffering=True)
    scopes = list(dict.fromkeys(BASE_SCOPES + GMAIL_SCOPES + DRIVE_SCOPES + DOCS_SCOPES + SHEETS_SCOPES))
    flow = InstalledAppFlow.from_client_secrets_file(TARGET_CONFIG_PATH, scopes)
    flow.redirect_uri = "http://localhost:8080/"
    
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    
    with open("c:/isg/auth_url.txt", "w", encoding="utf-8") as f:
        f.write(auth_url)
        
    print(f"AUTH_URL_READY: {auth_url}", flush=True)

    # Windows'ta varsayılan tarayıcıda doğrudan aç
    try:
        subprocess.Popen(["cmd.exe", "/c", "start", "", auth_url], shell=False)
    except Exception as e:
        print(f"Tarayıcı açılırken uyarı: {e}", flush=True)

    # Callback dinle (tek bir istek karşılayacak)
    server = HTTPServer(('localhost', 8080), OAuthCallbackHandler)
    server.handle_request()

    if not OAuthCallbackHandler.auth_response:
        print("HATA: Yetki yanıtı alınamadı.", flush=True)
        sys.exit(1)

    flow.fetch_token(authorization_response=OAuthCallbackHandler.auth_response)
    creds = flow.credentials

    service = build('oauth2', 'v2', credentials=creds)
    user_info = service.userinfo().get().execute()
    email = user_info.get('email', 'default_user')

    store = get_credential_store()
    store.store_credential(email, creds)
    
    print(f"SUCCESS: {email} başarıyla kaydedildi!", flush=True)

if __name__ == "__main__":
    main()
