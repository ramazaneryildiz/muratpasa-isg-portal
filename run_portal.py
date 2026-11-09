import os
import sys
import time
import socket
import subprocess
import webbrowser
import threading
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_EXE = sys.executable

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def main():
    print("=" * 68)
    print("      MURATPAŞA İLÇE MİLLÎ EĞİTİM MÜDÜRLÜĞÜ - İSG PORTALI")
    print("=" * 68)
    
    local_ip = get_local_ip()
    cloudflared_bin = os.path.join(BASE_DIR, "cloudflared.exe")
    tunnel_file = os.path.join(BASE_DIR, "tunnel_url.txt")
    
    # 1. Cloudflare Tünelini Başlat (Mobil SSL Erişimi İçin)
    tunnel_process = None
    tunnel_url = None
    
    if os.path.exists(cloudflared_bin):
        print("\n[*] Mobil cihazlar için güvenli SSL tüneli kuruluyor...")
        try:
            tunnel_process = subprocess.Popen(
                [cloudflared_bin, "tunnel", "--url", "http://127.0.0.1:8000"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace"
            )
            
            # Tünel URL'sini yakala
            start_time = time.time()
            while time.time() - start_time < 12:
                line = tunnel_process.stdout.readline()
                if not line:
                    break
                match = re.search(r"https://[a-zA-Z0-9\-]+\.trycloudflare\.com", line)
                if match:
                    tunnel_url = match.group(0)
                    with open(tunnel_file, "w", encoding="utf-8") as f:
                        f.write(tunnel_url)
                    break
        except Exception as e:
            print(f"[!] Tünel başlatılamadı: {e}")

    # 2. Terminal Bilgilendirme Ekranı
    print("\n" + "=" * 68)
    print("                      PORTAL BAŞARIYLA BAŞLATILDI")
    print("=" * 68)
    print(f" [*] Bilgisayar Erişimi (Yerel) : http://localhost:8000")
    print(f" [*] Aynı Wi-Fi Mobil Erişimi   : http://{local_ip}:8000/saha-asistani")
    if tunnel_url:
        print(f" [*] Her Yerden Mobil Erişim (SSL): {tunnel_url}/saha-asistani")
    print("-" * 68)
    print(" [i] Varsayılan Yetkili Girişi : ramazan.eryildiz / isg2026!")
    print(" [i] Kapatmak için bu pencereyi kapatabilir veya Ctrl+C yapabilirsiniz.")
    print("=" * 68 + "\n")

    # 3. Tarayıcıyı Otomatik Aç
    def open_browser():
        time.sleep(1.5)
        webbrowser.open("http://localhost:8000")
    
    threading.Thread(target=open_browser, daemon=True).start()

    # 4. Uvicorn Web Sunucusunu Başlat
    import uvicorn
    try:
        uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False, log_level="info")
    except KeyboardInterrupt:
        print("\n[*] Sunucu kapatılıyor...")
    finally:
        if tunnel_process:
            try:
                tunnel_process.terminate()
            except Exception:
                pass
        print("[OK] Sistem güvenle kapatıldı.")

if __name__ == "__main__":
    main()
