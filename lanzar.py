"""
Levanta el servidor del dashboard.

Uso:
    python lanzar.py                    # procesa los Excel si faltan los CSV y abre en http://localhost:8501
    python lanzar.py --procesar         # vuelve a procesar los Excel aunque ya existan los CSV
    python lanzar.py --tunel            # además abre un enlace temporal https://….trycloudflare.com
    python lanzar.py --tunel --segundo-plano   # (Colab) deja todo corriendo y devuelve el control

El puerto, las carpetas y los colores del tema se toman de config.py.
"""
import argparse
import os
import platform
import re
import shutil
import signal
import subprocess
import sys
import time
import urllib.request

import config as cfg
import procesamiento as pr

LOG_STREAMLIT = cfg.CARPETA_PROYECTO / "streamlit.log"
LOG_TUNEL = cfg.CARPETA_PROYECTO / "cloudflared.log"
URL_CLOUDFLARED = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64"


def tema():
    """Colores de los controles de Streamlit (tema INE), como opciones de línea de comandos."""
    return ["--theme.base", "light",
            "--theme.primaryColor", cfg.AZUL_OSC,
            "--theme.backgroundColor", "#FFFFFF",
            "--theme.secondaryBackgroundColor", cfg.AZUL_CLARO,
            "--theme.textColor", cfg.TEXTO]


def detener_anteriores():
    if platform.system() != "Windows":
        subprocess.run("pkill -f 'streamlit run' ; pkill -f 'cloudflared tunnel'", shell=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(1)


def iniciar_streamlit(segundo_plano):
    cmd = [sys.executable, "-m", "streamlit", "run", str(cfg.CARPETA_PROYECTO / "app.py"),
           "--server.port", str(cfg.PUERTO), "--server.headless", "true",
           "--browser.gatherUsageStats", "false"] + tema()
    if not segundo_plano:
        return subprocess.Popen(cmd, cwd=cfg.CARPETA_PROYECTO)
    # stdin=DEVNULL: si el proceso queda conectado a la terminal de Colab, la celda no termina nunca
    return subprocess.Popen(cmd, cwd=cfg.CARPETA_PROYECTO, stdin=subprocess.DEVNULL, stdout=open(LOG_STREAMLIT, "w"),
                            stderr=subprocess.STDOUT, start_new_session=True)


def binario_cloudflared():
    en_path = shutil.which("cloudflared")
    if en_path:
        return en_path
    local = cfg.CARPETA_PROYECTO / "cloudflared"
    if not local.exists():
        print("Descargando cloudflared...")
        urllib.request.urlretrieve(URL_CLOUDFLARED, local)
        local.chmod(0o755)
    return str(local)


def iniciar_tunel(espera=60):
    """Abre un túnel de Cloudflare al puerto del dashboard y devuelve la URL pública (o None)."""
    subprocess.Popen([binario_cloudflared(), "tunnel", "--url", f"http://localhost:{cfg.PUERTO}", "--no-autoupdate"],
                     stdin=subprocess.DEVNULL, stdout=open(LOG_TUNEL, "w"), stderr=subprocess.STDOUT,
                     start_new_session=True)
    for _ in range(espera):
        time.sleep(1)
        m = re.search(r"https://[a-z0-9-]+\.trycloudflare\.com", LOG_TUNEL.read_text(errors="ignore"))
        if m:
            return m.group(0)
    return None


def main():
    ap = argparse.ArgumentParser(description="Levanta el dashboard demográfico.")
    ap.add_argument("--procesar", action="store_true", help="volver a procesar los Excel del INE")
    ap.add_argument("--tunel", action="store_true", help="abrir un enlace público temporal con Cloudflare")
    ap.add_argument("--segundo-plano", action="store_true", help="no bloquear la terminal (útil en Colab)")
    args = ap.parse_args()

    if args.procesar or not pr.datos_procesados_listos():
        pr.procesar()
    sys.stdout.flush()

    detener_anteriores()
    proc = iniciar_streamlit(args.segundo_plano)
    time.sleep(5)
    print(f"Dashboard local: http://localhost:{cfg.PUERTO}")

    if args.tunel:
        url = iniciar_tunel()
        if url:
            print(f"Enlace público temporal: {url}")
            print("Si al abrirlo aparece un error, esperá unos segundos y recargá la página.")
        else:
            print("No se obtuvo el enlace. Últimas líneas del registro del túnel:")
            print(LOG_TUNEL.read_text(errors="ignore")[-2000:])

    if args.segundo_plano:
        print(f"Corriendo en segundo plano. Registro de Streamlit: {LOG_STREAMLIT}")
        return
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.send_signal(signal.SIGINT)
        detener_anteriores()


if __name__ == "__main__":
    os.chdir(cfg.CARPETA_PROYECTO)
    main()
