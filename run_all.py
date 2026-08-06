### Run all servers (API unifiée)

"""
============================================================
 RUN ALL — API unique + Streamlit
============================================================
Usage :
    python run_all.py

Lance :
  - reconc.py   (port 8002) — gateway + tous les partenaires
                montés sous /svc/...
  - streamlit   (port 8501)

Module_FED continue d'appeler http://127.0.0.1:8002 uniquement.
Ctrl+C arrête proprement tous les sous-processus.
============================================================
"""

import subprocess
import sys
import signal
import threading
import time

# En prod, nginx reverse-proxifie vers localhost (voir deploy/).
SERVICES = [
    ("reconc", ["uvicorn", "gateway.reconc:app", "--host", "127.0.0.1", "--port", "8002"]),
    (
        "streamlit",
        [
            "streamlit",
            "run",
            "ui/streamlit_app.py",
            "--server.headless",
            "true",
            "--server.address",
            "127.0.0.1",
            "--server.port",
            "8501",
        ],
    ),
]

processus = []
arret_demande = threading.Event()


def lire_et_prefixer(nom, pipe):
    """Lit la sortie d'un sous-processus ligne par ligne et l'affiche préfixée par son nom."""
    for ligne in iter(pipe.readline, ""):
        if not ligne:
            break
        print(f"[{nom}] {ligne.rstrip()}")
    pipe.close()


def demarrer_service(nom, commande):
    print(f"[run_all] Démarrage de '{nom}' : {' '.join(commande)}")
    p = subprocess.Popen(
        commande,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    t = threading.Thread(target=lire_et_prefixer, args=(nom, p.stdout), daemon=True)
    t.start()
    return p


def arreter_tout(*_args):
    if arret_demande.is_set():
        return
    arret_demande.set()

    print("\n[run_all] Arrêt demandé — extinction de tous les services...")

    for nom, p in processus:
        if p.poll() is None:
            print(f"[run_all] Arrêt de '{nom}' (pid={p.pid})")
            p.terminate()

    deadline = time.time() + 5
    for nom, p in processus:
        delai_restant = max(0, deadline - time.time())
        try:
            p.wait(timeout=delai_restant)
        except subprocess.TimeoutExpired:
            print(f"[run_all] '{nom}' ne répond pas, arrêt forcé (kill).")
            p.kill()

    print("[run_all] Tous les services sont arrêtés.")
    sys.exit(0)


def main():
    signal.signal(signal.SIGINT, arreter_tout)
    signal.signal(signal.SIGTERM, arreter_tout)

    for nom, commande in SERVICES:
        processus.append((nom, demarrer_service(nom, commande)))
        time.sleep(1)

    print("\n[run_all] API unifiée sur http://127.0.0.1:8002 — Ctrl+C pour arrêter.\n")

    deja_signales = set()

    try:
        while not arret_demande.is_set():
            for nom, p in processus:
                if p.poll() is not None and nom not in deja_signales:
                    print(f"[run_all] ATTENTION : le service '{nom}' s'est arrêté (code {p.returncode}).")
                    deja_signales.add(nom)
            time.sleep(2)
    except KeyboardInterrupt:
        arreter_tout()


if __name__ == "__main__":
    main()
