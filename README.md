# Reconciliation Flexcube — API unifiée

## Lancer (local)

```bash
source .venv/bin/activate
python run_all.py
```

- API : http://127.0.0.1:8002
- Streamlit : http://localhost:8501
- Secrets : `.env` (voir `.env.example`)

## Déploiement prod (nginx + systemd)

Sur le serveur (Debian/Ubuntu), depuis le dépôt :

```bash
sudo bash deploy/install.sh /opt/reconciliation reconc.votredomaine.local
```

Puis renseigner Oracle dans `/opt/reconciliation/.env` et redémarrer :

```bash
sudo nano /opt/reconciliation/.env
sudo systemctl restart reconciliation
```

- UI : `http://<serveur>/`
- API (via nginx) : `http://<serveur>/health`, `/svc/...`, `/charger`, …
- Module_FED sur la même machine peut toujours appeler `http://127.0.0.1:8002`
- Logs : `journalctl -u reconciliation -f`

Fichiers : `deploy/nginx-reconciliation.conf`, `deploy/reconciliation.service`, `deploy/install.sh`.

## Structure

```
.
├── run_all.py              # Lance API + Streamlit
├── reconc.py               # Entrée uvicorn (compat Module_FED)
├── config.py               # Shim → common/config.py
├── requirements.txt
├── .env / .env.example
│
├── common/                 # Code partagé
│   ├── config.py
│   ├── excel_common.py
│   ├── flex_common.py
│   ├── sqlite_io.py
│   ├── http_export.py
│   ├── reconciliation_engine.py
│   └── nettoyage_ussd_orange.py
│
├── gateway/                # API gateway
│   ├── reconc.py
│   └── mount_partners.py
│
├── partners/               # Un dossier par partenaire
│   ├── wave/
│   ├── wave_agence/
│   ├── orange/
│   ├── orange_ussd/
│   ├── wizz/
│   └── ria_agence/
│
├── ui/                     # Streamlit
│   ├── streamlit_app.py
│   └── credentials.json
│
├── data/
│   ├── base.db
│   └── mappings/           # Référentiels agences / MSISDN
│
└── _archive/               # Anciens fichiers (hors runtime)
```

## Ajouter un partenaire

1. Créer `partners/<nom>/` (Excel + Flex)
2. Entrées dans `common/config.py` (`PARTENAIRES`) et `gateway/mount_partners.py`
