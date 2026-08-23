import requests

url = "http://127.0.0.1:8002/charger"
fichier = r"C:\Users\ndick.faye\Documents\RECONCOR\reconciliation\Daily-ChannelUserTransactionReport-786256338-20260730 USSD ORANGE.xls"

params = {
    "partenaire": "ORANGE_USSD",
    "date_debut": "01/01/2000",
    "date_fin": "01/01/2100",
}

with open(fichier, "rb") as f:
    r = requests.post(url, params=params, files={"files": f})

print(r.status_code)
print(r.text)