import requests
from bs4 import BeautifulSoup
from datetime import datetime
import os
import json
import smtplib
from email.mime.text import MIMEText
from openpyxl import Workbook, load_workbook

EXCEL_FILE = "metal_rates.xlsx"
HISTORY_FILE = "last_prices.json"

METALS = ["Gold 24K", "Gold 22K", "Gold 18K", "Gold 14K", "Silver", "Silver Bar", "Platinum"]

API_URL = "https://goldpriceeditor.droidinfinity.com/api/external/metal-prices/1085"

HEADERS = {
    "accept": "application/json",
    "origin": "https://pngadgilandsons.com",
    "referer": "https://pngadgilandsons.com/",
    "user-agent": "Mozilla/5.0"
}

def send_email(subject, body):
    sender = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")

    if not sender or not password:
        print("No email credentials")
        return

    password = password.replace(" ", "")

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = sender

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        print("Email Sent")
    except Exception as e:
        print("Email Error:", e)

def load_last_prices():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return json.load(f)
    return {}

def save_last_prices(data):
    with open(HISTORY_FILE, "w") as f:
        json.dump(data, f)

def get_rates():
    try:
        res = requests.get(API_URL, headers=HEADERS)
        data = res.json()["rates"]

        return {
            "Gold 24K": str(data["goldPrice24K"]),
            "Gold 22K": str(data["goldPrice22K"]),
            "Gold 18K": str(data["goldPrice18K"]),
            "Gold 14K": str(data["goldPrice14K"]),
            "Silver": str(data["silverPrice"]),
            "Silver Bar": str(data["silverBarPrice"]),
            "Platinum": str(data["platinumPrice"])
        }
    except:
        return None

def save_excel(data):
    now = datetime.now()

    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Time"] + METALS)
        wb.save(EXCEL_FILE)

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    row = [now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")]

    for m in METALS:
        row.append(data.get(m, ""))

    ws.append(row)
    wb.save(EXCEL_FILE)

def main():
    data = get_rates()
    if not data:
        return

    save_excel(data)

    last = load_last_prices()

    if last:
        changes = []
        for m in METALS:
            if str(last.get(m)) != str(data.get(m)):
                changes.append(f"{m}: {last.get(m)} → {data.get(m)}")

        if changes:
            send_email("Price Changed", "\n".join(changes))

    save_last_prices(data)

if __name__ == "__main__":
    main()
