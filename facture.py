import os
import csv
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()
EMAIL = os.getenv("FACTURE_EMAIL")
PASSWORD = os.getenv("FACTURE_PASSWORD")
COMPANY_ID = os.getenv("FACTURE_COMPANY_ID")
CLIENT_ID = os.getenv("FACTURE_CLIENT_ID")
TAUX_HORAIRE = float(os.getenv("FACTURE_TAUX_HORAIRE", "45"))
CSV_FILE = "facture.csv"
BANK_ACCOUNT_ID = os.getenv("FACTURE_BANK_ACCOUNT_ID")

tasks = []
with open(CSV_FILE, newline='', encoding='utf-8') as f:
    lines = f.readlines()
    if len(lines) < 3:
        raise Exception("Le fichier CSV est trop court")

    period_line = lines[1].strip().strip('"\'')
    if not period_line.startswith("Period:"):
        raise Exception("La deuxième ligne ne contient pas la période: " + period_line)

    # Exemple : "Period: Jul 1 - Jul 31"
    _, period_raw = period_line.split(":", 1)
    start_str, end_str = period_raw.strip().split(" - ")
    year = datetime.now().year

    start_date = datetime.strptime(f"{start_str} {year}", "%b %d %Y")
    end_date = datetime.strptime(f"{end_str} {year}", "%b %d %Y")

    DESCRIPTION_PREFIX = f"Freelance du {start_date.day} {start_date.strftime('%B')} {start_date.year} au {end_date.day} {end_date.strftime('%B')} {end_date.year}"

    mois_fr = {
        'January': 'janvier', 'February': 'février', 'March': 'mars', 'April': 'avril',
        'May': 'mai', 'June': 'juin', 'July': 'juillet', 'August': 'août',
        'September': 'septembre', 'October': 'octobre', 'November': 'novembre', 'December': 'décembre'
    }
    DESCRIPTION_PREFIX = DESCRIPTION_PREFIX.replace(start_date.strftime('%B'), mois_fr[start_date.strftime('%B')])
    DESCRIPTION_PREFIX = DESCRIPTION_PREFIX.replace(end_date.strftime('%B'), mois_fr[end_date.strftime('%B')])

    f.seek(0)
    lines = f.readlines()
    reader = csv.DictReader(lines[2:])
    for row in reader:
        if row['Member'].strip().lower() == "total":
            continue
        time_hours = round(float(row['Time']), 2)
        hours = int(time_hours)
        minutes = int(round((time_hours - hours) * 60))
        if hours == 0:
            description_time = f"{minutes}min"
        else:
            description_time = f"{hours}h{minutes:02d}m"
        description = f"{row['Task']} – {description_time}"
        tasks.append({
            "description": description,
            "quantite": time_hours,
            "prix_unitaire": TAUX_HORAIRE
        })

if not tasks:
    raise Exception("Aucune tâche trouvée dans le CSV")

driver = webdriver.Chrome()
driver.get("https://www.facture.net/login")

driver.find_element(By.ID, "user_email").send_keys(EMAIL)
driver.find_element(By.ID, "user_password").send_keys(PASSWORD + Keys.RETURN)

time.sleep(1.5)

driver.get(f"https://www.facture.net/{COMPANY_ID}/bills/new")
time.sleep(1.5)

driver.find_element(By.CSS_SELECTOR, '[data-action="click->cms--popup-component#postHide:stop"]').click()
time.sleep(1)

client_input = driver.find_element(By.CLASS_NAME, "combo-select")
client_input.click()
time.sleep(2)

xo7_div = driver.find_element(By.XPATH, f"//div[@class='item' and @data-value='{CLIENT_ID}']")
xo7_div.click()

time.sleep(3)

for i, task in enumerate(tasks):
    if i > 0:
        driver.find_element(By.CSS_SELECTOR, "a.add_fields.btn-primary").click()
        time.sleep(0.5)

    description = driver.find_elements(By.CSS_SELECTOR, "[id^='bill_items_attributes_'][id$='_description']")[i]
    quantity = driver.find_elements(By.CSS_SELECTOR, "[id^='bill_items_attributes_'][id$='_quantity']")[i]
    price = driver.find_elements(By.CSS_SELECTOR, "[id^='bill_items_attributes_'][id$='_unit_price']")[i]

    description.send_keys(task["description"])
    time.sleep(0.1)
    quantity.clear(); quantity.send_keys(str(task["quantite"]))
    time.sleep(0.1)
    price.clear(); price.send_keys(str(task["prix_unitaire"]))
    time.sleep(0.1)

time.sleep(1)

dropdown_div = driver.find_element(By.CSS_SELECTOR, 'div.form-input.select.optional.selection.dropdown.dropdown-with-blank')
dropdown_div.click()

time.sleep(1)

option_div = driver.find_element(By.CSS_SELECTOR, 'div.item[data-value="'+BANK_ACCOUNT_ID+'"]') 
option_div.click()

time.sleep(1)

driver.find_element(By.ID, "bill_header").send_keys(DESCRIPTION_PREFIX + Keys.RETURN)

submit_button = driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="Créer la facture"]')
submit_button.click()

print("Facture créée depuis la feuille de temps !")
time.sleep(3)
driver.quit()