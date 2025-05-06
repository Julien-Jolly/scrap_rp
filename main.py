# redondance dans les if.else ---- OK --------
# refaire commentaires
# reduire timeout ---- OK --------
# format csv "," ---- OK --------
# 30 et 31 jours ---- A TESTER --------
# reformatage avec __init__ ---- OK --------
# ajout concat autres fichiers ---- OK --------
# gestion mdp ---- OK --------
# depot git ---- OK --------
# passer des listes hotel et hotel code à des dict

#python main.py --classements
#python main.py --notes-booking

from playwright.sync_api import sync_playwright
import os
import time
import shutil
import mdp
import concat_book_files
import concat_classement_files
import convert_format_csv
import argparse
from datetime import datetime
from dateutil.relativedelta import relativedelta

hotels = [
    "Amadil Ocean Club Agadir",
    "Atlas Essaouira Riad Resort",
    "Atlas Orient",
    "Atlas Volubilis",
    "Dunes d’Or Ocean Club Agadir",
    "Jaal Riad Resort Marrakech",
    "Kaan Casablanca",
    "Marina Bay City Center Tanger",
    "Palais Medina Riad Resort",
    "Relax Airport",
    "Relax Casa Voyageurs",
    "Relax Kenitra",
    "Relax Marrakech",
    "Relax Oujda",
    "Sky Casa Airport Casablanca",
    "Terminus City Center Oujda",
    "The View Agadir",
    "The View Bouznika",
]

hotels_code = [
    107298,
    148910,
    233862,
    150056,
    161034,
    139880,
    150687,
    150495,
    148875,
    233869,
    430592,
    430598,
    430587,
    453117,
    150665,
    234184,
    110037,
    433555,
]

hotels_id = [
    "A009",
    "A004",
    "A012",
    "S004",
    "S002",
    "S003",
    "TBD",
    "A011",
    "A002",
    "A013",
    "A020",
    "A016",
    "A015",
    "A017",
    "A007",
    "A005",
    "A003",
    "N002",
]

user = os.getenv("USERNAME")  # WINDOWS
PATH = (
    f"c:/users/{user}/onedrive/documents/Experienciah/Extract ReviewPro/"  # WINDOWS FOLDER
)
tmp_folder = f"c:/users/{user}/AppData/Local/Temp/"  # WINDOWS FOLDER

option1_path = "Classements/"
option2_path = "Notes Booking/"
mois_str = ""
chemin = PATH
choix = 1  # 1 = Classements, 2 = Notes booking

# Fonction pour vider le dossier temporaire
def clear_temp_folder(temp_folder):
    session_folder = os.path.join(temp_folder, "playwright_session")
    if os.path.exists(session_folder):
        try:
            shutil.rmtree(session_folder)
            print(f"Dossier {session_folder} vidé avec succès.")
        except Exception as e:
            print(f"Erreur lors du vidage du dossier {session_folder} : {e}")
    else:
        print(f"Dossier {session_folder} n'existe pas, poursuite sans vidage.")

# Fonction pour obtenir le mois précédent
def get_previous_month():
    current_date = datetime.now()
    previous_month = current_date - relativedelta(months=1)
    return f"{previous_month.month:02d}"

# Fonction pour valider les entrées via CLI interactive
def get_user_input():
    global mois_str, choix
    while True:
        try:
            mois = int(input("Entrez le mois à extraire (entre 1 et 12) : "))
            if mois < 1 or mois > 12:
                raise ValueError
            mois_str = f"{mois:02d}"
            break
        except ValueError:
            print("Erreur : Veuillez entrer un mois valide entre 1 et 12.")

    while True:
        try:
            choix_input = input("Choisissez une option (1 = Classements, 2 = Notes booking) : ")
            choix = int(choix_input)
            if choix not in [1, 2]:
                raise ValueError
            break
        except ValueError:
            print("Erreur : Veuillez entrer 1 pour Classements ou 2 pour Notes booking.")

def record_files(tmp_folder, file_old, hotel):
    temp_folder_list = os.listdir(tmp_folder)

    time.sleep(3)
    dossier = [dossier for dossier in temp_folder_list if "artifacts" in dossier]
    print(dossier)
    file = os.listdir(f"{tmp_folder}{dossier[0]}")

    print(f"ancien contenu {file_old}")
    print(f"nouveau contenu {file}")

    attempt = 0
    while set(file) == set(file_old) or any("crdown" in item for item in file):
        file = os.listdir(f"{tmp_folder}{dossier[0]}")
        if attempt % 100000 == 0:
            print(f"    -> attempt n°{attempt} : nouveau contenu {file}")
        attempt += 1

    added_element = set(file) - set(file_old)
    print(f"différence : {added_element}")
    added_element = next(iter(added_element))

    print(f"-> élément téléchargé added_element {added_element}")

    src = f"{tmp_folder}{dossier[0]}/{str(added_element)}"
    dest = f"{PATH}{option_path}{mois_str}-{years[1]}/{hotel}.xlsx"
    shutil.copy(src, dest)

    return file

def download_files(page, url, option_path, years, hotel, choix):
    page.goto(url)
    # Vérifier si on est redirigé vers la page de login
    if "login" in page.url:
        print(f"Redirection inattendue vers la page de login pour {hotel}. Tentative de reconnexion...")
        page.goto("https://app.reviewpro.com/login")
        # Vérifier si le bouton de consentement est présent
        try:
            accept_button = page.get_by_role("button", name="Accept All")
            accept_button.wait_for(timeout=5000)  # Attendre 5 secondes max
            print("Bouton 'Accept All' trouvé, clic en cours...")
            accept_button.click()
        except:
            try:
                accept_button = page.get_by_role("button", name="Tout accepter")
                accept_button.wait_for(timeout=5000)
                print("Bouton 'Tout accepter' trouvé, clic en cours...")
                accept_button.click()
            except:
                print("Aucun bouton de consentement trouvé, poursuite sans clic.")
        page.get_by_label("Email").fill(mdp.id)
        page.get_by_role("button", name="Next").click()
        page.get_by_label("Password").fill(mdp.mdp)
        page.get_by_role("button", name="Log in").click()
        page.wait_for_timeout(5000)  # Attendre 5 secondes après reconnexion
        page.goto(url)  # Réessayer l'URL

    page.get_by_role("button", name="Actions").click()
    page.get_by_role("menuitem", name="Télécharger rapport").click()
    page.wait_for_timeout(3000)  # Attendre le chargement

    if choix == 2:
        page.get_by_role("button", name="Télécharger Le Rapport").click()
    else:
        page.get_by_role("button", name="Download report").click()
    page.wait_for_timeout(3000)  # Attendre que le téléchargement commence

if __name__ == "__main__":
    # Vider le dossier temporaire avant de commencer
    clear_temp_folder(tmp_folder)

    # Configuration des arguments CLI
    parser = argparse.ArgumentParser(description="Script pour extraire des rapports ReviewPro.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--classements", action="store_true", help="Exécuter pour Classements (choix 1) avec le mois précédent")
    group.add_argument("--notes-booking", action="store_true", help="Exécuter pour Notes booking (choix 2) avec le mois précédent")
    args = parser.parse_args()

    # Déterminer le mois et le choix
    if args.classements:
        choix = 1
        mois_str = get_previous_month()  # Mois précédent (ex: "04" pour avril 2025)
        print(f"Mode Classements sélectionné, mois: {mois_str}")
    elif args.notes_booking:
        choix = 2
        mois_str = get_previous_month()  # Mois précédent (ex: "04" pour avril 2025)
        print(f"Mode Notes booking sélectionné, mois: {mois_str}")
    else:
        # Mode interactif si aucun argument n'est fourni
        get_user_input()

    print(f"Mois à extraire : {mois_str}")
    print(f"Chemin de récupération : {chemin}")
    print(f"Choix sélectionné : {choix}")

    years = ["2024", "2025"]
    month_day_start = f"{mois_str}-01"
    months_30_days = {"04", "06", "09", "11"}

    if mois_str == "02":
        day_end = "28"
    elif mois_str in months_30_days:
        day_end = "30"
    else:
        day_end = "31"

    month_day_end = f"{mois_str}-{day_end}"

    with sync_playwright() as p:
        # Utiliser un contexte persistant pour maintenir la session
        context = p.chromium.launch_persistent_context(
            user_data_dir=f"{tmp_folder}playwright_session",
            headless=False
        )
        page = context.new_page()

        # Connexion initiale
        url = "https://app.reviewpro.com/login"
        page.goto(url)
        # Vérifier si le bouton de consentement est présent
        try:
            accept_button = page.get_by_role("button", name="Accept All")
            accept_button.wait_for(timeout=5000)  # Attendre 5 secondes max
            print("Bouton 'Accept All' trouvé, clic en cours...")
            accept_button.click()
        except:
            try:
                accept_button = page.get_by_role("button", name="Tout accepter")
                accept_button.wait_for(timeout=5000)
                print("Bouton 'Tout accepter' trouvé, clic en cours...")
                accept_button.click()
            except:
                print("Aucun bouton de consentement trouvé, poursuite sans clic.")
        page.get_by_label("Email").fill(mdp.id)
        page.get_by_role("button", name="Next").click()
        page.get_by_label("Password").fill(mdp.mdp)
        page.get_by_role("button", name="Log in").click()
        # Attendre un élément spécifique ou un délai fixe
        try:
            page.wait_for_selector('button:has-text("Actions")', timeout=10000)  # Attendre le bouton "Actions" (10s)
        except:
            print("Bouton 'Actions' non trouvé, poursuite avec délai fixe")
            page.wait_for_timeout(5000)  # Attendre 5 secondes comme solution de secours

        file_old = ()

        for i in range(len(hotels)):
            hotel = hotels[i]
            hotel_code = hotels_code[i]

            if choix == 1:
                url = f"https://app.reviewpro.com/reviews/dashboard?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotel_code}"
                option_path = option1_path
            else:
                url = f"https://app.reviewpro.com/reviews/results?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotel_code}&dataGrouping=daily&indexType=GRI"
                option_path = option2_path

            print(f"demarrage extraction {hotel}")
            print(f"url utilisée : {url}")

            if not os.path.exists(f"{PATH}{option_path}{mois_str}-{years[1]}"):
                os.mkdir(f"{PATH}{option_path}{mois_str}-{years[1]}")
                print(
                    f"Création du répertoire : {PATH}{option_path}{mois_str}-{years[1]}"
                )

            if os.path.exists(f"{PATH}{option_path}{mois_str}-{years[1]}/{hotel}.xlsx"):
                print(f"le fichier {hotel} existe déjà")
                print("-------------------------------")
                continue

            download_files(page, url, option_path, years, hotel, choix)
            print(f"enregistrement du fichier {hotel}")
            file_old = record_files(tmp_folder, file_old, hotel)
            print(f"extraction et enregistrement de {hotel} terminés")
            print("---------------------------------------------------")

        context.close()

    year = years[1]
    if choix == 2:
        concat_book_files.concat(PATH, option_path, year, mois_str)
        path = f"{PATH}{option_path}"
        print(path)
        convert_format_csv.convert_csv(path, mois_str, year)
    else:
        concat_classement_files.concat(PATH, option_path, year, mois_str)