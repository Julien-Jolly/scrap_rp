# redondance dans les if.else ---- OK --------
# refaire commentaires
# reduire timeout ---- OK --------
# format csv "," ---- OK --------
# 30 et 31 jours ---- A TESTER --------
# reformatage avec __init__ ---- OK --------
# ajout concat autres fichiers ---- OK --------
# gestion mdp ---- OK --------
# depot git ---- OK --------
# passer des listes hotel et hotel codeà des dict
# integrer messages console dans tkinter

from playwright.sync_api import sync_playwright
import os
import time
import tkinter as tk
from tkinter import messagebox
import shutil
import mdp
import concat_book_files
import concat_classement_files
import convert_format_csv


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
    "The View Rabat",
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
    433554,
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
    "N001",
]

# user = os.getenv("USER")  ### LINUX
# tmp_folder = "/mnt/wslg/distro/tmp/"  ### LINUX FOLDER
# PATH = (f"/mnt/c/users/{user}/onedrive/documents/Experienciah/Extract ReviewPro/")  ### LINUX FOLDER
user = os.getenv("USERNAME")  ### WINDOWS
PATH = (
    f"c:/users/{user}/onedrive/documents/Experienciah/Extract ReviewPro/"  ### WINDOWS FOLDER
)
tmp_folder = f"c:/users/{user}/AppData/Local/Temp/"  ### WINDOWS FOLDER

option1_path = "Classements/"
option2_path = "Notes Booking/"
mois_str = ""
chemin = PATH
choix = 1  # 1 = Classements, 2 = Notes booking


# Fonction pour valider les entrées
def valider():
    global mois_str, chemin, choix  # Déclaration pour modifier les variables globales
    try:
        mois = int(entry_mois.get())
        if mois < 1 or mois > 12:
            raise ValueError

        # Conversion du mois au format "XX"
        mois_str = f"{mois:02d}"

        # Récupération du choix utilisateur
        choix = choix_var.get()

        root.quit()

    except ValueError:
        messagebox.showerror("Erreur", "Veuillez entrer un mois valide entre 1 et 12.")


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

    # while not file or any("crdown" in item for item in file) or len(file) > 1:     GARDER POUR WINDOWS
    #     file = os.listdir(f"{tmp_folder}{dossier[-1]}")                                   GARDER POUR WINDOWS
    #     print(f"fichier(s) trouvé(s){file}")                                              GARDER POUR WINDOWS

    # time.sleep(2)
    src = f"{tmp_folder}{dossier[0]}/{str(added_element)}"
    dest = f"{PATH}{option_path}{mois_str}-{years[1]}/{hotel}.xlsx"
    shutil.copy(src, dest)

    return file


def download_files(url, option_path, years, hotel):

    # page.wait_for_timeout(1000)
    page.goto(url)
    page.get_by_role("button", name="Actions").click()
    page.get_by_role("menuitem", name="Télécharger rapport").click()
    # page.wait_for_timeout(1000)
    page.get_by_role("button", name="Télécharger le rapport").click()


if __name__ == "__main__":

    # Création de la fenêtre principale
    root = tk.Tk()
    root.title("Paramètres de récupération")
    root.geometry("400x300")

    label_mois = tk.Label(root, text="Mois à extraire (entre 1 et 12) :")
    label_mois.pack(pady=5)
    entry_mois = tk.Entry(root)
    entry_mois.pack(pady=5)

    label_choix = tk.Label(root, text="Choisissez une option :")
    label_choix.pack(pady=5)

    choix_var = tk.IntVar()  # choix 1 ou 2
    choix_var.set(1)  # Valeur par défaut 1

    # Option 1
    radio_classements = tk.Radiobutton(
        root, text="Classements", variable=choix_var, value=1
    )
    radio_classements.pack(pady=5)

    # Option 2
    radio_booking = tk.Radiobutton(
        root, text="Notes booking", variable=choix_var, value=2
    )
    radio_booking.pack(pady=5)

    # Bouton pour valider
    btn_valider = tk.Button(root, text="Valider", command=valider)
    btn_valider.pack(pady=10)

    # Lancement de la boucle principale
    root.mainloop()

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

        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        url = "https://app.reviewpro.com/login"
        page.goto(url)

        page.get_by_role("button", name="Accept All").click()
        page.get_by_label("Email").fill(mdp.id)
        page.get_by_role("button", name="Next").click()
        page.get_by_label("Password").fill(mdp.mdp)
        page.get_by_role("button", name="Log in").click()

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

            download_files(url, option_path, years, hotel)
            print(f"enregistrement du fichier {hotel}")
            file_old = record_files(tmp_folder, file_old, hotel)
            print(f"extraction et enregistrement de {hotel} terminés")
            print("---------------------------------------------------")

        browser.close()

    year = years[1]
    if choix == 2 :
        concat_book_files.concat(PATH, option_path, year, mois_str)
        path = f"{PATH}{option_path}"
        print(path)
        convert_format_csv.convert_csv(path, mois_str, year)
    else:
        concat_classement_files.concat(PATH, option_path, year, mois_str)

