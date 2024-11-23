# redondance dans les if.else
# refaire commentaires
# reduire timeout
# format csv ","
# chemin d'écriture different selon if.else
# 30 et 31 jours ---- A TESTER --------
# reformatage avec __init__
# ajout concat autres fichiers
# voir comment intégrer concat
# gestion mdp ---- OK --------
# depot git



from playwright.sync_api import sync_playwright
import os
import time
import tkinter as tk
from tkinter import messagebox, filedialog
import stat
import subprocess
import shutil
import mdp



# LISTE HOTELS, FAIRE MODIF CODE POUR PRENDRE PLUTOT l'ID PRESENT DANS l'URL (+ AJOUTER ID HOTEL EXPERIENCIAH
hotels = ["Amadil Ocean Club Agadir", "Atlas Essaouira Riad Resort", "Atlas Orient",
         "Atlas Volubilis", "Dunes d’Or Ocean Club Agadir", "Jaal Riad Resort Marrakech", "Kaan Casablanca",
         "Marina Bay City Center Tanger", "Palais Medina Riad Resort", "Relax Airport", "Relax Casa Voyageurs",
         "Relax Kenitra", "Relax Marrakech", "Relax Oujda", "Sky Casa Airport Casablanca",
         "Terminus City Center Oujda", "The View Agadir", "The View Bouznika","The View Rabat"]

hotels_code = [107298, 148910, 233862,
         150056, 161034, 139880, 150687,
         150495, 148875, 233869, 430592,
         430598, 430587, 453117, 150665,
         234184, 110037, 433555, 433554]

#user = os.getenv("USER")
user="julien"
tmp_folder="/mnt/wslg/distro/tmp/"    ### LINUX FOLDER
PATH = f"/mnt/c/users/{user}/onedrive/documents/Atlas/Extract ReviewPro/"     ### LINUX FOLDER
#PATH = f"c:/users/{user}/onedrive/documents/"   ### WINDOWS FOLDER
option1_path = "Classements/"
option2_path = "Notes Booking/"
#tmp_folder  = os.listdir(f"c:/users/{user}/AppData/Local/Temp/") ### WINDOWS FOLDER

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
radio_classements = tk.Radiobutton(root, text="Classements", variable=choix_var, value=1)
radio_classements.pack(pady=5)

# Option 2
radio_booking = tk.Radiobutton(root, text="Notes booking", variable=choix_var, value=2)
radio_booking.pack(pady=5)

# Bouton pour valider
btn_valider = tk.Button(root, text="Valider", command=valider)
btn_valider.pack(pady=10)

# Lancement de la boucle principale
root.mainloop()

print(f"Mois à extraire : {mois_str}")
print(f"Chemin de récupération : {chemin}")
print(f"Choix sélectionné : {choix}")

years = ["2023","2024"]
month_day_start = f"{mois_str}-01"

match mois_str:
    case "02":
        day_end="28"
    case "04":
        day_end="30"
    case "06":
        day_end = "30"
    case "09":
        day_end = "30"
    case "11":
        day_end="30"
    case _:
        day_end="31"

month_day_end = f"{mois_str}-{day_end}"






with sync_playwright() as p:

    browser = p.chromium.launch(headless=False)
    page = browser.new_page()


    url="https://app.reviewpro.com/login"
    page.goto(url)

    try:
        page.get_by_role("button", name="Accept All").click()
    except:
        print("ok")

    page.get_by_label("Email").fill(mdp.id)
    page.get_by_role("button", name="Next").click()
    page.get_by_label("Password").fill(mdp.mdp)
    page.get_by_role("button", name="Log in").click()


    def record_files(tmp_folder, file_old, hotel):
        temp_folder_list = os.listdir(tmp_folder)

        time.sleep(10)
        dossier = [dossier for dossier in temp_folder_list if "artifacts" in dossier]
        print(f"dossier Artifacts : {dossier[0]}")

        file = os.listdir(f"{tmp_folder}{dossier[0]}")


        print(f"précédente liste des ficheirs présents dans le repertoire Artifacts : {file_old}")
        print(f"fichiers présents dans le repertoire Artifacts{file}")


        added_element = set(file) - set(file_old)
        added_element = next(iter(added_element))

        print(f"élément téléchargé added_element {added_element}")



        # while not file or any("crdown" in item for item in file) or len(file) > 1:     GARDER POUR WINDOWS
        #     file = os.listdir(f"{tmp_folder}{dossier[-1]}")                                   GARDER POUR WINDOWS
        #     print(f"fichier(s) trouvé(s){file}")                                              GARDER POUR WINDOWS


        time.sleep(2)
        src = f"{tmp_folder}{dossier[0]}/{str(added_element)}"
        dest = f"{PATH}{option_path}{mois_str}-{years[1]}/{hotel}.xlsx"
        shutil.copy(src, dest)

        return file


    def download_files(url, option_path, years, hotel):

        # page.get_by_role("button", name=hotel).click()
        page.wait_for_timeout(1000)
        #url = f"https://app.reviewpro.com/reviews/dashboard?fd={years[1]}-{month_day_start}&td=2024-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[h]}"
        page.goto(url)
        # page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
        # page.get_by_label("close").click()
        page.get_by_role("button", name="Actions").click()
        #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
        #             page.get_by_label("close").click()
        page.get_by_role("menuitem", name="Télécharger rapport").click()
        #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
        #             page.get_by_label("close").click()
        page.wait_for_timeout(1000)
        #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
        #             page.get_by_label("close").click()
        page.get_by_role("button", name="Télécharger le rapport").click()


    file_old=()

    for i in range(len(hotels)):
        hotel=hotels[i]
        hotel_code = hotels_code[i]



        if choix == 1:
            url = f"https://app.reviewpro.com/reviews/dashboard?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotel_code}"
            option_path = option1_path
        else:
            url = f"https://app.reviewpro.com/reviews/results?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotel_code}&dataGrouping=daily&indexType=GRI"
            option_path = option2_path

        # CREATION REPERTOIRES

        if not os.path.exists(f"{PATH}{option_path}{mois_str}-{years[1]}"):
            os.mkdir(f"{PATH}{option_path}{mois_str}-{years[1]}")
            print(f"Création du répertoire : {PATH}{option_path}{mois_str}-{years[1]}")

        if os.path.exists(f"{PATH}{option_path}{mois_str}-{years[1]}/{hotel}.xlsx"):
            print("le fichier existe déjà")

        download_files(url, option_path, years, hotel)
        file_old = record_files(tmp_folder, file_old, hotel)





    # if choix==1:
    #     h=0
    #     for hotel in hotels:
    #         if os.path.exists(
    #                 f"{PATH}{option1_path}{mois_str}-{years[1]}/{hotels[h]}.xlsx"):
    #             print("le fichier existe déjà")
    #         else:
    #             #page.get_by_role("button", name=hotel).click()
    #             page.wait_for_timeout(1000)
    #             url = f"https://app.reviewpro.com/reviews/dashboard?fd={years[1]}-{month_day_start}&td=2024-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[h]}"
    #             page.goto(url)
    #             # page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    #             # page.get_by_label("close").click()
    #             page.get_by_role("button", name="Actions").click()
    # #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    # #             page.get_by_label("close").click()
    #             page.get_by_role("menuitem", name="Télécharger rapport").click()
    # #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    # #             page.get_by_label("close").click()
    #             page.wait_for_timeout(1000)
    # #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    # #             page.get_by_label("close").click()
    #             page.get_by_role("button", name="Télécharger le rapport").click()
    # #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    # #             page.get_by_label("close").click()
    #             print(f"le fichier de l'hotel {hotels[h]} a bien été téléchargé")
    # #             page.get_by_label("close").click()
    #
    #
    #             temp_folder_list = os.listdir(tmp_folder)
    #             print(temp_folder_list)
    #
    #             time.sleep(10)
    #             dossier = [dossier for dossier in temp_folder_list if "artifacts" in dossier]
    #             print(dossier[0])
    #
    #
    #             file=os.listdir(f"{tmp_folder}{dossier[0]}")
    #
    #             if h==0:
    #                 file_histo=()
    #             added_element = set(file) - set(file_histo)
    #             added_element = next(iter(added_element))
    #
    #             print(file)
    #             print(added_element)
    #             file_histo=file
    #
    #             while not file or any("crdown" in item for item in file) :    #or len(file) > 1:
    #                 file = os.listdir(f"{tmp_folder}{dossier[-1]}")
    #                 print(f"fichier(s) trouvé(s){file}")
    #
    #             print(f"fichier retenu{file}")
    #             time.sleep(2)
    #             src = f"{tmp_folder}{dossier[0]}/{str(added_element)}"
    #             dest = f"{PATH}{option1_path}{mois_str}-{years[1]}/{hotels[h]}.xlsx"
    #             shutil.copy(src, dest)
    #
    #             h+=1
    #
    #
    # else:
    #     h=0
    #     for hotel in hotels:
    #         if os.path.exists(
    #                 f"{PATH}{option2_path}{month_day_start}-{years[1]}/{hotels[h]}.xlsx"):
    #             print("le fichier existe déjà")
    #         else:
    #             page.wait_for_timeout(2000)
    #             url = f"https://app.reviewpro.com/reviews/results?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[h]}&dataGrouping=daily&indexType=GRI"
    #             page.goto(url)
    #             page.get_by_role("button", name="Actions").click()
    #             page.get_by_role("menuitem", name="Télécharger rapport").click()
    #             page.wait_for_timeout(1000)
    #             page.get_by_role("button", name="Télécharger le rapport").click()
    #             print(f"le fichier de l'hotel {hotels[h]} a bien été téléchargé")
    #
    #             temp_folder_list = os.listdir(f"/mnt/c/users/{user}/AppData/Local/Temp/")
    #             dossier = [dossier for dossier in temp_folder_list if "artifacts" in dossier]
    #             print(dossier)
    #
    #             page.get_by_role("button").first.click()
    #             page.get_by_role("button", name=hotel).click()
    #
    #             file = os.listdir(f"/mnt/c/users/{user}/AppData/Local/Temp/{dossier[0]}")
    #             while not file or any("crdown" in item for item in file) or len(file) > 1:
    #                 file = os.listdir(f"/mnt/c/users/{user}/AppData/Local/Temp/{dossier[-1]}")
    #                 print(f"fichier(s) trouvé(s){file}")
    #
    #             print(f"fichier retenu{file}")
    #             time.sleep(2)
    #             src = f"/mnt/c/users/{user}/AppData/Local/Temp/{dossier[0]}/{file[0]}"
    #             dest = f"{PATH}{option2_path}{month_day_start}-{years[1]}/{hotels[h]}.xlsx"
    #             os.rename(src, dest)
    #
    #             for filename in os.listdir(f"/mnt/c/users/{user}/AppData/Local/Temp/{dossier[0]}/"):
    #                 file_path = os.path.join(f"/mnt/c/users/{user}/AppData/Local/Temp/{dossier[0]}/", filename)
    #                 try:
    #                     if os.path.isfile(file_path) or os.path.islink(file_path):
    #                         os.unlink(file_path)
    #                         print(f"Fichier supprimé : {file_path}")
    #                     elif os.path.isdir(file_path):
    #                         os.rmdir(file_path)
    #                         print(f"Répertoire vide supprimé : {file_path}")
    #                 except Exception as e:
    #                     print(f"Impossible de supprimer {file_path}. Raison : {e}")
    #
    #             print("Tous les fichiers ont été supprimés du répertoire.")
    #             h+=1

    browser.close()
