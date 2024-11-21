# redondance dans les if.else
# refaire commentaires
# reduire timeout
# format csv ","
# chemin d'écriture different selon if.else
# 30 et 31 jours
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



# Déclaration des variables globales
mois_str = ""
chemin = ""
choix = 1  # Par défaut : 1 = Classements, 2 = Notes booking

# Fonction pour valider les entrées
def valider():
    global mois_str, chemin, choix  # Déclaration pour modifier les variables globales
    try:
        # Récupération de la valeur du mois
        mois = int(entry_mois.get())
        if mois < 1 or mois > 12:
            raise ValueError

        # Conversion du mois au format "XX"
        mois_str = f"{mois:02d}"

        # Récupération du chemin de récupération
        chemin = entry_chemin.get()

        # Récupération du choix utilisateur
        choix = choix_var.get()

        # Affichage des résultats
        messagebox.showinfo("Informations", f"Mois à extraire : {mois_str}\n"
                                            f"Chemin de récupération : {chemin}\n"
                                            f"Choix sélectionné : {choix}")

        # Ferme la fenêtre après validation
        root.quit()

    except ValueError:
        messagebox.showerror("Erreur", "Veuillez entrer un mois valide entre 1 et 12.")

# Fonction pour parcourir un répertoire
def parcourir_chemin():
    chemin = filedialog.askdirectory(initialdir="C:/Users/Julien/Onedrive/Documents/Jul6/")
    if chemin:
        entry_chemin.delete(0, tk.END)  # Efface l'ancien chemin
        entry_chemin.insert(0, chemin)  # Insère le nouveau chemin sélectionné

# Création de la fenêtre principale
root = tk.Tk()
root.title("Paramètres de récupération")
root.geometry("400x300")

# Label et champ pour le mois à extraire
label_mois = tk.Label(root, text="Mois à extraire (entre 1 et 12) :")
label_mois.pack(pady=5)
entry_mois = tk.Entry(root)
entry_mois.pack(pady=5)

# Label et champ pour le chemin de récupération
label_chemin = tk.Label(root, text="Chemin de récupération :")
label_chemin.pack(pady=5)

# Valeur par défaut du chemin
default_chemin = "C:/Users/Julien/Onedrive/Documents/Jul6/"
entry_chemin = tk.Entry(root, width=50)
entry_chemin.insert(0, default_chemin)
entry_chemin.pack(pady=5)

# Bouton pour parcourir le répertoire
btn_parcourir = tk.Button(root, text="Parcourir...", command=parcourir_chemin)
btn_parcourir.pack(pady=5)

# Ajout des options avec Radiobuttons pour le choix utilisateur
label_choix = tk.Label(root, text="Choisissez une option :")
label_choix.pack(pady=5)

choix_var = tk.IntVar()  # Variable pour stocker le choix de l'utilisateur (1 ou 2)
choix_var.set(1)  # Valeur par défaut

# Option 1: Classements (valeur 1)
radio_classements = tk.Radiobutton(root, text="Classements", variable=choix_var, value=1)
radio_classements.pack(pady=5)

# Option 2: Notes booking (valeur 2)
radio_booking = tk.Radiobutton(root, text="Notes booking", variable=choix_var, value=2)
radio_booking.pack(pady=5)

# Bouton pour valider les informations
btn_valider = tk.Button(root, text="Valider", command=valider)
btn_valider.pack(pady=10)

# Lancement de la boucle principale
root.mainloop()

# Utilisation des variables globales après la fermeture de la fenêtre
print(f"Mois à extraire : {mois_str}")
print(f"Chemin de récupération : {chemin}")
print(f"Choix sélectionné : {choix}")

years = ["2023","2024"]
month_day_start = f"{mois_str}-01"
month_day_end = f"{mois_str}-31"


# CREATION REPERTOIRES

for y in range(len(years) - 1):
    for m_d in range(len(month_day_start)):
        if os.path.exists(f"{chemin}{month_day_start}-{years[1]}"):
           print("le repertoire existe déjà")
        else:
           os.mkdir(f"{chemin}{month_day_start}-{years[1]}")


# TELECHARGEMENT FICHIERS

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    if choix==1:
        url = f"https://app.reviewpro.com/reviews/dashboard?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[0]}"
    else:
        url = f"https://app.reviewpro.com/reviews/results?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[0]}&dataGrouping=daily&indexType=GRI"

    page.goto(url)
    page.get_by_role("button", name="Accept All").click()
    page.get_by_label("Email").fill(mdp.id)
    page.get_by_role("button", name="Next").click()
    page.get_by_label("Password").fill(mdp.mdp)
    page.get_by_role("button", name="Log in").click()


    if choix==1:
        h=0
        for hotel in hotels:
            if os.path.exists(
                    f"{chemin}{month_day_start}-{years[1]}/{hotels[h]}.xlsx"):
                print("le fichier existe déjà")
            else:
                #page.get_by_role("button", name=hotel).click()
                page.wait_for_timeout(2000)
                url = f"https://app.reviewpro.com/reviews/dashboard?fd={years[1]}-{month_day_start}&td=2024-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[h]}"
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
    #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    #             page.get_by_label("close").click()
                print(f"le fichier de l'hotel {hotels[h]} a bien été téléchargé")
    #             page.get_by_label("close").click()

                user = os.getlogin()
                liste_dossiers = os.listdir("/Users/Julien/AppData/Local/Temp/")
                dossier = [dossier for dossier in liste_dossiers if "artifacts" in dossier]
                print(dossier)

                # page.get_by_role("button").first.click()
                # page.get_by_role("button", name=hotel).click()

                file = os.listdir(f"/Users/Julien/AppData/Local/Temp/{dossier[0]}")
                while not file or any("crdown" in item for item in file) or len(file) > 1:
                    file = os.listdir(f"/Users/Julien/AppData/Local/Temp/{dossier[-1]}")
                    print(f"fichier(s) trouvé(s){file}")

                print(f"fichier retenu{file}")
                time.sleep(2)
                src = f"C:/Users/Julien/AppData/Local/Temp/{dossier[0]}/{file[0]}"
                dest = f"c:/Users/Julien/Onedrive/Documents/Jul6/{month_day_start}-{years[1]}/{hotels[h]}.xlsx"
                os.rename(src, dest)

                for filename in os.listdir(f"C:/Users/Julien/AppData/Local/Temp/{dossier[0]}/"):
                    file_path = os.path.join(f"C:/Users/Julien/AppData/Local/Temp/{dossier[0]}/", filename)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path):
                            os.unlink(file_path)
                            print(f"Fichier supprimé : {file_path}")
                        elif os.path.isdir(file_path):
                            os.rmdir(file_path)
                            print(f"Répertoire vide supprimé : {file_path}")
                    except Exception as e:
                        print(f"Impossible de supprimer {file_path}. Raison : {e}")

                print("Tous les fichiers ont été supprimés du répertoire.")
                h+=1
                # page.wait_for_timeout(5000)
    #             page.locator("div").filter(has_text="YOUR REVIEWPRO ACCOUNT HAS").nth(3).click()
    #             page.get_by_label("close").click()

    else:
        h=0
        for hotel in hotels:
            if os.path.exists(
                    f"c:/Users/Julien/Onedrive/Documents/Jul6/{month_day_start}-{years[1]}/{hotels[h]}.xlsx"):
                print("le fichier existe déjà")
            else:
                page.wait_for_timeout(2000)
                url = f"https://app.reviewpro.com/reviews/results?fd={years[1]}-{month_day_start}&td={years[1]}-{month_day_end}&prevFd={years[0]}-{month_day_start}&prevTd={years[0]}-{month_day_end}&fdManagement={years[1]}-{month_day_start}&tdManagement={years[1]}-{month_day_end}&lang=fr&pid={hotels_code[h]}&dataGrouping=daily&indexType=GRI"
                page.goto(url)
                page.get_by_role("button", name="Actions").click()
                page.get_by_role("menuitem", name="Télécharger rapport").click()
                page.wait_for_timeout(1000)
                page.get_by_role("button", name="Télécharger le rapport").click()
                print(f"le fichier de l'hotel {hotels[h]} a bien été téléchargé")

                user = os.getlogin()
                liste_dossiers = os.listdir("/Users/Julien/AppData/Local/Temp/")
                dossier = [dossier for dossier in liste_dossiers if "artifacts" in dossier]
                print(dossier)

                page.get_by_role("button").first.click()
                page.get_by_role("button", name=hotel).click()

                file = os.listdir(f"/Users/Julien/AppData/Local/Temp/{dossier[0]}")
                while not file or any("crdown" in item for item in file) or len(file) > 1:
                    file = os.listdir(f"/Users/Julien/AppData/Local/Temp/{dossier[-1]}")
                    print(f"fichier(s) trouvé(s){file}")

                print(f"fichier retenu{file}")
                time.sleep(2)
                src = f"C:/Users/Julien/AppData/Local/Temp/{dossier[0]}/{file[0]}"
                dest = f"c:/Users/Julien/Onedrive/Documents/Jul6/{month_day_start}-{years[1]}/{hotels[h]}.xlsx"
                os.rename(src, dest)

                for filename in os.listdir(f"C:/Users/Julien/AppData/Local/Temp/{dossier[0]}/"):
                    file_path = os.path.join(f"C:/Users/Julien/AppData/Local/Temp/{dossier[0]}/", filename)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path):
                            os.unlink(file_path)
                            print(f"Fichier supprimé : {file_path}")
                        elif os.path.isdir(file_path):
                            os.rmdir(file_path)
                            print(f"Répertoire vide supprimé : {file_path}")
                    except Exception as e:
                        print(f"Impossible de supprimer {file_path}. Raison : {e}")

                print("Tous les fichiers ont été supprimés du répertoire.")
                h+=1

