from playwright.sync_api import sync_playwright
import os
import time

#years = ["2018", "2019", "2020", "2021", "2022", "2023", "2024"]
#months_days_start = ["01-01", "02-01", "03-01", "04-01", "05-01", "06-01", "07-01", "08-01", "09-01", "10-01", "11-01",
#                    "12-01"]
#months_days_end = ["01-31", "02-28", "03-31", "04-30", "05-31", "06-30", "07-31", "08-31", "09-30", "10-31", "11-30",
#                  "12-31"]

years = ["2023","2024"]
months_days_start = ["09-01"]
months_days_end = ["09-30"]

# LISTE HOTELS, FAIRE MODIF CODE POUR PRENDRE PLUTOT l'ID PRESENT DANS l'URL (+ AJOUTER ID HOTEL EXPERIENCIAH
hotels = ["The View Rabat", "Amadil Ocean Club Agadir", "Atlas Essaouira Riad Resort", "Atlas Orient",
         "Atlas Volubilis", "Dunes d’Or Ocean Club Agadir", "Jaal Riad Resort Marrakech", "Kaan Casablanca",
         "Marina Bay City Center Tanger", "Palais Medina Riad Resort", "Relax Airport", "Relax Casa Voyageurs",
         "Relax Kenitra", "Relax Marrakech", "Relax Oujda", "Sky Casa Airport Casablanca",
         "Terminus City Center Oujda", "The View Agadir", "The View Bouznika"]

hotels_code = [433554, 107298, 148910, 233862,
         150056, 161034, 139880, 150687,
         150495, 148875, 233869, 430592,
         430598, 430587, 453117, 150665,
         234184, 110037, 433555]

"""
# CREATION REPERTOIRES
for y in range(len(years) - 1):
   for m_d in range(len(months_days_start)):
       if os.path.exists(f"c:/Users/Julien/Onedrive/Documents/Jul6/{months_days_start[m_d]}-{years[y + 1]}"):
           print("le repertoire existe déjà")
       else:
           os.mkdir(f"c:/Users/Julien/Onedrive/Documents/Jul6/{months_days_start[m_d]}-{years[y + 1]}")
"""


# TELECHARGEMENT FICHIERS
for y in range(len(years)):
   print(f"démarrage année {years[y + 1]}")
   for m_d in range(len(months_days_start)):

       with sync_playwright() as p:
           browser = p.chromium.launch(headless=False)
           page = browser.new_page()

           url = f"https://app.reviewpro.com/reviews/results?fd={years[y + 1]}-{months_days_start[m_d]}&td={years[y + 1]}-{months_days_end[m_d]}&prevFd={years[y]}-{months_days_start[m_d]}&prevTd={years[y]}-{months_days_end[m_d]}&fdManagement=2024-05-08&tdManagement=2024-06-07&lang=fr&pid=107298&dataGrouping=daily&indexType=GRI"

           page.goto(url)
           page.get_by_role("button", name="Accept All").click()
           page.get_by_label("Email").fill("rfarmawi@hotelsatlas.com")
           page.get_by_role("button", name="Next").click()
           page.get_by_label("Password").fill("Jadane3121")
           page.get_by_role("button", name="Log in").click()

           print(f"démarrage mois {months_days_start[m_d]}")

           page.get_by_role("button", name=hotels[0]).click()

           for hotel in hotels:
               if os.path.exists(
                       f"c:/Users/Julien/Onedrive/Documents/Jul6/{months_days_start[m_d]}-{years[y + 1]}/{hotel}.xlsx"):
                   print("le fichier existe déjà")
               else:
                   page.get_by_role("button", name=hotel).click()
                   page.get_by_role("button", name="Actions").click()
                   page.get_by_role("menuitem", name="Télécharger rapport").click()
                   page.wait_for_timeout(1000)
                   page.get_by_role("button", name="Télécharger le rapport").click()
                   print(f"le fichier de l'hotel {hotel} a bien été téléchargé")

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
                   dest = f"c:/Users/Julien/Onedrive/Documents/Jul6/{months_days_start[m_d]}-{years[y + 1]}/{hotel}.xlsx"
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