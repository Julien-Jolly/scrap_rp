# Scraping et Traitement des Données ReviewPro

Ce projet est une application Python qui automatise la récupération et le traitement des rapports depuis la plateforme ReviewPro. Il utilise Playwright pour la navigation web, Tkinter pour une interface graphique, et des scripts de traitement pour consolider les données dans des fichiers CSV.

## Table des Matières
- [Description](#description)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Utilisation](#utilisation)
- [Structure du Code](#structure-du-code)
- [Dépendances](#dépendances)
- [Configuration](#configuration)
- [Fonctionnalités](#fonctionnalités)
- [Notes Techniques](#notes-techniques)
- [Contribution](#contribution)
- [Licence](#licence)
- [Contact](#contact)

## Description

L'application permet de télécharger des rapports (Classements et Notes Booking) depuis ReviewPro pour une liste prédéfinie d'hôtels, en fonction d'un mois et d'une année spécifiés par l'utilisateur via une interface Tkinter. Les données sont ensuite extraites, formatées avec des virgules comme séparateurs décimaux, et consolidées dans des fichiers CSV organisés par hôtel.

## Prérequis

- Python 3.12 ou version supérieure
- Environnement Windows (le code est optimisé pour Windows, avec des chemins spécifiques)
- Accès à un compte ReviewPro avec identifiants valides

## Installation

1. Clone le dépôt :
   ```bash
   git clone https://github.com/Julien-Jolly/scrap_rp.git
   cd scrap_rp
   
2. Crée un environnement virtuel :
   ```bash
   python -m venv venv
   
3. Active l'environnement virtuel :
   ```bash
   venv\Scripts\activate
      
4. Installe les dépendances :
   ```bash
   pip install -r requirements.txt
   

## Utilisation

1. Lance l'application :
   ```bash
   python main.py
   
2. Une fenêtre Tkinter s'ouvre :
   Entre un mois (1-12).
   Choisis entre "Classements" ou "Notes Booking".
   Clique sur "Valider".

3. L'application se connecte à ReviewPro, télécharge les rapports, et les traite automatiquemen

4. Les logs s'affichent dans l'interface Tkinter, et les fichiers sont enregistrés dans Extract ReviewPro/Fait/.
   

## Structure du Code

scrap_rp/
│
├── main.py              # Point d'entrée avec interface Tkinter et logique de téléchargement
├── concat_book_files.py # Script de concaténation des données Notes Booking
├── concat_classement_files.py # Script de concaténation des données Classements
├── convert_format_csv.py # Conversion des fichiers CSV au format français
├── mdp.py               # Stockage sécurisé des identifiants (non versionné)
├── requirements.txt     # Liste des dépendances
├── .gitignore           # Fichiers à ignorer (ex. __pycache__, venv)
└── README.md            # Documentation

Contact
Auteur : Julien Jolly
Email : julien.jolly@gmail.com
GitHub : github.com/Julien-Jolly