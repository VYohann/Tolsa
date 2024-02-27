import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, simpledialog
import webbrowser
import os
import sys
import pyperclip
from unidecode import unidecode
import keyboard
import requests

def ouvrir_page():
    url = entry.get()
    if url == "sap":
       #url = "https://raw.githubusercontent.com/VYohann/Tolsa/main/maintenance.py"
        url="https://raw.githubusercontent.com/VYohann/Tolsa/main/sap.py"
        #url = input("Url du fichier python a executer: ")

        # Télécharger le contenu du fichier depuis l'URL
        response = requests.get(url)

        # Vérifier si le téléchargement s'est bien passé (code 200)
        if response.status_code == 200:
            # Execute le contenu du fichier
            exec(response.text)
        else:
            print(f"Échec du téléchargement. Code de statut : {response.status_code}")
    else:
        webbrowser.open_new(url)

def ouvrir_ecran_process():
    lien = "http://10.80.152.101"
    webbrowser.open_new(lien)

def ouvrir_ecran_marfil():
    lien = "http://192.168.4.52"
    webbrowser.open_new(lien)

def ouvrir_etiqueteuse_markem_imaje():
    lien = "http://192.168.4.221"
    webbrowser.open_new(lien)

def ouvrir_etiqueteuse_zebra_500():
    lien = "http://192.168.4.160"
    webbrowser.open_new(lien)

def ouvrir_etiqueteuse_zebra_bureau():
    lien = "http://10.80.128.54"
    webbrowser.open_new(lien)

# Créer la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Maintenance Utilitaire")

# Créer un champ de saisie pour l'URL
entry = ttk.Entry(fenetre, width=25)
entry.insert(0, "https://www.google.com")
entry.grid(row=0, column=0, padx=10, pady=10)

# Créer un bouton pour ouvrir la page
bouton_ouvrir_page = ttk.Button(fenetre, text="Ouvrir la page web", command=ouvrir_page)
bouton_ouvrir_page.grid(row=0, column=1, padx=10, pady=10)

# Créer un bouton pour ouvrir le lien "10.80.152.101"
bouton_ouvrir_lien = ttk.Button(fenetre, text="Ecran process", command=ouvrir_ecran_process)
bouton_ouvrir_lien.grid(row=1, column=0, padx=10, pady=10)

# Créer un bouton pour ouvrir le lien "192.168.4.52"
bouton_ouvrir_lien = ttk.Button(fenetre, text="Ecran marfil", command=ouvrir_ecran_marfil)
bouton_ouvrir_lien.grid(row=1, column=1, padx=10, pady=10)

# Créer un bouton pour ouvrir le lien "192.168.4.221"
bouton_ouvrir_lien = ttk.Button(fenetre, text="Etiqueteuse markem_imaje", command=ouvrir_etiqueteuse_markem_imaje)
bouton_ouvrir_lien.grid(row=2, column=0, padx=10, pady=10)

# Créer un bouton pour ouvrir le lien "192.168.4.160"
bouton_ouvrir_lien = ttk.Button(fenetre, text="Etiqueteuse Zebra 500", command=ouvrir_etiqueteuse_zebra_500)
bouton_ouvrir_lien.grid(row=2, column=1, padx=10, pady=10)

# Créer un bouton pour ouvrir le lien "10.80.128.54"
bouton_ouvrir_lien = ttk.Button(fenetre, text="Etiqueteuse Zebra bureau", command=ouvrir_etiqueteuse_zebra_bureau)
bouton_ouvrir_lien.grid(row=2, column=2, padx=10, pady=10)

# Démarrer la boucle principale
fenetre.mainloop()
