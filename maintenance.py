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

#def_sap
def copier_dans_pressepapier():
    contenu = zone_texte.get("1.0", "end-1c")  # Récupérer le contenu de la zone de texte
    pyperclip.copy(contenu)  # Copier le contenu dans le presse-papiers

def change_font_size(size):
    #zone_texte.configure(font=("Courier New", size))
    #zone_texte.configure(font=("Monaco", size))
    zone_texte.configure(font=("Consolas", size))

def get_file_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

def sauvegarder_contenu(contenu, nom_fichier):
    dossier_sauvegarde = get_file_path("SAP_import")
    if not os.path.exists(dossier_sauvegarde):
        os.makedirs(dossier_sauvegarde)

    chemin_fichier = os.path.join(dossier_sauvegarde, nom_fichier)
    
    with open(chemin_fichier, "w") as fichier:
        fichier.write(contenu)
        messagebox.showinfo("Sauvegarde réussie", "Le contenu a été sauvegardé dans le fichier.")

def importer_du_pressepapier(nom_fichier):
    contenu_pressepapier = pyperclip.paste()
    if contenu_pressepapier:
        sauvegarder_contenu(contenu_pressepapier, nom_fichier)
        afficher(contenu_pressepapier)
    else:
        messagebox.showwarning("Aucun contenu", "Il n'y a pas de contenu à sauvegarder.")

def afficher_contenu_fichier(nom_fichier):
    dossier_sap_import = get_file_path("SAP_import")
    chemin_fichier = os.path.join(dossier_sap_import, nom_fichier)

    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r") as fichier:
            contenu = fichier.read()
            if contenu:
                zone_texte.delete("1.0", "end")
                zone_texte.insert("1.0", contenu)
            else:
                messagebox.showinfo("Contenu vide", "Le fichier " + nom_fichier + " est vide.")
    else:
        messagebox.showwarning("Fichier introuvable", "Le fichier " + nom_fichier + " n'a pas été trouvé.")

def afficher(text):
    zone_texte.delete("1.0", "end")
    zone_texte.insert("1.0", text)

def convert_to_data_array(nom_fichier):
    dossier_sap_import = get_file_path("SAP_import")
    chemin_fichier = os.path.join(dossier_sap_import, nom_fichier)

    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r") as fichier:
            contenu = fichier.read()
            if contenu:
                pass
            else:
                messagebox.showinfo("Contenu vide", "Le fichier " + nom_fichier + " est vide.")
    else:
        messagebox.showwarning("Fichier introuvable", "Le fichier " + nom_fichier + " n'a pas été trouvé.")

    input_string = contenu
    # Séparer les lignes du texte
    lines = input_string.strip().split('\n')

    # Trouver l'indice de la ligne d'en-tête
    header_index = None
    for i, line in enumerate(lines):
        if 'Article' in line and 'Désignation article' in line:
            header_index = i
            break

    # Vérifier si la ligne d'en-tête a été trouvée
    if header_index is None:
        return []
    
    # Récupérer les en-têtes de colonne
    headers = [header.strip() for header in lines[header_index].split('|') if header.strip()]
    #afficher(headers)

    # Initialiser la liste de données
    data = [headers]
    #afficher(data)

    # Parcourir les lignes de données
    for i in range(header_index + 2, len(lines)):
        line = lines[i]
        if '|' in line:
            # Séparer les cellules de la ligne de données
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            if cells:
               if 'dlt-' not in cells[1] and cells[1] != '':  # Vérifier si 'dlt-' n'est pas présent dans la colonne 'Désignation article'
                    data.append(cells)

    return data

def data_array_to_string(data_array):
    # Vérifier si la liste de données est vide
    if not data_array or len(data_array) <= 1:
        return ''

    # Récupérer les en-têtes de colonne
    headers = data_array[0]

    # Calculer la largeur de chaque colonne
    column_widths = [max(len(header), max(len(row[i]) for row in data_array[1:])) for i, header in enumerate(headers)]

    # Générer la ligne de séparation
    separator = '-' * (sum(column_widths) + 3 * len(column_widths) + 1)

    # Générer la chaîne de caractères
    output_string = ''
    output_string += separator + '\n'
    output_string += '| ' + ' | '.join(headers[i].ljust(column_widths[i]) for i in range(len(headers))) + ' |\n'
    output_string += separator + '\n'

    for row in data_array[1:]:
        output_string += '| ' + ' | '.join(row[i].ljust(column_widths[i]) for i in range(len(row))) + ' |\n'

    output_string += separator + '\n'

    return output_string

def filter_data_array(data_array, filter_criteria):
    # Vérifier si la liste de données est vide
    if not data_array or len(data_array) <= 1:
        return []

    # Récupérer les en-têtes de colonne
    headers = data_array[0]

    # Récupérer l'indice de la colonne utilisée pour le filtrage
    filter_column_index = None
    for i, header in enumerate(headers):
        if header == filter_criteria:
            filter_column_index = i
            break

    # Vérifier si la colonne de filtrage a été trouvée
    if filter_column_index is None:
        return []

    # Filtrer les lignes de données qui satisfont le critère
    filtered_data_array = [headers]
    for row in data_array[1:]:
        if row[filter_column_index] == '0':
            filtered_data_array.append(row)

    return filtered_data_array

def filter_data_array_by_emplacement(data_array):
    # Vérifier si la liste de données est vide
    if not data_array or len(data_array) <= 1:
        return []

    # Récupérer les en-têtes de colonne
    headers = data_array[0]

    # Récupérer l'indice de la colonne "Emplacemt"
    emplacement_column_index = headers.index('Emplacemt')

    # Créer un nouveau data array pour stocker les lignes filtrées
    filtered_data_array = [headers]

    # Parcourir les lignes de données
    for row in data_array[1:]:
        # Vérifier si le caractère '[' est présent dans la colonne "Emplacemt"
        if '[' in row[emplacement_column_index]:
            filtered_data_array.append(row)

    return filtered_data_array

def filter_data_array_by_emplacement_cmd(data_array):
    # Vérifier si la liste de données est vide
    if not data_array or len(data_array) <= 1:
        return []

    # Récupérer les en-têtes de colonne
    headers = data_array[0]

    # Récupérer les index des colonnes "Emplacemt" et "Utilis.lib"
    emplacement_column_index = headers.index('Emplacemt')
    utilis_lib_column_index = headers.index('Utilis.lib')

    # Créer un nouveau data array pour stocker les lignes filtrées
    filtered_data_array = [headers + ['A commander']]

    # Parcourir les lignes de données
    for row in data_array[1:]:
        emplacement_value = row[emplacement_column_index]
        utilis_lib_value = row[utilis_lib_column_index]

        # Extraire les valeurs entre '[' et ']'
        emplacement_values = [val.strip() for val in emplacement_value.strip('[]').split(',') if val.strip()]

        if len(emplacement_values) >= 2 and emplacement_values[0].isdigit() and emplacement_values[1].isdigit():
            mini, maxi = float(emplacement_values[0]), float(emplacement_values[1])
            stock = float(utilis_lib_value.replace(',', '.'))

            # Vérifier si "A commander" est nécessaire
            a_commander = ''
            if stock < mini:
                a_commander = str(int(maxi - stock))

            # Ajouter la colonne "A commander" et les valeurs filtrées au nouveau data array
            if a_commander != '':  # Afficher uniquement si "A commander" est supérieur à 0
                filtered_data_array.append(row + [a_commander])


    return filtered_data_array


def filter_data_array_by_keywords(data_array, keywords):
    # Vérifier si la liste de données est vide
    if not data_array or len(data_array) <= 1:
        return []

    # Récupérer les en-têtes de colonne
    headers = data_array[0]

    # Diviser les mots-clés en une liste
    keyword_list = keywords.lower().split()

    # Créer un nouveau data array pour stocker les lignes filtrées
    filtered_data_array = [headers]

    # Parcourir les lignes de données
    for row in data_array[1:]:
        keywords_found = [False] * len(keyword_list)

        # Parcourir les cellules de la ligne
        for cell in row:
            # Normaliser la cellule et les mots-clés en caractères non accentués et en minuscules
            normalized_cell = unidecode(cell.lower())
            normalized_keywords = [unidecode(keyword) for keyword in keyword_list]

            # Vérifier si chaque mot-clé est présent dans la cellule normalisée
            for i, keyword in enumerate(normalized_keywords):
                if keyword in normalized_cell:
                    keywords_found[i] = True

        # Si tous les mots-clés sont présents dans la ligne, ajouter la ligne filtrée
        if all(keywords_found):
            filtered_data_array.append(row)

    return filtered_data_array



def filter_data_array_by_keywords_old(data_array, keywords):
    # Vérifier si la liste de données est vide
    if not data_array or len(data_array) <= 1:
        return []

    # Récupérer les en-têtes de colonne
    headers = data_array[0]

    # Diviser les mots-clés en une liste
    keyword_list = keywords.lower().split()

    # Créer un nouveau data array pour stocker les lignes filtrées
    filtered_data_array = [headers]

    # Parcourir les lignes de données
    for row in data_array[1:]:
        keywords_found = False

        # Parcourir les cellules de la ligne
        for cell in row:
            # Normaliser la cellule et les mots-clés en caractères non accentués et en minuscules
            normalized_cell = unidecode(cell.lower())
            normalized_keywords = [unidecode(keyword) for keyword in keyword_list]

            # Vérifier si au moins un mot-clé est présent dans la cellule normalisée
            if any(keyword in normalized_cell for keyword in normalized_keywords):
                keywords_found = True
                break

        if keywords_found:
            filtered_data_array.append(row)

    return filtered_data_array


def filter_0():
    data = convert_to_data_array("MB52.txt")
    afficher(data_array_to_string(data))

def filter_1():
    data = convert_to_data_array("MB52.txt")
    filter_criteria = 'Utilis.lib'
    filtered_data = filter_data_array(data, filter_criteria)
    afficher(data_array_to_string(filtered_data))

def filter_2():
    data = convert_to_data_array("MB52.txt")
    filtered_data = filter_data_array_by_emplacement(data)
    afficher(data_array_to_string(filtered_data))
    
def filter_3():
    data = convert_to_data_array("MB52.txt")
    filtered_data = filter_data_array_by_emplacement_cmd(data)
    afficher(data_array_to_string(filtered_data))

def filter_4():
    data = convert_to_data_array("MB52.txt")
    keyword = tk.simpledialog.askstring("Mots-clés", "Mots-clés séparés par un espace:")
    if keyword == "key":
        afficher(key)
    elif keyword == "clear":
        afficher("")
    elif keyword is not None:  # Vérifier si l'utilisateur a saisi quelque chose
        filtered_data = filter_data_array_by_keywords(data, keyword)
        afficher(data_array_to_string(filtered_data))

key = ''
ctrl_pressed = False

def on_press(event):
    global key, ctrl_pressed
    cle = event.name
    key = key + ' ' + cle

def on_release(event):
    global ctrl_pressed
    cle = event.name
#/def_sap



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
