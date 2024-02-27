import os
import sys
import tkinter as tk
from tkinter import messagebox, simpledialog
import pyperclip
from unidecode import unidecode
import keyboard

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

keyboard.on_press(on_press)
keyboard.on_release(on_release)

# Créer la fenêtre principale
fenetre = tk.Tk()
fenetre.title("SAP")
fenetre.geometry("600x300")

# Créer la barre de menu
barre_menu = tk.Menu(fenetre)
fenetre.config(menu=barre_menu)

# Créer le menu "Import"
menu_import = tk.Menu(barre_menu, tearoff=0)
barre_menu.add_cascade(label="Import", menu=menu_import)

# Ajouter les options de sous-menu
menu_import.add_command(label="Importer ME2N:F01 depuis presse papier", command=lambda: importer_du_pressepapier("F01.txt"))
menu_import.add_command(label="Importer ME2N:F02 depuis presse papier", command=lambda: importer_du_pressepapier("F02.txt"))
menu_import.add_command(label="Importer ME2N:121 depuis presse papier", command=lambda: importer_du_pressepapier("121.txt"))
menu_import.add_command(label="Importer MB52 depuis presse papier", command=lambda: importer_du_pressepapier("MB52.txt"))

# Créer le menu "Afficher"
menu_afficher = tk.Menu(barre_menu, tearoff=0)
barre_menu.add_cascade(label="Afficher", menu=menu_afficher)

# Ajouter les options de sous-menu
menu_afficher.add_command(label="Afficher ME2N:F01", command=lambda: afficher_contenu_fichier("F01.txt"))
menu_afficher.add_command(label="Afficher ME2N:F02", command=lambda: afficher_contenu_fichier("F02.txt"))
menu_afficher.add_command(label="Afficher ME2N:121", command=lambda: afficher_contenu_fichier("121.txt"))
menu_afficher.add_command(label="Afficher MB52", command=lambda: afficher_contenu_fichier("MB52.txt"))

# Créer le menu "Taille"
menu_taille_police = tk.Menu(barre_menu, tearoff=0)
barre_menu.add_cascade(label="Taille de police", menu=menu_taille_police)

# Ajouter les options de sous-menu pour différentes tailles de police
menu_taille_police.add_command(label="6", command=lambda: change_font_size(6))
#menu_taille_police.add_command(label="7", command=lambda: change_font_size(7))
menu_taille_police.add_command(label="8", command=lambda: change_font_size(8))
#menu_taille_police.add_command(label="9", command=lambda: change_font_size(9))
menu_taille_police.add_command(label="10", command=lambda: change_font_size(10))
menu_taille_police.add_command(label="11", command=lambda: change_font_size(11))
menu_taille_police.add_command(label="12", command=lambda: change_font_size(12))





# Créer le menu "MB52"
menu_mb52 = tk.Menu(barre_menu, tearoff=0)
barre_menu.add_cascade(label="MB52", menu=menu_mb52)
menu_mb52.add_command(label="Voir liste complete", command=lambda: filter_0())
menu_mb52.add_command(label="Voir filtre 'Utilis.lib' = 0", command=lambda: filter_1())
menu_mb52.add_command(label="Voir filtre 'Emplacement' [mini,maxi]", command=lambda: filter_2())
menu_mb52.add_command(label="Voir filtre a commander", command=lambda: filter_3())
menu_mb52.add_command(label="Filtrer par mot(s) clé", command=lambda: filter_4())

# Ajouter une nouvelle option de menu pour copier le contenu
barre_menu.add_command(label="Copier dans le presse-papiers", command=lambda: copier_dans_pressepapier())


# Créer une zone de texte pour afficher le contenu
zone_texte = tk.Text(fenetre)
zone_texte.grid(row=1, column=0, sticky="nsew")

# Créer une barre de défilement verticale
scrollbar = tk.Scrollbar(fenetre, command=zone_texte.yview)
scrollbar.grid(row=1, column=1, sticky="ns")
zone_texte.config(yscrollcommand=scrollbar.set)

# Configurer le dimensionnement de la grille
fenetre.grid_rowconfigure(1, weight=1)  # Permet à la ligne 1 de la grille de prendre tout l'espace vertical disponible
fenetre.grid_columnconfigure(0, weight=1)  # Permet à la colonne 0 de la grille de prendre tout l'espace horizontal disponible


# Lancer la boucle principale
fenetre.mainloop()
