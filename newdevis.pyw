import tkinter as tk
from tkinter import messagebox, Toplevel
import xlwt
import os
from xlrd import open_workbook
from xlutils.copy import copy
from tkinter import ttk

# Fonction pour enregistrer les données dans un fichier Excel
def save_data():
    try:
        devis_number = entry_devis_number.get()
        client_name = entry_client_name.get()
        project_name = entry_project_name.get()
        date = entry_date.get()
        sale_price = float(entry_sale_price.get())
        purchase_budget = float(entry_purchase_budget.get())
        labor_cost = float(entry_labor_cost.get())
        general_expenses = float(entry_general_expenses.get())  # Traité comme un nombre

        # Vérifier si le fichier Excel existe déjà
        if not os.path.exists('form_data.xls'):
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('Data')
            sheet.write(0, 0, 'Numéro de devis')
            sheet.write(0, 1, 'Nom du client')
            sheet.write(0, 2, 'Nom du projet')
            sheet.write(0, 3, 'Date')
            sheet.write(0, 4, 'Prix de vente')
            sheet.write(0, 5, 'Budget achat')
            sheet.write(0, 6, 'Main d\'œuvre prévue')
            sheet.write(0, 7, 'Frais généraux')
            row = 1
        else:
            rb = open_workbook('form_data.xls')
            workbook = copy(rb)
            sheet = workbook.get_sheet(0)
            row = sheet.nrows  # Ajoute les données à la première ligne vide

        # Écrire les données dans la nouvelle ligne
        sheet.write(row, 0, devis_number)
        sheet.write(row, 1, client_name)
        sheet.write(row, 2, project_name)
        sheet.write(row, 3, date)
        sheet.write(row, 4, sale_price)
        sheet.write(row, 5, purchase_budget)
        sheet.write(row, 6, labor_cost)
        sheet.write(row, 7, general_expenses)  # Stocker en tant que nombre

        workbook.save('form_data.xls')
        
        # Réinitialiser les champs du formulaire
        entry_devis_number.delete(0, tk.END)
        entry_client_name.delete(0, tk.END)
        entry_project_name.delete(0, tk.END)
        entry_date.delete(0, tk.END)
        entry_sale_price.delete(0, tk.END)
        entry_purchase_budget.delete(0, tk.END)
        entry_labor_cost.delete(0, tk.END)
        entry_general_expenses.delete(0, tk.END)

        # Afficher un message de confirmation
        messagebox.showinfo("Succès", "Données enregistrées avec succès dans form_data.xls")
    
    except ValueError:
        messagebox.showerror("Erreur", "Veuillez entrer des valeurs numériques valides pour le prix de vente, le budget achat, la main d'œuvre prévue, et les frais généraux.")

# Fonction pour quitter l'application
def quit_app():
    root.destroy()

# Fonction pour afficher les devis
def show_devis():
    if not os.path.exists('form_data.xls'):
        messagebox.showinfo("Info", "Aucune donnée disponible.")
        return
    
    rb = open_workbook('form_data.xls')
    sheet = rb.sheet_by_index(0)

    # Créer une nouvelle fenêtre pour afficher les devis
    top = Toplevel(root)
    top.title("Devis")

    # Ajuster les dimensions de la nouvelle fenêtre
    top.geometry("1000x400")

    # Créer un Treeview (tableau) pour afficher les données
    tree = ttk.Treeview(top)
    
    # Définir les colonnes
    tree["columns"] = ("devis_number", "client_name", "project_name", "date", "sale_price", "purchase_budget", "labor_cost", "general_expenses")
    
    # Configurer les en-têtes de colonnes
    tree.heading("devis_number", text="Numéro de devis")
    tree.heading("client_name", text="Nom du client")
    tree.heading("project_name", text="Nom du projet")
    tree.heading("date", text="Date")
    tree.heading("sale_price", text="Prix de vente")
    tree.heading("purchase_budget", text="Budget achat")
    tree.heading("labor_cost", text="Main d'œuvre prévue")
    tree.heading("general_expenses", text="Frais généraux")

    tree.column("#0", width=0, stretch=tk.NO)  # Hide the first column
    tree.column("devis_number", anchor=tk.W, width=120)
    tree.column("client_name", anchor=tk.W, width=150)
    tree.column("project_name", anchor=tk.W, width=150)
    tree.column("date", anchor=tk.W, width=100)
    tree.column("sale_price", anchor=tk.E, width=100)
    tree.column("purchase_budget", anchor=tk.E, width=100)
    tree.column("labor_cost", anchor=tk.E, width=100)
    tree.column("general_expenses", anchor=tk.E, width=100)

    tree.pack(fill=tk.BOTH, expand=1)

    # Insérer les données du fichier Excel dans le tableau
    for row_idx in range(1, sheet.nrows):  # Skip the header row
        row_values = sheet.row_values(row_idx)
        tree.insert("", tk.END, values=row_values)

# Initialiser la fenêtre tkinter
root = tk.Tk()
root.title("Formulaire de Saisie")

# Agrandir la taille de la fenêtre de 25%
root.geometry("600x500")

# Définir une taille de police plus grande pour les labels et les champs
label_font = ("Arial", 14)
entry_font = ("Arial", 14)

# Créer les labels et champs de saisie
tk.Label(root, text="Numéro de devis", font=label_font).grid(row=0, sticky=tk.W, padx=10, pady=5)
tk.Label(root, text="Nom du client", font=label_font).grid(row=1, sticky=tk.W, padx=10, pady=5)
tk.Label(root, text="Nom du projet", font=label_font).grid(row=2, sticky=tk.W, padx=10, pady=5)
tk.Label(root, text="Date", font=label_font).grid(row=3, sticky=tk.W, padx=10, pady=5)
tk.Label(root, text="Prix de vente", font=label_font).grid(row=4, sticky=tk.W, padx=10, pady=5)
tk.Label(root, text="Budget achat", font=label_font).grid(row=5, sticky=tk.W, padx=10, pady=5)
tk.Label(root, text="Main d'œuvre prévue", font=label_font).grid(row=6, sticky=tk.W, padx=10, pady=5)

# Le texte "Frais généraux" en bleu
tk.Label(root, text="Frais généraux", font=label_font, fg="blue").grid(row=7, sticky=tk.W, padx=10, pady=5)

entry_devis_number = tk.Entry(root, font=entry_font)
entry_client_name = tk.Entry(root, font=entry_font)
entry_project_name = tk.Entry(root, font=entry_font)
entry_date = tk.Entry(root, font=entry_font)
entry_sale_price = tk.Entry(root, font=entry_font)
entry_purchase_budget = tk.Entry(root, font=entry_font)
entry_labor_cost = tk.Entry(root, font=entry_font)
entry_general_expenses = tk.Entry(root, font=entry_font)

entry_devis_number.grid(row=0, column=1, padx=10, pady=5)
entry_client_name.grid(row=1, column=1, padx=10, pady=5)
entry_project_name.grid(row=2, column=1, padx=10, pady=5)
entry_date.grid(row=3, column=1, padx=10, pady=5)
entry_sale_price.grid(row=4, column=1, padx=10, pady=5)
entry_purchase_budget.grid(row=5, column=1, padx=10, pady=5)
entry_labor_cost.grid(row=6, column=1, padx=10, pady=5)
entry_general_expenses.grid(row=7, column=1, padx=10, pady=5)

# Créer le bouton pour afficher les devis
tk.Button(root, text='Afficher', command=show_devis, font=label_font).grid(row=8, column=0, pady=10)

# Créer le bouton pour soumettre les données
tk.Button(root, text='Soumettre', command=save_data, font=label_font).grid(row=8, column=1, pady=10)

# Créer le bouton pour quitter l'application
tk.Button(root, text='Quitter', command=quit_app, font=label_font).grid(row=8, column=2, pady=10)

# Ajouter une étiquette en bas avec la version du programme
tk.Label(root, text="Version 1.0 - YR", font=("Arial", 10), fg="grey").grid(row=9, column=1, pady=10)

# Démarrer la boucle principale de l'interface graphique
root.mainloop()
