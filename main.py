import tkinter as tk
from tkinter import ttk
import json
import os

FICHIER = "donnees.json"

# ----------------- Fonctions -----------------

def ajouter_ligne():
    tree.insert("", "end", values=("Janvier", 0, 0, 0, 0, 0, 0, 0, "0 %", 0))
    recalculer_tout()
    sauvegarder_donnees()

def supprimer_ligne():
    selected = tree.selection()
    for item in selected:
        tree.delete(item)
    recalculer_tout()
    sauvegarder_donnees()

def convertir_nombre(valeur):
    valeur = valeur.replace(",", ".")
    try:
        return float(valeur)
    except:
        return None

def recalculer_total_ligne(item):
    values = tree.item(item)["values"]
    try:
        total = sum(float(values[i]) for i in range(1, 7))
    except:
        total = 0
    tree.set(item, "Total", round(total, 2))

def recalculer_pourcentage():
    items = tree.get_children()
    for i in range(len(items)):
        if i == 0:
            tree.set(items[i], "PctGainPerte", "0 %")
            tree.set(items[i], "GainPerte", 0)
        else:
            total_actuel = float(tree.set(items[i], "Total"))
            total_precedent = float(tree.set(items[i-1], "Total"))

            if total_precedent == 0:
                pourcentage = 0
            else:
                pourcentage = ((total_actuel - total_precedent) / total_precedent) * 100

            gain_perte = total_actuel - total_precedent

            tree.set(items[i], "PctGainPerte", f"{round(pourcentage, 2)} %")
            tree.set(items[i], "GainPerte", round(gain_perte, 2))

def recalculer_tout():
    for item in tree.get_children():
        recalculer_total_ligne(item)
    recalculer_pourcentage()
    appliquer_tags()
    mettre_a_jour_dernier_total()

def modifier_cellule(event):
    selected_item = tree.identify_row(event.y)
    col = tree.identify_column(event.x)
    if not selected_item or not col:
        return
    col_index = int(col.replace("#", "")) - 1
    if col_index >= 7:
        return

    x, y, width, height = tree.bbox(selected_item, col)
    entry = tk.Entry(tree, font=("Arial", 11), bg="#2e2e2e", fg="white", insertbackground="white")
    entry.place(x=x, y=y, width=width, height=height)

    valeur_actuelle = tree.item(selected_item)["values"][col_index]
    entry.insert(0, valeur_actuelle)
    entry.focus()

    def sauvegarder(event=None):
        nouvelle_valeur = entry.get()
        if col_index == 0:
            tree.set(selected_item, col, nouvelle_valeur)
        elif 1 <= col_index <= 6:
            nombre = convertir_nombre(nouvelle_valeur)
            if nombre is None:
                entry.destroy()
                return
            tree.set(selected_item, col, round(nombre, 2))
            recalculer_tout()
        entry.destroy()
        sauvegarder_donnees()

    entry.bind("<Return>", sauvegarder)
    entry.bind("<FocusOut>", lambda e: entry.destroy())

# ----------------- Dernier total -----------------

def mettre_a_jour_dernier_total():
    dernier_total = 0
    dernier_mois = ""
    for item in reversed(tree.get_children()):
        total = float(tree.set(item, "Total"))
        if total > 0:
            dernier_total = total
            dernier_mois = tree.set(item, "Mois")
            break
    label_dernier_total.config(text=f"Dernier total ({dernier_mois}): {dernier_total} €")

# ----------------- Sauvegarde / Chargement -----------------

def sauvegarder_donnees():
    data = []
    for item in tree.get_children():
        data.append(tree.item(item)["values"])
    with open(FICHIER, "w") as f:
        json.dump(data, f)

def charger_donnees():
    if not os.path.exists(FICHIER):
        return
    with open(FICHIER, "r") as f:
        data = json.load(f)
        for row in data:
            tree.insert("", "end", values=row)
    recalculer_tout()

# ----------------- Fenêtre -----------------

fenetre = tk.Tk()
fenetre.title("Patrimonial")
fenetre.geometry("1300x650")
fenetre.configure(bg="#1e1e1e")  # Fond foncé

# --- Label dernier total ---
label_dernier_total = tk.Label(fenetre, text="Dernier total: 0 €", font=("Arial", 14, "bold"), bg="#1e1e1e", fg="white")
label_dernier_total.pack(pady=10)

# --- Frame boutons discret ---
frame_boutons = tk.Frame(fenetre, bg="#1e1e1e")
frame_boutons.pack(pady=5)



# --- Boutons discrets légèrement plus gros ---
def style_bouton_discret(btn):
    btn.configure(bg="#1e1e1e", fg="#aaaaaa", font=("Arial", 13, "bold"), bd=0, padx=10, pady=5,
                  activebackground="#1e1e1e", activeforeground="white")
    btn.bind("<Enter>", lambda e: btn.config(fg="white"))
    btn.bind("<Leave>", lambda e: btn.config(fg="#aaaaaa"))

btn_ajouter = tk.Button(frame_boutons, text="+", command=ajouter_ligne)
style_bouton_discret(btn_ajouter)
btn_ajouter.pack(side="left", padx=5)

btn_supprimer = tk.Button(frame_boutons, text="-", command=supprimer_ligne)
style_bouton_discret(btn_supprimer)
btn_supprimer.pack(side="left", padx=5)

# --- Treeview ---
colonnes = (
    "Mois",
    "Binance",
    "IBKR",
    "Tiktok",
    "Compte Bancaire",
    "Liquide",
    "Vinted",
    "Total",
    "PctGainPerte",
    "GainPerte"
)

tree = ttk.Treeview(fenetre, columns=colonnes, show="headings", height=20)

# Style sombre
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview",
                background="#2e2e2e",
                foreground="white",
                fieldbackground="#2e2e2e",
                rowheight=28,
                font=("Arial", 11))
style.configure("Treeview.Heading",
                font=("Arial", 12, "bold"),
                background="#3a3a3a",
                foreground="white")
style.map("Treeview", background=[("selected", "#4a90e2")], foreground=[("selected", "white")])

# Lignes alternées
tree.tag_configure('oddrow', background="#2e2e2e")
tree.tag_configure('evenrow', background="#252525")

for col in colonnes:
    if col == "PctGainPerte":
        tree.heading(col, text="% Gain/Perte")
    elif col == "GainPerte":
        tree.heading(col, text="€ Gain/Perte")
    else:
        tree.heading(col, text=col)
    tree.column(col, width=130, anchor="center")
tree.column("GainPerte", width=150)
tree.pack(expand=True, fill="both", padx=20, pady=10)

tree.bind("<Double-1>", modifier_cellule)

# Lignes alternées
def appliquer_tags():
    for i, item in enumerate(tree.get_children()):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        tree.item(item, tags=(tag,))

charger_donnees()
appliquer_tags()

fenetre.mainloop()