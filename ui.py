import tkinter as tk
from tkinter import ttk, filedialog
from openpyxl import Workbook
import pandas as pd

def save_to_excel(dependencies):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        return

    wb = Workbook()
    
    # Première feuille : données de dépendance
    ws1 = wb.active
    ws1.title = "Données de Dépendance"
    
    # En-têtes
    ws1.append(["xmi:id", "Supplier", "Client", "Supplier Description", "Client Description"])
    
    for dep in dependencies:
        ws1.append(dep)

    # Préparation des données pour la matrice de couverture
    suppliers = set()
    clients = set()
    
    for _, supplier, client, _, _ in dependencies:
        if supplier:
            suppliers.add(supplier)
        if client:
            clients.add(client)
    
    suppliers = sorted(suppliers)
    clients = sorted(clients)
    
    # Création de la matrice de couverture
    matrix_data = pd.DataFrame(index=suppliers, columns=clients).fillna('')

    for _, supplier, client, _, _ in dependencies:
        if supplier and client:
            matrix_data.at[supplier, client] = 'X'  # 'X' indique un impact
    
    # Deuxième feuille : matrice de couverture
    ws2 = wb.create_sheet(title="Matrice de Couverture")

    # Ajout des en-têtes de colonne
    ws2.append([''] + clients)
    
    # Ajouter les lignes de la matrice
    for supplier in suppliers:
        row = [supplier] + [matrix_data.at[supplier, client] for client in clients]
        ws2.append(row)
    
    wb.save(file_path)
    print(f"Matrice sauvegardée dans le fichier {file_path}")

def create_gui(dependencies, all_elements, elements_without_dependencies):
    root = tk.Tk()
    root.title("Matrice de Traçabilité")

    # Configure la fenêtre principale pour utiliser tout l'espace
    root.geometry("800x600")  # Taille initiale, peut être ajustée
    root.minsize(600, 400)    # Taille minimale
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    main_frame = ttk.Frame(root)
    main_frame.grid(row=0, column=0, sticky="nsew")

    main_frame.rowconfigure(0, weight=1)
    main_frame.columnconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=1)

    left_frame = ttk.Frame(main_frame, padding="10", relief="solid")
    left_frame.grid(row=0, column=0, sticky="nsew")

    right_frame = ttk.Frame(main_frame, padding="10", relief="solid")
    right_frame.grid(row=0, column=1, sticky="nsew")

    left_frame.rowconfigure(0, weight=0)  # Label header
    left_frame.rowconfigure(1, weight=1)  # Treeview
    left_frame.rowconfigure(2, weight=0)  # Label
    left_frame.rowconfigure(3, weight=1)  # Listbox

    left_frame.columnconfigure(0, weight=1)

    right_frame.rowconfigure(0, weight=1)  # Canvas
    right_frame.rowconfigure(1, weight=0)  # Label
    right_frame.rowconfigure(2, weight=0)  # Button

    right_frame.columnconfigure(0, weight=1)

    ttk.Label(left_frame, text="Matrice de Traçabilité", font=("Arial", 14)).grid(row=0, column=0, pady=10, sticky="ew")

    tree = ttk.Treeview(left_frame, columns=("ID", "Supplier", "Client"), show='headings')
    tree.heading("ID", text="xmi:id")
    tree.heading("Supplier", text="Supplier")
    tree.heading("Client", text="Client")
    tree.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    for dep in dependencies:
        tree.insert("", tk.END, values=dep[:3])

    ttk.Label(left_frame, text="Cas d'utilisation ou exigences sans dépendances", font=("Arial", 12)).grid(row=2, column=0, pady=10, sticky="ew")

    listbox = tk.Listbox(left_frame)
    listbox.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")

    for element_id in elements_without_dependencies:
        name = all_elements.get(element_id, 'Inconnu')
        listbox.insert(tk.END, f"{element_id}: {name}")

    # Canvas pour le cercle
    canvas = tk.Canvas(right_frame, bg="white")
    canvas.grid(row=0, column=0, pady=20, sticky="nsew")

    right_frame.rowconfigure(0, weight=1)
    right_frame.columnconfigure(0, weight=1)

    total_elements = len(all_elements)
    covered_elements = total_elements - len(elements_without_dependencies)
    coverage_percentage = (covered_elements / total_elements) * 100 if total_elements > 0 else 0

    # Fonction pour redessiner un cercle plus esthétique
    def draw_circle(event):
        canvas.delete("all")  # Supprimer le contenu précédent

        width = event.width
        height = event.height
        diameter = min(width, height) - 40  # Garde un peu plus de marge
        radius = diameter // 2

        center_x = width // 2
        center_y = height // 2

        # Ombre du cercle pour l'effet 3D
        canvas.create_oval(center_x - radius + 5, center_y - radius + 5,
                           center_x + radius + 5, center_y + radius + 5,
                           outline="", fill="#d9d9d9")  # Ombre grise

        # Cercle principal (arrière-plan)
        canvas.create_oval(center_x - radius, center_y - radius,
                           center_x + radius, center_y + radius,
                           outline="", fill="#f0f0f0")  # Couleur douce de fond

        # Arc représentant la couverture
        extent_angle = (covered_elements / total_elements) * 360 if total_elements > 0 else 0
        canvas.create_arc(center_x - radius, center_y - radius,
                          center_x + radius, center_y + radius,
                          start=90, extent=-extent_angle,
                          fill="#3cb371", outline="")  # Arc vert doux

        # Cercle central pour masquer la partie intérieure (donne un effet "donut")
        inner_radius = radius * 0.75
        canvas.create_oval(center_x - inner_radius, center_y - inner_radius,
                           center_x + inner_radius, center_y + inner_radius,
                           outline="", fill="white")  # Centre blanc

    # Lier le redimensionnement du Canvas à la fonction draw_circle
    canvas.bind("<Configure>", draw_circle)

    ttk.Label(right_frame, text=f"Taux de couverture: {coverage_percentage:.2f}%", font=("Arial", 12)).grid(row=1, column=0, pady=10, sticky="ew")

    save_button = ttk.Button(right_frame, text="Exporter en Excel", command=lambda: save_to_excel(dependencies))
    save_button.grid(row=2, column=0, pady=20, sticky="ew")

    root.mainloop()