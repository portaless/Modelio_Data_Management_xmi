import tkinter as tk
from tkinter import ttk, filedialog
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import Workbook
import textwrap

def load_xml(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()
    except Exception as e:
        print(f"Erreur lors du chargement du fichier XML : {e}")
        return None
    return xml_content

def parse_xml(xml_content):
    soup = BeautifulSoup(xml_content, 'xml')

    # Récupération des noms des cas d'utilisation (UseCase)
    use_case_names = {elem.get('xmi:id'): elem.get('name') 
                      for elem in soup.find_all(['packagedElement', 'nestedClassifier'], {'xmi:type': 'uml:UseCase'})}

    # Récupérer les exigences (xmi:type="uml:Class", name="Exigence")
    exigence_elements = soup.find_all('packagedElement', {'xmi:type': 'uml:Class', 'name': 'Exigence'})
    
    # Extraire les identifiants des exigences trouvées
    exigence_ids = [elem.get('xmi:id') for elem in exigence_elements]

    # Récupération des exigences associées
    requirement_elements = {elem.get('xmi:id'): elem.get('name') for elem in soup.find_all('packagedElement', {'classifier': exigence_ids})}
    
    # Fusionner les cas d'utilisation et les exigences
    all_elements = {**use_case_names, **requirement_elements}

    # Récupération des descriptions à partir de ownedComment -> body pour les cas d'utilisation et exigences
    descriptions = {}
    for elem_id, name in all_elements.items():
        elem = soup.find('packagedElement', {'xmi:id': elem_id})
        if elem:
            owned_comment = elem.find('ownedComment')
            if owned_comment:
                body = owned_comment.find('body')
                if body:
                    descriptions[elem_id] = body.text  # Enregistre la description
                else:
                    descriptions[elem_id] = ''  # Pas de description trouvée
            else:
                descriptions[elem_id] = ''  # Pas de commentaire trouvé

    # Rechercher les dépendances
    dependency_elements = soup.find_all('packagedElement', {'xmi:type': 'uml:Dependency'})

    dependencies = []
    elements_with_dependencies = set()
    all_element_ids = set(all_elements.keys())
    
    # Traitement des dépendances
    for dep_elem in dependency_elements:
        xmi_id = dep_elem.get('xmi:id')
        supplier_id = dep_elem.get('supplier')
        client_id = dep_elem.get('client')
        
        supplier_valid = supplier_id in all_elements
        client_valid = client_id in all_elements
        
        if not supplier_valid or not client_valid:
            continue
        
        supplier_name = all_elements.get(supplier_id, 'Inconnu')
        client_name = all_elements.get(client_id, 'Inconnu')
        supplier_desc = descriptions.get(supplier_id, '')
        client_desc = descriptions.get(client_id, '')
        
        if xmi_id:
            dependencies.append((xmi_id, supplier_name, client_name, supplier_desc, client_desc))
            elements_with_dependencies.add(supplier_id)
            elements_with_dependencies.add(client_id)
    
    # Élément sans dépendance
    elements_without_dependencies = all_element_ids - elements_with_dependencies

    return dependencies, all_elements, elements_without_dependencies

def save_to_excel(dependencies):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if not file_path:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Matrice de Traçabilité"

    ws.append(["xmi:id", "Supplier", "Client", "Supplier Description", "Client Description"])

    for dep in dependencies:
        ws.append(dep)

    wb.save(file_path)
    print(f"Matrice sauvegardée dans le fichier {file_path}")

# Fonction pour ajuster le texte à une largeur spécifique
def adjust_text_to_width(text, width, font):
    # Créer une instance d'un widget temporaire pour obtenir la largeur des caractères
    temp_label = tk.Label(text=text, font=font)
    # Largeur moyenne des caractères
    char_width = temp_label.winfo_reqwidth() / len(text) if len(text) > 0 else 1
    # Nombre de caractères par ligne
    max_chars_per_line = int(width / char_width)
    # Couper le texte en fonction de la largeur disponible
    wrapped_text = "\n".join(textwrap.wrap(text, width=max_chars_per_line))
    return wrapped_text

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

    # Ajouter une barre de défilement
    tree_frame = ttk.Frame(left_frame)
    tree_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    tree_scrollbar = ttk.Scrollbar(tree_frame)
    tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    tree = ttk.Treeview(tree_frame, columns=("ID", "Supplier", "Client"), show='headings', yscrollcommand=tree_scrollbar.set)
    tree.pack(expand=True, fill='both')

    tree_scrollbar.config(command=tree.yview)

    tree.heading("ID", text="xmi:id")
    tree.heading("Supplier", text="Supplier")
    tree.heading("Client", text="Client")

    # Taille maximale en pixels de la colonne Supplier et Client (à ajuster selon vos besoins)
    column_width = 200  # Exemple de largeur en pixels
    font = ("Arial", 10)  # Police utilisée dans le Treeview

    for dep in dependencies:
        supplier_name = adjust_text_to_width(dep[1], column_width, font)
        client_name = adjust_text_to_width(dep[2], column_width, font)
        tree.insert("", tk.END, values=(dep[0], supplier_name, client_name))

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



def main():
    file_path = 'Modele.xmi'
    xml_content = load_xml(file_path)
    if xml_content:
        dependencies, all_elements, elements_without_dependencies = parse_xml(xml_content)
        create_gui(dependencies, all_elements, elements_without_dependencies)

if __name__ == "__main__":
    main()
