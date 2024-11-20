
from bs4 import BeautifulSoup
import ui #IHM pour la visualisation

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
    
    # Fusionner les dicts des cas d'utilisation et des exigences
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



def main():
    file_path = 'Modele.xmi'
    xml_content = load_xml(file_path)
    if xml_content:
        dependencies, all_elements, elements_without_dependencies = parse_xml(xml_content)
        ui.create_gui(dependencies, all_elements, elements_without_dependencies)

if __name__ == "__main__":
    main()
