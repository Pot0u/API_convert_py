import os
import json
from typing import List, Dict, Any, Union
from collections import defaultdict
from pathlib import Path

import pdfplumber
import re
from geopy.geocoders import Nominatim
import pandas as pd
from datetime import datetime
import time

# Ajouter après les autres imports globaux
geolocator = Nominatim(user_agent="detect-ville")


class InvoiceParser:
    def __init__(self):
        self.global_delivery_date = None  # Variable pour stocker la date globale
        self.last_result = None  # Variable pour stocker le dernier résultat
        self.pdf_path = None

        # Configuration du géocodeur avec timeout plus long
        self.geolocator = Nominatim(
            user_agent="detect-ville",
            timeout=5  # Augmenter le timeout à 5 secondes
        )

    def _is_valid_date(self, day: str, month: str, year: str) -> bool:
        """Vérifie si la date est valide"""
        try:
            day = int(day)
            month = int(month)
            year = int(year)
            
            if not (1 <= day <= 31):
                return False
            if not (1 <= month <= 12):
                return False
            if not (1900 <= year <= 3000):
                return False
                
            # Vérification supplémentaire pour les mois à 30 jours
            if month in [4, 6, 9, 11] and day > 30:
                return False
                
            # Vérification pour février
            if month == 2:
                if (year % 4 == 0 and year % 100 != 0) or year % 400 == 0:
                    return day <= 29
                return day <= 28
                
            return True
        except ValueError:
            return False

    def _extract_date_from_text(self, text: str, pos: int = None) -> Union[str, None]:
        """Cherche une date valide au format dd.mm.yyyy ou dd/mm/yyyy"""
        # D'abord chercher le format dd.mm.yyyy
        dates = []
        for match in re.finditer(r"(\d{2})\.(\d{2})\.(\d{4})", text):
            if self._is_valid_date(match.group(1), match.group(2), match.group(3)):
                dates.append((match.group(0), match.start()))
        
        # Si aucune date valide trouvée, chercher le format dd/mm/yyyy
        if not dates:
            for match in re.finditer(r"(\d{2})/(\d{2})/(\d{4})", text):
                if self._is_valid_date(match.group(1), match.group(2), match.group(3)):
                    dates.append((match.group(0), match.start()))
        
        if dates:
            if pos is not None:
                # Retourner la date la plus proche de pos
                return min(dates, key=lambda x: abs(x[1] - pos))[0]
            # Sans pos, retourner la première date valide
            return dates[0][0]
        return None

    def _extract_date_from_line(self, line: str) -> Union[str, None]:
        """Recherche une date valide dans la ligne"""
        return self._extract_date_from_text(line)

    def _extract_global_date(self, text: str) -> Union[str, None]:
        """Retourne la date valide la plus proche du mot 'livraison'"""
        # Position du mot livraison (insensible à la casse)
        liv = re.search(r"livraison", text, re.IGNORECASE)
        
        if liv:
            pos = liv.start()
            # Chercher une date valide près de "livraison"
            if date := self._extract_date_from_text(text, pos):
                print(f"Date valide trouvée : {date}")
                return date
        
        # Si aucune date valide n'est trouvée près de "livraison"
        print("Aucune date valide trouvée près du mot 'livraison'")
        return None

    def _extract_total(self, text: str) -> Union[str, None]:
        """Recherche la ligne contenant 'total' et retourne le montant format FR"""
        for line in text.splitlines():
            if re.search(r"total", line, re.IGNORECASE):
                # Pattern plus flexible pour les montants incluant les millions
                patterns = [
                    # Millions avec séparateurs (ex: 1.234.567,89)
                    r"(\d{1,3}(?:\.\d{3}){2,},\d{2})",
                    # Millions sans séparateurs (ex: 1234567,89)
                    r"(\d{7,},\d{2})",
                    # Format standard avec séparateur (ex: 123.456,78)
                    r"(\d{1,3}(?:\.\d{3})*,\d{2})",
                    # Format sans séparateur (ex: 123456,78)
                    r"(\d{1,6},\d{2})",
                ]
                
                for pattern in patterns:
                    m = re.search(pattern, line)
                    if m:
                        total = m.group(1)
                        print(f"Total trouvé avec pattern '{pattern}': {total}")
                        return total
        return None

    def _extract_date_from_context(self, text: str, start_pos: int, window: int = 200) -> Union[str, None]:
        """
        Recherche une date dans une fenêtre de texte autour d'une position donnée
        """
        # Définir la fenêtre de recherche
        start = max(0, start_pos - window)
        end = min(len(text), start_pos + window)
        search_text = text[start:end]
        
        # Chercher toutes les dates dans la fenêtre
        dates = re.finditer(r"\d{1,2}[\.\-/]\d{1,2}[\.\-/]\d{2,4}", search_text)
        dates = list(dates)
        
        # Retourner la date la plus proche de la position de référence
        if dates:
            return min(dates, key=lambda m: abs(m.start() + start - start_pos)).group(0)
        return None

    def _heuristic_parse(self, text: str, next_page_text: str = None) -> List[Dict[str, Any]]:
        """Extraction ligne-à-ligne par regex/keywords si Donut n'a pas tout"""
        results = []
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        current = None
        
        # Pattern pour détecter "page n / n" et "Montant total HT"
        page_pattern = re.compile(r"^page\s+\d+\s*/\s*\d+$", re.IGNORECASE)
        montant_total_pattern = re.compile(r"montant\s+total\s+ht", re.IGNORECASE)
        
        def find_product_name_after_montant_ht(text: str) -> Union[str, None]:
            """Cherche le nom du produit après 'Montant HT' dans le texte donné"""
            if not text:
                return None
                
            lines = text.splitlines()
            for i, line in enumerate(lines):
                if "Montant HT" in line:
                    # Parcourir les lignes suivantes jusqu'à trouver une ligne non vide
                    for next_line in lines[i+1:]:
                        if next_line.strip():
                            # Vérifier que ce n'est pas une ligne d'item
                            if not (next_line.split() + [None, None])[:2][0].isdigit():
                                return next_line.strip()
            return None

        for i, line in enumerate(lines):
            parts = line.split()
            # ligne item: 2 premiers tokens numériques
            if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                # sauvegarde item précédent
                if current:
                    results.append(current)
                    
                current = {
                    "position": parts[0],
                    "designation": parts[1],
                    "nom_produit": None,
                    "quantite": None,
                    "unite": None,
                    "prix_unitaire": None,
                    "date_livraison": None
                }
                
                # Chercher d'abord une date dans la ligne courante
                date_in_line = self._extract_date_from_line(line)
                if date_in_line:
                    current["date_livraison"] = date_in_line
                
                # Si pas de date, chercher près du mot "livraison"
                if not current["date_livraison"]:
                    # Chercher "livraison" dans le contexte
                    context = "\n".join(lines[i:min(i+5, len(lines))])
                    liv_match = re.search(r"livraison", context, re.IGNORECASE)
                    if liv_match:
                        date = self._extract_date_from_context(context, liv_match.start())
                        if date:
                            current["date_livraison"] = date
                
                # Extraire les informations numériques de la ligne
                for j, tok in enumerate(parts[2:], start=2):
                    # Si on a déjà trouvé une quantité et une unité, ne pas les écraser
                    if not current["quantite"] and re.fullmatch(r"\d+[\.,]?\d*", tok):
                        if j + 1 < len(parts) and not re.search(r"\d", parts[j+1]):
                            current["quantite"] = tok
                            current["unite"] = parts[j+1]
                            # Chercher le prix unitaire juste après l'unité
                            if j + 2 < len(parts) and re.fullmatch(r"\d{1,3}(?:[\.]\d{3})*,\d{2}", parts[j+2]):
                                current["prix_unitaire"] = parts[j+2]
                
                # Gestion du nom du produit
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if page_pattern.match(next_line) or montant_total_pattern.search(next_line):
                        print(f"Détection de 'page n / n' ou 'Montant total HT': {next_line}")
                        # Chercher le nom dans la page suivante
                        if next_page_text:
                            product_name = find_product_name_after_montant_ht(next_page_text)
                            if product_name:
                                print(f"Nom de produit trouvé sur la page suivante: {product_name}")
                                current["nom_produit"] = product_name
                    else:
                        # Comportement normal - prendre la ligne suivante si ce n'est pas un nouvel item
                        if not (next_line.split() + [None, None])[:2][0].isdigit():
                            current["nom_produit"] = next_line

        # Ne pas oublier le dernier item
        if current:
            results.append(current)
            
        return results

    def _merge_items(self, a: List[Dict[str, Any]], b: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Fusionne deux listes d'items selon la clé 'position'"""
        merged: Dict[str, Dict[str, Any]] = {}
        # démarrer avec heuristique
        for it in b:
            merged[it['position']] = {**it}
        # fusionner Donut
        for it in a:
            pos = it.get('position')
            if not pos:
                continue
            if pos in merged:
                for k, v in it.items():
                    if v and not merged[pos].get(k):
                        merged[pos][k] = v
            else:
                merged[pos] = it
        # trier par position numérique
        return [merged[k] for k in sorted(merged, key=lambda x: int(x))]

    def parse_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """Traite un PDF page par page en utilisant pdfplumber"""
        all_items: List[Dict[str, Any]] = []
        all_texts: List[str] = []

        with pdfplumber.open(pdf_path) as pdf:
            # Extraire le texte de la première page pour la date globale
            first_page_text = pdf.pages[0].extract_text()
            # Stocker la date globale trouvée près du mot "livraison"
            self.global_delivery_date = self._extract_global_date(first_page_text)
            print(f"Date de livraison globale trouvée: {self.global_delivery_date}")
            
            # Extraire les informations globales
            numero_commande = self._extract_order_number(first_page_text)
            ville = self.find_city(pdf_path)
            
            # Traiter toutes les pages pour les items
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"Traitement de la page {page_num}...")
                
                # Extraire le texte de la page
                text = page.extract_text()
                if text:
                    all_texts.append(text)
                
                try:
                    # Obtenir le texte de la page suivante si elle existe
                    next_page_text = pdf.pages[page_num].extract_text() if page_num < len(pdf.pages) else None
                    
                    # Analyse heuristique du texte
                    page_items = self._heuristic_parse(text, next_page_text)
                    print(f"Heuristique a trouvé {len(page_items)} items sur la page {page_num}")
                    
                    # Ajouter la date de livraison
                    for item in page_items:
                        if not item.get('date_livraison'):
                            item['date_livraison'] = self.global_delivery_date
                    
                    all_items.extend(page_items)
                    print(f"Total: {len(page_items)} items trouvés sur la page {page_num}\n")
                    
                except Exception as e:
                    print(f"Erreur lors du traitement de la page {page_num}: {str(e)}")
                    continue

        # Recherche du total HT dans le texte complet
        total_text = "\n".join(all_texts)
        total_ht = self._extract_total(total_text)

        result = {
            "items": all_items,
            "total_ht": total_ht,
            "numero_commande": numero_commande,
            "ville": ville
        }
        
        # Sauvegarder le résultat pour l'export Excel
        self.last_result = result
        
        return result

    def _extract_order_number(self, text: str) -> Union[str, None]:
        """Recherche du numéro de commande dans le texte."""
        order_patterns = [
            r'Commande\s*N°\s*(\d+/[A-Z]+)',  # Format "4500791137/ROTI"
            r'N°\s*commande\s*:\s*(\d+)',     # Format "N° commande : 4500791137"
            r'Commande\s*:\s*(\d+)',          # Format "Commande : 4500791137"
            r'N°\s*(\d{8,})'                  # Format générique (8+ chiffres)
        ]
        for pattern in order_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)
        return None

    def _extract_ville(self, text: str) -> str:
        """Extrait le nom de ville potentiel d'une ligne"""
        def is_valid_city(word: str) -> Union[str, None]:
            """Vérifie si le mot est une ville"""
            try:
                print(f"Vérification du mot: '{word}'")
                # Liste des mots à exclure (tout en minuscules)
                excluded_words = ["poste de","poste rte", "poste electrique", "rte", "(rte)", 
                                 "poteau", "rouge", "poste" , "je", "electrique", "électrique", 
                                 "droite", "colonnes", "trois", "90kv", "smart", "jacquelin", "jeremy"]
                
                # Vérifier si le premier mot est "de"
                first_word = word.strip().split()[0].lower() if word.strip() else ""
                if first_word == "de":
                    print(f"Mot exclu (commence par le mot 'de'): '{word}'")
                    return None

                # Vérifier si la chaîne ne contient que des nombres et des espaces
                if all(c.isdigit() or c.isspace() for c in word):
                    print(f"Mot exclu (nombres uniquement): '{word}'")
                    return None
                
                # Vérifier si le mot commence par "de" (insensible à la casse)
                if word.lower().startswith('de '):
                    print(f"Mot exclu (commence par 'de'): '{word}'")
                    return None

                # Vérifier si la chaîne ne contient que des nombres et des espaces
                if all(c.isdigit() or c.isspace() for c in word):
                    print(f"Mot exclu (nombres uniquement): '{word}'")
                    return None
                
                # Vérifier si le mot contient "rte" ou "kv" (insensible à la casse)
                if "rte" in word.lower() or "kv" in word.lower():
                    print(f"Mot exclu (contient 'rte' ou 'kv'): '{word}'")
                    return None
                
                # Comparaison en minuscules
                if (len(word) < 3 or word.isdigit() or word.lower() in excluded_words):
                    print(f"Mot exclu: '{word}'")
                    return None
                    
                time.sleep(1)
                
                location = self.geolocator.geocode(word, exactly_one=True)
                if location:
                    print(f"Ville trouvée: '{word}'")
                    return word
                print(f"Pas une ville: '{word}'")
                return None
            except Exception as e:
                print(f"Erreur de géocodage pour '{word}': {str(e)}")
                return None

        # Découper la ligne en mots
        words = text.strip().split()
        if not words:
            return ""
        
        # 1. Essayer des combinaisons de mots depuis la fin
        print("\nRecherche par combinaisons de mots depuis la fin:")
        for length in range(3, 0, -1):  # Essayer d'abord 3 mots, puis 2, puis 1
            for i in range(len(words) - 1, length - 2, -1):  # Parcourir depuis la fin
                # Prendre length mots consécutifs en partant de i vers l'arrière
                word_group = " ".join(words[i - length + 1:i + 1])
                word_group = word_group.strip('.,;:-').strip()
                print(f"Essai groupe de {length} mots depuis la fin: '{word_group}'")
                
                if city := is_valid_city(word_group):
                    print(f"✅ Ville trouvée avec {length} mots: {city}")
                    return city

        return ""

    def extract_left_lines_after_obj(self, page_number: int = 0, n_lines: int = 3) -> List[str]:
        """Extrait les lignes de gauche après 'OBJET'"""
        print(f"Analyse de la page {page_number}")
        with pdfplumber.open(self.pdf_path) as pdf:
            page = pdf.pages[page_number]
            words = page.extract_words(use_text_flow=True)
            
            # Localiser 'OBJET'
            try:
                obj = next(w for w in words if w['text'].strip().lower().startswith('objet'))
                y_obj = obj['top']
                print(f"'OBJET' trouvé à y={y_obj}")
            except StopIteration:
                print("❌ Mot 'OBJET' non trouvé")
                return []

            # Filtrer mots en dessous et utiliser 60% de la page
            left_boundary = page.width * 0.6  # 60% de la largeur de la page
            left_words = [w for w in words if w['top'] > y_obj and w['x1'] < left_boundary]
            print(f"Nombre de mots dans la zone de gauche (60%): {len(left_words)}")

            # Regrouper par ligne
            lines_dict = defaultdict(list)
            for w in left_words:
                key = round(w['top'], 1)
                lines_dict[key].append(w)

            # Trier et assembler les lignes
            sorted_tops = sorted(lines_dict.keys())
            left_lines = []
            for top in sorted_tops[:n_lines]:
                line_words = sorted(lines_dict[top], key=lambda w: w['x0'])
                line = " ".join(w['text'] for w in line_words)
                left_lines.append(line)
                print(f"Ligne détectée: {line}")

            return left_lines

    def is_valid_city(self, city_name: str) -> bool:
        """Vérifie si le nom correspond à une ville valide"""
        try:
            # Vérifier si le mot contient "kV"
            if "kv" in city_name.lower():
                print(f"❌ '{city_name}' contient 'kV' - ignoré.")
                return False

            location = self.geolocator.geocode(city_name, exactly_one=True)
            if not location:
                print(f"❌ '{city_name}' n'est pas une ville valide.")
                return False

            # Vérification du type de lieu
            lieu_class = location.raw.get('class')
            lieu_type = location.raw.get('type')

            # Accepter les villes, zones administratives et cours d'eau et highway
            is_valid = (lieu_class == 'place' and lieu_type in ['city', 'town', 'village']) or \
                      (lieu_class == 'boundary' and lieu_type == 'administrative') or \
                      (lieu_class in ['waterway', "highway"]) 

            if is_valid:
                print(f"✅ '{city_name}' est un lieu valide.")
                print(f"Informations :")
                print(f" - Adresse : {location.address}")
                print(f" - Type : {lieu_class}/{lieu_type}")
                return True
            else:
                print(f"❌ '{city_name}' n'est pas un lieu valide ({lieu_class}/{lieu_type})")
                return False

        except Exception as e:
            print(f"❌ Erreur lors de la vérification de '{city_name}': {e}")
            return False

    def _find_city_after_keywords(self, text: str) -> Union[str, None]:
        """Cherche une ville après des mots-clés spécifiques"""
        patterns = [
            r"POSTE\s+ELECTRIQUE\s+DE\s+([^\n\.,]+)",
            r"POSTE\s+RTE\s+DE\s+([^\n\.,]+)"
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                potential_city = match.group(1).strip()
                print(f"\nRecherche après motif '{match.group(0)}':")
                
                # Essayer le mot directement après
                if city := self._extract_ville(potential_city):
                    return city
                    
                # Si pas de ville trouvée, essayer les deux prochains mots
                words = potential_city.split()
                if len(words) >= 2:
                    two_words = " ".join(words[:2])
                    if city := self._extract_ville(two_words):
                        return city
        
        return None

    def find_city(self, pdf_path: str) -> str:
        """Processus complet d'extraction et validation de ville"""
        self.pdf_path = pdf_path
        
        # 1. D'abord essayer dans les lignes après OBJET
        lines = self.extract_left_lines_after_obj(page_number=0, n_lines=2)
        print("\nRecherche de villes dans les lignes...")
        
        for line in lines:
            print(f"\nAnalyse de la ligne : {line}")
            if city := self._extract_ville(line):
                return city
        
        # 2. Si aucune ville trouvée, chercher après les mots-clés dans tout le texte
        print("\nRecherche après les mots-clés spécifiques...")
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join(page.extract_text() for page in pdf.pages)
            if city := self._find_city_after_keywords(full_text):
                return city
        
        # 3. Si toujours rien trouvé
        print("\nAucune ville trouvée.")
        return None

    def _clean_number(self, value: str) -> str:
        """Nettoie un nombre en supprimant les points et gardant la virgule"""
        if not value:
            return value
        return str(value).replace('.', '')

    def export_to_excel(self, output_path: str = None) -> None:
        try:
            # Créer une copie des données pour ne pas modifier l'original
            items_data = self.last_result["items"].copy()
            
            # Ajouter le numéro de commande et la ville à chaque ligne
            for item in items_data:
                item['numero_commande'] = self.last_result.get('numero_commande', '')
                item['ville'] = self.last_result.get('ville', '')
            
            # Nettoyer les nombres dans les items
            for item in items_data:
                if item.get('prix_unitaire'):
                    item['prix_unitaire'] = self._clean_number(item['prix_unitaire'])
                if item.get('quantite'):
                    item['quantite'] = self._clean_number(item['quantite'])
            
            # Créer le DataFrame pour les items
            df_items = pd.DataFrame(items_data)
            
            # Définir les nouveaux noms de colonnes
            column_names = {
                'numero_commande': 'Numéro de commande',
                'ville': 'Ville',
                'position': 'Position',
                'designation': 'Référence',
                'nom_produit': 'Désignation',
                'quantite': 'Quantité',
                'unite': 'Unité',
                'prix_unitaire': 'Prix unitaire',
                'date_livraison': 'Date de livraison'
            }
            
            # Réorganiser et renommer les colonnes
            columns_order = [
                'numero_commande',
                'ville',
                'position',
                'designation',
                'nom_produit',
                'quantite',
                'unite',
                'prix_unitaire',
                'date_livraison'
            ]
            
            # Réorganiser et renommer
            df_items = df_items[columns_order].rename(columns=column_names)
            
            # Si le chemin n'est pas spécifié, créer un nom par défaut
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"facture_{timestamp}.xlsx"
            
            # Si le fichier existe et est ouvert, créer un nouveau nom
            if os.path.exists(output_path):
                try:
                    # Essayer d'ouvrir le fichier en écriture pour tester s'il est verrouillé
                    with open(output_path, 'a+b'):
                        pass
                except PermissionError:
                    # Si le fichier est verrouillé, créer un nouveau nom avec timestamp
                    base, ext = os.path.splitext(output_path)
                    timestamp = datetime.now().strftime("_%Y-%m-%d_%H-%M-%S")
                    output_path = f"{base}{timestamp}{ext}"
                    print(f"Fichier existant verrouillé, utilisation du nouveau nom: {output_path}")

            # Créer un writer Excel
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
                # Écrire les items dans la première feuille
                df_items.to_excel(writer, sheet_name='Items', index=False)
                
                # Écrire les informations globales dans la deuxième feuille
                df_global = pd.DataFrame({
                    'Numéro de commande': [self.last_result.get('numero_commande', '')],
                    'Ville': [self.last_result.get('ville', '')],
                    'Total HT': [self._clean_number(self.last_result.get('total_ht', ''))]
                })
                df_global.to_excel(writer, sheet_name='Informations globales', index=False)
                
                # Ajuster automatiquement la largeur des colonnes
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
            print(f"Fichier Excel créé avec succès : {output_path}")
            
        except Exception as e:
            print(f"Erreur lors de la création du fichier Excel : {str(e)}")
            raise

if __name__ == "__main__":
    print("=== Convertisseur PDF vers Excel ===")
    
    # Demander le chemin du fichier PDF
    while True:
        pdf_path = input("\nVeuillez entrer le chemin complet du fichier PDF (ou 'q' pour quitter) : ")
        
        if pdf_path.lower() == 'q':
            break
            
        if not os.path.exists(pdf_path):
            print(f"Erreur: Le fichier '{pdf_path}' n'existe pas.")
            continue
            
        try:
            parser = InvoiceParser()
            print(f"\nTraitement du fichier : {pdf_path}")
            print("Veuillez patienter...")
            
            res = parser.parse_pdf(pdf_path)
            excel_path = pdf_path.replace('.pdf', '.xlsx')
            parser.export_to_excel(excel_path)
            
            print("\nTraitement terminé !")
            print(f"Fichier Excel créé : {excel_path}")
            
        except Exception as e:
            print(f"\nErreur lors du traitement : {str(e)}")
            
        print("\n" + "="*50)


