"""
facture_to_excel.py

Module principal pour l'extraction d'informations structurées à partir de factures PDF
et l'export de ces données vers un fichier Excel.

Fonctionnalités principales :
- Extraction des lignes d'items, des dates, des montants et de la ville depuis un PDF.
- Validation des villes via Nominatim (OpenStreetMap).
- Export des résultats dans un fichier Excel structuré.

Auteur  : Lam Clément
Date    : 2024-06
Licence : Usage interne VINCI Energies

Dépendances :
- pdfplumber
- geopy
- pandas
- openpyxl
"""

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

geolocator = Nominatim(user_agent="detect-ville")

class InvoiceParser:
    """
    Permet d'extraire les informations structurées d'une facture PDF (items, ville, date, etc.)
    et de les exporter vers Excel.

    Attributs :
        global_delivery_date (str|None) : Date de livraison globale trouvée dans le PDF.
        last_result (dict|None) : Dernier résultat d'extraction.
        pdf_path (str|None) : Chemin du PDF en cours de traitement.
        geolocator (Nominatim) : Instance du géocodeur pour la validation des villes.
    """

    def __init__(self):
        """
        Initialise le parser et le géocodeur.
        """
        self.global_delivery_date = None
        self.last_result = None
        self.pdf_path = None
        self.geolocator = Nominatim(user_agent="detect-ville", timeout=5)

    def _is_valid_date(self, day: str, month: str, year: str) -> bool:
        """
        Vérifie si une date (jour, mois, année) est valide.

        Args:
            day (str): Jour.
            month (str): Mois.
            year (str): Année.

        Returns:
            bool: True si la date est valide, False sinon.
        """
        try:
            day = int(day)
            month = int(month)
            year = int(year)
            if not (1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 3000):
                return False
            # Gère les mois à 30 jours et les années bissextiles pour février
            if month in [4, 6, 9, 11] and day > 30:
                return False
            if month == 2:
                if (year % 4 == 0 and year % 100 != 0) or year % 400 == 0:
                    return day <= 29
                return day <= 28
            return True
        except ValueError:
            return False

    def _extract_date_from_text(self, text: str, pos: int = None) -> Union[str, None]:
        """
        Cherche une date valide au format dd.mm.yyyy ou dd/mm/yyyy dans le texte.
        Si pos est fourni, retourne la date la plus proche de cette position.

        Args:
            text (str): Texte à analyser.
            pos (int, optional): Position de référence.

        Returns:
            str|None: Date trouvée ou None.
        """
        dates = []
        for match in re.finditer(r"(\d{2})\.(\d{2})\.(\d{4})", text):
            if self._is_valid_date(match.group(1), match.group(2), match.group(3)):
                dates.append((match.group(0), match.start()))
        if not dates:
            for match in re.finditer(r"(\d{2})/(\d{2})/(\d{4})", text):
                if self._is_valid_date(match.group(1), match.group(2), match.group(3)):
                    dates.append((match.group(0), match.start()))
        if dates:
            if pos is not None:
                # Prend la date la plus proche de la position donnée (utile pour "livraison")
                return min(dates, key=lambda x: abs(x[1] - pos))[0]
            return dates[0][0]
        return None

    def _extract_date_from_line(self, line: str) -> Union[str, None]:
        """
        Recherche une date valide dans une ligne de texte.

        Args:
            line (str): Ligne à analyser.

        Returns:
            str|None: Date trouvée ou None.
        """
        return self._extract_date_from_text(line)

    def _extract_global_date(self, text: str) -> Union[str, None]:
        """
        Retourne la date valide la plus proche du mot 'livraison' dans le texte.

        Args:
            text (str): Texte à analyser.

        Returns:
            str|None: Date trouvée ou None.
        """
        liv = re.search(r"livraison", text, re.IGNORECASE)
        if liv:
            pos = liv.start()
            if date := self._extract_date_from_text(text, pos):
                print(f"Date valide trouvée : {date}")
                return date
        print("Aucune date valide trouvée près du mot 'livraison'")
        return None

    def _extract_total(self, text: str) -> Union[str, None]:
        """
        Recherche la ligne contenant 'total' et retourne le montant au format français.

        Args:
            text (str): Texte à analyser.

        Returns:
            str|None: Montant trouvé ou None.
        """
        for line in text.splitlines():
            if re.search(r"total", line, re.IGNORECASE):
                # Plusieurs patterns pour gérer les différents formats de montants
                patterns = [
                    r"(\d{1,3}(?:\.\d{3}){2,},\d{2})",
                    r"(\d{7,},\d{2})",
                    r"(\d{1,3}(?:\.\d{3})*,\d{2})",
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
        Recherche une date dans une fenêtre de texte autour d'une position donnée.
        Utile pour trouver une date proche d'un mot-clé.

        Args:
            text (str): Texte à analyser.
            start_pos (int): Position centrale.
            window (int): Taille de la fenêtre autour de la position.

        Returns:
            str|None: Date trouvée ou None.
        """
        start = max(0, start_pos - window)
        end = min(len(text), start_pos + window)
        search_text = text[start:end]
        dates = list(re.finditer(r"\d{1,2}[\.\-/]\d{1,2}[\.\-/]\d{2,4}", search_text))
        if dates:
            return min(dates, key=lambda m: abs(m.start() + start - start_pos)).group(0)
        return None

    def _heuristic_parse(self, text: str, next_page_text: str = None) -> List[Dict[str, Any]]:
        """
        Extraction ligne-à-ligne des items par heuristique (regex/keywords).
        Permet de récupérer les lignes d'items même si la structure du PDF varie.

        Args:
            text (str): Texte de la page.
            next_page_text (str, optional): Texte de la page suivante (pour certains cas multi-pages).

        Returns:
            list[dict]: Liste d'items extraits.
        """
        results = []
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        current = None
        page_pattern = re.compile(r"^page\s+\d+\s*/\s*\d+$", re.IGNORECASE)
        montant_total_pattern = re.compile(r"montant\s+total\s+ht", re.IGNORECASE)

        def find_product_name_after_montant_ht(text: str) -> Union[str, None]:
            """
            Cherche le nom du produit après 'Montant HT' dans le texte donné.
            Utile pour les cas où la désignation est sur la page suivante.
            """
            if not text:
                return None
            lines = text.splitlines()
            for i, line in enumerate(lines):
                if "Montant HT" in line:
                    for next_line in lines[i+1:]:
                        if next_line.strip() and not (next_line.split() + [None, None])[:2][0].isdigit():
                            return next_line.strip()
            return None

        for i, line in enumerate(lines):
            parts = line.split()
            # On considère une ligne d'item si les deux premiers tokens sont numériques
            if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
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
                # On tente d'extraire la date de livraison sur la même ligne
                date_in_line = self._extract_date_from_line(line)
                if date_in_line:
                    current["date_livraison"] = date_in_line
                # Sinon, on cherche une date proche du mot "livraison" dans les lignes suivantes
                if not current["date_livraison"]:
                    context = "\n".join(lines[i:min(i+5, len(lines))])
                    liv_match = re.search(r"livraison", context, re.IGNORECASE)
                    if liv_match:
                        date = self._extract_date_from_context(context, liv_match.start())
                        if date:
                            current["date_livraison"] = date
                # Extraction de la quantité, unité et prix unitaire
                for j, tok in enumerate(parts[2:], start=2):
                    if not current["quantite"] and re.fullmatch(r"\d+(?:[.,]\d+)?", tok):
                        if j + 1 < len(parts) and not re.search(r"\d", parts[j+1]):
                            current["quantite"] = tok
                            current["unite"] = parts[j+1]
                            # Le prix unitaire suit généralement l'unité
                            if j + 2 < len(parts) and re.fullmatch(r"\d{1,3}(?:[\.]\d{3})*,\d{2}", parts[j+2]):
                                current["prix_unitaire"] = parts[j+2]
                # Gestion du nom du produit (ligne suivante ou page suivante)
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if page_pattern.match(next_line) or montant_total_pattern.search(next_line):
                        # Cas où la désignation est sur la page suivante
                        if next_page_text:
                            product_name = find_product_name_after_montant_ht(next_page_text)
                            if product_name:
                                current["nom_produit"] = product_name
                    else:
                        # Cas classique : la ligne suivante contient la désignation
                        if not (next_line.split() + [None, None])[:2][0].isdigit():
                            current["nom_produit"] = next_line
        if current:
            results.append(current)
        return results

    def _merge_items(self, a: List[Dict[str, Any]], b: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Fusionne deux listes d'items selon la clé 'position'.
        Prend les valeurs non nulles de la première liste si absentes dans la seconde.

        Args:
            a (list[dict]): Première liste d'items.
            b (list[dict]): Deuxième liste d'items.

        Returns:
            list[dict]: Liste fusionnée.
        """
        merged: Dict[str, Dict[str, Any]] = {}
        for it in b:
            merged[it['position']] = {**it}
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
        return [merged[k] for k in sorted(merged, key=lambda x: int(x))]

    def parse_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """
        Traite un PDF page par page et extrait les informations structurées.

        Args:
            pdf_path (str): Chemin du fichier PDF.

        Returns:
            dict: Résultat contenant items, total_ht, numero_commande, ville.
        """
        all_items: List[Dict[str, Any]] = []
        all_texts: List[str] = []
        with pdfplumber.open(pdf_path) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            self.global_delivery_date = self._extract_global_date(first_page_text)
            print(f"Date de livraison globale trouvée: {self.global_delivery_date}")
            numero_commande = self._extract_order_number(first_page_text)
            ville = self.find_city(pdf_path)
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"Traitement de la page {page_num}...")
                text = page.extract_text()
                if text:
                    all_texts.append(text)
                try:
                    next_page_text = pdf.pages[page_num].extract_text() if page_num < len(pdf.pages) else None
                    page_items = self._heuristic_parse(text, next_page_text)
                    print(f"Heuristique a trouvé {len(page_items)} items sur la page {page_num}")
                    for item in page_items:
                        if not item.get('date_livraison'):
                            item['date_livraison'] = self.global_delivery_date
                    all_items.extend(page_items)
                    print(f"Total: {len(page_items)} items trouvés sur la page {page_num}\n")
                except Exception as e:
                    print(f"Erreur lors du traitement de la page {page_num}: {str(e)}")
                    continue
        total_text = "\n".join(all_texts)
        total_ht = self._extract_total(total_text)
        result = {
            "items": all_items,
            "total_ht": total_ht,
            "numero_commande": numero_commande,
            "ville": ville
        }
        self.last_result = result
        return result

    def _extract_order_number(self, text: str) -> Union[str, None]:
        """
        Recherche du numéro de commande dans le texte.

        Args:
            text (str): Texte à analyser.

        Returns:
            str|None: Numéro de commande trouvé ou None.
        """
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
        """
        Extrait le nom de ville potentiel d'une ligne de texte.
        Utilise le géocodeur pour vérifier si la chaîne correspond à une ville réelle.

        Args:
            text (str): Ligne à analyser.

        Returns:
            str: Ville trouvée ou chaîne vide.
        """
        def is_valid_city(word: str) -> Union[str, None]:
            """
            Vérifie si le mot est une ville réelle via Nominatim.
            Exclut certains mots-clés métiers ou techniques.
            """
            try:
                excluded_words = [
                    "poste de", "poste rte", "poste electrique", "rte", "(rte)",
                    "poteau", "rouge", "poste", "je", "electrique", "électrique",
                    "droite", "colonnes", "trois", "90kv", "smart", "jacquelin", "jeremy"
                ]
                first_word = word.strip().split()[0].lower() if word.strip() else ""
                if first_word == "de":
                    return None
                if all(c.isdigit() or c.isspace() for c in word):
                    return None
                if word.lower().startswith('de '):
                    return None
                if all(c.isdigit() or c.isspace() for c in word):
                    return None
                if "rte" in word.lower() or "kv" in word.lower():
                    return None
                if (len(word) < 3 or word.isdigit() or word.lower() in excluded_words):
                    return None
                time.sleep(1)  # Limite le débit des requêtes Nominatim
                location = self.geolocator.geocode(word, exactly_one=True)
                if location:
                    return word
                return None
            except Exception:
                return None

        words = text.strip().split()
        if not words:
            return ""
        # On teste d'abord les groupes de 3, puis 2, puis 1 mots depuis la fin de la ligne
        for length in range(3, 0, -1):
            for i in range(len(words) - 1, length - 2, -1):
                word_group = " ".join(words[i - length + 1:i + 1]).strip('.,;:-').strip()
                if city := is_valid_city(word_group):
                    return city
        return ""

    def extract_left_lines_after_obj(self, page_number: int = 0, n_lines: int = 3) -> List[str]:
        """
        Extrait les lignes de gauche après le mot 'OBJET' sur une page PDF.
        Permet de cibler la zone où la ville est souvent mentionnée.

        Args:
            page_number (int): Numéro de la page à analyser.
            n_lines (int): Nombre de lignes à extraire.

        Returns:
            list[str]: Lignes extraites.
        """
        with pdfplumber.open(self.pdf_path) as pdf:
            page = pdf.pages[page_number]
            words = page.extract_words(use_text_flow=True)
            try:
                obj = next(w for w in words if w['text'].strip().lower().startswith('objet'))
                y_obj = obj['top']
            except StopIteration:
                return []
            # On ne prend que les mots situés sous "OBJET" et dans la partie gauche de la page
            left_boundary = page.width * 0.6
            left_words = [w for w in words if w['top'] > y_obj and w['x1'] < left_boundary]
            lines_dict = defaultdict(list)
            for w in left_words:
                key = round(w['top'], 1)
                lines_dict[key].append(w)
            sorted_tops = sorted(lines_dict.keys())
            left_lines = []
            for top in sorted_tops[:n_lines]:
                line_words = sorted(lines_dict[top], key=lambda w: w['x0'])
                line = " ".join(w['text'] for w in line_words)
                left_lines.append(line)
            return left_lines

    def is_valid_city(self, city_name: str) -> bool:
        """
        Vérifie si le nom correspond à une ville valide via Nominatim.

        Args:
            city_name (str): Nom à vérifier.

        Returns:
            bool: True si c'est une ville valide, False sinon.
        """
        try:
            if "kv" in city_name.lower():
                return False
            location = self.geolocator.geocode(city_name, exactly_one=True)
            if not location:
                return False
            lieu_class = location.raw.get('class')
            lieu_type = location.raw.get('type')
            # On accepte les villes, zones administratives, cours d'eau et routes
            is_valid = (lieu_class == 'place' and lieu_type in ['city', 'town', 'village']) or \
                       (lieu_class == 'boundary' and lieu_type == 'administrative') or \
                       (lieu_class in ['waterway', "highway"])
            return is_valid
        except Exception:
            return False

    def _find_city_after_keywords(self, text: str) -> Union[str, None]:
        """
        Cherche une ville après des mots-clés spécifiques dans le texte (ex : "POSTE ELECTRIQUE DE ...").

        Args:
            text (str): Texte à analyser.

        Returns:
            str|None: Ville trouvée ou None.
        """
        patterns = [
            r"POSTE\s+ELECTRIQUE\s+DE\s+([^\n\.,]+)",
            r"POSTE\s+RTE\s+DE\s+([^\n\.,]+)"
        ]
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                potential_city = match.group(1).strip()
                if city := self._extract_ville(potential_city):
                    return city
                words = potential_city.split()
                if len(words) >= 2:
                    two_words = " ".join(words[:2])
                    if city := self._extract_ville(two_words):
                        return city
        return None

    def find_city(self, pdf_path: str) -> str:
        """
        Processus complet d'extraction et validation de ville à partir du PDF.

        Args:
            pdf_path (str): Chemin du fichier PDF.

        Returns:
            str|None: Ville trouvée ou None.
        """
        self.pdf_path = pdf_path
        # On tente d'abord dans les lignes après "OBJET"
        lines = self.extract_left_lines_after_obj(page_number=0, n_lines=2)
        for line in lines:
            if city := self._extract_ville(line):
                return city
        # Sinon, on cherche après les mots-clés dans tout le texte
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join(page.extract_text() for page in pdf.pages)
            if city := self._find_city_after_keywords(full_text):
                return city
        return None

    def _clean_number(self, value: str) -> str:
        """
        Nettoie un nombre en supprimant les points et gardant la virgule.

        Args:
            value (str): Nombre à nettoyer.

        Returns:
            str: Nombre nettoyé.
        """
        if not value:
            return value
        return str(value).replace('.', '')

    def export_to_excel(self, output_path: str = None) -> None:
        """
        Exporte les résultats extraits au format Excel.

        Args:
            output_path (str, optional): Chemin du fichier Excel de sortie.

        Raises:
            Exception: En cas d'erreur lors de l'écriture du fichier.
        """
        try:
            items_data = self.last_result["items"].copy()
            for item in items_data:
                item['numero_commande'] = self.last_result.get('numero_commande', '')
                item['ville'] = self.last_result.get('ville', '')
            for item in items_data:
                if item.get('prix_unitaire'):
                    item['prix_unitaire'] = self._clean_number(item['prix_unitaire'])
            df_items = pd.DataFrame(items_data)
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
            columns_order = [
                'numero_commande', 'ville', 'position', 'designation', 'nom_produit',
                'quantite', 'unite', 'prix_unitaire', 'date_livraison'
            ]
            df_items = df_items[columns_order].rename(columns=column_names)
            if not output_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"facture_{timestamp}.xlsx"
            if os.path.exists(output_path):
                try:
                    with open(output_path, 'a+b'):
                        pass
                except PermissionError:
                    base, ext = os.path.splitext(output_path)
                    timestamp = datetime.now().strftime("_%Y-%m-%d_%H-%M-%S")
                    output_path = f"{base}{timestamp}{ext}"
                    print(f"Fichier existant verrouillé, utilisation du nouveau nom: {output_path}")
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
                df_items.to_excel(writer, sheet_name='Items', index=False)
                df_global = pd.DataFrame({
                    'Numéro de commande': [self.last_result.get('numero_commande', '')],
                    'Ville': [self.last_result.get('ville', '')],
                    'Total HT': [self._clean_number(self.last_result.get('total_ht', ''))]
                })
                df_global.to_excel(writer, sheet_name='Informations globales', index=False)
                # Ajuste automatiquement la largeur des colonnes pour une meilleure lisibilité
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


