"""
api_pdf_convert.py

API Flask pour recevoir un PDF encodé en base64, extraire les informations de facture,
et retourner le résultat en JSON.

Fonctionnalités principales :
- Décode le PDF reçu en base64 et le sauvegarde temporairement.
- Utilise InvoiceParser pour extraire les données structurées (items, objet, lieu, etc.).
- Nettoie les champs numériques pour un format français cohérent.
- Retourne un JSON structuré avec les items, l'objet, le lieu de livraison et les informations globales.

Auteur  : Lam Clément
Date    : 2024-06
Licence : Usage interne VINCI Energies

Dépendances :
- Flask
- facture_to_excel (InvoiceParser)
"""

from flask import Flask, request, jsonify
import base64
import os
from facture_to_excel import InvoiceParser
 
app = Flask(__name__)

def clean_fr_number(val: str) -> str:
    """
    Nettoie un nombre au format français : enlève les points, garde la virgule.

    Args:
        val (str): Nombre à nettoyer.

    Returns:
        str: Nombre nettoyé.
    """
    if not isinstance(val, str):
        return val
    val = val.replace('.', '')  # Enlève les séparateurs de milliers
    if ',' in val:
        entier, dec = val.split(',')
        dec = dec.rstrip('0')
        if dec == '':
            return entier
        return f"{entier},{dec}"
    return val
 
@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Endpoint principal de l'API.
    Reçoit un PDF encodé en base64, l'analyse et retourne les informations extraites.

    Entrée attendue (JSON):
        - filename: nom du fichier PDF (optionnel)
        - filecontent: contenu du PDF encodé en base64

    Retour:
        - 200: JSON structuré avec items, globalité, objet, lieu_livraison
        - 400: JSON d'erreur en cas de problème

    Exceptions:
        - Retourne une erreur 400 si le décodage ou le parsing échoue.
    """
    data = request.get_json()
    filename = data.get('filename', 'fichier.pdf')
    filecontent_base64 = data.get('filecontent')

    try:
        # On sauvegarde le PDF décodé dans un fichier temporaire pour le parser.
        pdf_path = f"/tmp/{filename}"
        with open(pdf_path, 'wb') as f:
            f.write(base64.b64decode(filecontent_base64))

        parser = InvoiceParser()
        result = parser.parse_pdf(pdf_path)

        # Nettoyage des champs numériques pour garantir un format homogène côté client.
        for item in result["items"]:
            if "quantite" in item and item["quantite"]:
                item["quantite"] = clean_fr_number(item["quantite"])
            if "prix_unitaire" in item and item["prix_unitaire"]:
                item["prix_unitaire"] = clean_fr_number(item["prix_unitaire"])
            if "montant_total" in item and item.get("montant_total"):
                item["montant_total"] = clean_fr_number(item["montant_total"])

        if "total_ht" in result and result["total_ht"]:
            result["total_ht"] = clean_fr_number(result["total_ht"])

        # Extraction des lignes pour objet et lieu de livraison (pour affichage détaillé)
        objet_lines = parser.extract_lines_after_objet(pdf_path)
        objet_json = [{"objet": line} for line in objet_lines]

        lieu_livraison_lines = result.get("lieu_livraison", "").replace('\r', '').split('\n')
        lieu_livraison_lines = [line.strip() for line in lieu_livraison_lines if line.strip()]
        lieu_livraison_json = [{"lieu_livraison": line} for line in lieu_livraison_lines]

        custom_response = {
            "items": result["items"],
            "globalite": {
                "numero_commande": result.get("numero_commande"),
                "total_ht": result.get("total_ht"),
            },
            "objet": objet_json,
            "lieu_livraison": lieu_livraison_json
        }

        return jsonify(custom_response), 200

    except Exception as e:
        # On retourne l'erreur au client pour faciliter le debug côté front ou client API.
        return jsonify({"error": str(e)}), 400
 
if __name__ == '__main__':
    # Permet de lancer l'API en local ou sur un port défini par la variable d'environnement PORT.
    port = int(os.environ.get("PORT", 3000))
    app.run(host='0.0.0.0', port=port)