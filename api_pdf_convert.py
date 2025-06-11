"""
api_pdf_convert.py

API Flask pour recevoir un PDF encodé en base64, extraire les informations de facture,
et retourner le résultat en JSON.

Fonctionnalités principales :
- Décode le PDF reçu en base64 et le sauvegarde temporairement.
- Utilise InvoiceParser pour extraire les données structurées.
- Nettoie les champs numériques pour un format français cohérent.
- Retourne un JSON structuré avec les items et les informations globales.

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
    Nettoie un nombre au format français : garde la virgule, enlève les zéros inutiles après la virgule.

    Args:
        val (str): Nombre à nettoyer.

    Returns:
        str: Nombre nettoyé.
    """
    if not isinstance(val, str):
        return val
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
    """
    data = request.get_json()
    filename = data.get('filename', 'fichier.pdf')
    filecontent_base64 = data.get('filecontent')

    try:
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

        custom_response = {
            "items": result["items"],
            "globalite": {
                "numero_commande": result.get("numero_commande"),
                "objet": result.get("objet"),
                "lieu_livraison": result.get("lieu_livraison"),
                "total_ht": result.get("total_ht"),
            }
        }

        return jsonify(custom_response), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 400
 
if __name__ == '__main__':
    # Permet de lancer l'API en local ou sur un port défini par la variable d'environnement PORT.
    port = int(os.environ.get("PORT", 3000))
    app.run(host='0.0.0.0', port=port)