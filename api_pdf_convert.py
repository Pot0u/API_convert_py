from flask import Flask, request, jsonify
import base64
import os
from facture_to_excel import InvoiceParser
 
app = Flask(__name__)

def clean_fr_number(val: str) -> str:
    """Garde la virgule, enlève les zéros inutiles après la virgule."""
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
    data = request.get_json()
    filename = data.get('filename', 'fichier.pdf')
    filecontent_base64 = data.get('filecontent')

    try:
        # Sauvegarde du PDF temporaire
        pdf_path = f"/tmp/{filename}"
        with open(pdf_path, 'wb') as f:
            f.write(base64.b64decode(filecontent_base64))

        # Utiliser InvoiceParser pour parser le PDF
        parser = InvoiceParser()
        result = parser.parse_pdf(pdf_path)

        # Nettoyer les quantités/prix dans items
        for item in result["items"]:
            if "quantite" in item and item["quantite"]:
                item["quantite"] = clean_fr_number(item["quantite"])
            if "prix_unitaire" in item and item["prix_unitaire"]:
                item["prix_unitaire"] = clean_fr_number(item["prix_unitaire"])
            if "montant_total" in item and item["montant_total"]:
                item["montant_total"] = clean_fr_number(item["montant_total"])

        # Nettoyer le total global
        if "total_ht" in result and result["total_ht"]:
            result["total_ht"] = clean_fr_number(result["total_ht"])

        custom_response = {
            "items": result["items"],
            "globalite": {
                "numero_commande": result.get("numero_commande"),
                "ville": result.get("ville"),
                "total_ht": result.get("total_ht"),
            }
        }
 
        # Retourner le JSON directement
        return jsonify(custom_response), 200
 
    except Exception as e:
        return jsonify({"error": str(e)}), 400
 
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 3000))
    app.run(host='0.0.0.0', port=port)