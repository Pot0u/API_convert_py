from flask import Flask, request, jsonify
import base64
import os
from facture_to_excel import InvoiceParser
 
app = Flask(__name__)
 
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

        custom_response = {
            "items": result["items"],
            "globalite": {
                "numero_commande": result.get("numero_commande"),
                "ville": result.get("ville"),
                "total_ht": result.get("total_ht"),
            },
            "excel_file_base64": excel_base64
        }
 
        # Retourner le JSON directement
        return jsonify(custom_response), 200
 
    except Exception as e:
        return jsonify({"error": str(e)}), 400
 
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 3000))
    app.run(host='0.0.0.0', port=port)