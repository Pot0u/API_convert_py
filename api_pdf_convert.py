from flask import Flask, request, jsonify, send_file
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
        # Sauvegarder le fichier PDF temporairement
        pdf_path = f"/tmp/{filename}"
        with open(pdf_path, 'wb') as f:
            f.write(base64.b64decode(filecontent_base64))
 
        # Utiliser ton script existant pour convertir en Excel
        parser = InvoiceParser()
        parser.parse_pdf(pdf_path)
        excel_path = pdf_path.replace('.pdf', '.xlsx')
        parser.export_to_excel(excel_path)
 
        # Renvoyer le fichier Excel en réponse
        return send_file(
            excel_path,
            as_attachment=True,
            download_name=os.path.basename(excel_path),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
 
    except Exception as e:
        return jsonify({"error": str(e)}), 400
 
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 3000))
    app.run(host='0.0.0.0', port=port)