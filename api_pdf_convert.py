from flask import Flask, request, jsonify
import base64
import os

# Importer directement ta classe existante
from facture_to_excel import InvoiceParser

app = Flask(__name__)

@app.route('/convert-pdf', methods=['POST'])
def convert_pdf():
    try:
        file_bytes = request.data
        if not file_bytes or len(file_bytes) < 100:
            raise ValueError("Fichier vide ou trop petit")

        os.makedirs('received_files', exist_ok=True)
        pdf_path = os.path.join('received_files', 'document.pdf')
        with open(pdf_path, 'wb') as f:
            f.write(file_bytes)

        # Appeler ton traitement avec ton script
        from facture_to_excel import InvoiceParser
        parser = InvoiceParser()
        parser.parse_pdf(pdf_path)
        excel_path = pdf_path.replace('.pdf', '.xlsx')
        parser.export_to_excel(excel_path)

        # Lire l'Excel et renvoyer
        with open(excel_path, 'rb') as f:
            excel_base64 = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({
            "status": "success",
            "excel_file_base64": excel_base64
        }), 200

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
