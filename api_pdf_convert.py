from flask import Flask, request, jsonify
import base64
import os

app = Flask(__name__)

@app.route('/status', methods=['GET'])
def check_status():
    bonjour_value = request.headers.get('bonjour', 'Valeur par défaut')
    return jsonify({"message": f"L'application fonctionne correctement.", "bonjour": bonjour_value}), 200

@app.route('/upload', methods=['POST'])
def upload_file():
    data = request.get_json()
    filename = data.get('filename')
    filecontent_base64 = data.get('filecontent')
    
    try:
        file_bytes = base64.b64decode(filecontent_base64)
        with open(filename, 'wb') as f:
            f.write(file_bytes)
        
        return jsonify({"message": f"Fichier {filename} reçu et sauvegardé avec succès."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 400

if __name__ == '__main__':
    # Replit fournit automatiquement un PORT dans les variables d'environnement
    port = int(os.environ.get("PORT", 3000))
    app.run(host='0.0.0.0', port=port)
