import pdfplumber

pdf_path = r"c:\Users\clement.lam\Downloads\4500784755.pdf"  # Remplace par le chemin réel

with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    print(f"Largeur : {page.width} points, Hauteur : {page.height} points")
    # taille du pdf 595, 842
    # Exemple d'extraction d'une zone (adapte les coordonnées à ton besoin)
    x0, top, x1, bottom = 20, 425, 228, 514
    zone = page.crop((x0, top, x1, bottom))
    texte = zone.extract_text()
    print("Texte extrait de la zone :")
    print(texte)