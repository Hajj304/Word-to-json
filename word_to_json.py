📜 محتوى ملف `word_to_json.py`

import aspose.words as aw
from bs4 import BeautifulSoup
import json
from google.colab import files

# Étape 1 : Uploader un fichier Word depuis l'utilisateur
print("📤 Veuillez téléverser un fichier Word (.docx)")
uploaded = files.upload()

# Récupérer le nom du fichier Word
filename = list(uploaded.keys())[0]

# Étape 2 : Convertir le document Word en HTML
print("📄 Conversion du Word en HTML...")
doc = aw.Document(filename)
doc.save("converted.html", aw.SaveFormat.HTML)

# Étape 3 : Lire le HTML
with open("converted.html", "r", encoding="utf-8") as file:
    soup = BeautifulSoup(file.read(), "html.parser")

# Étape 4 : Extraire les sections et paragraphes
sections = []
current_section = {"title": "Document", "content": []}

for tag in soup.body.find_all(["h1", "h2", "h3", "p"]):
    if tag.name.startswith("h"):
        if current_section["content"]:
            sections.append(current_section)
        current_section = {
            "title": tag.get_text(strip=True),
            "content": []
        }
    elif tag.name == "p":
        current_section["content"].append(tag.get_text(strip=True))

if current_section["content"]:
    sections.append(current_section)

# Étape 5 : Sauvegarder le JSON
with open("output.json", "w", encoding="utf-8") as f:
    json.dump(sections, f, ensure_ascii=False, indent=2)

# Étape 6 : Télécharger le fichier JSON
print("📥 Téléchargement du fichier JSON...")
files.download("output.json")
