import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def load_data(excel_file):
    df = pd.read_excel(excel_file)
    return df


# Fonction pour générer un graphique
def generate_chart(data, file_name='chart.png'):
    # Créer un graphique (par exemple, un histogramme des morceaux écoutés)
    plt.figure(figsize=(10, 6))
    plt.bar(data['Produit'], data['Quantite'], color='skyblue')
    plt.title("Produit les plus consommés", fontsize=14)
    plt.xlabel("Produits", fontsize=12)
    plt.ylabel("Quantité de commandes", fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    # Sauvegarder le graphique en tant qu'image
    plt.savefig(file_name)
    plt.close()

# Fonction pour générer le visuel dans un PDF au format A4
def create_visual(ficname):
    for i in range(2):
        df = pd.read_excel(ficname)
        columns_as_lists = {column: df[column].tolist() for column in df.columns}
        nom = columns_as_lists['nom'][i]
        pdf_file=f'repertoire_annuel_{nom}.pdf'
    # Initialiser le canvas pour le PDF
        c = canvas.Canvas(pdf_file, pagesize=A4)
        width, height = A4  # Dimensions de la page A4
        
    
    # Vérifiez les noms des colonnes pour vous assurer d'utiliser les bons
        print(f"Colonnes disponibles : {df.columns}")

    # Charger les données nécessaires
          # Exemple d'accès à la première ligne de la colonne 'nom'
        prenom = columns_as_lists['prenom'][i]
        top_conso = columns_as_lists['top1_produit'][i]
        top_qtte = columns_as_lists['top1_quantite'][i]
        place = columns_as_lists['favorite_place'][i]
        top2 = columns_as_lists['top2_produit'][i]
        top2_qtte = columns_as_lists['top2_quantite'][i]
        top3 = columns_as_lists['top3_produit'][i]
        top3_qtte = columns_as_lists['top3_quantite'][i]
        # Ajouter un titre principal
        c.setFont("Helvetica-Bold", 24)
        c.drawString(50, height - 50, f"Récapitulatif de l'année Comif de {prenom} {nom}")

    # Ajouter un texte avec des infos principales
        c.setFont("Helvetica", 14)
        c.drawString(100, height - 100, f"Tu as consommé {top_conso} {top_qtte} fois")
        c.drawString(100, height - 160, f"Tu préfères {place}")

    
        chart_file = 'chart.png'
        generate_chart(pd.DataFrame({'Produit': [top_conso,top2,top3],'Quantite': [top_qtte,top2_qtte,top3_qtte]}))

        c.drawImage(chart_file, 100, height - 500, width=300, height=250)

    # Ajouter une image ou une icône
        icon = Image.open("logo.png")  # Assurez-vous d'avoir une image dans le même dossier
        icon = icon.resize((100, 100))  # Redimensionner l'icône si nécessaire
        icon.save("resized_icon.png")

        c.drawImage("resized_icon.png", 500, height - 450, width=100, height=100)

    # Sauvegarder le fichier PDF
        c.save()

# Exécution de la création du visuel
if __name__ == "__main__":
    create_visual("data_comif_wrapped.xlsx")
