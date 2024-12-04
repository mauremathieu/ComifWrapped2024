import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from matplotlib import font_manager
from email.mime.base import MIMEBase
from email import encoders


pdfmetrics.registerFont(TTFont('abibas', 'Abibas.ttf'))
pdfmetrics.registerFont(TTFont('aloja', 'Aloja.ttf'))
prop = font_manager.FontProperties(fname='Aloja.ttf')

# Fonction pour générer un graphique
def generate_chart_glob(data, file_name='chartprod.png'):
    plt.figure(figsize=(10, 6))
    bars = plt.bar(data['Produit'], data['Quantite'], color="#AB7100")
    for bar in bars:
        yval = bar.get_height() 
        plt.text(bar.get_x() + bar.get_width() / 2, yval ,str(yval),ha='center', va='bottom', fontsize=30,fontproperties = prop,color = '#006666')
    plt.ylabel("Quantité consommée", fontsize=14)
    plt.tight_layout()
    plt.savefig(file_name,transparent=True)
    plt.close()

def generate_chart_biere(data, file_name='chartbiere.png'):
    plt.figure(figsize=(10, 6))
    bars = plt.bar(data['Produit'], data['Quantite'], color="#AB7100")
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, yval,str(yval),ha='center', va='bottom', fontsize=20,fontproperties = prop,color = '#006666')
    plt.xticks([])
    plt.xlabel("Bières", fontsize=14)
    plt.ylabel("Quantité avalée", fontsize=14)
    plt.tight_layout()
    plt.savefig(file_name,transparent=True)
    plt.close()



# Fonction pour générer le visuel dans un PDF au format A4
def create_visual(ficname):
    df = pd.read_excel(ficname)# data loading
    columns_as_lists = {column: df[column].tolist() for column in df.columns}
    print(f"Colonnes disponibles : {df.columns}")#check nom de colonne
    for i in range(3):
        nom = columns_as_lists['nom'][i]#création de variable
        prenom = columns_as_lists['prenom'][i]
        top_conso = columns_as_lists['top1_produit'][i]
        top_qtte = int(columns_as_lists['top1_quantite'][i])
        place = columns_as_lists['favorite_place'][i]
        top2 = columns_as_lists['top2_produit'][i]
        top2_qtte = int(columns_as_lists['top2_quantite'][i])
        top3 = columns_as_lists['top3_produit'][i]
        top3_qtte = int(columns_as_lists['top3_quantite'][i])
        pos = int(columns_as_lists['rang_alcool'][i])
        money = columns_as_lists['total_depense'][i]/100
        b1 = columns_as_lists['top1_biere'][i]
        b2 = columns_as_lists['top2_biere'][i]
        b3 = columns_as_lists['top3_biere'][i]
        bq1 = int(columns_as_lists['top1_biere_quantite'][i])
        bq2 = int(columns_as_lists['top2_biere_quantite'][i])
        bq3 = int(columns_as_lists['top3_biere_quantite'][i])
        email = b3 = columns_as_lists['email'][i]

        pdf_file=f'wrapped_blaze/wrapped_{prenom}_{nom}.pdf'
        c = canvas.Canvas(pdf_file, pagesize=A4) 
        width, height = A4  #canva PDF A4
        c.drawImage("Background_1.png", 0, 0, width=width, height=height)
        c.setFont("abibas", 15)
        c.setFillColor(HexColor(0xab7100)) #marron
        c.drawString(40, height-465, f"{top_conso}")
        c.setFont("abibas", 40)
        c.drawString(445, height-450, f"{top_qtte}") #informations 
        c.setFont("aloja", 18)
        c.setFillColor(HexColor(0x006666)) #bleu vert
        c.drawString(270, height-290, f"{prenom}")
        c.drawString(350, height-290, f"{nom}")
        c.drawString(420, height-326, f"{pos}")
        c.drawString(375, height-543, f"{money}")

        chart_file = 'chartprod.png'
        generate_chart_glob(pd.DataFrame({'Produit': [top_conso,top2,top3],'Quantite': [top_qtte,top2_qtte,top3_qtte]}))
        c.drawImage(chart_file, 40, height - 785, width=250, height=150)

        chart_file = 'chartbiere.png'
        generate_chart_biere(pd.DataFrame({'Produit': [b1,b2,b3],'Quantite': [bq1,bq2,bq3]}))
        c.drawImage(chart_file, 310, height - 785, width=250, height=150)

        c.setFont("abibas", 10)
        c.setFillColor(HexColor(0xab7100)) #marron
        c.drawString(40, height-800, f"1.{top_conso}  2.{top2}  3.{top3}")
        c.drawString(310, height-800, f"1.{b1}  2.{b2}  3.{b3}")
        c.save() # sauvegarde

# Exécution de la création du visuel
if __name__ == "__main__":
    create_visual("data_comif_wrapped.xlsx")
