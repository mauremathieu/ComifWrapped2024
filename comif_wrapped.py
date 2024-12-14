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
import os

pdfmetrics.registerFont(TTFont('stinger', 'fonts/StingerFit-Bold.ttf'))
pdfmetrics.registerFont(TTFont('tan', 'fonts/TANHEADLINE.ttf'))



def create_visual(ficname):
    df = pd.read_excel(ficname)  # data loading
    columns_as_lists = {column: df[column].tolist() for column in df.columns}

    team_tibbar = (df['favorite_place'].str.contains('Tibbar').sum() / len(df)) * 100
    team_titpause = 100 - team_tibbar
    
    for i in range(5):
        nom = columns_as_lists['nom'][i].upper()
        prenom = columns_as_lists['prenom'][i].upper()
        top_conso = columns_as_lists['top1_produit'][i].upper()
        top_qtte = int(columns_as_lists['top1_quantite'][i])
        place = columns_as_lists['favorite_place'][i].upper()
        top2 = columns_as_lists['top2_produit'][i].upper()
        top2_qtte = int(columns_as_lists['top2_quantite'][i])
        top3 = columns_as_lists['top3_produit'][i].upper()
        top3_qtte = int(columns_as_lists['top3_quantite'][i])
        position_alcool = int(columns_as_lists['rang_alcool'][i])
        money = f"{int(columns_as_lists['total_depense'][i] / 100)}€"
        position_depense = int(columns_as_lists['rang_consommateur'][i])
        b1 = columns_as_lists['top1_biere'][i].upper()
        b2 = columns_as_lists['top2_biere'][i].upper()
        b3 = columns_as_lists['top3_biere'][i].upper()
        bq1 = int(columns_as_lists['top1_biere_quantite'][i])
        bq2 = int(columns_as_lists['top2_biere_quantite'][i])
        bq3 = int(columns_as_lists['top3_biere_quantite'][i])
        email = columns_as_lists['email'][i]

        pdf_file = f'wrapped_export/wrapped_comif_2024_{prenom}_{nom}.pdf'
        c = canvas.Canvas(pdf_file, pagesize=A4)
        width, height = A4  # canva PDF A4
        c.drawImage("template_wrapped_comif.png", 0, 0, width=width, height=height)
        
        #nom et prenom
        c.setFillColor(HexColor(0x595757))  # gris
        if len(prenom) > 7:
            c.setFont("tan", 30)
        else:
            c.setFont("tan", 40)
        c.drawString(90, height - 290, f"{prenom}")
        
        if len(nom) > 8:
            c.setFont("tan", 25)
        else:
            c.setFont("tan", 35)
        c.drawString(80, height - 340, f"{nom}")


        #prix et top consommation
        c.setFillColor(HexColor(0x06873D))  # vert
        c.setFont("tan", 40)
        c.drawString(375, height - 235, f"{money}")
        c.setFont("stinger", 15)
        c.drawString(404, height - 347, f"{position_depense}")


        #top buveur
        c.setFillColor(HexColor(0xfae8c4)) # beige
        c.setFont("tan", 37)
        c.drawString(415, height - 460, f"{position_alcool}")


        #favourite place
        c.setFillColor(HexColor(0x282828))  # noir
        c.setFont("tan", 30)
        c.drawString(370, 300, f"{place}")  # X = 25cm, Y = 36cm
        c.setFont("stinger", 18)
        if place == "TIBBAR":
            c.drawString(320, 300, f"{team_tibbar}")
            c.drawImage("visual/beer.png", 370, 300, width=50, height=50)
        elif place == "TITPAUSE":
            c.drawString(320, 300, f"{team_titpause}")
            c.drawImage("visual/beer.png", 370, 300, width=50, height=50)
        else:
            raise ValueError(f"Place {place} not recognized")


        #top consommation
        c.setFillColor(HexColor(0xfae8c4)) # beige
        c.setFont("tan", 20)
        c.drawString(40, height - 465, f"{top_conso}")
        c.drawString(40, height - 465, f"{top_qtte}")
        
        
        #top biere
        c.setFillColor(HexColor(0xfae8c4)) # beige
        c.setFont("tan", 15)
        c.drawString(40, height - 465, f"{b1}")
        c.drawString(40, height - 465, f"{bq1}")        

        
        c.save()  # sauvegarde

        print(f"PDF {pdf_file} created")

# Exécution de la création du visuel
if __name__ == "__main__":
    create_visual("data_comif_wrapped.xlsx")
