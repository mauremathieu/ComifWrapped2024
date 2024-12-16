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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from password import password

pdfmetrics.registerFont(TTFont('stinger', 'fonts/StingerFit-Bold.ttf'))
pdfmetrics.registerFont(TTFont('tan', 'fonts/TANHEADLINE.ttf'))

def send_email(to_email, pdf_file, prenom, nom, solde):
    from_email = "comif13120@gmail.com"
    subject = "Ton Comif Wrapped 2024 !"
    prenom = prenom.capitalize()
    nom = nom.capitalize()
    body = f"Bonjour {prenom} {nom},\n\nVoici ton Comif Wrapped 2024 ! En te souhaitant de bonne fête de fin d'année.\n\n"
    if solde < 0:
        body += f"En revanche, ton solde est négatif: {solde}€. Merci de bien vouloir régulariser ta situation au plus vite auprès de nos serveurs !\nIBAN de la Comif: FR7630004007020001006208160. Envoie nous un message si tu fait un virement pour qu'on puisse le valider.\n\n"
    body += "A bientôt, \n\nComif is Love, Comif is Life"

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(pdf_file, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(pdf_file))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_file)}"'
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, password) # password is imported from password.py
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

def create_visual(ficname):
    df = pd.read_excel(ficname)  # data loading
    columns_as_lists = {column: df[column].tolist() for column in df.columns}

    team_tibbar = (df['favorite_place'].str.contains('Tibbar').sum() / len(df)) * 100
    team_tibbar = int(team_tibbar)
    team_titpause = f"{100 - team_tibbar}%"
    team_tibbar = f"{team_tibbar}%"
    
    for i in range(len(columns_as_lists['nom'])):
        nom = columns_as_lists['nom'][i].upper()
        prenom = columns_as_lists['prenom'][i].upper()
        solde = int(columns_as_lists['solde'][i] / 100)
        top_conso = columns_as_lists['top1_produit'][i]
        top_conso = top_conso.upper() if pd.notna(top_conso) else "N/A"
        top_qtte = int(columns_as_lists['top1_quantite'][i]) if pd.notna(columns_as_lists['top1_quantite'][i]) else 0
        place = columns_as_lists['favorite_place'][i].upper()
        top2 = columns_as_lists['top2_produit'][i]
        top2 = top2.upper() if pd.notna(top2) else "N/A"
        top2_qtte = int(columns_as_lists['top2_quantite'][i]) if pd.notna(columns_as_lists['top2_quantite'][i]) else 0
        top3 = columns_as_lists['top3_produit'][i]
        top3 = top3.upper() if pd.notna(top3) else "N/A"
        top3_qtte = int(columns_as_lists['top3_quantite'][i]) if pd.notna(columns_as_lists['top3_quantite'][i]) else 0
        top4 = columns_as_lists['top4_produit'][i]
        top4 = top4.upper() if pd.notna(top4) else "N/A"
        top4_qtte = int(columns_as_lists['top4_quantite'][i]) if pd.notna(columns_as_lists['top4_quantite'][i]) else 0
        top5 = columns_as_lists['top5_produit'][i]
        top5 = top5.upper() if pd.notna(top5) else "N/A"
        top5_qtte = int(columns_as_lists['top5_quantite'][i]) if pd.notna(columns_as_lists['top5_quantite'][i]) else 0
        position_alcool = int(columns_as_lists['rang_alcool'][i]) if pd.notna(columns_as_lists['rang_alcool'][i]) else 295
        money = f"{int(columns_as_lists['total_depense'][i] / 100)}€"
        position_depense = int(columns_as_lists['rang_consommateur'][i])
        b1 = columns_as_lists['top1_biere'][i]
        b1 = b1.upper() if pd.notna(b1) else "N/A"
        b2 = columns_as_lists['top2_biere'][i]
        b2 = b2.upper() if pd.notna(b2) else "N/A"
        b3 = columns_as_lists['top3_biere'][i]
        b3 = b3.upper() if pd.notna(b3) else "N/A"
        bq1 = int(columns_as_lists['top1_biere_quantite'][i]) if pd.notna(columns_as_lists['top1_biere_quantite'][i]) else 0
        bq2 = int(columns_as_lists['top2_biere_quantite'][i]) if pd.notna(columns_as_lists['top2_biere_quantite'][i]) else 0
        bq3 = int(columns_as_lists['top3_biere_quantite'][i]) if pd.notna(columns_as_lists['top3_biere_quantite'][i]) else 0
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
        c.drawString(415, height - 457, f"{position_alcool}")


        #favourite place
        c.setFillColor(HexColor(0x282828))  # noir
        c.setFont("tan", 30)
        c.drawString(359, 288, f"{place}") 
        c.setFont("stinger", 18)
        if place == "TIBBAR":
            c.drawString(207, 288, f"{team_tibbar}")
            c.drawImage("visual/beer.png", 495, 275, width=45, height=45)
        elif place == "TITPAUSE":
            c.drawString(207, 288, f"{team_titpause}")
            c.drawImage("visual/croissant.jpg", 540, 275, width=45, height=45)
        else:
            raise ValueError(f"Place {place} not recognized")


        #top consommation
        c.setFillColor(HexColor(0xfae8c4)) # beige
        c.setFont("tan", 13)
        if len(top_conso) > 16:
            c.setFont("tan", 😎
        c.drawString(70, 160, f"{top_conso}")
        c.setFont("tan", 13)
        c.drawString(250, 160, f"{top_qtte}")

        if len(top2) > 16:
            c.setFont("tan", 😎
        c.drawString(70, 130, f"{top2}")
        c.setFont("tan", 13)
        c.drawString(250, 130, f"{top2_qtte}")

        if len(top3) > 16:
            c.setFont("tan", 😎
        c.drawString(70, 100, f"{top3}")
        c.setFont("tan", 13)
        c.drawString(250, 100, f"{top3_qtte}")

        if len(top4) > 16:
            c.setFont("tan", 😎
        c.drawString(70, 70, f"{top4}")
        c.setFont("tan", 13)
        c.drawString(250, 70, f"{top4_qtte}")

        if len(top5) > 16:
            c.setFont("tan", 😎
        c.drawString(70, 40, f"{top5}")
        c.setFont("tan", 13)
        c.drawString(250, 40, f"{top5_qtte}")
        
        
        #top biere
        c.setFillColor(HexColor(0xfae8c4)) # beige
        c.setFont("tan", 14)
        if len(b1) > 17:
            c.setFont("tan", 10)
        c.drawString(298, 188, f"{b1}")
        c.setFont("tan", 14)
        c.drawString(505, 188, f"{bq1}")  
        
        if len(b2) > 17:
            c.setFont("tan", 10)
        c.drawString(298, 158, f"{b2}")
        c.setFont("tan", 14)
        c.drawString(505, 158, f"{bq2}") 
        
        if len(b3) > 17:
            c.setFont("tan", 10)
        c.drawString(298, 128, f"{b3}")
        c.setFont("tan", 14)
        c.drawString(505, 128, f"{bq3}")     

        c.save() 
        print(f"PDF {pdf_file} created")
        send_email(email, pdf_file, prenom, nom, solde)

# Exécution de la création du visuel
if _name_ == "_main_":
    create_visual("data_comif_wrapped.xlsx")