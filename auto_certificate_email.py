import pandas as pd
from PIL import Image, ImageDraw, ImageFont

data = pd.read_excel('names.xlsx')

name_list = data["Name"].tolist()
id_list = data['ID'].tolist()
email_list = data['Email'].tolist()

for name, i, e in zip(name_list, id_list, email_list):
    im = Image.open("cert.jpg")
    d = ImageDraw.Draw(im)
    location = (320, 216)
    text_color = (0, 0, 0)
    font = ImageFont.truetype("CHOPS___.TTF", 50)
    d.text(location, name, fill=text_color, font=font)
    location = (405, 284)
    text_color = (0, 0, 0)
    font = ImageFont.truetype("CHOPS___.TTF", 25)
    d.text(location, f"ID: {i}", fill=text_color, font=font)
    file_name = f"certificate_{i}.pdf"
    im.save(file_name)