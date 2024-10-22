import openpyxl
from PIL import Image, ImageDraw, ImageFont

wb_honorarios = openpyxl.load_workbook('relacao_honorarios.xlsx')
sheet_honorarios = wb_honorarios['honorario']

for indice, linha in enumerate (sheet_honorarios.iter_rows(min_row=2,max_row=2)):
    empresa=linha[1].value #nome da empresa
    valor=linha[2].value #valor em R$
    mes=linha[3].value #mes de referencia
    total=valor
    
    fonte_geral = ImageFont.truetype('./Roboto-MediumItalic.ttf',50)

    image= Image.open('./honorario_padrao.jpg')
    desenhar=ImageDraw.Draw(image)

    desenhar.text((840,550),str(valor),font=fonte_geral, fill='black')
    desenhar.text((357,383),str(empresa),font=fonte_geral,fill='black')
    desenhar.text((623,570),str(mes),font=fonte_geral,fill='black')

    image.save(f'./{indice} {empresa} honorarios.png')

