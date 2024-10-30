import openpyxl
from PIL import Image, ImageDraw, ImageFont
import os

wb_alteracoes = openpyxl.load_workbook('balancete_alteracoes.xlsx', data_only=True)
sheet_alteracoes = wb_alteracoes['alteracoes']



for indice, linha in enumerate(sheet_alteracoes.iter_rows(min_row=4,max_row=4)):

    #BOLETOS DE ALTERACOES
    empresa_alteracoes =linha[0].value #nome da empresa
    valor_alteracoes =linha[1].value
    ref = linha[2].value #mes de referencia da alteracao
    vencimento ='10/11/2024'
    cnpj =linha[4].value

    empresa_sanitizada = empresa_alteracoes.replace("/", "").replace("\\", "").replace(":", "").replace("*", "")

    fonte_geral = ImageFont.truetype('./Roboto-MediumItalic.ttf', 16)
    fonte_mes = ImageFont.truetype('./Roboto-MediumItalic.ttf', 15)

    #define a imagem q vai abrir pra desenhar
    image_alteracoes = Image.open('boleto_alteracao.jpg')
    desenhar_alteracoes = ImageDraw.Draw(image_alteracoes)


    # Desenhar no boleto de alteracoes
    desenhar_alteracoes.text((271, 776), str(ref), font=fonte_mes, fill='black')
    desenhar_alteracoes.text((1018, 189), 'R$'+ str(linha[1].value if linha[1].value is not None else 0)+',00', font=fonte_geral, fill='black')
    desenhar_alteracoes.text((1018, 628), 'R$'+ str(linha[1].value if linha[1].value is not None else 0)+',00', font=fonte_geral, fill='black')
    desenhar_alteracoes.text((645,189), str(vencimento),font=fonte_geral,fill='black')
    desenhar_alteracoes.text((75,282), empresa_alteracoes, font=fonte_geral,fill='black')
    desenhar_alteracoes.text((75,875), empresa_alteracoes, font=fonte_geral,fill='black')
    desenhar_alteracoes.text((75,305),str(cnpj),font=fonte_geral,fill='black')
    desenhar_alteracoes.text((75,900),str(cnpj),font=fonte_geral,fill='black')
    

    # Definir o caminho para a pasta "Boleto_Alteracoes"
    pasta_boletos_alteracoes = os.path.join(os.path.expanduser("~"), "Desktop", "Boletos_Alteracoes")
    
    # garantir q o arquivo seja salvo
    os.makedirs(pasta_boletos_alteracoes, exist_ok=True)

    caminho_arquivo = os.path.join(pasta_boletos_alteracoes, f'{empresa_sanitizada}_boleto.pdf')
    image_alteracoes.save(caminho_arquivo) #fim