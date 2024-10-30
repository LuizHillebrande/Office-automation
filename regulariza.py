import openpyxl
from PIL import Image, ImageDraw, ImageFont
import os

wb_alteracoes = openpyxl.load_workbook('balancete_alteracoes.xlsx', data_only=True)
sheet_alteracoes = wb_alteracoes['alteracoes']



for indice, linha in enumerate(sheet_alteracoes.iter_rows(min_row=2,max_row=16)):
    
    #BOLETOS DE ALTERACOES
    empresa_alteracoes =linha[0].value #nome da empresa
    valor_alteracoes =linha[1].value
    ref = linha[2].value #mes de referencia da alteracao
    vencimento ='15/11/2024'
    cnpj =linha[4].value
    descricao_outros=linha[5].value #descricao de outros
    obs=linha[3].value #so pra validar

    empresa_sanitizada = empresa_alteracoes.replace("/", "").replace("\\", "").replace(":", "").replace("*", "")

    fonte_geral = ImageFont.truetype('./Roboto-MediumItalic.ttf', 16)
    fonte_mes = ImageFont.truetype('./Roboto-MediumItalic.ttf', 15)

    #define a imagem q vai abrir pra desenhar
    image_regularizacoes = Image.open('boleto_regularizacao.jpg')
    desenhar_regularizacoes = ImageDraw.Draw(image_regularizacoes)

    algo_desenhado = False

     # Definir o caminho para a pasta "Boletos_Regularizacoes"
    pasta_boletos_regularizacao = os.path.join(os.path.expanduser("~"), "Desktop", "Boletos_Regularizacoes")

    if obs and 'outro' in obs.strip(): # se tiver escrito outro no boleto, ele desenha, se nao n
        algo_desenhado = True # vira verdadeiro
    # Desenhar no boleto de regularizacoes
        desenhar_regularizacoes.text((1018, 189), 'R$'+ str(linha[1].value if linha[1].value is not None else 0)+',00', font=fonte_geral, fill='black')
        desenhar_regularizacoes.text((1018, 628), 'R$'+ str(linha[1].value if linha[1].value is not None else 0)+',00', font=fonte_geral, fill='black')
        desenhar_regularizacoes.text((645,189), str(vencimento),font=fonte_geral,fill='black')
        desenhar_regularizacoes.text((75,282), empresa_alteracoes, font=fonte_geral,fill='black')
        desenhar_regularizacoes.text((75,875), empresa_alteracoes, font=fonte_geral,fill='black')
        desenhar_regularizacoes.text((75,305),str(cnpj),font=fonte_geral,fill='black')
        desenhar_regularizacoes.text((75,900),str(cnpj),font=fonte_geral,fill='black')
        desenhar_regularizacoes.text((70, 776), str(descricao_outros), font=fonte_mes, fill='black')

    if algo_desenhado:
            caminho_arquivo = os.path.join(pasta_boletos_regularizacao, f'{empresa_sanitizada}_boleto.pdf')
            image_regularizacoes.save(caminho_arquivo)
    
    # garantir q o arquivo seja salvo
    os.makedirs(pasta_boletos_regularizacao, exist_ok=True)

    