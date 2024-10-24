import openpyxl
from PIL import Image, ImageDraw, ImageFont
import os

# Função para dividir o texto em várias linhas com base na largura máxima
def quebra_texto(texto, fonte, largura_maxima, desenhar):
    linhas = []
    palavras = texto.split()
    linha_atual = ""

    for palavra in palavras:
        linha_com_palavra = linha_atual + " " + palavra if linha_atual else palavra
        largura_texto, altura_texto = desenhar.textbbox((0, 0), linha_com_palavra, font=fonte)[2:]

        if largura_texto <= largura_maxima:
            linha_atual = linha_com_palavra
        else:
            linhas.append(linha_atual)
            linha_atual = palavra

    if linha_atual:
        linhas.append(linha_atual)

    return linhas

# Carregando a planilha com os valores calculados das fórmulas
wb_honorarios = openpyxl.load_workbook('relacao_honorarios.xlsx', data_only=True)
sheet_honorarios = wb_honorarios['honorario']

for indice, linha in enumerate(sheet_honorarios.iter_rows(min_row=2,max_row=94)):
    empresa = linha[1].value  # nome da empresa
    valor = linha[2].value  # valor em R$
    mes = 'oct/24'
    total = linha[21].value  # valor total calculado
    recalc_fgts = 'RECALC.FGTS'
    desconto = 'DESCONTO'
    descricao_outros= linha[20].value #descricao de outros
    valor_outros = 'R$'+ str(linha[18].value) #valor do "outros"
    simone = linha[23].value #linha p verificar se os recibos sao da simone ou n

    fonte_geral = ImageFont.truetype('./Roboto-MediumItalic.ttf', 50)

    image = Image.open('./honorario_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # Definir a largura máxima permitida para o nome da empresa
    largura_maxima_empresa = 800  # ajuste conforme necessário
    coordenada_inicial_empresa = (340, 350)

    # Sanitizar o nome da empresa para evitar problemas ao salvar o arquivo
    empresa_sanitizada = empresa.replace("/", "").replace("\\", "").replace(":", "").replace("*", "")

    # Dividir o nome da empresa em várias linhas, se necessário
    linhas_empresa = quebra_texto(empresa_sanitizada, fonte_geral, largura_maxima_empresa, desenhar)

    # Verificar se o texto excede a largura máxima e ajustar a posição Y se necessário
    if len(linhas_empresa) > 1:  # Se houver mais de uma linha
        coordenada_inicial_empresa = (340, 320)  # Ajusta Y (mais acima)
    
    # Desenhar cada linha do nome da empresa, ajustando a altura
    altura_linha = 60  # distância entre as linhas
    for i, linha_empresa in enumerate(linhas_empresa):
        desenhar.text((coordenada_inicial_empresa[0], coordenada_inicial_empresa[1] + i * altura_linha),
                      linha_empresa, font=fonte_geral, fill='black')

    #ver se o recibo contém algo além do honorário padrao.
    if descricao_outros and 'RECALC. FGTS' in descricao_outros.strip():
        desenhar.text((187,929), recalc_fgts, font=fonte_geral, fill='black')
        desenhar.text((795,925), str(valor_outros), font=fonte_geral, fill='black')

    if descricao_outros and 'DESCONTO' in descricao_outros.strip():
        desenhar.text((187,929), desconto, font=fonte_geral, fill='black')
        desenhar.text((795,925), str(valor_outros), font=fonte_geral, fill='black')

     # Desenhar outros valores na imagem
    desenhar.text((840, 550), str(valor), font=fonte_geral, fill='black')
    desenhar.text((610, 550), str(mes), font=fonte_geral, fill='black')
    desenhar.text((835, 1045), str(total), font=fonte_geral, fill='black')

   # Definir o caminho para a pasta "Recibos"
    pasta_recibos = os.path.join(os.path.expanduser("~"), "Desktop", "Recibos")
    # Definir o caminho para a pasta "Recibos_Simone"
    pasta_recibos_simone = os.path.join(os.path.expanduser("~"), "Desktop", "Recibos_Simone")

    # Se você quiser garantir que o arquivo seja salvo lá
    os.makedirs(pasta_recibos, exist_ok=True)
    os.makedirs(pasta_recibos_simone,exist_ok=True)

    # ... [seu código anterior] ...

    # Salvar a imagem com o nome sanitizado na pasta Resultado
    if simone and 'sim' in simone.strip():
        caminho_arquivo = os.path.join(pasta_recibos_simone, f'{empresa_sanitizada}_recibo.pdf')
        image.save(caminho_arquivo)
    else:
        caminho_arquivo = os.path.join(pasta_recibos, f'{empresa_sanitizada}_recibo.pdf')
        image.save(caminho_arquivo)