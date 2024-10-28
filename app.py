import openpyxl
from PIL import Image, ImageDraw, ImageFont
import os
import tkinter as tk
from tkinter import messagebox, font

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

def gerar_boletos():
    try:
        # Carregando a planilha com os valores calculados das fórmulas
        wb_honorarios = openpyxl.load_workbook('relacao_honorarios.xlsx', data_only=True)
        sheet_honorarios = wb_honorarios['honorario']

        for indice, linha in enumerate(sheet_honorarios.iter_rows(min_row=2, max_row=20)):
            # BOLETOS DE HONORARIOS
            empresa = linha[1].value  # nome da empresa
            valor = linha[2].value  # valor em R$
            mes = '10/24'
            total = linha[21].value if linha[21].value is not None else 0  # valor total calculado
            recalc_fgts = 'RECALC.FGTS'
            desconto = 'DESCONTO'
            descricao_outros = linha[20].value  # descricao de outros
            valor_outros = 'R$' + str(linha[18].value)  # valor do "outros"
            simone = linha[23].value  # linha p verificar se os recibos sao da simone ou n
            claudio = linha[24].value  # linha p verificar se os recibos sao do claudio ou n
            email = linha[25].value  # linha p verificar se os recibos sao por email ou n
            vencimento = '05/11/2024'
            cnpj = 'CNPJ ' + str(linha[26].value)

            fonte_geral = ImageFont.truetype('./Roboto-MediumItalic.ttf', 16)
            fonte_mes = ImageFont.truetype('./Roboto-MediumItalic.ttf', 15)

            image = Image.open('./boleto_padrao.jpg')
            desenhar = ImageDraw.Draw(image)

            # Definir a largura máxima permitida para o nome da empresa
            largura_maxima_empresa = 800  # ajuste conforme necessário
            coordenada_inicial_empresa = (75, 875)

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

            # Ver se o recibo contém algo além do honorário padrão.
            if descricao_outros and 'RECALC. FGTS' in descricao_outros.strip():
                desenhar.text((650, 230), recalc_fgts, font=fonte_mes, fill='black')
                desenhar.text((775, 230), str(valor_outros), font=fonte_mes, fill='black')
                desenhar.text((945, 795), recalc_fgts, font=fonte_mes, fill='black')
                desenhar.text((1045, 795), str(valor_outros) + ',00', font=fonte_mes, fill='black')

            if descricao_outros and 'DESCONTO' in descricao_outros.strip():
                desenhar.text((90, 230), str(valor_outros) + ',00', font=fonte_geral, fill='black')
                desenhar.text((1018, 670), str(valor_outros) + ',00', font=fonte_geral, fill='black')

            # Desenhar outros valores na imagem
            desenhar.text((1018, 189), 'R$' + str(linha[21].value if linha[21].value is not None else 0) + ',00', font=fonte_geral, fill='black')
            desenhar.text((271, 776), str(mes), font=fonte_mes, fill='black')
            desenhar.text((1018, 628), 'R$' + str(linha[21].value if linha[21].value is not None else 0) + ',00', font=fonte_geral, fill='black')
            desenhar.text((645, 189), str(vencimento), font=fonte_geral, fill='black')
            desenhar.text((75, 282), empresa, font=fonte_geral, fill='black')
            desenhar.text((75, 305), str(cnpj), font=fonte_geral, fill='black')
            desenhar.text((75, 900), str(cnpj), font=fonte_geral, fill='black')
            desenhar.text((1003,516),str(vencimento),font=fonte_geral,fill='black')

            # Definir o caminho para as pastas
            pasta_boletos_wpp = os.path.join(os.path.expanduser("~"), "Desktop", "Boletos_Wpp")
            pasta_boletos_simone = os.path.join(os.path.expanduser("~"), "Desktop", "Boletos_Simone")
            pasta_boletos_claudio = os.path.join(os.path.expanduser("~"), "Desktop", "Boletos_Claudio")
            pasta_boletos_email = os.path.join(os.path.expanduser("~"), "Desktop", "Boletos_email")

            # garantir que as pastas existam
            os.makedirs(pasta_boletos_wpp, exist_ok=True)
            os.makedirs(pasta_boletos_simone, exist_ok=True)
            os.makedirs(pasta_boletos_claudio, exist_ok=True)
            os.makedirs(pasta_boletos_email, exist_ok=True)

            # Salvar a imagem com o nome sanitizado na pasta correspondente
            if simone and 'sim' in simone.strip():
                caminho_arquivo = os.path.join(pasta_boletos_simone, f'{empresa_sanitizada}_boleto.pdf')
                image.save(caminho_arquivo)
            elif claudio and 'sim' in claudio.strip():
                caminho_arquivo = os.path.join(pasta_boletos_claudio, f'{empresa_sanitizada}_boleto.pdf')
                image.save(caminho_arquivo)
            elif email and 'sim' in email.strip():
                caminho_arquivo = os.path.join(pasta_boletos_email, f'{empresa_sanitizada}_boleto.pdf')
                image.save(caminho_arquivo)
            else:
                caminho_arquivo = os.path.join(pasta_boletos_wpp, f'{empresa_sanitizada}_boleto.pdf')
                image.save(caminho_arquivo)

        messagebox.showinfo("Sucesso", "Todos os boletos foram gerados.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

def centralizar_janela(tamanho):
    largura, altura = tamanho
    x = (root.winfo_screenwidth() // 2) - (largura // 2)
    y = (root.winfo_screenheight() // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{x}+{y}")

root = tk.Tk()
root.title("Gerador de Boletos")
root.configure(bg="#f0f0f0")

# Define a fonte
fonte = font.Font(family="Arial", size=12, weight="bold")

# Tamanho da janela
tamanho_janela = (400, 200)  # Largura x Altura
centralizar_janela(tamanho_janela)

# Configuração do texto de orientação
texto_orientacao = tk.Label(root, text="Clique aqui para gerar todos os boletos:", font=fonte, bg="#f0f0f0")
texto_orientacao.pack(pady=10)

# Configuração do botão
botao_gerar = tk.Button(root, text="Gerar Boletos", command=gerar_boletos, font=fonte, bg="#4CAF50", fg="white")
botao_gerar.pack(pady=10)

root.mainloop() #loop