import pandas as pd
import tkinter as tk
from tkinter import filedialog

def carregar_arquivo():
    global df
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            info_label.config(text="Arquivo carregado com sucesso!")
            print(df.head())  # Exibir as primeiras linhas dos dados
        except FileNotFoundError:
            info_label.config(text="Arquivo não encontrado. Verifique o caminho e o nome do arquivo.")
        except Exception as e:
            info_label.config(text=f"Erro ao carregar arquivo: {e}")

def verifica_codigos(row):
    colunas = ['FO 0180', 'FO 90270', 'FO 0', 'FO 90', 'FO 180', 'FO 270', 'FO AD']
    for coluna in colunas:
        valor = str(row[coluna])
        if valor:
            return valor.split()
    return []

def gerar_scripts_coordenadas():
    global df
    if 'Leste' in df.columns and 'Norte' in df.columns:
        try:
            # Geração dos scripts para coordenadas
            scripts = []
            for index, row in df.iterrows():
                leste = row['Leste']
                norte = row['Norte']

                # Verificação sequencial das colunas para encontrar o primeiro valor não vazio
                valores = []
                colunas = ['FO 0180', 'FO 90270', 'FO 0', 'FO 90', 'FO 180', 'FO 270', 'FO AD']
                valor_encontrado = False  # Flag para identificar se foi encontrado algum valor não vazio
                valores_por_coluna = {}  # Armazenar valores por coluna

                for coluna in colunas:
                    valor = str(row[coluna])
                    if valor.strip():
                        valores_por_coluna[coluna] = valor.split()  # Adiciona os valores identificados na coluna
                        valor_encontrado = True  # Marca que um valor foi encontrado

                # Se valores forem encontrados, considerar a coluna com mais códigos
                if valor_encontrado:
                    coluna_mais_codigos = max(valores_por_coluna, key=lambda k: len(valores_por_coluna[k]))
                    valores = valores_por_coluna[coluna_mais_codigos]

                # Formar o script de coordenadas
                if valores:
                    posicoes = [f"{i}ª_POSIÇÃO_{codigo}" for i, codigo in enumerate(valores, start=1)]
                    script = f"MTEXT {leste},{norte} 1.6 {leste} {norte} {' '.join(posicoes)}\n\n"
                    scripts.append(script)
                else:
                    # Se nenhum valor for encontrado, criar o script sem a última parte
                    script = f"MTEXT {leste},{norte} 1.6 {leste} {norte}\n\n"
                    scripts.append(script)

            # Selecionar o local para salvar o arquivo .scr
            file_to_save = filedialog.asksaveasfile(defaultextension=".scr")
            if file_to_save:
                # Escrever os scripts gerados no arquivo
                for script in scripts:
                    file_to_save.write(script)
                file_to_save.close()
                info_label.config(text="Scripts para COORDENADAS gerados e salvos com sucesso!")
        except Exception as e:
            info_label.config(text=f"Erro ao gerar/salvar scripts para COORDENADAS: {e}")
    else:
        info_label.config(text="Colunas 'Leste' e 'Norte' necessárias para gerar COORDENADAS não encontradas.")

def gerar_scripts_postes():
    global df
    if 'Leste' in df.columns and 'Norte' in df.columns:
        try:
            # Geração dos scripts para retângulos
            scripts = []
            for index, row in df.iterrows():
                leste = row['Leste']
                norte = row['Norte']
                script = f"RECTANGLE {leste},{norte} D 2.11 3.65 0.0"
                scripts.append(script)

            # Selecionar o local para salvar o arquivo .scr
            file_to_save = filedialog.asksaveasfile(defaultextension=".scr")
            if file_to_save:
                # Escrever os scripts gerados no arquivo
                for script in scripts:
                    file_to_save.write(script + '\n')
                file_to_save.close()
                info_label.config(
                    text="Scripts para POSTES gerados e salvos com sucesso!")
        except Exception as e:
            info_label.config(
                text=f"Erro ao gerar/salvar scripts para POSTES: {e}")
    else:
        info_label.config(text="Colunas 'Leste' e 'Norte' não encontradas.")


def gerar_scripts_estruturas():
    global df
    if 'Leste' in df.columns and 'Norte' in df.columns and 'ESTRUTURA' in df.columns:
        try:
            # Geração dos scripts para estruturas
            scripts = []
            for index, row in df.iterrows():
                leste = row['Leste']
                norte = row['Norte']
                estrutura = row['ESTRUTURA']
                script = f"MTEXT {leste},{norte} 1.6 DT_{estrutura}\n\n"
                scripts.append(script)

            # Selecionar o local para salvar o arquivo .scr
            file_to_save = filedialog.asksaveasfile(defaultextension=".scr")
            if file_to_save:
                # Escrever os scripts gerados no arquivo
                for script in scripts:
                    file_to_save.write(script)
                file_to_save.close()
                info_label.config(
                    text="Scripts para ESTRUTURAS gerados e salvos com sucesso!")
        except Exception as e:
            info_label.config(
                text=f"Erro ao gerar/salvar scripts para ESTRUTURAS: {e}")
    else:
        info_label.config(
            text="Colunas 'Leste', 'Norte' ou 'ESTRUTURA' não encontradas.")


def gerar_scripts_circulos():
    global df
    if 'Leste' in df.columns and 'Norte' in df.columns:
        try:
            # Geração dos scripts para círculos
            scripts = []
            for index, row in df.iterrows():
                leste = row['Leste']
                norte = row['Norte']
                script = f"CIRCLE {leste},{norte} 2.1"
                scripts.append(script)

            # Selecionar o local para salvar o arquivo .scr
            file_to_save = filedialog.asksaveasfile(defaultextension=".scr")
            if file_to_save:
                # Escrever os scripts gerados no arquivo
                for script in scripts:
                    file_to_save.write(script + '\n')
                file_to_save.close()
                info_label.config(
                    text="Scripts para CÍRCULOS gerados e salvos com sucesso!")
        except Exception as e:
            info_label.config(
                text=f"Erro ao gerar/salvar scripts para CÍRCULOS: {e}")
    else:
        info_label.config(text="Colunas 'Leste' e 'Norte' não encontradas.")


def gerar_scripts_mtext():
    global df
    if 'Leste' in df.columns and 'Norte' in df.columns and 'POSTE' in df.columns and 'POSIÇÃO' in df.columns:
        try:
            # Geração dos scripts para MTEXT
            scripts = []
            for index, row in df.iterrows():
                leste = row['Leste']
                norte = row['Norte']
                poste = row['POSTE'] if 'POSTE' in df.columns else ''
                posicao = row['POSIÇÃO'] if 'POSIÇÃO' in df.columns else ''
                script = f"MTEXT {leste},{norte} 1.6 DT_{poste} {posicao}\n\n"
                scripts.append(script)

            # Selecionar o local para salvar o arquivo .scr
            file_to_save = filedialog.asksaveasfile(defaultextension=".scr")
            if file_to_save:
                # Escrever os scripts gerados no arquivo
                for script in scripts:
                    file_to_save.write(script)
                file_to_save.close()
                info_label.config(
                    text="Scripts para MTEXT gerados e salvos com sucesso!")
        except Exception as e:
            info_label.config(
                text=f"Erro ao gerar/salvar scripts para MTEXT: {e}")
    else:
        info_label.config(
            text="Colunas 'Leste', 'Norte', 'POSTE' ou 'POSIÇÃO' não encontradas.")


def gerar_scripts_numeracao_levantamento():
    global df
    if 'Leste' in df.columns and 'Norte' in df.columns and 'Caracte' in df.columns:
        try:
            # Geração dos scripts para numeração de levantamento
            scripts = []
            for index, row in df.iterrows():
                leste = row['Leste']
                norte = row['Norte']
                caracte = row['Caracte']
                script = f"TEXT {leste},{norte} 1.6 0 {caracte}"
                scripts.append(script)

            # Selecionar o local para salvar o arquivo .scr
            file_to_save = filedialog.asksaveasfile(defaultextension=".scr")
            if file_to_save:
                # Escrever os scripts gerados no arquivo
                file_to_save.write('\n'.join(scripts))
                file_to_save.close()
                info_label.config(
                    text="Scripts para NUMERAÇÃO LEVANTAMENTO gerados e salvos com sucesso!")
        except Exception as e:
            info_label.config(
                text=f"Erro ao gerar/salvar scripts para NUMERAÇÃO LEVANTAMENTO: {e}")
    else:
        info_label.config(
            text="Colunas 'Leste', 'Norte' ou 'Caracte' não encontradas.")


# Configuração da interface gráfica
root = tk.Tk()
root.title("Gerador de Scripts AutoCAD")

# Botão para carregar o arquivo Excel
load_button = tk.Button(
    root, text="Carregar Arquivo Excel", command=carregar_arquivo)
load_button.pack()

# Botão para gerar scripts para POSTES
generate_postes_button = tk.Button(
    root, text="POSTES", command=gerar_scripts_postes)
generate_postes_button.pack()

# Botão para gerar scripts para ESTRUTURAS
generate_estruturas_button = tk.Button(
    root, text="ESTRUTURAS", command=gerar_scripts_estruturas)
generate_estruturas_button.pack()

# Botão para gerar scripts para CÍRCULOS
generate_circulos_button = tk.Button(
    root, text="CÍRCULOS", command=gerar_scripts_circulos)
generate_circulos_button.pack()

# Botão para gerar scripts para MTEXT
generate_mtext_button = tk.Button(
    root, text="MTEXT", command=gerar_scripts_mtext)
generate_mtext_button.pack()

# Botão para gerar scripts para NUMERAÇÃO LEVANTAMENTO
generate_numeracao_button = tk.Button(
    root, text="NUMERAÇÃO LEVANTAMENTO", command=gerar_scripts_numeracao_levantamento)
generate_numeracao_button.pack()

# Botão para gerar scripts para COORDENADAS
generate_coordenadas_button = tk.Button(
    root, text="COORDENADAS", command=gerar_scripts_coordenadas)
generate_coordenadas_button.pack()

# Label para exibir informações
info_label = tk.Label(root, text="")
info_label.pack()

root.mainloop()