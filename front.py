import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import functions
from PIL import Image, ImageTk

import pandas as pd

def selecionar_arquivo_1():
    global bm
    bm = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if bm:
        messagebox.showinfo("Seleção de Arquivo", "BM selecionado com sucesso!")

def selecionar_arquivo_2():
    global drake
    drake = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls")])
    if drake:
        messagebox.showinfo("Seleção de Arquivo", "Relatório Drake selecionado com sucesso!")

def comparar_arquivos():
    if bm and drake:
        try:
            lista_plat = ['PPG1', 'PCP-2', 'PCP-1/3', 'PVM1', 'PVM3', 'FSO']
            df = pd.read_excel(drake)
            df_final = functions.comparar(lista_plat, df, bm)

            treeview.delete(*treeview.get_children())

            # Adicionar as colunas no treeview
            treeview["columns"] = list(df_final.columns)
            for column in df_final.columns:
                treeview.column(column, width=950)
                treeview.heading(column, text=column)

            # Adicionar as linhas no treeview
            for index, row in df_final.iterrows():
                treeview.insert("", tk.END, values=list(row))

            # Ativar o botão para exportar o dataframe para um arquivo Excel
            export_button.config(state=tk.NORMAL)
            janela_comparar_bm.destroy()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    else:
        messagebox.showwarning("Comparar Arquivos", "Selecione os dois arquivos antes de compará-los!")

def abrir_tela_comparar_bm():
    global bm
    global drake
    global janela_comparar_bm
    bm = None
    drake = None

    largura = 400
    altura = 200

    # Criar a janela para comparar BM
    janela_comparar_bm = tk.Toplevel(root)
    janela_comparar_bm.title("Comparar BM")
    janela_comparar_bm.grab_set()

    # Obter a resolução do monitor
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()

    # Calcular as coordenadas para centralizar a janela na tela
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2

    # Definir a geometria da janela
    janela_comparar_bm.geometry(f"{largura}x{altura}+{x}+{y}")

    # Especifique o caminho para o arquivo de ícone (.ico)
    icone_path = "icone.ico"

    # Defina o ícone da janela principal
    janela_comparar_bm.iconbitmap(icone_path)

    # Criar um widget Label para exibir a mensagem
    obs_label = tk.Label(janela_comparar_bm, text="OBS.: No relatório BM a coluna 'Conferência Matriz' \n tem que estar na coluna 'AY' do excel para que funcione.")
    obs_label.pack()

    # Botões para selecionar arquivos
    btn_arquivo1 = tk.Button(janela_comparar_bm, text="Selecionar o BM", command=selecionar_arquivo_1)
    btn_arquivo1.pack(pady=10)

    btn_arquivo2 = tk.Button(janela_comparar_bm, text="Selecionar o relatório Drake", command=selecionar_arquivo_2)
    btn_arquivo2.pack(pady=10)

    # Botão para comparar os arquivos
    btn_comparar = tk.Button(janela_comparar_bm, text="Comparar", command=comparar_arquivos, state=tk.DISABLED)
    btn_comparar.pack(pady=10)

    # Verificar a seleção dos arquivos para habilitar o botão de comparação
    def verificar_selecao_arquivos():
        if bm and drake:
            btn_comparar.config(state=tk.NORMAL)
        else:
            btn_comparar.config(state=tk.DISABLED)
        janela_comparar_bm.after(200, verificar_selecao_arquivos)

    verificar_selecao_arquivos()

# Função para gerar o relatório
def botao_gerar():
    try:
        # Abrir uma janela para selecionar um arquivo Excel
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls")])

        # Verificar se o usuário selecionou um arquivo
        if file_path:
            # Ler o arquivo Excel e obter os dados em um dataframe
            df = pd.read_excel(file_path)
            df.rename(columns={'index': 'indice', 'Matrícula do Trabalhador': 'Matricula', 'Nome do Trabalhador': 'Nome',
                               'Uop do Trabalhador': 'PlatTrab',
                               'Situação do Trabalhador': 'Situacao', 'Função de Folha do Trabalhador': 'Funcao',
                               'Data de Início do Evento': 'Data_ini',
                               'Data de Término do Evento': 'Data_fin', 'Descrição do Evento': 'Evento',
                               'Uop do Evento': 'Uop', 'Quantidade de Dias no Período': 'Dias'}, inplace=True)
            df.query(
                "Uop == 'PPG1' or Uop == 'PCP-2' or Uop == 'PCP-1/3' or Uop == 'PVM1' or Uop == 'PVM3' or Uop == 'FSO'",
                inplace=True)
            df.query("Evento != 'FOLGA'", inplace=True)
            df.query("Evento != 'RESERVA'", inplace=True)
            df.query("Evento != 'FALTA'", inplace=True)
            df.query("Evento != 'Desligamento'", inplace=True)
            df.query("Evento != 'ABONO ÓBITO'", inplace=True)
            df.query("Evento != 'LICENÇA MÉDICA VENCIDA'", inplace=True)
            df.drop(['Situacao'], inplace=True, axis=1)
            df = df.reset_index(drop=True)

            lista_plat = ['PPG1', 'PCP-2', 'PCP-1/3', 'PVM1', 'PVM3', 'FSO']

            df_final = functions.gerar_relatorio(lista_plat, df)

            # Limpar os dados anteriores
            treeview.delete(*treeview.get_children())

            # Adicionar as colunas no treeview
            treeview["columns"] = list(df_final.columns)
            for column in df_final.columns:
                treeview.column(column, width=950)
                treeview.heading(column, text=column)

            # Adicionar as linhas no treeview
            for index, row in df_final.iterrows():
                treeview.insert("", tk.END, values=list(row))

            # Ativar o botão para exportar o dataframe para um arquivo Excel
            export_button.config(state=tk.NORMAL)

    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Função para exportar o dataframe para um arquivo Excel
def exportar_para_excel():
    # Obter todos os itens do treeview
    items = treeview.get_children()

    if len(items) > 0:
        # Criar um dicionário vazio para armazenar os dados
        data = {}

        # Obter as colunas do treeview
        columns = treeview["columns"]

        # Iterar sobre os itens do treeview
        for item in items:
            # Obter os valores de cada coluna para o item atual
            values = treeview.item(item, "values")

            # Adicionar os valores ao dicionário de dados
            for i in range(len(columns)):
                column = columns[i]
                value = values[i]
                if column not in data:
                    data[column] = []
                data[column].append(value)

        # Criar o dataframe a partir dos dados
        df = pd.DataFrame(data)

        # Abrir uma janela para salvar o arquivo Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

        # Verificar se o usuário selecionou um local para salvar o arquivo
        if file_path:
            # Salvar o dataframe como um arquivo Excel
            df.to_excel(file_path, index=False)
            tk.messagebox.showinfo("Exportar para Excel", "O arquivo foi exportado com sucesso!")
    else:
        tk.messagebox.showwarning("Exportar para Excel", "Não há dados para exportar!")

def encerrar_programa():
    root.destroy()

# Criar a janela principal
root = tk.Tk()
root.title("Gerador BM")

# Definir as dimensões da janela
largura = 800
altura = 400

# Criar um Frame como container
container = tk.Frame(root)
container.pack(fill=tk.BOTH, expand=True)

# Obter a resolução do monitor
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

# Calcular as coordenadas para centralizar a janela na tela
x = (largura_tela - largura) // 2
y = (altura_tela - altura) // 2

# Definir a geometria da janela
root.geometry(f"{largura}x{altura}+{x}+{y}")

# Botão para gerar o relatório
gerar_button = tk.Button(root, text="Gerar Relatório", command=botao_gerar)
gerar_button.pack()
gerar_button.place(x=75, y=270)

# Criar o treeview
treeview = ttk.Treeview(container)
treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
treeview.place(width= 765, height= 250)
treeview.place(x=10, y=10)

# Botão para exportar o dataframe para um arquivo Excel
export_button = tk.Button(root, text="Exportar para Excel", command=exportar_para_excel, state=tk.DISABLED)
export_button.pack()
export_button.place(x=610, y=270)

# Botão para abrir a tela de comparação de BM
comparar_button = tk.Button(root, text="Comparar BM", command=abrir_tela_comparar_bm)
comparar_button.pack(pady=20)
comparar_button.place(x=355, y=270)

# Criar um widget Label para exibir a mensagem
mensagem_label = tk.Label(root, text="Ao clicar, selecione o relatório gerado pelo Drake.")
mensagem_label.pack()

# Definir a posição da mensagem
mensagem_label.place(x=10, y=310)

scrollbar = ttk.Scrollbar(container, orient="vertical", command=treeview.yview)
treeview.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Configurar a função para encerrar o programa corretamente
root.protocol("WM_DELETE_WINDOW", encerrar_programa)

# Especifique o caminho para o arquivo de ícone (.ico)
icone_path = "icone.ico"

# Defina o ícone da janela principal
root.iconbitmap(icone_path)

# Carregar a imagem
image = Image.open("logo.png")
photo = image.resize((200, 70), Image.ANTIALIAS)

photo = ImageTk.PhotoImage(photo)
# Criar um widget Label para exibir a imagem
label = tk.Label(root, image=photo)
label.pack()
label.place(x=570, y=320)

# Iniciar a interface
root.mainloop()
