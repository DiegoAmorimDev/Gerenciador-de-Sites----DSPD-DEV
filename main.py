
import os
import sys
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import customtkinter as ctk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk


manutencao_opcoes = ["Sim", "Não"]
padronizado_opcoes = ["Padronizado", "Não padronizado", "Com erro", "Diferente", "Desenvolvimento"]
servidor_opcoes = ["59", "115"]


# Configurações de autenticação e planilha
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SAMPLE_SPREADSHEET_ID = "1AerAPBJanhwTSyhC7b3H1Gx8WDV5Jb0uubFM5f42GXI"
SAMPLE_RANGE_NAME = "Dados_Gerais!A1:F1015"

# Funções para autenticação e acesso ao Google Sheets
def get_google_sheets_service():
    creds = None
    token_path = get_file_path('token.json')  # Certifique-se de que o caminho do token seja correto
    credentials_path = get_file_path('credentials.json')  # Certifique-se de que o caminho do arquivo de credenciais seja correto

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as token:
            token.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


def get_file_path(filename):
    # Verifica se o programa está sendo executado como executável
    if getattr(sys, 'frozen', False):
        # Caminho correto quando o script é empacotado
        base_path = sys._MEIPASS
    else:
        # Caminho correto durante o desenvolvimento
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

# Use o get_file_path para acessar os arquivos
credentials_path = get_file_path('credentials.json')
token_path = get_file_path('token.json')

if os.path.exists(token_path):
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    print(f"Token encontrado: {token_path}")
else:
    print("Token não encontrado.")


def carregar_dados():
    try:
        service = get_google_sheets_service()
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
        values = result.get("values", [])
        if not values:
            raise ValueError("Nenhum dado encontrado na planilha.")
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
    except (HttpError, ValueError) as e:
        messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
        return pd.DataFrame(columns=["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"])

def salvar_dados_edit(df):

    selected_item = tabela.selection()

    item_values = tabela.item(selected_item)["values"]
    index = df[df["NOME DO SITE"] == item_values[0]].index[0]
  
    valores_originais = item_values

    # Variável para controlar se a janela já está aberta
    if hasattr(salvar_dados_edit, "janela_edicao_aberta") and salvar_dados_edit.janela_edicao_aberta:
        return  # Se a janela já foi aberta, não faz nada


    def confirmar_edicao():
        janela_edicao.destroy()  
        df.loc[index, "NOME DO SITE"] = entry_nome.get()
        df.loc[index, "URL"] = entry_url.get()
        df.loc[index, "PADRONIZADO"] = entry_padronizado.get()
        df.loc[index, "MANUTENÇÃO"] = entry_manutencao.get()
        df.loc[index, "SERVIDOR"] = entry_servidor.get()
        df.loc[index, "COMENTÁRIOS"] = entry_comentarios.get()
        atualizar_tabela()
        limpar_campos()
        messagebox.showinfo("Sucesso", "Registro editado com sucesso.")

        

        try:
            service = get_google_sheets_service()
            sheet = service.spreadsheets()
            values = [df.columns.tolist()] + df.values.tolist()
            body = {"values": values}
            sheet.values().update(
                spreadsheetId=SAMPLE_SPREADSHEET_ID,
                range=SAMPLE_RANGE_NAME,
                valueInputOption="RAW",
                body=body
            ).execute()
            messagebox.showinfo("Sucesso", "Dados salvos na planilha do Google Sheets.")
        except HttpError as e:
            messagebox.showerror("Erro", f"Erro ao salvar dados: {str(e)}")

    # messagebox.showinfo("Erro", "Edite um registro para salvar!")

    def cancelar_edicao():
        
        entry_nome.delete(0, ctk.END)
        entry_nome.insert(0, valores_originais[0])
        entry_url.delete(0, ctk.END)
        entry_url.insert(0, valores_originais[1])
        entry_padronizado.set(valores_originais[2])
        entry_manutencao.set(valores_originais[3])
        entry_servidor.set(valores_originais[4])
        entry_comentarios.delete(0, ctk.END)
        entry_comentarios.insert(0, valores_originais[5])
        janela_edicao.destroy()

    janela_edicao = ctk.CTkToplevel(root)
    janela_edicao.title("Confirmar Edição")
    janela_edicao.transient(root) 
    janela_edicao.lift()  # Coloca a janela de confirmação no topo, sem bloquear a interação
    
    ctk.CTkLabel(janela_edicao, text="Tem certeza de que deseja salvar a edição?").pack(padx=10, pady=10)
    ctk.CTkButton(janela_edicao, text="Confirmar", command=confirmar_edicao).pack(padx=10, pady=10)
    ctk.CTkButton(janela_edicao, text="Cancelar", command=cancelar_edicao).pack(padx=10, pady=5)

    janela_edicao.resizable(False, False)

    # Marca que a janela está aberta
    salvar_dados_edit.janela_edicao_aberta = True


def adicionar_dado():
    nome = entry_nome.get()
    url = entry_url.get() or ""
    padronizado = entry_padronizado.get()
    manutencao = entry_manutencao.get()
    servidor = entry_servidor.get()
    comentarios = entry_comentarios.get() or ""
    print("Dados capturados:", nome, url, comentarios)
    
    if nome and url and padronizado and manutencao and servidor:
        novo_dado = pd.DataFrame([[nome, url, padronizado, manutencao, servidor, comentarios]], 
                                 columns=["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"])
        global df
        df = pd.concat([df, novo_dado], ignore_index=True)
        atualizar_tabela()
        limpar_campos()
    else:
        messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!")

    try:
        service = get_google_sheets_service()
        sheet = service.spreadsheets()
        values = [df.columns.tolist()] + df.values.tolist()
        body = {"values": values}
        sheet.values().update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=SAMPLE_RANGE_NAME,
            valueInputOption="RAW",
            body=body
        ).execute()
        messagebox.showinfo("Sucesso", "Dados salvos na planilha do Google Sheets.")
    except HttpError as e:
        messagebox.showerror("Erro", f"Erro ao salvar dados: {str(e)}")

    


def atualizar_tabela():
    for i in tabela.get_children():
        tabela.delete(i)
    for idx, row in df.iterrows():
        tabela.insert("", "end", values=list(row))


def atualizar_tabela_filtrada(df_filtrado):
    for i in tabela.get_children():
        tabela.delete(i)
    for idx, row in df_filtrado.iterrows():
        tabela.insert("", "end", values=list(row))


def limpar_campos():
    entry_nome.delete(0, ctk.END)
    entry_url.delete(0, ctk.END)
    entry_padronizado.set("")
    entry_manutencao.set("")
    entry_servidor.set("")
    entry_comentarios.delete(0, ctk.END)


def filtrar_tabela(search_var, df, tabela):
    texto_pesquisa = search_var.get().lower()
    df_filtrado = df[df.apply(lambda row: texto_pesquisa in row.astype(str).str.lower().to_string(), axis=1)]
    atualizar_tabela_filtrada(df_filtrado)

def editar_dado():
    selected_item = tabela.selection()
    if not selected_item:
        messagebox.showwarning("Aviso", "Selecione um registro para editar.")
        return
    
    item_values = tabela.item(selected_item)["values"]
    index = df[df["NOME DO SITE"] == item_values[0]].index[0] 
    
  
    valores_originais = item_values

    entry_nome.delete(0, ctk.END)
    entry_nome.insert(0, item_values[0])
    entry_url.delete(0, ctk.END)
    entry_url.insert(0, item_values[1])
    entry_padronizado.set(item_values[2])
    entry_manutencao.set(item_values[3])
    entry_servidor.set(item_values[4])
    entry_comentarios.delete(0, ctk.END)
    entry_comentarios.insert(0, item_values[5])
    
    

# Configuração do caminho correto para o ícone
def resource_path(relative_path):
    """ Retorna o caminho absoluto para recursos em uma pasta ou arquivo dentro do executável """
    if hasattr(sys, "_MEIPASS"):
        # Caso o código esteja sendo executado como um executável
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)





# Configuração da interface 
root = ctk.CTk(fg_color="#262C40")
root.title("WebManager - DSPD")
root.state("zoomed")


# Ajuste no caminho do ícone, usando a função resource_path para localizar corretamente
icon_path = resource_path("LOGO-UEPA.ico")
# Para Windows e sistemas que requerem um formato de imagem específico:
icon = Image.open(icon_path)  # Carrega a imagem
icon = icon.resize((32, 32))  # Ajuste o tamanho do ícone, se necessário
root.iconphoto(True, ImageTk.PhotoImage(icon))  # Usando iconphoto ao invés de iconbitmap



df = carregar_dados()

# Função para mostrar ou esconder o placeholder
def toggle_placeholder(event, entry, placeholder_text, is_focused):
    if is_focused:
        # Remover placeholder quando o campo ganha foco
        if entry.get() == placeholder_text:
            entry.delete(0, "end")
            entry.configure(fg_color="#C3CDEA", text_color="black")
    else:
        # Adicionar placeholder quando o campo perde foco e está vazio
        if entry.get() == "":
            entry.insert(0, placeholder_text)
            entry.configure(fg_color="#C3CDEA", text_color="gray")

placeholder = "Pesquisar sites"

search_frame = ctk.CTkFrame(root, fg_color="#8093F1")
search_frame.pack(pady=30)

search_var = ctk.StringVar()
search_entry = ctk.CTkEntry(search_frame, textvariable=search_var, width=400)
search_entry.pack(side="left")
search_entry.bind("<KeyRelease>", lambda event: filtrar_tabela(search_var, df, tabela))

search_entry.configure(
    fg_color="#C3CDEA",
    text_color="gray",  # Inicialmente cinza para o placeholder
    font=("Arial", 12),
    border_width=2,
    border_color="#3550E9"
)

search_entry.insert(0, placeholder)

search_entry.bind("<FocusIn>", lambda event: toggle_placeholder(event, search_entry, placeholder, True))
search_entry.bind("<FocusOut>", lambda event: toggle_placeholder(event, search_entry, placeholder, False))

# Ícone de lupa (usar PIL para carregar imagem)
icon_path = get_file_path("lupa.png")  # ou resource_path("lupa.png")
icon = Image.open(icon_path).resize((20, 20))  # Redimensiona conforme necessário
icon = ImageTk.PhotoImage(icon)

# Label para exibir o ícone
icon_label = ctk.CTkLabel(search_frame, image=icon, text="")
icon_label.pack(side="left", padx=10)


style = ttk.Style()
style.theme_use("default")
style.configure("Treeview", background="#262C40", foreground="white", fieldbackground="#262C40", font=("Arial", 12)) 
style.configure("Treeview.Heading", background="#8093F1", foreground="white", font=("Arial", 14, "bold"))
style.map("Treeview", background=[("selected", "#5A70ED")], foreground=[("selected", "black")])
style.map("Treeview.Heading", background=[("active", "#8093F1")], foreground=[("active", "white")])
style.configure("Treeview", rowheight=40)  # Ajusta a altura de todas as linhas

tabela = ttk.Treeview(root, columns=["NOME DO SITE", "URL", "STATUS", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"], show="headings", height=15)

# Configurar colunas com largura fixa
colunas = [
    ("NOME DO SITE", 200),
    ("URL", 400),
    ("STATUS", 150),
    ("MANUTENÇÃO", 200),
    ("SERVIDOR", 150),
    ("COMENTÁRIOS", 200),
]

for col, width in colunas:
    tabela.heading(col, text=col)  # Configurar cabeçalhos
    tabela.column(col, width=width, anchor="center", stretch=False, minwidth=width)  # Desabilita redimensionamento





# Função para bloquear o redimensionamento
def bloquear_redimensionamento(event):
    return "break"  # Impede o redimensionamento da coluna

# Bind de redimensionamento da coluna
tabela.bind("<Button-1>", bloquear_redimensionamento)  # Impede o redimensionamento de colunas

def selecionar_linha(event):
    item = tabela.selection()  
    if item:
        print(f"Selecionado: {tabela.item(item)['values']}")

tabela.bind("<Button-1>", selecionar_linha) 
tabela.pack(side="left", fill="both", expand=False, padx=(5, 0))


for col in tabela["columns"]:
    tabela.heading(col, text=col)

atualizar_tabela()








frame_campos = ctk.CTkFrame(root, fg_color="#8093F1")
frame_campos.pack(side="top", fill="none", expand=False, padx=20, pady=(150, 0))

ctk.CTkLabel(frame_campos, text="Nome do site:", text_color="white").grid(row=0, column=0, padx=10, pady=5, sticky="e")
entry_nome = ctk.CTkEntry(frame_campos)
entry_nome.grid(row=0, column=1, padx=10, pady=5)

entry_nome.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)

ctk.CTkLabel(frame_campos, text="URL:", text_color="white").grid(row=1, column=0, padx=10, pady=5, sticky="e")
entry_url = ctk.CTkEntry(frame_campos)
entry_url.grid(row=1, column=1, padx=10, pady=5)

entry_url.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)


ctk.CTkLabel(frame_campos, text="Status:", text_color="white").grid(row=2, column=0, padx=10, pady=5, sticky="e")
entry_padronizado = ctk.CTkComboBox(frame_campos, values=padronizado_opcoes, state="readonly")
entry_padronizado.grid(row=2, column=1, padx=10, pady=5)

entry_padronizado.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    button_color="#3550E9",  # Cor do botão do dropdown
    dropdown_fg_color="#8093F1",  # Cor de fundo do dropdown
    dropdown_text_color="white",  # Cor do texto dentro do dropdown
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)

ctk.CTkLabel(frame_campos, text="Manutenção:", text_color="white").grid(row=3, column=0, padx=10, pady=5, sticky="e")
entry_manutencao = ctk.CTkComboBox(frame_campos, values=manutencao_opcoes, state="readonly")
entry_manutencao.grid(row=3, column=1, padx=10, pady=5)

entry_manutencao.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    button_color="#3550E9",  # Cor do botão do dropdown
    dropdown_fg_color="#8093F1",  # Cor de fundo do dropdown
    dropdown_text_color="white",  # Cor do texto dentro do dropdown
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)

ctk.CTkLabel(frame_campos, text="Servidor:", text_color="white").grid(row=4, column=0, padx=10, pady=5, sticky="e")
entry_servidor = ctk.CTkComboBox(frame_campos, values=servidor_opcoes, state="readonly")
entry_servidor.grid(row=4, column=1, padx=10, pady=5)

entry_servidor.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    button_color="#3550E9",  # Cor do botão do dropdown
    dropdown_fg_color="#8093F1",  # Cor de fundo do dropdown
    dropdown_text_color="white",  # Cor do texto dentro do dropdown
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)

ctk.CTkLabel(frame_campos, text="Comentários:", text_color="white").grid(row=5, column=0, padx=10, pady=5, sticky="e")
entry_comentarios = ctk.CTkEntry(frame_campos)
entry_comentarios.grid(row=5, column=1, padx=10, pady=5)

entry_comentarios.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)


# Botões
frame_botoes = ctk.CTkFrame(root, fg_color="#262C40")
frame_botoes.pack(pady=10)

ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Adicionar", command=adicionar_dado).grid(row=0, column=0, pady=10)
ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Editar", command=editar_dado).grid(row=1, column=0, pady=10)
ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Salvar edição", command=lambda: salvar_dados_edit(df)).grid(row=2, column=0, pady=10)
ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Limpar campos", command=lambda: limpar_campos()).grid(row=3, column=0, pady=10)

root.mainloop()