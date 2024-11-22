import pandas as pd
from openpyxl import load_workbook
import customtkinter as ctk
from tkinter import ttk, messagebox

manutencao_opcoes = ["Sim", "Não"]
padronizado_opcoes = ["Padronizado", "Não padronizado", "Com erro", "Diferente", "Desenvolvimento"]
servidor_opcoes = ["59", "115"]


def carregar_dados():
    try:
        with pd.ExcelWriter("AMBIENTE DE TESTES.xlsx", engine="openpyxl", mode="a") as writer:
            if "Testes" not in writer.book.sheetnames:
                pd.DataFrame(columns=["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"]).to_excel(writer, sheet_name="Testes", index=False)
        df = pd.read_excel("AMBIENTE DE TESTES.xlsx", sheet_name="Testes")
        return df
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo Excel não encontrado.")
        return pd.DataFrame(columns=["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"])


def salvar_dados(df):
    with pd.ExcelWriter("AMBIENTE DE TESTES.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df = df[["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"]]
        df.to_excel(writer, sheet_name="Testes", index=False)
    messagebox.showinfo("Sucesso", "Dados salvos no arquivo Excel.")


def adicionar_dado():
    nome = entry_nome.get()
    url = entry_url.get()
    padronizado = entry_padronizado.get()
    manutencao = entry_manutencao.get()
    servidor = entry_servidor.get()
    comentarios = entry_comentarios.get()
    
    if nome and url and padronizado and manutencao and servidor:
        novo_dado = pd.DataFrame([[nome, url, padronizado, manutencao, servidor, comentarios]], 
                                 columns=["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"])
        global df
        df = pd.concat([df, novo_dado], ignore_index=True)
        atualizar_tabela()
        limpar_campos()
    else:
        messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!")


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
    


def selecionar_linha(event):
    selected_item = tabela.selection()
    if selected_item:
        item_values = tabela.item(selected_item)["values"]
        entry_nome.delete(0, ctk.END)
        entry_nome.insert(0, item_values[0])
        entry_url.delete(0, ctk.END)
        entry_url.insert(0, item_values[1])
        entry_padronizado.set(item_values[2])
        entry_manutencao.set(item_values[3])
        entry_servidor.set(item_values[4])
        entry_comentarios.delete(0, ctk.END)
        entry_comentarios.insert(0, item_values[5])

# Configuração da interface 
root = ctk.CTk(fg_color="#262C40")
root.title("Gerenciador de Sites - AMBIENTE DE TESTES")
root.state("zoomed")  
root.iconbitmap("LOGO-UEPA.ico")

df = carregar_dados()

search_frame = ctk.CTkFrame(root, fg_color="#8093F1")
search_frame.pack(pady=30)

search_var = ctk.StringVar()
search_label = ctk.CTkLabel(search_frame, text="Buscar:", text_color="white")
search_label.pack(side="left", padx=10)
search_entry = ctk.CTkEntry(search_frame, textvariable=search_var, width=200)
search_entry.pack(side="left")
search_entry.bind("<KeyRelease>", lambda event: filtrar_tabela(search_var, df, tabela))

search_entry.configure(
    fg_color="#C3CDEA",  # Cor de fundo do combo box
    text_color="black",  # Cor do texto
    font=("Arial", 12),  # Fonte do texto
    border_width=2,  # Largura da borda
    border_color="#3550E9"  # Cor da borda
)


style = ttk.Style()
style.theme_use("default")
style.configure("Treeview", background="#A3AFF5", foreground="black", fieldbackground="#A3AFF5", font=("Arial", 12)) 
style.configure("Treeview.Heading", background="#8093F1", foreground="white", font=("Arial", 14, "bold"))
style.map("Treeview", background=[("selected", "#5A70ED")], foreground=[("selected", "black")])
style.map("Treeview.Heading", background=[("active", "#8093F1")], foreground=[("active", "white")])

tabela = ttk.Treeview(root, columns=["NOME DO SITE", "URL", "STATUS", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"], show="headings", height=15)


tabela.pack(side="top", fill="none", expand=False)


for col in tabela["columns"]:
    tabela.heading(col, text=col)

atualizar_tabela()


tabela.bind("<ButtonRelease-1>", selecionar_linha)





frame_campos = ctk.CTkFrame(root, fg_color="#8093F1")
frame_campos.pack(side="top", fill="none", expand=False, pady=20)

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

ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Adicionar", command=adicionar_dado).grid(row=0, column=0, padx=10)
ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Editar", command=editar_dado).grid(row=0, column=1, padx=10)
ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Salvar", command=lambda: salvar_dados(df)).grid(row=0, column=2, padx=10)
ctk.CTkButton(frame_botoes, fg_color="#8093F1", hover_color="#5A70ED", text="Limpar campos", command=lambda: limpar_campos()).grid(row=0, column=3, padx=10)

root.mainloop()