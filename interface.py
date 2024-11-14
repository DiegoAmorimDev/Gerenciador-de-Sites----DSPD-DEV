import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# Funções de manipulação de dados (carregar, salvar, adicionar, editar)

def atualizar_tabela(tabela, df):
    """Atualiza a tabela Tkinter com os dados do DataFrame."""
    for i in tabela.get_children():
        tabela.delete(i)
    for idx, row in df.iterrows():
        tabela.insert("", "end", values=list(row))

def limpar_campos(entries):
    """Limpa os campos de entrada."""
    for entry in entries:
        entry.delete(0, tk.END)

def adicionar_dado(entries, df, tabela):
    """Adiciona um novo registro ao DataFrame e atualiza a tabela."""
    dados = [entry.get() for entry in entries]
    if all(dados[:-1]):  # Verifica se todos os campos obrigatórios estão preenchidos
        novo_dado = pd.DataFrame([dados], columns=df.columns)
        df = pd.concat([df, novo_dado], ignore_index=True)
        atualizar_tabela(tabela, df)
        limpar_campos(entries)
    else:
        messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!")
    return df

def editar_dado(entries, df, tabela):
    """Edita o registro selecionado na tabela."""
    selected_item = tabela.selection()
    if not selected_item:
        messagebox.showwarning("Aviso", "Selecione um registro para editar.")
        return df
    
    item_values = tabela.item(selected_item)["values"]
    index = tabela.index(selected_item)

    # Preenche os campos com valores selecionados para edição
    for entry, value in zip(entries, item_values):
        entry.delete(0, tk.END)
        entry.insert(0, value)
    
    def confirmar_edicao():
        novo_dado = [entry.get() for entry in entries]
        df.iloc[index] = novo_dado
        atualizar_tabela(tabela, df)
        limpar_campos(entries)
        messagebox.showinfo("Sucesso", "Registro editado com sucesso.")
        janela_edicao.destroy()

    janela_edicao = tk.Toplevel()
    janela_edicao.title("Confirmar Edição")
    tk.Label(janela_edicao, text="Tem certeza de que deseja salvar a edição?").pack()
    tk.Button(janela_edicao, text="Confirmar", command=confirmar_edicao).pack()
    
    return df

#interface-----

def criar_interface(df, salvar_dados):
    
    root = tk.Tk()
    root.title("Gerenciador de Sites UEPA - TESTE")

    
    tabela = ttk.Treeview(root, columns=list(df.columns), show="headings")
    for col in tabela["columns"]:
        tabela.heading(col, text=col)
    tabela.pack()

    
    frame_inputs = tk.Frame(root)
    frame_inputs.pack()
    labels = ["Nome do Site", "URL", "Padronizado", "Manutenção", "Servidor", "Comentários"]
    entries = []

    for i, label_text in enumerate(labels):
        tk.Label(frame_inputs, text=label_text).grid(row=i, column=0)
        entry = tk.Entry(frame_inputs)
        entry.grid(row=i, column=1)
        entries.append(entry)

    
    tk.Button(root, text="Adicionar", command=lambda: adicionar_dado(entries, df, tabela)).pack()
    tk.Button(root, text="Editar Selecionado", command=lambda: editar_dado(entries, df, tabela)).pack()
    tk.Button(root, text="Salvar", command=lambda: salvar_dados(df)).pack()

   
    atualizar_tabela(tabela, df)

    return root
