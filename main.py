import pandas as pd
from openpyxl import load_workbook
from tkinter import messagebox  # Corrigido: Importação do messagebox
from interface import criar_interface

# Função para carregar dados do Excel
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

# Função para salvar dados no Excel
def salvar_dados(df):
    with pd.ExcelWriter("AMBIENTE DE TESTES.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df = df[["NOME DO SITE", "URL", "PADRONIZADO", "MANUTENÇÃO", "SERVIDOR", "COMENTÁRIOS"]]
        df.to_excel(writer, sheet_name="Testes", index=False)
    messagebox.showinfo("Sucesso", "Dados salvos no arquivo Excel.")

# Executa o programa
if __name__ == "__main__":
    df = carregar_dados()
    root = criar_interface(df, salvar_dados)
    root.mainloop()
