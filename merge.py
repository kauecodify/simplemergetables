# -*- coding: utf-8 -*-
"""

Created on Sun Nov 30 23:23:56 2025

@author: k

run in spyder...

"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE

df_global = None

# ---------- FUNÇÕES ----------

def carregar_arquivo():
    global df_global
    caminho = filedialog.askopenfilename(
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
    )
    
    if not caminho:
        return
    
    try:
        if caminho.endswith(".csv"):
            df_global = pd.read_csv(caminho)
        else:
            df_global = pd.read_excel(caminho)
        
        lbl_status["text"] = "Arquivo carregado com sucesso!"
        atualizar_lista_colunas()

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{e}")


def atualizar_lista_colunas():
    listbox_colunas.delete(0, tk.END)
    for col in df_global.columns:
        listbox_colunas.insert(tk.END, col)


def mesclar_colunas():
    if df_global is None:
        messagebox.showerror("Erro", "Nenhum arquivo carregado.")
        return

    indices = listbox_colunas.curselection()
    if not indices:
        messagebox.showerror("Erro", "Selecione ao menos duas colunas.")
        return

    colunas = [df_global.columns[i] for i in indices]
    nome_coluna = entry_nome.get().strip()
    
    if not nome_coluna:
        messagebox.showerror("Erro", "Digite um nome para a nova coluna.")
        return

    separador = entry_sep.get()

    df_global[nome_coluna] = df_global[colunas].astype(str).agg(separador.join, axis=1)

    salvar_arquivo()


def salvar_arquivo():
    caminho = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")]
    )
    
    if not caminho:
        return
    
    try:
        df_global.to_excel(caminho, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar arquivo:\n{e}")

# ---------- INTERFACE ----------

root = tk.Tk()
root.title("Merge de Colunas")
root.geometry("600x500")
root.configure(bg="#000000")

fg_green = "#00FF00"
font_padrao = ("Consolas", 12)

btn_carregar = tk.Button(
    root, text="Carregar Arquivo", command=carregar_arquivo,
    bg="#003300", fg=fg_green, font=font_padrao, width=20
)
btn_carregar.pack(pady=10)

label1 = tk.Label(root, text="Selecione colunas para mesclar:", bg="#000000", fg=fg_green, font=font_padrao)
label1.pack()

listbox_colunas = Listbox(
    root, selectmode=MULTIPLE, bg="#001100", fg=fg_green,
    font=font_padrao, width=40, height=10
)
listbox_colunas.pack(pady=10)

label2 = tk.Label(root, text="Nome da nova coluna:", bg="#000000", fg=fg_green, font=font_padrao)
label2.pack()

entry_nome = tk.Entry(root, bg="#001100", fg=fg_green, font=font_padrao, width=30)
entry_nome.pack(pady=5)

label3 = tk.Label(root, text="Separador (ex: espaço, vírgula, hífen):", bg="#000000", fg=fg_green, font=font_padrao)
label3.pack()

entry_sep = tk.Entry(root, bg="#001100", fg=fg_green, font=font_padrao, width=30)
entry_sep.insert(0, " ")
entry_sep.pack(pady=5)

btn_mesclar = tk.Button(
    root, text="Mesclar Colunas", command=mesclar_colunas,
    bg="#003300", fg=fg_green, font=font_padrao, width=20
)
btn_mesclar.pack(pady=20)

lbl_status = tk.Label(root, text="", bg="#000000", fg=fg_green, font=font_padrao)
lbl_status.pack()

root.mainloop()
