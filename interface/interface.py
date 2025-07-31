import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd
import json

import sys, os

def caminho_recurso(relativo):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relativo)
    return os.path.join(os.path.abspath("."), relativo)

class SeletorEmpresas:
    def __init__(self, master, empresas):
        self.master = master
        self.master.title("Seleção de Empresas e Período")
        self.empresas = empresas
        self.checkbox_vars = {}

        # Campo de datas
        self.frame_datas = ttk.Frame(master)
        self.frame_datas.pack(pady=10)

        ttk.Label(self.frame_datas, text="Data Inicial:").grid(row=0, column=0, padx=5)
        self.data_inicial = DateEntry(self.frame_datas, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.data_inicial.grid(row=0, column=1, padx=5)

        ttk.Label(self.frame_datas, text="Data Final:").grid(row=0, column=2, padx=5)
        self.data_final = DateEntry(self.frame_datas, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.data_final.grid(row=0, column=3, padx=5)

        # Campo de e-mail
        ttk.Label(master, text="E-mail para envio:").pack(pady=5)
        self.email_var = tk.StringVar()
        ttk.Entry(master, textvariable=self.email_var, width=40).pack()

        # Filtro de busca
        self.filtro_var = tk.StringVar()
        self.filtro_var.trace("w", self.atualizar_lista_empresas)

        filtro_frame = ttk.Frame(master)
        filtro_frame.pack(pady=5)
        ttk.Label(filtro_frame, text="Filtrar empresas:").pack(side=tk.LEFT, padx=5)
        self.entry_filtro = ttk.Entry(filtro_frame, textvariable=self.filtro_var)
        self.entry_filtro.pack(side=tk.LEFT, padx=5)

        # Frame com scrollbar
        self.frame_lista = ttk.Frame(master)
        self.frame_lista.pack(padx=10, pady=5, fill="both", expand=True)

        canvas = tk.Canvas(self.frame_lista, height=400)
        scrollbar = ttk.Scrollbar(self.frame_lista, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Botões de seleção
        botoes_frame = ttk.Frame(master)
        botoes_frame.pack(pady=5)
        ttk.Button(botoes_frame, text="Marcar Todos", command=self.marcar_todos).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_frame, text="Desmarcar Todos", command=self.desmarcar_todos).pack(side=tk.LEFT, padx=5)

        # Botão confirmar
        ttk.Button(master, text="Confirmar", command=self.confirmar).pack(pady=10)

        self.atualizar_lista_empresas()

    def atualizar_lista_empresas(self, *args):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        texto_filtro = self.filtro_var.get().lower()
        for i, empresa in enumerate(self.empresas):
            nome = empresa['Nome_Empresa']
            if texto_filtro in nome.lower():
                var = self.checkbox_vars.get(nome, tk.BooleanVar())
                cb = ttk.Checkbutton(self.scroll_frame, text=nome, variable=var)
                cb.pack(anchor="w")
                self.checkbox_vars[nome] = var

    def marcar_todos(self):
        for var in self.checkbox_vars.values():
            var.set(True)

    def desmarcar_todos(self):
        for var in self.checkbox_vars.values():
            var.set(False)

    def confirmar(self):
        selecionadas = [nome for nome, var in self.checkbox_vars.items() if var.get()]
        dados = {
            "empresas": selecionadas,
            "data_inicial": self.data_inicial.get(),
            "data_final": self.data_final.get(),
            "email": self.email_var.get().strip()
        }
        with open(caminho_recurso("selecionadas.json"), "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)

        print("Empresas selecionadas:", selecionadas)
        print("Data Inicial:", self.data_inicial.get())
        print("Data Final:", self.data_final.get())
        print("E-mail:", self.email_var.get().strip())
        self.master.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    df_empresas = pd.read_excel(caminho_recurso("config/Empresas.xlsx"))
    empresas = df_empresas.to_dict(orient="records")
    app = SeletorEmpresas(root, empresas)
    root.mainloop()
    
    import main
    main.main()
