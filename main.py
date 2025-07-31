import sys
import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import time
import json

def caminho_recurso(relativo):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relativo)
    return os.path.join(os.path.abspath('.'), relativo)

import traceback

from automacao.login import realizar_login, abrir_sistema, reiniciar_sistema
from automacao.baixar_xml import selecionar_empresa, baixar_xml, gerar_relatorio, renomear_ultimo_zip
from automacao.enviar_drive import autenticar_drive, enviar_arquivos, enviar_email_com_link
from utils.excel import carregar_status, atualizar_status, ler_empresas

import os
from datetime import datetime
import pyautogui
import pandas as pd

is_running = True
is_paused = False

def carregar_dados_selecionados():
    with open(caminho_recurso("selecionadas.json"), "r", encoding="utf-8") as f:
        return json.load(f)

def executar_automacao(log_func):
    global is_running, is_paused

    try:
        dados = carregar_dados_selecionados()
        empresas_selecionadas = dados["empresas"]
        data_inicial = dados["data_inicial"]
        data_final = dados["data_final"]
        email_destino = dados["email"]

        todas_empresas = ler_empresas("config/Empresas.xlsx")
        empresas_filtradas = [e for e in todas_empresas if e['Nome_Empresa'] in empresas_selecionadas]

        status_df = carregar_status("status_empresas.xlsx")
        drive_service = autenticar_drive()

        abrir_sistema()
        log_func(f"Data: {data_inicial} até {data_final}")

        for empresa in empresas_filtradas:
            if not is_running:
                log_func("❌ Execução interrompida pelo usuário.")
                return

            while is_paused:
                time.sleep(0.5)

            nome = empresa['Nome_Empresa']
            codigo = empresa['Codigo']
            filial = empresa['Filial']
            login = empresa['Usuario']
            senha = empresa['Senha']

            log_func(f"\n▶️ Processando: {nome} [{codigo}-{filial}]")

            if nome in status_df["NOME EMPRESA"].values:
                status_existente = status_df.loc[status_df["NOME EMPRESA"] == nome, "STATUS"].values[0]
                if pd.notna(status_existente):
                    log_func(f"⏩ Empresa {nome} já processada. Pulando...")
                    continue

            try:
                realizar_login(login, senha)
                selecionar_empresa(codigo, filial)

                mes_ano = datetime.strptime(data_inicial, '%d/%m/%Y').strftime('%m-%y')
                pasta_destino = os.path.join(r"C:\Scripts_e_Automacoes\Automacao_Mercadao\XML", nome, mes_ano)
                os.makedirs(pasta_destino, exist_ok=True)

                baixar_xml(pasta_destino, data_inicial, data_final)
                renomear_ultimo_zip(pasta_destino, nome)
                gerar_relatorio(pasta_destino)

                link = enviar_arquivos(drive_service, os.path.join(r"C:\Scripts_e_Automacoes\Automacao_Mercadao\XML", nome), nome)
                log_func(f"📎 Link de acesso: {link}")

                atualizar_status(nome, "SUCESSO", "status_empresas.xlsx")
                enviar_email_com_link(email_destino, link, nome)

            except Exception as e:
                atualizar_status(nome, f"Erro: {str(e)}", "status_empresas.xlsx")
                log_func(f"❌ Erro ao processar {nome}: {str(e)}")
                log_func(traceback.format_exc())

            reiniciar_sistema()

        log_func("\n✅ Todas as empresas processadas com sucesso.")
        pyautogui.alert("Automação finalizada com sucesso!", title="Conclusão")

    except Exception as e:
        log_func(f"⚠️ Erro geral na automação: {str(e)}")
        log_func(traceback.format_exc())

def iniciar_interface_execucao():
    global is_running, is_paused
    is_running = True
    is_paused = False

    exec_window = tk.Tk()
    exec_window.title("Execução da Automação")

    log_text = scrolledtext.ScrolledText(exec_window, width=90, height=28)
    log_text.pack(padx=10, pady=10)

    def log(msg):
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)

    def pausar():
        global is_paused
        is_paused = True
        log("⏸️ Execução pausada pelo usuário.")

    def continuar():
        global is_paused
        if not is_paused:
            return
        log("🖱️ Por favor, selecione a tela do aplicativo do Mercadão.")
        for i in range(5, 0, -1):
            log(f"⏳ Retomando em {i} segundos...")
            exec_window.update()
            time.sleep(1)
        is_paused = False
        log("▶️ Execução retomada.\n")

    def finalizar():
        global is_running
        if messagebox.askyesno("Finalizar", "Deseja realmente finalizar a execução? A automação será interrompida após a empresa atual."):
            is_running = False
            log("⚠️ A automação será interrompida ao final da empresa atual.")


    botoes_frame = tk.Frame(exec_window)
    botoes_frame.pack(pady=10)

    tk.Button(botoes_frame, text="⏸️ Pausar", width=15, command=pausar).pack(side=tk.LEFT, padx=10)
    tk.Button(botoes_frame, text="▶️ Continuar", width=15, command=continuar).pack(side=tk.LEFT, padx=10)
    tk.Button(botoes_frame, text="❌ Finalizar", width=15, command=finalizar).pack(side=tk.LEFT, padx=10)

    thread = threading.Thread(target=executar_automacao, args=(log,))
    thread.start()

    exec_window.mainloop()

def main():
    iniciar_interface_execucao()
