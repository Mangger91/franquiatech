import pyautogui
import time
import os
from pyautogui import ImageNotFoundException
import pygetwindow as gw
import glob
import shutil

pyautogui.PAUSE = 1

def focar_chrome():
    for window in gw.getWindowsWithTitle("Google Chrome"):
        window.activate()
        window.maximize()
        time.sleep(1)
        break

def consultando_dados():
    print("Consultando Dados da Empresa...")
    while True:
        try:
            consultando = pyautogui.locateCenterOnScreen("imagens/consultandoDados.png", confidence=0.8, grayscale=True)
            if consultando is None:
                print("Dados carregados. Prosseguindo...")
                break
            else:
                print("Aguardando dados da empresa...")
        except ImageNotFoundException:
            print("dados carregados (imagem desapareceu). Prosseguindo...")
            break
        except Exception as e:
            print(f"⚠️ Erro inesperado ao verificar imagem: {e}")
            break
        time.sleep(1)

def selecionar_empresa(codigo, filial):
    pyautogui.click(pyautogui.locateCenterOnScreen("imagens/campoEmpresa.png", confidence=0.5))
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('del')
    pyautogui.write(str(codigo))
    pyautogui.press('tab')
    pyautogui.press('del')
    pyautogui.write(str(filial))
    pyautogui.press('tab')
    pyautogui.click(pyautogui.locateCenterOnScreen("imagens/btnSelecionar.png", confidence=0.9))
    consultando_dados()

def aguardar_fim_consulta(timeout_segundos=120):
    timeout = time.time() + timeout_segundos
    print("Buscando XML...")
    while True:
        try:
            carregando = pyautogui.locateCenterOnScreen("imagens/carregando.png", confidence=0.8, grayscale=True)
            if carregando is None:
                print("Notas carregadas. Prosseguindo...")
                break
            else:
                print("Aguardando processamento das notas...")
        except ImageNotFoundException:
            print("Notas carregadas (imagem desapareceu). Prosseguindo...")
            break
        except Exception as e:
            print(f"⚠️ Erro inesperado ao verificar imagem: {e}")
            break
        if time.time() > timeout:
            print("⛔ Tempo limite excedido ao buscar XML.")
            break
        time.sleep(1)

def baixar_xml(pasta_destino, data_inicial, data_final):
    try:
        pyautogui.click(pyautogui.locateCenterOnScreen("imagens/btnMenu.png", confidence=0.9))

        pos_fiscal = pyautogui.locateCenterOnScreen("imagens/opcaoFiscal.png", confidence=0.9)
        if pos_fiscal:
            pyautogui.moveTo(pos_fiscal)
        else:
            raise Exception("Imagem 'opcaoFiscal.png' não localizada.")

        pos_area = pyautogui.locateCenterOnScreen("imagens/areaContador.png", confidence=0.9)
        if pos_area:
            pyautogui.moveTo(pos_area)
        else:
            raise Exception("Imagem 'areaContador.png' não localizada.")

        pyautogui.click(pyautogui.locateCenterOnScreen("imagens/exportarXML.png", confidence=0.9))
    except Exception as e:
        print(f"⛔ Erro ao navegar até exportação de XML: {e}")
        return

    time.sleep(2)

    try:
        campo_data_inicio = pyautogui.locateCenterOnScreen("imagens/dataInicio.png", confidence=0.7)
        if campo_data_inicio:
            pyautogui.click(campo_data_inicio)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('del')
            pyautogui.write(data_inicial)
            pyautogui.press('tab')
            pyautogui.press('del')
            pyautogui.write(data_final)
        else:
            print("⛔ Campo de data não localizado. Encerrando empresa.")
            return
    except Exception as e:
        print(f"⛔ Erro ao preencher datas: {e}")
        return

    try:
        btn_pesquisar = pyautogui.locateCenterOnScreen("imagens/btnPesquisar.png", confidence=0.9)
        if btn_pesquisar:
            pyautogui.click(btn_pesquisar)
        else:
            print("⛔ Botão 'Pesquisar' não encontrado.")
            return
    except Exception as e:
        print(f"⛔ Erro ao clicar em pesquisar: {e}")
        return

    print("Buscando XML...")
    time.sleep(5)
    aguardar_fim_consulta()

    try:
        pyautogui.click(pyautogui.locateCenterOnScreen("imagens/salvarXML.png", confidence=0.9))
        time.sleep(2)
        pyautogui.write(pasta_destino)
        pyautogui.press('tab')
        pyautogui.press('enter')
        print(f"XML salvo em: {pasta_destino}")
        time.sleep(5)
    except Exception as e:
        print(f"⛔ Erro ao salvar XML: {e}")

def renomear_ultimo_zip(destino, nome_empresa):
    zips = glob.glob(os.path.join(destino, '*.zip'))
    if not zips:
        print("⛔ Nenhum arquivo .zip encontrado para renomear.")
        return None
    zip_mais_recente = max(zips, key=os.path.getctime)
    nome_limpo = nome_empresa.replace(" ", "_").replace("/", "-")
    novo_caminho = os.path.join(destino, f"{nome_limpo}.zip")
    try:
        shutil.move(zip_mais_recente, novo_caminho)
        print(f"✅ Zip renomeado para: {novo_caminho}")
        return novo_caminho
    except Exception as e:
        print(f"⛔ Erro ao renomear zip: {e}")
        return None

def gerar_relatorio(pasta_destino):
    print("Preparando relatório...")
    time.sleep(10)
    try:
        pyautogui.click(pyautogui.locateCenterOnScreen("imagens/relatorioNotas.png", confidence=0.9))
        pyautogui.press('tab')
        pyautogui.press('enter')
    except Exception as e:
        print(f"⛔ Erro ao abrir relatório: {e}")
        return
    gerando_relatorio()
    if focar_visualizador_pdf():
        pyautogui.hotkey('ctrl', 's')
        time.sleep(2)
        pyautogui.write(os.path.join(pasta_destino, "relatorio.pdf"))
        pyautogui.press('enter')
        time.sleep(10)
        pyautogui.hotkey('alt', 'f4')
        print("Relatório salvo.")
    else:
        print("⛔ Não foi possível focar o visualizador de PDF. Relatório não salvo.")

def gerando_relatorio(timeout_segundos=120):
    timeout = time.time() + timeout_segundos
    print("Gerando relatório...")
    while True:
        try:
            carregando = pyautogui.locateCenterOnScreen("imagens/gerandoRelatorio.png", confidence=0.8, grayscale=True)
            if carregando is None:
                print("Relatório gerado.")
                break
            else:
                print("Aguardando o relatório ser gerado...")
        except ImageNotFoundException:
            print("Relatório gerado. Prosseguindo...")
            time.sleep(15)
            break
        except Exception as e:
            print(f"⚠️ Erro inesperado ao verificar imagem: {e}")
            break
        if time.time() > timeout:
            print("⛔ Tempo limite excedido ao gerar relatório.")
            break
        time.sleep(1)

def focar_visualizador_pdf():
    possiveis_titulos = ["Google Chrome", "chrome", "Adobe Acrobat", "Adobe Reader", "Visualizador", "relatório", "report"]
    for titulo in possiveis_titulos:
        janelas = gw.getWindowsWithTitle(titulo)
        if janelas:
            janela = janelas[0]
            janela.activate()
            janela.maximize()
            time.sleep(1)
            print(f"🔎 Foco definido para: {janela.title}")
            return True
    print("⚠️ Nenhuma janela de visualizador PDF encontrada.")
    return False
