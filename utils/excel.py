import pandas as pd
import os
import sys

def caminho_recurso(relativo):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relativo)
    return os.path.join(os.path.abspath("."), relativo)

def ler_parametros(path='config/Parametros.xlsx'):
    try:
        df = pd.read_excel(caminho_recurso(path), header=None)
        parametros = {df.iloc[i, 0]: df.iloc[i, 1] for i in range(len(df))}
        return parametros
    except Exception as e:
        print(f"Erro ao ler Parametros.xlsx: {e}")
        return {}

def ler_empresas(path='config/Empresas.xlsx'):
    try:
        df = pd.read_excel(caminho_recurso(path), dtype=str)
        return df.to_dict(orient='records')
    except Exception as e:
        print(f"Erro ao ler Empresas.xlsx: {e}")
        return []

def carregar_status(path='status_empresas.xlsx'):
    full_path = caminho_recurso(path)
    if os.path.exists(full_path):
        return pd.read_excel(full_path)
    else:
        df = pd.DataFrame(columns=["NOME EMPRESA", "STATUS"])
        df.to_excel(full_path, index=False)
        return df

def salvar_status(df, path='status_empresas.xlsx'):
    df.to_excel(caminho_recurso(path), index=False)

def atualizar_status(nome_empresa, status, path='status_empresas.xlsx'):
    df = carregar_status(path)
    if nome_empresa in df["NOME EMPRESA"].values:
        df.loc[df["NOME EMPRESA"] == nome_empresa, "STATUS"] = status
    else:
        df.loc[len(df)] = [nome_empresa, status]
    salvar_status(df, path)
