import os
import smtplib
import re
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request


# Escopo mais amplo e persistente
SCOPES = ['https://www.googleapis.com/auth/drive']

def autenticar_drive():
    creds = None
    token_path = 'credenciais/token.json'
    from utils.excel import caminho_recurso  # ou defina localmente se neces
    cred_path = caminho_recurso('credenciais/credentials.json')

    # Garante que a pasta existe
    os.makedirs('credenciais', exist_ok=True)

    # Usa o token salvo, se existir
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    # Se não tiver token ou for inválido, faz nova autenticação
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_path, SCOPES)
            creds = flow.run_local_server(port=0, open_browser=True)  # ou False, se preferir terminal
        # Salva token
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())

    # Constrói o serviço do Drive
    return build('drive', 'v3', credentials=creds)

def criar_pasta_drive(drive_service, nome_pasta, parent_id=None):
    query = f"name='{nome_pasta}' and mimeType='application/vnd.google-apps.folder'"
    if parent_id:
        query += f" and '{parent_id}' in parents"

    resposta = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    arquivos = resposta.get('files', [])

    if arquivos:
        return arquivos[0]['id']

    pasta_metadata = {
        'name': nome_pasta,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent_id:
        pasta_metadata['parents'] = [parent_id]

    pasta = drive_service.files().create(body=pasta_metadata, fields='id').execute()
    return pasta.get('id')

def enviar_arquivos(drive_service, pasta_empresa_local, nome_empresa):
    # Cria (ou encontra) a pasta da empresa no Drive
    pasta_empresa_id = criar_pasta_drive(drive_service, nome_empresa)

    # Percorre todas as subpastas e arquivos da empresa localmente
    for root, dirs, files in os.walk(pasta_empresa_local):
        caminho_relativo = os.path.relpath(root, pasta_empresa_local)
        
        # Cria subpasta no Drive (caso não seja a raiz)
        parent_id = pasta_empresa_id
        if caminho_relativo != ".":
            parent_id = criar_pasta_drive(drive_service, caminho_relativo, parent_id=pasta_empresa_id)

        # Envia todos os arquivos da pasta atual
        for nome_arquivo in files:
            caminho_arquivo = os.path.join(root, nome_arquivo)
            media = MediaFileUpload(caminho_arquivo, resumable=True)
            drive_service.files().create(
                body={'name': nome_arquivo, 'parents': [parent_id]},
                media_body=media,
                fields='id'
            ).execute()
            print(f"📁 Enviado: {caminho_relativo}/{nome_arquivo}")

    return gerar_link_compartilhamento(drive_service, pasta_empresa_id)

            
def gerar_link_compartilhamento(drive_service, folder_id):
    permission = {
        'type': 'anyone',
        'role': 'reader'
    }
    drive_service.permissions().create(
        fileId=folder_id,
        body=permission,
        fields='id'
    ).execute()
    
    return f"https://drive.google.com/drive/folders/{folder_id}"

def enviar_email_com_link(destinatario, link, nome_empresa):
    import re  # garante que está importado
    remetente = "auto.macaomercadao1@gmail.com"
    senha = "awgy evjr xlxi rzpd"  # senha de app

    corpo = f"Olá!\n\nSegue o link com os arquivos da empresa **{nome_empresa}**:\n{link}\n\nAtt,\nAutomação Mercadão"
    msg = MIMEText(corpo)
    msg['Subject'] = "Arquivos enviados para o Drive"
    msg['From'] = remetente

    # Suporte a múltiplos e-mails separados por vírgula ou ponto e vírgula
    destinatarios = [email.strip() for email in re.split(',|;', destinatario) if email.strip()]
    msg['To'] = ', '.join(destinatarios)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(remetente, senha)
            smtp.sendmail(remetente, destinatarios, msg.as_string())
            print("E-mail enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
