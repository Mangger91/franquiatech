# Imagem base com Python
FROM python:3.10-slim

# Instala dependências do sistema necessárias para o PyAutoGUI
RUN apt-get update && apt-get install -y \
    scrot \
    python3-tk \
    python3-dev \
    libx11-dev \
    libxtst-dev \
    libpng-dev \
    libjpeg-dev \
    x11-utils \
    && apt-get clean

# Cria a pasta da aplicação
WORKDIR /app

# Copia todos os arquivos do projeto
COPY . .

# Instala dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

# Define a execução padrão
CMD ["python", "main.py"]
