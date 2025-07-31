# Automação Franquiatech

Automação Python para extração de XMLs, envio ao Google Drive e compartilhamento por e-mail.

## Estrutura

- `main.py`: ponto de entrada do sistema
- `automacao/`: módulos de ação (login, baixar, enviar, etc.)
- `utils/`: funções de apoio (excel, logger, navegador)
- `config/`: planilhas e variáveis
- `imagens/`: imagens usadas pelo PyAutoGUI

## Execução

### Local

```bash
python main.py

## Docker

bash
Copiar
Editar
docker build -t automacao-franquia .
docker run --rm automacao-franquia

Projeto de automação com PyAutoGUI e Google Drive.

Testando integração de testes com GitHub Actions.