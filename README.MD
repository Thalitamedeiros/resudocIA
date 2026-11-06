# Resumidor de documentos com IA

Site em Flask: você envia um documento (TXT, MD, PDF ou DOCX) e a IA devolve um resumo, pontos-chave e tópicos.

## Como rodar

```bash
# 1. criar e ativar o ambiente virtual
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 2. instalar as dependências
pip install -r requirements.txt

# 3. definir a chave da API (https://console.anthropic.com)
export ANTHROPIC_API_KEY="sua-chave"   # Windows (PowerShell): $env:ANTHROPIC_API_KEY="sua-chave"

# 4. iniciar
python app.py
```

Abra http://127.0.0.1:5000 no navegador.

Para trocar o modelo: `export ANTHROPIC_MODEL="nome-do-modelo"`.