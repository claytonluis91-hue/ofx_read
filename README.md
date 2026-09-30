# Nexus • Conferência Financeira

Aplicação Streamlit para converter arquivos OFX em planilhas Excel de conferência e extratos PDF Bradesco Empresas em OFX.

## Recursos

- OFX → Excel com abas de lançamentos, resumo e instruções.
- Totais de créditos, débitos, resultado, valor conferido e total visível por filtro.
- Coluna **Conferido** com seleção Sim/Não e destaque automático.
- PDF → OFX com extração de dados bancários e lançamentos.
- Tabela editável antes da geração do OFX.
- Processamento local, sem envio dos arquivos a serviços externos.

## Executar no Windows

1. Instale o Python 3.11 ou superior.
2. Dê duplo clique em `iniciar_local.bat`.
3. Na primeira execução, aguarde a criação do ambiente virtual e instalação das dependências.

Para executar manualmente:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

## Testes

```powershell
python -m pytest -q
```

Os testes de integração usam os arquivos existentes na pasta local `Exemplos`. Essa pasta é ignorada pelo Git para impedir a publicação de dados financeiros reais; em ambientes sem esses arquivos, os respectivos testes são ignorados automaticamente.

## Escopo do PDF

O parser foi calibrado para o layout de extrato **Bradesco Empresas** presente em `Exemplos`. PDFs escaneados não possuem texto selecionável e precisam de OCR. Outros bancos ou mudanças relevantes de layout exigem um adaptador adicional.

## Deploy no Streamlit Community Cloud

1. Publique o projeto em um repositório GitHub.
2. No Streamlit Community Cloud, escolha o repositório e defina `app.py` como arquivo principal.
3. Não é necessário cadastrar segredos para esta versão.

> Antes de importar um OFX no sistema financeiro, valide quantidade de lançamentos, totais e saldo final contra o PDF original.
