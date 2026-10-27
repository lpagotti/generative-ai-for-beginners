# Exportar clientes do MOGO Gourmet

Exporta o cadastro **Cadastros → Pessoas → Clientes** do MOGO Gourmet para uma planilha Excel.

## Como funciona

1. Abre o Chromium (Playwright) em modo headless e faz login com `MOGO_USER` / `MOGO_PASS`.
2. Escolhe o modo **Gerencial**. Se o usuário já estiver conectado em outro dispositivo, marca "Finalizar sessão ativa", porque o MOGO só permite uma sessão por usuário.
3. Lê a lista pela mesma chamada interna da tela (`POST /Cadastros/GenerioDadosJqGrid`), em páginas de 500.
4. Sai do sistema (`/Account/LogOff`) e grava `saida/clientes_<empresa>_<data>.xlsx`.

## Pré-requisitos

```bash
pip install playwright==1.63.0 openpyxl
# Em máquina sem Chromium do Playwright: python -m playwright install chromium
export MOGO_USER=...   # use um usuário dedicado: o login derruba outra sessão aberta dele
export MOGO_PASS=...
```

## Uso

```bash
python exportar_clientes.py                       # empresa cidinhabolos
python exportar_clientes.py --empresa outra --saida clientes.xlsx --por-pagina 500
```

Uma exportação completa (cerca de 15,6 mil clientes) leva uns 2 minutos.

## LGPD

A planilha contém dados pessoais (nome, celular, endereço, CPF). A pasta `saida/` e os arquivos `.xlsx`/`.csv` estão no `.gitignore`: **nunca versione a exportação**. Usar esses dados em campanhas exige base legal (por exemplo, legítimo interesse ou consentimento) e uma forma de o cliente pedir para sair da lista.
