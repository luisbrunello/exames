# Painel de exames - GitHub + Render

Aplicação web simples para:
- receber um PDF de exames laboratoriais
- receber uma planilha Excel no modelo do painel
- localizar a linha `DATA`
- adicionar os valores na próxima coluna disponível
- manter os dados anteriores intactos
- devolver um ZIP com a planilha atualizada e um relatório JSON

## Rodar localmente

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Depois abra no navegador:

```text
http://127.0.0.1:5000
```

## Deploy no Render

1. Suba estes arquivos para um repositório no GitHub.
2. No Render, crie um novo Web Service a partir do repositório.
3. O `render.yaml` já define os comandos de build e start.
4. Após o deploy, abra a URL do serviço.

## Observações

- O app usa a linha `DATA` na coluna B para decidir onde escrever a nova data.
- A nova coluna é a próxima coluna livre depois da última data já preenchida.
- O app tenta copiar a formatação da coluna anterior.
- Se a data não for encontrada no PDF, o sistema usa a data atual.
- A extração ainda é baseada em regras/regex. Se aparecer um laboratório com layout muito diferente, será preciso adicionar novos padrões.
