# Consolidador de Estoque

Cruza as vendas (NFC-e) com as notas de entrada (Consulta NF-e da SEFAZ), ajuda a validar
quais produtos são o mesmo e gera o inventário do Bloco H (H010) de cada ano.

**Usar no navegador:** abra a página do GitHub Pages deste repositório. O processamento roda
no próprio navegador (Python via Pyodide) — os arquivos não são enviados para nenhum servidor.

**Testar localmente:** nesta pasta, `python -m http.server` e abra http://localhost:8000
(abrir o index.html direto do disco não funciona).

**Linha de comando:** `pip install -r requirements.txt` e
`python consolidacao_estoque.py --data_dir "XML SAIDAS"` (`--help` para as opções).

Este repositório é público: não coloque aqui XMLs, planilhas ou o cadastro_produtos.csv.
