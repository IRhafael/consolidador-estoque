# para testar a validação manual --validacao_preenchida <arquivo>


import os
import glob
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule, DataBarRule
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.datavalidation import DataValidation
import logging
import argparse
import unicodedata
from datetime import date

# Configuração de log
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Cadastro de produtos: liga as CHAVES naturais de cada lado (venda: PDV:cProd; compra:
# EAN:gtin ou FORN:cnpj:código) a um PRODUTO_ID próprio. Vale para todos os anos — é o que
# faz 2022, 2023... reaproveitarem a validação de 2021. Esta é a ordem das colunas no CSV.
CADASTRO_COLUNAS = ['CHAVE', 'LADO', 'PRODUTO_ID', 'DESCRICAO', 'FATOR_CONV', 'ORIGEM', 'SCORE', 'DATA']
# Pares respondidos 'Não': nunca mais são sugeridos nem casados no automático.
REJEITADOS_COLUNAS = ['CHAVE_VENDA', 'CHAVE_COMPRA', 'DESCRICAO_VENDA', 'DESCRICAO_COMPRA', 'DATA']

# ─────────────────────────────────────────────────────────────────────────────
# TEMA VISUAL DOS RELATÓRIOS (cores, fontes e helpers de formatação Excel)
# ─────────────────────────────────────────────────────────────────────────────
COR_AZUL_ESCURO = "1F3B57"
COR_AZUL_MEDIO  = "2E5A87"
COR_TEAL        = "137D74"
COR_VERMELHO    = "C0392B"
COR_ROXO        = "8E7CC3"
COR_AMARELO     = "F1C40F"
COR_CINZA_BANDA = "EEF3F8"
COR_CINZA_CLARO = "F7F9FB"
COR_BRANCO      = "FFFFFF"
COR_TXT         = "1A1A1A"
FONTE_PADRAO    = "Arial"

FORMATOS_KARDEX = {c: '#,##0.###' for c in ('SALDO_INI_QTD', 'ENTRADAS_QTD', 'SAIDAS_QTD', 'SALDO_FIM_QTD')}
FORMATOS_KARDEX.update({c: 'R$ #,##0.00' for c in ('SALDO_INI_VLR', 'ENTRADAS_VLR', 'CMV_SAIDAS_VLR', 'SALDO_FIM_VLR')})
FORMATOS_KARDEX['CUSTO_MEDIO'] = 'R$ #,##0.0000'

_LADO_FINO = Side(style="thin", color="D5DCE4")
BORDA_FINA = Border(left=_LADO_FINO, right=_LADO_FINO, top=_LADO_FINO, bottom=_LADO_FINO)


def _fmt_num(n):
    return f"{n:,}".replace(",", ".")


def estilizar_cabecalho(ws, ncols, cor=COR_AZUL_ESCURO, linha=1, altura=28, congelar=True):
    fill = PatternFill("solid", fgColor=cor)
    fonte = Font(name=FONTE_PADRAO, size=10, bold=True, color=COR_BRANCO)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=linha, column=c)
        cell.fill = fill
        cell.font = fonte
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDA_FINA
    ws.row_dimensions[linha].height = altura
    if congelar:
        ws.freeze_panes = ws.cell(row=linha + 1, column=1)


def auto_largura(ws, ncols, header_row=1, max_row_scan=250, minw=9, maxw=48):
    for c in range(1, ncols + 1):
        letra = get_column_letter(c)
        maxlen = 0
        for r in range(header_row, min(ws.max_row, max_row_scan) + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                maxlen = max(maxlen, len(str(v)))
        ws.column_dimensions[letra].width = max(minw, min(maxw, maxlen + 3))


def banda_zebra(ws, ncols, first_data_row=2, last_row=None):
    last_row = last_row or ws.max_row
    if last_row < first_data_row:
        return
    rng = f"A{first_data_row}:{get_column_letter(ncols)}{last_row}"
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=["ISEVEN(ROW())"],
                         fill=PatternFill("solid", fgColor=COR_CINZA_BANDA)))


def corpo_fonte(ws, ncols, first_data_row=2, last_row=None, size=10):
    last_row = last_row or ws.max_row
    f = Font(name=FONTE_PADRAO, size=size, color=COR_TXT)
    for r in range(first_data_row, last_row + 1):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).font = f


def escrever_aba_formatada(wb, nome, df, cor_header, formatos=None, col_amarela=None,
                           databar_col=None, colorscale_col=None, larguras_fixas=None,
                           limitar_fonte=None):
    """Escreve um DataFrame numa aba já com cabeçalho colorido, filtro, zebra, fonte,
    larguras e formatos numéricos. Retorna (ws, mapa_colunas)."""
    ws = wb.create_sheet(nome)
    ws.sheet_view.showGridLines = False
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    ncols = df.shape[1]
    if ncols == 0:
        return ws, {}
    estilizar_cabecalho(ws, ncols, cor=cor_header)
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}1"
    cols = {c: i + 1 for i, c in enumerate(df.columns)}
    last = ws.max_row
    if formatos:
        for cname, fmt_ in formatos.items():
            if cname in cols:
                ci = cols[cname]
                for r in range(2, last + 1):
                    ws.cell(row=r, column=ci).number_format = fmt_
    if limitar_fonte is None and last > 5000:
        limitar_fonte = 1
    corpo_fonte(ws, ncols, last_row=(limitar_fonte or last))
    banda_zebra(ws, ncols, last_row=last)
    auto_largura(ws, ncols)
    if larguras_fixas:
        for cname, w in larguras_fixas.items():
            if cname in cols:
                ws.column_dimensions[get_column_letter(cols[cname])].width = w
    if col_amarela and col_amarela in cols:
        ci = cols[col_amarela]
        ws.cell(row=1, column=ci).fill = PatternFill("solid", fgColor=COR_AMARELO)
        ws.cell(row=1, column=ci).font = Font(name=FONTE_PADRAO, size=10, bold=True, color="4A3B00")
        for r in range(2, last + 1):
            ws.cell(row=r, column=ci).fill = PatternFill("solid", fgColor="FFFCEC")
    if databar_col and databar_col in cols and last >= 2:
        letra = get_column_letter(cols[databar_col])
        ws.conditional_formatting.add(
            f"{letra}2:{letra}{last}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100,
                        color=COR_TEAL, showValue=True))
    if colorscale_col and colorscale_col in cols and last >= 2:
        letra = get_column_letter(cols[colorscale_col])
        ws.conditional_formatting.add(
            f"{letra}2:{letra}{last}",
            ColorScaleRule(start_type="num", start_value=-50, start_color="F8696B",
                           mid_type="num", mid_value=0, mid_color="FFEB9C",
                           end_type="num", end_value=50, end_color="63BE7B"))
    return ws, cols

def _pintar_linhas_por_resposta(ws, cols, last_row, ncols_total):
    """Pinta a LINHA INTEIRA conforme a resposta digitada na coluna amarela:

        "Sim"  -> verde   (match aceito)
        "Não"  -> vermelho (descartado)
        vazio  -> sem cor (ainda não decidido)

    Usa formatação condicional do Excel com referência mista ($D2), então a cor
    aparece na hora em que a pessoa digita, sem rodar o script de novo. Aceita
    "Sim", "sim", "S", "Não", "nao", "N" — compara só a primeira letra em maiúscula.
    """
    col_resp = 'ACEITAR_MATCH (Sim/Não)'
    if last_row < 2 or col_resp not in cols or ncols_total < 1:
        return

    letra = get_column_letter(cols[col_resp])
    faixa = f"A2:{get_column_letter(ncols_total)}{last_row}"
    ref = f'UPPER(LEFT(TRIM(${letra}2),1))'

    # Empurra as regras que já existem (zebra, barra de similaridade, status) para
    # prioridades mais altas = menos importantes. Sem isso a listra zebrada empata
    # em priority=1 com o realce e o Excel decide por desempate — o verde sairia
    # só em metade das linhas.
    for faixa_existente in ws.conditional_formatting:
        for regra in faixa_existente.rules:
            if regra.priority is not None:
                regra.priority = regra.priority + 10

    # priority 1 e 2 e stopIfTrue=True: a cor da resposta tem de VENCER a listra
    # zebrada da tabela, que já foi registrada antes nesta mesma faixa. Sem isso a
    # linha "Sim" ficaria verde só nas linhas ímpares.
    regra_sim = FormulaRule(
        formula=[f'AND(${letra}2<>"",{ref}="S")'],
        fill=PatternFill("solid", fgColor="C9EBD0"),
        font=Font(name=FONTE_PADRAO, size=10, color="14532D"),
        stopIfTrue=True)
    regra_sim.priority = 1
    ws.conditional_formatting.add(faixa, regra_sim)

    regra_nao = FormulaRule(
        formula=[f'AND(${letra}2<>"",{ref}="N")'],
        fill=PatternFill("solid", fgColor="F8CBCB"),
        font=Font(name=FONTE_PADRAO, size=10, color="7F1D1D"),
        stopIfTrue=True)
    regra_nao.priority = 2
    ws.conditional_formatting.add(faixa, regra_nao)


def _validacao_lista_suspensa(ws, cols, last_row):
    """Põe uma lista suspensa Sim/Não na coluna de resposta. Evita 'sin', 'SIm ',
    'ok' e outras variações que o leitor de volta não reconheceria."""
    col_resp = 'ACEITAR_MATCH (Sim/Não)'
    if last_row < 2 or col_resp not in cols:
        return
    letra = get_column_letter(cols[col_resp])
    dv = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=True,
                        showDropDown=False, errorStyle="warning",
                        error="Responda apenas Sim ou Não.",
                        errorTitle="Resposta inválida",
                        prompt="Sim = é o mesmo produto. Não = sugestão errada.",
                        promptTitle="Validação do match")
    ws.add_data_validation(dv)
    dv.add(f"{letra}2:{letra}{last_row}")


def _adicionar_colunas_formula_validacao(ws, cols, last_row):
    """Acrescenta DUAS COLUNAS COM FÓRMULA DE EXCEL (recalculam sozinhas, ao vivo,
    sem precisar rodar o script de novo) na aba de validação manual:

      - SALDO_SE_RESPOSTA: qual seria o saldo desse item considerando a resposta
        atual da coluna 'ACEITAR_MATCH (Sim/Não)' (Sim -> ENTRADA_QTD_AJUST - SAIDA_QTD;
        Não/vazio -> -SAIDA_QTD, ou seja, continua só saída).
      - STATUS_AUTOMATICO: em texto e com cor de fundo:
          "OK"                 -> aceitou o match e o saldo resultante é positivo/zero
          "NEGATIVO"           -> aceitou o match mas ainda sobra saldo negativo
          "REVISAR PROPORCAO"  -> aceitou um match onde a venda é mais de 15x a
                                   compra casada (sinal forte de sugestão errada)
          "DESCARTADO"         -> marcado como "Não"
          "AGUARDANDO"         -> célula ainda vazia

    Isso dá feedback instantâneo pra quem está validando, direto no Excel, sem
    precisar esperar o script rodar de novo pra saber se a resposta faz sentido.
    """
    obrigatorias = ('SAIDA_QTD', 'ENTRADA_QTD_AJUST', 'ACEITAR_MATCH (Sim/Não)')
    if last_row < 2 or not all(c in cols for c in obrigatorias):
        return

    col_saida = get_column_letter(cols['SAIDA_QTD'])
    col_entrada = get_column_letter(cols['ENTRADA_QTD_AJUST'])
    col_resp = get_column_letter(cols['ACEITAR_MATCH (Sim/Não)'])

    idx_saldo = max(cols.values()) + 1
    idx_status = idx_saldo + 1
    let_saldo = get_column_letter(idx_saldo)
    let_status = get_column_letter(idx_status)

    for idx, nome in ((idx_saldo, 'SALDO_SE_RESPOSTA'), (idx_status, 'STATUS_AUTOMATICO')):
        cel = ws.cell(row=1, column=idx, value=nome)
        cel.font = Font(name=FONTE_PADRAO, size=10, bold=True, color=COR_BRANCO)
        cel.fill = PatternFill("solid", fgColor=COR_AZUL_MEDIO)
        cel.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(2, last_row + 1):
        resp = f'UPPER(TRIM({col_resp}{r}))'
        f_saldo = f'=IF(LEFT({resp},1)="S",{col_entrada}{r}-{col_saida}{r},-{col_saida}{r})'
        c_saldo = ws.cell(row=r, column=idx_saldo, value=f_saldo)
        c_saldo.number_format = '#,##0.###'
        c_saldo.alignment = Alignment(horizontal="center")

        f_status = (
            f'=IF({col_resp}{r}="","AGUARDANDO",'
            f'IF(LEFT({resp},1)="N","DESCARTADO",'
            f'IF(LEFT({resp},1)="S",'
            f'IF(AND({col_entrada}{r}>0,({col_saida}{r}/{col_entrada}{r})>15),'
            f'"REVISAR PROPORCAO",IF({let_saldo}{r}>=0,"OK","NEGATIVO")),'
            f'"?")))'
        )
        c_status = ws.cell(row=r, column=idx_status, value=f_status)
        c_status.alignment = Alignment(horizontal="center")
        c_status.font = Font(name=FONTE_PADRAO, size=10, bold=True)

    ws.column_dimensions[let_saldo].width = 17
    ws.column_dimensions[let_status].width = 20

    faixa = f"{let_status}2:{let_status}{last_row}"
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'{let_status}2="OK"'], fill=PatternFill("solid", fgColor="C6EFCE"),
        font=Font(color="006100")))
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'{let_status}2="NEGATIVO"'], fill=PatternFill("solid", fgColor="FFEB9C"),
        font=Font(color="9C6500")))
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'{let_status}2="REVISAR PROPORCAO"'], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'{let_status}2="DESCARTADO"'], fill=PatternFill("solid", fgColor="EEEEEE"),
        font=Font(color="808080")))


class EstoqueProcessor:
    def __init__(self, output_dir="outputs", cadastro_path="cadastro_produtos.csv"):
        """Inicializa o processador de estoque com os diretórios de saída."""
        self.output_dir = output_dir
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        self.nfce_saidas = pd.DataFrame()
        self.entradas_xls = pd.DataFrame()
        self.resumo_estoque = pd.DataFrame()
        self.validacao_produtos = pd.DataFrame()
        self.matches_suspeitos_pos_validacao = pd.DataFrame()
        self.entradas_descartadas = pd.DataFrame()

        # Freios dos casamentos automáticos (ajustáveis pela linha de comando — ver o final).
        self.limiar_ean = 0.30      # EAN: chave forte, tolera nome diferente
        self.limiar_marca = 0.55    # palavra-chave de marca em comum
        self.usar_match_similaridade = True
        self.limiar_auto_similaridade = 0.88  # aceite automático sem chave nenhuma
        self.margem_ambiguidade = 0.05        # distância mínima para o 2º colocado
        self.usar_trava_preco = True
        self.trava_preco_min = 0.5         # venda/custo abaixo disso -> aborta o match automático
        self.trava_preco_max = 5.0         # venda/custo acima disso  -> aborta o match automático
        self.usar_veto_volume = True
        # PALAVRA-CHAVE e SIMILARIDADE ALTA só são GRAVADOS no cadastro com score >= isto.
        # Abaixo disso o casamento vale nesta execução, mas vai para a planilha de validação
        # como AUTO A CONFIRMAR até alguém responder. DESCRICAO EXATA e EAN são sempre gravados.
        self.cadastro_auto_min_sim = 0.98
        # Kardex e base de entradas dentro da planilha de validação (abas grandes e lentas)
        self.exportar_abas_pesadas = True

        # ── CADASTRO DE PRODUTOS: memória entre execuções E ENTRE ANOS ──
        self.cadastro_path = cadastro_path
        self.rejeitados_path = os.path.splitext(cadastro_path)[0] + '_rejeitados.csv'
        self._recusas_auto = {}    # (xProd_NORM, DESCRICAO_NORM) -> motivo da recusa (vai p/ OBSERVACAO)
        self._prov = {}            # CHAVE -> PRODUTO_ID dos casamentos AUTO A CONFIRMAR (não gravados)
        self._prov_origem = {}     # CHAVE -> origem desses casamentos
        self._provisorios = []     # (chaves_venda, chaves_compra, tipo, score) desta execução
        self._desc_chave = {}      # CHAVE -> descrição principal (preenchido no catálogo)
        self._carregar_cadastro()

    @staticmethod
    def _listar_arquivos(pasta, extensoes, periodo=None):
        """Lista os arquivos de `pasta` E DE TODAS AS SUBPASTAS, em ordem. Assim os dados
        podem ficar organizados por ano (XML SAIDAS/2021, XML SAIDAS/2022, ...) e todos
        os anos entram juntos no kardex. Com `periodo` (ex: '01-2021' ou '2021'), só
        entram os arquivos com esse texto no nome."""
        achados = []
        for raiz, _, nomes in os.walk(pasta):
            for nome in nomes:
                # '~$...' é o arquivo temporário de uma planilha que está aberta no Excel
                if nome.startswith('~$') or not nome.lower().endswith(extensoes):
                    continue
                if periodo and periodo not in nome:
                    continue
                achados.append(os.path.join(raiz, nome))
        return sorted(achados)

    def extrair_xmls_nfce(self, zip_dir, periodo=None):
        """Extrai os itens das NFC-e (saídas) dos .zip — e de .xml soltos — da pasta e subpastas.

        A mesma nota em dois arquivos (XML baixado de novo, ou o lote de um ano trazendo
        notas do fim do ano anterior) conta uma vez só: chave de acesso + número do item.
        Notas com evento de cancelamento presente no lote são excluídas."""
        logging.info(f"Iniciando extração de XMLs de NFC-e em: {zip_dir}")
        all_data = []
        canceladas = set()

        # Expressão regular para remover namespaces do XML
        ns_pattern = re.compile(r'\{.*\}')
        NS = '{http://www.portalfiscal.inf.br/nfe}'

        def processar(xml_file):
            try:
                root = ET.parse(xml_file).getroot()
            except Exception:
                return  # Ignora erros de parse individuais para não poluir o log

            chNFe = dhEmi = tpEvento = ""
            for elem in root.iter():
                tag = ns_pattern.sub('', elem.tag)
                if tag == 'chNFe':
                    chNFe = elem.text
                elif tag == 'dhEmi':
                    dhEmi = elem.text
                elif tag == 'tpEvento':
                    tpEvento = elem.text

            if tpEvento == '110111':  # evento de cancelamento: só anota a chave
                if chNFe:
                    canceladas.add(chNFe)
                return

            if not chNFe:
                for infNFe in root.iter(f'{NS}infNFe'):
                    _id = infNFe.get('Id')
                    if _id and _id.startswith('NFe'):
                        chNFe = _id[3:]
                        break

            for det in root.iter(f'{NS}det'):
                prod = det.find(f'{NS}prod')
                if prod is None:
                    continue
                qCom = self._get_text(prod, 'qCom')
                vUnCom = self._get_text(prod, 'vUnCom')
                vProd = self._get_text(prod, 'vProd')
                all_data.append({
                    'Chave_NFe': chNFe or '', 'N_ITEM': det.get('nItem', ''), 'Data_Emissao': dhEmi,
                    'cProd': self._get_text(prod, 'cProd'),
                    'cEAN': self._get_text(prod, 'cEAN'),
                    'xProd': self._get_text(prod, 'xProd'),
                    'NCM': self._get_text(prod, 'NCM'),
                    'CFOP': self._get_text(prod, 'CFOP'),
                    'uCom': self._get_text(prod, 'uCom'),
                    'QTD_SAIDA': float(qCom) if qCom else 0.0,
                    'VLR_UNIT_SAIDA': float(vUnCom) if vUnCom else 0.0,
                    'VLR_TOTAL_SAIDA': float(vProd) if vProd else 0.0,
                })

        arquivos = self._listar_arquivos(zip_dir, ('.zip', '.xml'), periodo)
        if not arquivos:
            logging.warning(f"Nenhum ZIP/XML encontrado para o período '{periodo or 'TODOS'}' na pasta {zip_dir}")

        for arq in arquivos:
            if arq.lower().endswith('.xml'):
                processar(arq)
                continue
            logging.info(f"Processando arquivo ZIP: {os.path.relpath(arq, zip_dir)}")
            try:
                with zipfile.ZipFile(arq, 'r') as z:
                    for filename in z.namelist():
                        if filename.lower().endswith('.xml'):
                            with z.open(filename) as xml_file:
                                processar(xml_file)
            except Exception as e:
                logging.error(f"Erro ao abrir ZIP {arq}: {e}")

        df = pd.DataFrame(all_data)
        if not df.empty:
            dup = df.duplicated(['Chave_NFe', 'N_ITEM']) & (df['Chave_NFe'] != '')
            if dup.any():
                logging.warning(f"{int(dup.sum())} item(ns) de NFC-e repetido(s) em mais de um arquivo — contados uma vez só.")
                df = df[~dup]
            canc = df['Chave_NFe'].isin(canceladas)
            if canc.any():
                logging.warning(f"{int(canc.sum())} item(ns) de NFC-e CANCELADA(s) removido(s) das saídas.")
                df = df[~canc]
            df = df.reset_index(drop=True)
            # Período AAAA-MM direto da data de emissão (não depende do nome do arquivo)
            df['PERIODO'] = df['Data_Emissao'].astype(str).str.extract(r'^(\d{4}-\d{2})')[0].fillna('')
            por_ano = df['PERIODO'].str[:4].value_counts().sort_index().to_dict()
            logging.info(f"Itens de NFC-e por ano: {por_ano}")

        self.nfce_saidas = df
        logging.info(f"Total de {len(self.nfce_saidas)} registros extraídos das NFC-es (Saídas).")
        return self.nfce_saidas
        
    def _get_text(self, element, tag_name):
        """Helper para buscar texto de uma tag ignorando namespace."""
        for child in element.iter():
            if child.tag.endswith(f'}}{tag_name}') or child.tag == tag_name:
                return child.text
        return None

    def _detectar_linha_cabecalho(self, xls_file, max_scan=25):
        """Varre as primeiras linhas do arquivo procurando a linha real de cabeçalho.

        Relatórios exportados de sistemas do governo (ex: Consulta NF-e da SEFAZ)
        costumam ter várias linhas de identificação institucional (nome do órgão,
        usuário, data/hora, título do relatório) ANTES da tabela de produtos.
        Se lermos direto com header=0, essas linhas viram os nomes das colunas
        e a tabela real de produtos nunca é reconhecida (colunas 'unnamed: N'),
        fazendo o cruzamento de estoque falhar silenciosamente (0 matches).
        """
        palavras_chave = [
            'cód', 'codigo', 'código', 'descri', 'ean', 'gtin', 'qtd', 'quantidade',
            'valor', 'vlr', 'ncm', 'cfop', 'unid', 'cst', 'csosn'
        ]
        try:
            try:
                preview = pd.read_excel(xls_file, header=None, nrows=max_scan, engine='xlrd')
            except Exception:
                preview = pd.read_excel(xls_file, header=None, nrows=max_scan)
        except Exception as e:
            logging.warning(f"Não foi possível pré-visualizar {xls_file} para detectar cabeçalho: {e}")
            return 0

        melhor_linha, melhor_score = 0, 0
        for idx, row in preview.iterrows():
            celulas = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            if not celulas:
                continue
            score = sum(1 for kw in palavras_chave if any(kw in c for c in celulas))
            if score > melhor_score:
                melhor_score = score
                melhor_linha = idx

        if melhor_score < 3:
            logging.warning(
                f"Não foi possível identificar com confiança a linha de cabeçalho em "
                f"{os.path.basename(xls_file)} (melhor score={melhor_score}). Usando linha 0 (padrão)."
            )
            return 0

        logging.info(f"Cabeçalho de '{os.path.basename(xls_file)}' detectado na linha {melhor_linha} (score={melhor_score}).")
        return melhor_linha

    def ler_planilhas_entradas(self, xls_dir, periodo=None):
        """Lê e processa planilhas .xls de entradas (pasta e subpastas) e mapeia colunas."""
        logging.info(f"Iniciando leitura de planilhas de entrada em: {xls_dir}")
        all_entradas = []

        col_mapping = {
            'cancelada': 'CANCELADA',
            'situação': 'CANCELADA',
            'situacao': 'CANCELADA',
            'cód. prod.': 'COD_ITEM',
            'descrição prod.': 'DESCRICAO',
            'ean': 'EAN',
            'ncm': 'NCM',
            'cfop': 'CFOP',
            'und': 'UNID',
            'qtd com.': 'QTD_ENTRADA',
            'qtd trib.': 'QTD_TRIBUTAVEL',
            'vlr prod.': 'VLR_TOTAL_ENTRADA',
            'cst': 'CST',
            'código': 'COD_ITEM',
            'cód. produto': 'COD_ITEM',
            'código produto': 'COD_ITEM',
            'código de barras': 'EAN',
            'código de barras/ean': 'EAN',
            'cod. barras': 'EAN',
            'gtin': 'EAN',
            'descrição': 'DESCRICAO',
            'descrição do produto': 'DESCRICAO',
            'unidade': 'UNID',
            'unidade com.': 'UNID',
            'valor total': 'VLR_TOTAL_ENTRADA',
            'valor do produto': 'VLR_TOTAL_ENTRADA',
            'valor prod.': 'VLR_TOTAL_ENTRADA',
            'quantidade': 'QTD_ENTRADA',
            'quantidade com.': 'QTD_ENTRADA',
            'cst/csosn': 'CST',
            # Identificação do fornecedor e da nota. O CNPJ do emitente é o que torna o
            # código do produto único: fornecedores diferentes reusam a mesma numeração.
            'emitente': 'CNPJ_EMITENTE',
            'cnpj emitente': 'CNPJ_EMITENTE',
            'cpf/cnpj emitente': 'CNPJ_EMITENTE',
            'razao_emitente': 'RAZAO_EMITENTE',
            'chave': 'CHAVE_ACESSO',
            'chave de acesso': 'CHAVE_ACESSO',
            'item': 'N_ITEM',
            'dt. emissão': 'DATA_EMISSAO',
            'data emissão': 'DATA_EMISSAO',
            'data de emissão': 'DATA_EMISSAO',
        }

        arquivos_xls = self._listar_arquivos(xls_dir, ('.xls', '.xlsx'), periodo)

        if not arquivos_xls:
            logging.warning(f"Nenhum arquivo XLS encontrado para o período '{periodo or 'TODOS'}' na pasta {xls_dir}")

        for xls_file in arquivos_xls:
            logging.info(f"Lendo arquivo: {os.path.relpath(xls_file, xls_dir)}")
            try:
                header_row = self._detectar_linha_cabecalho(xls_file)
                try:
                    df = pd.read_excel(xls_file, engine='xlrd', header=header_row)
                except Exception:
                    df = pd.read_excel(xls_file, header=header_row)

                # Descarta linhas totalmente vazias (comuns em relatórios exportados)
                df = df.dropna(how='all')

                cols = [str(c).strip().lower() for c in df.columns]
                # No relatório da SEFAZ o cabeçalho EMITENTE ocupa 3 colunas mescladas
                # (CNPJ | razão social | UF): a 1ª é o CNPJ, a seguinte a razão social.
                if 'emitente' in cols:
                    i = cols.index('emitente')
                    if i + 1 < len(cols) and cols[i + 1].startswith('unnamed'):
                        cols[i + 1] = 'razao_emitente'
                df.columns = cols
                rename_dict = {col: col_mapping[col] for col in df.columns if col in col_mapping}
                df = df.rename(columns=rename_dict)

                required_cols = ['COD_ITEM', 'DESCRICAO', 'QTD_ENTRADA', 'VLR_TOTAL_ENTRADA', 'EAN']
                for col in required_cols:
                    if col not in df.columns:
                        df[col] = 0 if 'VLR' in col or 'QTD' in col else ''

                # Período (AAAA-MM) pela DATA DE EMISSÃO da nota; o nome do arquivo
                # (ex: 01-2021.xls -> 2021-01) só é usado quando a data não vier.
                # É o que permite montar o kardex mês a mês e rolar o saldo entre anos.
                per_nome = self._extrair_periodo_do_nome(os.path.basename(xls_file))
                if 'DATA_EMISSAO' in df.columns:
                    dt = pd.to_datetime(df['DATA_EMISSAO'], errors='coerce', dayfirst=True)
                    df['PERIODO'] = dt.dt.strftime('%Y-%m').fillna(per_nome)
                else:
                    df['PERIODO'] = per_nome

                all_entradas.append(df)
            except Exception as e:
                logging.error(f"Erro ao ler {xls_file}: {e}")

        if all_entradas:
            self.entradas_xls = pd.concat(all_entradas, ignore_index=True)

            # Remove linhas de "lixo" que sobraram de cabeçalhos/rodapés institucionais
            # (sem descrição de produto e sem quantidade), evitando poluir o cruzamento.
            desc = self.entradas_xls['DESCRICAO']
            desc_vazia = desc.isna() | desc.astype(str).str.strip().isin(['', 'nan', 'None', '<NA>'])
            qtd_zero = pd.to_numeric(self.entradas_xls['QTD_ENTRADA'], errors='coerce').fillna(0) == 0
            linhas_antes = len(self.entradas_xls)
            self.entradas_xls = self.entradas_xls[~(desc_vazia & qtd_zero)].reset_index(drop=True)
            if linhas_antes != len(self.entradas_xls):
                logging.info(f"Descartadas {linhas_antes - len(self.entradas_xls)} linhas sem descrição/quantidade (lixo de cabeçalho/rodapé).")

            self.entradas_xls['COD_ITEM'] = self.entradas_xls['COD_ITEM'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).replace(['nan', 'None', ''], 'N/A')
            self.entradas_xls['QTD_ENTRADA'] = pd.to_numeric(self.entradas_xls['QTD_ENTRADA'], errors='coerce').fillna(0)
            self.entradas_xls['VLR_TOTAL_ENTRADA'] = pd.to_numeric(self.entradas_xls['VLR_TOTAL_ENTRADA'], errors='coerce').fillna(0)
            self.entradas_xls['EAN'] = self.entradas_xls['EAN'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).replace(['nan', 'None', ''], 'N/A')
            self.entradas_xls['DESCRICAO_NORM'] = self.entradas_xls['DESCRICAO'].map(self._normalizar_desc)
            if 'PERIODO' not in self.entradas_xls.columns:
                self.entradas_xls['PERIODO'] = ''
            self.entradas_xls['PERIODO'] = self.entradas_xls['PERIODO'].fillna('').astype(str)

            # Só dígitos no CNPJ, na chave de acesso e no nº do item ('…882'' -> '…882')
            for col in ('CNPJ_EMITENTE', 'CHAVE_ACESSO', 'N_ITEM'):
                if col not in self.entradas_xls.columns:
                    self.entradas_xls[col] = ''
                self.entradas_xls[col] = (self.entradas_xls[col].astype(str).str.replace(r'\.0$', '', regex=True)
                                          .str.replace(r'\D', '', regex=True))
            if 'RAZAO_EMITENTE' not in self.entradas_xls.columns:
                self.entradas_xls['RAZAO_EMITENTE'] = ''
            self.entradas_xls['RAZAO_EMITENTE'] = (self.entradas_xls['RAZAO_EMITENTE'].astype(str).str.strip()
                                                   .replace(['nan', 'None'], ''))

            # A mesma nota em duas planilhas (ex: relatório baixado de novo, ou meses que se
            # sobrepõem na virada do ano) entra uma vez só: chave de acesso + nº do item.
            dup = (self.entradas_xls.duplicated(['CHAVE_ACESSO', 'N_ITEM'])
                   & (self.entradas_xls['CHAVE_ACESSO'] != ''))
            if dup.any():
                logging.warning(f"{int(dup.sum())} item(ns) de nota de entrada repetido(s) em mais de uma "
                                f"planilha — contados uma vez só.")
                self.entradas_xls = self.entradas_xls[~dup].reset_index(drop=True)

            self._filtrar_canceladas_e_cfop()

            por_ano = self.entradas_xls['PERIODO'].str[:4].value_counts().sort_index().to_dict()
            logging.info(f"Total de {len(self.entradas_xls)} registros de entrada consolidados. Por ano: {por_ano}")
        return self.entradas_xls

    # CFOPs de aquisição de mercadoria para revenda/industrialização e transferências.
    # O relatório da SEFAZ mostra o CFOP como o FORNECEDOR emitiu (saída dele = nossa
    # entrada), por isso a lista é de códigos 5xxx/6xxx/7xxx.
    _CFOP_COMPRA = {
        5101, 5102, 5103, 5104, 5105, 5106, 5109, 5110, 5111, 5112, 5113, 5114,
        5115, 5116, 5117, 5118, 5119, 5120, 5122, 5123,
        5401, 5402, 5403, 5405, 5451,
        5151, 5152, 5155, 5156,
        6101, 6102, 6103, 6104, 6105, 6106, 6107, 6108, 6109, 6110, 6111, 6112,
        6113, 6114, 6115, 6116, 6117, 6118, 6119, 6120, 6122, 6123,
        6401, 6402, 6403, 6404,
        6151, 6152, 6155, 6156,
        7101, 7102,
    }
    # Bonificação/doação/brinde: a mercadoria ENTRA de verdade (soma quantidade), mas
    # costuma vir com valor simbólico ou zerado, o que derruba o custo médio. Entra,
    # porém sinalizada.
    _CFOP_BONIFICACAO = {5910, 5911, 6910, 6911}
    # Não são compra: remessa, conserto, comodato, demonstração, industrialização por
    # conta de terceiro, devolução. Se entrarem como estoque, inflam o inventário.
    _CFOP_NAO_COMPRA = {
        5901, 5902, 5903, 5904, 5905, 5906, 5907, 5908, 5909, 5912, 5913, 5914,
        5915, 5916, 5917, 5918, 5919, 5920, 5921, 5922, 5923, 5924, 5925, 5926,
        5551, 5552, 5553, 5554, 5555, 5556, 5557,
        5201, 5202, 5208, 5209, 5210, 5410, 5411, 5412, 5413, 5503, 5553,
        6901, 6902, 6903, 6904, 6905, 6906, 6907, 6908, 6909, 6912, 6913, 6914,
        6915, 6916, 6917, 6918, 6919, 6920, 6921, 6922, 6923, 6924, 6925,
        6551, 6552, 6553, 6554, 6555, 6556, 6557,
        6201, 6202, 6208, 6209, 6210, 6410, 6411, 6412, 6413,
    }

    def _filtrar_canceladas_e_cfop(self):
        """Tira do estoque o que não é compra de mercadoria.

        Dois furos que existiam antes:
          1. Notas CANCELADAS entravam como compra normal (a coluna existe no
             relatório da SEFAZ mas nunca era lida).
          2. Todo CFOP era tratado como aquisição. Remessa para conserto, comodato,
             demonstração e devolução inflavam o inventário; bonificação com valor
             zerado derrubava o custo médio sem aviso.

        O que não é compra vai para self.entradas_descartadas (vira aba de auditoria)
        em vez de sumir calado."""
        if self.entradas_xls.empty:
            return

        descartes = []
        antes = len(self.entradas_xls)

        # 1. Notas canceladas
        if 'CANCELADA' in self.entradas_xls.columns:
            flag = (self.entradas_xls['CANCELADA'].astype(str).str.strip().str.upper()
                    .isin(['S', 'SIM', 'TRUE', '1', 'CANCELADA', 'X']))
            if flag.any():
                d = self.entradas_xls[flag].copy()
                d['MOTIVO_DESCARTE'] = 'Nota cancelada'
                descartes.append(d)
                self.entradas_xls = self.entradas_xls[~flag]
                logging.warning(f"{int(flag.sum())} linha(s) de NOTA CANCELADA removidas das entradas.")

        # 2. CFOP
        if 'CFOP' in self.entradas_xls.columns:
            cfop = pd.to_numeric(self.entradas_xls['CFOP'], errors='coerce')
            self.entradas_xls['CFOP_NUM'] = cfop

            nao_compra = cfop.isin(self._CFOP_NAO_COMPRA)
            if nao_compra.any():
                d = self.entradas_xls[nao_compra].copy()
                d['MOTIVO_DESCARTE'] = 'CFOP não é aquisição de mercadoria (remessa/devolução/conserto)'
                descartes.append(d)
                self.entradas_xls = self.entradas_xls[~nao_compra]
                cfop = cfop[~nao_compra]
                resumo = d['CFOP'].value_counts().head(6).to_dict()
                logging.warning(f"{len(d)} linha(s) removidas por CFOP de não-compra: {resumo}")

            bonif = cfop.isin(self._CFOP_BONIFICACAO)
            self.entradas_xls['BONIFICACAO'] = bonif.reindex(self.entradas_xls.index).fillna(False)
            if bonif.any():
                zeradas = int((bonif & (pd.to_numeric(self.entradas_xls['VLR_TOTAL_ENTRADA'],
                                                      errors='coerce').fillna(0) <= 0)).sum())
                logging.warning(f"{int(bonif.sum())} linha(s) de BONIFICAÇÃO/BRINDE mantidas no estoque "
                                f"({zeradas} sem valor). Elas somam quantidade e puxam o custo médio "
                                f"para baixo — confira se é isso que a contabilidade quer.")

            conhecidos = self._CFOP_COMPRA | self._CFOP_BONIFICACAO
            desconhecido = (~cfop.isin(conhecidos)) & cfop.notna()
            desconhecido = desconhecido.reindex(self.entradas_xls.index).fillna(False)
            if desconhecido.any():
                lista = sorted(self.entradas_xls.loc[desconhecido, 'CFOP'].dropna().unique().tolist())
                logging.warning(f"{int(desconhecido.sum())} linha(s) com CFOP fora da lista conhecida "
                                f"{lista[:10]} foram MANTIDAS como compra. Confira se alguma dessas "
                                f"operações não deveria entrar no estoque.")
            self.entradas_xls['CFOP_REVISAR'] = desconhecido

        self.entradas_xls = self.entradas_xls.reset_index(drop=True)
        self.entradas_descartadas = (pd.concat(descartes, ignore_index=True)
                                     if descartes else pd.DataFrame())
        if antes != len(self.entradas_xls):
            logging.info(f"Entradas após filtro de cancelamento/CFOP: {len(self.entradas_xls)} "
                         f"(eram {antes}).")

    @staticmethod
    def _extrair_periodo_do_nome(nome):
        """Extrai o período AAAA-MM do nome de um arquivo. Reconhece MM-AAAA, AAAA-MM,
        MM_AAAA, MM.AAAA etc. Ex: '01-2021.xls' -> '2021-01'. Retorna '' se não achar."""
        base = str(nome)
        # MM-AAAA (ex: 01-2021)
        m = re.search(r'(?<!\d)(0?[1-9]|1[0-2])[\-_./ ](\d{4})(?!\d)', base)
        if m:
            return f"{int(m.group(2)):04d}-{int(m.group(1)):02d}"
        # AAAA-MM (ex: 2021-01)
        m = re.search(r'(?<!\d)(\d{4})[\-_./ ](0?[1-9]|1[0-2])(?!\d)', base)
        if m:
            return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
        return ''

    # ------------------------------------------------------------------
    # CONVERSÃO DE EMBALAGENS
    # ------------------------------------------------------------------
    # Overrides manuais por palavra-chave: quando a descrição da NOTA DE ENTRADA
    # não traz o número de unidades da embalagem (ex: só "FLOCAO FLOCUSCUZ 500G"
    # comprado em fardo), informe aqui a palavra-chave e o fator. O override só é
    # aplicado quando NENHUM padrão numérico foi encontrado na descrição (fator=1),
    # para não multiplicar duas vezes.
    OVERRIDES_EMBALAGEM = {
        'FLOCUSCUZ': 20,  # flocão de milho comprado em fardo c/ 20 un (conferência)
    }

    # Unidades comerciais que indicam embalagem múltipla (se não acharmos o número
    # de unidades na descrição, a linha é marcada para revisão manual do fator).
    _UNIDADES_EMBALAGEM = {'FD', 'FDO', 'FARDO', 'CX', 'CXA', 'CAIXA', 'PCT', 'PC', 'PACOTE', 'EMB', 'SC', 'SACO'}

    def carregar_fatores_externos(self, caminho):
        """Carrega overrides de fator de conversão de um arquivo externo (CSV ou XLSX)
        com colunas PALAVRA_CHAVE e FATOR. Ex de linha: FLOCUSCUZ;20
        Assim você corrige fatores sem editar o código."""
        try:
            if caminho.lower().endswith(('.xls', '.xlsx')):
                df = pd.read_excel(caminho)
            else:
                df = pd.read_csv(caminho, sep=None, engine='python')
            df.columns = [str(c).strip().upper() for c in df.columns]
            if 'PALAVRA_CHAVE' not in df.columns or 'FATOR' not in df.columns:
                logging.error(f"Arquivo de fatores '{caminho}' precisa das colunas PALAVRA_CHAVE e FATOR.")
                return
            novos = 0
            for _, r in df.iterrows():
                chave = str(r['PALAVRA_CHAVE']).strip().upper()
                fator = pd.to_numeric(r['FATOR'], errors='coerce')
                if chave and pd.notna(fator) and fator > 0:
                    self.OVERRIDES_EMBALAGEM[chave] = float(fator)
                    novos += 1
            logging.info(f"{novos} fatores de conversão externos carregados de {os.path.basename(caminho)}.")
        except Exception as e:
            logging.error(f"Não foi possível carregar fatores externos de '{caminho}': {e}")

    def converter_embalagens_entradas(self):
        """Aplica fator de conversão de fardos/pacotes/caixas para unidades.

        Cobre agora: FD/FARDO C/20, CX C/12, PCT C/10, C/20 genérico, 20X500G,
        12UN, dúzia (DZ) etc. Quando a unidade comercial indica embalagem (FD, CX,
        PCT...) mas não há número na descrição nem override, marca a linha em
        EMBALAGEM_SEM_FATOR=True para revisão manual (sai numa aba própria)."""
        logging.info("Aplicando conversão de embalagens nas entradas...")
        if self.entradas_xls.empty:
            return

        num = lambda m: int(m.group(1))
        # Ordem importa: padrões mais específicos primeiro. Compilados uma vez.
        # Cada padrão tem um rótulo de ORIGEM e um flag 'arriscado': padrões que
        # dependem de números colados na descrição (C500, TKC6, CARPLAST12X400G)
        # acertam muita coisa que antes ficava com fator 1, mas também podem pegar
        # um número que é modelo/bitola. Esses saem marcados para conferência em vez
        # de entrarem calados no estoque.
        _pats = [
            (r'\b(?:FD|FDO|FARDO)\s*C?[\s/.\-]*(\d+)\b', num, 'FD/FARDO', False),
            (r'\b(?:CX|CXA|CAIXA)\s*C?[\s/.\-]*(\d+)\b', num, 'CX/CAIXA', False),
            (r'\b(?:PCT|PACOTE|PC)\s*C?[\s/.\-]*(\d+)\b', num, 'PCT/PACOTE', False),
            (r'\b(?:SC|SACO|SACA)\s*C?[\s/.\-]*(\d+)\b', num, 'SC/SACO', False),
            (r'C[/.\-]\s*(\d+)\s*UN', num, 'C/NN UN', False),
            (r'\bC[/.\-]\s*(\d+)\b', num, 'C/NN', False),
            # NxN com unidade, SEM exigir fronteira antes do número: pega
            # "CARPLAST12X400G", "TIXAN20X1KG" — que antes escapavam porque
            # letra+dígito não formam \b e o padrão nunca casava.
            (r'(\d+)\s*X\s*\d+\s*(?:G|GR|KG|ML|L|LT|UN|UND|M)\b', num, 'NxN+unidade', False),
            (r'\b(\d+)\s*X\s*\d', num, 'NxN', False),
            (r'\b(\d+)\s*UN(?:ID|D)?\b', num, 'NN UN', False),
            (r'\bDZ\b|\bDUZIA\b|\bDÚZIA\b', 12, 'DUZIA', False),
            # --- padrões "grudados", aplicados só depois de todos os acima ---
            # "DISCO LIXA FLAP 4.1/2 040 THOM.C10" -> 10
            # "PARAFUSO CHIP4,0X40           C500" -> 500
            (r'(?:^|[\s.\-])C(\d{1,4})(?![A-Z0-9])', num, 'C+NN colado', True),
            # "ADESIVO EPOX ARALD.TRAN16G10M TKC6" -> 6 ; "TRENA ... FITA3/4 C6" -> 6
            (r'[A-Z]C(\d{1,3})(?![A-Z0-9])', num, 'letra+C+NN colado', True),
        ]
        patterns = [(re.compile(p), a, rot, risco) for p, a, rot, risco in _pats]

        def get_pack_factor(desc, un):
            # Blindagem: NaN vira '' (astype(str) de NaN pode retornar float em pandas recente)
            if not isinstance(desc, str):
                desc = '' if desc is None or (isinstance(desc, float) and pd.isna(desc)) else str(desc)
            if not isinstance(un, str):
                un = '' if un is None or (isinstance(un, float) and pd.isna(un)) else str(un)

            if un in ('DZ', 'DUZIA', 'DÚZIA'):
                return 12, False, 'UNIDADE=DZ'

            for pattern, action, rotulo, arriscado in patterns:
                match = pattern.search(desc)
                if match:
                    fator = action(match) if not isinstance(action, int) else action
                    # Fator absurdo é quase sempre número que não era embalagem
                    # (bitola, modelo, ano). Ignora e segue tentando.
                    if fator > 1 and fator <= 1000:
                        # sem_fator = False (temos um fator!). O 'arriscado' vira
                        # FATOR_REVISAR, um estado diferente de "não sei o fator".
                        return fator, False, (rotulo + ' [REVISAR]') if arriscado else rotulo
                    if fator == 1:
                        return 1, False, rotulo

            # Nenhum número na descrição: tenta override por palavra-chave
            for chave, fator in self.OVERRIDES_EMBALAGEM.items():
                if chave in desc:
                    return fator, False, 'OVERRIDE'

            # Unidade indica embalagem múltipla, mas não sabemos quantas unidades tem
            if un in self._UNIDADES_EMBALAGEM:
                return 1, True, 'SEM FATOR'

            return 1, False, 'UNITARIO'

        # Calcula o fator uma única vez por combinação (descrição, unidade) — descrições
        # repetidas (ex: milhares de "MASSA DE MILHO FLOCUSCUZ 500G") são resolvidas via cache.
        # fillna('') ANTES de astype(str): em pandas recente astype(str) de NaN pode devolver
        # float, então garantimos string preenchendo os nulos primeiro.
        desc_series = self.entradas_xls['DESCRICAO_NORM'].fillna('').map(lambda v: str(v).upper())
        if 'UNID' in self.entradas_xls.columns:
            un_series = self.entradas_xls['UNID'].fillna('').map(lambda v: str(v).upper())
        else:
            un_series = pd.Series([''] * len(self.entradas_xls), index=self.entradas_xls.index)
        un_series = un_series.str.replace(r'[^A-ZÀ-Ú]', '', regex=True)

        cache = {}
        fatores, flags, origens = [], [], []
        for d, u in zip(desc_series, un_series):
            chave = (d, u)
            if chave not in cache:
                cache[chave] = get_pack_factor(d, u)
            f, flag, origem = cache[chave]
            fatores.append(f)
            flags.append(flag)
            origens.append(origem)

        self.entradas_xls['FATOR_CONV'] = fatores
        self.entradas_xls['EMBALAGEM_SEM_FATOR'] = flags
        self.entradas_xls['FATOR_ORIGEM'] = origens
        self.entradas_xls['QTD_ENTRADA_AJUSTADA'] = self.entradas_xls['QTD_ENTRADA'] * self.entradas_xls['FATOR_CONV']
        self.entradas_xls['VLR_UNIT_ENTRADA_AJUSTADO'] = (self.entradas_xls['VLR_TOTAL_ENTRADA'] / self.entradas_xls['QTD_ENTRADA_AJUSTADA']).fillna(0).replace([float('inf'), -float('inf')], 0)

        n_override = int((self.entradas_xls['FATOR_CONV'] > 1).sum())
        logging.info(f"{n_override} entrada(s) tiveram fator de conversão > 1 aplicado (fardo/caixa/pacote/override).")

        self.entradas_xls['FATOR_REVISAR'] = self.entradas_xls['FATOR_ORIGEM'].str.contains(
            r'\[REVISAR\]', na=False, regex=True)
        n_arriscado = int(self.entradas_xls['FATOR_REVISAR'].sum())
        if n_arriscado:
            logging.warning(f"{n_arriscado} entrada(s) tiveram o fator deduzido de número COLADO na "
                            f"descrição (ex: 'THOM.C10' -> 10, 'CARPLAST12X400G' -> 12). "
                            f"Confira na aba 'Revisar_Fatores_Conversao' — pode ser bitola/modelo.")

        n_sem_fator = int(self.entradas_xls['EMBALAGEM_SEM_FATOR'].sum())
        if n_sem_fator:
            exemplos = self.entradas_xls.loc[self.entradas_xls['EMBALAGEM_SEM_FATOR'], 'DESCRICAO_NORM'].unique()[:5]
            logging.warning(
                f"{n_sem_fator} entrada(s) com unidade de EMBALAGEM (FD/CX/PCT...) sem número de unidades na descrição. "
                f"Fator ficou = 1 e elas saem na aba 'Revisar_Fatores_Conversao'. Exemplos: {list(exemplos)}"
            )
        logging.info("Conversão de embalagens concluída.")

    # Palavras genéricas demais para ajudar a diferenciar produtos (não contam como "match" de token)
    _STOPWORDS_PROD = {
        'DE', 'DA', 'DO', 'DOS', 'DAS', 'EM', 'NO', 'NA', 'C', 'COM', 'P', 'PARA', 'UN', 'UND', 'UNID',
        'KG', 'G', 'GR', 'ML', 'L', 'LT', 'X', 'CX', 'PCT', 'FD', 'CT', 'UNI', 'UNIDADE'
    }

    def _tokenizar_desc(self, descricao):
        """Quebra a descrição em palavras/números relevantes, descartando conectivos e
        unidades soltas que não ajudam a identificar QUAL produto é (evita que 'C/12 UN'
        conte como sinal de similaridade entre dois produtos diferentes).
        Usa cache: a mesma descrição só é tokenizada uma vez em toda a execução."""
        cache = self.__dict__.setdefault('_cache_tokens', {})
        if descricao is None or (isinstance(descricao, float) and pd.isna(descricao)):
            return []
        chave = str(descricao)
        cached = cache.get(chave)
        if cached is not None:
            return cached
        texto = chave.upper()
        tokens = re.findall(r'[A-ZÀ-Ú0-9]+', texto)
        resultado = [t for t in tokens if t not in self._STOPWORDS_PROD and len(t) > 1]
        cache[chave] = resultado
        return resultado

    @staticmethod
    def _tokens_casam(a, b):
        """Dois tokens 'casam' se forem iguais OU se um for prefixo do outro com pelo
        menos 4 letras. Isso resolve abreviações típicas de PDV: 'INSET' ~ 'INSETICIDA',
        'REFRIG' ~ 'REFRIGERANTE', 'MARG' não casa (só 4+ letras evita falso positivo
        curto demais tipo 'CA' ~ 'CAFE')."""
        if a == b:
            return True
        if len(a) >= 4 and len(b) >= 4 and (a.startswith(b) or b.startswith(a)):
            return True
        return False

    def _tokens_fortes(self, descricao):
        """Tokens 'de marca': palavras com 5+ caracteres e pelo menos uma letra
        (descarta números puros como '500G' já filtrado e códigos). São o melhor
        sinal de que duas descrições falam do MESMO produto (ex: FLOCUSCUZ, BAYGON,
        VERNIZ, SPARLACK)."""
        return {t for t in self._tokenizar_desc(descricao) if len(t) >= 5 and not t.isdigit()}

    def _tokens_set(self, descricao):
        """Set de tokens (cacheado) para cálculos de similaridade."""
        cache = self.__dict__.setdefault('_cache_tokenset', {})
        chave = str(descricao)
        s = cache.get(chave)
        if s is None:
            s = frozenset(self._tokenizar_desc(chave))
            cache[chave] = s
        return s

    def _bigramas(self, texto):
        """Conjunto de bigramas de caracteres (cacheado), usado como substituto barato
        do SequenceMatcher para medir semelhança de texto bruto. Dá resultado
        praticamente equivalente para nomes de produto a uma fração do custo."""
        cache = self.__dict__.setdefault('_cache_bigramas', {})
        s = cache.get(texto)
        if s is None:
            t = str(texto).upper().strip()
            s = frozenset(t[i:i+2] for i in range(len(t) - 1)) if len(t) > 1 else frozenset([t])
            cache[texto] = s
        return s

    def _sim_texto(self, desc1, desc2):
        """Similaridade de texto bruto via coeficiente de Dice sobre bigramas.
        Rápida (operações de conjunto) e sem o overhead do difflib."""
        b1, b2 = self._bigramas(desc1), self._bigramas(desc2)
        if not b1 or not b2:
            return 0.0
        inter = len(b1 & b2)
        return 2.0 * inter / (len(b1) + len(b2))

    def _similaridade_produto(self, desc1, desc2):
        """Similaridade combinando sobreposição de palavras-chave (Jaccard com
        equivalência por prefixo) + similaridade de texto bruto (bigramas/Dice)
        + bônus quando as duas descrições compartilham uma palavra 'de marca'
        (5+ letras). O Jaccard pesa mais porque é o que realmente diferencia produtos
        diferentes que só coincidem em números/unidades.

        Otimizado: quando não há NENHUM token em comum, retorna 0 sem calcular o texto."""
        # >>> NOVO: VETO POR VOLUME/PESO — 'PITU 600ML' x 'PITU 350ML' => similaridade 0, ponto.
        if getattr(self, 'usar_veto_volume', True) and self._extrair_veto_volume(desc1, desc2):
            return 0.0
        set1, set2 = self._tokens_set(desc1), self._tokens_set(desc2)

        if not set1 or not set2:
            return 0.0

        # Interseção rápida por igualdade
        iguais = set1 & set2
        inter = len(iguais)
        marca_compartilhada = any(len(t) >= 5 and not t.isdigit() for t in iguais)

        # Tokens que sobraram de cada lado tentam casar por prefixo (INSET ~ INSETICIDA)
        rem1 = [t for t in set1 if t not in iguais]
        rem2 = [t for t in set2 if t not in iguais]
        if rem1 and rem2:
            usados = set()
            for a in rem1:
                la = len(a)
                if la < 4:
                    continue
                for b in rem2:
                    if b in usados or len(b) < 4:
                        continue
                    if a.startswith(b) or b.startswith(a):
                        inter += 1
                        usados.add(b)
                        if min(la, len(b)) >= 5 and not a.isdigit():
                            marca_compartilhada = True
                        break

        # Sem token em comum: o score jamais passaria dos limiares -> retorna cedo.
        if inter == 0:
            return 0.0

        uniao = len(set1) + len(set2) - inter
        jaccard = inter / uniao if uniao else 0.0

        seq = self._sim_texto(desc1, desc2)
        score = 0.7 * jaccard + 0.3 * seq
        if marca_compartilhada:
            score += 0.15  # bônus: compartilhar 'FLOCUSCUZ'/'BAYGON'/'VERNIZ' é forte evidência
        return min(score, 1.0)

    @staticmethod
    def _custo_unit_ponderado(qtd_total, vlr_total, fallback=0.0):
        """Custo unitário por média PONDERADA: valor total comprado / quantidade total
        comprada. A média simples dos unitários (usada antes) distorce sempre que as
        quantidades das notas diferem — e esse número vai direto no VL_UNIT do H010."""
        try:
            q = float(qtd_total)
            v = float(vlr_total)
        except (TypeError, ValueError):
            q = v = 0.0
        if q and not pd.isna(q) and q != 0 and not pd.isna(v):
            return v / q
        try:
            fb = float(fallback)
            return 0.0 if pd.isna(fb) else fb
        except (TypeError, ValueError):
            return 0.0

    # ═════════════════════════════════════════════════════════════════════════
    # CHAVES NATURAIS + CADASTRO DE PRODUTOS (memória entre execuções e entre anos)
    # ═════════════════════════════════════════════════════════════════════════
    # Os dois lados não têm nenhum campo em comum. A venda (NFC-e) só identifica o
    # produto pelo código do PDV — o EAN vem 'SEM GTIN' e o NCM é genérico. A compra
    # identifica pelo EAN ou pelo código DO FORNECEDOR, que só é único junto com o
    # CNPJ dele (fornecedores diferentes reusam a mesma numeração). Por isso cada
    # linha ganha a CHAVE natural do seu lado:
    #
    #     venda : PDV:<cProd>
    #     compra: EAN:<gtin>            GTIN válido e de circulação global
    #             FORN:<cnpj>:<código>  sem EAN
    #
    # e o cadastro (cadastro_produtos.csv) liga cada CHAVE a um PRODUTO_ID próprio
    # (P000001, P000002...). Um produto pode ter várias chaves (vários fornecedores,
    # cProd duplicado no PDV, EAN da caixa e da unidade); uma chave só pertence a UM
    # produto. Casar é decidir isso uma vez por chave — depois todo movimento, de
    # qualquer ano, só consulta o cadastro.

    @staticmethod
    def _limpar_codigo(c):
        """Normaliza código: '001' -> '1', 1234.0 -> '1234', nulos/'N/A' -> ''."""
        if c is None or (isinstance(c, float) and pd.isna(c)):
            return ''
        if isinstance(c, float) and c.is_integer():
            c = int(c)
        c = str(c).strip()
        if c.lower() in ('', 'nan', 'none', 'n/a'):
            return ''
        return str(int(c)) if c.isdigit() else c

    @staticmethod
    def _normalizar_desc(txt):
        """Maiúsculas, sem acento e com espaços simples: 'Açúcar  Cristal' -> 'ACUCAR CRISTAL'.
        É o que permite a DESCRIÇÃO EXATA casar mesmo com diferença de acento/espaço."""
        if txt is None or (isinstance(txt, float) and pd.isna(txt)):
            return ''
        t = unicodedata.normalize('NFKD', str(txt)).encode('ascii', 'ignore').decode('ascii')
        return re.sub(r'\s+', ' ', t.upper()).strip()

    @staticmethod
    def _gtin_canonico(ean):
        """GTIN na forma canônica, ou '' se não serve como chave. O mesmo código chega
        como '78909052' ou '0000078909052' conforme o sistema do fornecedor: tira os
        zeros da esquerda e completa para 8, 13 ou 14 dígitos. Recusa dígito verificador
        errado e os GTINs de uso interno (prefixo 2, balança/etiqueta da loja, e EAN-8
        começado em 0): dois fornecedores podem usar o mesmo número para coisas diferentes."""
        e = str(ean or '').strip()
        if not e.isdigit() or len(e) > 14:
            return ''
        sig = e.lstrip('0')
        if len(sig) < 7:
            return ''  # '0000000000239': código interno preenchido com zeros
        g = sig.zfill(8) if len(sig) <= 8 else sig.zfill(13) if len(sig) <= 13 else sig
        if (len(g) == 8 and g[0] in '02') or (len(g) == 13 and (g[0] == '2' or g[:2] in ('02', '04')))                 or (len(g) == 14 and g[1] == '2'):
            return ''
        digitos = [int(c) for c in g]
        soma = sum(d * (3 if i % 2 == 0 else 1) for i, d in enumerate(reversed(digitos[:-1])))
        return g if (10 - soma % 10) % 10 == digitos[-1] else ''

    @staticmethod
    def _num(x, padrao=0.0):
        try:
            v = float(str(x).replace(',', '.'))
            return padrao if pd.isna(v) else v
        except (TypeError, ValueError):
            return padrao

    def atribuir_chaves(self):
        """Dá a cada linha de saída e de entrada a sua CHAVE natural (ver acima)."""
        if not self.nfce_saidas.empty:
            s = self.nfce_saidas
            s['xProd_NORM'] = s['xProd'].map(self._normalizar_desc)
            cod = s['cProd'].map(self._limpar_codigo)
            s['CHAVE'] = 'PDV:' + cod.where(cod != '', 'DESC:' + s['xProd_NORM'])
            ean = s['cEAN'].fillna('').astype(str).str.strip()
            s['cEAN'] = ean.map(self._gtin_canonico).replace('', 'N/A')
            logging.info(f"Saídas: {s['CHAVE'].nunique()} chave(s) de venda (código do PDV).")

        if not self.entradas_xls.empty:
            e = self.entradas_xls
            if (e['CNPJ_EMITENTE'] == '').all():
                logging.warning("Planilhas de entrada sem a coluna EMITENTE (CNPJ do fornecedor): produtos sem "
                                "EAN ficam identificados só pelo código do fornecedor, que se repete entre fornecedores.")
            cod = e['COD_ITEM'].map(self._limpar_codigo)
            forn = 'FORN:' + e['CNPJ_EMITENTE'] + ':' + cod.where(cod != '', 'DESC:' + e['DESCRICAO_NORM'])
            gtin = e['EAN'].map(self._gtin_canonico)
            ean_ok = gtin != ''
            # Linhas sem EAN de um (fornecedor, código) que em outras notas veio com UM EAN só
            # herdam esse EAN: é o mesmo produto, e assim juntam com as compras de outros fornecedores.
            eans_por_forn = gtin[ean_ok].groupby(forn[ean_ok]).agg(lambda x: set(x))
            unico = {k: next(iter(v)) for k, v in eans_por_forn.items() if len(v) == 1}
            herdado = forn.map(unico)
            chave = ('EAN:' + herdado).where(herdado.notna(), forn)
            e['CHAVE'] = ('EAN:' + gtin).where(ean_ok, chave)
            n_ean = int(e['CHAVE'].str.startswith('EAN:').sum())
            logging.info(f"Entradas: {e['CHAVE'].nunique()} chave(s) de compra "
                         f"({n_ean} de {len(e)} linhas identificadas por EAN, o resto por CNPJ+código do fornecedor).")

    def aplicar_fatores_do_cadastro(self):
        """O fator de conversão gravado no cadastro (coluna FATOR_CONV, preenchida pela aba
        Revisar_Fatores_Conversao ou à mão) vence o fator deduzido da descrição."""
        if self.entradas_xls.empty or 'CHAVE' not in self.entradas_xls.columns:
            return
        fatores = {c: self._num(r.get('FATOR_CONV')) for c, r in self._cad.items()
                   if self._num(r.get('FATOR_CONV')) > 0}
        e = self.entradas_xls
        f = e['CHAVE'].map(fatores)
        m = f.notna()
        if not m.any():
            return
        e['FATOR_CONV'] = e['FATOR_CONV'].astype(float)
        e.loc[m, 'FATOR_CONV'] = f[m].astype(float)
        e.loc[m, 'FATOR_ORIGEM'] = 'CADASTRO'
        e.loc[m, 'FATOR_REVISAR'] = False
        e.loc[m, 'EMBALAGEM_SEM_FATOR'] = False
        e['QTD_ENTRADA_AJUSTADA'] = e['QTD_ENTRADA'] * e['FATOR_CONV']
        e['VLR_UNIT_ENTRADA_AJUSTADO'] = (e['VLR_TOTAL_ENTRADA'] / e['QTD_ENTRADA_AJUSTADA']).fillna(0).replace(
            [float('inf'), -float('inf')], 0)
        logging.info(f"{int(m.sum())} linha(s) de entrada usaram o fator de conversão gravado no cadastro.")

    # ── leitura e gravação do cadastro ────────────────────────────────────────
    @staticmethod
    def _ler_csv_memoria(path, colunas):
        """Lê um CSV de memória (';', utf-8 com BOM, tudo como texto). Inexistente -> vazio.
        Corrompido -> move para '.bak' (nunca sobrescreve sem querer) e começa vazio."""
        vazio = pd.DataFrame(columns=colunas)
        if not path or not os.path.exists(path) or os.path.getsize(path) == 0:
            return vazio
        try:
            df = pd.read_csv(path, sep=';', dtype=str, keep_default_na=False, encoding='utf-8-sig')
            if len(df.columns) == 1:  # arquivo salvo pelo Excel com vírgula
                df = pd.read_csv(path, sep=',', dtype=str, keep_default_na=False, encoding='utf-8-sig')
            df.columns = [str(c).strip() for c in df.columns]
            faltam = [c for c in colunas if c not in df.columns]
            if faltam:
                raise ValueError(f"colunas ausentes: {faltam}")
            return df[colunas].fillna('')
        except Exception as e:
            bak = path + '.bak'
            logging.error(f"Não consegui ler '{path}' ({e}). Arquivo movido para '{bak}'; começando vazio.")
            try:
                os.replace(path, bak)
            except OSError:
                pass
            return vazio

    @staticmethod
    def _gravar_csv(df, path):
        pasta = os.path.dirname(path)
        if pasta:
            os.makedirs(pasta, exist_ok=True)
        tmp = path + '.tmp'
        df.to_csv(tmp, sep=';', index=False, encoding='utf-8-sig')
        os.replace(tmp, path)

    def _carregar_cadastro(self):
        self._cad = {}          # CHAVE -> registro (colunas de CADASTRO_COLUNAS)
        self._rejeitados = {}   # (CHAVE_VENDA, CHAVE_COMPRA) -> registro
        for r in self._ler_csv_memoria(self.cadastro_path, CADASTRO_COLUNAS).to_dict('records'):
            r = {k: str(v).strip() for k, v in r.items()}
            if r['CHAVE']:
                self._cad[r['CHAVE']] = r
        for r in self._ler_csv_memoria(self.rejeitados_path, REJEITADOS_COLUNAS).to_dict('records'):
            r = {k: str(v).strip() for k, v in r.items()}
            if r['CHAVE_VENDA'] and r['CHAVE_COMPRA']:
                self._rejeitados[(r['CHAVE_VENDA'], r['CHAVE_COMPRA'])] = r
        ids = [int(r['PRODUTO_ID'][1:]) for r in self._cad.values() if re.fullmatch(r'P\d+', r['PRODUTO_ID'])]
        self._ultimo_id = max(ids, default=0)
        if self._cad or self._rejeitados:
            logging.info(f"Cadastro: {len(self._cad)} chave(s) em {len(set(ids))} produto(s) e "
                         f"{len(self._rejeitados)} par(es) rejeitado(s), de '{self.cadastro_path}'.")
        else:
            logging.info(f"Cadastro: '{self.cadastro_path}' ainda não existe — começando vazio.")

    def salvar_cadastro(self):
        """Grava o cadastro e os pares rejeitados (escrita atômica: .tmp + replace). Os
        casamentos AUTO A CONFIRMAR não entram: só o que é MANUAL ou AUTO seguro."""
        df = pd.DataFrame([{k: r.get(k, '') for k in CADASTRO_COLUNAS} for r in self._cad.values()],
                          columns=CADASTRO_COLUNAS)
        df = df.sort_values(['PRODUTO_ID', 'LADO', 'CHAVE'], ascending=[True, False, True], kind='stable')
        self._gravar_csv(df, self.cadastro_path)
        rej = pd.DataFrame(list(self._rejeitados.values()), columns=REJEITADOS_COLUNAS)
        self._gravar_csv(rej, self.rejeitados_path)
        logging.info(f"Cadastro gravado: {len(df)} chave(s) em "
                     f"{df.loc[df['PRODUTO_ID'] != '', 'PRODUTO_ID'].nunique()} produto(s) -> '{self.cadastro_path}'.")

    def cadastro_como_df(self):
        """Cadastro gravado + casamentos a confirmar desta execução (aba Cadastro_Produtos)."""
        linhas = [dict({k: r.get(k, '') for k in CADASTRO_COLUNAS}, GRAVADO='Sim') for r in self._cad.values()]
        for c, pid in self._prov.items():
            linhas.append({'CHAVE': c, 'LADO': 'V' if c.startswith('PDV:') else 'C', 'PRODUTO_ID': pid,
                           'DESCRICAO': self._desc_chave.get(c, ''), 'FATOR_CONV': '',
                           'ORIGEM': self._prov_origem.get(c, ''), 'SCORE': '', 'DATA': '',
                           'GRAVADO': 'Não (a confirmar)'})
        df = pd.DataFrame(linhas, columns=CADASTRO_COLUNAS + ['GRAVADO'])
        return df.sort_values(['PRODUTO_ID', 'LADO', 'CHAVE'], ascending=[True, False, True],
                              kind='stable').reset_index(drop=True)

    # ── vínculo de chaves a produtos ──────────────────────────────────────────
    def _pid(self, chave):
        """PRODUTO_ID da chave (gravado ou a confirmar), ou None se ela ainda não tem produto."""
        r = self._cad.get(chave)
        if r and r.get('PRODUTO_ID'):
            return r['PRODUTO_ID']
        return self._prov.get(chave)

    def _novo_produto_id(self):
        self._ultimo_id += 1
        return f"P{self._ultimo_id:06d}"

    def _vincular(self, chaves_v, chaves_c, origem, score=None, gravar=True, descricoes=None):
        """Põe as chaves de venda e de compra no MESMO produto. Retorna o PRODUTO_ID ou None.

        Produto de destino: o da chave de venda, se ela já tiver; senão o da compra; senão um novo.
          * MANUAL: a decisão do usuário vale — a chave de compra sai do produto em que estava
            e vai para o da venda (é assim que se corrige um vínculo errado).
          * AUTO: nunca mexe em chave que já é de outro produto, nem casa par rejeitado."""
        chaves_v = [c for c in dict.fromkeys(chaves_v) if c]
        chaves_c = [c for c in dict.fromkeys(chaves_c) if c]
        if not chaves_v or not chaves_c:
            return None
        manual = origem.startswith('MANUAL')
        if not manual and any((v, c) in self._rejeitados for v in chaves_v for c in chaves_c):
            return None
        pids_v = sorted({p for p in map(self._pid, chaves_v) if p})
        pids_c = sorted({p for p in map(self._pid, chaves_c) if p})
        if manual:
            alvo = (pids_v or pids_c or [None])[0]
        else:
            pids = set(pids_v) | set(pids_c)
            if len(pids) > 1:
                return None
            alvo = next(iter(pids), None)
        alvo = alvo or self._novo_produto_id()
        if descricoes:
            self._desc_chave.update({k: v for k, v in descricoes.items() if v})

        hoje = date.today().isoformat()
        for c in chaves_v + chaves_c:
            atual = self._cad.get(c)
            if atual and atual.get('PRODUTO_ID') == alvo and not (manual and not atual.get('ORIGEM', '').startswith('MANUAL')):
                continue  # já está neste produto
            if not gravar:
                if not (atual and atual.get('PRODUTO_ID')):
                    self._prov[c] = alvo
                    self._prov_origem[c] = origem
                continue
            rec = atual or {'CHAVE': c, 'FATOR_CONV': ''}
            rec.update({'LADO': 'V' if c.startswith('PDV:') else 'C', 'PRODUTO_ID': alvo,
                        'DESCRICAO': self._desc_chave.get(c, rec.get('DESCRICAO', '')), 'ORIGEM': origem,
                        'SCORE': '' if score is None else f"{float(score):.2f}", 'DATA': hoje})
            self._cad[c] = rec
            self._prov.pop(c, None)
            self._prov_origem.pop(c, None)
        return alvo

    # ═════════════════════════════════════════════════════════════════════════
    # TRAVA 3 — VETO POR VOLUME/PESO
    # ═════════════════════════════════════════════════════════════════════════
    # "12X350ML" -> "12 350ML" (o multiplicador de embalagem não é volume)
    _RE_X_ENTRE_NUMEROS = re.compile(r'(?<=\d)X(?=\d)')
    # número + unidade: 600ML, 1,5L, 2LT, 500G, 500 GR, 1KG ...
    _RE_MEDIDA = re.compile(
        r'(?<![A-Z0-9.,])(\d+(?:[.,]\d+)?)\s*'
        r'(KGS?|KILOS?|QUILOS?|GRAMAS?|GRS?|MG|MLS?|LITROS?|LTS?|L|G)\b')

    @staticmethod
    def _unidade_para_base(un):
        """Converte a unidade para a base: volume em ML ('V') e massa em G ('M')."""
        if un in ('ML', 'MLS'):
            return 'V', 1.0
        if un in ('L', 'LT', 'LTS', 'LITRO', 'LITROS'):
            return 'V', 1000.0
        if un in ('G', 'GR', 'GRS', 'GRAMA', 'GRAMAS'):
            return 'M', 1.0
        if un in ('KG', 'KGS', 'KILO', 'KILOS', 'QUILO', 'QUILOS'):
            return 'M', 1000.0
        if un == 'MG':
            return 'M', 0.001
        return None, 0.0

    def _extrair_volume_regex(self, descricao):
        """Extrai as medidas de peso/volume da descrição, já convertidas para a base.
        Retorna frozenset de tuplas (grandeza, valor): 'PITU 600ML' -> {('V', 600.0)};
        '1,5L' -> {('V', 1500.0)}; '1KG' -> {('M', 1000.0)}. Cacheado por descrição."""
        cache = self.__dict__.setdefault('_cache_volume', {})
        if descricao is None or (isinstance(descricao, float) and pd.isna(descricao)):
            return frozenset()
        chave = str(descricao)
        res = cache.get(chave)
        if res is None:
            texto = self._RE_X_ENTRE_NUMEROS.sub(' ', chave.upper())
            medidas = set()
            for num, un in self._RE_MEDIDA.findall(texto):
                grandeza, fator = self._unidade_para_base(un)
                if grandeza is None:
                    continue
                try:
                    valor = float(num.replace(',', '.')) * fator
                except ValueError:
                    continue
                if valor > 0:
                    medidas.add((grandeza, round(valor, 3)))
            res = frozenset(medidas)
            cache[chave] = res
        return res

    def _extrair_veto_volume(self, desc_a, desc_b):
        """True = VETO. Só veta quando as DUAS descrições trazem medida da MESMA grandeza
        (volume x volume, peso x peso) e nenhuma medida coincide: 600ML x 350ML -> veta;
        1L x 1000ML -> não veta; 'PITU' x 'PITU 600ML' (um lado sem medida) -> não veta;
        500G x 500ML (grandezas diferentes) -> não veta."""
        va, vb = self._extrair_volume_regex(desc_a), self._extrair_volume_regex(desc_b)
        if not va or not vb:
            return False
        for g in ('V', 'M'):
            ga = {v for (gr, v) in va if gr == g}
            gb = {v for (gr, v) in vb if gr == g}
            if ga and gb and ga.isdisjoint(gb):
                return True
        return False

    # ═════════════════════════════════════════════════════════════════════════
    # TRAVA 2 — PREÇO (triangulação financeira / guard-rail)
    # ═════════════════════════════════════════════════════════════════════════
    @staticmethod
    def _razao_preco(venda_vlr_unit, custo_vlr_unit):
        """Venda_VLR_UNIT / Custo_VLR_UNIT, ou None se algum dos preços for inválido/zero."""
        try:
            v, c = float(venda_vlr_unit), float(custo_vlr_unit)
        except (TypeError, ValueError):
            return None
        if pd.isna(v) or pd.isna(c) or v <= 0 or c <= 0:
            return None
        return v / c

    def _validar_trava_preco(self, venda_vlr_unit, custo_vlr_unit):
        """True = preço plausível, pode seguir no automático.
        False = ABORTA o match automático (vai para a validação manual).
        Sem preço de venda ou de custo para conferir também aborta: na dúvida, pergunta."""
        if not getattr(self, 'usar_trava_preco', True):
            return True
        razao = self._razao_preco(venda_vlr_unit, custo_vlr_unit)
        if razao is None:
            return False
        return float(self.trava_preco_min) <= razao <= float(self.trava_preco_max)

    def _anotar_recusa_auto(self, s_out, s_in):
        """Guarda o motivo da recusa; ele aparece na coluna OBSERVACAO da validação manual."""
        razao = self._razao_preco(s_out.get('VLR_UNIT_SAIDA'), s_in.get('VLR_UNIT_ENTRADA_AJUSTADO'))
        if razao is None:
            nota = 'TRAVA DE PREÇO: sem preço de venda/custo para conferir — match automático abortado'
        else:
            nota = (f"TRAVA DE PREÇO: venda/custo = {razao:.2f}x, fora de "
                    f"{self.trava_preco_min}x–{self.trava_preco_max}x — match automático abortado")
        self._recusas_auto[(str(s_out.get('xProd_NORM', '')), str(s_in.get('DESCRICAO_NORM', '')))] = nota

    def _parear_por_chave(self, sobras_saida, sobras_entrada, col_saida, col_entrada,
                          limiar_similaridade, tipo_match, motivo_rejeicao=''):
        """Pareia saídas x entradas que compartilham a MESMA chave exata (código interno,
        EAN ou descrição), consolidando N compras em 1 produto.

        Substitui _resolver_1_para_1_com_freio, que tinha dois defeitos graves:

          * removia das sobras TODAS as linhas com aquela chave, mas casava só uma —
            as demais entradas sumiam do relatório (nem somavam saldo, nem apareciam
            como 'APENAS ENTRADA'). Furo silencioso de estoque.
          * o drop_duplicates por chave colapsava saídas distintas que dividem o mesmo
            cProd/cEAN, descartando quantidade vendida.

        Agora tudo é feito por ÍNDICE: só sai das sobras o que foi efetivamente casado.
        O freio de similaridade é aplicado entrada a entrada — uma compra que colidiu
        por código continua disponível para casar por EAN/descrição/marca depois.

        Retorna (df_matches, sobras_saida_restantes, sobras_entrada_restantes, rejeitados).
        """
        vazio = pd.DataFrame()
        if sobras_saida.empty or sobras_entrada.empty:
            return vazio, sobras_saida, sobras_entrada, vazio

        lado_s = sobras_saida[sobras_saida[col_saida].astype(str) != 'N/A']
        lado_e = sobras_entrada[sobras_entrada[col_entrada].astype(str) != 'N/A']
        if lado_s.empty or lado_e.empty:
            return vazio, sobras_saida, sobras_entrada, vazio

        grupos_e = dict(list(lado_e.groupby(col_entrada)))
        linhas, rejeitados = [], []
        idx_s_usados, idx_e_usados = [], []

        for chave, g_s in lado_s.groupby(col_saida):
            g_e = grupos_e.get(chave)
            if g_e is None or g_e.empty:
                continue

            descs_saida = [d for d in g_s['xProd_NORM'].tolist() if d]
            if not descs_saida:
                continue

            aceitas, sims = [], []
            for j, r_e in g_e.iterrows():
                sim = max(self._similaridade_produto(d, r_e['DESCRICAO_NORM'])
                          for d in descs_saida)
                if sim >= limiar_similaridade:
                    aceitas.append(j)
                    sims.append(sim)
                elif motivo_rejeicao:
                    rejeitados.append({
                        'xProd': g_s.iloc[0]['xProd'],
                        'DESCRICAO_NORM': r_e['DESCRICAO_NORM'],
                        'DESC_SIMILARIDADE': sim,
                        'MOTIVO': motivo_rejeicao,
                    })

            if not aceitas:
                continue  # nada removido das sobras: a chave inteira segue disponível

            g_ok = g_e.loc[aceitas]
            qtd_ent = float(pd.to_numeric(g_ok['QTD_ENTRADA_AJUSTADA'], errors='coerce').fillna(0).sum())
            vlr_ent = float(pd.to_numeric(g_ok['VLR_TOTAL_ENTRADA'], errors='coerce').fillna(0).sum())
            base = g_s.iloc[0]

            linha = {
                'cProd': base['cProd'],
                'cEAN': base['cEAN'],
                'xProd': base['xProd'],
                'xProd_NORM': base['xProd_NORM'],
                'QTD_SAIDA': float(pd.to_numeric(g_s['QTD_SAIDA'], errors='coerce').fillna(0).sum()),
                'COD_ITEM': g_ok.iloc[0]['COD_ITEM'],
                'EAN': g_ok.iloc[0]['EAN'],
                'DESCRICAO_NORM': g_ok.iloc[0]['DESCRICAO_NORM'],
                'QTD_ENTRADA_AJUSTADA': qtd_ent,
                'VLR_TOTAL_ENTRADA': vlr_ent,
                'VLR_UNIT_ENTRADA_AJUSTADO': self._custo_unit_ponderado(
                    qtd_ent, vlr_ent, g_ok['VLR_UNIT_ENTRADA_AJUSTADO'].mean()),
                'DESC_SIMILARIDADE': max(sims),
                'TIPO_MATCH': tipo_match,
                'CPRODS_MATCH': '|'.join(sorted({str(x) for x in g_s['cProd']})),
                'COD_ITENS_MATCH': '|'.join(sorted({str(x) for x in g_ok['COD_ITEM']})),
                'N_ENTRADAS_SOMADAS': len(aceitas),
                'N_SAIDAS_SOMADAS': len(g_s),
            }
            for col in ('FATOR_CONV', 'EMBALAGEM_SEM_FATOR'):
                if col in g_ok.columns:
                    linha[col] = g_ok[col].max()

            linhas.append(linha)
            idx_s_usados.extend(g_s.index.tolist())
            idx_e_usados.extend(aceitas)

        df_matches = pd.DataFrame(linhas) if linhas else vazio
        df_rej = pd.DataFrame(rejeitados) if rejeitados else vazio
        return (df_matches,
                sobras_saida.drop(index=idx_s_usados),
                sobras_entrada.drop(index=idx_e_usados),
                df_rej)

    def _match_por_palavra_chave(self, sobras_saida, sobras_entrada, limiar_guarda=None):
        """Pareia automaticamente sobras de saída x entrada que compartilham uma
        palavra-chave forte (5+ letras, com equivalência por prefixo) que aparece em
        EXATAMENTE uma saída e EXATAMENTE uma entrada não pareadas.

        Ex: 'FLOCUSCUZ 500G' (saída) x 'FLOCAO MILHO FLOCUSCUZ FD C/20' (entrada);
            'INSET BAYGON' x 'INSETICIDA BAYGON MULTI ACAO 300ML';
            'VERNIZ' x 'VERNIZ SPARLACK 350ML'.

        Ambiguidade (a mesma palavra em 2+ saídas ou 2+ entradas, ou a mesma saída
        apontando para entradas diferentes) NÃO é resolvida aqui: fica para a
        validação manual. Um freio de similaridade mínima evita casar produtos
        totalmente diferentes que por azar dividem uma palavra rara.

        Retorna (df_matches, sobras_saida_restantes, sobras_entrada_restantes)."""
        vazio = pd.DataFrame()
        if limiar_guarda is None:
            limiar_guarda = getattr(self, 'limiar_marca', 0.55)
        if sobras_saida.empty or sobras_entrada.empty:
            return vazio, sobras_saida, sobras_entrada

        # Índice: palavra forte -> conjunto de índices onde ela aparece
        idx_saida, idx_entrada = {}, {}
        for i, r in sobras_saida.iterrows():
            for t in self._tokens_fortes(r['xProd_NORM']):
                idx_saida.setdefault(t, set()).add(i)
        for j, r in sobras_entrada.iterrows():
            for t in self._tokens_fortes(r['DESCRICAO_NORM']):
                idx_entrada.setdefault(t, set()).add(j)

        # Casa palavras equivalentes por prefixo (INSET ~ INSETICIDA) unificando índices
        pares_por_saida = {}   # i -> set de j candidatos por palavra única
        for t_s, s_ids in idx_saida.items():
            for t_e, e_ids in idx_entrada.items():
                if not self._tokens_casam(t_s, t_e):
                    continue
                if len(s_ids) == 1 and len(e_ids) == 1:
                    i, j = next(iter(s_ids)), next(iter(e_ids))
                    pares_por_saida.setdefault(i, set()).add(j)

        # Só aceita saídas que apontam para UMA única entrada; e cada entrada só pode
        # ser reivindicada por UMA saída (senão vira ambiguidade -> validação manual).
        candidatos = {i: next(iter(js)) for i, js in pares_por_saida.items() if len(js) == 1}
        uso_entrada = {}
        for i, j in candidatos.items():
            uso_entrada.setdefault(j, []).append(i)
        pares_finais = [(i, j) for j, is_ in uso_entrada.items() if len(is_) == 1 for i in is_]

        linhas, usados_i, usados_j = [], [], []
        for i, j in pares_finais:
            s_out, s_in = sobras_saida.loc[i], sobras_entrada.loc[j]
            sim = self._similaridade_produto(s_out['xProd_NORM'], s_in['DESCRICAO_NORM'])
            if sim < limiar_guarda:
                continue  # palavra em comum mas resto incompatível demais: deixa para validação

            # >>> NOVO: TRAVA DE PREÇO — venda/custo absurdo aborta o automático
            if not self._validar_trava_preco(s_out.get('VLR_UNIT_SAIDA'),
                                             s_in.get('VLR_UNIT_ENTRADA_AJUSTADO')):
                self._anotar_recusa_auto(s_out, s_in)
                continue
            linha = {**s_out.to_dict(), **s_in.to_dict()}
            linha['DESC_SIMILARIDADE'] = sim
            linha['TIPO_MATCH'] = 'PALAVRA-CHAVE (MARCA)'
            linha['CPRODS_MATCH'] = str(s_out.get('cProd', ''))
            linha['COD_ITENS_MATCH'] = str(s_in.get('COD_ITEM', ''))
            linha['N_ENTRADAS_SOMADAS'] = 1
            linha['N_SAIDAS_SOMADAS'] = 1
            linhas.append(linha)
            usados_i.append(i)
            usados_j.append(j)

        if not linhas:
            return vazio, sobras_saida, sobras_entrada

        df_matches = pd.DataFrame(linhas)
        return (
            df_matches,
            sobras_saida.drop(index=usados_i),
            sobras_entrada.drop(index=usados_j),
        )

    def _match_por_similaridade_alta(self, sobras_saida, sobras_entrada,
                                     limiar=None, margem=None):
        """Aceita automaticamente pares cuja descrição é quase idêntica, mesmo sem
        nenhuma chave em comum. É a "muleta" para o caso mais frequente desta base:
        o mesmo produto escrito abreviado no PDV e por extenso na nota do fornecedor.

            'ENERGETICO RED BULL 250ML'  x  'ENERG RED BULL 250ML'    -> 100%
            'LIMPA VIDROS AZULIM 500ML'  x  'LIMPA VIDRO AZULIM 500ML' -> 100%

        Esses dois já pontuavam 1.0 no cálculo de similaridade, mas caíam na
        validação manual porque o único caminho automático sem chave era o
        _match_por_palavra_chave — que exige a palavra forte aparecendo em
        EXATAMENTE uma saída e uma entrada. Com várias notas de Red Bull no mês,
        'ENERG' aparece em várias entradas, vira ambiguidade e o par era descartado
        mesmo batendo 100%.

        Regra de aceite (casamento mútuo, para não errar no automático):
          * o score tem de ser >= limiar (padrão 0.88);
          * a entrada J tem de ser a MELHOR opção da saída I **e** a saída I tem de
            ser a melhor opção da entrada J;
          * o segundo colocado de cada lado tem de ficar pelo menos `margem` abaixo.
            Se dois produtos empatam (ex: 'REFRI COLA 2L' contra '...2L' e '...1,5L'),
            ninguém é aceito e o par vai para a validação manual, como antes.

        Retorna (df_matches, sobras_saida_restantes, sobras_entrada_restantes)."""
        vazio = pd.DataFrame()
        if limiar is None:
            limiar = float(getattr(self, 'limiar_auto_similaridade', 0.88))
        if margem is None:
            margem = float(getattr(self, 'margem_ambiguidade', 0.05))
        if sobras_saida.empty or sobras_entrada.empty or limiar <= 0:
            return vazio, sobras_saida, sobras_entrada

        entradas = list(sobras_entrada.iterrows())

        # Índice invertido token -> posições, para não comparar todos contra todos.
        indice = {}
        for pos, (j, r) in enumerate(entradas):
            for t in self._tokens_set(r['DESCRICAO_NORM']):
                indice.setdefault(t, []).append(pos)
        total = max(len(entradas), 1)
        limite_freq = max(50, int(0.30 * total))
        genericos = {t for t, lst in indice.items() if len(lst) > limite_freq}

        # melhor e segundo melhor de cada saída
        melhor_da_saida = {}   # i -> (score, pos)
        segundo_da_saida = {}  # i -> score
        # melhor e segundo melhor de cada entrada
        melhor_da_entrada = {}   # pos -> (score, i)
        segundo_da_entrada = {}  # pos -> score

        for i, s_out in sobras_saida.iterrows():
            candidatos = set()
            for t in self._tokens_set(s_out['xProd_NORM']):
                if t not in genericos:
                    candidatos.update(indice.get(t, ()))
            if not candidatos:
                continue

            for pos in candidatos:
                score = self._similaridade_produto(s_out['xProd_NORM'],
                                                   entradas[pos][1]['DESCRICAO_NORM'])
                if score < limiar:
                    continue

                atual = melhor_da_saida.get(i)
                if atual is None or score > atual[0]:
                    if atual is not None:
                        segundo_da_saida[i] = max(segundo_da_saida.get(i, 0.0), atual[0])
                    melhor_da_saida[i] = (score, pos)
                else:
                    segundo_da_saida[i] = max(segundo_da_saida.get(i, 0.0), score)

                atual_e = melhor_da_entrada.get(pos)
                if atual_e is None or score > atual_e[0]:
                    if atual_e is not None:
                        segundo_da_entrada[pos] = max(segundo_da_entrada.get(pos, 0.0), atual_e[0])
                    melhor_da_entrada[pos] = (score, i)
                else:
                    segundo_da_entrada[pos] = max(segundo_da_entrada.get(pos, 0.0), score)

        linhas, usados_i, usados_j = [], [], []
        ambiguos = 0
        for i, (score, pos) in melhor_da_saida.items():
            # casamento mútuo: a entrada também tem de eleger esta saída
            m_e = melhor_da_entrada.get(pos)
            if not m_e or m_e[1] != i:
                ambiguos += 1
                continue
            # margem sobre o segundo colocado dos dois lados
            if (score - segundo_da_saida.get(i, 0.0) < margem or
                    score - segundo_da_entrada.get(pos, 0.0) < margem):
                ambiguos += 1
                continue

            j, s_in = entradas[pos]
            s_out = sobras_saida.loc[i]
            # >>> NOVO: TRAVA DE PREÇO — venda/custo absurdo aborta o automático
            if not self._validar_trava_preco(s_out.get('VLR_UNIT_SAIDA'),
                                             s_in.get('VLR_UNIT_ENTRADA_AJUSTADO')):
                self._anotar_recusa_auto(s_out, s_in)
                continue
            linha = {**s_out.to_dict(), **s_in.to_dict()}
            linha['DESC_SIMILARIDADE'] = score
            linha['TIPO_MATCH'] = 'SIMILARIDADE ALTA (AUTO)'
            linha['CPRODS_MATCH'] = str(s_out.get('cProd', ''))
            linha['COD_ITENS_MATCH'] = str(s_in.get('COD_ITEM', ''))
            linha['N_ENTRADAS_SOMADAS'] = 1
            linha['N_SAIDAS_SOMADAS'] = 1
            linhas.append(linha)
            usados_i.append(i)
            usados_j.append(j)

        if ambiguos:
            logging.info(f"{ambiguos} par(es) acima do limiar ficaram AMBÍGUOS (empate entre "
                         f"candidatos) e seguiram para a validação manual.")
        if not linhas:
            return vazio, sobras_saida, sobras_entrada

        return (pd.DataFrame(linhas),
                sobras_saida.drop(index=usados_i),
                sobras_entrada.drop(index=usados_j))

    def _montar_catalogo(self):
        """Uma linha por CHAVE de cada lado, com a descrição mais usada e os totais de todos os
        períodos carregados. É sobre ESTE catálogo — e não sobre cada nota — que o casamento
        é decidido. Retorna (cat_venda, cat_compra) com as colunas que as funções de
        pareamento esperam (cProd/xProd_NORM de um lado, COD_ITEM/DESCRICAO_NORM do outro)."""
        s = self.nfce_saidas
        desc_v = (s.groupby(['CHAVE', 'xProd_NORM'])['QTD_SAIDA'].sum().reset_index()
                   .sort_values(['CHAVE', 'QTD_SAIDA'], ascending=[True, False], kind='stable')
                   .drop_duplicates('CHAVE').set_index('CHAVE')['xProd_NORM'])
        cat_v = s.groupby('CHAVE').agg(
            QTD_SAIDA=('QTD_SAIDA', 'sum'), VLR_TOTAL_SAIDA=('VLR_TOTAL_SAIDA', 'sum'),
            cEAN=('cEAN', lambda x: next((v for v in x if v != 'N/A'), 'N/A')))
        cat_v['xProd_NORM'] = desc_v
        cat_v = cat_v.reset_index()
        cat_v['cProd'] = cat_v['CHAVE']
        cat_v['xProd'] = cat_v['xProd_NORM']
        cat_v['VLR_UNIT_SAIDA'] = (cat_v['VLR_TOTAL_SAIDA'] /
                                   cat_v['QTD_SAIDA'].where(cat_v['QTD_SAIDA'] != 0)).fillna(0.0)

        e = self.entradas_xls
        desc_c = (e.groupby(['CHAVE', 'DESCRICAO_NORM'])['QTD_ENTRADA_AJUSTADA'].sum().reset_index()
                   .sort_values(['CHAVE', 'QTD_ENTRADA_AJUSTADA'], ascending=[True, False], kind='stable')
                   .drop_duplicates('CHAVE').set_index('CHAVE')['DESCRICAO_NORM'])
        agg = {'QTD_ENTRADA_AJUSTADA': ('QTD_ENTRADA_AJUSTADA', 'sum'),
               'VLR_TOTAL_ENTRADA': ('VLR_TOTAL_ENTRADA', 'sum'),
               'FATOR_CONV': ('FATOR_CONV', 'max'),
               'EMBALAGEM_SEM_FATOR': ('EMBALAGEM_SEM_FATOR', 'max'),
               'FORNECEDOR': ('RAZAO_EMITENTE',
                              lambda x: ' | '.join(sorted({str(v) for v in x if str(v).strip()}))[:120])}
        cat_c = e.groupby('CHAVE').agg(**agg)
        cat_c['DESCRICAO_NORM'] = desc_c
        cat_c = cat_c.reset_index()
        cat_c['COD_ITEM'] = cat_c['CHAVE']
        cat_c['EAN'] = cat_c['CHAVE'].where(cat_c['CHAVE'].str.startswith('EAN:'), 'N/A').str.replace('EAN:', '', regex=False)
        cat_c['VLR_UNIT_ENTRADA_AJUSTADO'] = [
            self._custo_unit_ponderado(q, v) for q, v in zip(cat_c['QTD_ENTRADA_AJUSTADA'], cat_c['VLR_TOTAL_ENTRADA'])]

        self._desc_chave.update(dict(zip(cat_v['CHAVE'], cat_v['xProd_NORM'])))
        self._desc_chave.update(dict(zip(cat_c['CHAVE'], cat_c['DESCRICAO_NORM'])))
        return cat_v, cat_c

    def _aplicar_matches(self, df):
        """Leva ao cadastro os casamentos automáticos de um passo. DESCRICAO EXATA e EAN são
        gravados; PALAVRA-CHAVE e SIMILARIDADE só com score >= cadastro_auto_min_sim — abaixo
        disso valem nesta execução e vão para a validação como AUTO A CONFIRMAR."""
        if df is None or df.empty:
            return 0
        n = 0
        for r in df.to_dict('records'):
            cv = [c for c in str(r.get('CPRODS_MATCH') or '').split('|') if c]
            cc = [c for c in str(r.get('COD_ITENS_MATCH') or '').split('|') if c]
            tipo = str(r.get('TIPO_MATCH', ''))
            score = self._num(r.get('DESC_SIMILARIDADE'))
            gravar = tipo in ('DESCRICAO EXATA', 'EAN EXATO') or score >= float(self.cadastro_auto_min_sim)
            origem = f"AUTO - {tipo}" if gravar else f"AUTO A CONFIRMAR - {tipo}"
            novo = any(not self._pid(c) for c in cv + cc)
            if self._vincular(cv, cc, origem, score, gravar=gravar):
                n += novo
                if not gravar:
                    self._provisorios.append((cv, cc, tipo, score))
        return n

    def _sem_produto(self, cat):
        return cat[[not self._pid(c) for c in cat['CHAVE']]]

    def comparar_estoque(self):
        """Casa as chaves de venda com as de compra e monta o estoque por PRODUTO_ID.

        0) o que já está no cadastro não é questionado de novo;
        1) DESCRIÇÃO IDÊNTICA (qualquer descrição que a chave já tenha usado);
        2) EAN;
        3) PALAVRA-CHAVE de marca única;
        4) SIMILARIDADE ALTA com aceite mútuo + travas de preço e volume.
        O que sobrar vai para a planilha de validação, ordenado pelo valor envolvido."""
        logging.info("Iniciando cruzamento de Entradas vs Saídas (cadastro de produtos)...")
        if self.nfce_saidas.empty or self.entradas_xls.empty:
            logging.error("Dados de entrada ou saída ausentes. Execute as extrações primeiro.")
            return
        if 'CHAVE' not in self.nfce_saidas.columns or 'CHAVE' not in self.entradas_xls.columns:
            self.atribuir_chaves()

        self._recusas_auto, self._provisorios = {}, []
        self._prov, self._prov_origem = {}, {}
        cat_v, cat_c = self._montar_catalogo()
        self.catalogo_venda, self.catalogo_compra = cat_v, cat_c  # a tela usa para buscar outra chave
        nv = sum(1 for c in cat_v['CHAVE'] if self._pid(c))
        nc = sum(1 for c in cat_c['CHAVE'] if self._pid(c))
        logging.info(f"Catálogo: {len(cat_v)} chave(s) de venda e {len(cat_c)} de compra. "
                     f"Já no cadastro: {nv} de venda e {nc} de compra.")

        # 1. DESCRIÇÃO IDÊNTICA — todas as descrições que cada chave já usou, dos dois lados
        pares_v = self.nfce_saidas[['CHAVE', 'xProd_NORM']].drop_duplicates()
        pares_c = self.entradas_xls[['CHAVE', 'DESCRICAO_NORM']].drop_duplicates()
        m = pares_v.merge(pares_c, left_on='xProd_NORM', right_on='DESCRICAO_NORM', suffixes=('_V', '_C'))
        n = 0
        for desc, g in m.groupby('xProd_NORM'):
            if len(desc) < 5:
                continue  # 'OLEO', 'X' etc.: curto demais para provar que é o mesmo produto
            cv, cc = sorted(set(g['CHAVE_V'])), sorted(set(g['CHAVE_C']))
            novo = any(not self._pid(c) for c in cv + cc)
            if self._vincular(cv, cc, 'AUTO - DESCRICAO EXATA', 1.0, gravar=True):
                n += novo
        logging.info(f"Descrição idêntica: {n} casamento(s) novo(s).")

        # 2. EAN (raro na NFC-e desta base, mas forte quando existe)
        match_ean, _, _, rej_ean = self._parear_por_chave(
            cat_v, cat_c, 'cEAN', 'EAN', float(self.limiar_ean), 'EAN EXATO',
            motivo_rejeicao='EAN igual, mas descrição muito diferente (possível erro de cadastro)')
        logging.info(f"EAN: {self._aplicar_matches(match_ean)} casamento(s) novo(s).")

        # 3. PALAVRA-CHAVE DE MARCA — só entre chaves que ainda não têm produto
        match_kw, _, _ = self._match_por_palavra_chave(self._sem_produto(cat_v), self._sem_produto(cat_c))
        logging.info(f"Palavra-chave de marca: {self._aplicar_matches(match_kw)} casamento(s) novo(s).")

        # 4. SIMILARIDADE ALTA — compras sem produto contra TODAS as vendas: uma compra de
        # fornecedor novo (comum nos anos seguintes) junta ao produto que já existe.
        if self.usar_match_similaridade:
            match_sim, _, _ = self._match_por_similaridade_alta(cat_v, self._sem_produto(cat_c))
            logging.info(f"Similaridade alta: {self._aplicar_matches(match_sim)} casamento(s) novo(s).")

        if self._recusas_auto:
            logging.info(f"Trava de preço: {len(self._recusas_auto)} casamento(s) automático(s) abortado(s) "
                         f"(venda/custo fora de {self.trava_preco_min}x–{self.trava_preco_max}x) "
                         f"— seguiram para a validação manual.")

        self._montar_resumo()
        self._gerar_sugestoes(cat_v, cat_c, rej_ean)

    def _montar_resumo(self):
        """Estoque consolidado por PRODUTO_ID. Chave sem produto vira um produto sozinha (o
        texto da própria chave é o ID) e aparece como APENAS SAIDA / APENAS ENTRADA."""
        s, e = self.nfce_saidas, self.entradas_xls
        for df in (s, e):
            ids = {c: (self._pid(c) or c) for c in df['CHAVE'].unique()}
            df['PRODUTO_ID'] = df['CHAVE'].map(ids)

        r = s.groupby('PRODUTO_ID').agg(QTD_SAIDA=('QTD_SAIDA', 'sum'), VLR_TOTAL_SAIDA=('VLR_TOTAL_SAIDA', 'sum'))
        ec = e.groupby('PRODUTO_ID').agg(QTD_ENTRADA_AJUSTADA=('QTD_ENTRADA_AJUSTADA', 'sum'),
                                          VLR_TOTAL_ENTRADA=('VLR_TOTAL_ENTRADA', 'sum'),
                                          FATOR_CONV=('FATOR_CONV', 'max'),
                                          EMBALAGEM_SEM_FATOR=('EMBALAGEM_SEM_FATOR', 'max'))
        r = r.join(ec, how='outer')
        for c in ('QTD_SAIDA', 'VLR_TOTAL_SAIDA', 'QTD_ENTRADA_AJUSTADA', 'VLR_TOTAL_ENTRADA'):
            r[c] = r[c].fillna(0.0)
        r['FATOR_CONV'] = r['FATOR_CONV'].fillna(1)
        r['EMBALAGEM_SEM_FATOR'] = r['EMBALAGEM_SEM_FATOR'].eq(True)

        def chaves(df, qcol):
            g = (df.groupby(['PRODUTO_ID', 'CHAVE'])[qcol].sum().reset_index()
                   .sort_values(['PRODUTO_ID', qcol], ascending=[True, False], kind='stable'))
            todas = g.groupby('PRODUTO_ID')['CHAVE'].agg(lambda x: '|'.join(sorted(x)))
            principal = g.drop_duplicates('PRODUTO_ID').set_index('PRODUTO_ID')['CHAVE']
            return todas.reindex(r.index).fillna(''), principal.reindex(r.index).fillna('')

        r['CHAVES_VENDA'], kv = chaves(s, 'QTD_SAIDA')
        r['CHAVES_COMPRA'], kc = chaves(e, 'QTD_ENTRADA_AJUSTADA')

        def origem(chave):
            rec = self._cad.get(chave)
            if rec and rec.get('PRODUTO_ID'):
                return rec.get('ORIGEM', '')
            return 'AUTO A CONFIRMAR' if chave in self._prov else ''

        rotulos, cods, descs_c, origens, tipos = [], [], [], [], []
        for pid, cv_todas, cc_todas, v, c in zip(r.index, r['CHAVES_VENDA'], r['CHAVES_COMPRA'], kv, kc):
            rotulos.append(self._desc_chave.get(v or c, ''))
            descs_c.append(self._desc_chave.get(c, ''))
            k = v or c
            cods.append(k.split(':', 1)[1] if k.startswith(('PDV:', 'EAN:')) else k)
            ors = sorted({o for o in map(origem, [x for x in (cv_todas + '|' + cc_todas).split('|') if x]) if o})
            origens.append('; '.join(ors))
            if cv_todas and cc_todas:
                if any(o.startswith('AUTO A CONFIRMAR') for o in ors):
                    tipos.append('PAREADO - AUTO A CONFIRMAR')
                elif any(o.startswith('MANUAL') for o in ors):
                    tipos.append('PAREADO - MANUAL')
                else:
                    tipos.append('PAREADO - AUTO')
            else:
                tipos.append('APENAS SAIDA' if cv_todas else 'APENAS ENTRADA')

        r['ROTULO_PRODUTO'], r['cProd'], r['DESCRICAO_NORM'] = rotulos, cods, descs_c
        r['ORIGEM_VINCULO'], r['TIPO_MATCH'] = origens, tipos
        r['SALDO_QTD'] = r['QTD_ENTRADA_AJUSTADA'] - r['QTD_SAIDA']
        r['VLR_UNIT_ENTRADA_AJUSTADO'] = [self._custo_unit_ponderado(q, v)
                                          for q, v in zip(r['QTD_ENTRADA_AJUSTADA'], r['VLR_TOTAL_ENTRADA'])]
        r = r.reset_index()
        # Apelidos usados pelo kardex e pelo H010
        r['CHAVE_PRODUTO'] = r['PRODUTO_ID']
        r['xProd'] = r['ROTULO_PRODUTO']
        ordem = ['PRODUTO_ID', 'ROTULO_PRODUTO', 'TIPO_MATCH', 'ORIGEM_VINCULO', 'QTD_ENTRADA_AJUSTADA',
                 'QTD_SAIDA', 'SALDO_QTD', 'VLR_TOTAL_ENTRADA', 'VLR_UNIT_ENTRADA_AJUSTADO', 'VLR_TOTAL_SAIDA',
                 'FATOR_CONV', 'EMBALAGEM_SEM_FATOR', 'cProd', 'CHAVES_VENDA', 'CHAVES_COMPRA',
                 'DESCRICAO_NORM', 'xProd', 'CHAVE_PRODUTO']
        self.resumo_estoque = r[ordem]

        # Venda muito maior que a compra casada (ex: vendeu 4223, "comprou" 1) quase sempre
        # é casamento errado ou fator de embalagem faltando: vai para reconferência.
        par = r[r['TIPO_MATCH'].str.startswith('PAREADO') & (r['QTD_ENTRADA_AJUSTADA'] > 0)]
        self.matches_suspeitos_pos_validacao = par[par['QTD_SAIDA'] / par['QTD_ENTRADA_AJUSTADA'] > 15].copy()

        cont = r['TIPO_MATCH'].value_counts().to_dict()
        logging.info(f"Estoque por produto: {len(r)} produto(s) — {cont}")

    def _melhores_candidatos(self, origem, col_origem, destino, col_destino, proibido):
        """Para cada chave de `origem`, a chave de `destino` de descrição mais parecida
        (índice invertido de palavras, para não comparar todos contra todos). Pares em que
        `proibido(chave_origem, chave_destino)` é True são ignorados.
        Retorna {chave_origem: (chave_destino ou None, score)}."""
        dest = list(zip(destino['CHAVE'], destino[col_destino]))
        indice = {}
        for pos, (_, d) in enumerate(dest):
            for t in self._tokens_set(d):
                indice.setdefault(t, []).append(pos)
        limite = max(50, int(0.30 * max(len(dest), 1)))
        genericos = {t for t, lst in indice.items() if len(lst) > limite}

        # O veto de volume é freio do AUTOMÁTICO. Na sugestão ele atrapalha: 'CALABRESA 1KG'
        # (vendida por kg) x 'CALABRESA CX 4X2,5KG' é o par certo e seria zerado. Lá ele vira
        # só uma observação.
        veto, self.usar_veto_volume = self.usar_veto_volume, False
        try:
            return self._melhores_candidatos_sem_veto(origem, col_origem, dest, indice, genericos, proibido)
        finally:
            self.usar_veto_volume = veto

    def _melhores_candidatos_sem_veto(self, origem, col_origem, dest, indice, genericos, proibido):
        res = {}
        for chave, desc in zip(origem['CHAVE'], origem[col_origem]):
            toks = self._tokens_set(desc)
            cand = set()
            for t in toks:
                if t not in genericos:
                    cand.update(indice.get(t, ()))
            if not cand:  # todas as palavras eram genéricas: usa-as mesmo assim
                for t in toks:
                    cand.update(indice.get(t, ()))
            melhor, maior = None, 0.10  # mínimo de 10% só para ter algo a mostrar
            for pos in cand:
                cd, dd = dest[pos]
                if proibido(chave, cd):
                    continue
                sc = self._similaridade_produto(desc, dd)
                if sc > maior:
                    melhor, maior = cd, sc
            res[chave] = (melhor, maior if melhor else 0.0)
        return res

    def _gerar_sugestoes(self, cat_v, cat_c, rejeitados_ean):
        """Planilha de validação: uma linha por decisão pendente, com a melhor sugestão.
          AUTO A CONFIRMAR  — casou no automático com score < 98%: já vale e vem marcado Sim;
          VENDA SEM COMPRA  — chave de venda sem produto: compra mais parecida;
          COMPRA SEM VENDA  — chave de compra sem produto: venda mais parecida (ex: fornecedor novo).
        Ordenada pelo valor envolvido: as primeiras linhas são as que mais mexem no estoque."""
        vrow = cat_v.set_index('CHAVE').to_dict('index')
        crow = cat_c.set_index('CHAVE').to_dict('index')

        avisos = {}
        if rejeitados_ean is not None and not rejeitados_ean.empty:
            for r in rejeitados_ean.to_dict('records'):
                avisos.setdefault(r['xProd'], []).append(
                    f"{r['MOTIVO']} (compra em questão: '{r['DESCRICAO_NORM']}', "
                    f"similaridade {round(r['DESC_SIMILARIDADE'] * 100, 2)}%)")

        def linha(tipo, cv, cc, score, resposta='', obs_extra=None):
            v, c = vrow.get(cv, {}), crow.get(cc, {}) if cc else {}
            obs = list(obs_extra or []) + list(avisos.get(v.get('xProd'), []))
            if cc:
                nota = self._recusas_auto.get((v.get('xProd_NORM'), c.get('DESCRICAO_NORM')))
                if nota:
                    obs.append(nota)
                if score < 0.25 and tipo != 'AUTO A CONFIRMAR':
                    obs.append('SIMILARIDADE BAIXA — confira com atenção antes de aceitar')
                if self._extrair_veto_volume(v.get('xProd_NORM'), c.get('DESCRICAO_NORM')):
                    obs.append('PESO/VOLUME DIFERENTE na descrição — se for o mesmo produto, confira o fator de conversão')
                if c.get('EMBALAGEM_SEM_FATOR'):
                    obs.append('COMPRA EM EMBALAGEM (FD/CX/PCT) SEM FATOR — quantidade pode estar em fardos')
                if tipo == 'VENDA SEM COMPRA' and self._pid(cc):
                    obs.append(f"Esta compra já é do produto {self._pid(cc)}: 'Sim' junta esta venda a ele")
                if tipo == 'COMPRA SEM VENDA' and self._pid(cv):
                    obs.append(f"Esta venda já é do produto {self._pid(cv)}: 'Sim' junta esta compra a ele "
                               f"(outro fornecedor/embalagem)")
            vlr_v, vlr_c = float(v.get('VLR_TOTAL_SAIDA', 0) or 0), float(c.get('VLR_TOTAL_ENTRADA', 0) or 0)
            impacto = {'VENDA SEM COMPRA': vlr_v, 'COMPRA SEM VENDA': vlr_c}.get(tipo, max(vlr_v, vlr_c))
            return {
                'TIPO_PENDENCIA': tipo,
                'SAIDA_DESC': v.get('xProd_NORM', ''), 'ENTRADA_DESC': c.get('DESCRICAO_NORM', '') if cc else
                ('NENHUM CANDIDATO — digite a chave certa em ENTRADA_CHAVE, se souber' if tipo == 'VENDA SEM COMPRA' else ''),
                'SIMILARIDADE (%)': round(score * 100, 2),
                'ACEITAR_MATCH (Sim/Não)': resposta,
                'IMPACTO_R$': round(impacto, 2),
                'SAIDA_CHAVE': cv or '', 'SAIDA_QTD': float(v.get('QTD_SAIDA', 0) or 0),
                'PRECO_MEDIO_VENDA': round(float(v.get('VLR_UNIT_SAIDA', 0) or 0), 4),
                'ENTRADA_CHAVE': cc or '', 'ENTRADA_FORNECEDOR': c.get('FORNECEDOR', ''),
                'ENTRADA_QTD_AJUST': float(c.get('QTD_ENTRADA_AJUSTADA', 0) or 0),
                'FATOR_CONV_ENTRADA': c.get('FATOR_CONV', '') if cc else '',
                'CUSTO_MEDIO_ENTRADA': round(float(c.get('VLR_UNIT_ENTRADA_AJUSTADO', 0) or 0), 4),
                'OBSERVACAO': '; '.join(obs),
            }

        linhas = []
        for cv, cc, tipo, score in self._provisorios:
            for v in cv:
                for c in cc:
                    linhas.append(linha('AUTO A CONFIRMAR', v, c, score, 'Sim', [f"Casado automaticamente por {tipo}"]))

        proibido_vc = lambda v, c: (v, c) in self._rejeitados
        sobra_v = self._sem_produto(cat_v)
        for cv, (cc, sc) in self._melhores_candidatos(sobra_v, 'xProd_NORM', cat_c, 'DESCRICAO_NORM', proibido_vc).items():
            linhas.append(linha('VENDA SEM COMPRA', cv, cc, sc))

        sobra_c = self._sem_produto(cat_c)
        proibido_cv = lambda c, v: (v, c) in self._rejeitados
        for cc, (cv, sc) in self._melhores_candidatos(sobra_c, 'DESCRICAO_NORM', cat_v, 'xProd_NORM', proibido_cv).items():
            if cv:  # compra sem nenhuma venda parecida é só estoque parado: não pede decisão
                linhas.append(linha('COMPRA SEM VENDA', cv, cc, sc))

        ordem = ['TIPO_PENDENCIA', 'SAIDA_DESC', 'ENTRADA_DESC', 'SIMILARIDADE (%)', 'ACEITAR_MATCH (Sim/Não)',
                 'IMPACTO_R$', 'IMPACTO_ACUM_%', 'SAIDA_CHAVE', 'SAIDA_QTD', 'PRECO_MEDIO_VENDA',
                 'ENTRADA_CHAVE', 'ENTRADA_FORNECEDOR', 'ENTRADA_QTD_AJUST', 'FATOR_CONV_ENTRADA',
                 'CUSTO_MEDIO_ENTRADA', 'OBSERVACAO']
        val = pd.DataFrame(linhas)
        if val.empty:
            self.validacao_produtos = pd.DataFrame(columns=ordem)
            return
        val = val.drop_duplicates(['SAIDA_CHAVE', 'ENTRADA_CHAVE'])
        val['_ord'] = val['TIPO_PENDENCIA'].map({'AUTO A CONFIRMAR': 0, 'VENDA SEM COMPRA': 1, 'COMPRA SEM VENDA': 2})
        val = val.sort_values(['_ord', 'IMPACTO_R$'], ascending=[True, False], kind='stable').drop(columns='_ord')
        # % acumulado do valor dentro de cada tipo: mostra até onde vale a pena validar
        tot = val.groupby('TIPO_PENDENCIA')['IMPACTO_R$'].transform('sum').replace(0, 1)
        val['IMPACTO_ACUM_%'] = (val.groupby('TIPO_PENDENCIA')['IMPACTO_R$'].cumsum() / tot * 100).round(1)
        self.validacao_produtos = val[ordem].reset_index(drop=True)
        logging.info(f"Validação manual: {val['TIPO_PENDENCIA'].value_counts().to_dict()}")

    def gerar_kardex_mensal(self):
        """Monta o KARDEX (ficha de estoque) mês a mês, rolando QUANTIDADE **E VALOR**
        de um período para o próximo, atravessando anos, pelo método do
        CUSTO MÉDIO PONDERADO MÓVEL — que é o critério aceito para valorar o
        inventário do Bloco H.

        Para cada produto e cada período:

            qtd_ini, val_ini  = saldo do mês anterior
            qtd_disp = qtd_ini + entradas_qtd
            val_disp = val_ini + entradas_valor
            custo_medio = val_disp / qtd_disp
            cmv (custo da mercadoria vendida) = saidas_qtd * custo_medio
            qtd_fim = qtd_disp - saidas_qtd
            val_fim = val_disp - cmv

        Ou seja: o dinheiro que sobra no fim do mês vira o saldo inicial do mês
        seguinte, e uma compra mais cara em março reajusta o custo de abril em
        diante — o que o custo fixo da versão anterior não fazia.

        Depende de comparar_estoque() já ter rodado. Para o carry-over funcionar,
        processe TODOS os períodos de uma vez (sem filtrar --periodo).

        Preenche self.kardex, self.saldo_final_kardex e self.saldo_kardex_por_ano
        (usado para emitir um H010 por ano)."""
        logging.info("Montando kardex mensal (custo médio móvel, saldo rolando entre períodos)...")
        if self.resumo_estoque.empty:
            logging.warning("resumo_estoque vazio — kardex não gerado.")
            self.kardex = pd.DataFrame()
            return self.kardex

        rotulo_por_chave = (self.resumo_estoque.dropna(subset=['CHAVE_PRODUTO'])
                            .drop_duplicates('CHAVE_PRODUTO')
                            .set_index('CHAVE_PRODUTO')['ROTULO_PRODUTO'].to_dict())

        # COD_ITEM a exibir no H010 de cada produto (o código do PDV quando ele é vendido;
        # o EAN/chave da compra quando só houve compra).
        cod_por_chave = dict(zip(self.resumo_estoque['CHAVE_PRODUTO'], self.resumo_estoque['cProd']))

        # Custo unitário de referência: usado só quando um produto tem saída mas
        # nenhuma entrada valorada no histórico (saldo negativo sem compra).
        custo_referencia = {c: float(v) for c, v in zip(self.resumo_estoque['CHAVE_PRODUTO'],
                                                        self.resumo_estoque['VLR_UNIT_ENTRADA_AJUSTADO']) if v}

        movimentos = []  # (produto, periodo, ent_qtd, ent_vlr, sai_qtd)

        # Cada linha crua já carrega o PRODUTO_ID (dado pelo cadastro em comparar_estoque):
        # não há mais mapa por código, então código repetido entre fornecedores não mistura nada.
        # --- SAÍDAS por período (das NFC-e) ---
        if not self.nfce_saidas.empty and 'PRODUTO_ID' in self.nfce_saidas.columns:
            s = self.nfce_saidas[self.nfce_saidas['PERIODO'].astype(str) != '']
            gs = s.groupby(['PRODUTO_ID', 'PERIODO'])['QTD_SAIDA'].sum().reset_index()
            for p, per, q in zip(gs['PRODUTO_ID'], gs['PERIODO'], gs['QTD_SAIDA']):
                movimentos.append((p, per, 0.0, 0.0, float(q)))

        # --- ENTRADAS por período (das XLS), com QUANTIDADE E VALOR ---
        if not self.entradas_xls.empty and 'PRODUTO_ID' in self.entradas_xls.columns:
            e = self.entradas_xls[self.entradas_xls['PERIODO'].astype(str) != '']
            ge = e.groupby(['PRODUTO_ID', 'PERIODO'])[['QTD_ENTRADA_AJUSTADA', 'VLR_TOTAL_ENTRADA']].sum().reset_index()
            for p, per, q, v in zip(ge['PRODUTO_ID'], ge['PERIODO'], ge['QTD_ENTRADA_AJUSTADA'], ge['VLR_TOTAL_ENTRADA']):
                movimentos.append((p, per, float(q), float(v), 0.0))

        if not movimentos:
            logging.warning("Nenhum movimento com período identificado — kardex vazio. "
                            "Verifique se os nomes dos arquivos XLS têm o mês/ano (ex: 01-2021.xls).")
            self.kardex = pd.DataFrame()
            return self.kardex

        mov = pd.DataFrame(movimentos, columns=['CHAVE_PRODUTO', 'PERIODO',
                                                'ENT_QTD', 'ENT_VLR', 'SAI_QTD'])
        mov = mov.groupby(['CHAVE_PRODUTO', 'PERIODO'], as_index=False)[['ENT_QTD', 'ENT_VLR', 'SAI_QTD']].sum()

        periodos = sorted(mov['PERIODO'].unique())
        logging.info(f"Kardex cobrindo {len(periodos)} período(s): {periodos[0]} a {periodos[-1]}.")

        linhas = []
        saldo_final_por_chave = {}
        saldo_por_ano = {}   # {ano: {chave: {'QTD':..,'VLR':..,'CUSTO':..}}}

        for chave, g in mov.groupby('CHAVE_PRODUTO'):
            gmap = g.set_index('PERIODO')[['ENT_QTD', 'ENT_VLR', 'SAI_QTD']].to_dict('index')
            qtd = 0.0
            val = 0.0
            custo = float(custo_referencia.get(chave, 0.0))
            rotulo = rotulo_por_chave.get(chave, chave)

            for per in periodos:
                mv = gmap.get(per)
                if mv is None and qtd == 0 and val == 0:
                    continue  # produto ainda não existe: não polui o kardex

                ent_q = float(mv['ENT_QTD']) if mv else 0.0
                ent_v = float(mv['ENT_VLR']) if mv else 0.0
                sai_q = float(mv['SAI_QTD']) if mv else 0.0

                qtd_ini, val_ini = qtd, val

                # Disponível = saldo que veio do mês anterior + compras do mês
                qtd_disp = qtd_ini + ent_q
                val_disp = val_ini + ent_v

                # Custo médio móvel. Se não há quantidade disponível positiva
                # (produto só com venda), mantém o último custo conhecido.
                # Saldo inicial negativo (vendeu antes de comprar) não entra na média: somar
                # valor negativo com a compra dá custo negativo/absurdo, e isso rolaria para
                # os anos seguintes. Nesse caso a compra do mês define o custo.
                if qtd_ini <= 0 and ent_q > 0 and ent_v > 0:
                    custo = ent_v / ent_q
                    val_disp = qtd_disp * custo
                elif qtd_disp > 0:
                    custo = val_disp / qtd_disp
                elif ent_q > 0:
                    custo = ent_v / ent_q

                cmv = sai_q * custo
                qtd = qtd_disp - sai_q
                val = val_disp - cmv

                # Saldo negativo: valoriza pelo custo corrente para o valor não
                # descolar da quantidade (fica negativo também, sinalizando o furo).
                if qtd <= 0:
                    val = qtd * custo

                linhas.append({
                    'PRODUTO_ID': chave,
                    'PRODUTO': rotulo,
                    'PERIODO': per,
                    'SALDO_INI_QTD': round(qtd_ini, 3),
                    'SALDO_INI_VLR': round(val_ini, 2),
                    'ENTRADAS_QTD': round(ent_q, 3),
                    'ENTRADAS_VLR': round(ent_v, 2),
                    'SAIDAS_QTD': round(sai_q, 3),
                    'CMV_SAIDAS_VLR': round(cmv, 2),
                    'CUSTO_MEDIO': round(custo, 4),
                    'SALDO_FIM_QTD': round(qtd, 3),
                    'SALDO_FIM_VLR': round(val, 2),
                })

                # Fotografa o saldo de dezembro de cada ano para o H010 anual.
                ano, mes = per.split('-')
                if mes == '12':
                    saldo_por_ano.setdefault(int(ano), {})[chave] = {
                        'QTD': qtd, 'VLR': val, 'CUSTO': custo,
                        'PRODUTO': rotulo, 'COD_ITEM': cod_por_chave.get(chave, ''),
                    }

            saldo_final_por_chave[chave] = {'QTD': qtd, 'VLR': val, 'CUSTO': custo}

        self.kardex = pd.DataFrame(linhas)
        self.saldo_final_kardex = saldo_final_por_chave

        # Se o último período do lote não for dezembro, ainda assim guarda o saldo
        # final daquele ano (inventário parcial) para não ficar sem H010.
        if periodos:
            ultimo_ano, ultimo_mes = periodos[-1].split('-')
            if ultimo_mes != '12':
                snap = {}
                for chave, s in saldo_final_por_chave.items():
                    snap[chave] = dict(s, PRODUTO=rotulo_por_chave.get(chave, chave),
                                       COD_ITEM=cod_por_chave.get(chave, ''))
                saldo_por_ano.setdefault(int(ultimo_ano), snap)
                logging.warning(f"O lote termina em {periodos[-1]} (não é dezembro). O inventário de "
                                f"{ultimo_ano} sairá com o saldo dessa data, não de 31/12.")

        self.saldo_kardex_por_ano = saldo_por_ano
        logging.info(f"Kardex gerado: {len(self.kardex)} linhas de movimento, "
                     f"{len(saldo_final_por_chave)} produtos, "
                     f"inventário disponível para: {sorted(saldo_por_ano)}.")
        return self.kardex

    def _detectar_ano_inventario(self):
        """Detecta o ano de referência a partir das datas de emissão das NFC-e extraídas,
        para não depender do usuário lembrar de passar --ano toda vez. Se houver notas de
        mais de um ano no mesmo lote, avisa (pode indicar mistura indevida de períodos)."""
        if self.nfce_saidas.empty or 'Data_Emissao' not in self.nfce_saidas.columns:
            return None
        anos = pd.to_datetime(self.nfce_saidas['Data_Emissao'], errors='coerce').dt.year.dropna()
        if anos.empty:
            return None
        contagem = anos.value_counts()
        ano_predominante = int(contagem.idxmax())
        if len(contagem) > 1:
            logging.warning(
                f"As NFC-e processadas misturam mais de um ano ({dict(contagem)}). "
                f"Usando o ano predominante ({ano_predominante}) como data do inventário. "
                f"Se isso não fizer sentido, informe --ano explicitamente."
            )
        return ano_predominante

    def gerar_inventarios_por_ano(self, desconsiderar_saidas_sem_entrada=False, mot_inv='01'):
        """Emite um H010 para CADA ano coberto pelo kardex, usando o saldo de 31/12
        daquele ano — quantidade e valor já rolados pelo custo médio móvel.

        É assim que 2021 → 2022 → ... → 2025 se encadeiam: o estoque que sobra em
        31/12/2021 (com o dinheiro que ele vale) é o saldo inicial de janeiro/2022.
        Rode o script UMA vez com todos os anos juntos, não um ano por vez."""
        anos = sorted(getattr(self, 'saldo_kardex_por_ano', {}) or {})
        if not anos:
            logging.warning("Sem kardex por ano — emitindo inventário único pelo saldo agregado.")
            return [self.gerar_inventario_efd_h010(desconsiderar_saidas_sem_entrada, mot_inv=mot_inv)]
        gerados = []
        for a in anos:
            gerados.append(self.gerar_inventario_efd_h010(
                desconsiderar_saidas_sem_entrada, ano=a, mot_inv=mot_inv))
        logging.info(f"Inventários emitidos para os anos: {anos}")
        return gerados

    def gerar_inventario_efd_h010(self, desconsiderar_saidas_sem_entrada=False, ano=None, mot_inv='01'):
        logging.info("Gerando planilhas de inventário EFD ICMS/IPI (Registro H010)...")
        if self.resumo_estoque.empty: return

        if ano is None:
            ano = self._detectar_ano_inventario()
        if ano:
            logging.info(f"Data do inventário (H005/DT_INV) definida como 31/12/{ano}.")
            dt_inv_exibicao = f"31/12/{ano}"
            dt_inv_sped = f"31122{ano}" if len(str(ano)) == 3 else f"3112{ano}"  # ddmmaaaa (ex: 31122021)
        else:
            logging.warning(
                "Não foi possível detectar o ano das NFC-e para preencher a data do inventário. "
                "Informe --ano AAAA na linha de comando. DT_INV ficará em branco."
            )
            dt_inv_exibicao, dt_inv_sped = '', ''

        # Preferência: saldo de 31/12 do ano vindo do KARDEX (quantidade e valor já
        # acumulados pelo custo médio móvel). O saldo agregado do resumo_estoque só é
        # usado como fallback quando o kardex não cobriu o período — ele ignora a
        # ordem cronológica e não sabe encadear um ano no outro.
        snapshot = (getattr(self, 'saldo_kardex_por_ano', {}) or {}).get(ano)
        if snapshot:
            entradas_por_chave = {}
            if 'CHAVE_PRODUTO' in self.resumo_estoque.columns:
                entradas_por_chave = (self.resumo_estoque
                                      .groupby('CHAVE_PRODUTO')['QTD_ENTRADA_AJUSTADA']
                                      .sum().to_dict())
            df = pd.DataFrame([{
                'cProd': dados.get('COD_ITEM', ''),
                'xProd': dados.get('PRODUTO', chave),
                'CHAVE_PRODUTO': chave,
                'SALDO_QTD': float(dados.get('QTD', 0) or 0),
                'VLR_UNIT_ENTRADA_AJUSTADO': float(dados.get('CUSTO', 0) or 0),
                'QTD_ENTRADA_AJUSTADA': float(entradas_por_chave.get(chave, 0) or 0),
                'VL_ITEM': float(dados.get('VLR', 0) or 0),
            } for chave, dados in snapshot.items()])
            logging.info(f"Inventário de {ano} montado a partir do kardex "
                         f"({len(df)} produtos com saldo em 31/12/{ano}).")
        else:
            df = self.resumo_estoque.copy()
            df['SALDO_QTD'] = df['SALDO_QTD'].fillna(0)
            df['VLR_UNIT_ENTRADA_AJUSTADO'] = df['VLR_UNIT_ENTRADA_AJUSTADO'].fillna(0)
            df['QTD_ENTRADA_AJUSTADA'] = df['QTD_ENTRADA_AJUSTADA'].fillna(0)
            df['VL_ITEM'] = df['SALDO_QTD'] * df['VLR_UNIT_ENTRADA_AJUSTADO']
        
        df_positivos = df[df['SALDO_QTD'] > 0].copy()
        df_negativos = df[df['SALDO_QTD'] < 0].copy()
        
        if desconsiderar_saidas_sem_entrada:
            logging.info("REGRA: Desconsiderando estoque negativo proveniente do ano anterior.")
            df_negativos = df_negativos[df_negativos['QTD_ENTRADA_AJUSTADA'] > 0]
        
        def format_h010(data_subset):
            if data_subset.empty: return pd.DataFrame()
            # BUG CORRIGIDO: atribuir um escalar a um DataFrame VAZIO não cria linha
            # nenhuma — era por isso que a coluna REG saía toda em branco no arquivo.
            # Agora o DF nasce com o índice do subconjunto.
            h010 = pd.DataFrame(index=data_subset.index)
            h010['REG'] = 'H010'
            h010['COD_ITEM'] = data_subset['cProd'] 
            h010['UNID'] = 'UN' 
            h010['QTD'] = data_subset['SALDO_QTD'].abs() 
            h010['VL_UNIT'] = data_subset['VLR_UNIT_ENTRADA_AJUSTADO'].round(4)
            h010['VL_ITEM'] = data_subset['VL_ITEM'].abs().round(2)
            h010['IND_PROP'] = '0' 
            h010['COD_PART'] = ''
            h010['TXT_COMPL'] = data_subset['xProd']
            h010['COD_CTA'] = ''
            h010['VL_ITEM_IR'] = ''
            h010['DT_INV'] = dt_inv_exibicao  # data do inventário (último dia do ano) - referência
            return h010

        h005_df = pd.DataFrame([{
            'REG': 'H005',
            'DT_INV': dt_inv_sped,
            'VL_INV': round(df_positivos['VL_ITEM'].abs().sum(), 2) if not df_positivos.empty else 0,
            'MOT_INV': mot_inv,
        }])
        h010_pos = format_h010(df_positivos)
        h010_neg = format_h010(df_negativos)

        sufixo = f"_{ano}" if ano else ""
        outfile = os.path.join(self.output_dir, f"Inventario_EFD_ICMS_IPI_Saldos{sufixo}.xlsx")
        self._exportar_inventario_formatado(outfile, h005_df, h010_pos, h010_neg,
                                            df_positivos, df_negativos, dt_inv_sped)
        logging.info("Inventário H010 salvo com sucesso (com dashboard e formatação)!")
        return outfile

    def _exportar_inventario_formatado(self, outfile, h005_df, h010_pos, h010_neg,
                                       df_positivos, df_negativos, dt_inv_sped):
        MONEY2 = 'R$ #,##0.00'; MONEY4 = 'R$ #,##0.0000'; NUM = '#,##0.###'
        vl_inv = float(df_positivos['VL_ITEM'].abs().sum()) if not df_positivos.empty else 0.0
        qtd_pos = float(df_positivos['SALDO_QTD'].abs().sum()) if not df_positivos.empty else 0.0
        n_pos, n_neg = len(h010_pos), len(h010_neg)
        qtd_neg = float(df_negativos['SALDO_QTD'].abs().sum()) if not df_negativos.empty else 0.0

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resumo"
        ws.sheet_view.showGridLines = False

        ws.merge_cells('B2:H2')
        ws['B2'].value = "INVENTÁRIO DE ESTOQUE — EFD ICMS/IPI (Bloco H)"
        ws['B2'].font = Font(name=FONTE_PADRAO, size=18, bold=True, color=COR_BRANCO)
        ws['B2'].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in "BCDEFGH":
            ws[f'{col}2'].fill = PatternFill("solid", fgColor=COR_AZUL_ESCURO)
        ws.row_dimensions[2].height = 40

        dt = str(dt_inv_sped)
        data_fmt = f"{dt[:2]}/{dt[2:4]}/{dt[4:]}" if len(dt) == 8 else dt
        ws.merge_cells('B3:H3')
        ws['B3'].value = f"Data do inventário: {data_fmt}     •     MOT_INV: 01 — Final do período"
        ws['B3'].font = Font(name=FONTE_PADRAO, size=10, italic=True, color="5A6B7B")
        ws['B3'].alignment = Alignment(horizontal="left", indent=1)

        self._card(ws, 'B5', "VALOR TOTAL DO INVENTÁRIO", f"R$ {vl_inv:,.2f}", COR_TEAL,
                   "Soma dos saldos positivos (H005/VL_INV)")
        self._card(ws, 'D5', "ITENS COM SALDO POSITIVO", _fmt_num(n_pos), COR_AZUL_MEDIO,
                   f"{qtd_pos:,.0f} unidades em estoque".replace(",", "."))
        self._card(ws, 'F5', "ITENS COM SALDO NEGATIVO", _fmt_num(n_neg), COR_VERMELHO,
                   f"{qtd_neg:,.0f} un — pendências a revisar".replace(",", "."))

        ws.merge_cells('B9:H12')
        ws['B9'].value = ("COMO LER: o 'Valor do Inventário' considera apenas saldos positivos — o estoque real declarável. "
                          "Saldos negativos indicam itens vendidos sem entrada correspondente no período (estoque de anos "
                          "anteriores ou compras lançadas com nome/código diferente ainda não pareadas). Revise-os antes de fechar o SPED.")
        ws['B9'].font = Font(name=FONTE_PADRAO, size=9, color="42525F")
        ws['B9'].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        for r in range(9, 13):
            for col in "BCDEFGH":
                ws[f'{col}{r}'].fill = PatternFill("solid", fgColor=COR_CINZA_CLARO)

        # Top 8 por valor + gráfico
        if not h010_pos.empty:
            top = h010_pos.nlargest(8, 'VL_ITEM')[['TXT_COMPL', 'VL_ITEM']].copy()
            top['TXT_COMPL'] = top['TXT_COMPL'].astype(str).str.slice(0, 28)
            start = 15
            ws.cell(row=start - 1, column=2, value="TOP 8 ITENS POR VALOR EM ESTOQUE").font = Font(name=FONTE_PADRAO, size=11, bold=True, color=COR_AZUL_ESCURO)
            ws.cell(row=start, column=2, value="Item").font = Font(name=FONTE_PADRAO, size=9, bold=True)
            ws.cell(row=start, column=3, value="Valor (R$)").font = Font(name=FONTE_PADRAO, size=9, bold=True)
            for i, (_, r) in enumerate(top.iterrows(), start=start + 1):
                ws.cell(row=i, column=2, value=r['TXT_COMPL']).font = Font(name=FONTE_PADRAO, size=9)
                vc = ws.cell(row=i, column=3, value=round(float(r['VL_ITEM']), 2))
                vc.number_format = MONEY2; vc.font = Font(name=FONTE_PADRAO, size=9)
            chart = BarChart(); chart.type = "bar"; chart.title = "Top 8 itens por valor (R$)"
            chart.height = 7.5; chart.width = 18
            data = Reference(ws, min_col=3, min_row=start, max_row=start + len(top))
            cats = Reference(ws, min_col=2, min_row=start + 1, max_row=start + len(top))
            chart.add_data(data, titles_from_data=True); chart.set_categories(cats); chart.legend = None
            try: chart.series[0].graphicalProperties.solidFill = COR_TEAL
            except Exception: pass
            ws.add_chart(chart, "E15")

        for idx, w in {1: 3, 2: 20, 3: 16, 4: 20, 5: 16, 6: 20, 7: 16, 8: 16}.items():
            ws.column_dimensions[get_column_letter(idx)].width = w
        ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_area = 'A1:H40'

        # H005
        wsh, _ = escrever_aba_formatada(wb, "H005 (Cabecalho)", h005_df, COR_AZUL_ESCURO,
                                        formatos={'VL_INV': MONEY2})
        # H010 positivos / negativos
        escrever_aba_formatada(wb, "Saldos Positivos (H010)", h010_pos, COR_TEAL,
                               formatos={'QTD': NUM, 'VL_UNIT': MONEY4, 'VL_ITEM': MONEY2},
                               larguras_fixas={'TXT_COMPL': 42})
        escrever_aba_formatada(wb, "Saldos Negativos (H010)", h010_neg, COR_VERMELHO,
                               formatos={'QTD': NUM, 'VL_UNIT': MONEY4, 'VL_ITEM': MONEY2},
                               larguras_fixas={'TXT_COMPL': 42})

        # Kardex mensal (se gerado; a tela web baixa o kardex em CSV à parte)
        if self.exportar_abas_pesadas and getattr(self, 'kardex', None) is not None and not self.kardex.empty:
            escrever_aba_formatada(wb, "Kardex Mensal", self.kardex, COR_ROXO,
                                   formatos=FORMATOS_KARDEX, colorscale_col='SALDO_FIM_QTD',
                                   larguras_fixas={'PRODUTO': 40, 'PERIODO': 12})

        wb.move_sheet("Resumo", -(len(wb.sheetnames) - 1))
        wb.save(outfile)

    @staticmethod
    def _card(ws, ancora, titulo, valor, cor, sub=""):
        col = ws[ancora].column; row = ws[ancora].row
        c1 = get_column_letter(col); c2 = get_column_letter(col + 1)
        for dr in range(3):
            ws.merge_cells(f'{c1}{row+dr}:{c2}{row+dr}')
        ws[f'{c1}{row}'].value = titulo
        ws[f'{c1}{row}'].font = Font(name=FONTE_PADRAO, size=9, bold=True, color=COR_BRANCO)
        ws[f'{c1}{row}'].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws[f'{c1}{row+1}'].value = valor
        ws[f'{c1}{row+1}'].font = Font(name=FONTE_PADRAO, size=17, bold=True, color=COR_BRANCO)
        ws[f'{c1}{row+1}'].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws[f'{c1}{row+2}'].value = sub
        ws[f'{c1}{row+2}'].font = Font(name=FONTE_PADRAO, size=8, color=COR_BRANCO)
        ws[f'{c1}{row+2}'].alignment = Alignment(horizontal="left", vertical="top", indent=1, wrap_text=True)
        for rr in (row, row + 1, row + 2):
            for cc in (col, col + 1):
                ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=cor)
        ws.row_dimensions[row].height = 18
        ws.row_dimensions[row + 1].height = 26
        ws.row_dimensions[row + 2].height = 24

    def aplicar_validacao_manual(self, arquivo_validacao_preenchido):
        """Lê a planilha de validação devolvida e grava as respostas no CADASTRO:

          Sim   -> as duas chaves da linha viram o mesmo produto (MANUAL, vale para sempre,
                   inclusive nos próximos anos);
          Não   -> o par vai para os rejeitados: nunca mais é sugerido nem casado no automático
                   (na próxima rodada aparece a 2ª melhor sugestão);
          vazio -> nada muda.

        Para corrigir uma sugestão errada: troque o texto de ENTRADA_CHAVE pela chave certa
        (copiada de outra linha ou da aba Cadastro_Produtos) e responda Sim.
        Também lê a coluna FATOR_CORRETO da aba Revisar_Fatores_Conversao.

        Chame ANTES de comparar_estoque(): o cruzamento já sai com as respostas incorporadas."""
        logging.info(f"Lendo validações manuais em: {arquivo_validacao_preenchido}")
        try:
            abas = pd.read_excel(arquivo_validacao_preenchido, sheet_name=None, dtype=str)
        except Exception as e:
            logging.error(f"Não foi possível ler {arquivo_validacao_preenchido}: {e}")
            return

        val = abas.get('Validacao_Manual')
        col = 'ACEITAR_MATCH (Sim/Não)'
        if val is None or col not in val.columns or 'SAIDA_CHAVE' not in val.columns:
            logging.error("Aba 'Validacao_Manual' (com SAIDA_CHAVE/ENTRADA_CHAVE) não encontrada. "
                          "A planilha é de uma versão antiga do script? Gere uma nova e responda nela.")
            return

        rev = abas.get('Revisar_Fatores_Conversao')
        fatores = rev.to_dict('records') if rev is not None and {'CHAVE', 'FATOR_CORRETO'}.issubset(rev.columns) else []
        return self.aplicar_respostas(val.to_dict('records'), fatores)

    def aplicar_respostas(self, respostas, fatores=()):
        """Grava no cadastro as respostas da validação — vindas da planilha Excel ou da tela
        (index.html). `respostas`: dicts com SAIDA_CHAVE, ENTRADA_CHAVE, SAIDA_DESC, ENTRADA_DESC,
        SIMILARIDADE (%) e 'ACEITAR_MATCH (Sim/Não)'. `fatores`: dicts com CHAVE, DESCRICAO e
        FATOR_CORRETO. Retorna {'sim': n, 'nao': n, 'fatores': n, 'invalidas': n}."""
        col = 'ACEITAR_MATCH (Sim/Não)'
        txt = lambda x: '' if x is None or (isinstance(x, float) and pd.isna(x)) or str(x) == 'nan' else str(x).strip()
        hoje = date.today().isoformat()
        sim = nao = invalidas = 0
        for r in respostas:
            resp = txt(r.get(col)).upper()[:1]
            if resp not in ('S', 'N'):
                continue
            cv, cc = txt(r.get('SAIDA_CHAVE')), txt(r.get('ENTRADA_CHAVE'))
            if not cv.startswith('PDV:') or not cc or cc.startswith('PDV:'):
                invalidas += 1
                continue
            dv, dc = txt(r.get('SAIDA_DESC')), txt(r.get('ENTRADA_DESC'))
            if resp == 'S':
                self._rejeitados.pop((cv, cc), None)
                if self._vincular([cv], [cc], 'MANUAL', self._num(r.get('SIMILARIDADE (%)')) / 100,
                                  gravar=True, descricoes={cv: dv, cc: dc}):
                    sim += 1
            else:
                self._rejeitados[(cv, cc)] = {'CHAVE_VENDA': cv, 'CHAVE_COMPRA': cc, 'DESCRICAO_VENDA': dv,
                                              'DESCRICAO_COMPRA': dc, 'DATA': hoje}
                nao += 1

        # Fatores de conversão confirmados na aba de revisão
        n_fatores = 0
        for r in fatores:
            f = self._num(r.get('FATOR_CORRETO'))
            if f <= 0:
                continue
            for c in [x.strip() for x in txt(r.get('CHAVE')).split('|') if x.strip()]:
                rec = self._cad.setdefault(c, {'CHAVE': c, 'LADO': 'C', 'PRODUTO_ID': '', 'ORIGEM': 'FATOR MANUAL',
                                               'DESCRICAO': txt(r.get('DESCRICAO')), 'SCORE': '', 'DATA': hoje})
                rec['FATOR_CONV'] = f"{f:g}"
                n_fatores += 1

        self.salvar_cadastro()
        logging.info(f"Validação incorporada ao cadastro: {sim} 'Sim', {nao} 'Não', {n_fatores} fator(es) de conversão.")
        if invalidas:
            logging.warning(f"{invalidas} linha(s) respondida(s) sem SAIDA_CHAVE (PDV:...) e ENTRADA_CHAVE válidas foram ignoradas.")
        return {'sim': sim, 'nao': nao, 'fatores': n_fatores, 'invalidas': invalidas}

    def exportar_planilha_validacao(self):
        logging.info("Exportando Planilha de Validação...")
        if self.resumo_estoque.empty:
            logging.warning("resumo_estoque vazio — planilha de validação não gerada.")
            return
        outfile = os.path.join(self.output_dir, "Analise_NFCE_Entradas_Saidas_VALIDACAO.xlsx")
        self._exportar_validacao_formatada(outfile, self.resumo_estoque.copy(), self.validacao_produtos.copy(),
                                           self.tabela_revisao_fatores(), self.entradas_xls.copy())
        logging.info("Planilha de Validação formatada e salva com sucesso (com dashboard)!")
        return outfile

    def tabela_revisao_fatores(self):
        """Chaves de compra cujo fator de conversão precisa de conferência (aba
        Revisar_Fatores_Conversao da planilha e aba 'Fatores' da tela)."""
        # Aba de revisão de fatores de conversão — UMA LINHA POR CHAVE DE COMPRA. O fator
        # digitado em FATOR_CORRETO vai para o cadastro e vale para todas as notas dessa
        # chave, em todos os anos:
        # (a1) fator deduzido de número COLADO na descrição: acertou na maioria, mas pode
        #      ter pego bitola/modelo (ex: 'BROCA 5,0MM IRC10' -> 10 é caixa com 10?);
        # (a2) unidade de embalagem (FD/CX/PCT...) sem número de unidades -> fator ficou 1;
        # (b)  produto pareado com saldo negativo e fator 1 -> sintoma clássico de compra
        #      em embalagem vendida por unidade (entrada 5 fardos, saída 100 unidades).
        e = self.entradas_xls
        colunas_rev = ['CHAVE', 'DESCRICAO', 'FORNECEDOR', 'UNID', 'QTD_ENTRADA', 'FATOR_CONV',
                       'QTD_ENTRADA_AJUSTADA', 'MOTIVO_REVISAO', 'FATOR_CORRETO']
        if 'UNID' not in e.columns:
            e['UNID'] = ''

        def por_chave(mask, motivo):
            if not mask.any():
                return None
            g = e[mask].groupby('CHAVE').agg(
                DESCRICAO=('DESCRICAO_NORM', 'first'), FORNECEDOR=('RAZAO_EMITENTE', 'first'),
                UNID=('UNID', 'first'), QTD_ENTRADA=('QTD_ENTRADA', 'sum'), FATOR_CONV=('FATOR_CONV', 'max'),
                QTD_ENTRADA_AJUSTADA=('QTD_ENTRADA_AJUSTADA', 'sum')).reset_index()
            g['MOTIVO_REVISAO'] = [motivo(f) for f in g['FATOR_CONV']]
            return g

        frames_revisao = [
            por_chave(e.get('FATOR_REVISAR', pd.Series(False, index=e.index)).astype(bool),
                      lambda f: f"Fator {f:g} deduzido de número COLADO na descrição — confirme se é embalagem e não bitola/modelo"),
            por_chave(e.get('EMBALAGEM_SEM_FATOR', pd.Series(False, index=e.index)).astype(bool),
                      lambda f: 'Unidade de embalagem (FD/CX/PCT/SC) sem número de unidades na descrição — fator ficou 1'),
        ]
        r = self.resumo_estoque
        suspeitos = r[r['TIPO_MATCH'].str.startswith('PAREADO') & (r['SALDO_QTD'] < 0) &
                      (pd.to_numeric(r['FATOR_CONV'], errors='coerce').fillna(1) == 1) &
                      (r['QTD_ENTRADA_AJUSTADA'] > 0)]
        if not suspeitos.empty:
            frames_revisao.append(pd.DataFrame({
                'CHAVE': suspeitos['CHAVES_COMPRA'], 'DESCRICAO': suspeitos['DESCRICAO_NORM'], 'FORNECEDOR': '',
                'UNID': '', 'QTD_ENTRADA': suspeitos['QTD_ENTRADA_AJUSTADA'], 'FATOR_CONV': 1,
                'QTD_ENTRADA_AJUSTADA': suspeitos['QTD_ENTRADA_AJUSTADA'],
                'MOTIVO_REVISAO': 'Pareado com saldo NEGATIVO e fator=1: possível compra em embalagem vendida por '
                                  'unidade (informe o fator, ex: fardo c/20)'}))
        frames_revisao = [f for f in frames_revisao if f is not None and not f.empty]
        df_revisao_fatores = (pd.concat(frames_revisao, ignore_index=True).drop_duplicates('CHAVE')
                              if frames_revisao else pd.DataFrame(columns=colunas_rev))
        df_revisao_fatores['FATOR_CORRETO'] = ''
        return df_revisao_fatores[colunas_rev].reset_index(drop=True)

    def _exportar_validacao_formatada(self, outfile, est, val, rev, base):
        MONEY2 = 'R$ #,##0.00'; MONEY4 = 'R$ #,##0.0000'; NUM = '#,##0.###'
        n_pareados = int(est['TIPO_MATCH'].str.startswith('PAREADO').sum())
        n_so_entrada = int((est['TIPO_MATCH'] == 'APENAS ENTRADA').sum()) if 'TIPO_MATCH' in est.columns else 0
        n_so_saida = int((est['TIPO_MATCH'] == 'APENAS SAIDA').sum()) if 'TIPO_MATCH' in est.columns else 0
        n_val = len(val)
        n_conf = int((val['TIPO_PENDENCIA'] == 'AUTO A CONFIRMAR').sum())
        vendas = val[val['TIPO_PENDENCIA'] == 'VENDA SEM COMPRA']
        n_80 = int((vendas['IMPACTO_ACUM_%'] <= 80).sum()) + (1 if len(vendas) else 0)
        n_rev = len(rev)
        dist = est['TIPO_MATCH'].value_counts() if 'TIPO_MATCH' in est.columns else pd.Series(dtype=int)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resumo"
        ws.sheet_view.showGridLines = False

        ws.merge_cells('B2:I2')
        ws['B2'].value = "ANÁLISE DE ESTOQUE — ENTRADAS × SAÍDAS (NFC-e)"
        ws['B2'].font = Font(name=FONTE_PADRAO, size=18, bold=True, color=COR_BRANCO)
        ws['B2'].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in "BCDEFGHI":
            ws[f'{col}2'].fill = PatternFill("solid", fgColor=COR_AZUL_ESCURO)
        ws.row_dimensions[2].height = 40

        ws.merge_cells('B3:I3')
        ws['B3'].value = ("Cadastro de produtos: memória → descrição idêntica → EAN → palavra-chave → similaridade alta. "
                          "O restante vai para validação, ordenado pelo valor.")
        ws['B3'].font = Font(name=FONTE_PADRAO, size=10, italic=True, color="5A6B7B")
        ws['B3'].alignment = Alignment(horizontal="left", indent=1)

        self._card(ws, 'B5', "PRODUTOS PAREADOS", _fmt_num(n_pareados), COR_TEAL, "Venda e compra no mesmo PRODUTO_ID")
        self._card(ws, 'D5', "AGUARDANDO VALIDAÇÃO", _fmt_num(n_val), COR_AZUL_MEDIO, f"{_fmt_num(n_conf)} automáticos a confirmar")
        self._card(ws, 'F5', "SÓ ENTRADA (compra)", _fmt_num(n_so_entrada), COR_ROXO, "Compras sem venda casada")
        self._card(ws, 'H5', "SÓ SAÍDA (venda)", _fmt_num(n_so_saida), COR_VERMELHO, "Vendas sem entrada — gera saldo negativo")

        ws.merge_cells('B9:I11')
        ws['B9'].value = (f"AÇÃO RECOMENDADA:  na aba 'Validacao_Manual', responda Sim/Não na coluna amarela. Comece pelas "
                          f"{_fmt_num(n_conf)} linhas AUTO A CONFIRMAR (já vêm 'Sim'; troque para 'Não' se estiver errado) e "
                          f"pelas primeiras {_fmt_num(n_80)} VENDA SEM COMPRA, que somam 80% do valor vendido sem par. Depois rode "
                          f"o script com --validacao_preenchida: as respostas vão para o cadastro e valem para todos os anos. "
                          f"{_fmt_num(n_rev)} chaves de compra pedem conferência de fator (aba 'Revisar_Fatores_Conversao').")
        ws['B9'].font = Font(name=FONTE_PADRAO, size=10, color="7A5C00")
        ws['B9'].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        for r in range(9, 12):
            for col in "BCDEFGHI":
                ws[f'{col}{r}'].fill = PatternFill("solid", fgColor="FEF7DC")

        # Tabela + gráfico de distribuição
        if not dist.empty:
            start = 14
            ws.cell(row=start - 1, column=2, value="COMO OS ITENS FORAM CLASSIFICADOS").font = Font(name=FONTE_PADRAO, size=11, bold=True, color=COR_AZUL_ESCURO)
            ws.cell(row=start, column=2, value="Tipo de classificação").font = Font(name=FONTE_PADRAO, size=9, bold=True)
            ws.cell(row=start, column=3, value="Qtd. itens").font = Font(name=FONTE_PADRAO, size=9, bold=True)
            ordem_tipo = ['PAREADO - AUTO', 'PAREADO - MANUAL', 'PAREADO - AUTO A CONFIRMAR', 'APENAS ENTRADA', 'APENAS SAIDA']
            rr = start + 1
            for tp in ordem_tipo:
                if tp in dist.index:
                    ws.cell(row=rr, column=2, value=tp).font = Font(name=FONTE_PADRAO, size=9)
                    ws.cell(row=rr, column=3, value=int(dist[tp])).font = Font(name=FONTE_PADRAO, size=9)
                    rr += 1
            chart = BarChart(); chart.type = "bar"; chart.title = "Itens por tipo de classificação"
            chart.height = 7; chart.width = 17
            data = Reference(ws, min_col=3, min_row=start, max_row=rr - 1)
            cats = Reference(ws, min_col=2, min_row=start + 1, max_row=rr - 1)
            chart.add_data(data, titles_from_data=True); chart.set_categories(cats); chart.legend = None
            try: chart.series[0].graphicalProperties.solidFill = COR_AZUL_MEDIO
            except Exception: pass
            ws.add_chart(chart, "E14")

        for idx, w in {1: 3, 2: 30, 3: 14, 4: 20, 5: 14, 6: 20, 7: 12, 8: 20, 9: 14}.items():
            ws.column_dimensions[get_column_letter(idx)].width = w
        ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_area = 'A1:I34'

        # Abas de dados — os NOMES 'Validacao_Manual' e 'Revisar_Fatores_Conversao' (e suas
        # colunas de chave) são o que aplicar_validacao_manual lê de volta.
        est = est.drop(columns=['xProd', 'CHAVE_PRODUTO'], errors='ignore').rename(columns={'cProd': 'COD_ITEM_H010'})
        escrever_aba_formatada(wb, "Estoque_por_Produto", est, COR_AZUL_ESCURO,
                               formatos={'QTD_SAIDA': NUM, 'QTD_ENTRADA_AJUSTADA': NUM, 'VLR_TOTAL_ENTRADA': MONEY2,
                                         'VLR_TOTAL_SAIDA': MONEY2, 'VLR_UNIT_ENTRADA_AJUSTADO': MONEY4, 'SALDO_QTD': NUM},
                               colorscale_col='SALDO_QTD',
                               larguras_fixas={'ROTULO_PRODUTO': 34, 'DESCRICAO_NORM': 34, 'ORIGEM_VINCULO': 30,
                                               'CHAVES_VENDA': 24, 'CHAVES_COMPRA': 30})

        ws_val, cols_val = escrever_aba_formatada(wb, "Validacao_Manual", val, COR_AZUL_MEDIO,
                               formatos={'SAIDA_QTD': NUM, 'ENTRADA_QTD_AJUST': NUM, 'CUSTO_MEDIO_ENTRADA': MONEY4,
                                         'PRECO_MEDIO_VENDA': MONEY4, 'IMPACTO_R$': MONEY2, 'IMPACTO_ACUM_%': '0.0',
                                         'FATOR_CONV_ENTRADA': '0', 'SIMILARIDADE (%)': '0.00'},
                               col_amarela='ACEITAR_MATCH (Sim/Não)', databar_col='SIMILARIDADE (%)',
                               larguras_fixas={'TIPO_PENDENCIA': 20, 'SAIDA_DESC': 38, 'ENTRADA_DESC': 40,
                                               'SIMILARIDADE (%)': 15, 'ACEITAR_MATCH (Sim/Não)': 16,
                                               'SAIDA_CHAVE': 14, 'ENTRADA_CHAVE': 30, 'ENTRADA_FORNECEDOR': 30,
                                               'OBSERVACAO': 50})
        # Congela as 5 primeiras colunas (as de decisão) além do cabeçalho, para elas
        # ficarem visíveis mesmo rolando para os dados de apoio à direita.
        ws_val.freeze_panes = 'F2'
        _validacao_lista_suspensa(ws_val, cols_val, ws_val.max_row)
        _adicionar_colunas_formula_validacao(ws_val, cols_val, ws_val.max_row)
        # Pintura da linha inteira DEPOIS de criar as colunas de fórmula, para o
        # realce cobrir também SALDO_SE_RESPOSTA e STATUS_AUTOMATICO.
        _pintar_linhas_por_resposta(ws_val, cols_val, ws_val.max_row, ws_val.max_column)

        escrever_aba_formatada(wb, "Revisar_Fatores_Conversao", rev, COR_VERMELHO,
                               formatos={'QTD_ENTRADA': NUM, 'FATOR_CONV': '0', 'QTD_ENTRADA_AJUSTADA': NUM},
                               col_amarela='FATOR_CORRETO',
                               larguras_fixas={'CHAVE': 30, 'DESCRICAO': 40, 'FORNECEDOR': 30, 'MOTIVO_REVISAO': 50,
                                               'FATOR_CORRETO': 16})

        # Cadastro inteiro (gravado + a confirmar): é aqui que se confere a qual produto
        # cada chave pertence e de onde se copia uma chave para corrigir uma sugestão.
        escrever_aba_formatada(wb, "Cadastro_Produtos", self.cadastro_como_df(), COR_TEAL,
                               larguras_fixas={'CHAVE': 30, 'DESCRICAO': 42, 'ORIGEM': 34})

        # Produtos pareados cuja venda é muito maior que a compra (ex: vendeu 4223,
        # "comprou" 1) — quase sempre casamento errado ou fator de embalagem faltando.
        # Fica numa aba separada para reconferência, sem desfazer nada sozinho.
        suspeitos_prop = getattr(self, 'matches_suspeitos_pos_validacao', pd.DataFrame())
        if suspeitos_prop is not None and not suspeitos_prop.empty:
            sp = suspeitos_prop.copy()
            sp['RATIO_VENDA_COMPRA'] = (sp['QTD_SAIDA'] / sp['QTD_ENTRADA_AJUSTADA'].replace(0, 0.001)).round(1)
            colunas_sp = ['cProd', 'xProd', 'QTD_SAIDA', 'DESCRICAO_NORM', 'QTD_ENTRADA_AJUSTADA',
                          'VLR_UNIT_ENTRADA_AJUSTADO', 'RATIO_VENDA_COMPRA', 'SALDO_QTD']
            colunas_sp = [c for c in colunas_sp if c in sp.columns]
            escrever_aba_formatada(wb, "REVISAR_Proporcao_Implausivel", sp[colunas_sp], COR_VERMELHO,
                                   formatos={'QTD_SAIDA': NUM, 'QTD_ENTRADA_AJUSTADA': NUM,
                                             'VLR_UNIT_ENTRADA_AJUSTADO': MONEY4, 'RATIO_VENDA_COMPRA': '0.0',
                                             'SALDO_QTD': NUM},
                                   larguras_fixas={'xProd': 34, 'DESCRICAO_NORM': 34})

        # Kardex mensal (se gerado) — também disponível aqui. Esta aba e a da base de entradas
        # são as mais pesadas (dezenas de milhares de linhas); a tela web as dispensa.
        if self.exportar_abas_pesadas and getattr(self, 'kardex', None) is not None and not self.kardex.empty:
            escrever_aba_formatada(wb, "Kardex_Mensal", self.kardex, COR_ROXO, formatos=FORMATOS_KARDEX,
                                   colorscale_col='SALDO_FIM_QTD', larguras_fixas={'PRODUTO': 40, 'PERIODO': 12})

        # Base (header estilizado + larguras; muitas linhas/colunas -> corpo leve)
        if self.exportar_abas_pesadas:
            escrever_aba_formatada(wb, "Base_Entradas_Convertidas", base, "5A6B7B", limitar_fonte=1)

        wb.move_sheet("Resumo", -(len(wb.sheetnames) - 1))
        wb.save(outfile)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Consolidador de Estoque - NFC-e e Entradas")
    parser.add_argument('--data_dir', type=str, default='XML SAIDAS',
                        help='Pasta com os ZIPs de NFC-e e as planilhas XLS de entrada. As subpastas também são '
                             'lidas, então dá para organizar por ano (XML SAIDAS/2021, XML SAIDAS/2022, ...).')
    parser.add_argument('--periodo', type=str, default=None,
                        help='Filtra os arquivos pelo nome. Ex: 01-2021 (um mês) ou 2021 (um ano). Sem filtro '
                             'processa todos os anos de uma vez — é o que faz o saldo rolar de um ano para o outro.')
    parser.add_argument('--out_dir', type=str, default='outputs_estoque', help='Diretório de saída principal')
    parser.add_argument('--ignorar_estoque_anterior', action='store_true', help='Ignora saídas sem entrada (ex: para Jan/2021 considerando estoque de 2020)')
    parser.add_argument('--validacao_preenchida', type=str, default=None,
                         help='Caminho da planilha Analise_..._VALIDACAO.xlsx respondida (Sim/Não). As respostas '
                              'são gravadas no cadastro de produtos ANTES do cruzamento e valem para sempre.')
    parser.add_argument('--fatores', type=str, default=None,
                         help='Arquivo CSV/XLSX com colunas PALAVRA_CHAVE e FATOR para fatores de conversão de '
                              'embalagem que não constam na descrição (ex: FLOCUSCUZ;20). Se não informado, o '
                              'script procura automaticamente "fatores_conversao.csv" ou ".xlsx" na pasta de dados.')
    parser.add_argument('--sem_kardex', action='store_true',
                         help='Não gera o kardex mensal (saldo rolando entre períodos). Por padrão o kardex É gerado. '
                              'Para o carry-over funcionar entre meses/anos, rode SEM filtrar --periodo (processe tudo de uma vez).')

    parser.add_argument('--cadastro', type=str, default=None,
                         help='Caminho do cadastro_produtos.csv (memória CHAVE -> PRODUTO_ID). Padrão: '
                              '<out_dir>/cadastro_produtos.csv — na pasta BASE de saída, para valer em todos os anos.')
    parser.add_argument('--cadastro_auto_min_sim', type=float, default=0.98,
                         help='Score mínimo para um casamento automático por palavra-chave/similaridade ser GRAVADO '
                              'no cadastro. Abaixo disso ele vale na execução mas pede confirmação. Padrão 0.98.')
    parser.add_argument('--limiar_ean', type=float, default=0.30,
                         help='Similaridade mínima para aceitar match por EAN igual. Padrão 0.30.')
    parser.add_argument('--limiar_marca', type=float, default=0.55,
                         help='Similaridade mínima para aceitar match por palavra-chave de marca. Padrão 0.55.')
    parser.add_argument('--limiar_auto_similaridade', type=float, default=0.88,
                         help='Similaridade mínima para PAREAR AUTOMATICAMENTE sem chave nenhuma. '
                              'Resolve abreviação e plural (ENERG/ENERGETICO, VIDRO/VIDROS). Padrão 0.88.')
    parser.add_argument('--margem_ambiguidade', type=float, default=0.05,
                         help='Distância mínima entre o melhor e o segundo melhor candidato para o '
                              'aceite automático. Empate vai para a validação manual. Padrão 0.05.')
    parser.add_argument('--sem_match_similaridade', action='store_true',
                         help='Desliga o aceite automático por similaridade alta.')
    parser.add_argument('--trava_preco_min', type=float, default=0.5,
                         help='Venda/custo mínimo aceito num match automático. Padrão 0.5.')
    parser.add_argument('--trava_preco_max', type=float, default=5.0,
                         help='Venda/custo máximo aceito num match automático. Padrão 5.0.')
    parser.add_argument('--sem_trava_preco', action='store_true', help='Desliga a trava de preço.')
    parser.add_argument('--sem_veto_volume', action='store_true', help='Desliga o veto por volume/peso.')
    args = parser.parse_args()

    final_out_dir = os.path.join(args.out_dir, args.periodo) if args.periodo else args.out_dir
    caminho_cadastro = args.cadastro or os.path.join(args.out_dir, 'cadastro_produtos.csv')
    processador = EstoqueProcessor(output_dir=final_out_dir, cadastro_path=caminho_cadastro)
    processador.cadastro_auto_min_sim = args.cadastro_auto_min_sim
    processador.usar_trava_preco = not args.sem_trava_preco
    processador.trava_preco_min = args.trava_preco_min
    processador.trava_preco_max = args.trava_preco_max
    processador.usar_veto_volume = not args.sem_veto_volume
    processador.limiar_ean = args.limiar_ean
    processador.limiar_marca = args.limiar_marca
    processador.usar_match_similaridade = not args.sem_match_similaridade
    processador.limiar_auto_similaridade = args.limiar_auto_similaridade
    processador.margem_ambiguidade = args.margem_ambiguidade

    # Fatores de conversão externos (arquivo informado ou auto-detectado na pasta de dados)
    arquivo_fatores = args.fatores
    if not arquivo_fatores:
        for nome in ('fatores_conversao.csv', 'fatores_conversao.xlsx'):
            candidato = os.path.join(args.data_dir, nome)
            if os.path.exists(candidato):
                arquivo_fatores = candidato
                break
    if arquivo_fatores:
        processador.carregar_fatores_externos(arquivo_fatores)

    processador.extrair_xmls_nfce(args.data_dir, periodo=args.periodo)
    processador.ler_planilhas_entradas(args.data_dir, periodo=args.periodo)
    processador.converter_embalagens_entradas()
    processador.atribuir_chaves()

    # As respostas da validação entram no cadastro ANTES do cruzamento
    if args.validacao_preenchida:
        processador.aplicar_validacao_manual(args.validacao_preenchida)
    processador.aplicar_fatores_do_cadastro()

    processador.comparar_estoque()
    processador.salvar_cadastro()

    if not args.sem_kardex:
        processador.gerar_kardex_mensal()
        if args.periodo:
            logging.warning("Você filtrou --periodo: o kardex terá só esse período e o saldo NÃO rola entre meses/anos. "
                            "Para o carry-over, rode sem --periodo (processando todos os arquivos juntos).")

    processador.exportar_planilha_validacao()
    if args.sem_kardex:
        processador.gerar_inventario_efd_h010(desconsiderar_saidas_sem_entrada=args.ignorar_estoque_anterior)
    else:
        # Um H010 por ano, encadeados pelo saldo de 31/12 do ano anterior.
        processador.gerar_inventarios_por_ano(desconsiderar_saidas_sem_entrada=args.ignorar_estoque_anterior)

    print("\nProcesso concluído com sucesso! Verifique a pasta:", final_out_dir)
