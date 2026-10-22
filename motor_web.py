"""Ponte entre a tela (index.html) e o consolidacao_estoque.py.

Roda DENTRO DO NAVEGADOR, no Pyodide (Python em WebAssembly), num Web Worker: os
arquivos fiscais nunca saem do computador de quem usa a tela. As funções públicas
recebem/devolvem JSON (texto) para atravessar a fronteira Python <-> JavaScript.

Pastas no sistema de arquivos virtual do Pyodide:
    /dados     ZIPs/XMLs de NFC-e, planilhas de entrada, fatores_conversao.csv
    /memoria   cadastro_produtos.csv e cadastro_produtos_rejeitados.csv
    /validacao planilha de validação respondida no Excel (opcional)
    /saida     planilhas geradas para download
"""
import glob
import json
import logging
import os
import shutil

import pandas as pd

import consolidacao_estoque as ce

DADOS, MEMORIA, SAIDA = '/dados', '/memoria', '/saida'
VALIDACAO = '/validacao'  # planilha respondida fica FORA de /dados: senão seria lida como entrada
CADASTRO = os.path.join(MEMORIA, 'cadastro_produtos.csv')

_estado = {'proc': None}


class _LogParaTela(logging.Handler):
    """Manda cada mensagem de log do script para a tela (callback JavaScript)."""

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        try:
            self.callback(record.getMessage(), record.levelname)
        except Exception:
            pass


def configurar_log(callback):
    raiz = logging.getLogger()
    raiz.handlers = [_LogParaTela(callback)]
    raiz.setLevel(logging.INFO)


def _registros(df):
    """DataFrame -> lista de dicts com NaN virando null (JSON válido)."""
    if df is None or df.empty:
        return []
    return json.loads(df.to_json(orient='records', force_ascii=False))


def _texto(caminho):
    if not os.path.exists(caminho):
        return ''
    with open(caminho, encoding='utf-8-sig') as f:
        return f.read()


def _resultado(p, extra=None):
    r = p.resumo_estoque
    anos = []
    for ano, saldos in sorted((getattr(p, 'saldo_kardex_por_ano', {}) or {}).items()):
        pos = [s for s in saldos.values() if s['QTD'] > 0]
        neg = [s for s in saldos.values() if s['QTD'] < 0]
        anos.append({'ano': int(ano), 'valor': round(sum(s['VLR'] for s in pos), 2),
                     'itens_pos': len(pos), 'itens_neg': len(neg),
                     'valor_neg': round(sum(abs(s['VLR']) for s in neg), 2)})

    por_ano = lambda df: df['PERIODO'].astype(str).str[:4].value_counts().sort_index().to_dict() if not df.empty else {}
    cad = p.cadastro_como_df()
    gravado = cad[cad['GRAVADO'] == 'Sim']
    res = {
        'contagem': r['TIPO_MATCH'].value_counts().to_dict() if not r.empty else {},
        'n_produtos': int(len(r)),
        'chaves_venda': int(p.nfce_saidas['CHAVE'].nunique()),
        'chaves_compra': int(p.entradas_xls['CHAVE'].nunique()),
        'itens_por_ano': {'vendas': por_ano(p.nfce_saidas), 'compras': por_ano(p.entradas_xls)},
        'anos': anos,
        'pendencias': _registros(p.validacao_produtos),
        'fatores': _registros(p.tabela_revisao_fatores()),
        'suspeitos': _registros(p.matches_suspeitos_pos_validacao[
            [c for c in ('PRODUTO_ID', 'ROTULO_PRODUTO', 'QTD_ENTRADA_AJUSTADA', 'QTD_SAIDA', 'SALDO_QTD',
                         'CHAVES_VENDA', 'CHAVES_COMPRA') if c in p.matches_suspeitos_pos_validacao.columns]]
                                  if not p.matches_suspeitos_pos_validacao.empty else None),
        # Catálogos enxutos para a busca da tela (corrigir uma sugestão escolhendo outra chave)
        'catalogo_compra': [{'c': c, 'd': d, 'f': f, 'q': round(float(q), 3), 'p': p._pid(c) or ''}
                            for c, d, f, q in zip(p.catalogo_compra['CHAVE'], p.catalogo_compra['DESCRICAO_NORM'],
                                                  p.catalogo_compra['FORNECEDOR'], p.catalogo_compra['QTD_ENTRADA_AJUSTADA'])],
        'catalogo_venda': [{'c': c, 'd': d, 'q': round(float(q), 3), 'p': p._pid(c) or ''}
                           for c, d, q in zip(p.catalogo_venda['CHAVE'], p.catalogo_venda['xProd_NORM'],
                                              p.catalogo_venda['QTD_SAIDA'])],
        'memoria': {'chaves': int(len(gravado)),
                    'produtos': int(gravado.loc[gravado['PRODUTO_ID'] != '', 'PRODUTO_ID'].nunique()),
                    'rejeitados': len(p._rejeitados),
                    'cadastro_csv': _texto(CADASTRO),
                    'rejeitados_csv': _texto(p.rejeitados_path)},
    }
    if extra:
        res.update(extra)
    return json.dumps(res, ensure_ascii=False, default=str)


def _cruzar(p):
    """Do cadastro em diante: fatores -> cruzamento -> grava memória -> kardex."""
    p.converter_embalagens_entradas()
    p.atribuir_chaves()
    p.aplicar_fatores_do_cadastro()
    p.comparar_estoque()
    p.salvar_cadastro()
    p.gerar_kardex_mensal()


def processar(opcoes_json='{}'):
    """Lê tudo o que está em /dados, aplica a planilha de validação respondida (se veio)
    e cruza. Devolve o resultado em JSON."""
    op = json.loads(opcoes_json or '{}')
    shutil.rmtree(SAIDA, ignore_errors=True)
    os.makedirs(MEMORIA, exist_ok=True)

    p = ce.EstoqueProcessor(output_dir=SAIDA, cadastro_path=CADASTRO)
    p.exportar_abas_pesadas = False
    p.usar_trava_preco = bool(op.get('trava_preco', True))
    p.limiar_auto_similaridade = float(op.get('limiar_auto_similaridade', 0.88))
    p.ignorar_estoque_anterior = bool(op.get('ignorar_estoque_anterior', False))

    fatores = os.path.join(DADOS, 'fatores_conversao.csv')
    if os.path.exists(fatores):
        p.carregar_fatores_externos(fatores)

    p.extrair_xmls_nfce(DADOS)
    p.ler_planilhas_entradas(DADOS)
    if p.nfce_saidas.empty or p.entradas_xls.empty:
        falta = 'ZIPs/XMLs de NFC-e (vendas)' if p.nfce_saidas.empty else 'planilhas de entrada (compras)'
        raise ValueError(f"Não encontrei {falta} nos arquivos enviados.")

    extra = {}
    planilha = op.get('planilha_validacao')
    if planilha:
        p.converter_embalagens_entradas()
        p.atribuir_chaves()
        extra['aplicadas'] = p.aplicar_validacao_manual(os.path.join(VALIDACAO, planilha))

    _cruzar(p)
    _estado['proc'] = p
    return _resultado(p, extra)


def revalidar(respostas_json):
    """Grava as respostas dadas na tela e refaz o cruzamento SEM reler os arquivos."""
    p = _estado['proc']
    if p is None:
        raise ValueError("Processe os arquivos primeiro.")
    d = json.loads(respostas_json)
    aplicadas = p.aplicar_respostas(d.get('respostas', []), d.get('fatores', []))
    _cruzar(p)
    return _resultado(p, {'aplicadas': aplicadas})


def gerar_arquivo(tipo, ano=0):
    """Gera um arquivo para download e devolve o caminho dele no sistema virtual."""
    p = _estado['proc']
    if p is None:
        raise ValueError("Processe os arquivos primeiro.")
    os.makedirs(SAIDA, exist_ok=True)
    if tipo == 'validacao':
        return p.exportar_planilha_validacao()
    if tipo == 'h010':
        return p.gerar_inventario_efd_h010(getattr(p, 'ignorar_estoque_anterior', False), ano=int(ano))
    if tipo == 'kardex':
        caminho = os.path.join(SAIDA, 'Kardex_Mensal.csv')
        p.kardex.to_csv(caminho, sep=';', index=False, decimal=',', encoding='utf-8-sig')
        return caminho
    if tipo == 'estoque':
        caminho = os.path.join(SAIDA, 'Estoque_por_Produto.csv')
        p.resumo_estoque.drop(columns=['xProd', 'CHAVE_PRODUTO'], errors='ignore').to_csv(
            caminho, sep=';', index=False, decimal=',', encoding='utf-8-sig')
        return caminho
    raise ValueError(f"Tipo de arquivo desconhecido: {tipo}")


def limpar_dados():
    for pasta in (DADOS, VALIDACAO):
        shutil.rmtree(pasta, ignore_errors=True)
        os.makedirs(pasta, exist_ok=True)
    _estado['proc'] = None


def definir_memoria(cadastro_csv='', rejeitados_csv=''):
    """Grava (ou apaga, se vier vazio) a memória que a tela guardou no navegador."""
    os.makedirs(MEMORIA, exist_ok=True)
    for nome, conteudo in (('cadastro_produtos.csv', cadastro_csv), ('cadastro_produtos_rejeitados.csv', rejeitados_csv)):
        caminho = os.path.join(MEMORIA, nome)
        if conteudo:
            with open(caminho, 'w', encoding='utf-8-sig') as f:
                f.write(conteudo)
        elif os.path.exists(caminho):
            os.remove(caminho)
