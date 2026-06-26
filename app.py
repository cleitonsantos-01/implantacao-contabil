"""
Implantação de saldo e Implantação de saldo e Gerador de Carta de Responsabilidade – BHub
==================================================================
Extrai dados de balancetes PDF → Gera De-Para ECD + Domínio + Carta de Responsabilidade
"""

import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import difflib
import unicodedata
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Implantação de saldo e Gerador de Carta de Responsabilidade – BHub",
    page_icon="📝",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# FUNÇÕES UTILITÁRIAS
# ─────────────────────────────────────────────────────────────────────────────

def parse_br_value(s: str):
    """'1.234,56D' → (1234.56, 'D')   |   '0,00' → (0.0, None)"""
    s = str(s).strip()
    if not s:
        return 0.0, None
    ind = None
    if s and s[-1] in ("D", "C"):
        ind = s[-1]
        s = s[:-1]
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s), ind
    except ValueError:
        return 0.0, ind


def format_br(v: float) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def cls_depth(cls: str) -> int:
    if not cls:
        return 0
    return len(cls.split("."))


def find_leaf_accounts(records: list[dict]) -> set[str]:
    """
    Detecta contas analíticas pela estrutura do plano de contas.

    Critério: contas analíticas estão na PROFUNDIDADE MÁXIMA de classificação
    do seu grupo principal (1º dígito: 1=Ativo, 2=Passivo, 3+=Resultado).
    Contas em profundidade menor são tratadas como sintéticas, mesmo sem
    filhas visíveis no PDF — pois o PDF pode não ter extraído todos os filhos.

    Ex.: se o grupo 1 tem contas até profundidade 5 (ex. 1.1.3.07.000001),
    então "1.1.3.07" (profundidade 4) é sintética mesmo que pareça folha.
    """
    all_cls = [r["classification"] for r in records if r["classification"]]
    if not all_cls:
        return set()

    all_cls_set = set(all_cls)

    # Profundidade máxima por grupo principal (1º dígito da classificação)
    group_max: dict[str, int] = {}
    for cls in all_cls_set:
        g = cls.split(".")[0]
        d = len(cls.split("."))
        if d > group_max.get(g, 0):
            group_max[g] = d

    # Analítica = profundidade igual ao máximo do seu grupo
    leaves: set[str] = set()
    for cls in all_cls_set:
        g = cls.split(".")[0]
        d = len(cls.split("."))
        if d == group_max.get(g, d):
            leaves.add(cls)

    return leaves


def get_depth_summary(records: list[dict]) -> dict[str, int]:
    """Retorna a profundidade analítica detectada por grupo principal."""
    all_cls = [r["classification"] for r in records if r["classification"]]
    group_max: dict[str, int] = {}
    for cls in all_cls:
        g = cls.split(".")[0]
        d = len(cls.split("."))
        if d > group_max.get(g, 0):
            group_max[g] = d
    return group_max


def get_account_group(cls: str) -> str:
    """Retorna 'Ativo', 'Passivo/PL' ou 'Resultado' com base no primeiro dígito da classificação."""
    if not cls:
        return "Outros"
    first = cls.split(".")[0]
    if first == "1":
        return "Ativo"
    elif first == "2":
        return "Passivo/PL"
    else:
        return "Resultado"


def annotate_tipo(records: list[dict]) -> list[dict]:
    """Adiciona campo 'tipo': 'A' (analítica/folha) ou 'S' (sintética/grupo)."""
    leaves = find_leaf_accounts(records)
    # Sem nenhuma classificação no plano, não há hierarquia para determinar
    # sintético vs analítico — tratar todos como analíticos.
    if not leaves:
        for r in records:
            r["tipo"] = "A"
        return records
    for r in records:
        r["tipo"] = "A" if r.get("classification", "") in leaves else "S"
    return records


def signed_total(df_group: pd.DataFrame) -> float:
    """Soma com sinal: D=positivo, C=negativo (convenção contábil para Ativo/Resultado)."""
    d = df_group[df_group["current_indicator"] == "D"]["current_value"].sum()
    c = df_group[df_group["current_indicator"] == "C"]["current_value"].sum()
    return d - c


# ─────────────────────────────────────────────────────────────────────────────
# PLANO DE CONTAS BHUB (contas analíticas — fonte: Plano Padrão BHub Wise PJ)
# Total: 1730 contas | Ativo: 360 | Passivo/PL: 298
# Resultado: Receitas(3.x)=77 | Custo(4.x)=686 | Despesas(5.x)=307
# ─────────────────────────────────────────────────────────────────────────────

BHUB_ACCOUNTS = [
    # ── Ativo ─────────────────────────────────────────────────────────────────
    {"code": "1001", "name": "CAIXA GERAL", "grupo": "Ativo", "mask": "1.1.01.01.00001"},
    {"code": "1002", "name": "CAIXA MATRIZ", "grupo": "Ativo", "mask": "1.1.01.01.00002"},
    {"code": "1003", "name": "CAIXA FILIAIS", "grupo": "Ativo", "mask": "1.1.01.01.00003"},
    {"code": "1004", "name": "CAIXA FUNDO FIXO", "grupo": "Ativo", "mask": "1.1.01.01.00004"},
    {"code": "1005", "name": "NUMERÁRIOS EM TRÂNSITO", "grupo": "Ativo", "mask": "1.1.01.01.00005"},
    {"code": "1006", "name": "MOEDAS ESTRANGEIRAS 1", "grupo": "Ativo", "mask": "1.1.01.01.00006"},
    {"code": "1007", "name": "MOEDAS ESTRANGEIRAS 2", "grupo": "Ativo", "mask": "1.1.01.01.00007"},
    {"code": "1008", "name": "MOEDAS ESTRANGEIRAS 3", "grupo": "Ativo", "mask": "1.1.01.01.00008"},
    {"code": "1018", "name": "BANCO DO BRASIL", "grupo": "Ativo", "mask": "1.1.01.02.00001"},
    {"code": "1019", "name": "BANCO BRADESCO", "grupo": "Ativo", "mask": "1.1.01.02.00002"},
    {"code": "1020", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Ativo", "mask": "1.1.01.02.00003"},
    {"code": "1021", "name": "BANCO SANTANDER", "grupo": "Ativo", "mask": "1.1.01.02.00004"},
    {"code": "1022", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Ativo", "mask": "1.1.01.02.00005"},
    {"code": "1023", "name": "BANCO INTER", "grupo": "Ativo", "mask": "1.1.01.02.00006"},
    {"code": "1024", "name": "CONTA SIMPLES", "grupo": "Ativo", "mask": "1.1.01.02.00007"},
    {"code": "1025", "name": "BANCO BTG", "grupo": "Ativo", "mask": "1.1.01.02.00008"},
    {"code": "1026", "name": "BANCO SICOOB", "grupo": "Ativo", "mask": "1.1.01.02.00009"},
    {"code": "1027", "name": "BANCO C6", "grupo": "Ativo", "mask": "1.1.01.02.00010"},
    {"code": "1028", "name": "BANCO XP", "grupo": "Ativo", "mask": "1.1.01.02.00011"},
    {"code": "1029", "name": "BANCO CORA", "grupo": "Ativo", "mask": "1.1.01.02.00012"},
    {"code": "1030", "name": "BANCO STARKBANK", "grupo": "Ativo", "mask": "1.1.01.02.00013"},
    {"code": "1031", "name": "BANCO NUBANK", "grupo": "Ativo", "mask": "1.1.01.02.00014"},
    {"code": "1032", "name": "BANCO SAFRA", "grupo": "Ativo", "mask": "1.1.01.02.00015"},
    {"code": "1033", "name": "BANCO BS2", "grupo": "Ativo", "mask": "1.1.01.02.00016"},
    {"code": "1034", "name": "BANCO ASAAS", "grupo": "Ativo", "mask": "1.1.01.02.00017"},
    {"code": "1035", "name": "BANCO SICREDI", "grupo": "Ativo", "mask": "1.1.01.02.00018"},
    {"code": "1036", "name": "BANCO PAGSEGURO", "grupo": "Ativo", "mask": "1.1.01.02.00019"},
    {"code": "1037", "name": "BANCO MERCADO PAGO", "grupo": "Ativo", "mask": "1.1.01.02.00020"},
    {"code": "1052", "name": "BANCO DO BRASIL", "grupo": "Ativo", "mask": "1.1.01.03.00001"},
    {"code": "1053", "name": "BANCO BRADESCO", "grupo": "Ativo", "mask": "1.1.01.03.00002"},
    {"code": "1054", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Ativo", "mask": "1.1.01.03.00003"},
    {"code": "1055", "name": "BANCO SANTANDER", "grupo": "Ativo", "mask": "1.1.01.03.00004"},
    {"code": "1056", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Ativo", "mask": "1.1.01.03.00005"},
    {"code": "1057", "name": "BANCO INTER", "grupo": "Ativo", "mask": "1.1.01.03.00006"},
    {"code": "1058", "name": "CONTA SIMPLES", "grupo": "Ativo", "mask": "1.1.01.03.00007"},
    {"code": "1059", "name": "BANCO BTG", "grupo": "Ativo", "mask": "1.1.01.03.00008"},
    {"code": "1060", "name": "BANCO C6", "grupo": "Ativo", "mask": "1.1.01.03.00009"},
    {"code": "1061", "name": "BANCO SICREDI", "grupo": "Ativo", "mask": "1.1.01.03.00010"},
    {"code": "1062", "name": "BANCO XP", "grupo": "Ativo", "mask": "1.1.01.03.00011"},
    {"code": "1063", "name": "BANCO CORA", "grupo": "Ativo", "mask": "1.1.01.03.00012"},
    {"code": "1064", "name": "BANCO STARKBANK", "grupo": "Ativo", "mask": "1.1.01.03.00013"},
    {"code": "1065", "name": "BANCO NUBANK", "grupo": "Ativo", "mask": "1.1.01.03.00014"},
    {"code": "1066", "name": "BANCO SAFRA", "grupo": "Ativo", "mask": "1.1.01.03.00015"},
    {"code": "1067", "name": "BANCO BS2", "grupo": "Ativo", "mask": "1.1.01.03.00016"},
    {"code": "1068", "name": "BANCO ASAAS", "grupo": "Ativo", "mask": "1.1.01.03.00017"},
    {"code": "1069", "name": "BANCO SICOOB", "grupo": "Ativo", "mask": "1.1.01.03.00018"},
    {"code": "1070", "name": "BANCO PAGSEGURO", "grupo": "Ativo", "mask": "1.1.01.03.00019"},
    {"code": "1071", "name": "BANCO MERCADO PAGO", "grupo": "Ativo", "mask": "1.1.01.03.00020"},
    {"code": "1086", "name": "CLIENTES", "grupo": "Ativo", "mask": "1.1.02.01.00001"},
    {"code": "1087", "name": "CLIENTES DO EXTERIOR", "grupo": "Ativo", "mask": "1.1.02.02.00001"},
    {"code": "1088", "name": "CLIENTES NACIONAIS - PARTES RELACIONADAS", "grupo": "Ativo", "mask": "1.1.02.03.00001"},
    {"code": "1089", "name": "CLIENTES DO EXTERIOR - EMPRESAS LIGADAS", "grupo": "Ativo", "mask": "1.1.02.03.00002"},
    {"code": "1090", "name": "CLIENTES DO EXTERIOR - OUTROS", "grupo": "Ativo", "mask": "1.1.02.03.00003"},
    {"code": "1091", "name": "PLATAFORMAS DIGITAIS", "grupo": "Ativo", "mask": "1.1.02.03.00004"},
    {"code": "1092", "name": "CARTÃO DE CRÉDITO", "grupo": "Ativo", "mask": "1.1.02.03.00005"},
    {"code": "1093", "name": "(-) DUPLICATAS DESCONTADAS", "grupo": "Ativo", "mask": "1.1.02.03.00006"},
    {"code": "1094", "name": "VENDAS PARA ENTREGA FUTURA", "grupo": "Ativo", "mask": "1.1.02.03.00007"},
    {"code": "1095", "name": "(-) FATURAMENTO PARA ENTREGA FUTURA", "grupo": "Ativo", "mask": "1.1.02.03.00008"},
    {"code": "1096", "name": "CPC 47 - ATIVOS DE CONTRATO", "grupo": "Ativo", "mask": "1.1.02.03.00009"},
    {"code": "1097", "name": "(-) JUROS A APROPRIAR AVP", "grupo": "Ativo", "mask": "1.1.02.03.00010"},
    {"code": "1107", "name": "(-) PROVISÃO PARA CRÉDITO DE LIQUIDAÇÃO DUVIDOSA", "grupo": "Ativo", "mask": "1.1.02.04.00001"},
    {"code": "1117", "name": "CONSÓRCIOS A RECEBER", "grupo": "Ativo", "mask": "1.1.03.01.00001"},
    {"code": "1118", "name": "DUPLICATAS A RECEBER", "grupo": "Ativo", "mask": "1.1.03.01.00002"},
    {"code": "1119", "name": "CHEQUES A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.01.00003"},
    {"code": "1120", "name": "OUTROS CRÉDITOS", "grupo": "Ativo", "mask": "1.1.03.01.00004"},
    {"code": "1130", "name": "EMPRÉSTIMOS A EMPREGADOS", "grupo": "Ativo", "mask": "1.1.03.02.00001"},
    {"code": "1131", "name": "EMPRÉSTIMOS A SÓCIOS", "grupo": "Ativo", "mask": "1.1.03.02.00002"},
    {"code": "1132", "name": "EMPRÉSTIMOS A RECEBER - PARTES RELACIONADAS", "grupo": "Ativo", "mask": "1.1.03.02.00003"},
    {"code": "1133", "name": "EMPRÉSTIMOS A RECEBER - EMPRESAS LIGADAS", "grupo": "Ativo", "mask": "1.1.03.02.00004"},
    {"code": "1134", "name": "EMPRÉSTIMOS A RECEBER - OUTROS", "grupo": "Ativo", "mask": "1.1.03.02.00005"},
    {"code": "1144", "name": "TÍTULOS A RECEBER", "grupo": "Ativo", "mask": "1.1.03.03.00001"},
    {"code": "1145", "name": "TÍTULOS A RECEBER- PARTES RELACIONADAS", "grupo": "Ativo", "mask": "1.1.03.03.00002"},
    {"code": "1146", "name": "TÍTULOS A RECEBER- EMPRESAS LIGADAS", "grupo": "Ativo", "mask": "1.1.03.03.00003"},
    {"code": "1147", "name": "TÍTULOS A RECEBER- OUTROS", "grupo": "Ativo", "mask": "1.1.03.03.00004"},
    {"code": "1157", "name": "ADIANTAMENTO A FORNECEDORES", "grupo": "Ativo", "mask": "1.1.03.04.00001"},
    {"code": "1158", "name": "ADIANTAMENTO DE IMPORTAÇÃO/EXPORTAÇÃO", "grupo": "Ativo", "mask": "1.1.03.04.00002"},
    {"code": "1159", "name": "ADIANTAMENTOS A IDENTIFICAR", "grupo": "Ativo", "mask": "1.1.03.04.00003"},
    {"code": "1169", "name": "ADIANTAMENTO DE SALÁRIO", "grupo": "Ativo", "mask": "1.1.03.05.00001"},
    {"code": "1170", "name": "ADIANTAMENTO DE 13º SALÁRIO", "grupo": "Ativo", "mask": "1.1.03.05.00002"},
    {"code": "1171", "name": "ADIANTAMENTO DE FÉRIAS", "grupo": "Ativo", "mask": "1.1.03.05.00003"},
    {"code": "1172", "name": "ADIANTAMENTO DE VIAGEM", "grupo": "Ativo", "mask": "1.1.03.05.00004"},
    {"code": "1173", "name": "ADIANTAMENTO DE VIAGEM - EXTERIOR", "grupo": "Ativo", "mask": "1.1.03.05.00005"},
    {"code": "1174", "name": "ADIANTAMENTO A FUNCIONÁRIOS", "grupo": "Ativo", "mask": "1.1.03.05.00006"},
    {"code": "1184", "name": "ADIANTAMENTO A SÓCIOS", "grupo": "Ativo", "mask": "1.1.03.06.00001"},
    {"code": "1185", "name": "ADIANTAMENTO A SÓCIOS - I", "grupo": "Ativo", "mask": "1.1.03.06.00002"},
    {"code": "1186", "name": "ADIANTAMENTO A SÓCIOS - II", "grupo": "Ativo", "mask": "1.1.03.06.00003"},
    {"code": "1187", "name": "ADIANTAMENTO A SÓCIOS - III", "grupo": "Ativo", "mask": "1.1.03.06.00004"},
    {"code": "1197", "name": "IPI A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00001"},
    {"code": "1198", "name": "ICMS A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00002"},
    {"code": "1199", "name": "IRPJ A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00003"},
    {"code": "1200", "name": "IRPJ PAGO POR ESTIMATIVA", "grupo": "Ativo", "mask": "1.1.03.07.00004"},
    {"code": "1201", "name": "CONTRIBUIÇÃO SOCIAL PAGA ESTIMATIVA", "grupo": "Ativo", "mask": "1.1.03.07.00005"},
    {"code": "1202", "name": "CSLL A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00006"},
    {"code": "1203", "name": "COFINS A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00007"},
    {"code": "1204", "name": "PIS A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00008"},
    {"code": "1205", "name": "INSS A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00009"},
    {"code": "1206", "name": "ISS A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00010"},
    {"code": "1207", "name": "IRRF A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00011"},
    {"code": "1208", "name": "CSRF A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00012"},
    {"code": "1209", "name": "IMPOSTOS DE IMPORTAÇÃO A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.07.00013"},
    {"code": "1229", "name": "IPI A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00001"},
    {"code": "1230", "name": "ICMS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00002"},
    {"code": "1231", "name": "IRPJ A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00003"},
    {"code": "1232", "name": "IMPOSTO DE RENDA PAGO POR ESTIMATIVA", "grupo": "Ativo", "mask": "1.1.03.08.00004"},
    {"code": "1233", "name": "CONTRIBUIÇÃO SOCIAL PAGA ESTIMATIVA", "grupo": "Ativo", "mask": "1.1.03.08.00005"},
    {"code": "1234", "name": "CSLL A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00006"},
    {"code": "1235", "name": "COFINS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00007"},
    {"code": "1236", "name": "PIS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00008"},
    {"code": "1237", "name": "INSS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00009"},
    {"code": "1238", "name": "ISS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00010"},
    {"code": "1239", "name": "FGTS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00011"},
    {"code": "1240", "name": "CONTRIBUIÇÃO SINDICAL A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00012"},
    {"code": "1241", "name": "IRRF A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00013"},
    {"code": "1242", "name": "IRRF SOBRE APLICAÇÕES FINANCEIRAS", "grupo": "Ativo", "mask": "1.1.03.08.00014"},
    {"code": "1243", "name": "SIMPLES NACIONAL A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00015"},
    {"code": "1244", "name": "CSRF A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00016"},
    {"code": "1245", "name": "CREDITO DE ICMS - CIAP", "grupo": "Ativo", "mask": "1.1.03.08.00017"},
    {"code": "1246", "name": "IRPJ PAGO A MAIOR", "grupo": "Ativo", "mask": "1.1.03.08.00018"},
    {"code": "1247", "name": "CSLL PAGO A MAIOR", "grupo": "Ativo", "mask": "1.1.03.08.00019"},
    {"code": "1248", "name": "OUTROS TRIBUTOS A COMPENSAR", "grupo": "Ativo", "mask": "1.1.03.08.00020"},
    {"code": "1258", "name": "SALDO NEGATIVO IRPJ", "grupo": "Ativo", "mask": "1.1.03.09.00001"},
    {"code": "1259", "name": "SALDO NEGATIVO CSLL", "grupo": "Ativo", "mask": "1.1.03.09.00002"},
    {"code": "1260", "name": "SALDO NEGATIVO IRPJ", "grupo": "Ativo", "mask": "1.1.03.09.00003"},
    {"code": "1261", "name": "SALDO NEGATIVO CSLL", "grupo": "Ativo", "mask": "1.1.03.09.00004"},
    {"code": "1262", "name": "SALDO NEGATIVO IRPJ", "grupo": "Ativo", "mask": "1.1.03.09.00005"},
    {"code": "1263", "name": "SALDO NEGATIVO CSLL", "grupo": "Ativo", "mask": "1.1.03.09.00006"},
    {"code": "1264", "name": "SALDO NEGATIVO IRPJ", "grupo": "Ativo", "mask": "1.1.03.09.00007"},
    {"code": "1265", "name": "SALDO NEGATIVO CSLL", "grupo": "Ativo", "mask": "1.1.03.09.00008"},
    {"code": "1266", "name": "SALDO NEGATIVO IRPJ", "grupo": "Ativo", "mask": "1.1.03.09.00009"},
    {"code": "1267", "name": "SALDO NEGATIVO CSLL", "grupo": "Ativo", "mask": "1.1.03.09.00010"},
    {"code": "1312", "name": "PRORROGAÇÃO DE LICENÇA MATERNIDADE", "grupo": "Ativo", "mask": "1.1.03.10.00001"},
    {"code": "1313", "name": "PRORROGAÇÃO DE LICENÇA PATERNIDADE", "grupo": "Ativo", "mask": "1.1.03.10.00002"},
    {"code": "1314", "name": "BÔNUS DE ADIMPLÊNCIA FISCAL A RECUPERAR", "grupo": "Ativo", "mask": "1.1.03.10.00003"},
    {"code": "1324", "name": "JUROS SOBRE O CAPITAL PRÓPRIO A RECEBER", "grupo": "Ativo", "mask": "1.1.03.11.00001"},
    {"code": "1334", "name": "PETROBRÁS", "grupo": "Ativo", "mask": "1.1.04.01.00001"},
    {"code": "1335", "name": "MERCADO FINANCEIRO", "grupo": "Ativo", "mask": "1.1.04.01.00002"},
    {"code": "1336", "name": "MERCADO FINANCEIRO - EXTERIOR", "grupo": "Ativo", "mask": "1.1.04.01.00003"},
    {"code": "1346", "name": "VALOR NOMINAL", "grupo": "Ativo", "mask": "1.1.04.02.00001"},
    {"code": "1347", "name": "(-) DESÁGIO A APROPRIAR", "grupo": "Ativo", "mask": "1.1.04.02.00002"},
    {"code": "1357", "name": "MERCADORIAS PARA REVENDA", "grupo": "Ativo", "mask": "1.1.05.01.00001"},
    {"code": "1358", "name": "MATÉRIA-PRIMA", "grupo": "Ativo", "mask": "1.1.05.01.00002"},
    {"code": "1359", "name": "OUTROS MATERIAIS DE CONSUMO", "grupo": "Ativo", "mask": "1.1.05.01.00003"},
    {"code": "1360", "name": "PRODUTOS ACABADOS", "grupo": "Ativo", "mask": "1.1.05.01.00004"},
    {"code": "1361", "name": "PRODUTOS EM ELABORAÇÃO", "grupo": "Ativo", "mask": "1.1.05.01.00005"},
    {"code": "1362", "name": "INSUMOS AGROPECUÁRIOS", "grupo": "Ativo", "mask": "1.1.05.01.00006"},
    {"code": "1363", "name": "PRODUTOS AGROPECUARIOS", "grupo": "Ativo", "mask": "1.1.05.01.00007"},
    {"code": "1364", "name": "ANIMAIS", "grupo": "Ativo", "mask": "1.1.05.01.00008"},
    {"code": "1365", "name": "(-) PROVISÃO PARA AJUSTES DO ESTOQUE", "grupo": "Ativo", "mask": "1.1.05.01.00009"},
    {"code": "1366", "name": "(-) PROVISÃO PARA PERDAS COM ESTOQUE", "grupo": "Ativo", "mask": "1.1.05.01.00010"},
    {"code": "1376", "name": "ALMOXARIFADO DE MANUTENÇÃO", "grupo": "Ativo", "mask": "1.1.05.02.00001"},
    {"code": "1377", "name": "ALMOXARIFADO ADMINISTRATIVO", "grupo": "Ativo", "mask": "1.1.05.02.00002"},
    {"code": "1378", "name": "(-) PROVISÃO PARA AJUSTES DO ESTOQUE", "grupo": "Ativo", "mask": "1.1.05.02.00003"},
    {"code": "1388", "name": "MERCADORIAS PARA REVENDA", "grupo": "Ativo", "mask": "1.1.05.03.00001"},
    {"code": "1389", "name": "MATÉRIA-PRIMA", "grupo": "Ativo", "mask": "1.1.05.03.00002"},
    {"code": "1390", "name": "ENCOMENDA PARA ENTREGA FUTURA - MERCADORIA", "grupo": "Ativo", "mask": "1.1.05.03.00003"},
    {"code": "1391", "name": "ENCOMENDA PARA ENTREGA FUTURA - MATÉRIA PRIMA", "grupo": "Ativo", "mask": "1.1.05.03.00004"},
    {"code": "1392", "name": "(-) PROVISÃO PARA AJUSTES DO ESTOQUE", "grupo": "Ativo", "mask": "1.1.05.03.00005"},
    {"code": "1400", "name": "ESTOQUE EM TRÂNSITO", "grupo": "Ativo", "mask": "1.1.05.04.00001"},
    {"code": "1401", "name": "IMPORTAÇÃO EM ANDAMENTO - MERCADORIAS", "grupo": "Ativo", "mask": "1.1.05.04.00002"},
    {"code": "1402", "name": "IMPORTAÇÃO EM ANDAMENTO - PRODUTOS", "grupo": "Ativo", "mask": "1.1.05.04.00003"},
    {"code": "1403", "name": "IMPORTAÇÃO EM ANDAMENTO - OUTROS", "grupo": "Ativo", "mask": "1.1.05.04.00004"},
    {"code": "1404", "name": "ADIANTAMENTO A FORNEC.EXTERIOR - ESTOQUE", "grupo": "Ativo", "mask": "1.1.05.04.00005"},
    {"code": "1405", "name": "(-) PROVISÕES E AJUSTE ESTOQUE EM TRANSITO", "grupo": "Ativo", "mask": "1.1.05.04.00006"},
    {"code": "1415", "name": "TRANSFERÊNCIAS ENTRE MATRIZ/FILIAIS - ENTRADA", "grupo": "Ativo", "mask": "1.1.05.05.00001"},
    {"code": "1416", "name": "TRANSFERÊNCIAS ENTRE MATRIZ/FILIAIS - SAIDA", "grupo": "Ativo", "mask": "1.1.05.05.00002"},
    {"code": "1417", "name": "ENERGIA ELÉTRICA", "grupo": "Ativo", "mask": "1.1.05.05.00003"},
    {"code": "1418", "name": "COMUNICAÇÕES", "grupo": "Ativo", "mask": "1.1.05.05.00004"},
    {"code": "1419", "name": "TRANSPORTES", "grupo": "Ativo", "mask": "1.1.05.05.00005"},
    {"code": "1420", "name": "OUTROS ESTOQUES - IV", "grupo": "Ativo", "mask": "1.1.05.05.00006"},
    {"code": "1425", "name": "IMOVEIS PARA REVENDA - UNIDADES I", "grupo": "Ativo", "mask": "1.1.05.06.00001"},
    {"code": "1426", "name": "IMOVEIS PARA REVENDA - UNIDADES II", "grupo": "Ativo", "mask": "1.1.05.06.00002"},
    {"code": "1427", "name": "IMOVEIS PARA REVENDA - UNIDADES III", "grupo": "Ativo", "mask": "1.1.05.06.00003"},
    {"code": "1428", "name": "IMOVEIS PARA REVENDA - UNIDADES IV", "grupo": "Ativo", "mask": "1.1.05.06.00004"},
    {"code": "1438", "name": "PARTICIPAÇÕES SOCIETARIAS - I", "grupo": "Ativo", "mask": "1.1.05.07.00001"},
    {"code": "1439", "name": "PARTICIPAÇÕES SOCIETARIAS - II", "grupo": "Ativo", "mask": "1.1.05.07.00002"},
    {"code": "1440", "name": "PARTICIPAÇÕES SOCIETARIAS - III", "grupo": "Ativo", "mask": "1.1.05.07.00003"},
    {"code": "1460", "name": "PRÊMIOS DE SEGUROS A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00001"},
    {"code": "1461", "name": "ASSINATURAS E ANUIDADES", "grupo": "Ativo", "mask": "1.1.06.01.00002"},
    {"code": "1462", "name": "IPVA A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00003"},
    {"code": "1463", "name": "IPTU A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00004"},
    {"code": "1464", "name": "LICENCIAMENTO A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00005"},
    {"code": "1465", "name": "JUROS A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00006"},
    {"code": "1466", "name": "MULTAS A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00007"},
    {"code": "1467", "name": "ALUGUÉIS A VENCER", "grupo": "Ativo", "mask": "1.1.06.01.00008"},
    {"code": "1468", "name": "OUTRAS DESPESAS A APROPRIAR", "grupo": "Ativo", "mask": "1.1.06.01.00009"},
    {"code": "1469", "name": "SEGUROS A VENCER", "grupo": "Ativo", "mask": "1.1.06.01.00010"},
    {"code": "1470", "name": "CARTÃO DE CRÉDITO", "grupo": "Ativo", "mask": "1.1.06.01.00011"},
    {"code": "1480", "name": "IMPOSTO DE RENDA DIFERIDO", "grupo": "Ativo", "mask": "1.1.07.01.00001"},
    {"code": "1481", "name": "CONTRIBUIÇÃO SOCIAL DIFERIDA", "grupo": "Ativo", "mask": "1.1.07.01.00002"},
    {"code": "1482", "name": "PIS DIFERIDO", "grupo": "Ativo", "mask": "1.1.07.01.00003"},
    {"code": "1483", "name": "COFINS DIFERIDO", "grupo": "Ativo", "mask": "1.1.07.01.00004"},
    {"code": "1484", "name": "ISS DIFERIDO", "grupo": "Ativo", "mask": "1.1.07.01.00005"},
    {"code": "1485", "name": "INSS DIFERIDO", "grupo": "Ativo", "mask": "1.1.07.01.00006"},
    {"code": "1486", "name": "OUTROS IMPOSTOS DIFERIDOS", "grupo": "Ativo", "mask": "1.1.07.01.00007"},
    {"code": "1496", "name": "ATIVOS NÃO CIRCULANTES MANTIDOS PARA VENDA", "grupo": "Ativo", "mask": "1.1.08.01.00001"},
    {"code": "1497", "name": "AJUSTE A VALOR PRESENTE DE ATIVOS MANTIDOS PARA A VENDA", "grupo": "Ativo", "mask": "1.1.08.01.00002"},
    {"code": "1498", "name": "(-) REDUÇÃO AO VALOR RECUPERÁVEL DE ATIVOS", "grupo": "Ativo", "mask": "1.1.08.01.00003"},
    {"code": "1508", "name": "CONTA CORRENTE - PESSOAS LIGADAS - PAÍS", "grupo": "Ativo", "mask": "1.1.09.01.00001"},
    {"code": "1509", "name": "CONTA CORRENTE - PARTES NÃO RELACIONADAS", "grupo": "Ativo", "mask": "1.1.09.01.00002"},
    {"code": "1510", "name": "CONTA CORRENTE - EMPRESAS LIGADAS - EXT.", "grupo": "Ativo", "mask": "1.1.09.01.00003"},
    {"code": "1560", "name": "CLIENTES", "grupo": "Ativo", "mask": "1.2.01.01.00001"},
    {"code": "1561", "name": "CLIENTES NACIONAIS - PARTES RELACIONADAS", "grupo": "Ativo", "mask": "1.2.01.01.00002"},
    {"code": "1562", "name": "CLIENTES DO EXTERIOR - EMPRESAS LIGADAS", "grupo": "Ativo", "mask": "1.2.01.01.00003"},
    {"code": "1563", "name": "CLIENTES DO EXTERIOR - OUTROS", "grupo": "Ativo", "mask": "1.2.01.01.00004"},
    {"code": "1564", "name": "DUPLICATAS A RECEBER", "grupo": "Ativo", "mask": "1.2.01.01.00005"},
    {"code": "1565", "name": "(-) DUPLICATAS DESCONTADAS", "grupo": "Ativo", "mask": "1.2.01.01.00006"},
    {"code": "1566", "name": "VENDAS PARA ENTREGA FUTURA", "grupo": "Ativo", "mask": "1.2.01.01.00007"},
    {"code": "1567", "name": "(-) FATURAMENTO PARA ENTREGA FUTURA", "grupo": "Ativo", "mask": "1.2.01.01.00008"},
    {"code": "1568", "name": "CPC 47 - ATIVOS DE CONTRATO", "grupo": "Ativo", "mask": "1.2.01.01.00009"},
    {"code": "1569", "name": "(-) JUROS A APROPRIAR AVP", "grupo": "Ativo", "mask": "1.2.01.01.00010"},
    {"code": "1579", "name": "(-) PROVISÃO PARA CRÉDITO DE LIQUIDAÇÃO DUVIDOSA", "grupo": "Ativo", "mask": "1.2.01.02.00001"},
    {"code": "1589", "name": "CONSÓRCIOS A RECEBER", "grupo": "Ativo", "mask": "1.2.02.01.00001"},
    {"code": "1590", "name": "CAUÇÕES", "grupo": "Ativo", "mask": "1.2.02.01.00002"},
    {"code": "1591", "name": "OUTROS CRÉDITOS", "grupo": "Ativo", "mask": "1.2.02.01.00003"},
    {"code": "1601", "name": "EMPRÉSTIMOS A EMPREGADOS", "grupo": "Ativo", "mask": "1.2.02.02.00001"},
    {"code": "1602", "name": "EMPRÉSTIMOS A SÓCIOS", "grupo": "Ativo", "mask": "1.2.02.02.00002"},
    {"code": "1603", "name": "EMPRÉSTIMOS A TERCEIROS", "grupo": "Ativo", "mask": "1.2.02.02.00003"},
    {"code": "1613", "name": "BANCO DO BRASIL", "grupo": "Ativo", "mask": "1.2.02.04.00001"},
    {"code": "1614", "name": "BANCO BRADESCO", "grupo": "Ativo", "mask": "1.2.02.04.00002"},
    {"code": "1615", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Ativo", "mask": "1.2.02.04.00003"},
    {"code": "1616", "name": "BANCO SANTANDER", "grupo": "Ativo", "mask": "1.2.02.04.00004"},
    {"code": "1617", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Ativo", "mask": "1.2.02.04.00005"},
    {"code": "1618", "name": "BANCO INTER", "grupo": "Ativo", "mask": "1.2.02.04.00006"},
    {"code": "1619", "name": "CONTA SIMPLES", "grupo": "Ativo", "mask": "1.2.02.04.00007"},
    {"code": "1620", "name": "BANCO BTG", "grupo": "Ativo", "mask": "1.2.02.04.00008"},
    {"code": "1621", "name": "BANCO SICOOB", "grupo": "Ativo", "mask": "1.2.02.04.00009"},
    {"code": "1622", "name": "BANCO C6", "grupo": "Ativo", "mask": "1.2.02.04.00010"},
    {"code": "1623", "name": "BANCO XP", "grupo": "Ativo", "mask": "1.2.02.04.00011"},
    {"code": "1624", "name": "BANCO CORA", "grupo": "Ativo", "mask": "1.2.02.04.00012"},
    {"code": "1625", "name": "BANCO STARKBANK", "grupo": "Ativo", "mask": "1.2.02.04.00013"},
    {"code": "1626", "name": "BANCO NUBANK", "grupo": "Ativo", "mask": "1.2.02.04.00014"},
    {"code": "1627", "name": "BANCO SAFRA", "grupo": "Ativo", "mask": "1.2.02.04.00015"},
    {"code": "1628", "name": "BANCO BS2", "grupo": "Ativo", "mask": "1.2.02.04.00016"},
    {"code": "1629", "name": "BANCO ASAAS", "grupo": "Ativo", "mask": "1.2.02.04.00017"},
    {"code": "1630", "name": "BANCO SICREDI", "grupo": "Ativo", "mask": "1.2.02.04.00018"},
    {"code": "1631", "name": "BANCO PAGSEGURO", "grupo": "Ativo", "mask": "1.2.02.04.00019"},
    {"code": "1632", "name": "BANCO MERCADO PAGO", "grupo": "Ativo", "mask": "1.2.02.04.00020"},
    {"code": "1647", "name": "DEPÓSITOS JUDICIAIS", "grupo": "Ativo", "mask": "1.2.02.05.00001"},
    {"code": "1657", "name": "IPI A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00001"},
    {"code": "1658", "name": "ICMS A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00002"},
    {"code": "1659", "name": "IRPJ A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00003"},
    {"code": "1660", "name": "IMPOSTO DE RENDA PAGO POR ESTIMATIVA", "grupo": "Ativo", "mask": "1.2.02.06.00004"},
    {"code": "1661", "name": "CONTRIBUIÇÃO SOCIAL PAGA ESTIMATIVA", "grupo": "Ativo", "mask": "1.2.02.06.00005"},
    {"code": "1662", "name": "CSLL A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00006"},
    {"code": "1663", "name": "COFINS A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00007"},
    {"code": "1664", "name": "PIS A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00008"},
    {"code": "1665", "name": "INSS A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00009"},
    {"code": "1666", "name": "ISS A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00010"},
    {"code": "1667", "name": "FGTS A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00011"},
    {"code": "1668", "name": "CONTRIBUIÇÃO SINDICAL A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00012"},
    {"code": "1669", "name": "IRRF A RECUPERAR/COMPENSAR", "grupo": "Ativo", "mask": "1.2.02.06.00013"},
    {"code": "1679", "name": "PRÊMIOS DE SEGUROS A APROPRIAR", "grupo": "Ativo", "mask": "1.2.02.07.00001"},
    {"code": "1680", "name": "ASSINATURAS E ANUIDADES", "grupo": "Ativo", "mask": "1.2.02.07.00002"},
    {"code": "1681", "name": "JUROS A APROPRIAR", "grupo": "Ativo", "mask": "1.2.02.07.00003"},
    {"code": "1682", "name": "MULTAS A APROPRIAR", "grupo": "Ativo", "mask": "1.2.02.07.00004"},
    {"code": "1683", "name": "ALUGUÉIS A VENCER", "grupo": "Ativo", "mask": "1.2.02.07.00005"},
    {"code": "1684", "name": "OUTRAS DESPESAS A APROPRIAR", "grupo": "Ativo", "mask": "1.2.02.07.00006"},
    {"code": "1685", "name": "SEGUROS A VENCER", "grupo": "Ativo", "mask": "1.2.02.07.00007"},
    {"code": "1686", "name": "CARTÃO DE CRÉDITO", "grupo": "Ativo", "mask": "1.2.02.07.00008"},
    {"code": "1716", "name": "CONTROLADA - VALOR PATRIMONIAL", "grupo": "Ativo", "mask": "1.2.03.01.00001"},
    {"code": "1717", "name": "COLIGADAS - VALOR PATRIMONIAL", "grupo": "Ativo", "mask": "1.2.03.01.00002"},
    {"code": "1718", "name": "(-) CONTROL/COLIG - AMORTIZAÇÃO DO ÁGIO", "grupo": "Ativo", "mask": "1.2.03.01.00003"},
    {"code": "1719", "name": "(-) CONTROL/COLIG - DESÁGIO NA AQUISIÇÃO", "grupo": "Ativo", "mask": "1.2.03.01.00004"},
    {"code": "1720", "name": "CONTROLADA - AMORTIZAÇÃO DO DESÁGIO", "grupo": "Ativo", "mask": "1.2.03.01.00005"},
    {"code": "1721", "name": "ADIANTAMENTO P/FUTURO AUMENTO DE CAPITAL", "grupo": "Ativo", "mask": "1.2.03.01.00006"},
    {"code": "1722", "name": "PARTICIPAÇÕES SOCIETÁRIAS COLIGADAS-PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00007"},
    {"code": "1723", "name": "PARTICIPAÇÕES SOCIET. CONTROLADAS - PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00008"},
    {"code": "1724", "name": "PARTICIPAÇÕES SOCIET.CONTROLADA-EXTERIOR", "grupo": "Ativo", "mask": "1.2.03.01.00009"},
    {"code": "1725", "name": "OUTRAS PARTICIPAÇÕES SOCIETÁRIAS-PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00010"},
    {"code": "1726", "name": "PARTIC.EM SOCIEDADES SCP-SÓCIO OSTENSIVO", "grupo": "Ativo", "mask": "1.2.03.01.00011"},
    {"code": "1727", "name": "CERTIFICADOS DE INVESTIMENTOS", "grupo": "Ativo", "mask": "1.2.03.01.00012"},
    {"code": "1728", "name": "PARTIC.EM SOCIEDADES SCP-SÓCIO PARTIC.", "grupo": "Ativo", "mask": "1.2.03.01.00013"},
    {"code": "1729", "name": "AÇÕES DE OUTRAS EMPRESAS", "grupo": "Ativo", "mask": "1.2.03.01.00014"},
    {"code": "1730", "name": "GOODWILL EM INVESTIMENTOS NO PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00015"},
    {"code": "1731", "name": "MAIS VALIA EM INVESTIMENTOS - NO PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00016"},
    {"code": "1732", "name": "(-) PERDAS ESTIMADAS POR REDUÇÃO VR.REC.", "grupo": "Ativo", "mask": "1.2.03.01.00017"},
    {"code": "1733", "name": "MENOS VALIA EM INVESTIMENTOS - NO PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00018"},
    {"code": "1734", "name": "OUTROS INVESTIMENTOS - PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00019"},
    {"code": "1735", "name": "OUTRAS CONTAS - PAÍS", "grupo": "Ativo", "mask": "1.2.03.01.00020"},
    {"code": "1736", "name": "(-) AJUSTES NEGAT.", "grupo": "Ativo", "mask": "1.2.03.01.00021"},
    {"code": "1737", "name": "PROPRIEDADE PARA INVEST.VALOR JUSTO", "grupo": "Ativo", "mask": "1.2.03.01.00022"},
    {"code": "1738", "name": "PROPRIEDADE PAR INVESTIMENTO - CUSTO", "grupo": "Ativo", "mask": "1.2.03.01.00023"},
    {"code": "1739", "name": "PROPRIEDADE PARA INVEST.(-)DEPREC.ACUM", "grupo": "Ativo", "mask": "1.2.03.01.00024"},
    {"code": "1740", "name": "EQUIVALÊNCIA PATRIMONIAL - CONTROLADA", "grupo": "Ativo", "mask": "1.2.03.01.00025"},
    {"code": "1741", "name": "EQUIVALÊNCIA PATRIMONIAL - COLIGADA", "grupo": "Ativo", "mask": "1.2.03.01.00026"},
    {"code": "1749", "name": "IMÓVEIS NÃO DESTINADOS AO USO", "grupo": "Ativo", "mask": "1.2.03.02.00001"},
    {"code": "1759", "name": "TERRENOS", "grupo": "Ativo", "mask": "1.2.04.01.00001"},
    {"code": "1760", "name": "IMÓVEIS", "grupo": "Ativo", "mask": "1.2.04.01.00002"},
    {"code": "1761", "name": "MÓVEIS E UTENSÍLIOS", "grupo": "Ativo", "mask": "1.2.04.01.00003"},
    {"code": "1762", "name": "MÁQUINAS E EQUIPAMENTOS", "grupo": "Ativo", "mask": "1.2.04.01.00004"},
    {"code": "1763", "name": "FERRAMENTAS E ACESSÓRIOS", "grupo": "Ativo", "mask": "1.2.04.01.00005"},
    {"code": "1764", "name": "VEÍCULOS", "grupo": "Ativo", "mask": "1.2.04.01.00006"},
    {"code": "1765", "name": "MÁQUINAS DE PROCESSAMENTOS DE DADOS", "grupo": "Ativo", "mask": "1.2.04.01.00007"},
    {"code": "1766", "name": "INSTALAÇÕES", "grupo": "Ativo", "mask": "1.2.04.01.00008"},
    {"code": "1767", "name": "BENFEITORIAS EM IMÓVEIS DE TERCEIROS", "grupo": "Ativo", "mask": "1.2.04.01.00009"},
    {"code": "1768", "name": "MOLDES E MATRIZES", "grupo": "Ativo", "mask": "1.2.04.01.00010"},
    {"code": "1769", "name": "BIBLIOTECAS", "grupo": "Ativo", "mask": "1.2.04.01.00011"},
    {"code": "1770", "name": "EMBARCAÇÕES", "grupo": "Ativo", "mask": "1.2.04.01.00012"},
    {"code": "1771", "name": "AERONAVES", "grupo": "Ativo", "mask": "1.2.04.01.00013"},
    {"code": "1772", "name": "DUTOS E TUBULAÇÕES", "grupo": "Ativo", "mask": "1.2.04.01.00014"},
    {"code": "1773", "name": "MAQUINAS DE ATIVIDADE RURAL", "grupo": "Ativo", "mask": "1.2.04.01.00015"},
    {"code": "1774", "name": "OUTRAS IMOBILIZAÇÕES", "grupo": "Ativo", "mask": "1.2.04.01.00016"},
    {"code": "1824", "name": "CONSTRUÇÃO EM ANDAMENTO", "grupo": "Ativo", "mask": "1.2.04.02.00001"},
    {"code": "1825", "name": "ADIANTAMENTO A FORNECEDORES DE IMOBILIZADO", "grupo": "Ativo", "mask": "1.2.04.02.00002"},
    {"code": "1826", "name": "(-) DEVOLUÇÕES - BENS DO IMOBILIZADO", "grupo": "Ativo", "mask": "1.2.04.02.00003"},
    {"code": "1835", "name": "(-) DEPRECIAÇÕES DE IMÓVEIS", "grupo": "Ativo", "mask": "1.2.04.03.00001"},
    {"code": "1836", "name": "(-) DEPRECIAÇÕES DE MÓVEIS E UTENSÍLIOS", "grupo": "Ativo", "mask": "1.2.04.03.00002"},
    {"code": "1837", "name": "(-) DEPRECIAÇÕES DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Ativo", "mask": "1.2.04.03.00003"},
    {"code": "1838", "name": "(-) DEPRECIAÇÕES DE FERRAMENTAS E ACESSÓRIOS", "grupo": "Ativo", "mask": "1.2.04.03.00004"},
    {"code": "1839", "name": "(-) DEPRECIAÇÕES DE VEÍCULOS", "grupo": "Ativo", "mask": "1.2.04.03.00005"},
    {"code": "1840", "name": "(-) DEPRECIAÇÕES DE MÁQUINAS DE PROCESSAMENTOS DE DADOS", "grupo": "Ativo", "mask": "1.2.04.03.00006"},
    {"code": "1841", "name": "(-) DEPRECIAÇÕES DE INSTALAÇÕES", "grupo": "Ativo", "mask": "1.2.04.03.00007"},
    {"code": "1842", "name": "(-) AMORTIZAÇÃO DE BENFEITORIA EM IMÓVEIS DE TERCEIROS", "grupo": "Ativo", "mask": "1.2.04.03.00008"},
    {"code": "1843", "name": "(-) DEPRECIAÇÕES DE MOLDES E MATRIZES", "grupo": "Ativo", "mask": "1.2.04.03.00009"},
    {"code": "1844", "name": "(-) BIBLIOTECAS", "grupo": "Ativo", "mask": "1.2.04.03.00010"},
    {"code": "1845", "name": "(-) EMBARCAÇÕES", "grupo": "Ativo", "mask": "1.2.04.03.00011"},
    {"code": "1846", "name": "(-) AERONAVES", "grupo": "Ativo", "mask": "1.2.04.03.00012"},
    {"code": "1847", "name": "(-) DUTOS E TUBULAÇÕES", "grupo": "Ativo", "mask": "1.2.04.03.00013"},
    {"code": "1848", "name": "(-) MAQUINAS DE ATIVIDADE RURAL", "grupo": "Ativo", "mask": "1.2.04.03.00014"},
    {"code": "1849", "name": "(-) OUTRAS IMOBILIZAÇÕES", "grupo": "Ativo", "mask": "1.2.04.03.00015"},
    {"code": "1869", "name": "SOFTWARES", "grupo": "Ativo", "mask": "1.2.05.01.00001"},
    {"code": "1870", "name": "MARCAS E PATENTES", "grupo": "Ativo", "mask": "1.2.05.01.00002"},
    {"code": "1871", "name": "DIREITOS", "grupo": "Ativo", "mask": "1.2.05.01.00003"},
    {"code": "1872", "name": "GOODWILL - NO PAÍS", "grupo": "Ativo", "mask": "1.2.05.01.00004"},
    {"code": "1873", "name": "FUNDO DE COMÉRCIO", "grupo": "Ativo", "mask": "1.2.05.01.00005"},
    {"code": "1874", "name": "MAIS VALIA EM INVESTIMENTOS NO PAÍS", "grupo": "Ativo", "mask": "1.2.05.01.00006"},
    {"code": "1875", "name": "CONCESSÕES", "grupo": "Ativo", "mask": "1.2.05.01.00007"},
    {"code": "1876", "name": "DIREITOS DE EXPLORAÇÃO FLORESTAL", "grupo": "Ativo", "mask": "1.2.05.01.00008"},
    {"code": "1877", "name": "DIREITOS AUTORAIS", "grupo": "Ativo", "mask": "1.2.05.01.00009"},
    {"code": "1878", "name": "DIREITOS DE EXPLORAÇÃO MINERAL", "grupo": "Ativo", "mask": "1.2.05.01.00010"},
    {"code": "1879", "name": "FRANQUIAS", "grupo": "Ativo", "mask": "1.2.05.01.00011"},
    {"code": "1880", "name": "OUTRAS AQUISIÇÕES INTANGÍVEIS", "grupo": "Ativo", "mask": "1.2.05.01.00012"},
    {"code": "1881", "name": "DESENVOLVIMENTO DE PRODUTOS", "grupo": "Ativo", "mask": "1.2.05.01.00013"},
    {"code": "1882", "name": "(-) PERDAS ESTIMADAS POR REDUÇÃO VR.REC.", "grupo": "Ativo", "mask": "1.2.05.01.00014"},
    {"code": "1883", "name": "(-) MENOS VALIA EM INVESTIMENTO - PAÍS", "grupo": "Ativo", "mask": "1.2.05.01.00015"},
    {"code": "1884", "name": "CONTRATOS DE ALUGUEL (DIREITO DE USO)", "grupo": "Ativo", "mask": "1.2.05.01.00016"},
    {"code": "1885", "name": "(-) AVP DE CONTRATOS DE ALUGUEL", "grupo": "Ativo", "mask": "1.2.05.01.00017"},
    {"code": "1895", "name": "(-) AMORTIZAÇÃO DE SOFTWARE", "grupo": "Ativo", "mask": "1.2.05.02.00001"},
    {"code": "1896", "name": "(-) AMORTIZAÇÃO DE MARCAS E PATENTES", "grupo": "Ativo", "mask": "1.2.05.02.00002"},
    {"code": "1897", "name": "(-) AMORTIZAÇÃO DE DIREITOS", "grupo": "Ativo", "mask": "1.2.05.02.00003"},
    {"code": "1898", "name": "(-) CONCESSÕES", "grupo": "Ativo", "mask": "1.2.05.02.00004"},
    {"code": "1899", "name": "(-) AMORTIZAÇÃO DE ÁGIOS EM INVESTIMENTOS", "grupo": "Ativo", "mask": "1.2.05.02.00005"},
    {"code": "1900", "name": "(-) FRANQUIAS", "grupo": "Ativo", "mask": "1.2.05.02.00006"},
    {"code": "1901", "name": "(-) DIREITOS DE EXPLORAÇÃO FLORESTAL", "grupo": "Ativo", "mask": "1.2.05.02.00007"},
    {"code": "1902", "name": "(-) DIREITOS AUTORAIS", "grupo": "Ativo", "mask": "1.2.05.02.00008"},
    {"code": "1903", "name": "(-) DIREITOS DE USO EXPLORAÇÃO MINERAL", "grupo": "Ativo", "mask": "1.2.05.02.00009"},
    {"code": "1904", "name": "(-) OUTRAS AMORTIZAÇÕES DO INTANGÍVEL", "grupo": "Ativo", "mask": "1.2.05.02.00010"},
    {"code": "1905", "name": "(-) AMORTIZAÇÃO CONTRATOS DE ALUGUEL", "grupo": "Ativo", "mask": "1.2.05.02.00011"},
    {"code": "1906", "name": "(-) AMORTIZACAO AVP CONTRATOS DE ALUGUEL", "grupo": "Ativo", "mask": "1.2.05.02.00012"},
    {"code": "1926", "name": "ARRENDAMENTO FINANCEIRO - IMOVEIS IFRS 16", "grupo": "Ativo", "mask": "1.2.06.01.00001"},
    {"code": "1927", "name": "(-) JUROS S/ ARRENDAMENTO IMOVEL IFRS 16", "grupo": "Ativo", "mask": "1.2.06.01.00002"},
    {"code": "1928", "name": "(-) AMORTIZAÇÃO ARRENDAMENTO FINANCEIRO - IMOVEIS IFRS 16", "grupo": "Ativo", "mask": "1.2.06.01.00003"},
    {"code": "1929", "name": "ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Ativo", "mask": "1.2.06.02.00001"},
    {"code": "1930", "name": "(-) JUROS S/ ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Ativo", "mask": "1.2.06.02.00002"},
    {"code": "1931", "name": "(-) AMORTIZAÇÃO ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Ativo", "mask": "1.2.06.02.00003"},
    {"code": "1941", "name": "DEMAIS TITULOS A RECEBER", "grupo": "Ativo", "mask": "1.2.07.01.00001"},
    {"code": "1951", "name": "COMPRA PARA ENTREGA FUTURA", "grupo": "Ativo", "mask": "1.2.99.01.00001"},
    {"code": "1952", "name": "COMPRAS PARA ENTREGA FUTURA", "grupo": "Ativo", "mask": "1.2.99.01.00002"},
    {"code": "1953", "name": "REMESSA AO ESTABELECIMENTO", "grupo": "Ativo", "mask": "1.2.99.01.00003"},
    {"code": "1955", "name": "CONSIGNAÇÃO", "grupo": "Ativo", "mask": "1.2.99.01.00005"},
    # ── Passivo/PL ────────────────────────────────────────────────────────────
    {"code": "2001", "name": "SALÁRIOS E ORDENADOS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00001"},
    {"code": "2002", "name": "PRÓ-LABORE A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00002"},
    {"code": "2003", "name": "FÉRIAS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00003"},
    {"code": "2004", "name": "13º SALARIO A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00004"},
    {"code": "2005", "name": "ABONO DE PIS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00005"},
    {"code": "2006", "name": "RESCISÃO A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00006"},
    {"code": "2007", "name": "PARTICIPAÇÃO NOS LUCROS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00007"},
    {"code": "2008", "name": "GRATIFICAÇÕES A A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00008"},
    {"code": "2009", "name": "BÔNUS", "grupo": "Passivo/PL", "mask": "2.1.01.01.00009"},
    {"code": "2010", "name": "PENSÃO ALIMENTÍCIA A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00010"},
    {"code": "2011", "name": "AUTÔNOMOS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00011"},
    {"code": "2012", "name": "ESTAGIÁRIOS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.01.01.00012"},
    {"code": "2022", "name": "INSS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.01.02.00001"},
    {"code": "2023", "name": "FGTS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.01.02.00002"},
    {"code": "2024", "name": "PIS SOBRE FOLHA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.01.02.00003"},
    {"code": "2029", "name": "APROPRIAÇÕES PARA FÉRIAS", "grupo": "Passivo/PL", "mask": "2.1.01.03.00001"},
    {"code": "2030", "name": "APROPRIAÇÕES PARA 13º SALÁRIO", "grupo": "Passivo/PL", "mask": "2.1.01.03.00002"},
    {"code": "2031", "name": "INSS SOBRE APROPRIAÇÕES PARA FÉRIAS", "grupo": "Passivo/PL", "mask": "2.1.01.03.00003"},
    {"code": "2032", "name": "INSS SOBRE APROPRIAÇÕES PARA 13º SALÁRIO", "grupo": "Passivo/PL", "mask": "2.1.01.03.00004"},
    {"code": "2033", "name": "FGTS SOBRE APROPRIAÇÕES PARA FÉRIAS", "grupo": "Passivo/PL", "mask": "2.1.01.03.00005"},
    {"code": "2034", "name": "FGTS SOBRE APROPRIAÇÕES PARA 13º SALÁRIO", "grupo": "Passivo/PL", "mask": "2.1.01.03.00006"},
    {"code": "2035", "name": "PIS SOBRE APROPRIAÇÕES PARA FÉRIAS", "grupo": "Passivo/PL", "mask": "2.1.01.03.00007"},
    {"code": "2036", "name": "PIS SOBRE APROPRIAÇÕES PARA 13º SALÁRIO", "grupo": "Passivo/PL", "mask": "2.1.01.03.00008"},
    {"code": "2041", "name": "PROVISÃO PARA CUSTOS E DESPESAS", "grupo": "Passivo/PL", "mask": "2.1.01.04.00001"},
    {"code": "2042", "name": "PROVISÃO PARA CONTINGÊNCIAS", "grupo": "Passivo/PL", "mask": "2.1.01.04.00002"},
    {"code": "2043", "name": "PROVISÃO PARA COMISSÕES DE TERCEIROS", "grupo": "Passivo/PL", "mask": "2.1.01.04.00003"},
    {"code": "2044", "name": "PROVISÕES DE NATUREZA FISCAL", "grupo": "Passivo/PL", "mask": "2.1.01.04.00004"},
    {"code": "2045", "name": "PROVISÕES DE NATUREZA CÍVEL", "grupo": "Passivo/PL", "mask": "2.1.01.04.00005"},
    {"code": "2046", "name": "CPC 47 - PASSIVOS DE CONTRATO", "grupo": "Passivo/PL", "mask": "2.1.01.04.00006"},
    {"code": "2047", "name": "(-) JUROS A APROPRIAR AVP", "grupo": "Passivo/PL", "mask": "2.1.01.04.00007"},
    {"code": "2057", "name": "IPI A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00001"},
    {"code": "2058", "name": "ICMS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00002"},
    {"code": "2059", "name": "ICMS ST A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00003"},
    {"code": "2060", "name": "ICMS DIFAL A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00004"},
    {"code": "2061", "name": "ICMS ANTECIPADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00005"},
    {"code": "2062", "name": "ISS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00006"},
    {"code": "2063", "name": "ISS RETIDO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00007"},
    {"code": "2064", "name": "ISS SOBRE IMPORTAÇÃO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00008"},
    {"code": "2065", "name": "PROVISÃO PARA IMPOSTO DE RENDA SOBRE O LUCRO", "grupo": "Passivo/PL", "mask": "2.1.02.01.00009"},
    {"code": "2066", "name": "PROVISÃO PARA CONTRIBUIÇÃO SOCIAL SOBRE O LUCRO", "grupo": "Passivo/PL", "mask": "2.1.02.01.00010"},
    {"code": "2067", "name": "IMPOSTO DE RENDA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00011"},
    {"code": "2068", "name": "CONTRIBUIÇÃO SOCIAL A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00012"},
    {"code": "2069", "name": "CONTRIBUIÇÃO SOCIAL RETIDA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00013"},
    {"code": "2070", "name": "CBS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00014"},
    {"code": "2071", "name": "IBS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00015"},
    {"code": "2072", "name": "IRRF SOBRE FOLHA A RECOLHER COD. 0561", "grupo": "Passivo/PL", "mask": "2.1.02.01.00016"},
    {"code": "2073", "name": "IRRF SOBRE SERVIÇOS A RECOLHER COD. 1708", "grupo": "Passivo/PL", "mask": "2.1.02.01.00017"},
    {"code": "2074", "name": "IRRF SOBRE SERVIÇOS A RECOLHER COD. 8045", "grupo": "Passivo/PL", "mask": "2.1.02.01.00018"},
    {"code": "2075", "name": "IRRF SOBRE ALUGUEL A RECOLHER COD. 3208", "grupo": "Passivo/PL", "mask": "2.1.02.01.00019"},
    {"code": "2076", "name": "IRRF SOBRE AUTONOMOS A RECOLHER COD. 0588", "grupo": "Passivo/PL", "mask": "2.1.02.01.00020"},
    {"code": "2077", "name": "IRRF SOBRE IMPORTAÇÃO A RECOLHER COD. 0422", "grupo": "Passivo/PL", "mask": "2.1.02.01.00021"},
    {"code": "2078", "name": "CIDE A RECOLHER COD. 8741", "grupo": "Passivo/PL", "mask": "2.1.02.01.00022"},
    {"code": "2079", "name": "PIS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00023"},
    {"code": "2080", "name": "PIS RETIDO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00024"},
    {"code": "2081", "name": "PIS SOBRE IMPORTAÇÃO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00025"},
    {"code": "2082", "name": "COFINS A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00026"},
    {"code": "2083", "name": "COFINS RETIDO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00027"},
    {"code": "2084", "name": "COFINS SOBRE IMPORTAÇÃO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00028"},
    {"code": "2085", "name": "CSRF A RECOLHER COD. 5952", "grupo": "Passivo/PL", "mask": "2.1.02.01.00029"},
    {"code": "2086", "name": "IOF A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00030"},
    {"code": "2087", "name": "INSS RETIDO A RECOLHER COD. 2631", "grupo": "Passivo/PL", "mask": "2.1.02.01.00031"},
    {"code": "2088", "name": "INSS RECEITA BRUTA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00032"},
    {"code": "2089", "name": "SIMPLES NACIONAL A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00033"},
    {"code": "2090", "name": "CONTRIBUIÇÃO SINDICAL A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00034"},
    {"code": "2091", "name": "IRPJ POR ESTIMATIVA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00035"},
    {"code": "2092", "name": "CSLL POR ESTIMATIVA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00036"},
    {"code": "2093", "name": "IPTU A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00037"},
    {"code": "2094", "name": "IPVA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00038"},
    {"code": "2095", "name": "FIA A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00039"},
    {"code": "2096", "name": "FUNRURAL A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.02.01.00040"},
    {"code": "2110", "name": "FORNECEDORES", "grupo": "Passivo/PL", "mask": "2.1.03.01.00001"},
    {"code": "2130", "name": "FORNECEDORES", "grupo": "Passivo/PL", "mask": "2.1.03.02.00001"},
    {"code": "2140", "name": "BANCO DO BRASIL", "grupo": "Passivo/PL", "mask": "2.1.04.01.00001"},
    {"code": "2141", "name": "BANCO BRADESCO", "grupo": "Passivo/PL", "mask": "2.1.04.01.00002"},
    {"code": "2142", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Passivo/PL", "mask": "2.1.04.01.00003"},
    {"code": "2143", "name": "BANCO SANTANDER", "grupo": "Passivo/PL", "mask": "2.1.04.01.00004"},
    {"code": "2144", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Passivo/PL", "mask": "2.1.04.01.00005"},
    {"code": "2145", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.04.01.00006"},
    {"code": "2154", "name": "BANCO DO BRASIL", "grupo": "Passivo/PL", "mask": "2.1.04.02.00001"},
    {"code": "2155", "name": "BANCO BRADESCO", "grupo": "Passivo/PL", "mask": "2.1.04.02.00002"},
    {"code": "2156", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Passivo/PL", "mask": "2.1.04.02.00003"},
    {"code": "2157", "name": "BANCO SANTANDER", "grupo": "Passivo/PL", "mask": "2.1.04.02.00004"},
    {"code": "2158", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Passivo/PL", "mask": "2.1.04.02.00005"},
    {"code": "2159", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.04.02.00006"},
    {"code": "2163", "name": "DUPLICATAS DESCONTADAS", "grupo": "Passivo/PL", "mask": "2.1.04.03.00001"},
    {"code": "2164", "name": "EMPRÉSTIMOS DE TERCEIROS", "grupo": "Passivo/PL", "mask": "2.1.04.03.00002"},
    {"code": "2165", "name": "EMPRÉSTIMOS SÓCIOS", "grupo": "Passivo/PL", "mask": "2.1.04.03.00003"},
    {"code": "2166", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.04.03.00004"},
    {"code": "2185", "name": "BANCO DO BRASIL", "grupo": "Passivo/PL", "mask": "2.1.04.04.00001"},
    {"code": "2186", "name": "BANCO BRADESCO", "grupo": "Passivo/PL", "mask": "2.1.04.04.00002"},
    {"code": "2187", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Passivo/PL", "mask": "2.1.04.04.00003"},
    {"code": "2188", "name": "BANCO SANTANDER", "grupo": "Passivo/PL", "mask": "2.1.04.04.00004"},
    {"code": "2189", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Passivo/PL", "mask": "2.1.04.04.00005"},
    {"code": "2190", "name": "BANCO VOLKSWAGEN", "grupo": "Passivo/PL", "mask": "2.1.04.04.00006"},
    {"code": "2191", "name": "FINANCIAMENTOS", "grupo": "Passivo/PL", "mask": "2.1.04.04.00007"},
    {"code": "2192", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.04.04.00008"},
    {"code": "2201", "name": "EMPRÉSTIMOS ESTRANGEIROS", "grupo": "Passivo/PL", "mask": "2.1.04.05.00001"},
    {"code": "2202", "name": "FINANCIAMENTOS ESTRANGEIROS", "grupo": "Passivo/PL", "mask": "2.1.04.05.00002"},
    {"code": "2203", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.04.05.00003"},
    {"code": "2212", "name": "ADIANTAMENTO DE CLIENTES", "grupo": "Passivo/PL", "mask": "2.1.05.01.00001"},
    {"code": "2222", "name": "ADIANTAMENTO DE CLIENTES NO EXTERIOR", "grupo": "Passivo/PL", "mask": "2.1.05.02.00001"},
    {"code": "2232", "name": "RECEBIMENTOS A IDENTIFICAR", "grupo": "Passivo/PL", "mask": "2.1.05.03.00001"},
    {"code": "2242", "name": "HONORÁRIOS CONTÁBEIS", "grupo": "Passivo/PL", "mask": "2.1.06.01.00001"},
    {"code": "2243", "name": "ALUGUEL A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00002"},
    {"code": "2244", "name": "CONDOMÍNIO A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00003"},
    {"code": "2245", "name": "SEGUROS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00004"},
    {"code": "2246", "name": "SERVIÇOS PRESTADOS POR TERCEIROS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00005"},
    {"code": "2247", "name": "ENERGIA ELÉTRICA A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00006"},
    {"code": "2248", "name": "CONSUMO DE ÁGUA A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00007"},
    {"code": "2249", "name": "TELEFONE A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00008"},
    {"code": "2250", "name": "CARTÃO DE CRÉDITO A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00009"},
    {"code": "2251", "name": "CHEQUES A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00010"},
    {"code": "2252", "name": "OUTRAS CONTAS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.06.01.00011"},
    {"code": "2272", "name": "ICMS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00001"},
    {"code": "2273", "name": "ISS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00002"},
    {"code": "2274", "name": "PIS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00003"},
    {"code": "2275", "name": "COFINS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00004"},
    {"code": "2276", "name": "IRPJ PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00005"},
    {"code": "2277", "name": "CSLL PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00006"},
    {"code": "2278", "name": "IPI PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00007"},
    {"code": "2279", "name": "SIMPLES NACIONAL PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00008"},
    {"code": "2280", "name": "INSS DESONERAÇÃO PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.01.00009"},
    {"code": "2281", "name": "PARCELAMENTO ESPECIAL PERT", "grupo": "Passivo/PL", "mask": "2.1.07.01.00010"},
    {"code": "2282", "name": "PARCELAMENTO CONVENCIONAL PGFN", "grupo": "Passivo/PL", "mask": "2.1.07.01.00011"},
    {"code": "2283", "name": "PARCELAMENTOS MUNICIPAIS", "grupo": "Passivo/PL", "mask": "2.1.07.01.00012"},
    {"code": "2284", "name": "(-) JUROS A APROPRIAR S/ OBRIGAÇÕES MUNICIPAIS", "grupo": "Passivo/PL", "mask": "2.1.07.01.00013"},
    {"code": "2285", "name": "(-) JUROS A APROPRIAR S/ OBRIGAÇÕES FEDERAIS", "grupo": "Passivo/PL", "mask": "2.1.07.01.00014"},
    {"code": "2293", "name": "INSS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.02.00001"},
    {"code": "2294", "name": "FGTS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.02.00002"},
    {"code": "2295", "name": "IRRF S/ FOLHA PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.1.07.02.00003"},
    {"code": "2296", "name": "PROCESSOS TRABALHISTAS", "grupo": "Passivo/PL", "mask": "2.1.07.02.00004"},
    {"code": "2297", "name": "PARCELAMENTO ESPECIAL PERT", "grupo": "Passivo/PL", "mask": "2.1.07.02.00005"},
    {"code": "2298", "name": "PARCELAMENTO CONVENCIONAL PGFN", "grupo": "Passivo/PL", "mask": "2.1.07.02.00006"},
    {"code": "2299", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.07.02.00007"},
    {"code": "2308", "name": "OUTROS FORNECEDORES A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.07.03.00001"},
    {"code": "2309", "name": "DEMAIS PROCESSOS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.07.03.00002"},
    {"code": "2319", "name": "ARRENDAMENTO FINANCEIRO - IMOVEIS IFRS 16", "grupo": "Passivo/PL", "mask": "2.1.08.01.00001"},
    {"code": "2320", "name": "(-) JUROS S/ ARRENDAMENTO IMOVEL IFRS 16", "grupo": "Passivo/PL", "mask": "2.1.08.01.00002"},
    {"code": "2321", "name": "(-) AMORTIZAÇÃO ARRENDAMENTO FINANCEIRO - IMOVEIS IFRS 16", "grupo": "Passivo/PL", "mask": "2.1.08.01.00003"},
    {"code": "2341", "name": "ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Passivo/PL", "mask": "2.1.08.02.00001"},
    {"code": "2342", "name": "(-) JUROS S/ ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Passivo/PL", "mask": "2.1.08.02.00002"},
    {"code": "2343", "name": "(-) AMORTIZAÇÃO ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Passivo/PL", "mask": "2.1.08.02.00003"},
    {"code": "2363", "name": "EMPRÉSTIMOS ENTRE PARTES LIGADAS", "grupo": "Passivo/PL", "mask": "2.1.09.01.00001"},
    {"code": "2373", "name": "DEBÊNTURES CONVERSÍVEIS EM AÇÕES", "grupo": "Passivo/PL", "mask": "2.1.10.01.00001"},
    {"code": "2383", "name": "DEBÊNTURES NÃO CONVERSÍVEIS", "grupo": "Passivo/PL", "mask": "2.1.10.02.00001"},
    {"code": "2393", "name": "DIVIDENDOS PROPOSTOS", "grupo": "Passivo/PL", "mask": "2.1.11.01.00001"},
    {"code": "2394", "name": "DIVIDENDOS A PAGAR", "grupo": "Passivo/PL", "mask": "2.1.11.01.00002"},
    {"code": "2403", "name": "PARTICIPAÇÕES PROPOSTA A ADMINISTRADORES", "grupo": "Passivo/PL", "mask": "2.1.11.02.00001"},
    {"code": "2404", "name": "PARTICIPAÇÕES PROPOSTA A EMPREGADOS", "grupo": "Passivo/PL", "mask": "2.1.11.02.00002"},
    {"code": "2409", "name": "JUROS SOBRE CAPITAL PRÓPRIO", "grupo": "Passivo/PL", "mask": "2.1.11.03.00001"},
    {"code": "2414", "name": "IMPOSTO DE RENDA DIFERIDO", "grupo": "Passivo/PL", "mask": "2.1.12.01.00001"},
    {"code": "2415", "name": "CONTRIBUIÇÃO SOCIAL DIFERIDA", "grupo": "Passivo/PL", "mask": "2.1.12.01.00002"},
    {"code": "2416", "name": "PIS DIFERIDO", "grupo": "Passivo/PL", "mask": "2.1.12.01.00003"},
    {"code": "2417", "name": "COFINS DIFERIDO", "grupo": "Passivo/PL", "mask": "2.1.12.01.00004"},
    {"code": "2418", "name": "ISS DIFERIDO", "grupo": "Passivo/PL", "mask": "2.1.12.01.00005"},
    {"code": "2419", "name": "INSS DIFERIDO", "grupo": "Passivo/PL", "mask": "2.1.12.01.00006"},
    {"code": "2420", "name": "OUTROS IMPOSTOS DIFERIDO", "grupo": "Passivo/PL", "mask": "2.1.12.01.00007"},
    {"code": "2425", "name": "RECEITA DE VENDAS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.13.01.00001"},
    {"code": "2426", "name": "IMPOSTOS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.13.01.00002"},
    {"code": "2436", "name": "RECEITAS FINANCEIRAS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.13.02.00001"},
    {"code": "2437", "name": "IMPOSTOS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.1.13.02.00002"},
    {"code": "2447", "name": "CONTA CORRENTE - PESSOAS LIGADAS - PAÍS", "grupo": "Passivo/PL", "mask": "2.1.14.01.00001"},
    {"code": "2448", "name": "CONTA CORRENTE - PARTES NÃO RELACIONADAS", "grupo": "Passivo/PL", "mask": "2.1.14.01.00002"},
    {"code": "2449", "name": "CONTA CORRENTE - EMPRESAS LIGADAS - EXT.", "grupo": "Passivo/PL", "mask": "2.1.14.01.00003"},
    {"code": "2469", "name": "FORNECEDORES NACIONAIS", "grupo": "Passivo/PL", "mask": "2.2.01.01.00001"},
    {"code": "2479", "name": "FORNECEDORES ESTRANGEIROS", "grupo": "Passivo/PL", "mask": "2.2.01.02.00001"},
    {"code": "2489", "name": "BANCO DO BRASIL", "grupo": "Passivo/PL", "mask": "2.2.02.01.00001"},
    {"code": "2490", "name": "BANCO BRADESCO", "grupo": "Passivo/PL", "mask": "2.2.02.01.00002"},
    {"code": "2491", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Passivo/PL", "mask": "2.2.02.01.00003"},
    {"code": "2492", "name": "BANCO SANTANDER", "grupo": "Passivo/PL", "mask": "2.2.02.01.00004"},
    {"code": "2493", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Passivo/PL", "mask": "2.2.02.01.00005"},
    {"code": "2494", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.02.01.00006"},
    {"code": "2503", "name": "DUPLICATAS DESCONTADAS", "grupo": "Passivo/PL", "mask": "2.2.02.02.00001"},
    {"code": "2504", "name": "EMPRÉSTIMOS DE TERCEIROS", "grupo": "Passivo/PL", "mask": "2.2.02.02.00002"},
    {"code": "2505", "name": "EMPRÉSTIMOS SÓCIOS", "grupo": "Passivo/PL", "mask": "2.2.02.02.00003"},
    {"code": "2506", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.02.02.00004"},
    {"code": "2515", "name": "BANCO DO BRASIL", "grupo": "Passivo/PL", "mask": "2.2.02.03.00001"},
    {"code": "2516", "name": "BANCO BRADESCO", "grupo": "Passivo/PL", "mask": "2.2.02.03.00002"},
    {"code": "2517", "name": "BANCO ITAÚ UNIBANCO", "grupo": "Passivo/PL", "mask": "2.2.02.03.00003"},
    {"code": "2518", "name": "BANCO SANTANDER", "grupo": "Passivo/PL", "mask": "2.2.02.03.00004"},
    {"code": "2519", "name": "CAIXA ECONÔMICA FEDERAL", "grupo": "Passivo/PL", "mask": "2.2.02.03.00005"},
    {"code": "2520", "name": "BANCO VOLKSWAGEN", "grupo": "Passivo/PL", "mask": "2.2.02.03.00006"},
    {"code": "2521", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.02.03.00007"},
    {"code": "2531", "name": "EMPRÉSTIMOS ESTRANGEIROS", "grupo": "Passivo/PL", "mask": "2.2.02.04.00001"},
    {"code": "2532", "name": "FINANCIAMENTOS ESTRANGEIROS", "grupo": "Passivo/PL", "mask": "2.2.02.04.00002"},
    {"code": "2533", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.02.04.00003"},
    {"code": "2542", "name": "ADIANTAMENTO DE CLIENTES", "grupo": "Passivo/PL", "mask": "2.2.03.01.00001"},
    {"code": "2543", "name": "ADIANTAMENTO DE CLIENTES NO EXTERIOR", "grupo": "Passivo/PL", "mask": "2.2.03.01.00002"},
    {"code": "2544", "name": "RECEBIMENTOS A IDENTIFICAR", "grupo": "Passivo/PL", "mask": "2.2.03.01.00003"},
    {"code": "2554", "name": "ALUGUEL A PAGAR", "grupo": "Passivo/PL", "mask": "2.2.03.02.00001"},
    {"code": "2555", "name": "SEGUROS A PAGAR", "grupo": "Passivo/PL", "mask": "2.2.03.02.00002"},
    {"code": "2556", "name": "OUTRAS CONTAS A PAGAR", "grupo": "Passivo/PL", "mask": "2.2.03.02.00003"},
    {"code": "2566", "name": "ICMS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00001"},
    {"code": "2567", "name": "ISS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00002"},
    {"code": "2568", "name": "PIS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00003"},
    {"code": "2569", "name": "COFINS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00004"},
    {"code": "2570", "name": "IRPJ PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00005"},
    {"code": "2571", "name": "CSLL PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00006"},
    {"code": "2572", "name": "IPI PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00007"},
    {"code": "2573", "name": "SIMPLES NACIONAL PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00008"},
    {"code": "2574", "name": "INSS DESONERAÇÃO PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.01.00009"},
    {"code": "2575", "name": "PARCELAMENTO ESPECIAL PERT", "grupo": "Passivo/PL", "mask": "2.2.04.01.00010"},
    {"code": "2576", "name": "PARCELAMENTO CONVENCIONAL PGFN", "grupo": "Passivo/PL", "mask": "2.2.04.01.00011"},
    {"code": "2577", "name": "PARCELAMENTOS MUNICIPAIS", "grupo": "Passivo/PL", "mask": "2.2.04.01.00012"},
    {"code": "2578", "name": "(-) JUROS A APROPRIAR S/ OBRIGAÇÕES MUNICIPAIS", "grupo": "Passivo/PL", "mask": "2.2.04.01.00013"},
    {"code": "2579", "name": "(-) JUROS A APROPRIAR S/ OBRIGAÇÕES FEDERAIS", "grupo": "Passivo/PL", "mask": "2.2.04.01.00014"},
    {"code": "2587", "name": "INSS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.02.00001"},
    {"code": "2588", "name": "FGTS PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.02.00002"},
    {"code": "2589", "name": "IRRF S/ FOLHA PARCELADO A RECOLHER", "grupo": "Passivo/PL", "mask": "2.2.04.02.00003"},
    {"code": "2590", "name": "PROCESSOS TRABALHISTAS", "grupo": "Passivo/PL", "mask": "2.2.04.02.00004"},
    {"code": "2591", "name": "PARCELAMENTO ESPECIAL PERT", "grupo": "Passivo/PL", "mask": "2.2.04.02.00005"},
    {"code": "2592", "name": "PARCELAMENTO CONVENCIONAL PGFN", "grupo": "Passivo/PL", "mask": "2.2.04.02.00006"},
    {"code": "2593", "name": "(-) JUROS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.04.02.00007"},
    {"code": "2602", "name": "FORNECEDORES A PAGAR", "grupo": "Passivo/PL", "mask": "2.2.04.03.00001"},
    {"code": "2603", "name": "DEMAIS PROCESSOS A PAGAR", "grupo": "Passivo/PL", "mask": "2.2.04.03.00002"},
    {"code": "2613", "name": "PROVISÃO PARA CUSTOS E DESPESAS", "grupo": "Passivo/PL", "mask": "2.2.04.04.00001"},
    {"code": "2614", "name": "PROVISÃO PARA CONTINGÊNCIAS", "grupo": "Passivo/PL", "mask": "2.2.04.04.00002"},
    {"code": "2615", "name": "PROVISÃO PARA COMISSÕES DE TERCEIROS", "grupo": "Passivo/PL", "mask": "2.2.04.04.00003"},
    {"code": "2616", "name": "PROVISÕES DE NATUREZA FISCAL", "grupo": "Passivo/PL", "mask": "2.2.04.04.00004"},
    {"code": "2617", "name": "PROVISÕES DE NATUREZA CÍVEL", "grupo": "Passivo/PL", "mask": "2.2.04.04.00005"},
    {"code": "2618", "name": "CPC 47 - PASSIVOS DE CONTRATO", "grupo": "Passivo/PL", "mask": "2.2.04.04.00006"},
    {"code": "2619", "name": "(-) JUROS A APROPRIAR AVP", "grupo": "Passivo/PL", "mask": "2.2.04.04.00007"},
    {"code": "2629", "name": "ARRENDAMENTO FINANCEIRO - IMOVEIS IFRS 16", "grupo": "Passivo/PL", "mask": "2.2.05.01.00001"},
    {"code": "2630", "name": "(-) JUROS S/ ARRENDAMENTO IMOVEL IFRS 16", "grupo": "Passivo/PL", "mask": "2.2.05.01.00002"},
    {"code": "2631", "name": "(-) AMORTIZAÇÃO ARRENDAMENTO FINANCEIRO - IMOVEIS IFRS 16", "grupo": "Passivo/PL", "mask": "2.2.05.01.00003"},
    {"code": "2641", "name": "ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Passivo/PL", "mask": "2.2.05.02.00001"},
    {"code": "2642", "name": "(-) JUROS S/ ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Passivo/PL", "mask": "2.2.05.02.00002"},
    {"code": "2643", "name": "(-) AMORTIZAÇÃO ARRENDAMENTO OPERACIONAL - IFRS 16", "grupo": "Passivo/PL", "mask": "2.2.05.02.00003"},
    {"code": "2653", "name": "CONTA CORRENTE - PESSOAS LIGADAS - PAÍS", "grupo": "Passivo/PL", "mask": "2.2.06.01.00001"},
    {"code": "2654", "name": "CONTA CORRENTE - PARTES NÃO RELACIONADAS", "grupo": "Passivo/PL", "mask": "2.2.06.01.00002"},
    {"code": "2655", "name": "CONTA CORRENTE - EMPRESAS LIGADAS - EXT.", "grupo": "Passivo/PL", "mask": "2.2.06.01.00003"},
    {"code": "2675", "name": "EMPRÉSTIMOS ENTRE PARTES LIGADAS", "grupo": "Passivo/PL", "mask": "2.2.07.01.00001"},
    {"code": "2676", "name": "EMPRÉSTIMOS ENTRE PARTES LIGADAS - I", "grupo": "Passivo/PL", "mask": "2.2.07.01.00002"},
    {"code": "2677", "name": "EMPRÉSTIMOS ENTRE PARTES LIGADAS - II", "grupo": "Passivo/PL", "mask": "2.2.07.01.00003"},
    {"code": "2678", "name": "EMPRÉSTIMOS ENTRE PARTES LIGADAS - III", "grupo": "Passivo/PL", "mask": "2.2.07.01.00004"},
    {"code": "2698", "name": "DEBÊNTURES CONVERSÍVEIS EM AÇÕES", "grupo": "Passivo/PL", "mask": "2.2.08.01.00001"},
    {"code": "2708", "name": "DEBÊNTURES NÃO CONVERSÍVEIS", "grupo": "Passivo/PL", "mask": "2.2.07.02.00001"},
    {"code": "2718", "name": "DIVIDENDOS PROPOSTOS", "grupo": "Passivo/PL", "mask": "2.2.09.01.00001"},
    {"code": "2728", "name": "PARTICIPAÇÕES PROPOSTA A ADMINISTRADORES", "grupo": "Passivo/PL", "mask": "2.2.09.02.00001"},
    {"code": "2729", "name": "PARTICIPAÇÕES PROPOSTA A EMPREGADOS", "grupo": "Passivo/PL", "mask": "2.2.09.02.00002"},
    {"code": "2739", "name": "JUROS SOBRE CAPITAL PRÓPRIO A PAGAR", "grupo": "Passivo/PL", "mask": "2.2.09.03.00001"},
    {"code": "2749", "name": "IMPOSTO DE RENDA DIFERIDO", "grupo": "Passivo/PL", "mask": "2.2.10.01.00001"},
    {"code": "2750", "name": "CONTRIBUIÇÃO SOCIAL DIFERIDA", "grupo": "Passivo/PL", "mask": "2.2.10.01.00002"},
    {"code": "2760", "name": "RECEITA DE VENDAS A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.11.01.00001"},
    {"code": "2761", "name": "IMPOSTOS SOBRE RECEITA A APROPRIAR", "grupo": "Passivo/PL", "mask": "2.2.11.01.00002"},
    {"code": "2771", "name": "PASSIVOS FINANCEIROS DE LONGO PRAZO", "grupo": "Passivo/PL", "mask": "2.2.12.01.00001"},
    {"code": "2772", "name": "PROVISÕES DE IMPOSTOS  DE LONGO PRAZO", "grupo": "Passivo/PL", "mask": "2.2.12.01.00002"},
    {"code": "2773", "name": "AJUSTE A VALOR PRESENTE", "grupo": "Passivo/PL", "mask": "2.2.12.01.00003"},
    {"code": "2783", "name": "VENDAS PARA ENTREGA FUTURA", "grupo": "Passivo/PL", "mask": "2.2.99.01.00001"},
    {"code": "2784", "name": "REMESSA AO ESTABELECIMENTO", "grupo": "Passivo/PL", "mask": "2.2.99.01.00002"},
    {"code": "2786", "name": "CONSIGNAÇÃO", "grupo": "Passivo/PL", "mask": "2.2.99.01.00004"},
    {"code": "2793", "name": "CAPITAL SOCIAL SUBSCRITO", "grupo": "Passivo/PL", "mask": "2.3.01.01.00001"},
    {"code": "2794", "name": "CAPITAL SOCIAL SUBSCRITO - SOCIO I", "grupo": "Passivo/PL", "mask": "2.3.01.01.00002"},
    {"code": "2795", "name": "CAPITAL SOCIAL SUBSCRITO - SOCIO II", "grupo": "Passivo/PL", "mask": "2.3.01.01.00003"},
    {"code": "2796", "name": "CAPITAL SOCIAL SUBSCRITO - SÓCIO NO EXTERIOR", "grupo": "Passivo/PL", "mask": "2.3.01.01.00004"},
    {"code": "2806", "name": "CAPITAL A INTEGRALIZAR", "grupo": "Passivo/PL", "mask": "2.3.01.02.00001"},
    {"code": "2807", "name": "CAPITAL A INTEGRALIZAR - SOCIO I", "grupo": "Passivo/PL", "mask": "2.3.01.02.00002"},
    {"code": "2808", "name": "CAPITAL A INTEGRALIZAR - SOCIO II", "grupo": "Passivo/PL", "mask": "2.3.01.02.00003"},
    {"code": "2809", "name": "CAPITAL A INTEGRALIZAR - SOCIO NO EXTERIOR", "grupo": "Passivo/PL", "mask": "2.3.01.02.00004"},
    {"code": "2819", "name": "AFAC - ADIANTAMENTO PARA FUTURO AUMENTO  DE CAPITAL", "grupo": "Passivo/PL", "mask": "2.3.01.03.00001"},
    {"code": "2820", "name": "AFAC - ADIANTAMENTO PARA FUTURO AUMENTO  DE CAPITAL - SÓCIO I", "grupo": "Passivo/PL", "mask": "2.3.01.03.00002"},
    {"code": "2821", "name": "AFAC - ADIANTAMENTO PARA FUTURO AUMENTO  DE CAPITAL - SÓCIO II", "grupo": "Passivo/PL", "mask": "2.3.01.03.00003"},
    {"code": "2822", "name": "AFAC - ADIANTAMENTO PARA FUTURO AUMENTO  DE CAPITAL - SÓCIO NO EXTERIOR", "grupo": "Passivo/PL", "mask": "2.3.01.03.00004"},
    {"code": "2832", "name": "PRÊMIO NA EMISSÃO DE DEBÊNTURES", "grupo": "Passivo/PL", "mask": "2.3.02.01.00001"},
    {"code": "2833", "name": "ÁGIO NA EMISSÃO DE AÇÕES", "grupo": "Passivo/PL", "mask": "2.3.02.01.00002"},
    {"code": "2834", "name": "(-) AÇÕES/QUOTAS EM TESOURARIA", "grupo": "Passivo/PL", "mask": "2.3.02.01.00003"},
    {"code": "2844", "name": "AJUSTE DE AVALIAÇÃO PATRIMONIAL DE ATIVOS", "grupo": "Passivo/PL", "mask": "2.3.02.02.00001"},
    {"code": "2845", "name": "AJUSTE DE AVALIAÇÃO PATRIMONIAL DE PASSIVOS", "grupo": "Passivo/PL", "mask": "2.3.02.02.00002"},
    {"code": "2855", "name": "RESERVA ESTATUTÁRIA", "grupo": "Passivo/PL", "mask": "2.3.02.03.00001"},
    {"code": "2856", "name": "RESERVA PARA CONTINGÊNCIAS", "grupo": "Passivo/PL", "mask": "2.3.02.03.00002"},
    {"code": "2857", "name": "RESERVA DE LUCROS A REALIZAR", "grupo": "Passivo/PL", "mask": "2.3.02.03.00003"},
    {"code": "2858", "name": "RESERVA ESPECIAL", "grupo": "Passivo/PL", "mask": "2.3.02.03.00004"},
    {"code": "2859", "name": "RESERVA LEGAL", "grupo": "Passivo/PL", "mask": "2.3.02.03.00005"},
    {"code": "2860", "name": "RESERVA PARA NOVOS INVESTIMENTOS", "grupo": "Passivo/PL", "mask": "2.3.02.03.00006"},
    {"code": "2861", "name": "REVERSÃO DE RESERVAS", "grupo": "Passivo/PL", "mask": "2.3.02.03.00007"},
    {"code": "2862", "name": "PARCELA DOS LUCROS INCORPORADAS AO CAPITAL", "grupo": "Passivo/PL", "mask": "2.3.02.03.00008"},
    {"code": "2870", "name": "LUCROS ACUMULADOS", "grupo": "Passivo/PL", "mask": "2.3.03.01.00001"},
    {"code": "2871", "name": "(-) PREJUÍZOS ACUMULADOS", "grupo": "Passivo/PL", "mask": "2.3.03.01.00002"},
    {"code": "2872", "name": "LUCRO DO EXERCÍCIO EM CURSO", "grupo": "Passivo/PL", "mask": "2.3.03.01.00003"},
    {"code": "2873", "name": "(-) PREJUÍZO DO EXERCÍCIO EM CURSO", "grupo": "Passivo/PL", "mask": "2.3.03.01.00004"},
    {"code": "2874", "name": "(-) DISTRIBUIÇÃO DE LUCROS", "grupo": "Passivo/PL", "mask": "2.3.03.01.00005"},
    {"code": "2875", "name": "AJUSTES DE EXERCÍCIOS ANTERIORES", "grupo": "Passivo/PL", "mask": "2.3.03.01.00006"},
    {"code": "2876", "name": "RESULTADO DO PERÍODO", "grupo": "Passivo/PL", "mask": "2.3.03.01.00007"},
    {"code": "2877", "name": "AJUSTES CREDORES DE PERÍODOS-BASE ANTERIORES", "grupo": "Passivo/PL", "mask": "2.3.03.01.00008"},
    {"code": "2878", "name": "(-) AJUSTES DEVEDORES DE PERÍODOS-BASE ANTERIORES", "grupo": "Passivo/PL", "mask": "2.3.03.01.00009"},
    {"code": "2879", "name": "PARCELA DOS LUCROS INCORPORADAS AO CAPITAL", "grupo": "Passivo/PL", "mask": "2.3.03.01.00010"},
    {"code": "2885", "name": "(-) DÉFICITS ACUMULADOS", "grupo": "Passivo/PL", "mask": "2.3.04.01.00001"},
    {"code": "2886", "name": "SUPERÁVITS ACUMULADOS", "grupo": "Passivo/PL", "mask": "2.3.04.01.00002"},
    {"code": "2887", "name": "(-) DÉFICIT DO PERÍODO", "grupo": "Passivo/PL", "mask": "2.3.04.01.00003"},
    {"code": "2888", "name": "SUPERÁVIT DO PERÍODO", "grupo": "Passivo/PL", "mask": "2.3.04.01.00004"},
    {"code": "2889", "name": "AJUSTES DE CONVERSÃO", "grupo": "Passivo/PL", "mask": "2.3.05.01.00001"},
    {"code": "2890", "name": "PERDAS EM PLANOS DE PENSÃO", "grupo": "Passivo/PL", "mask": "2.3.05.01.00002"},
    {"code": "2891", "name": "AJUSTES NO VALOR JUSTO DE INVESTIMENTOS", "grupo": "Passivo/PL", "mask": "2.3.05.01.00003"},
    {"code": "2892", "name": "AJUSTES DE HEDGE", "grupo": "Passivo/PL", "mask": "2.3.05.01.00004"},
    {"code": "2893", "name": "OUTROS RESULTADOS ABRANGENTES", "grupo": "Passivo/PL", "mask": "2.3.05.01.00005"},
    {"code": "2894", "name": "RESULTADOS ABRANGENTES - II", "grupo": "Passivo/PL", "mask": "2.3.05.01.00006"},
    {"code": "2895", "name": "RESULTADOS ABRANGENTES - II", "grupo": "Passivo/PL", "mask": "2.3.05.01.00007"},
    {"code": "2905", "name": "OUTRAS CONTAS DO PATRIMONIO LÍQUIDO", "grupo": "Passivo/PL", "mask": "2.3.06.01.00001"},
    # ── Resultado - Receitas (3.x) ────────────────────────────────────────────
    {"code": "3001", "name": "VENDA DE PRODUTOS", "grupo": "Resultado", "mask": "3.1.01.01.00001"},
    {"code": "3002", "name": "VENDA DE PRODUTOS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.01.01.00002"},
    {"code": "3022", "name": "VENDA DE MERCADORIAS", "grupo": "Resultado", "mask": "3.1.01.02.00001"},
    {"code": "3023", "name": "VENDA DE MERCADORIAS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.01.02.00002"},
    {"code": "3024", "name": "VENDA DE COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "3.1.01.02.00003"},
    {"code": "3025", "name": "(-) CPC 47 - OBRIGAÇÕES DE DESEMPENHO", "grupo": "Resultado", "mask": "3.1.01.02.00004"},
    {"code": "3045", "name": "SERVIÇOS PRESTADOS", "grupo": "Resultado", "mask": "3.1.01.03.00001"},
    {"code": "3046", "name": "SERVIÇOS PRESTADOS MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.01.03.00002"},
    {"code": "3047", "name": "PRESTAÇÃO DE SERVIÇOS DE TRANSPORTES", "grupo": "Resultado", "mask": "3.1.01.03.00003"},
    {"code": "3048", "name": "PRESTAÇÃO DE SERVIÇOS DE TELECOMUNICAÇÃO", "grupo": "Resultado", "mask": "3.1.01.03.00004"},
    {"code": "3049", "name": "VENDA DE ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "3.1.01.03.00005"},
    {"code": "3069", "name": "RECEITA DE LOCAÇÃO", "grupo": "Resultado", "mask": "3.1.01.04.00001"},
    {"code": "3070", "name": "RECEITA DE LOCAÇÃO NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.01.04.00002"},
    {"code": "3090", "name": "VENDAS DE ANIMAIS NO MERCADO INTERNO", "grupo": "Resultado", "mask": "3.1.01.05.00001"},
    {"code": "3091", "name": "VENDAS DE ANIMAIS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.01.05.00002"},
    {"code": "3092", "name": "VENDAS PRODUTOS PECUÁRIOS MERC.INTERNO", "grupo": "Resultado", "mask": "3.1.01.05.00003"},
    {"code": "3093", "name": "VENDAS PRODUTOS PECUÁRIOS MERC.EXTERNO", "grupo": "Resultado", "mask": "3.1.01.05.00004"},
    {"code": "3094", "name": "VENDAS PRODUTOS AGRÍCOLAS MERC.INTERNO", "grupo": "Resultado", "mask": "3.1.01.05.00005"},
    {"code": "3095", "name": "VENDAS PRODUTOS AGRÍCOLAS MERC.EXTERNO", "grupo": "Resultado", "mask": "3.1.01.05.00006"},
    {"code": "3105", "name": "RECEITA OUTRAS ATIVIDADES", "grupo": "Resultado", "mask": "3.1.01.06.00001"},
    {"code": "3106", "name": "RECEITA OUTRAS ATIVIDADES NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.01.06.00002"},
    {"code": "3116", "name": "LOCAÇÃO DE BENS MÓVEIS E IMÓVEIS", "grupo": "Resultado", "mask": "3.1.01.07.00001"},
    {"code": "3117", "name": "SUBCONTRATAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "3.1.01.07.00002"},
    {"code": "3118", "name": "CONTRIBUIÇÕES DE ASSOCIADOS - PJ", "grupo": "Resultado", "mask": "3.1.01.07.00003"},
    {"code": "3119", "name": "CONTRIBUIÇÕES DE ASSOCIADOS - PF", "grupo": "Resultado", "mask": "3.1.01.07.00004"},
    {"code": "3120", "name": "BENEFICIAMENTO", "grupo": "Resultado", "mask": "3.1.01.07.00005"},
    {"code": "3121", "name": "VENDA DE UNIDADES IMOBILIÁRIAS", "grupo": "Resultado", "mask": "3.1.01.07.00006"},
    {"code": "3122", "name": "CONTRATO DE CONSTRUÇÃO", "grupo": "Resultado", "mask": "3.1.01.07.00007"},
    {"code": "3123", "name": "SECURITIZAÇÃO DE CRÉDITOS", "grupo": "Resultado", "mask": "3.1.01.07.00008"},
    {"code": "3124", "name": "DOAÇÕES E SUBVENÇÕES PARA INVESTIMENTOS", "grupo": "Resultado", "mask": "3.1.01.07.00009"},
    {"code": "3125", "name": "OUTRAS RECEITAS DA ATIVIDADE GERAL", "grupo": "Resultado", "mask": "3.1.01.07.00010"},
    {"code": "3135", "name": "ALUGUÉIS E ARRENDAMENTOS", "grupo": "Resultado", "mask": "3.1.02.01.00001"},
    {"code": "3136", "name": "DIVIDENDOS E LUCROS RECEBIDOS", "grupo": "Resultado", "mask": "3.1.02.01.00002"},
    {"code": "3137", "name": "AMORTIZAÇÃO DE DESÁGIO", "grupo": "Resultado", "mask": "3.1.02.01.00003"},
    {"code": "3138", "name": "VENDA DE SUCATA", "grupo": "Resultado", "mask": "3.1.02.01.00004"},
    {"code": "3139", "name": "BRINDES E BONIFICAÇÕES", "grupo": "Resultado", "mask": "3.1.02.01.00005"},
    {"code": "3140", "name": "AMOSTRA GRÁTIS", "grupo": "Resultado", "mask": "3.1.02.01.00006"},
    {"code": "3141", "name": "PROVISÃO DE RECEITAS", "grupo": "Resultado", "mask": "3.1.02.01.00007"},
    {"code": "3142", "name": "RENDAS DIVERSAS", "grupo": "Resultado", "mask": "3.1.02.01.00008"},
    {"code": "3143", "name": "RECUPERAÇÃO DE DESPESAS - TRIBUTÁVEL", "grupo": "Resultado", "mask": "3.1.02.01.00009"},
    {"code": "3144", "name": "REVERSÃO DE PROVISÕES", "grupo": "Resultado", "mask": "3.1.02.01.00010"},
    {"code": "3145", "name": "DESCONTOS", "grupo": "Resultado", "mask": "3.1.02.01.00011"},
    {"code": "3165", "name": "(-) DEVOLUÇÃO DE VENDA DE PRODUTOS", "grupo": "Resultado", "mask": "3.1.03.01.00001"},
    {"code": "3166", "name": "(-) DEVOLUÇÃO VENDA DE PRODUTOS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.03.01.00002"},
    {"code": "3167", "name": "(-) DEVOLUÇÃO DE VENDA DE MERCADORIAS", "grupo": "Resultado", "mask": "3.1.03.01.00003"},
    {"code": "3168", "name": "(-) DEVOLUÇÃO DE VENDA DE MERCADORIAS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.03.01.00004"},
    {"code": "3169", "name": "(-) DEVOLUÇÃO DE VENDA DE SERVIÇOS", "grupo": "Resultado", "mask": "3.1.03.01.00005"},
    {"code": "3170", "name": "(-) DEVOLUÇÃO DE VENDA DE SERVIÇOS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.03.01.00006"},
    {"code": "3171", "name": "(-) ANULAÇÃO DE VALOR RELATIVO À PREST. SERV. DE TRANSPORTE", "grupo": "Resultado", "mask": "3.1.03.01.00007"},
    {"code": "3172", "name": "(-) ANULAÇÃO DE VALOR RELATIVO À PREST. SERV. DE COMUNICAÇÃO", "grupo": "Resultado", "mask": "3.1.03.01.00008"},
    {"code": "3173", "name": "(-) ANULAÇÃO DE VALOR RELATIVO À PREST. SERV. DE CONSUMO DE ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "3.1.03.01.00009"},
    {"code": "3174", "name": "(-) DEVOLUÇÃO DE VENDA DE COMBUSTÍVEL", "grupo": "Resultado", "mask": "3.1.03.01.00010"},
    {"code": "3181", "name": "(-) DESCONTO VENDA DE PRODUTOS", "grupo": "Resultado", "mask": "3.1.03.02.00001"},
    {"code": "3182", "name": "(-) DESCONTO VENDA DE PRODUTO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.03.02.00002"},
    {"code": "3183", "name": "(-) DESCONTO VENDA DE MERCADORIA", "grupo": "Resultado", "mask": "3.1.03.02.00003"},
    {"code": "3184", "name": "(-) DESCONTO VENDA DE MERCADORIAS MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.03.02.00004"},
    {"code": "3185", "name": "(-) DESCONTO DE VENDA DE SERVIÇOS", "grupo": "Resultado", "mask": "3.1.03.02.00005"},
    {"code": "3186", "name": "(-) DESCONTO DE VENDA DE SERVIÇOS NO MERCADO EXTERNO", "grupo": "Resultado", "mask": "3.1.03.02.00006"},
    {"code": "3187", "name": "(-) ANULAÇÃO VALOR REFERENTE A VENDA OU SERVIÇOS PRESTADOS", "grupo": "Resultado", "mask": "3.1.03.02.00007"},
    {"code": "3196", "name": "(-) IPI", "grupo": "Resultado", "mask": "3.1.03.03.00001"},
    {"code": "3197", "name": "(-) ICMS", "grupo": "Resultado", "mask": "3.1.03.03.00002"},
    {"code": "3198", "name": "(-) ISS", "grupo": "Resultado", "mask": "3.1.03.03.00003"},
    {"code": "3199", "name": "(-) PIS", "grupo": "Resultado", "mask": "3.1.03.03.00004"},
    {"code": "3200", "name": "(-) COFINS", "grupo": "Resultado", "mask": "3.1.03.03.00005"},
    {"code": "3201", "name": "(-) INSS SOBRE RECEITA BRUTA", "grupo": "Resultado", "mask": "3.1.03.03.00006"},
    {"code": "3202", "name": "(-) SIMPLES NACIONAL", "grupo": "Resultado", "mask": "3.1.03.03.00007"},
    {"code": "3203", "name": "(-) FIA", "grupo": "Resultado", "mask": "3.1.03.03.00008"},
    {"code": "3204", "name": "(-) ICMS SUBSTITUIÇÃO TRIBUTÁRIA", "grupo": "Resultado", "mask": "3.1.03.03.00009"},
    {"code": "3205", "name": "(-) ICMS DIFAL", "grupo": "Resultado", "mask": "3.1.03.03.00010"},
    {"code": "3206", "name": "(-) CBS", "grupo": "Resultado", "mask": "3.1.03.03.00011"},
    {"code": "3207", "name": "(-) IBS", "grupo": "Resultado", "mask": "3.1.03.03.00012"},
    {"code": "3208", "name": "(-) ICMS SOBRE OUTRAS SAÍDAS", "grupo": "Resultado", "mask": "3.1.03.03.00013"},
    {"code": "3209", "name": "(-) IPI SOBRE OUTRAS SAÍDAS", "grupo": "Resultado", "mask": "3.1.03.03.00014"},
    {"code": "3210", "name": "DEMAIS IMPOSTOS INCIDENTES SOBE A VENDAS E SERVIÇOS", "grupo": "Resultado", "mask": "3.1.03.03.00015"},
    {"code": "3221", "name": "CPC 47 - RECONHECIMENTOS E AJUSTES", "grupo": "Resultado", "mask": "3.1.03.04.00001"},
    {"code": "3222", "name": "CPC 47 - AJUSTES A VALOR PRESENTE", "grupo": "Resultado", "mask": "3.1.03.04.00002"},
    {"code": "3223", "name": "DEMAIS IMPOSTOS E CONTRIBUIÇÕES", "grupo": "Resultado", "mask": "3.1.03.04.00003"},
    # ── Resultado - Custo (4.x) ───────────────────────────────────────────────
    {"code": "4001", "name": "CPV", "grupo": "Resultado", "mask": "4.1.01.01.00001"},
    {"code": "4002", "name": "BRINDES E BONIFICAÇÕES", "grupo": "Resultado", "mask": "4.1.01.01.00002"},
    {"code": "4003", "name": "MATERIAL DE USO E CONSUMO", "grupo": "Resultado", "mask": "4.1.01.01.00003"},
    {"code": "4004", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "4.1.01.01.00004"},
    {"code": "4005", "name": "FRETES E CARRETOS", "grupo": "Resultado", "mask": "4.1.01.01.00005"},
    {"code": "4006", "name": "ÁGUA", "grupo": "Resultado", "mask": "4.1.01.01.00006"},
    {"code": "4007", "name": "ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "4.1.01.01.00007"},
    {"code": "4008", "name": "PERDA COM ESTOQUES", "grupo": "Resultado", "mask": "4.1.01.01.00008"},
    {"code": "4009", "name": "MATERIAL DE EMBALAGEM", "grupo": "Resultado", "mask": "4.1.01.01.00009"},
    {"code": "4010", "name": "ARMAZENAGEM E DEPOSITO", "grupo": "Resultado", "mask": "4.1.01.01.00010"},
    {"code": "4011", "name": "CUSTO COM INSTALAÇÕES", "grupo": "Resultado", "mask": "4.1.01.01.00011"},
    {"code": "4012", "name": "LANCHES E REFEIÇÕES", "grupo": "Resultado", "mask": "4.1.01.01.00012"},
    {"code": "4013", "name": "LANCHES E REFEIÇÕES - DIRIGENTES", "grupo": "Resultado", "mask": "4.1.01.01.00013"},
    {"code": "4014", "name": "COPA E COZINHA", "grupo": "Resultado", "mask": "4.1.01.01.00014"},
    {"code": "4015", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.1.01.01.00015"},
    {"code": "4016", "name": "FARMÁCIA", "grupo": "Resultado", "mask": "4.1.01.01.00016"},
    {"code": "4017", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.1.01.01.00017"},
    {"code": "4018", "name": "MATERIAIS DE HIGIENE E LIMPEZA", "grupo": "Resultado", "mask": "4.1.01.01.00018"},
    {"code": "4019", "name": "MATERIAIS DE ESCRITÓRIO", "grupo": "Resultado", "mask": "4.1.01.01.00019"},
    {"code": "4020", "name": "MATERIAIS DE INFORMÁTICA", "grupo": "Resultado", "mask": "4.1.01.01.00020"},
    {"code": "4022", "name": "MATERIAIS PUBLICITÁRIOS", "grupo": "Resultado", "mask": "4.1.01.01.00022"},
    {"code": "4023", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.1.01.01.00023"},
    {"code": "4024", "name": "BENS DE PEQUENO VALOR", "grupo": "Resultado", "mask": "4.1.01.01.00024"},
    {"code": "4025", "name": "CUSTO COM VEÍCULOS", "grupo": "Resultado", "mask": "4.1.01.01.00025"},
    {"code": "4026", "name": "COMPUTADORES E PERIFÉRICOS", "grupo": "Resultado", "mask": "4.1.01.01.00026"},
    {"code": "4028", "name": "BENFEITORIAS EM IMÓVEIS DE TERCEIROS", "grupo": "Resultado", "mask": "4.1.01.01.00028"},
    {"code": "4029", "name": "COMUNICAÇÕES", "grupo": "Resultado", "mask": "4.1.01.01.00029"},
    {"code": "4030", "name": "TELEFONE FIXO", "grupo": "Resultado", "mask": "4.1.01.01.00030"},
    {"code": "4031", "name": "TELEFONE MOVEL", "grupo": "Resultado", "mask": "4.1.01.01.00031"},
    {"code": "4032", "name": "INTERNET", "grupo": "Resultado", "mask": "4.1.01.01.00032"},
    {"code": "4033", "name": "SERVIÇOS PRESTADOS DE PROGRAMAÇÃO E TI", "grupo": "Resultado", "mask": "4.1.01.01.00033"},
    {"code": "4034", "name": "CUSTO COM INFRAESTRUTURA", "grupo": "Resultado", "mask": "4.1.01.01.00034"},
    {"code": "4035", "name": "CUSTO COM HONORÁRIOS", "grupo": "Resultado", "mask": "4.1.01.01.00035"},
    {"code": "4036", "name": "MANUTENÇÃO DE SOFTWARE", "grupo": "Resultado", "mask": "4.1.01.01.00036"},
    {"code": "4037", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "4.1.01.01.00037"},
    {"code": "4038", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA FISICA", "grupo": "Resultado", "mask": "4.1.01.01.00038"},
    {"code": "4039", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.1.01.01.00039"},
    {"code": "4040", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.01.01.00040"},
    {"code": "4041", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.01.01.00041"},
    {"code": "4042", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.1.01.01.00042"},
    {"code": "4043", "name": "ROYALTIES E ASSISTÊNCIA TÉCNICA PAÍS", "grupo": "Resultado", "mask": "4.1.01.01.00043"},
    {"code": "4044", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.01.01.00044"},
    {"code": "4045", "name": "INDUSTRIALIZAÇÃO EFETUADO POR TERCEIROS", "grupo": "Resultado", "mask": "4.1.01.01.00045"},
    {"code": "4046", "name": "ASSESSORIA CONTÁBIL", "grupo": "Resultado", "mask": "4.1.01.01.00046"},
    {"code": "4047", "name": "CONSULTORIA E SUPORTE TÉCNICO", "grupo": "Resultado", "mask": "4.1.01.01.00047"},
    {"code": "4048", "name": "SERVIÇOS DE ADVOGADOS", "grupo": "Resultado", "mask": "4.1.01.01.00048"},
    {"code": "4049", "name": "SERVIÇOS DE AUDITORIA", "grupo": "Resultado", "mask": "4.1.01.01.00049"},
    {"code": "4050", "name": "SERVIÇOS DE CONSULTORIA", "grupo": "Resultado", "mask": "4.1.01.01.00050"},
    {"code": "4051", "name": "SERVIÇOS DE DESPACHANTES ADUANEIRO", "grupo": "Resultado", "mask": "4.1.01.01.00051"},
    {"code": "4052", "name": "ASSESSORIA ADMINISTRATIVA", "grupo": "Resultado", "mask": "4.1.01.01.00052"},
    {"code": "4053", "name": "ADMINISTRAÇÃO DE BENS", "grupo": "Resultado", "mask": "4.1.01.01.00053"},
    {"code": "4054", "name": "SERVIÇOS E ANALISES TÉCNICAS", "grupo": "Resultado", "mask": "4.1.01.01.00054"},
    {"code": "4055", "name": "PESQUISA/PROJETO", "grupo": "Resultado", "mask": "4.1.01.01.00055"},
    {"code": "4056", "name": "CURSOS E TREINAMENTOS", "grupo": "Resultado", "mask": "4.1.01.01.00056"},
    {"code": "4057", "name": "DATILOGRAFIA", "grupo": "Resultado", "mask": "4.1.01.01.00057"},
    {"code": "4058", "name": "TRADUÇÕES E VERSÕES", "grupo": "Resultado", "mask": "4.1.01.01.00058"},
    {"code": "4059", "name": "SERVIÇOS DE INTERMEDIAÇÃO", "grupo": "Resultado", "mask": "4.1.01.01.00059"},
    {"code": "4060", "name": "ASSESSORIA E CONSULTORIA EM INFORMÁTICA", "grupo": "Resultado", "mask": "4.1.01.01.00060"},
    {"code": "4061", "name": "LICENÇA DE SOFTWARE", "grupo": "Resultado", "mask": "4.1.01.01.00061"},
    {"code": "4062", "name": "ANALISE E DESENVOLVIMENTO DE SISTEMAS", "grupo": "Resultado", "mask": "4.1.01.01.00062"},
    {"code": "4063", "name": "PROCESSAMENTO E ARMAZENAGEM DE DADOS", "grupo": "Resultado", "mask": "4.1.01.01.00063"},
    {"code": "4064", "name": "PROGRAMAÇÃO E COMUNICAÇÃO VISUAL", "grupo": "Resultado", "mask": "4.1.01.01.00064"},
    {"code": "4065", "name": "ASSISTÊNCIA TÉCNICA", "grupo": "Resultado", "mask": "4.1.01.01.00065"},
    {"code": "4067", "name": "VIAGENS E ESTADIAS", "grupo": "Resultado", "mask": "4.1.01.01.00067"},
    {"code": "4068", "name": "FEIRAS E EVENTOS", "grupo": "Resultado", "mask": "4.1.01.01.00068"},
    {"code": "4069", "name": "PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "4.1.01.01.00069"},
    {"code": "4070", "name": "SERVIÇOS GRÁFICOS", "grupo": "Resultado", "mask": "4.1.01.01.00070"},
    {"code": "4071", "name": "ANÚNCIOS E PUBLICAÇÕES", "grupo": "Resultado", "mask": "4.1.01.01.00071"},
    {"code": "4072", "name": "CARTÓRIO", "grupo": "Resultado", "mask": "4.1.01.01.00072"},
    {"code": "4073", "name": "DESPESAS COM ANIMAIS E MÉDICOS VETERINÁRIOS", "grupo": "Resultado", "mask": "4.1.01.01.00073"},
    {"code": "4074", "name": "SERVIÇOS MÉDICOS E EXAMES LABORATORIAS", "grupo": "Resultado", "mask": "4.1.01.01.00074"},
    {"code": "4075", "name": "ENGENHARIA E ARQUITETURA", "grupo": "Resultado", "mask": "4.1.01.01.00075"},
    {"code": "4076", "name": "MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.1.01.01.00076"},
    {"code": "4077", "name": "MANUTENÇÃO E CONSERVAÇÃO", "grupo": "Resultado", "mask": "4.1.01.01.00077"},
    {"code": "4078", "name": "INSTALAÇÕES E MONTAGENS", "grupo": "Resultado", "mask": "4.1.01.01.00078"},
    {"code": "4079", "name": "RECRUTAMENTO E SELEÇÃO", "grupo": "Resultado", "mask": "4.1.01.01.00079"},
    {"code": "4080", "name": "LOCAÇÃO DE MÃO DE OBRA", "grupo": "Resultado", "mask": "4.1.01.01.00080"},
    {"code": "4081", "name": "SERVIÇOS DE LIMPEZA", "grupo": "Resultado", "mask": "4.1.01.01.00081"},
    {"code": "4082", "name": "DEDETIZAÇÃO", "grupo": "Resultado", "mask": "4.1.01.01.00082"},
    {"code": "4083", "name": "SEGURANÇA", "grupo": "Resultado", "mask": "4.1.01.01.00083"},
    {"code": "4084", "name": "SEGUROS GERAIS", "grupo": "Resultado", "mask": "4.1.01.01.00084"},
    {"code": "4085", "name": "SEGUROS - RESPONSABILIDADE CÍVIL", "grupo": "Resultado", "mask": "4.1.01.01.00085"},
    {"code": "4086", "name": "SEGUROS - IMOBILIZADO", "grupo": "Resultado", "mask": "4.1.01.01.00086"},
    {"code": "4088", "name": "ESTACIONAMENTOS E PEDÁGIOS", "grupo": "Resultado", "mask": "4.1.01.01.00088"},
    {"code": "4089", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.01.01.00089"},
    {"code": "4090", "name": "ALUGUEL DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "4.1.01.01.00090"},
    {"code": "4091", "name": "ALUGUEL DE VEÍCULOS", "grupo": "Resultado", "mask": "4.1.01.01.00091"},
    {"code": "4092", "name": "ALUGUEL - PJ", "grupo": "Resultado", "mask": "4.1.01.01.00092"},
    {"code": "4093", "name": "ALUGUEL - PF", "grupo": "Resultado", "mask": "4.1.01.01.00093"},
    {"code": "4094", "name": "FOTOGRAFIA", "grupo": "Resultado", "mask": "4.1.01.01.00094"},
    {"code": "4095", "name": "BENS DE USO PERMANENTE", "grupo": "Resultado", "mask": "4.1.01.01.00095"},
    {"code": "4096", "name": "CUSTOS COM DEPRECIAÇÃO", "grupo": "Resultado", "mask": "4.1.01.01.00096"},
    {"code": "4097", "name": "CUSTOS COM AMORTIZAÇÃO", "grupo": "Resultado", "mask": "4.1.01.01.00097"},
    {"code": "4115", "name": "SALÁRIOS E ORDENADOS", "grupo": "Resultado", "mask": "4.1.01.02.00001"},
    {"code": "4116", "name": "FÉRIAS", "grupo": "Resultado", "mask": "4.1.01.02.00002"},
    {"code": "4117", "name": "13º SALÁRIO", "grupo": "Resultado", "mask": "4.1.01.02.00003"},
    {"code": "4118", "name": "HORA EXTRA", "grupo": "Resultado", "mask": "4.1.01.02.00004"},
    {"code": "4119", "name": "DESCANSO SEMANAL REMUNERADO", "grupo": "Resultado", "mask": "4.1.01.02.00005"},
    {"code": "4120", "name": "BÔNUS", "grupo": "Resultado", "mask": "4.1.01.02.00006"},
    {"code": "4121", "name": "ABONO PECUNIÁRIO", "grupo": "Resultado", "mask": "4.1.01.02.00007"},
    {"code": "4122", "name": "VERBAS RESCISÓRIAS", "grupo": "Resultado", "mask": "4.1.01.02.00008"},
    {"code": "4123", "name": "ADICIONAL NOTURNO", "grupo": "Resultado", "mask": "4.1.01.02.00009"},
    {"code": "4124", "name": "ADICIONAL DE PERICULOSIDADE", "grupo": "Resultado", "mask": "4.1.01.02.00010"},
    {"code": "4125", "name": "ADICIONAL POR TEMPO DE SERVIÇO", "grupo": "Resultado", "mask": "4.1.01.02.00011"},
    {"code": "4126", "name": "ADICIONAL DE INSALUBRIDADE", "grupo": "Resultado", "mask": "4.1.01.02.00012"},
    {"code": "4127", "name": "AUXÍLIO HOME OFFICE", "grupo": "Resultado", "mask": "4.1.01.02.00013"},
    {"code": "4128", "name": "CONTRIBUIÇÃO ASSISTENCIAL", "grupo": "Resultado", "mask": "4.1.01.02.00014"},
    {"code": "4129", "name": "INDENIZAÇÕES", "grupo": "Resultado", "mask": "4.1.01.02.00015"},
    {"code": "4130", "name": "MULTA ESTABILIDADE", "grupo": "Resultado", "mask": "4.1.01.02.00016"},
    {"code": "4131", "name": "CONTRIBUIÇÃO SINDICAL", "grupo": "Resultado", "mask": "4.1.01.02.00017"},
    {"code": "4132", "name": "ANUÊNIO E TRIÊNIO", "grupo": "Resultado", "mask": "4.1.01.02.00018"},
    {"code": "4133", "name": "QUINQUENIO", "grupo": "Resultado", "mask": "4.1.01.02.00019"},
    {"code": "4134", "name": "AUXÍLIO CRECHE", "grupo": "Resultado", "mask": "4.1.01.02.00020"},
    {"code": "4135", "name": "BOLSA GRATIFICAÇÃO", "grupo": "Resultado", "mask": "4.1.01.02.00021"},
    {"code": "4136", "name": "PROVISÕES E ENCARGOS", "grupo": "Resultado", "mask": "4.1.01.02.00022"},
    {"code": "4137", "name": "BOLSA AUXÍLIO", "grupo": "Resultado", "mask": "4.1.01.02.00023"},
    {"code": "4138", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.1.01.02.00024"},
    {"code": "4139", "name": "SERVIÇOS AUTÔNOMO", "grupo": "Resultado", "mask": "4.1.01.02.00025"},
    {"code": "4140", "name": "PARTICIPAÇÃO NOS LUCROS E RESULTADOS", "grupo": "Resultado", "mask": "4.1.01.02.00026"},
    {"code": "4141", "name": "PRÊMIOS E GRATIFICAÇÕES", "grupo": "Resultado", "mask": "4.1.01.02.00027"},
    {"code": "4142", "name": "DONATIVOS E GORJETAS", "grupo": "Resultado", "mask": "4.1.01.02.00028"},
    {"code": "4143", "name": "COMISSÕES", "grupo": "Resultado", "mask": "4.1.01.02.00029"},
    {"code": "4144", "name": "PRÓ-LABORE", "grupo": "Resultado", "mask": "4.1.01.02.00030"},
    {"code": "4145", "name": "FGTS", "grupo": "Resultado", "mask": "4.1.01.02.00031"},
    {"code": "4146", "name": "INSS", "grupo": "Resultado", "mask": "4.1.01.02.00032"},
    {"code": "4147", "name": "INSS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.1.01.02.00033"},
    {"code": "4148", "name": "INSS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.1.01.02.00034"},
    {"code": "4149", "name": "FGTS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.1.01.02.00035"},
    {"code": "4150", "name": "FGTS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.1.01.02.00036"},
    {"code": "4151", "name": "MULTA FGTS", "grupo": "Resultado", "mask": "4.1.01.02.00037"},
    {"code": "4152", "name": "AJUDA DE CUSTO", "grupo": "Resultado", "mask": "4.1.01.02.00038"},
    {"code": "4153", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.1.01.02.00039"},
    {"code": "4154", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.1.01.02.00040"},
    {"code": "4155", "name": "VALE ALIMENTAÇÃO/REFEIÇÃO", "grupo": "Resultado", "mask": "4.1.01.02.00041"},
    {"code": "4156", "name": "VALE TRANSPORTE", "grupo": "Resultado", "mask": "4.1.01.02.00042"},
    {"code": "4157", "name": "VALE COMBUSTÍVEL", "grupo": "Resultado", "mask": "4.1.01.02.00043"},
    {"code": "4158", "name": "BENEFÍCIOS TRABALHISTAS FLEXÍVEIS", "grupo": "Resultado", "mask": "4.1.01.02.00044"},
    {"code": "4159", "name": "SEGURO DE VIDA", "grupo": "Resultado", "mask": "4.1.01.02.00045"},
    {"code": "4160", "name": "ASSISTÊNCIA MÉDICA", "grupo": "Resultado", "mask": "4.1.01.02.00046"},
    {"code": "4161", "name": "ASSISTÊNCIA ODONTOLÓGICA", "grupo": "Resultado", "mask": "4.1.01.02.00047"},
    {"code": "4162", "name": "PIS SOBRE FOLHA", "grupo": "Resultado", "mask": "4.1.01.02.00048"},
    {"code": "4163", "name": "FUNDO DE APOSENTADORIA", "grupo": "Resultado", "mask": "4.1.01.02.00049"},
    {"code": "4164", "name": "PLANO DE POUPANÇA E INVESTIMENTO", "grupo": "Resultado", "mask": "4.1.01.02.00050"},
    {"code": "4165", "name": "VALE CULTURA", "grupo": "Resultado", "mask": "4.1.01.02.00051"},
    {"code": "4166", "name": "REFEIÇÕES PAT", "grupo": "Resultado", "mask": "4.1.01.02.00052"},
    {"code": "4167", "name": "PROCESSOS TRABALHISTAS", "grupo": "Resultado", "mask": "4.1.01.02.00053"},
    {"code": "4168", "name": "PREVIDÊNCIA PRIVADA", "grupo": "Resultado", "mask": "4.1.01.02.00054"},
    {"code": "4169", "name": "OUTRAS CUSTOS COM PESSOAL", "grupo": "Resultado", "mask": "4.1.01.02.00055"},
    {"code": "4170", "name": "OUTROS PROVENTOS RPA", "grupo": "Resultado", "mask": "4.1.01.02.00056"},
    {"code": "4189", "name": "SERVIÇOS PRESTADOS DE SUPORTE DE OPERAÇÕES", "grupo": "Resultado", "mask": "4.1.01.03.00001"},
    {"code": "4190", "name": "CUSTOS COM IMPORTAÇÃO", "grupo": "Resultado", "mask": "4.1.01.03.00002"},
    {"code": "4191", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.01.03.00003"},
    {"code": "4192", "name": "OUTROS CUSTOS COM SERVIÇOS", "grupo": "Resultado", "mask": "4.1.01.03.00004"},
    {"code": "4193", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.1.01.03.00005"},
    {"code": "4194", "name": "CRÉDITO DE PIS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "4.1.01.03.00006"},
    {"code": "4195", "name": "CRÉDITO DE COFINS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "4.1.01.03.00007"},
    {"code": "4196", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.01.03.00008"},
    {"code": "4197", "name": "CUSTO DA LOCAÇÃO DE BENS MÓVEIS E IMÓVEIS", "grupo": "Resultado", "mask": "4.1.01.03.00009"},
    {"code": "4198", "name": "CUSTO DA SUBCONTRATAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.01.03.00010"},
    {"code": "4199", "name": "CUSTO DA VENDA DE UNIDADES IMOBILIÁRIAS", "grupo": "Resultado", "mask": "4.1.01.03.00011"},
    {"code": "4200", "name": "CUSTO DO CONTRATO DE CONSTRUÇÃO", "grupo": "Resultado", "mask": "4.1.01.03.00012"},
    {"code": "4201", "name": "CUSTO DA SECURITIZAÇÃO DE CRÉDITOS", "grupo": "Resultado", "mask": "4.1.01.03.00013"},
    {"code": "4202", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.01.03.00014"},
    {"code": "4203", "name": "OUTROS CUSTOS DA ATIVIDADE GERAL", "grupo": "Resultado", "mask": "4.1.01.03.00015"},
    {"code": "4204", "name": "CUSTOS DIVERSOS", "grupo": "Resultado", "mask": "4.1.01.03.00016"},
    {"code": "4205", "name": "CUSTOS INDEDUTÍVEIS", "grupo": "Resultado", "mask": "4.1.01.03.00017"},
    {"code": "4206", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.01.03.00018"},
    {"code": "4207", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.01.03.00019"},
    {"code": "4208", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.1.01.03.00020"},
    {"code": "4218", "name": "MATERIAL APLICADO NO SERVIÇO", "grupo": "Resultado", "mask": "4.1.02.01.00001"},
    {"code": "4219", "name": "CUSTO COM INSTALAÇÕES", "grupo": "Resultado", "mask": "4.1.02.01.00002"},
    {"code": "4220", "name": "MATERIAL DE USO E CONSUMO", "grupo": "Resultado", "mask": "4.1.02.01.00003"},
    {"code": "4221", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "4.1.02.01.00004"},
    {"code": "4222", "name": "FRETES E CARRETOS", "grupo": "Resultado", "mask": "4.1.02.01.00005"},
    {"code": "4223", "name": "ÁGUA", "grupo": "Resultado", "mask": "4.1.02.01.00006"},
    {"code": "4224", "name": "ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "4.1.02.01.00007"},
    {"code": "4225", "name": "PERDA COM ESTOQUES", "grupo": "Resultado", "mask": "4.1.02.01.00008"},
    {"code": "4226", "name": "MATERIAL DE EMBALAGEM", "grupo": "Resultado", "mask": "4.1.02.01.00009"},
    {"code": "4227", "name": "MATERIAL DE MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.1.02.01.00010"},
    {"code": "4228", "name": "ARMAZENAGEM E DEPOSITO", "grupo": "Resultado", "mask": "4.1.02.01.00011"},
    {"code": "4229", "name": "CUSTO COM INSTALAÇÕES", "grupo": "Resultado", "mask": "4.1.02.01.00012"},
    {"code": "4230", "name": "LANCHES E REFEIÇÕES", "grupo": "Resultado", "mask": "4.1.02.01.00013"},
    {"code": "4231", "name": "LANCHES E REFEIÇÕES - DIRIGENTES", "grupo": "Resultado", "mask": "4.1.02.01.00014"},
    {"code": "4232", "name": "COPA E COZINHA", "grupo": "Resultado", "mask": "4.1.02.01.00015"},
    {"code": "4233", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.1.02.01.00016"},
    {"code": "4234", "name": "FARMÁCIA", "grupo": "Resultado", "mask": "4.1.02.01.00017"},
    {"code": "4235", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.1.02.01.00018"},
    {"code": "4236", "name": "MATERIAIS DE HIGIENE E LIMPEZA", "grupo": "Resultado", "mask": "4.1.02.01.00019"},
    {"code": "4237", "name": "MATERIAIS DE ESCRITÓRIO", "grupo": "Resultado", "mask": "4.1.02.01.00020"},
    {"code": "4238", "name": "MATERIAIS DE INFORMÁTICA", "grupo": "Resultado", "mask": "4.1.02.01.00021"},
    {"code": "4240", "name": "MATERIAIS PUBLICITÁRIOS", "grupo": "Resultado", "mask": "4.1.02.01.00023"},
    {"code": "4241", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.1.02.01.00024"},
    {"code": "4242", "name": "BENS DE PEQUENO VALOR", "grupo": "Resultado", "mask": "4.1.02.01.00025"},
    {"code": "4243", "name": "DESPESAS COM VEÍCULOS", "grupo": "Resultado", "mask": "4.1.02.01.00026"},
    {"code": "4244", "name": "COMPUTADORES E PERIFÉRICOS", "grupo": "Resultado", "mask": "4.1.02.01.00027"},
    {"code": "4246", "name": "BENFEITORIAS EM IMÓVEIS DE TERCEIROS", "grupo": "Resultado", "mask": "4.1.02.01.00029"},
    {"code": "4247", "name": "COMUNICAÇÕES", "grupo": "Resultado", "mask": "4.1.02.01.00030"},
    {"code": "4248", "name": "TELEFONE FIXO", "grupo": "Resultado", "mask": "4.1.02.01.00031"},
    {"code": "4249", "name": "TELEFONE MOVEL", "grupo": "Resultado", "mask": "4.1.02.01.00032"},
    {"code": "4250", "name": "INTERNET", "grupo": "Resultado", "mask": "4.1.02.01.00033"},
    {"code": "4251", "name": "SERVIÇOS PRESTADOS DE PROGRAMAÇÃO E TI", "grupo": "Resultado", "mask": "4.1.02.01.00034"},
    {"code": "4252", "name": "CUSTO COM INFRAESTRUTURA", "grupo": "Resultado", "mask": "4.1.02.01.00035"},
    {"code": "4253", "name": "CUSTO COM HONORÁRIOS", "grupo": "Resultado", "mask": "4.1.02.01.00036"},
    {"code": "4254", "name": "MANUTENÇÃO DE SOFTWARE", "grupo": "Resultado", "mask": "4.1.02.01.00037"},
    {"code": "4255", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "4.1.02.01.00038"},
    {"code": "4256", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA FISICA", "grupo": "Resultado", "mask": "4.1.02.01.00039"},
    {"code": "4257", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.1.02.01.00040"},
    {"code": "4258", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.02.01.00041"},
    {"code": "4259", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.02.01.00042"},
    {"code": "4260", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.1.02.01.00043"},
    {"code": "4261", "name": "ROYALTIES E ASSISTÊNCIA TÉCNICA PAÍS", "grupo": "Resultado", "mask": "4.1.02.01.00044"},
    {"code": "4262", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.02.01.00045"},
    {"code": "4263", "name": "INDUSTRIALIZAÇÃO EFETUADO POR TERCEIROS", "grupo": "Resultado", "mask": "4.1.02.01.00046"},
    {"code": "4264", "name": "ASSESSORIA CONTÁBIL", "grupo": "Resultado", "mask": "4.1.02.01.00047"},
    {"code": "4265", "name": "CONSULTORIA E SUPORTE TÉCNICO", "grupo": "Resultado", "mask": "4.1.02.01.00048"},
    {"code": "4266", "name": "SERVIÇOS DE ADVOGADOS", "grupo": "Resultado", "mask": "4.1.02.01.00049"},
    {"code": "4267", "name": "SERVIÇOS DE AUDITORIA", "grupo": "Resultado", "mask": "4.1.02.01.00050"},
    {"code": "4268", "name": "SERVIÇOS DE CONSULTORIA", "grupo": "Resultado", "mask": "4.1.02.01.00051"},
    {"code": "4269", "name": "SERVIÇOS DE DESPACHANTES ADUANEIRO", "grupo": "Resultado", "mask": "4.1.02.01.00052"},
    {"code": "4270", "name": "ASSESSORIA ADMINISTRATIVA", "grupo": "Resultado", "mask": "4.1.02.01.00053"},
    {"code": "4271", "name": "ADMINISTRAÇÃO DE BENS", "grupo": "Resultado", "mask": "4.1.02.01.00054"},
    {"code": "4272", "name": "SERVIÇOS E ANALISES TÉCNICAS", "grupo": "Resultado", "mask": "4.1.02.01.00055"},
    {"code": "4273", "name": "PESQUISA/PROJETO", "grupo": "Resultado", "mask": "4.1.02.01.00056"},
    {"code": "4274", "name": "CURSOS E TREINAMENTOS", "grupo": "Resultado", "mask": "4.1.02.01.00057"},
    {"code": "4275", "name": "DATILOGRAFIA", "grupo": "Resultado", "mask": "4.1.02.01.00058"},
    {"code": "4276", "name": "TRADUÇÕES E VERSÕES", "grupo": "Resultado", "mask": "4.1.02.01.00059"},
    {"code": "4277", "name": "SERVIÇOS DE INTERMEDIAÇÃO", "grupo": "Resultado", "mask": "4.1.02.01.00060"},
    {"code": "4278", "name": "ASSESSORIA E CONSULTORIA EM INFORMÁTICA", "grupo": "Resultado", "mask": "4.1.02.01.00061"},
    {"code": "4279", "name": "LICENÇA DE SOFTWARE", "grupo": "Resultado", "mask": "4.1.02.01.00062"},
    {"code": "4280", "name": "ANALISE E DESENVOLVIMENTO DE SISTEMAS", "grupo": "Resultado", "mask": "4.1.02.01.00063"},
    {"code": "4281", "name": "PROCESSAMENTO E ARMAZENAGEM DE DADOS", "grupo": "Resultado", "mask": "4.1.02.01.00064"},
    {"code": "4282", "name": "PROGRAMAÇÃO E COMUNICAÇÃO VISUAL", "grupo": "Resultado", "mask": "4.1.02.01.00065"},
    {"code": "4283", "name": "ASSISTÊNCIA TÉCNICA", "grupo": "Resultado", "mask": "4.1.02.01.00066"},
    {"code": "4285", "name": "VIAGENS E ESTADIAS", "grupo": "Resultado", "mask": "4.1.02.01.00068"},
    {"code": "4286", "name": "FEIRAS E EVENTOS", "grupo": "Resultado", "mask": "4.1.02.01.00069"},
    {"code": "4287", "name": "PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "4.1.02.01.00070"},
    {"code": "4288", "name": "SERVIÇOS GRÁFICOS", "grupo": "Resultado", "mask": "4.1.02.01.00071"},
    {"code": "4289", "name": "ANÚNCIOS E PUBLICAÇÕES", "grupo": "Resultado", "mask": "4.1.02.01.00072"},
    {"code": "4290", "name": "CARTÓRIO", "grupo": "Resultado", "mask": "4.1.02.01.00073"},
    {"code": "4291", "name": "DESPESAS COM ANIMAIS E MÉDICOS VETERINÁRIOS", "grupo": "Resultado", "mask": "4.1.02.01.00074"},
    {"code": "4292", "name": "SERVIÇOS MÉDICOS E EXAMES LABORATORIAS", "grupo": "Resultado", "mask": "4.1.02.01.00075"},
    {"code": "4293", "name": "ENGENHARIA E ARQUITETURA", "grupo": "Resultado", "mask": "4.1.02.01.00076"},
    {"code": "4294", "name": "MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.1.02.01.00077"},
    {"code": "4295", "name": "MANUTENÇÃO E CONSERVAÇÃO", "grupo": "Resultado", "mask": "4.1.02.01.00078"},
    {"code": "4296", "name": "INSTALAÇÕES E MONTAGENS", "grupo": "Resultado", "mask": "4.1.02.01.00079"},
    {"code": "4297", "name": "RECRUTAMENTO E SELEÇÃO", "grupo": "Resultado", "mask": "4.1.02.01.00080"},
    {"code": "4298", "name": "LOCAÇÃO DE MÃO DE OBRA", "grupo": "Resultado", "mask": "4.1.02.01.00081"},
    {"code": "4299", "name": "SERVIÇOS DE LIMPEZA", "grupo": "Resultado", "mask": "4.1.02.01.00082"},
    {"code": "4300", "name": "DEDETIZAÇÃO", "grupo": "Resultado", "mask": "4.1.02.01.00083"},
    {"code": "4301", "name": "SEGURANÇA", "grupo": "Resultado", "mask": "4.1.02.01.00084"},
    {"code": "4302", "name": "SEGUROS GERAIS", "grupo": "Resultado", "mask": "4.1.02.01.00085"},
    {"code": "4303", "name": "SEGUROS - RESPONSABILIDADE CÍVIL", "grupo": "Resultado", "mask": "4.1.02.01.00086"},
    {"code": "4304", "name": "SEGUROS - IMOBILIZADO", "grupo": "Resultado", "mask": "4.1.02.01.00087"},
    {"code": "4306", "name": "ESTACIONAMENTOS E PEDÁGIOS", "grupo": "Resultado", "mask": "4.1.02.01.00089"},
    {"code": "4307", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.02.01.00090"},
    {"code": "4308", "name": "ALUGUEL DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "4.1.02.01.00091"},
    {"code": "4309", "name": "ALUGUEL DE VEÍCULOS", "grupo": "Resultado", "mask": "4.1.02.01.00092"},
    {"code": "4310", "name": "ALUGUEL - PJ", "grupo": "Resultado", "mask": "4.1.02.01.00093"},
    {"code": "4311", "name": "ALUGUEL - PF", "grupo": "Resultado", "mask": "4.1.02.01.00094"},
    {"code": "4312", "name": "OUTROS CUSTOS", "grupo": "Resultado", "mask": "4.1.02.01.00095"},
    {"code": "4313", "name": "FOTOGRAFIA", "grupo": "Resultado", "mask": "4.1.02.01.00096"},
    {"code": "4314", "name": "OUTROS CUSTOS - PROVISÕES", "grupo": "Resultado", "mask": "4.1.02.01.00097"},
    {"code": "4315", "name": "MANUTENÇÃO DE VEÍCULOS", "grupo": "Resultado", "mask": "4.1.02.01.00098"},
    {"code": "4316", "name": "BRINDES E BONIFICAÇÕES", "grupo": "Resultado", "mask": "4.1.02.01.00099"},
    {"code": "4317", "name": "BENS DE USO PERMANENTE", "grupo": "Resultado", "mask": "4.1.02.01.00100"},
    {"code": "4318", "name": "CUSTOS COM DEPRECIAÇÃO", "grupo": "Resultado", "mask": "4.1.02.01.00101"},
    {"code": "4319", "name": "CUSTOS COM AMORTIZAÇÃO", "grupo": "Resultado", "mask": "4.1.02.01.00102"},
    {"code": "4337", "name": "SALÁRIOS E ORDENADOS", "grupo": "Resultado", "mask": "4.1.02.02.00001"},
    {"code": "4338", "name": "FÉRIAS", "grupo": "Resultado", "mask": "4.1.02.02.00002"},
    {"code": "4339", "name": "13º SALÁRIO", "grupo": "Resultado", "mask": "4.1.02.02.00003"},
    {"code": "4340", "name": "HORA EXTRA", "grupo": "Resultado", "mask": "4.1.02.02.00004"},
    {"code": "4341", "name": "DESCANSO SEMANAL REMUNERADO", "grupo": "Resultado", "mask": "4.1.02.02.00005"},
    {"code": "4342", "name": "BÔNUS", "grupo": "Resultado", "mask": "4.1.02.02.00006"},
    {"code": "4343", "name": "ABONO PECUNIÁRIO", "grupo": "Resultado", "mask": "4.1.02.02.00007"},
    {"code": "4344", "name": "VERBAS RESCISÓRIAS", "grupo": "Resultado", "mask": "4.1.02.02.00008"},
    {"code": "4345", "name": "ADICIONAL NOTURNO", "grupo": "Resultado", "mask": "4.1.02.02.00009"},
    {"code": "4346", "name": "ADICIONAL DE PERICULOSIDADE", "grupo": "Resultado", "mask": "4.1.02.02.00010"},
    {"code": "4347", "name": "ADICIONAL POR TEMPO DE SERVIÇO", "grupo": "Resultado", "mask": "4.1.02.02.00011"},
    {"code": "4348", "name": "ADICIONAL DE INSALUBRIDADE", "grupo": "Resultado", "mask": "4.1.02.02.00012"},
    {"code": "4349", "name": "AUXÍLIO HOME OFFICE", "grupo": "Resultado", "mask": "4.1.02.02.00013"},
    {"code": "4350", "name": "CONTRIBUIÇÃO ASSISTENCIAL", "grupo": "Resultado", "mask": "4.1.02.02.00014"},
    {"code": "4351", "name": "INDENIZAÇÕES", "grupo": "Resultado", "mask": "4.1.02.02.00015"},
    {"code": "4352", "name": "MULTA ESTABILIDADE", "grupo": "Resultado", "mask": "4.1.02.02.00016"},
    {"code": "4353", "name": "CONTRIBUIÇÃO SINDICAL", "grupo": "Resultado", "mask": "4.1.02.02.00017"},
    {"code": "4354", "name": "ANUÊNIO E TRIÊNIO", "grupo": "Resultado", "mask": "4.1.02.02.00018"},
    {"code": "4355", "name": "QUINQUENIO", "grupo": "Resultado", "mask": "4.1.02.02.00019"},
    {"code": "4356", "name": "OUTROS PROVENTOS RPA", "grupo": "Resultado", "mask": "4.1.02.02.00020"},
    {"code": "4357", "name": "BOLSA GRATIFICAÇÃO", "grupo": "Resultado", "mask": "4.1.02.02.00021"},
    {"code": "4358", "name": "PROVISÕES E ENCARGOS", "grupo": "Resultado", "mask": "4.1.02.02.00022"},
    {"code": "4359", "name": "AUXÍLIO CRECHE", "grupo": "Resultado", "mask": "4.1.02.02.00023"},
    {"code": "4360", "name": "BOLSA AUXÍLIO", "grupo": "Resultado", "mask": "4.1.02.02.00024"},
    {"code": "4361", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.1.02.02.00025"},
    {"code": "4362", "name": "SERVIÇOS AUTÔNOMO", "grupo": "Resultado", "mask": "4.1.02.02.00026"},
    {"code": "4363", "name": "PARTICIPAÇÃO NOS LUCROS E RESULTADOS", "grupo": "Resultado", "mask": "4.1.02.02.00027"},
    {"code": "4364", "name": "PRÊMIOS E GRATIFICAÇÕES", "grupo": "Resultado", "mask": "4.1.02.02.00028"},
    {"code": "4365", "name": "DONATIVOS E GORJETAS", "grupo": "Resultado", "mask": "4.1.02.02.00029"},
    {"code": "4366", "name": "COMISSÕES", "grupo": "Resultado", "mask": "4.1.02.02.00030"},
    {"code": "4367", "name": "PRÓ-LABORE", "grupo": "Resultado", "mask": "4.1.02.02.00031"},
    {"code": "4368", "name": "FGTS", "grupo": "Resultado", "mask": "4.1.02.02.00032"},
    {"code": "4369", "name": "INSS", "grupo": "Resultado", "mask": "4.1.02.02.00033"},
    {"code": "4370", "name": "INSS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.1.02.02.00034"},
    {"code": "4371", "name": "INSS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.1.02.02.00035"},
    {"code": "4372", "name": "FGTS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.1.02.02.00036"},
    {"code": "4373", "name": "FGTS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.1.02.02.00037"},
    {"code": "4374", "name": "MULTA FGTS", "grupo": "Resultado", "mask": "4.1.02.02.00038"},
    {"code": "4375", "name": "AJUDA DE CUSTO", "grupo": "Resultado", "mask": "4.1.02.02.00039"},
    {"code": "4376", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.1.02.02.00040"},
    {"code": "4377", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.1.02.02.00041"},
    {"code": "4378", "name": "VALE ALIMENTAÇÃO/REFEIÇÃO", "grupo": "Resultado", "mask": "4.1.02.02.00042"},
    {"code": "4379", "name": "VALE TRANSPORTE", "grupo": "Resultado", "mask": "4.1.02.02.00043"},
    {"code": "4380", "name": "VALE COMBUSTÍVEL", "grupo": "Resultado", "mask": "4.1.02.02.00044"},
    {"code": "4381", "name": "BENEFÍCIOS TRABALHISTAS FLEXÍVEIS", "grupo": "Resultado", "mask": "4.1.02.02.00045"},
    {"code": "4382", "name": "SEGURO DE VIDA", "grupo": "Resultado", "mask": "4.1.02.02.00046"},
    {"code": "4383", "name": "ASSISTÊNCIA MÉDICA", "grupo": "Resultado", "mask": "4.1.02.02.00047"},
    {"code": "4384", "name": "ASSISTÊNCIA ODONTOLÓGICA", "grupo": "Resultado", "mask": "4.1.02.02.00048"},
    {"code": "4385", "name": "PIS SOBRE FOLHA", "grupo": "Resultado", "mask": "4.1.02.02.00049"},
    {"code": "4386", "name": "FUNDO DE APOSENTADORIA", "grupo": "Resultado", "mask": "4.1.02.02.00050"},
    {"code": "4387", "name": "PLANO DE POUPANÇA E INVESTIMENTO", "grupo": "Resultado", "mask": "4.1.02.02.00051"},
    {"code": "4388", "name": "VALE CULTURA", "grupo": "Resultado", "mask": "4.1.02.02.00052"},
    {"code": "4389", "name": "REFEIÇÕES PAT", "grupo": "Resultado", "mask": "4.1.02.02.00053"},
    {"code": "4390", "name": "PROCESSOS TRABALHISTAS", "grupo": "Resultado", "mask": "4.1.02.02.00054"},
    {"code": "4391", "name": "PREVIDÊNCIA PRIVADA", "grupo": "Resultado", "mask": "4.1.02.02.00055"},
    {"code": "4392", "name": "OUTRAS DESPESAS COM PESSOAL", "grupo": "Resultado", "mask": "4.1.02.02.00056"},
    {"code": "4402", "name": "SERVIÇOS PRESTADOS DE SUPORTE DE OPERAÇÕES", "grupo": "Resultado", "mask": "4.1.02.03.00001"},
    {"code": "4403", "name": "CUSTOS COM IMPORTAÇÃO", "grupo": "Resultado", "mask": "4.1.02.03.00002"},
    {"code": "4404", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.02.03.00003"},
    {"code": "4405", "name": "OUTROS CUSTOS COM SERVIÇOS", "grupo": "Resultado", "mask": "4.1.02.03.00004"},
    {"code": "4406", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.1.02.03.00005"},
    {"code": "4407", "name": "CRÉDITO DE PIS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "4.1.02.03.00006"},
    {"code": "4408", "name": "CRÉDITO DE COFINS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "4.1.02.03.00007"},
    {"code": "4409", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.02.03.00008"},
    {"code": "4410", "name": "CUSTO DA LOCAÇÃO DE BENS MÓVEIS E IMÓVEIS", "grupo": "Resultado", "mask": "4.1.02.03.00009"},
    {"code": "4411", "name": "CUSTO DA SUBCONTRATAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.02.03.00010"},
    {"code": "4412", "name": "CUSTO DA VENDA DE UNIDADES IMOBILIÁRIAS", "grupo": "Resultado", "mask": "4.1.02.03.00011"},
    {"code": "4413", "name": "CUSTO DO CONTRATO DE CONSTRUÇÃO", "grupo": "Resultado", "mask": "4.1.02.03.00012"},
    {"code": "4414", "name": "CUSTO DA SECURITIZAÇÃO DE CRÉDITOS", "grupo": "Resultado", "mask": "4.1.02.03.00013"},
    {"code": "4415", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.02.03.00014"},
    {"code": "4416", "name": "OUTROS CUSTOS DA ATIVIDADE GERAL", "grupo": "Resultado", "mask": "4.1.02.03.00015"},
    {"code": "4417", "name": "CUSTOS DIVERSOS", "grupo": "Resultado", "mask": "4.1.02.03.00016"},
    {"code": "4418", "name": "CUSTOS INDEDUTÍVEIS", "grupo": "Resultado", "mask": "4.1.02.03.00017"},
    {"code": "4419", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.02.03.00018"},
    {"code": "4420", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.02.03.00019"},
    {"code": "4421", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.1.02.03.00020"},
    {"code": "4431", "name": "CMV", "grupo": "Resultado", "mask": "4.1.03.01.00001"},
    {"code": "4432", "name": "BRINDES BONIFICAÇÕES", "grupo": "Resultado", "mask": "4.1.03.01.00002"},
    {"code": "4433", "name": "MATERIAL DE USO E CONSUMO", "grupo": "Resultado", "mask": "4.1.03.01.00003"},
    {"code": "4434", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "4.1.03.01.00004"},
    {"code": "4435", "name": "FRETES E CARRETOS", "grupo": "Resultado", "mask": "4.1.03.01.00005"},
    {"code": "4436", "name": "ÁGUA", "grupo": "Resultado", "mask": "4.1.03.01.00006"},
    {"code": "4437", "name": "ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "4.1.03.01.00007"},
    {"code": "4438", "name": "PERDA COM ESTOQUES", "grupo": "Resultado", "mask": "4.1.03.01.00008"},
    {"code": "4439", "name": "MATERIAL DE EMBALAGEM", "grupo": "Resultado", "mask": "4.1.03.01.00009"},
    {"code": "4440", "name": "MATERIAL DE MANUTENÇÃO E REPARO", "grupo": "Resultado", "mask": "4.1.03.01.00010"},
    {"code": "4441", "name": "MATERIAL DE MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.1.03.01.00011"},
    {"code": "4442", "name": "ARMAZENAGEM E DEPOSITO", "grupo": "Resultado", "mask": "4.1.03.01.00012"},
    {"code": "4443", "name": "CUSTO COM INSTALAÇÕES", "grupo": "Resultado", "mask": "4.1.03.01.00013"},
    {"code": "4444", "name": "LANCHES E REFEIÇÕES", "grupo": "Resultado", "mask": "4.1.03.01.00014"},
    {"code": "4445", "name": "LANCHES E REFEIÇÕES - DIRIGENTES", "grupo": "Resultado", "mask": "4.1.03.01.00015"},
    {"code": "4446", "name": "COPA E COZINHA", "grupo": "Resultado", "mask": "4.1.03.01.00016"},
    {"code": "4447", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.1.03.01.00017"},
    {"code": "4448", "name": "FARMÁCIA", "grupo": "Resultado", "mask": "4.1.03.01.00018"},
    {"code": "4449", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.1.03.01.00019"},
    {"code": "4450", "name": "MATERIAIS DE HIGIENE E LIMPEZA", "grupo": "Resultado", "mask": "4.1.03.01.00020"},
    {"code": "4451", "name": "MATERIAIS DE ESCRITÓRIO", "grupo": "Resultado", "mask": "4.1.03.01.00021"},
    {"code": "4452", "name": "MATERIAIS DE INFORMÁTICA", "grupo": "Resultado", "mask": "4.1.03.01.00022"},
    {"code": "4454", "name": "MATERIAIS PUBLICITÁRIOS", "grupo": "Resultado", "mask": "4.1.03.01.00024"},
    {"code": "4455", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.1.03.01.00025"},
    {"code": "4456", "name": "BENS DE PEQUENO VALOR", "grupo": "Resultado", "mask": "4.1.03.01.00026"},
    {"code": "4457", "name": "DESPESA COM VEÍCULOS", "grupo": "Resultado", "mask": "4.1.03.01.00027"},
    {"code": "4458", "name": "COMPUTADORES E PERIFÉRICOS", "grupo": "Resultado", "mask": "4.1.03.01.00028"},
    {"code": "4460", "name": "BENFEITORIAS EM IMÓVEIS DE TERCEIROS", "grupo": "Resultado", "mask": "4.1.03.01.00030"},
    {"code": "4461", "name": "COMUNICAÇÕES", "grupo": "Resultado", "mask": "4.1.03.01.00031"},
    {"code": "4462", "name": "TELEFONE FIXO", "grupo": "Resultado", "mask": "4.1.03.01.00032"},
    {"code": "4463", "name": "TELEFONE MOVEL", "grupo": "Resultado", "mask": "4.1.03.01.00033"},
    {"code": "4464", "name": "INTERNET", "grupo": "Resultado", "mask": "4.1.03.01.00034"},
    {"code": "4465", "name": "SERVIÇOS PRESTADOS DE PROGRAMAÇÃO E TI", "grupo": "Resultado", "mask": "4.1.03.01.00035"},
    {"code": "4466", "name": "CUSTO COM INFRAESTRUTURA", "grupo": "Resultado", "mask": "4.1.03.01.00036"},
    {"code": "4467", "name": "CUSTO COM HONORÁRIOS", "grupo": "Resultado", "mask": "4.1.03.01.00037"},
    {"code": "4468", "name": "MANUTENÇÃO DE SOFTWARE", "grupo": "Resultado", "mask": "4.1.03.01.00038"},
    {"code": "4469", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "4.1.03.01.00039"},
    {"code": "4470", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA FISICA", "grupo": "Resultado", "mask": "4.1.03.01.00040"},
    {"code": "4471", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.1.03.01.00041"},
    {"code": "4472", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.03.01.00042"},
    {"code": "4473", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.03.01.00043"},
    {"code": "4474", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.1.03.01.00044"},
    {"code": "4475", "name": "ROYALTIES E ASSISTÊNCIA TÉCNICA PAÍS", "grupo": "Resultado", "mask": "4.1.03.01.00045"},
    {"code": "4476", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.03.01.00046"},
    {"code": "4477", "name": "INDUSTRIALIZAÇÃO EFETUADO POR TERCEIROS", "grupo": "Resultado", "mask": "4.1.03.01.00047"},
    {"code": "4478", "name": "ASSESSORIA CONTÁBIL", "grupo": "Resultado", "mask": "4.1.03.01.00048"},
    {"code": "4479", "name": "CONSULTORIA E SUPORTE TÉCNICO", "grupo": "Resultado", "mask": "4.1.03.01.00049"},
    {"code": "4480", "name": "SERVIÇOS DE ADVOGADOS", "grupo": "Resultado", "mask": "4.1.03.01.00050"},
    {"code": "4481", "name": "SERVIÇOS DE AUDITORIA", "grupo": "Resultado", "mask": "4.1.03.01.00051"},
    {"code": "4482", "name": "SERVIÇOS DE CONSULTORIA", "grupo": "Resultado", "mask": "4.1.03.01.00052"},
    {"code": "4483", "name": "SERVIÇOS DE DESPACHANTES ADUANEIRO", "grupo": "Resultado", "mask": "4.1.03.01.00053"},
    {"code": "4484", "name": "ASSESSORIA ADMINISTRATIVA", "grupo": "Resultado", "mask": "4.1.03.01.00054"},
    {"code": "4485", "name": "ADMINISTRAÇÃO DE BENS", "grupo": "Resultado", "mask": "4.1.03.01.00055"},
    {"code": "4486", "name": "SERVIÇOS E ANALISES TÉCNICAS", "grupo": "Resultado", "mask": "4.1.03.01.00056"},
    {"code": "4487", "name": "PESQUISA/PROJETO", "grupo": "Resultado", "mask": "4.1.03.01.00057"},
    {"code": "4488", "name": "CURSOS E TREINAMENTOS", "grupo": "Resultado", "mask": "4.1.03.01.00058"},
    {"code": "4489", "name": "DATILOGRAFIA", "grupo": "Resultado", "mask": "4.1.03.01.00059"},
    {"code": "4490", "name": "TRADUÇÕES E VERSÕES", "grupo": "Resultado", "mask": "4.1.03.01.00060"},
    {"code": "4491", "name": "SERVIÇOS DE INTERMEDIAÇÃO", "grupo": "Resultado", "mask": "4.1.03.01.00061"},
    {"code": "4492", "name": "ASSESSORIA E CONSULTORIA EM INFORMÁTICA", "grupo": "Resultado", "mask": "4.1.03.01.00062"},
    {"code": "4493", "name": "LICENÇA DE SOFTWARE", "grupo": "Resultado", "mask": "4.1.03.01.00063"},
    {"code": "4494", "name": "ANALISE E DESENVOLVIMENTO DE SISTEMAS", "grupo": "Resultado", "mask": "4.1.03.01.00064"},
    {"code": "4495", "name": "PROCESSAMENTO E ARMAZENAGEM DE DADOS", "grupo": "Resultado", "mask": "4.1.03.01.00065"},
    {"code": "4496", "name": "PROGRAMAÇÃO E COMUNICAÇÃO VISUAL", "grupo": "Resultado", "mask": "4.1.03.01.00066"},
    {"code": "4497", "name": "ASSISTÊNCIA TÉCNICA", "grupo": "Resultado", "mask": "4.1.03.01.00067"},
    {"code": "4499", "name": "VIAGENS E ESTADIAS", "grupo": "Resultado", "mask": "4.1.03.01.00069"},
    {"code": "4500", "name": "FEIRAS E EVENTOS", "grupo": "Resultado", "mask": "4.1.03.01.00070"},
    {"code": "4501", "name": "PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "4.1.03.01.00071"},
    {"code": "4502", "name": "SERVIÇOS GRÁFICOS", "grupo": "Resultado", "mask": "4.1.03.01.00072"},
    {"code": "4503", "name": "ANÚNCIOS E PUBLICAÇÕES", "grupo": "Resultado", "mask": "4.1.03.01.00073"},
    {"code": "4504", "name": "CARTÓRIO", "grupo": "Resultado", "mask": "4.1.03.01.00074"},
    {"code": "4505", "name": "DESPESAS COM ANIMAIS E MÉDICOS VETERINÁRIOS", "grupo": "Resultado", "mask": "4.1.03.01.00075"},
    {"code": "4506", "name": "SERVIÇOS MÉDICOS E EXAMES LABORATORIAS", "grupo": "Resultado", "mask": "4.1.03.01.00076"},
    {"code": "4507", "name": "ENGENHARIA E ARQUITETURA", "grupo": "Resultado", "mask": "4.1.03.01.00077"},
    {"code": "4508", "name": "MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.1.03.01.00078"},
    {"code": "4509", "name": "MANUTENÇÃO E CONSERVAÇÃO", "grupo": "Resultado", "mask": "4.1.03.01.00079"},
    {"code": "4510", "name": "INSTALAÇÕES E MONTAGENS", "grupo": "Resultado", "mask": "4.1.03.01.00080"},
    {"code": "4511", "name": "RECRUTAMENTO E SELEÇÃO", "grupo": "Resultado", "mask": "4.1.03.01.00081"},
    {"code": "4512", "name": "LOCAÇÃO DE MÃO DE OBRA", "grupo": "Resultado", "mask": "4.1.03.01.00082"},
    {"code": "4513", "name": "SERVIÇOS DE LIMPEZA", "grupo": "Resultado", "mask": "4.1.03.01.00083"},
    {"code": "4514", "name": "DEDETIZAÇÃO", "grupo": "Resultado", "mask": "4.1.03.01.00084"},
    {"code": "4515", "name": "SEGURANÇA", "grupo": "Resultado", "mask": "4.1.03.01.00085"},
    {"code": "4516", "name": "SEGUROS GERAIS", "grupo": "Resultado", "mask": "4.1.03.01.00086"},
    {"code": "4517", "name": "SEGUROS - RESPONSABILIDADE CÍVIL", "grupo": "Resultado", "mask": "4.1.03.01.00087"},
    {"code": "4518", "name": "SEGUROS - IMOBILIZADO", "grupo": "Resultado", "mask": "4.1.03.01.00088"},
    {"code": "4520", "name": "ESTACIONAMENTOS E PEDÁGIOS", "grupo": "Resultado", "mask": "4.1.03.01.00090"},
    {"code": "4521", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.03.01.00091"},
    {"code": "4522", "name": "ALUGUEL DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "4.1.03.01.00092"},
    {"code": "4523", "name": "ALUGUEL DE VEÍCULOS", "grupo": "Resultado", "mask": "4.1.03.01.00093"},
    {"code": "4524", "name": "ALUGUEL - PJ", "grupo": "Resultado", "mask": "4.1.03.01.00094"},
    {"code": "4525", "name": "ALUGUEL - PF", "grupo": "Resultado", "mask": "4.1.03.01.00095"},
    {"code": "4526", "name": "MANUTENÇÃO DE VEÍCULOS", "grupo": "Resultado", "mask": "4.1.03.01.00096"},
    {"code": "4527", "name": "OUTROS CUSTOS - PROVISÕES", "grupo": "Resultado", "mask": "4.1.03.01.00097"},
    {"code": "4528", "name": "FOTOGRAFIA", "grupo": "Resultado", "mask": "4.1.03.01.00098"},
    {"code": "4529", "name": "BENS DE USO PERMANENTE", "grupo": "Resultado", "mask": "4.1.03.01.00099"},
    {"code": "4530", "name": "CUSTOS COM DEPRECIAÇÃO", "grupo": "Resultado", "mask": "4.1.03.01.00100"},
    {"code": "4531", "name": "CUSTOS COM AMORTIZAÇÃO", "grupo": "Resultado", "mask": "4.1.03.01.00101"},
    {"code": "4549", "name": "SALÁRIOS E ORDENADOS", "grupo": "Resultado", "mask": "4.1.03.02.00001"},
    {"code": "4550", "name": "FÉRIAS", "grupo": "Resultado", "mask": "4.1.03.02.00002"},
    {"code": "4551", "name": "13º SALÁRIO", "grupo": "Resultado", "mask": "4.1.03.02.00003"},
    {"code": "4552", "name": "HORA EXTRA", "grupo": "Resultado", "mask": "4.1.03.02.00004"},
    {"code": "4553", "name": "DESCANSO SEMANAL REMUNERADO", "grupo": "Resultado", "mask": "4.1.03.02.00005"},
    {"code": "4554", "name": "BÔNUS", "grupo": "Resultado", "mask": "4.1.03.02.00006"},
    {"code": "4555", "name": "ABONO PECUNIÁRIO", "grupo": "Resultado", "mask": "4.1.03.02.00007"},
    {"code": "4556", "name": "VERBAS RESCISÓRIAS", "grupo": "Resultado", "mask": "4.1.03.02.00008"},
    {"code": "4557", "name": "ADICIONAL NOTURNO", "grupo": "Resultado", "mask": "4.1.03.02.00009"},
    {"code": "4558", "name": "ADICIONAL DE PERICULOSIDADE", "grupo": "Resultado", "mask": "4.1.03.02.00010"},
    {"code": "4559", "name": "ADICIONAL POR TEMPO DE SERVIÇO", "grupo": "Resultado", "mask": "4.1.03.02.00011"},
    {"code": "4560", "name": "ADICIONAL DE INSALUBRIDADE", "grupo": "Resultado", "mask": "4.1.03.02.00012"},
    {"code": "4561", "name": "AUXÍLIO HOME OFFICE", "grupo": "Resultado", "mask": "4.1.03.02.00013"},
    {"code": "4562", "name": "CONTRIBUIÇÃO ASSISTENCIAL", "grupo": "Resultado", "mask": "4.1.03.02.00014"},
    {"code": "4563", "name": "INDENIZAÇÕES", "grupo": "Resultado", "mask": "4.1.03.02.00015"},
    {"code": "4564", "name": "MULTA ESTABILIDADE", "grupo": "Resultado", "mask": "4.1.03.02.00016"},
    {"code": "4565", "name": "CONTRIBUIÇÃO SINDICAL", "grupo": "Resultado", "mask": "4.1.03.02.00017"},
    {"code": "4566", "name": "ANUÊNIO E TRIÊNIO", "grupo": "Resultado", "mask": "4.1.03.02.00018"},
    {"code": "4567", "name": "QUINQUENIO", "grupo": "Resultado", "mask": "4.1.03.02.00019"},
    {"code": "4568", "name": "OUTROS PROVENTOS RPA", "grupo": "Resultado", "mask": "4.1.03.02.00020"},
    {"code": "4569", "name": "BOLSA GRATIFICAÇÃO", "grupo": "Resultado", "mask": "4.1.03.02.00021"},
    {"code": "4570", "name": "PROVISÕES E ENCARGOS", "grupo": "Resultado", "mask": "4.1.03.02.00022"},
    {"code": "4571", "name": "AUXÍLIO CRECHE", "grupo": "Resultado", "mask": "4.1.03.02.00023"},
    {"code": "4572", "name": "BOLSA AUXÍLIO", "grupo": "Resultado", "mask": "4.1.03.02.00024"},
    {"code": "4573", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.1.03.02.00025"},
    {"code": "4574", "name": "SERVIÇOS AUTÔNOMO", "grupo": "Resultado", "mask": "4.1.03.02.00026"},
    {"code": "4575", "name": "PARTICIPAÇÃO NOS LUCROS E RESULTADOS", "grupo": "Resultado", "mask": "4.1.03.02.00027"},
    {"code": "4576", "name": "PRÊMIOS E GRATIFICAÇÕES", "grupo": "Resultado", "mask": "4.1.03.02.00028"},
    {"code": "4577", "name": "DONATIVOS E GORJETAS", "grupo": "Resultado", "mask": "4.1.03.02.00029"},
    {"code": "4578", "name": "COMISSÕES", "grupo": "Resultado", "mask": "4.1.03.02.00030"},
    {"code": "4579", "name": "PRÓ-LABORE", "grupo": "Resultado", "mask": "4.1.03.02.00031"},
    {"code": "4580", "name": "FGTS", "grupo": "Resultado", "mask": "4.1.03.02.00032"},
    {"code": "4581", "name": "INSS", "grupo": "Resultado", "mask": "4.1.03.02.00033"},
    {"code": "4582", "name": "INSS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.1.03.02.00034"},
    {"code": "4583", "name": "INSS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.1.03.02.00035"},
    {"code": "4584", "name": "FGTS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.1.03.02.00036"},
    {"code": "4585", "name": "FGTS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.1.03.02.00037"},
    {"code": "4586", "name": "MULTA FGTS", "grupo": "Resultado", "mask": "4.1.03.02.00038"},
    {"code": "4587", "name": "AJUDA DE CUSTO", "grupo": "Resultado", "mask": "4.1.03.02.00039"},
    {"code": "4588", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.1.03.02.00040"},
    {"code": "4589", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.1.03.02.00041"},
    {"code": "4590", "name": "VALE ALIMENTAÇÃO/REFEIÇÃO", "grupo": "Resultado", "mask": "4.1.03.02.00042"},
    {"code": "4591", "name": "VALE TRANSPORTE", "grupo": "Resultado", "mask": "4.1.03.02.00043"},
    {"code": "4592", "name": "VALE COMBUSTÍVEL", "grupo": "Resultado", "mask": "4.1.03.02.00044"},
    {"code": "4593", "name": "BENEFÍCIOS TRABALHISTAS FLEXÍVEIS", "grupo": "Resultado", "mask": "4.1.03.02.00045"},
    {"code": "4594", "name": "SEGURO DE VIDA", "grupo": "Resultado", "mask": "4.1.03.02.00046"},
    {"code": "4595", "name": "ASSISTÊNCIA MÉDICA", "grupo": "Resultado", "mask": "4.1.03.02.00047"},
    {"code": "4596", "name": "ASSISTÊNCIA ODONTOLÓGICA", "grupo": "Resultado", "mask": "4.1.03.02.00048"},
    {"code": "4597", "name": "PIS SOBRE FOLHA", "grupo": "Resultado", "mask": "4.1.03.02.00049"},
    {"code": "4598", "name": "FUNDO DE APOSENTADORIA", "grupo": "Resultado", "mask": "4.1.03.02.00050"},
    {"code": "4599", "name": "PLANO DE POUPANÇA E INVESTIMENTO", "grupo": "Resultado", "mask": "4.1.03.02.00051"},
    {"code": "4600", "name": "VALE CULTURA", "grupo": "Resultado", "mask": "4.1.03.02.00052"},
    {"code": "4601", "name": "REFEIÇÕES PAT", "grupo": "Resultado", "mask": "4.1.03.02.00053"},
    {"code": "4602", "name": "PROCESSOS TRABALHISTAS", "grupo": "Resultado", "mask": "4.1.03.02.00054"},
    {"code": "4603", "name": "PREVIDÊNCIA PRIVADA", "grupo": "Resultado", "mask": "4.1.03.02.00055"},
    {"code": "4604", "name": "OUTRAS DESPESAS COM PESSOAL", "grupo": "Resultado", "mask": "4.1.03.02.00056"},
    {"code": "4624", "name": "SERVIÇOS PRESTADOS DE SUPORTE DE OPERAÇÕES", "grupo": "Resultado", "mask": "4.1.04.01.00001"},
    {"code": "4625", "name": "CUSTOS COM IMPORTAÇÃO", "grupo": "Resultado", "mask": "4.1.04.01.00002"},
    {"code": "4626", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.04.01.00003"},
    {"code": "4627", "name": "OUTROS CUSTOS COM SERVIÇOS", "grupo": "Resultado", "mask": "4.1.04.01.00004"},
    {"code": "4628", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.1.04.01.00005"},
    {"code": "4629", "name": "CRÉDITO DE PIS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "4.1.04.01.00006"},
    {"code": "4630", "name": "CRÉDITO DE COFINS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "4.1.04.01.00007"},
    {"code": "4631", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.04.01.00008"},
    {"code": "4632", "name": "CUSTO DA LOCAÇÃO DE BENS MÓVEIS E IMÓVEIS", "grupo": "Resultado", "mask": "4.1.04.01.00009"},
    {"code": "4633", "name": "CUSTO DA SUBCONTRATAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.1.04.01.00010"},
    {"code": "4634", "name": "CUSTO DA VENDA DE UNIDADES IMOBILIÁRIAS", "grupo": "Resultado", "mask": "4.1.04.01.00011"},
    {"code": "4635", "name": "CUSTO DO CONTRATO DE CONSTRUÇÃO", "grupo": "Resultado", "mask": "4.1.04.01.00012"},
    {"code": "4636", "name": "CUSTO DA SECURITIZAÇÃO DE CRÉDITOS", "grupo": "Resultado", "mask": "4.1.04.01.00013"},
    {"code": "4637", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.1.04.01.00014"},
    {"code": "4638", "name": "OUTROS CUSTOS DA ATIVIDADE GERAL", "grupo": "Resultado", "mask": "4.1.04.01.00015"},
    {"code": "4639", "name": "CUSTOS DIVERSOS", "grupo": "Resultado", "mask": "4.1.04.01.00016"},
    {"code": "4640", "name": "CUSTOS INDEDUTÍVEIS", "grupo": "Resultado", "mask": "4.1.04.01.00017"},
    {"code": "4641", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.04.01.00018"},
    {"code": "4642", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.1.04.01.00019"},
    {"code": "4643", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.1.04.01.00020"},
    {"code": "4663", "name": "SALÁRIOS E ORDENADOS", "grupo": "Resultado", "mask": "4.2.01.01.00001"},
    {"code": "4664", "name": "FÉRIAS", "grupo": "Resultado", "mask": "4.2.01.01.00002"},
    {"code": "4665", "name": "13º SALÁRIO", "grupo": "Resultado", "mask": "4.2.01.01.00003"},
    {"code": "4666", "name": "HORA EXTRA", "grupo": "Resultado", "mask": "4.2.01.01.00004"},
    {"code": "4667", "name": "ABONO PECUNIÁRIO", "grupo": "Resultado", "mask": "4.2.01.01.00005"},
    {"code": "4668", "name": "DESCANSO SEMANAL REMUNERADO", "grupo": "Resultado", "mask": "4.2.01.01.00006"},
    {"code": "4669", "name": "BÔNUS", "grupo": "Resultado", "mask": "4.2.01.01.00007"},
    {"code": "4670", "name": "VERBAS RESCISÓRIAS", "grupo": "Resultado", "mask": "4.2.01.01.00008"},
    {"code": "4671", "name": "ADICIONAL NOTURNO", "grupo": "Resultado", "mask": "4.2.01.01.00009"},
    {"code": "4672", "name": "ADICIONAL DE PERICULOSIDADE", "grupo": "Resultado", "mask": "4.2.01.01.00010"},
    {"code": "4673", "name": "ADICIONAL POR TEMPO DE SERVIÇO", "grupo": "Resultado", "mask": "4.2.01.01.00011"},
    {"code": "4674", "name": "ADICIONAL DE INSALUBRIDADE", "grupo": "Resultado", "mask": "4.2.01.01.00012"},
    {"code": "4675", "name": "AUXÍLIO HOME OFFICE", "grupo": "Resultado", "mask": "4.2.01.01.00013"},
    {"code": "4676", "name": "CONTRIBUIÇÃO ASSISTENCIAL", "grupo": "Resultado", "mask": "4.2.01.01.00014"},
    {"code": "4677", "name": "INDENIZAÇÕES", "grupo": "Resultado", "mask": "4.2.01.01.00015"},
    {"code": "4678", "name": "MULTA ESTABILIDADE", "grupo": "Resultado", "mask": "4.2.01.01.00016"},
    {"code": "4679", "name": "CONTRIBUIÇÃO SINDICAL", "grupo": "Resultado", "mask": "4.2.01.01.00017"},
    {"code": "4680", "name": "ANUÊNIO E TRIÊNIO", "grupo": "Resultado", "mask": "4.2.01.01.00018"},
    {"code": "4681", "name": "QUINQUENIO", "grupo": "Resultado", "mask": "4.2.01.01.00019"},
    {"code": "4682", "name": "OUTROS PROVENTOS RPA", "grupo": "Resultado", "mask": "4.2.01.01.00020"},
    {"code": "4683", "name": "BOLSA GRATIFICAÇÃO", "grupo": "Resultado", "mask": "4.2.01.01.00021"},
    {"code": "4684", "name": "PROVISÕES E ENCARGOS", "grupo": "Resultado", "mask": "4.2.01.01.00022"},
    {"code": "4685", "name": "AUXÍLIO CRECHE", "grupo": "Resultado", "mask": "4.2.01.01.00023"},
    {"code": "4686", "name": "BOLSA AUXÍLIO", "grupo": "Resultado", "mask": "4.2.01.01.00024"},
    {"code": "4687", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.2.01.01.00025"},
    {"code": "4688", "name": "SERVIÇOS AUTÔNOMO", "grupo": "Resultado", "mask": "4.2.01.01.00026"},
    {"code": "4689", "name": "PARTICIPAÇÃO NOS LUCROS E RESULTADOS", "grupo": "Resultado", "mask": "4.2.01.01.00027"},
    {"code": "4690", "name": "PRÊMIOS E GRATIFICAÇÕES", "grupo": "Resultado", "mask": "4.2.01.01.00028"},
    {"code": "4691", "name": "DONATIVOS E GORJETAS", "grupo": "Resultado", "mask": "4.2.01.01.00029"},
    {"code": "4692", "name": "COMISSÕES", "grupo": "Resultado", "mask": "4.2.01.01.00030"},
    {"code": "4693", "name": "PRÓ-LABORE", "grupo": "Resultado", "mask": "4.2.01.01.00031"},
    {"code": "4694", "name": "FGTS", "grupo": "Resultado", "mask": "4.2.01.01.00032"},
    {"code": "4695", "name": "INSS", "grupo": "Resultado", "mask": "4.2.01.01.00033"},
    {"code": "4696", "name": "INSS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.2.01.01.00034"},
    {"code": "4697", "name": "INSS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.2.01.01.00035"},
    {"code": "4698", "name": "FGTS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "4.2.01.01.00036"},
    {"code": "4699", "name": "FGTS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "4.2.01.01.00037"},
    {"code": "4700", "name": "MULTA FGTS", "grupo": "Resultado", "mask": "4.2.01.01.00038"},
    {"code": "4701", "name": "AJUDA DE CUSTO", "grupo": "Resultado", "mask": "4.2.01.01.00039"},
    {"code": "4702", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.2.01.01.00040"},
    {"code": "4703", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.2.01.01.00041"},
    {"code": "4704", "name": "VALE ALIMENTAÇÃO/REFEIÇÃO", "grupo": "Resultado", "mask": "4.2.01.01.00042"},
    {"code": "4705", "name": "VALE TRANSPORTE", "grupo": "Resultado", "mask": "4.2.01.01.00043"},
    {"code": "4706", "name": "VALE COMBUSTÍVEL", "grupo": "Resultado", "mask": "4.2.01.01.00044"},
    {"code": "4707", "name": "BENEFÍCIOS TRABALHISTAS FLEXÍVEIS", "grupo": "Resultado", "mask": "4.2.01.01.00045"},
    {"code": "4708", "name": "SEGURO DE VIDA", "grupo": "Resultado", "mask": "4.2.01.01.00046"},
    {"code": "4709", "name": "ASSISTÊNCIA MÉDICA", "grupo": "Resultado", "mask": "4.2.01.01.00047"},
    {"code": "4710", "name": "ASSISTÊNCIA ODONTOLÓGICA", "grupo": "Resultado", "mask": "4.2.01.01.00048"},
    {"code": "4711", "name": "PIS SOBRE FOLHA", "grupo": "Resultado", "mask": "4.2.01.01.00049"},
    {"code": "4712", "name": "FUNDO DE APOSENTADORIA", "grupo": "Resultado", "mask": "4.2.01.01.00050"},
    {"code": "4713", "name": "PLANO DE POUPANÇA E INVESTIMENTO", "grupo": "Resultado", "mask": "4.2.01.01.00051"},
    {"code": "4714", "name": "VALE CULTURA", "grupo": "Resultado", "mask": "4.2.01.01.00052"},
    {"code": "4715", "name": "REFEIÇÕES PAT", "grupo": "Resultado", "mask": "4.2.01.01.00053"},
    {"code": "4716", "name": "PROCESSOS TRABALHISTAS", "grupo": "Resultado", "mask": "4.2.01.01.00054"},
    {"code": "4717", "name": "PREVIDÊNCIA PRIVADA", "grupo": "Resultado", "mask": "4.2.01.01.00055"},
    {"code": "4718", "name": "OUTRAS DESPESAS COM PESSOAL", "grupo": "Resultado", "mask": "4.2.01.01.00056"},
    {"code": "4738", "name": "AMOSTRA GRÁTIS", "grupo": "Resultado", "mask": "4.2.01.02.00001"},
    {"code": "4739", "name": "PROMOÇÕES E EVENTOS", "grupo": "Resultado", "mask": "4.2.01.02.00002"},
    {"code": "4740", "name": "PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "4.2.01.02.00003"},
    {"code": "4741", "name": "REPRESENTAÇÃO COMERCIAL", "grupo": "Resultado", "mask": "4.2.01.02.00004"},
    {"code": "4742", "name": "SERVIÇOS GRÁFICOS", "grupo": "Resultado", "mask": "4.2.01.02.00005"},
    {"code": "4743", "name": "ANÚNCIOS E PUBLICAÇÕES", "grupo": "Resultado", "mask": "4.2.01.02.00006"},
    {"code": "4744", "name": "PATROCÍNIO", "grupo": "Resultado", "mask": "4.2.01.02.00007"},
    {"code": "4745", "name": "FEIRAS E EXPOSIÇÕES", "grupo": "Resultado", "mask": "4.2.01.02.00008"},
    {"code": "4746", "name": "SIMPÓSIOS E CONFERÊNCIAS", "grupo": "Resultado", "mask": "4.2.01.02.00009"},
    {"code": "4747", "name": "BRINDES PROMOCIONAIS", "grupo": "Resultado", "mask": "4.2.01.02.00010"},
    {"code": "4748", "name": "TAXAS DE PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "4.2.01.02.00011"},
    {"code": "4768", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "4.2.01.03.00001"},
    {"code": "4769", "name": "CORREIOS E SIMILARES", "grupo": "Resultado", "mask": "4.2.01.03.00002"},
    {"code": "4770", "name": "FRETES E CARRETOS", "grupo": "Resultado", "mask": "4.2.01.03.00003"},
    {"code": "4771", "name": "CONDUÇÕES", "grupo": "Resultado", "mask": "4.2.01.03.00004"},
    {"code": "4772", "name": "LOCAÇÃO DE BENS", "grupo": "Resultado", "mask": "4.2.01.03.00005"},
    {"code": "4773", "name": "SEGUROS", "grupo": "Resultado", "mask": "4.2.01.03.00006"},
    {"code": "4774", "name": "DESPESAS COM MANUTENÇÃO DE VEICULOS", "grupo": "Resultado", "mask": "4.2.01.03.00007"},
    {"code": "4775", "name": "OUTRAS DESPESAS COM ENTREGAS", "grupo": "Resultado", "mask": "4.2.01.03.00008"},
    {"code": "4785", "name": "ALUGUEL DE VEÍCULOS", "grupo": "Resultado", "mask": "4.2.01.04.00001"},
    {"code": "4786", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "4.2.01.04.00002"},
    {"code": "4787", "name": "ESTACIONAMENTOS E PEDÁGIOS", "grupo": "Resultado", "mask": "4.2.01.04.00003"},
    {"code": "4788", "name": "ESTACIONAMENTOS", "grupo": "Resultado", "mask": "4.2.01.04.00004"},
    {"code": "4789", "name": "REFEIÇÕES", "grupo": "Resultado", "mask": "4.2.01.04.00005"},
    {"code": "4790", "name": "HOSPEDAGEM", "grupo": "Resultado", "mask": "4.2.01.04.00006"},
    {"code": "4791", "name": "VIAGENS E ESTADIAS", "grupo": "Resultado", "mask": "4.2.01.04.00007"},
    {"code": "4792", "name": "PASSAGENS E LOCOMOÇÕES", "grupo": "Resultado", "mask": "4.2.01.04.00008"},
    {"code": "4793", "name": "DESPESAS DIVERSAS", "grupo": "Resultado", "mask": "4.2.01.04.00009"},
    {"code": "4813", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "4.2.01.05.00001"},
    {"code": "4814", "name": "ÁGUA", "grupo": "Resultado", "mask": "4.2.01.05.00002"},
    {"code": "4815", "name": "ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "4.2.01.05.00003"},
    {"code": "4816", "name": "PERDA COM ESTOQUES", "grupo": "Resultado", "mask": "4.2.01.05.00004"},
    {"code": "4817", "name": "MATERIAL DE EMBALAGEM", "grupo": "Resultado", "mask": "4.2.01.05.00005"},
    {"code": "4818", "name": "MATERIAL DE MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.2.01.05.00006"},
    {"code": "4819", "name": "ARMAZENAGEM E DEPOSITO", "grupo": "Resultado", "mask": "4.2.01.05.00007"},
    {"code": "4820", "name": "CUSTO COM INSTALAÇÕES", "grupo": "Resultado", "mask": "4.2.01.05.00008"},
    {"code": "4821", "name": "LANCHES E REFEIÇÕES", "grupo": "Resultado", "mask": "4.2.01.05.00009"},
    {"code": "4822", "name": "LANCHES E REFEIÇÕES - DIRIGENTES", "grupo": "Resultado", "mask": "4.2.01.05.00010"},
    {"code": "4823", "name": "COPA E COZINHA", "grupo": "Resultado", "mask": "4.2.01.05.00011"},
    {"code": "4824", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "4.2.01.05.00012"},
    {"code": "4825", "name": "FARMÁCIA", "grupo": "Resultado", "mask": "4.2.01.05.00013"},
    {"code": "4826", "name": "UNIFORMES", "grupo": "Resultado", "mask": "4.2.01.05.00014"},
    {"code": "4827", "name": "MATERIAL DE USO E CONSUMO", "grupo": "Resultado", "mask": "4.2.01.05.00015"},
    {"code": "4828", "name": "MATERIAIS DE HIGIENE E LIMPEZA", "grupo": "Resultado", "mask": "4.2.01.05.00016"},
    {"code": "4829", "name": "MATERIAIS DE ESCRITÓRIO", "grupo": "Resultado", "mask": "4.2.01.05.00017"},
    {"code": "4830", "name": "MATERIAIS DE INFORMÁTICA", "grupo": "Resultado", "mask": "4.2.01.05.00018"},
    {"code": "4832", "name": "MATERIAIS PUBLICITÁRIOS", "grupo": "Resultado", "mask": "4.2.01.05.00020"},
    {"code": "4833", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "4.2.01.05.00021"},
    {"code": "4834", "name": "BENS DE PEQUENO VALOR", "grupo": "Resultado", "mask": "4.2.01.05.00022"},
    {"code": "4835", "name": "COMPUTADORES E PERIFÉRICOS", "grupo": "Resultado", "mask": "4.2.01.05.00023"},
    {"code": "4837", "name": "BENFEITORIAS EM IMÓVEIS DE TERCEIROS", "grupo": "Resultado", "mask": "4.2.01.05.00025"},
    {"code": "4838", "name": "COMUNICAÇÕES", "grupo": "Resultado", "mask": "4.2.01.05.00026"},
    {"code": "4839", "name": "TELEFONE FIXO", "grupo": "Resultado", "mask": "4.2.01.05.00027"},
    {"code": "4840", "name": "TELEFONE MOVEL", "grupo": "Resultado", "mask": "4.2.01.05.00028"},
    {"code": "4841", "name": "INTERNET", "grupo": "Resultado", "mask": "4.2.01.05.00029"},
    {"code": "4842", "name": "SERVIÇOS PRESTADOS DE PROGRAMAÇÃO E TI", "grupo": "Resultado", "mask": "4.2.01.05.00030"},
    {"code": "4843", "name": "DESPESAS COM INFRAESTRUTURA", "grupo": "Resultado", "mask": "4.2.01.05.00031"},
    {"code": "4844", "name": "CUSTO COM HONORÁRIOS", "grupo": "Resultado", "mask": "4.2.01.05.00032"},
    {"code": "4845", "name": "MANUTENÇÃO DE SOFTWARE", "grupo": "Resultado", "mask": "4.2.01.05.00033"},
    {"code": "4846", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "4.2.01.05.00034"},
    {"code": "4847", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA FISICA", "grupo": "Resultado", "mask": "4.2.01.05.00035"},
    {"code": "4848", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.2.01.05.00036"},
    {"code": "4849", "name": "EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.2.01.05.00037"},
    {"code": "4850", "name": "JUROS S/EMPRÉSTIMOS", "grupo": "Resultado", "mask": "4.2.01.05.00038"},
    {"code": "4851", "name": "ENCARGOS FINANCEIROS", "grupo": "Resultado", "mask": "4.2.01.05.00039"},
    {"code": "4852", "name": "ROYALTIES E ASSISTÊNCIA TÉCNICA PAÍS", "grupo": "Resultado", "mask": "4.2.01.05.00040"},
    {"code": "4853", "name": "AJUSTES - PREÇO DE TRANSFERÊNCIA", "grupo": "Resultado", "mask": "4.2.01.05.00041"},
    {"code": "4854", "name": "INDUSTRIALIZAÇÃO EFETUADO POR TERCEIROS", "grupo": "Resultado", "mask": "4.2.01.05.00042"},
    {"code": "4855", "name": "ASSESSORIA CONTÁBIL", "grupo": "Resultado", "mask": "4.2.01.05.00043"},
    {"code": "4856", "name": "CONSULTORIA E SUPORTE TÉCNICO", "grupo": "Resultado", "mask": "4.2.01.05.00044"},
    {"code": "4857", "name": "SERVIÇOS DE ADVOGADOS", "grupo": "Resultado", "mask": "4.2.01.05.00045"},
    {"code": "4858", "name": "SERVIÇOS DE AUDITORIA", "grupo": "Resultado", "mask": "4.2.01.05.00046"},
    {"code": "4859", "name": "SERVIÇOS DE CONSULTORIA", "grupo": "Resultado", "mask": "4.2.01.05.00047"},
    {"code": "4860", "name": "SERVIÇOS DE DESPACHANTES ADUANEIRO", "grupo": "Resultado", "mask": "4.2.01.05.00048"},
    {"code": "4861", "name": "ASSESSORIA ADMINISTRATIVA", "grupo": "Resultado", "mask": "4.2.01.05.00049"},
    {"code": "4862", "name": "ADMINISTRAÇÃO DE BENS", "grupo": "Resultado", "mask": "4.2.01.05.00050"},
    {"code": "4863", "name": "ANALISES TÉCNICAS", "grupo": "Resultado", "mask": "4.2.01.05.00051"},
    {"code": "4864", "name": "PESQUISA/PROJETO", "grupo": "Resultado", "mask": "4.2.01.05.00052"},
    {"code": "4865", "name": "CURSOS E TREINAMENTOS", "grupo": "Resultado", "mask": "4.2.01.05.00053"},
    {"code": "4866", "name": "DATILOGRAFIA", "grupo": "Resultado", "mask": "4.2.01.05.00054"},
    {"code": "4867", "name": "TRADUÇÕES E VERSÕES", "grupo": "Resultado", "mask": "4.2.01.05.00055"},
    {"code": "4868", "name": "SERVIÇOS DE INTERMEDIAÇÃO", "grupo": "Resultado", "mask": "4.2.01.05.00056"},
    {"code": "4869", "name": "ASSESSORIA E CONSULTORIA EM INFORMÁTICA", "grupo": "Resultado", "mask": "4.2.01.05.00057"},
    {"code": "4870", "name": "LICENÇA DE SOFTWARE", "grupo": "Resultado", "mask": "4.2.01.05.00058"},
    {"code": "4871", "name": "ANALISE E DESENVOLVIMENTO DE SISTEMAS", "grupo": "Resultado", "mask": "4.2.01.05.00059"},
    {"code": "4872", "name": "PROCESSAMENTO E ARMAZENAGEM DE DADOS", "grupo": "Resultado", "mask": "4.2.01.05.00060"},
    {"code": "4873", "name": "PROGRAMAÇÃO E COMUNICAÇÃO VISUAL", "grupo": "Resultado", "mask": "4.2.01.05.00061"},
    {"code": "4874", "name": "ASSISTÊNCIA TÉCNICA", "grupo": "Resultado", "mask": "4.2.01.05.00062"},
    {"code": "4876", "name": "FEIRAS E EVENTOS", "grupo": "Resultado", "mask": "4.2.01.05.00064"},
    {"code": "4877", "name": "PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "4.2.01.05.00065"},
    {"code": "4878", "name": "CARTÓRIO", "grupo": "Resultado", "mask": "4.2.01.05.00066"},
    {"code": "4879", "name": "DESPESAS COM ANIMAIS E MÉDICOS VETERINÁRIOS", "grupo": "Resultado", "mask": "4.2.01.05.00067"},
    {"code": "4880", "name": "SERVIÇOS MÉDICOS E EXAMES LABORATORIAS", "grupo": "Resultado", "mask": "4.2.01.05.00068"},
    {"code": "4881", "name": "ENGENHARIA E ARQUITETURA", "grupo": "Resultado", "mask": "4.2.01.05.00069"},
    {"code": "4882", "name": "MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "4.2.01.05.00070"},
    {"code": "4883", "name": "MANUTENÇÃO E CONSERVAÇÃO", "grupo": "Resultado", "mask": "4.2.01.05.00071"},
    {"code": "4884", "name": "INSTALAÇÕES E MONTAGENS", "grupo": "Resultado", "mask": "4.2.01.05.00072"},
    {"code": "4885", "name": "RECRUTAMENTO E SELEÇÃO", "grupo": "Resultado", "mask": "4.2.01.05.00073"},
    {"code": "4886", "name": "LOCAÇÃO DE MÃO DE OBRA", "grupo": "Resultado", "mask": "4.2.01.05.00074"},
    {"code": "4887", "name": "SERVIÇOS DE LIMPEZA", "grupo": "Resultado", "mask": "4.2.01.05.00075"},
    {"code": "4888", "name": "DEDETIZAÇÃO", "grupo": "Resultado", "mask": "4.2.01.05.00076"},
    {"code": "4889", "name": "SEGURANÇA", "grupo": "Resultado", "mask": "4.2.01.05.00077"},
    {"code": "4890", "name": "SEGUROS GERAIS", "grupo": "Resultado", "mask": "4.2.01.05.00078"},
    {"code": "4891", "name": "SEGUROS - RESPONSABILIDADE CÍVIL", "grupo": "Resultado", "mask": "4.2.01.05.00079"},
    {"code": "4892", "name": "SEGUROS - IMOBILIZADO", "grupo": "Resultado", "mask": "4.2.01.05.00080"},
    {"code": "4894", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "4.2.01.05.00082"},
    {"code": "4895", "name": "ALUGUEL DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "4.2.01.05.00083"},
    {"code": "4896", "name": "ALUGUEL - PJ", "grupo": "Resultado", "mask": "4.2.01.05.00084"},
    {"code": "4897", "name": "ALUGUEL - PF", "grupo": "Resultado", "mask": "4.2.01.05.00085"},
    {"code": "4898", "name": "OUTROS CUSTOS", "grupo": "Resultado", "mask": "4.2.01.05.00086"},
    {"code": "4899", "name": "CREDITO PRESUMIDO DE IMPOSTOS", "grupo": "Resultado", "mask": "4.2.01.05.00087"},
    {"code": "4900", "name": "DESPESAS COM DEPRECIAÇÃO", "grupo": "Resultado", "mask": "4.2.01.05.00088"},
    {"code": "4901", "name": "DESPESAS COM AMORTIZAÇÃO", "grupo": "Resultado", "mask": "4.2.01.05.00089"},
    {"code": "4902", "name": "REVERSÃO - CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "4.2.01.05.00090"},
    # ── Resultado - Despesas Gerais (5.x) ────────────────────────────────────
    {"code": "5000", "name": "SALÁRIOS E ORDENADOS", "grupo": "Resultado", "mask": "5.1.01.01.00001"},
    {"code": "5001", "name": "FÉRIAS", "grupo": "Resultado", "mask": "5.1.01.01.00002"},
    {"code": "5002", "name": "13º SALÁRIO", "grupo": "Resultado", "mask": "5.1.01.01.00003"},
    {"code": "5003", "name": "HORA EXTRA", "grupo": "Resultado", "mask": "5.1.01.01.00004"},
    {"code": "5004", "name": "ABONO PECUNIÁRIO", "grupo": "Resultado", "mask": "5.1.01.01.00005"},
    {"code": "5005", "name": "DESCANSO SEMANAL REMUNERADO", "grupo": "Resultado", "mask": "5.1.01.01.00006"},
    {"code": "5006", "name": "BÔNUS", "grupo": "Resultado", "mask": "5.1.01.01.00007"},
    {"code": "5007", "name": "VERBAS RESCISÓRIAS", "grupo": "Resultado", "mask": "5.1.01.01.00008"},
    {"code": "5008", "name": "ADICIONAL NOTURNO", "grupo": "Resultado", "mask": "5.1.01.01.00009"},
    {"code": "5009", "name": "ADICIONAL DE PERICULOSIDADE", "grupo": "Resultado", "mask": "5.1.01.01.00010"},
    {"code": "5010", "name": "ADICIONAL POR TEMPO DE SERVIÇO", "grupo": "Resultado", "mask": "5.1.01.01.00011"},
    {"code": "5011", "name": "ADICIONAL DE INSALUBRIDADE", "grupo": "Resultado", "mask": "5.1.01.01.00012"},
    {"code": "5012", "name": "AUXÍLIO HOME OFFICE", "grupo": "Resultado", "mask": "5.1.01.01.00013"},
    {"code": "5013", "name": "CONTRIBUIÇÃO ASSISTENCIAL", "grupo": "Resultado", "mask": "5.1.01.01.00014"},
    {"code": "5014", "name": "INDENIZAÇÕES", "grupo": "Resultado", "mask": "5.1.01.01.00015"},
    {"code": "5015", "name": "MULTA ESTABILIDADE", "grupo": "Resultado", "mask": "5.1.01.01.00016"},
    {"code": "5016", "name": "CONTRIBUIÇÃO SINDICAL", "grupo": "Resultado", "mask": "5.1.01.01.00017"},
    {"code": "5017", "name": "ANUÊNIO E TRIÊNIO", "grupo": "Resultado", "mask": "5.1.01.01.00018"},
    {"code": "5018", "name": "QUINQUENIO", "grupo": "Resultado", "mask": "5.1.01.01.00019"},
    {"code": "5019", "name": "OUTROS PROVENTOS RPA", "grupo": "Resultado", "mask": "5.1.01.01.00020"},
    {"code": "5020", "name": "BOLSA GRATIFICAÇÃO", "grupo": "Resultado", "mask": "5.1.01.01.00021"},
    {"code": "5021", "name": "PROVISÕES E ENCARGOS", "grupo": "Resultado", "mask": "5.1.01.01.00022"},
    {"code": "5022", "name": "AUXÍLIO CRECHE", "grupo": "Resultado", "mask": "5.1.01.01.00023"},
    {"code": "5023", "name": "BOLSA AUXÍLIO", "grupo": "Resultado", "mask": "5.1.01.01.00024"},
    {"code": "5024", "name": "UNIFORMES", "grupo": "Resultado", "mask": "5.1.01.01.00025"},
    {"code": "5025", "name": "SERVIÇOS AUTÔNOMO", "grupo": "Resultado", "mask": "5.1.01.01.00026"},
    {"code": "5026", "name": "PARTICIPAÇÃO NOS LUCROS E RESULTADOS", "grupo": "Resultado", "mask": "5.1.01.01.00027"},
    {"code": "5027", "name": "PRÊMIOS E GRATIFICAÇÕES", "grupo": "Resultado", "mask": "5.1.01.01.00028"},
    {"code": "5028", "name": "DONATIVOS E GORJETAS", "grupo": "Resultado", "mask": "5.1.01.01.00029"},
    {"code": "5029", "name": "COMISSÕES", "grupo": "Resultado", "mask": "5.1.01.01.00030"},
    {"code": "5030", "name": "PRÓ-LABORE", "grupo": "Resultado", "mask": "5.1.01.01.00031"},
    {"code": "5031", "name": "FGTS", "grupo": "Resultado", "mask": "5.1.01.01.00032"},
    {"code": "5032", "name": "INSS", "grupo": "Resultado", "mask": "5.1.01.01.00033"},
    {"code": "5033", "name": "INSS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "5.1.01.01.00034"},
    {"code": "5034", "name": "INSS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "5.1.01.01.00035"},
    {"code": "5035", "name": "FGTS - PROVISÃO DE FÉRIAS", "grupo": "Resultado", "mask": "5.1.01.01.00036"},
    {"code": "5036", "name": "FGTS - PROVISÃO DE 13O. SALÁRIO", "grupo": "Resultado", "mask": "5.1.01.01.00037"},
    {"code": "5037", "name": "MULTA FGTS", "grupo": "Resultado", "mask": "5.1.01.01.00038"},
    {"code": "5038", "name": "AJUDA DE CUSTO", "grupo": "Resultado", "mask": "5.1.01.01.00039"},
    {"code": "5039", "name": "KIT FUNCIONÁRIO", "grupo": "Resultado", "mask": "5.1.01.01.00040"},
    {"code": "5040", "name": "CESTA BÁSICA", "grupo": "Resultado", "mask": "5.1.01.01.00041"},
    {"code": "5041", "name": "VALE ALIMENTAÇÃO/REFEIÇÃO", "grupo": "Resultado", "mask": "5.1.01.01.00042"},
    {"code": "5042", "name": "VALE TRANSPORTE", "grupo": "Resultado", "mask": "5.1.01.01.00043"},
    {"code": "5043", "name": "VALE COMBUSTÍVEL", "grupo": "Resultado", "mask": "5.1.01.01.00044"},
    {"code": "5044", "name": "BENEFÍCIOS TRABALHISTAS FLEXÍVEIS", "grupo": "Resultado", "mask": "5.1.01.01.00045"},
    {"code": "5045", "name": "SEGURO DE VIDA", "grupo": "Resultado", "mask": "5.1.01.01.00046"},
    {"code": "5046", "name": "ASSISTÊNCIA MÉDICA", "grupo": "Resultado", "mask": "5.1.01.01.00047"},
    {"code": "5047", "name": "ASSISTÊNCIA ODONTOLÓGICA", "grupo": "Resultado", "mask": "5.1.01.01.00048"},
    {"code": "5048", "name": "PIS SOBRE FOLHA", "grupo": "Resultado", "mask": "5.1.01.01.00049"},
    {"code": "5049", "name": "FUNDO DE APOSENTADORIA", "grupo": "Resultado", "mask": "5.1.01.01.00050"},
    {"code": "5050", "name": "PLANO DE POUPANÇA E INVESTIMENTO", "grupo": "Resultado", "mask": "5.1.01.01.00051"},
    {"code": "5051", "name": "VALE CULTURA", "grupo": "Resultado", "mask": "5.1.01.01.00052"},
    {"code": "5052", "name": "REFEIÇÕES PAT", "grupo": "Resultado", "mask": "5.1.01.01.00053"},
    {"code": "5053", "name": "PROCESSOS TRABALHISTAS", "grupo": "Resultado", "mask": "5.1.01.01.00054"},
    {"code": "5054", "name": "PREVIDÊNCIA PRIVADA", "grupo": "Resultado", "mask": "5.1.01.01.00055"},
    {"code": "5055", "name": "OUTRAS DESPESAS COM PESSOAL", "grupo": "Resultado", "mask": "5.1.01.01.00056"},
    {"code": "5075", "name": "ÁGUA E ESGOTO", "grupo": "Resultado", "mask": "5.1.01.02.00001"},
    {"code": "5076", "name": "ALUGUEL - PJ", "grupo": "Resultado", "mask": "5.1.01.02.00002"},
    {"code": "5077", "name": "ALUGUEL - PF", "grupo": "Resultado", "mask": "5.1.01.02.00003"},
    {"code": "5078", "name": "CONDOMÍNIO", "grupo": "Resultado", "mask": "5.1.01.02.00004"},
    {"code": "5079", "name": "SEGUROS DE IMÓVEIS", "grupo": "Resultado", "mask": "5.1.01.02.00005"},
    {"code": "5080", "name": "ENERGIA ELÉTRICA", "grupo": "Resultado", "mask": "5.1.01.02.00006"},
    {"code": "5081", "name": "GÁS", "grupo": "Resultado", "mask": "5.1.01.02.00007"},
    {"code": "5082", "name": "MANUTENÇÃO E REPAROS", "grupo": "Resultado", "mask": "5.1.01.02.00008"},
    {"code": "5083", "name": "INTERNET", "grupo": "Resultado", "mask": "5.1.01.02.00009"},
    {"code": "5084", "name": "TV POR ASSINATURA", "grupo": "Resultado", "mask": "5.1.01.02.00010"},
    {"code": "5085", "name": "ALUGUEL DE VEÍCULOS", "grupo": "Resultado", "mask": "5.1.01.02.00011"},
    {"code": "5086", "name": "TELEFONE FIXO", "grupo": "Resultado", "mask": "5.1.01.02.00012"},
    {"code": "5087", "name": "TELEFONE MÓVEL", "grupo": "Resultado", "mask": "5.1.01.02.00013"},
    {"code": "5088", "name": "COMUNICAÇÕES", "grupo": "Resultado", "mask": "5.1.01.02.00014"},
    {"code": "5089", "name": "OUTRAS DESPESAS COM OCUPAÇÃO E UTILIDADES", "grupo": "Resultado", "mask": "5.1.01.02.00015"},
    {"code": "5105", "name": "BENS DE PEQUENO VALOR", "grupo": "Resultado", "mask": "5.1.01.03.00001"},
    {"code": "5106", "name": "MATERIAIS DE INFORMÁTICA", "grupo": "Resultado", "mask": "5.1.01.03.00002"},
    {"code": "5107", "name": "MATERIAIS DE SEGURANÇA", "grupo": "Resultado", "mask": "5.1.01.03.00003"},
    {"code": "5108", "name": "MATERIAL DE ESCRITÓRIO", "grupo": "Resultado", "mask": "5.1.01.03.00004"},
    {"code": "5109", "name": "MATERIAL DE USO E CONSUMO", "grupo": "Resultado", "mask": "5.1.01.03.00005"},
    {"code": "5110", "name": "MATERIAL PARA MANUTENÇÃO DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "5.1.01.03.00006"},
    {"code": "5111", "name": "MATERIAL PARA MANUTENÇÃO DE VEÍCULOS", "grupo": "Resultado", "mask": "5.1.01.03.00007"},
    {"code": "5112", "name": "MATERIAL PARA MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "5.1.01.03.00008"},
    {"code": "5113", "name": "MATERIAL DE LIMPEZA", "grupo": "Resultado", "mask": "5.1.01.03.00009"},
    {"code": "5114", "name": "MATERIAL DE EMBALAGEM", "grupo": "Resultado", "mask": "5.1.01.03.00010"},
    {"code": "5134", "name": "ASSESSORIA E CONSULTORIA EM INFORMÁTICA", "grupo": "Resultado", "mask": "5.1.01.04.00001"},
    {"code": "5135", "name": "PROCESSAMENTO E ARMAZENAGEM DE DADOS", "grupo": "Resultado", "mask": "5.1.01.04.00002"},
    {"code": "5136", "name": "LICENÇA DE SOFTWARE", "grupo": "Resultado", "mask": "5.1.01.04.00003"},
    {"code": "5137", "name": "ASSISTÊNCIA E SUPORTE TÉCNICO", "grupo": "Resultado", "mask": "5.1.01.04.00004"},
    {"code": "5138", "name": "REGISTRO DE MARCAS E PATENTES", "grupo": "Resultado", "mask": "5.1.01.04.00005"},
    {"code": "5139", "name": "PESQUISAS CIENTÍFICAS E TECNOLÓGICAS", "grupo": "Resultado", "mask": "5.1.01.04.00006"},
    {"code": "5140", "name": "PESQUISA E DESENVOLVIMENTO (P&D) - LEI DO BEM 11.196/2005", "grupo": "Resultado", "mask": "5.1.01.04.00007"},
    {"code": "5141", "name": "PROGRAMAÇÃO E COMUNICAÇÃO VISUAL", "grupo": "Resultado", "mask": "5.1.01.04.00008"},
    {"code": "5142", "name": "DESPESAS COM GERENCIAMENTO DE RISCOS E LGPD", "grupo": "Resultado", "mask": "5.1.01.04.00009"},
    {"code": "5143", "name": "DESPESAS COM INFRAESTRUTURA", "grupo": "Resultado", "mask": "5.1.01.04.00010"},
    {"code": "5144", "name": "ANALISE E DESENVOLVIMENTO DE SISTEMAS", "grupo": "Resultado", "mask": "5.1.01.04.00011"},
    {"code": "5164", "name": "ASSESSORIA ADMINISTRATIVA", "grupo": "Resultado", "mask": "5.1.01.05.00001"},
    {"code": "5165", "name": "ASSESSORIA JURÍDICA", "grupo": "Resultado", "mask": "5.1.01.05.00002"},
    {"code": "5166", "name": "ASSESSORIA CONTÁBIL", "grupo": "Resultado", "mask": "5.1.01.05.00003"},
    {"code": "5167", "name": "SERVIÇOS DE AUDITORIA", "grupo": "Resultado", "mask": "5.1.01.05.00004"},
    {"code": "5168", "name": "SERVIÇOS DE CONSULTORIA", "grupo": "Resultado", "mask": "5.1.01.05.00005"},
    {"code": "5169", "name": "SERVIÇOS DE ADVOGADOS", "grupo": "Resultado", "mask": "5.1.01.05.00006"},
    {"code": "5170", "name": "CURSOS E TREINAMENTOS", "grupo": "Resultado", "mask": "5.1.01.05.00007"},
    {"code": "5171", "name": "CONSULTORIA E SUPORTE TÉCNICO", "grupo": "Resultado", "mask": "5.1.01.05.00008"},
    {"code": "5172", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "5.1.01.05.00009"},
    {"code": "5173", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA FISICA", "grupo": "Resultado", "mask": "5.1.01.05.00010"},
    {"code": "5174", "name": "SERVIÇOS DE DESPACHANTES ADUANEIRO", "grupo": "Resultado", "mask": "5.1.01.05.00011"},
    {"code": "5175", "name": "SERVIÇOS DE DATILOGRAFIA", "grupo": "Resultado", "mask": "5.1.01.05.00012"},
    {"code": "5195", "name": "ADMINISTRAÇÃO DE BENS", "grupo": "Resultado", "mask": "5.1.01.06.00001"},
    {"code": "5196", "name": "AGENCIAMENTO E CORRETAGEM", "grupo": "Resultado", "mask": "5.1.01.06.00002"},
    {"code": "5197", "name": "ALUGUEL DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "5.1.01.06.00003"},
    {"code": "5198", "name": "ARMAZENAGEM E DEPOSITO", "grupo": "Resultado", "mask": "5.1.01.06.00004"},
    {"code": "5199", "name": "ASSINATURAS E MENSALIDADES", "grupo": "Resultado", "mask": "5.1.01.06.00005"},
    {"code": "5200", "name": "SERVIÇOS E ANALISES TÉCNICAS", "grupo": "Resultado", "mask": "5.1.01.06.00006"},
    {"code": "5201", "name": "DESPESAS COM DEPRECIAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00007"},
    {"code": "5202", "name": "DESPESAS COM AMORTIZAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00008"},
    {"code": "5203", "name": "CARTÓRIO", "grupo": "Resultado", "mask": "5.1.01.06.00009"},
    {"code": "5204", "name": "COMBUSTÍVEIS E LUBRIFICANTES", "grupo": "Resultado", "mask": "5.1.01.06.00010"},
    {"code": "5205", "name": "CONFRATERNIZAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00011"},
    {"code": "5206", "name": "CONSELHO REGIONAL", "grupo": "Resultado", "mask": "5.1.01.06.00012"},
    {"code": "5207", "name": "CORREIOS E SIMILARES", "grupo": "Resultado", "mask": "5.1.01.06.00013"},
    {"code": "5209", "name": "DEDETIZAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00015"},
    {"code": "5210", "name": "DESPESAS GERAIS COM FUNCIONÁRIOS", "grupo": "Resultado", "mask": "5.1.01.06.00016"},
    {"code": "5211", "name": "FRETES E CARRETOS", "grupo": "Resultado", "mask": "5.1.01.06.00017"},
    {"code": "5212", "name": "BENS DE USO PERMANENTE - DEDUTÍVEL", "grupo": "Resultado", "mask": "5.1.01.06.00018"},
    {"code": "5213", "name": "BRINDES E BONIFICAÇÕES", "grupo": "Resultado", "mask": "5.1.01.06.00019"},
    {"code": "5214", "name": "COPA E COZINHA", "grupo": "Resultado", "mask": "5.1.01.06.00020"},
    {"code": "5215", "name": "ESTACIONAMENTOS E PEDÁGIOS", "grupo": "Resultado", "mask": "5.1.01.06.00021"},
    {"code": "5216", "name": "FARMÁCIA", "grupo": "Resultado", "mask": "5.1.01.06.00022"},
    {"code": "5217", "name": "FEIRAS E EVENTOS", "grupo": "Resultado", "mask": "5.1.01.06.00023"},
    {"code": "5218", "name": "LOCAÇÃO DE MÃO DE OBRA", "grupo": "Resultado", "mask": "5.1.01.06.00024"},
    {"code": "5219", "name": "FOTOGRAFIA", "grupo": "Resultado", "mask": "5.1.01.06.00025"},
    {"code": "5221", "name": "MANUTENÇÃO E CONSERVAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00027"},
    {"code": "5222", "name": "MANUTENÇÃO DE SOFTWARE", "grupo": "Resultado", "mask": "5.1.01.06.00028"},
    {"code": "5223", "name": "MANUTENÇÃO DE MÁQUINAS E EQUIPAMENTOS", "grupo": "Resultado", "mask": "5.1.01.06.00029"},
    {"code": "5224", "name": "MANUTENÇÃO DE VEÍCULOS", "grupo": "Resultado", "mask": "5.1.01.06.00030"},
    {"code": "5225", "name": "MANUTENÇÃO PREDIAL", "grupo": "Resultado", "mask": "5.1.01.06.00031"},
    {"code": "5226", "name": "SERVIÇOS MÉDICOS E EXAMES LABORATORIAS", "grupo": "Resultado", "mask": "5.1.01.06.00032"},
    {"code": "5227", "name": "MULTAS CONTRATUAIS", "grupo": "Resultado", "mask": "5.1.01.06.00033"},
    {"code": "5228", "name": "MULTAS DE TRÂNSITO", "grupo": "Resultado", "mask": "5.1.01.06.00034"},
    {"code": "5229", "name": "PROCESSOS JUDICIAIS", "grupo": "Resultado", "mask": "5.1.01.06.00035"},
    {"code": "5230", "name": "LEGAIS E JUDICIAIS", "grupo": "Resultado", "mask": "5.1.01.06.00036"},
    {"code": "5231", "name": "RECRUTAMENTO E SELEÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00037"},
    {"code": "5232", "name": "ENGENHARIA E ARQUITETURA", "grupo": "Resultado", "mask": "5.1.01.06.00038"},
    {"code": "5233", "name": "SERVIÇOS DE INTERMEDIAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00039"},
    {"code": "5234", "name": "SEGURANÇA", "grupo": "Resultado", "mask": "5.1.01.06.00040"},
    {"code": "5235", "name": "SEGUROS GERAIS", "grupo": "Resultado", "mask": "5.1.01.06.00041"},
    {"code": "5236", "name": "SEGUROS - RESPONSABILIDADE CÍVIL", "grupo": "Resultado", "mask": "5.1.01.06.00042"},
    {"code": "5237", "name": "SEGUROS - IMOBILIZADO", "grupo": "Resultado", "mask": "5.1.01.06.00043"},
    {"code": "5238", "name": "TRADUÇÕES E VERSÕES", "grupo": "Resultado", "mask": "5.1.01.06.00044"},
    {"code": "5239", "name": "DESPESAS COM PATROCÍNIOS E PUBLICAÇÕES", "grupo": "Resultado", "mask": "5.1.01.06.00045"},
    {"code": "5240", "name": "REPRESENTAÇÕES", "grupo": "Resultado", "mask": "5.1.01.06.00046"},
    {"code": "5241", "name": "SERVIÇOS DE PESQUISAS E SERVIÇOS TÉCNICOS", "grupo": "Resultado", "mask": "5.1.01.06.00047"},
    {"code": "5242", "name": "PESQUISA/PROJETO", "grupo": "Resultado", "mask": "5.1.01.06.00048"},
    {"code": "5243", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "5.1.01.06.00049"},
    {"code": "5244", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA FISICA", "grupo": "Resultado", "mask": "5.1.01.06.00050"},
    {"code": "5245", "name": "DESENVOLVIMENTO", "grupo": "Resultado", "mask": "5.1.01.06.00051"},
    {"code": "5246", "name": "SERVIÇOS GRÁFICOS", "grupo": "Resultado", "mask": "5.1.01.06.00052"},
    {"code": "5247", "name": "ANÚNCIOS E PUBLICAÇÕES", "grupo": "Resultado", "mask": "5.1.01.06.00053"},
    {"code": "5248", "name": "CURSOS E TREINAMENTOS", "grupo": "Resultado", "mask": "5.1.01.06.00054"},
    {"code": "5249", "name": "LEASING", "grupo": "Resultado", "mask": "5.1.01.06.00055"},
    {"code": "5250", "name": "AMORTIZAÇÃO DE DIREITO DE USO", "grupo": "Resultado", "mask": "5.1.01.06.00056"},
    {"code": "5251", "name": "ARRENDAMENTO FINANCEIRO", "grupo": "Resultado", "mask": "5.1.01.06.00057"},
    {"code": "5252", "name": "SUPERMERCADO", "grupo": "Resultado", "mask": "5.1.01.06.00058"},
    {"code": "5253", "name": "CERTIFICAÇÃO DIGITAL", "grupo": "Resultado", "mask": "5.1.01.06.00059"},
    {"code": "5254", "name": "ALIMENTAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00060"},
    {"code": "5255", "name": "IMPORTAÇÃO DE SERVIÇOS", "grupo": "Resultado", "mask": "5.1.01.06.00061"},
    {"code": "5256", "name": "OUTRAS DESPESAS COM IMPORTAÇÃO / EXPORTAÇÃO", "grupo": "Resultado", "mask": "5.1.01.06.00062"},
    {"code": "5257", "name": "SERVIÇOS TOMADOS DO EXTERIOR", "grupo": "Resultado", "mask": "5.1.01.06.00063"},
    {"code": "5258", "name": "INSTALAÇÕES E MONTAGENS", "grupo": "Resultado", "mask": "5.1.01.06.00064"},
    {"code": "5259", "name": "SERVIÇOS DE LIMPEZA", "grupo": "Resultado", "mask": "5.1.01.06.00065"},
    {"code": "5260", "name": "DATILOGRAFIA", "grupo": "Resultado", "mask": "5.1.01.06.00066"},
    {"code": "5261", "name": "VIAGENS E ESTADIAS", "grupo": "Resultado", "mask": "5.1.01.06.00067"},
    {"code": "5262", "name": "PROPAGANDA E PUBLICIDADE", "grupo": "Resultado", "mask": "5.1.01.06.00068"},
    {"code": "5263", "name": "DESPESAS GERAIS", "grupo": "Resultado", "mask": "5.1.01.06.00069"},
    {"code": "5264", "name": "ORNAMENTOS E  DECORAÇÕES", "grupo": "Resultado", "mask": "5.1.01.06.00070"},
    {"code": "5265", "name": "CREDITO DE PIS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "5.1.01.06.00071"},
    {"code": "5266", "name": "CREDITO DE COFINS NÃO CUMULATIVO", "grupo": "Resultado", "mask": "5.1.01.06.00072"},
    {"code": "5267", "name": "AMORTIZAÇÃO DE DESPESAS PRÉ-OPERACIONAIS", "grupo": "Resultado", "mask": "5.1.01.06.00073"},
    {"code": "5268", "name": "DESPESAS COM ANIMAIS E MÉDICOS VETERINÁRIOS", "grupo": "Resultado", "mask": "5.1.01.06.00074"},
    {"code": "5269", "name": "LANCHES E REFEIÇÕES", "grupo": "Resultado", "mask": "5.1.01.06.00075"},
    {"code": "5270", "name": "DOAÇÕES E CONTRIBUIÇÕES DEDUTÍVEIS", "grupo": "Resultado", "mask": "5.1.01.06.00076"},
    {"code": "5271", "name": "AMOSTRA GRÁTIS", "grupo": "Resultado", "mask": "5.1.01.06.00077"},
    {"code": "5272", "name": "PERDAS POR CRÉDITO DE LIQUIDAÇÃO DUVIDOSA - PCLD", "grupo": "Resultado", "mask": "5.1.01.06.00078"},
    {"code": "5320", "name": "AUTO DE INFRAÇÃO", "grupo": "Resultado", "mask": "5.1.02.01.00001"},
    {"code": "5321", "name": "CONTRIBUIÇÕES SINDICAIS", "grupo": "Resultado", "mask": "5.1.02.01.00002"},
    {"code": "5322", "name": "PIS S/ OUTRAS RECEITAS", "grupo": "Resultado", "mask": "5.1.02.01.00003"},
    {"code": "5323", "name": "COFINS S/ OUTRAS RECEITAS", "grupo": "Resultado", "mask": "5.1.02.01.00004"},
    {"code": "5324", "name": "ICMS DIFAL", "grupo": "Resultado", "mask": "5.1.02.01.00005"},
    {"code": "5325", "name": "IMPOSTOS NÃO RECUPERÁVEIS", "grupo": "Resultado", "mask": "5.1.02.01.00006"},
    {"code": "5326", "name": "IOF", "grupo": "Resultado", "mask": "5.1.02.01.00007"},
    {"code": "5327", "name": "IPTU", "grupo": "Resultado", "mask": "5.1.02.01.00008"},
    {"code": "5328", "name": "IPVA", "grupo": "Resultado", "mask": "5.1.02.01.00009"},
    {"code": "5329", "name": "LICENÇA E ALVARÁ", "grupo": "Resultado", "mask": "5.1.02.01.00010"},
    {"code": "5330", "name": "LICENCIAMENTO", "grupo": "Resultado", "mask": "5.1.02.01.00011"},
    {"code": "5331", "name": "MULTAS POR INFRAÇÃO", "grupo": "Resultado", "mask": "5.1.02.01.00012"},
    {"code": "5332", "name": "ICMS SOBRE OUTRAS OPERACÕES", "grupo": "Resultado", "mask": "5.1.02.01.00013"},
    {"code": "5333", "name": "IPI S/ OUTRAS OPERAÇÕES", "grupo": "Resultado", "mask": "5.1.02.01.00014"},
    {"code": "5334", "name": "OUTROS TRIBUTOS E TAXAS", "grupo": "Resultado", "mask": "5.1.02.01.00015"},
    {"code": "5335", "name": "TAXAS AMBIENTAIS", "grupo": "Resultado", "mask": "5.1.02.01.00016"},
    {"code": "5336", "name": "TFE E TFA", "grupo": "Resultado", "mask": "5.1.02.01.00017"},
    {"code": "5337", "name": "TAXAS MUNICIPAIS", "grupo": "Resultado", "mask": "5.1.02.01.00018"},
    {"code": "5338", "name": "IR S/ REMESSA EXTERIOR", "grupo": "Resultado", "mask": "5.1.02.01.00019"},
    {"code": "5339", "name": "CIDE SOBRE IMPORTAÇÃO", "grupo": "Resultado", "mask": "5.1.02.01.00020"},
    {"code": "5340", "name": "PIS SOBRE IMPORTAÇÃO", "grupo": "Resultado", "mask": "5.1.02.01.00021"},
    {"code": "5341", "name": "COFINS SOBRE IMPORTAÇÃO", "grupo": "Resultado", "mask": "5.1.02.01.00022"},
    {"code": "5361", "name": "RECEITA SOBRE APLICAÇÕES FINANCEIRAS", "grupo": "Resultado", "mask": "5.1.03.01.00001"},
    {"code": "5362", "name": "JUROS ATIVOS", "grupo": "Resultado", "mask": "5.1.03.01.00002"},
    {"code": "5363", "name": "MULTAS ATIVAS", "grupo": "Resultado", "mask": "5.1.03.01.00003"},
    {"code": "5364", "name": "DESCONTOS OBTIDOS", "grupo": "Resultado", "mask": "5.1.03.01.00004"},
    {"code": "5365", "name": "JUROS SOBRE O CAPITAL PRÓPRIO", "grupo": "Resultado", "mask": "5.1.03.01.00005"},
    {"code": "5366", "name": "REMISSÃO DE DÍVIDA", "grupo": "Resultado", "mask": "5.1.03.01.00006"},
    {"code": "5367", "name": "REMUNERAÇÕES BANCÁRIAS", "grupo": "Resultado", "mask": "5.1.03.01.00007"},
    {"code": "5368", "name": "GANHO EM APLICAÇÃO SWAP", "grupo": "Resultado", "mask": "5.1.03.01.00008"},
    {"code": "5369", "name": "GANHOS COM OPERAÇÃO DAY-TRADE", "grupo": "Resultado", "mask": "5.1.03.01.00009"},
    {"code": "5370", "name": "RENDAS APLICAÇÕES FINANC. RENDA FIXA", "grupo": "Resultado", "mask": "5.1.03.01.00010"},
    {"code": "5371", "name": "RENDAS APLICAÇÕES FINANC. RENDA VARIÁVEL", "grupo": "Resultado", "mask": "5.1.03.01.00011"},
    {"code": "5372", "name": "VAR.CAMBIAIS/MONET.SUJEITAS A PIS/COFINS", "grupo": "Resultado", "mask": "5.1.03.01.00012"},
    {"code": "5373", "name": "RECEITAS FINANCEIRAS DECORRENTES DE AVP", "grupo": "Resultado", "mask": "5.1.03.01.00013"},
    {"code": "5374", "name": "JUROS SELIC S/ IMPOSTOS A RECUPERAR", "grupo": "Resultado", "mask": "5.1.03.01.00014"},
    {"code": "5381", "name": "VARIAÇÕES MONETÁRIAS ATIVAS", "grupo": "Resultado", "mask": "5.1.03.02.00001"},
    {"code": "5382", "name": "VARIAÇÕES CAMBIAIS ATIVAS", "grupo": "Resultado", "mask": "5.1.03.02.00002"},
    {"code": "5383", "name": "VARIAÇÕES CAMBIAIS ATIVAS - LIQUIDADAS", "grupo": "Resultado", "mask": "5.1.03.02.00003"},
    {"code": "5393", "name": "CUSTAS E EMOLUMENTOS", "grupo": "Resultado", "mask": "5.1.03.03.00001"},
    {"code": "5394", "name": "DESCONTOS CONCEDIDOS", "grupo": "Resultado", "mask": "5.1.03.03.00002"},
    {"code": "5395", "name": "DESPESAS BANCÁRIAS", "grupo": "Resultado", "mask": "5.1.03.03.00003"},
    {"code": "5396", "name": "DESPESAS COM FINANCIAMENTOS", "grupo": "Resultado", "mask": "5.1.03.03.00004"},
    {"code": "5397", "name": "ARRENDAMENTO MERCANTIL", "grupo": "Resultado", "mask": "5.1.03.03.00005"},
    {"code": "5398", "name": "DESPESAS FINANCEIRAS DE LEASING", "grupo": "Resultado", "mask": "5.1.03.03.00006"},
    {"code": "5399", "name": "JUROS", "grupo": "Resultado", "mask": "5.1.03.03.00007"},
    {"code": "5400", "name": "JUROS SOBRE O CAPITAL DE GIRO", "grupo": "Resultado", "mask": "5.1.03.03.00008"},
    {"code": "5401", "name": "JUROS SOBRE O CAPITAL PRÓPRIO", "grupo": "Resultado", "mask": "5.1.03.03.00009"},
    {"code": "5402", "name": "JUROS PASSIVOS-EMPR.PESSOA VINC./PAIS", "grupo": "Resultado", "mask": "5.1.03.03.00010"},
    {"code": "5403", "name": "JUROS S/OPERAÇÃO MÚTUO-PARTE RELACIONADA", "grupo": "Resultado", "mask": "5.1.03.03.00011"},
    {"code": "5404", "name": "MULTA", "grupo": "Resultado", "mask": "5.1.03.03.00012"},
    {"code": "5405", "name": "OUTRAS DESPESAS FINANCEIRAS", "grupo": "Resultado", "mask": "5.1.03.03.00013"},
    {"code": "5406", "name": "PERDAS SOBRE APLICAÇÕES FINANCEIRAS", "grupo": "Resultado", "mask": "5.1.03.03.00014"},
    {"code": "5407", "name": "REMISSÃO DE DÍVIDA", "grupo": "Resultado", "mask": "5.1.03.03.00015"},
    {"code": "5408", "name": "REMUNERAÇÃO DE DEBENTURES", "grupo": "Resultado", "mask": "5.1.03.03.00016"},
    {"code": "5409", "name": "TARIFAS DE EMPRÉSTIMOS", "grupo": "Resultado", "mask": "5.1.03.03.00017"},
    {"code": "5410", "name": "TAXAS ADMINISTRATIVAS DE CARTÃO", "grupo": "Resultado", "mask": "5.1.03.03.00018"},
    {"code": "5411", "name": "TAXAS ADMINISTRATIVAS DE CONSÓRCIO", "grupo": "Resultado", "mask": "5.1.03.03.00019"},
    {"code": "5412", "name": "VARIAÇÕES CAMBIAIS PASSIVAS", "grupo": "Resultado", "mask": "5.1.03.04.00001"},
    {"code": "5413", "name": "VARIAÇÕES CAMBIAIS PASSIVAS - LIQUIDADAS", "grupo": "Resultado", "mask": "5.1.03.04.00002"},
    {"code": "5423", "name": "LUCRO NA ALIENAÇÃO DE IMOBILIZADO", "grupo": "Resultado", "mask": "5.1.04.01.00001"},
    {"code": "5424", "name": "PREJUÍZO NA ALIENAÇÃO DE IMOBILIZADO", "grupo": "Resultado", "mask": "5.1.04.01.00002"},
    {"code": "5425", "name": "CUSTO DO BEM BAIXADO", "grupo": "Resultado", "mask": "5.1.04.01.00003"},
    {"code": "5426", "name": "VALOR ALIENAÇÃO IMOB.EM DESUSO/OBSOLETO", "grupo": "Resultado", "mask": "5.1.04.01.00004"},
    {"code": "5427", "name": "VALOR DE ALIENAÇÃO DO IMOBILIZADO", "grupo": "Resultado", "mask": "5.1.04.01.00005"},
    {"code": "5428", "name": "CUSTO DO BEM EM DESUSO/OBSOLETO BAIXADO", "grupo": "Resultado", "mask": "5.1.04.01.00006"},
    {"code": "5429", "name": "DEPRECIAÇÃO ACUM.BAIXADA BEM DESUSO/OBS", "grupo": "Resultado", "mask": "5.1.04.01.00007"},
    {"code": "5430", "name": "DEPRECIAÇÃO ACUMULADA BAIXADA", "grupo": "Resultado", "mask": "5.1.04.01.00008"},
    {"code": "5440", "name": "LUCRO NA ALIENAÇÃO DE PARTICIPAÇÕES SOCIETÁRIAS", "grupo": "Resultado", "mask": "5.1.05.01.00001"},
    {"code": "5441", "name": "RESULTADO POSIT.PELO MEP - PROPORÇÃO PL", "grupo": "Resultado", "mask": "5.1.05.01.00002"},
    {"code": "5442", "name": "RESULTADO POSIT.PELO MEP - DESPROP.PL", "grupo": "Resultado", "mask": "5.1.05.01.00003"},
    {"code": "5443", "name": "LUCROS E DIVID.PELO MÉTODO CUSTO AQUIS.", "grupo": "Resultado", "mask": "5.1.05.01.00004"},
    {"code": "5444", "name": "GANHO PROVENIENTE DE COMPRA VANTAJOSA", "grupo": "Resultado", "mask": "5.1.05.01.00005"},
    {"code": "5445", "name": "AMORTIZAÇÃO DE ÁGIO - CONTÁBIL", "grupo": "Resultado", "mask": "5.1.05.01.00006"},
    {"code": "5446", "name": "AMORTIZAÇÃO DE ÁGIO - FISCAL", "grupo": "Resultado", "mask": "5.1.05.01.00007"},
    {"code": "5447", "name": "RESULTADO POSITIVO EM SCP", "grupo": "Resultado", "mask": "5.1.05.01.00008"},
    {"code": "5448", "name": "RESULTADO POSIT.PARTICIPAÇÕES EXTERIOR", "grupo": "Resultado", "mask": "5.1.05.01.00009"},
    {"code": "5449", "name": "RESULTADO POSITIVO VAR. PART. MEP", "grupo": "Resultado", "mask": "5.1.05.01.00010"},
    {"code": "5455", "name": "PREJUÍZO NA ALIENAÇÃO DE PARTICIPAÇÕES SOCIETÁRIAS", "grupo": "Resultado", "mask": "5.1.05.02.00001"},
    {"code": "5456", "name": "RESULTADO NEGATIVO DE EQUIVALÊNCIA PATRIMONIAL", "grupo": "Resultado", "mask": "5.1.05.02.00002"},
    {"code": "5457", "name": "RESULTADO NEGAT.PELO MEP - PROPORÇÃO PL", "grupo": "Resultado", "mask": "5.1.05.02.00003"},
    {"code": "5458", "name": "RESULTADO NEGAT.PELO MEP - DESPROP.PL", "grupo": "Resultado", "mask": "5.1.05.02.00004"},
    {"code": "5459", "name": "AMORTIZAÇÃO DE DESÁGIO", "grupo": "Resultado", "mask": "5.1.05.02.00005"},
    {"code": "5460", "name": "RESULTADO NEGATIVO EM SCP", "grupo": "Resultado", "mask": "5.1.05.02.00006"},
    {"code": "5461", "name": "RESULTADO NEGAT.PARTICIPAÇÕES EXTERIOR", "grupo": "Resultado", "mask": "5.1.05.02.00007"},
    {"code": "5462", "name": "RESULTADO NEGATIVO VAR. PART. MEP", "grupo": "Resultado", "mask": "5.1.05.02.00008"},
    {"code": "5466", "name": "SINISTROS COM IMOBILIZADO", "grupo": "Resultado", "mask": "5.1.06.01.00001"},
    {"code": "5467", "name": "OUTRAS RECEITAS NÃO OPERACIONAIS", "grupo": "Resultado", "mask": "5.1.06.01.00002"},
    {"code": "5468", "name": "CONSTITUIÇÃO DE PROVISÕES", "grupo": "Resultado", "mask": "5.1.06.01.00003"},
    {"code": "5469", "name": "VALOR DE ALIENAÇÃO DOS INVESTIMENTOS", "grupo": "Resultado", "mask": "5.1.06.01.00004"},
    {"code": "5470", "name": "INDENIZAÇÃO DE SEGUROS", "grupo": "Resultado", "mask": "5.1.06.01.00005"},
    {"code": "5471", "name": "RECUPERAÇÃO DE DESPESAS", "grupo": "Resultado", "mask": "5.1.06.01.00006"},
    {"code": "5472", "name": "RECUPERAÇÃO DO ATIVO IMOBILIZADO", "grupo": "Resultado", "mask": "5.1.06.01.00007"},
    {"code": "5473", "name": "RENDAS DIVERSAS", "grupo": "Resultado", "mask": "5.1.06.01.00008"},
    {"code": "5474", "name": "REVERSÃO DE PROVISÕES", "grupo": "Resultado", "mask": "5.1.06.01.00009"},
    {"code": "5478", "name": "OUTRAS DESPESAS COM PRESTAÇÕES SERVIÇOS", "grupo": "Resultado", "mask": "5.1.06.02.00001"},
    {"code": "5479", "name": "MATERIAIS DE CONSUMO DIVERSOS", "grupo": "Resultado", "mask": "5.1.06.02.00002"},
    {"code": "5480", "name": "DESPESAS COM CONTINGÊNCIAS", "grupo": "Resultado", "mask": "5.1.06.02.00003"},
    {"code": "5481", "name": "PERDAS EXTRAORDINÁRIAS", "grupo": "Resultado", "mask": "5.1.06.02.00004"},
    {"code": "5482", "name": "DOAÇÕES E CONTRIBUIÇÕES", "grupo": "Resultado", "mask": "5.1.06.02.00005"},
    {"code": "5483", "name": "OUTRAS DESPESAS NÃO OPERACIONAIS", "grupo": "Resultado", "mask": "5.1.06.02.00006"},
    {"code": "5484", "name": "OUTRAS PROVISÕES", "grupo": "Resultado", "mask": "5.1.06.02.00007"},
    {"code": "5485", "name": "REVERSÃO - OUTRAS PROVISÕES", "grupo": "Resultado", "mask": "5.1.06.02.00008"},
    {"code": "5503", "name": "SERVIÇOS PRESTADOS POR TERCEIROS PESSOA JURIDICA", "grupo": "Resultado", "mask": "5.1.07.01.00001"},
    {"code": "5504", "name": "IMPOSTOS E TAXAS - INDEDUTÍVEIS", "grupo": "Resultado", "mask": "5.1.07.01.00002"},
    {"code": "5505", "name": "DESPESAS INDEDUTÍVEIS", "grupo": "Resultado", "mask": "5.1.07.01.00003"},
    {"code": "5506", "name": "MULTAS INDEDUTÍVEIS", "grupo": "Resultado", "mask": "5.1.07.01.00004"},
    {"code": "5507", "name": "BRINDES E BONIFICAÇÕES", "grupo": "Resultado", "mask": "5.1.07.01.00005"},
    {"code": "5508", "name": "DIFERENÇAS TEMPORÁRIAS - ARRENDAMENTOS FINANCEIROS", "grupo": "Resultado", "mask": "5.1.07.01.00006"},
    {"code": "5509", "name": "CUSTO DO BEM BAIXADO - INDEDUTÍVEL", "grupo": "Resultado", "mask": "5.1.07.01.00007"},
    {"code": "5510", "name": "MATERIAIS DE USO E CONSUMO - INDEDUTÍVEL", "grupo": "Resultado", "mask": "5.1.07.01.00008"},
    {"code": "5511", "name": "DOAÇÕES E CONTRIBUIÇÕES", "grupo": "Resultado", "mask": "5.1.07.01.00009"},
    {"code": "5512", "name": "OUTRAS DESPESAS NÃO OPERACIONAIS - INDEDUTÍVEL", "grupo": "Resultado", "mask": "5.1.07.01.00010"},
    {"code": "5513", "name": "SINISTROS E FURTOS", "grupo": "Resultado", "mask": "5.1.07.01.00011"},
    {"code": "5514", "name": "PAGAMENTOS SEM IDENTIFICAÇÃO", "grupo": "Resultado", "mask": "5.1.07.01.00012"},
    {"code": "5515", "name": "PROVISÕES", "grupo": "Resultado", "mask": "5.1.07.01.00013"},
    {"code": "5516", "name": "REVERSÃO DE PROVISÕES", "grupo": "Resultado", "mask": "5.1.07.01.00014"},
    {"code": "5517", "name": "DESPESAS PRÉ-OPERACIONAIS", "grupo": "Resultado", "mask": "5.1.07.01.00015"},
    {"code": "5518", "name": "DESPESAS FINANCEIRAS - PARCELA NÃO DEDUTÍVEL", "grupo": "Resultado", "mask": "5.1.07.01.00016"},
    {"code": "5519", "name": "PERDAS POR CRÉDITO DE LIQUIDAÇÃO DUVIDOSA - PCLD", "grupo": "Resultado", "mask": "5.1.07.01.00017"},
    {"code": "5520", "name": "REVERSÃO DE PERDAS POR CRÉDITO DE LIQUIDAÇÃO DUVIDOSA - PCLD", "grupo": "Resultado", "mask": "5.1.07.01.00018"},
    {"code": "5532", "name": "RESULTADO DE EXERCÍCIOS ANTERIORES DB", "grupo": "Resultado", "mask": "5.1.08.01.00001"},
    {"code": "5533", "name": "RESULTADO DE EXERCÍCIOS ANTERIORES CR", "grupo": "Resultado", "mask": "5.1.08.01.00002"},
    {"code": "5982", "name": "PROVISÃO DE CSLL", "grupo": "Resultado", "mask": "5.2.01.01.00001"},
    {"code": "5983", "name": "PROVISÃO DE IRPJ", "grupo": "Resultado", "mask": "5.2.01.01.00002"},
    {"code": "5992", "name": "PROVISÃO DE CSLL DIFERIDO", "grupo": "Resultado", "mask": "5.2.02.01.00001"},
    {"code": "5993", "name": "PROVISÃO DE IRPJ DIFERIDO", "grupo": "Resultado", "mask": "5.2.02.01.00002"},
]

# Lookup por código para auto-sugestão
_BHUB_BY_CODE: dict[str, dict] = {a["code"]: a for a in BHUB_ACCOUNTS if a.get("code")}


def _norm(s: str) -> str:
    """Normaliza string: maiúsculas, sem acentos, sem caracteres especiais."""
    s = unicodedata.normalize("NFD", s.upper())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    return " ".join(s.split())


# ── Sinônimos contábeis ───────────────────────────────────────────────────────
_SYNONYMS: dict[str, set[str]] = {
    "CAIXA":        {"NUMERARIO", "NUMERARIOS", "DINHEIRO", "ESPECIE", "FUNDO FIXO"},
    "BANCO":        {"BANCOS", "BANCARIO", "CONTA CORRENTE", "CONTA BANCARIA"},
    "APLICACAO":    {"APLICACOES", "INVESTIMENTO", "CDB", "LCI", "LCA", "POUPANCA", "RENDA FIXA"},
    "CLIENTE":      {"CLIENTES", "DUPLICATA", "DUPLICATAS", "RECEBIVEIS", "RECEBER", "FATURAMENTO"},
    "FORNECEDOR":   {"FORNECEDORES", "PAGAR", "CONTAS A PAGAR"},
    "EMPRESTIMO":   {"EMPRESTIMOS", "MUTUO", "FINANCIAMENTO", "FINANCIAMENTOS"},
    "SALARIO":      {"SALARIOS", "REMUNERACAO", "FOLHA", "PROVENTOS", "ORDENADO"},
    "ADIANTAMENTO": {"ADIANTAMENTOS", "ANTECIPACAO", "ADTO"},
    "SOCIO":        {"SOCIOS", "ACIONISTA", "QUOTISTA", "PARTICIPACAO"},
    "FUNCIONARIO":  {"FUNCIONARIOS", "EMPREGADO", "EMPREGADOS", "COLABORADOR"},
    "RECUPERAR":    {"RECUPERACAO", "COMPENSAR", "COMPENSACAO", "A RECUPERAR", "A COMPENSAR"},
    "IMPOSTO":      {"IMPOSTOS", "TRIBUTO", "TRIBUTOS", "TAXA", "TAXAS"},
    "DESPESA":      {"DESPESAS", "CUSTO", "CUSTOS", "GASTO", "GASTOS"},
    "RECEITA":      {"RECEITAS", "FATURAMENTO", "VENDA", "VENDAS"},
    "ICMS":         {"ICMS A RECUPERAR", "ICMS A COMPENSAR"},
    "PIS":          {"PIS PASEP", "PIS A RECUPERAR"},
    "COFINS":       {"COFINS A RECUPERAR"},
    "IRPJ":         {"IMPOSTO DE RENDA", "IR", "IRPJ A RECUPERAR"},
    "CSLL":         {"CONTRIBUICAO SOCIAL", "CSLL A RECUPERAR"},
    "INSS":         {"PREVIDENCIA", "PREVIDENCIA SOCIAL"},
    "ISS":          {"ISSQN", "ISS A RECUPERAR"},
    "IRRF":         {"IR RETIDO", "IMPOSTO RETIDO NA FONTE", "RETENCAO"},
    "CARTAO":       {"CARTOES", "CARTAO DE CREDITO", "CARTAO DE DEBITO", "ADMINISTRADORA"},
    "RECOLHER":     {"A RECOLHER", "A PAGAR", "RECOLHIMENTO", "RECOLHER"},
    "FORNECEDOR":   {"FORNECEDORES", "PAGAR", "CONTAS A PAGAR"},
    "CAPITAL":      {"CAPITAL SOCIAL", "PATRIMONIO", "QUOTA", "QUOTAS"},
    "LUCRO":        {"LUCROS", "RESULTADO", "SUPERAVIT", "LUCROS ACUMULADOS"},
    "PREJUIZO":     {"PREJUIZOS", "DEFICIT", "PERDA", "PREJUIZOS ACUMULADOS"},
    "RECEITA":      {"RECEITAS", "FATURAMENTO", "VENDA", "VENDAS", "PRESTACAO DE SERVICO"},
    "DESPESA":      {"DESPESAS", "CUSTO", "CUSTOS", "GASTO", "GASTOS", "ENCARGO"},
    "FOLHA":        {"FOLHA DE PAGAMENTO", "PESSOAL", "COLABORADORES"},
    "PRO LABORE":   {"PROLABORE", "HONORARIO", "HONORARIOS", "REMUNERACAO DE SOCIOS", "REMUNERACAO SOCIOS"},
    "CUSTO":        {"CUSTOS", "CMV", "CPV", "CSV", "CUSTO DOS PRODUTOS", "CUSTO DOS SERVICOS"},
    "LEASING":      {"ARRENDAMENTO", "ARRENDAMENTO MERCANTIL", "FINANCIAMENTO LEASING"},
    "DEPRECIACAO":  {"AMORTIZACAO", "DEPRECIACAO ACUMULADA", "DEPRECIACOES", "AMORTIZACOES", "DEPRECIACAO E AMORTIZACAO"},
    "PROVISAO":     {"PROVISOES", "ESTIMATIVA", "RESERVA"},
    "IMPOSTO":      {"IMPOSTOS", "TRIBUTO", "TRIBUTOS", "TAXA", "TAXAS", "IOF", "ISS", "ICMS", "IPI", "IRPJ", "CSLL", "CONTRIBUICAO", "CONTRIBUICOES"},
    "SERVICO":      {"SERVICOS", "PRESTACAO", "PRESTACOES", "CONSULTORIA", "AUDITORIA", "TERCEIROS"},
    "SALARIO":      {"SALARIOS", "ORDENADOS", "VENCIMENTOS", "REMUNERACAO"},
    "SEGURO":       {"SEGUROS", "APOLICE"},
    "ALUGUEL":      {"ALUGUEIS", "LOCACAO", "LOCACOES", "ARRENDAMENTO"},
    "MULTA":        {"MULTAS", "PENALIDADE", "PENALIDADES", "SANCAO"},
    "JUROS":        {"ENCARGOS", "FINANCEIROS", "BANCARIOS", "MORA"},
    # ── Grupos específicos para match preciso de despesas ─────────────────
    "ASSESSORIA":   {"ASSESSORIAS", "CONTABIL", "CONTABEIS", "CONTABILIDADE",
                     "HONORARIO CONTABIL", "HONORARIOS CONTABEIS", "SERVICOS CONTABEIS",
                     "ASSISTENCIA CONTABIL", "SERVICO CONTABIL", "ESCRITORIO CONTABIL"},
    "ADVOCAT":      {"ADVOCACIA", "ADVOCATICIOS", "ADVOCATICO", "ADVOCATICA",
                     "JURIDICO", "JURIDICA", "JURIDICOS", "JURIDICAS",
                     "ADVOGADO", "ADVOGADOS", "JURIDICAMENTE", "CURADORIA"},
    "COMBUSTIV":    {"COMBUSTIVEL", "COMBUSTIVEIS", "LUBRIFICANTE", "LUBRIFICANTES",
                     "GASOLINA", "DIESEL", "ETANOL", "ABASTECIMENTO", "COMBUSTAO"},
    "MANUTENCAO":   {"MANUTENCOES", "CONSERVACAO", "CONSERVACOES", "REPARO", "REPAROS",
                     "REFORMA", "REFORMAS", "CONSERTO", "CONSERTOS", "PREDIAL", "CORRETIVA"},
    "ESCRITORIO":   {"PAPELARIA", "PAPELARIAS", "SUPRIMENTO", "SUPRIMENTOS",
                     "MATERIAL ESCRITORIO", "MATERIAL CONSUMO", "EXPEDIENTE"},
    "LIMPEZA":      {"HIGIENE", "CONSERVACAO", "HIGIENIZACAO", "PRODUTO LIMPEZA",
                     "MATERIAL LIMPEZA", "DEDETIZACAO"},
    "VIGILANCIA":   {"SEGURANCA", "SEGURANCAS", "MONITORAMENTO", "PORTARIA", "RONDA"},
    "VIAGEM":       {"VIAGENS", "HOSPEDAGEM", "HOSPEDAGENS", "HOTEL", "HOTEIS",
                     "PASSAGEM", "PASSAGENS", "DESLOCAMENTO", "DIARIA", "DIARIAS",
                     "PASSAGEM AEREA", "TRANSLADO"},
    "FRETE":        {"FRETES", "CARRETO", "CARRETOS", "ENTREGA", "ENTREGAS",
                     "LOGISTICA", "TRANSPORTE CARGA", "REMESSA"},
    "INFORMATICA":  {"COMPUTADOR", "COMPUTADORES", "SISTEMA", "SISTEMAS", "SOFTWARE",
                     "SOFTWARES", "TI", "TECNOLOGIA", "INFORMATICO", "HARDWARE",
                     "LICENCA", "LICENCAS", "NUVEM", "CLOUD"},
    "CAPACITACAO":  {"TREINAMENTO", "TREINAMENTOS", "CURSO", "CURSOS", "PALESTRA",
                     "PALESTRAS", "QUALIFICACAO", "DESENVOLVIMENTO", "SEMINARIO",
                     "WORKSHOP", "FORMACAO"},
    "VALE":         {"VALES", "VT", "VR", "VA", "VALE TRANSPORTE", "VALE REFEICAO",
                     "VALE ALIMENTACAO", "TICKET"},
    "REPRESENTACAO": {"REPRESENTACOES", "HOSPITALIDADE", "ENTRETENIMENTO",
                      "RELACOES PUBLICAS", "RECEPCAO", "EVENTOS"},
}

# Reverse lookup: palavra normalizada → chave canônica do grupo
_SYN_REV: dict[str, str] = {}
for _key, _vals in _SYNONYMS.items():
    _SYN_REV[_key] = _key
    for _v in _vals:
        _SYN_REV[_norm(_v)] = _key


def _expand(words: set[str]) -> set[str]:
    """Expande conjunto de palavras com sinônimos do grupo."""
    expanded = set(words)
    for w in list(words):
        key = _SYN_REV.get(w)
        if key:
            expanded.add(key)
            expanded.update(_norm(v) for v in _SYNONYMS.get(key, set()))
    return expanded


_STOP_WORDS = frozenset({
    "E", "DE", "DA", "DO", "DAS", "DOS", "A", "O", "AS", "OS",
    "EM", "POR", "COM", "PARA", "NO", "NA", "NOS", "NAS", "SE",
})

def _word_overlap(a: str, b: str) -> float:
    """Jaccard sobre prefixos de palavras (7 chars) com expansão de sinônimos.
    Prefixo de 7 chars resolve plurais portugueses: DEPRECIACOES ≡ DEPRECIACAO."""
    def _stem_set(text: str) -> set[str]:
        expanded = _expand(set(text.split()))
        result: set[str] = set()
        for w in expanded:
            if " " in w or w in _STOP_WORDS:
                continue
            result.add(w)
            if len(w) > 7:
                result.add(w[:7])
        return result

    wa = _stem_set(a)
    wb = _stem_set(b)
    if not wa or not wb:
        return 0.0
    inter = wa & wb
    union = wa | wb
    return len(inter) / len(union) if union else 0.0


def _norm_seg(s: str) -> str:
    """Normaliza segmento de classificação removendo zeros à esquerda."""
    try:
        return str(int(s))
    except ValueError:
        return s


# Equivalência de grupo: BHub usa 3.x (Receitas), 4.x (Custo), 5.x (Despesas Gerais)
# para Resultado. Planos antigos tipicamente usam 3.x ou 4.x. Todos mapeados para "R".
_GRUPO_EQUIV: dict[str, str] = {"1": "1", "2": "2", "3": "R", "4": "R", "5": "R", "6": "R", "9": "R"}

def _structural_score(old_cls: str, bhub_mask: str) -> float:
    """
    Compara hierarquia de classificação segmento a segmento.
    Mapeia grupos 3/4/5 (Resultado) como equivalentes em ambos os lados:
    plano antigo 4.x (Despesas) ≡ BHub 4.x (Custo) ≡ BHub 5.x (Despesas Gerais).
    """
    if not old_cls or not bhub_mask:
        return 0.0
    op = [_norm_seg(p) for p in old_cls.split(".")]
    bp = [_norm_seg(p) for p in bhub_mask.split(".")]
    old_grp = _GRUPO_EQUIV.get(op[0], op[0])
    bhub_grp = _GRUPO_EQUIV.get(bp[0], bp[0])
    if old_grp != bhub_grp:
        return 0.0
    # Para Resultado (R), hierarquias divergem entre planos.
    # Bônus extra se o primeiro dígito coincide (ex.: old 5.x → BHub 5.x Despesas).
    if old_grp == "R":
        return 0.5 if op[0] == bp[0] else 0.3
    # Para Ativo e Passivo, comparar segmentos intermediários
    depth = min(len(op) - 1, len(bp) - 1, 3)
    if depth < 1:
        return 0.3
    matches = sum(1 for i in range(1, depth + 1) if i < len(op) and i < len(bp) and op[i] == bp[i])
    return matches / depth


def suggest_bhub(name: str, grupo: str, old_cls: str = "", n: int = 3) -> list[dict]:
    """
    Retorna até n sugestões BHub ordenadas por score combinado:
      50% similaridade de sequência de texto
      25% sobreposição de palavras com sinônimos contábeis
      25% alinhamento da estrutura de classificação
    Nunca retorna lista vazia — usa fallback para grupo mais próximo se necessário.
    """
    norm_name  = _norm(name)
    candidates = [a for a in BHUB_ACCOUNTS if a["grupo"] == grupo]

    # Fallback: se o grupo exato não tem candidatos, usa todos os accounts
    if not candidates:
        candidates = list(BHUB_ACCOUNTS)

    # Remove prefixos genéricos do início para SequenceMatcher não favorecer
    # acidentalmente contas com a mesma palavra genérica (ex.: "DESPESAS JURÍDICAS"
    # vs "DESPESAS FINANCEIRAS" — o prefixo "DESPESAS" é ruído para o match).
    _GENERIC_PREFIX = frozenset({
        "DESPESAS", "DESPESA", "RECEITAS", "RECEITA", "OUTROS", "OUTRAS",
        "DEMAIS", "CUSTO", "CUSTOS", "ENCARGOS", "SERVICOS", "SERVICO",
        "PROVISAO", "PROVISOES",
    })
    def _strip_prefix(s: str) -> str:
        words = s.split()
        while len(words) > 1 and words[0] in _GENERIC_PREFIX:
            words = words[1:]
        return " ".join(words)

    norm_stripped = _strip_prefix(norm_name)

    def _score(acc: dict) -> float:
        acc_norm     = _norm(acc["name"])
        acc_stripped = _strip_prefix(acc_norm)
        desc_sim   = difflib.SequenceMatcher(None, norm_stripped, acc_stripped).ratio()
        word_sim   = _word_overlap(norm_name, acc_norm)
        struct_sim = _structural_score(old_cls, acc["mask"])
        return 0.40 * desc_sim + 0.35 * word_sim + 0.25 * struct_sim

    scored = [{**acc, "score": round(_score(acc) * 100, 1)} for acc in candidates]
    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:n]


# ─────────────────────────────────────────────────────────────────────────────
# DETECÇÃO DE ESTRUTURA BHUB
# ─────────────────────────────────────────────────────────────────────────────

def _is_bhub_structure(df: pd.DataFrame, threshold: float = 0.35) -> tuple[bool, str]:
    """
    Verifica se o balancete já usa a estrutura de contas BHub.
    Retorna (is_bhub, motivo).

    Lógica: exige COMBINAÇÃO de critérios — um critério isolado não é suficiente.
      1. Prioridade: código BHub + (classificação OU descrição) bater no mesmo registro
      2. Fallback: classificação no formato BHub + descrição igual ao nome BHub
    Threshold padrão: 35% das contas precisam satisfazer a combinação.
    """
    import re as _re

    _bhub_mask_re = _re.compile(r'^\d\.\d{1,2}\.\d{2}\.\d{2}\.\d{5}$')

    # Índice por código: code -> lista de (mask, norm_name)
    _bhub_by_code: dict[str, list[tuple[str, str]]] = {}
    for a in BHUB_ACCOUNTS:
        c = str(a.get("code", "")).strip()
        if c:
            _bhub_by_code.setdefault(c, []).append(
                (a["mask"].strip(), _norm(a["name"]))
            )
    _bhub_names = {_norm(a["name"]) for a in BHUB_ACCOUNTS}

    total = len(df)
    if total == 0:
        return False, ""

    # ── Critério 1: código + (classificação OU descrição) no mesmo registro BHub ──
    triple_hits = 0
    mask_name_hits = 0  # classificação formato BHub + descrição BHub (sem código)

    for _, row in df.iterrows():
        code  = str(row.get("code", "")).strip()
        mask  = str(row.get("classification", "")).strip()
        name_n = _norm(str(row.get("description", "")))

        mask_is_bhub = bool(_bhub_mask_re.match(mask))
        name_is_bhub = name_n in _bhub_names

        if code in _bhub_by_code:
            for bhub_mask, bhub_name in _bhub_by_code[code]:
                if mask == bhub_mask or name_n == bhub_name:
                    triple_hits += 1
                    break

        if mask_is_bhub and name_is_bhub:
            mask_name_hits += 1

    triple_ratio    = triple_hits    / total
    mask_name_ratio = mask_name_hits / total

    if triple_ratio >= threshold:
        return True, (
            f"{triple_hits} de {total} contas ({triple_ratio:.0%}) têm código BHub "
            f"combinado com classificação ou descrição coincidentes"
        )
    if mask_name_ratio >= threshold:
        return True, (
            f"{mask_name_hits} de {total} contas ({mask_name_ratio:.0%}) têm "
            f"classificação no formato BHub e descrição idêntica ao plano"
        )
    return False, ""


# ─────────────────────────────────────────────────────────────────────────────
# SUGESTÃO VIA IA (Claude / Gemini)
# ─────────────────────────────────────────────────────────────────────────────

def _build_ai_prompt(bhub_catalog: str, accs_text: str) -> str:
    return f"""Você é um contador especialista em migração de planos de contas para o padrão BHub Contabilidade (ECD/SPED).

PLANO DE CONTAS BHUB (formato: ID|NOME|MÁSCARA|GRUPO):
{bhub_catalog}

CONTAS DO BALANCETE ANTERIOR A MAPEAR (formato: POS|CÓD_ANTIGO|DESCRIÇÃO|CLASSIF_ANTIGA|GRUPO):
{accs_text}

Para cada conta do balancete anterior, encontre o ID do plano BHub mais adequado.

REGRAS OBRIGATÓRIAS:
1. SEMPRE retorne um bhub_id para cada conta — NUNCA retorne null
2. Use SOMENTE os IDs listados no plano BHub acima — nunca invente contas fora da lista
3. O critério principal é o NOME/DESCRIÇÃO da conta — encontre a mais parecida semanticamente
4. Respeite obrigatoriamente o grupo: conta de Ativo → ID de Ativo; Passivo/PL → ID de Passivo/PL; Resultado → ID de Resultado
5. Use a classificação antiga para confirmar o grupo: 1.x = Ativo, 2.x = Passivo/PL, 3.x ou 4.x = Resultado
6. Para CLIENTES sem match exato → use a conta genérica de clientes do plano BHub
7. Para FORNECEDORES sem match exato → use a conta genérica de fornecedores do plano BHub
8. Para qualquer conta sem equivalente exato → sugira a conta de mesmo grupo com nome mais próximo
9. Em último caso (sem nada parecido no grupo correto) → escolha a conta mais genérica/totalizadora do grupo

Retorne APENAS o JSON a seguir, sem nenhum texto antes ou depois:
[{{"pos":0,"bhub_id":5}},{{"pos":1,"bhub_id":12}},...]"""


def _pip_install(package: str) -> None:
    """Instala pacote no mesmo Python que está rodando o Streamlit."""
    import subprocess, sys
    subprocess.run(
        [sys.executable, "-m", "pip", "install", package,
         "--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org",
         "--quiet"],
        check=True,
    )


def _call_openai(prompt: str, api_key: str, model: str = "gpt-4o-mini") -> str:
    try:
        import openai as _openai
    except ImportError:
        _pip_install("openai>=1.0.0")
        import openai as _openai
    client = _openai.OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=4096,
    )
    return response.choices[0].message.content.strip()


def _call_anthropic(prompt: str, api_key: str) -> str:
    try:
        import anthropic as _anthropic
    except ImportError:
        _pip_install("anthropic>=0.30.0")
        import anthropic as _anthropic
    client = _anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def _list_gemini_models(api_key: str) -> list[str]:
    """Retorna lista de modelos Gemini reais disponíveis para a chave informada."""
    try:
        from google import genai as _genai
    except ImportError:
        _pip_install("google-genai>=1.0.0")
        from google import genai as _genai
    client = _genai.Client(api_key=api_key)
    names = []
    for m in client.models.list():
        raw   = getattr(m, "name", "") or ""
        short = raw.replace("models/", "").strip()
        actions = getattr(m, "supported_actions", []) or []
        # Aceita apenas modelos Gemini reais (exclui modelos internos como "antigravity")
        if "generateContent" in actions and short.startswith("gemini-"):
            names.append(short)
    # Ordena preferindo flash → pro
    def _rank(n: str) -> tuple:
        return ("flash" not in n, "lite" not in n, n)
    return sorted(names, key=_rank)


def _call_gemini(prompt: str, api_key: str, model: str = "") -> str:
    try:
        from google import genai as _genai
    except ImportError:
        _pip_install("google-genai>=1.0.0")
        from google import genai as _genai
    import time as _time, re as _re

    client = _genai.Client(api_key=api_key)

    # Se modelo não especificado, tenta a lista de preferência em ordem
    candidates = (
        [model] if model
        else ["gemini-2.0-flash-lite", "gemini-1.5-flash", "gemini-1.5-flash-8b",
              "gemini-2.0-flash", "gemini-1.5-pro"]
    )

    last_err = None
    for candidate in candidates:
        for attempt in range(3):
            try:
                response = client.models.generate_content(model=candidate, contents=prompt)
                return response.text.strip()
            except Exception as _e:
                msg = str(_e)
                is_quota    = "429" in msg or "RESOURCE_EXHAUSTED" in msg
                is_zero_lim = "limit: 0" in msg          # quota zero → não adianta retry
                is_not_found = "404" in msg or "NOT_FOUND" in msg

                if is_not_found or (is_quota and is_zero_lim):
                    last_err = _e
                    break   # pula para o próximo candidato

                if is_quota and not is_zero_lim:
                    # Quota temporária → aguarda e tenta de novo
                    _delay_m = _re.search(r'retry[^\d]*(\d+)', msg, _re.IGNORECASE)
                    _delay   = int(_delay_m.group(1)) if _delay_m else 30
                    if attempt < 2:
                        _time.sleep(_delay + 2)
                        continue
                    last_err = _e
                    break

                raise   # outro erro → propaga imediatamente

    raise last_err or RuntimeError("Nenhum modelo Gemini disponível para esta chave API")


def suggest_bhub_ai(df_accounts: pd.DataFrame, api_key: str,
                    provider: str = "gemini", model: str = "") -> dict:
    """
    Usa IA (Gemini ou Claude) para sugerir contas BHub.
    Retorna {row_position: {"bhub_code": ..., "bhub_name": ..., "bhub_mask": ...}}.
    """
    import json as _json, re as _re

    # Catálogo BHub com IDs sequenciais
    bhub_by_id: dict[int, dict] = {}
    catalog_lines: list[str] = []
    for i, acc in enumerate(BHUB_ACCOUNTS, 1):
        bhub_by_id[i] = acc
        catalog_lines.append(f"{i}|{acc['name']}|{acc['mask']}|{acc['grupo']}")

    # Contas a mapear
    accs_lines: list[str] = []
    for pos, (_, row) in enumerate(df_accounts.iterrows()):
        accs_lines.append(
            f"{pos}|{row.get('code', '')}|{row.get('description', '')}|"
            f"{row.get('classification', '')}|{row.get('grupo', '')}"
        )

    prompt = _build_ai_prompt("\n".join(catalog_lines), "\n".join(accs_lines))

    if provider == "gemini":
        raw = _call_gemini(prompt, api_key, model=model or "")
    elif provider == "openai":
        raw = _call_openai(prompt, api_key, model=model or "gpt-4o-mini")
    else:
        raw = _call_anthropic(prompt, api_key)

    m = _re.search(r"\[[\s\S]*\]", raw)
    if not m:
        raise ValueError(f"IA não retornou JSON válido. Resposta recebida:\n{raw[:500]}")

    result: dict[int, dict] = {}
    for s in _json.loads(m.group()):
        pos = s.get("pos")
        bid = s.get("bhub_id")
        if pos is None:
            continue
        if bid and bid in bhub_by_id:
            acc = bhub_by_id[bid]
            result[int(pos)] = {
                "bhub_code": acc.get("code", ""),
                "bhub_name": acc["name"],
                "bhub_mask": acc["mask"],
            }
        else:
            result[int(pos)] = {"bhub_code": "", "bhub_name": "", "bhub_mask": ""}

    return result


# ─────────────────────────────────────────────────────────────────────────────
# PARSING DE PDF
# ─────────────────────────────────────────────────────────────────────────────

_SKIP = {
    # Cabeçalhos e rodapés do PDF (não têm 3 valores monetários, mas pulamos por segurança)
    "BALANCETE", "Período:", "Emissão:", "Hora:", "Folha:", "Número livro:",
    "Sistema licenciado", "Código Classificação", "Saldo Anterior",
    "RESUMO DO BALANCETE", "CONTAS DE APURAÇÃO", "RESULTADO DO MES",
    "RESULTADO DO EXERCÍCIO", "CONTAS DEVEDORAS", "CONTAS CREDORAS",
    "C.N.P.J.:", "Empresa:", "CPF:", "Reg. no CRC", "Assinado de forma",
    "SÓCIO-ADMINISTRADOR", "CONTAS DE RESULTADO",
    # "ADMINISTRADOR" removido → matchava "ADMINISTRADORA" em nomes de contas
    # "RESUMO" removido → muito curto, pode aparecer em descrições legítimas
    # Nomes de pessoas removidos → linhas de assinatura não têm 3 valores monetários
}

_AMT_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}[DC]?")
_CODE_RE = re.compile(r"^\d+$")
_CLS_RE = re.compile(r"^\d+(?:\.\d+)+$")


def _parse_lines(text: str) -> list[dict]:
    records = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if any(kw in line for kw in _SKIP):
            continue

        amounts = _AMT_RE.findall(line)
        if len(amounts) < 3:
            continue

        rest = _AMT_RE.sub("", line).strip()
        rest = " ".join(rest.split())
        if not rest:
            continue

        tokens = rest.split()
        code = cls = None
        i = 0

        # Pular tokens de texto inicial (cabeçalhos de seção como "COLIGADA I"
        # que o pdfplumber une na mesma linha do registro)
        while i < len(tokens) and not _CODE_RE.match(tokens[i]) and not _CLS_RE.match(tokens[i]):
            i += 1

        # Extrair código numérico e classificação
        if i < len(tokens) and _CODE_RE.match(tokens[i]):
            code = tokens[i]; i += 1
        if i < len(tokens) and _CLS_RE.match(tokens[i]):
            cls = tokens[i]; i += 1
        elif i < len(tokens) and _CODE_RE.match(tokens[i]) and not code:
            code = tokens[i]; i += 1

        description = " ".join(tokens[i:])
        if not description:
            continue

        # Últimos 3 valores: débito, crédito, saldo atual
        parsed = []
        for a in amounts:
            v, ind = parse_br_value(a)
            parsed.append((v, ind))

        cur_val, cur_ind = parsed[-1]
        cre_val = parsed[-2][0] if len(parsed) >= 2 else 0.0
        deb_val = parsed[-3][0] if len(parsed) >= 3 else 0.0

        if cur_ind is None:
            cur_ind = "D" if deb_val >= cre_val else "C"

        records.append(
            {
                "code": code or "",
                "classification": cls or "",
                "description": description,
                "debit_mv": deb_val,
                "credit_mv": cre_val,
                "current_value": cur_val,
                "current_indicator": cur_ind,
            }
        )
    return records


def extract_pdf(uploaded_file) -> list[dict]:
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    return _parse_lines(text)


# ─────────────────────────────────────────────────────────────────────────────
# GERAÇÃO DA PLANILHA DOMÍNIO (XLS-compatível via openpyxl)
# ─────────────────────────────────────────────────────────────────────────────

def _hdr(ws, row, cols, bg="1F4E79", fg="FFFFFF"):
    fill = PatternFill("solid", fgColor=bg)
    font = Font(color=fg, bold=True, size=10)
    for c, h in enumerate(cols, 1):
        cell = ws.cell(row, c, h)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def generate_dominio_excel(
    df: pd.DataFrame,
    cod_empresa: str,
    opening_date: date,
    contra_code: str = "",
    modo: str = "dobradas",
) -> io.BytesIO:
    """
    Gera Excel no formato da Planilha Modelo Domínio.
    modo='dobradas': 1 lançamento único em partidas dobradas (padrão).
    modo='simples':  1 lançamento por conta com conta de contrapartida.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Lancamentos"

    modo_label = (
        "Partidas dobradas (lançamento único)"
        if modo == "dobradas"
        else f"Simples | Contrapartida: {contra_code}"
    )
    ws["A1"] = (
        f"IMPLANTAÇÃO DE SALDOS | Empresa: {cod_empresa} | "
        f"Data: {opening_date.strftime('%d/%m/%Y')} | {modo_label}"
    )
    ws["A1"].font = Font(bold=True, color="1F4E79", size=11)
    ws.merge_cells("A1:J1")

    headers = [
        "Data", "Cód. Conta Debito", "Cód. Conta Credito", "Valor",
        "Cód. Histórico", "Complemento Histórico", "Inicia Lote",
        "Código Matriz/Filial", "Centro de Custo Débito", "Centro de Custo Crédito",
    ]
    _hdr(ws, 2, headers)

    col_w = [14, 18, 18, 16, 14, 55, 12, 20, 18, 18]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A3"

    alt  = PatternFill("solid", fgColor="EEF2F7")
    data_str = opening_date.strftime("%d/%m/%Y")

    row   = 3
    first = True   # controla INICIA_LOTE no modo dobradas

    for _, acc in df.iterrows():
        val  = acc.get("current_value", 0)
        ind  = acc.get("current_indicator", "D")
        code = str(acc.get("code", "")).strip()
        desc = str(acc.get("description", "")).strip()[:55]
        if val == 0 or not code:
            continue

        cents = int(round(val * 100))
        hist  = f"IMPLANTAÇÃO DE SALDO EM {data_str} - {desc}"
        fill  = alt if row % 2 == 0 else None

        def wr(r, c, v, _fill=fill):
            cell = ws.cell(r, c, v)
            if _fill:
                cell.fill = _fill

        if modo == "dobradas":
            # Uma linha por conta; INICIA_LOTE e empresa só na primeira linha do lote
            lote_val = 1 if first else ""
            emp_val  = cod_empresa if first else ""
            first    = False
            if ind == "D":
                wr(row, 1, data_str); wr(row, 2, code); wr(row, 3, "")
                wr(row, 4, cents);    wr(row, 5, "");   wr(row, 6, hist)
                wr(row, 7, lote_val); wr(row, 8, emp_val)
                wr(row, 9, "");       wr(row, 10, "")
            else:
                wr(row, 1, data_str); wr(row, 2, "");    wr(row, 3, code)
                wr(row, 4, cents);    wr(row, 5, "");    wr(row, 6, hist)
                wr(row, 7, lote_val); wr(row, 8, emp_val)
                wr(row, 9, "");       wr(row, 10, "")
            row += 1

        else:  # simples com contrapartida — 2 linhas por conta (lote próprio)
            if ind == "D":
                wr(row, 1, data_str); wr(row, 2, code);        wr(row, 3, "")
                wr(row, 4, cents);    wr(row, 5, "");           wr(row, 6, hist)
                wr(row, 7, 1);        wr(row, 8, cod_empresa);  wr(row, 9, ""); wr(row, 10, "")
                row += 1
                wr(row, 1, data_str); wr(row, 2, "");           wr(row, 3, contra_code)
                wr(row, 4, cents);    wr(row, 5, "");           wr(row, 6, hist)
                wr(row, 7, "");       wr(row, 8, "");           wr(row, 9, ""); wr(row, 10, "")
                row += 1
            else:
                wr(row, 1, data_str); wr(row, 2, contra_code); wr(row, 3, "")
                wr(row, 4, cents);    wr(row, 5, "");           wr(row, 6, hist)
                wr(row, 7, 1);        wr(row, 8, cod_empresa);  wr(row, 9, ""); wr(row, 10, "")
                row += 1
                wr(row, 1, data_str); wr(row, 2, "");           wr(row, 3, code)
                wr(row, 4, cents);    wr(row, 5, "");           wr(row, 6, hist)
                wr(row, 7, "");       wr(row, 8, "");           wr(row, 9, ""); wr(row, 10, "")
                row += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# GERAÇÃO DO TXT PARA DOMÍNIO
# ─────────────────────────────────────────────────────────────────────────────

def generate_dominio_txt(
    df: pd.DataFrame,
    cod_empresa: str,
    opening_date: date,
    contra_code: str = "",
    modo: str = "dobradas",
) -> str:
    """
    Formato: DATA;DEB;CRE;VALOR;COD_HIST;COMPLEMENTO;INICIA_LOTE;EMPRESA;CC_DEB;CC_CRE
    Valor em centavos (inteiro).
    modo='dobradas': todas as contas em 1 único lançamento (partidas dobradas).
    modo='simples':  1 lançamento por conta com contrapartida.
    """
    lines    = []
    date_str = opening_date.strftime("%d/%m/%Y")
    first    = True

    for _, acc in df.iterrows():
        val  = acc.get("current_value", 0)
        ind  = acc.get("current_indicator", "D")
        code = str(acc.get("code", "")).strip()
        desc = str(acc.get("description", "")).strip()[:55]
        if val == 0 or not code:
            continue

        cents = int(round(val * 100))
        hist  = f"IMPLANTAÇÃO DE SALDO EM {date_str} - {desc}"

        if modo == "dobradas":
            lote = "1" if first else ""
            emp  = cod_empresa if first else ""
            first = False
            if ind == "D":
                lines.append(f"{date_str};{code};;{cents};;{hist};{lote};{emp};;")
            else:
                lines.append(f"{date_str};;{code};{cents};;{hist};{lote};{emp};;")
        else:
            if ind == "D":
                lines.append(f"{date_str};{code};;{cents};;{hist};1;{cod_empresa};;")
                lines.append(f"{date_str};;{contra_code};{cents};;{hist};;;;")
            else:
                lines.append(f"{date_str};{contra_code};;{cents};;{hist};1;{cod_empresa};;")
                lines.append(f"{date_str};;{code};{cents};;{hist};;;;")

    return "\r\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# GERAÇÃO DA PLANILHA DE-PARA (ECD Ficha I057)
# ─────────────────────────────────────────────────────────────────────────────

def generate_depara_excel(df: pd.DataFrame) -> io.BytesIO:
    """
    Ficha Auxiliar - De x Para ECD ficha I057.
    Colunas: Tipo | Cod.Conta Anterior | Classif.Anterior | Cod.Conta Atual(BHub)
             | Classif.BHub | Saldo | Natureza | Centro de custo
    Linhas de dados a partir da linha 7.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "DePara_I057"

    # ── Título ──────────────────────────────────────────────────────────────
    ws["A2"] = "Ficha Auxiliar - De x Para ECD fiha I057"
    ws["A2"].font = Font(bold=True, size=13, color="1F4E79")
    ws.merge_cells("A2:H2")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # ── Cabeçalho linha 5 ───────────────────────────────────────────────────
    headers5 = [
        "Tipo",
        "Codigo da Conta Anterior",
        "Classificação da conta Anterior",
        "Codigo da Conta Atual (BHub)",
        "Classificação BHub",
        "Saldo",
        "Natureza\n(C=Credora/D=Devedora)",
        "Centro de custo",
    ]
    _hdr(ws, 5, headers5)
    ws.row_dimensions[5].height = 30

    # ── Sub-cabeçalho linha 6 ───────────────────────────────────────────────
    ws.cell(6, 1, "ID").font = Font(bold=True)

    # Larguras
    col_w = [6, 25, 30, 28, 30, 18, 22, 14]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A7"

    alt = PatternFill("solid", fgColor="EEF2F7")
    yellow = PatternFill("solid", fgColor="FFFACD")  # indica campo a preencher

    # ── Dados a partir da linha 7 ───────────────────────────────────────────
    for r, (_, acc) in enumerate(df.iterrows(), 7):
        fill = alt if r % 2 == 0 else None

        def wr(col, val, custom_fill=None, _r=r, _fill=fill):
            cell = ws.cell(_r, col, val)
            if custom_fill:
                cell.fill = custom_fill
            elif _fill:
                cell.fill = _fill

        bhub_pre  = str(acc.get("bhub_code", "")).strip()
        bhub_mask = str(acc.get("bhub_mask", "")).strip()
        wr(1, "C")                                             # Tipo
        wr(2, acc.get("code", ""))                             # Código anterior
        wr(3, acc.get("classification", ""))                   # Classif anterior
        wr(4, bhub_pre if bhub_pre else "", yellow)            # Código BHub
        wr(5, bhub_mask if bhub_mask else "", yellow)          # Classif BHub
        wr(6, acc.get("current_value", 0))                     # Saldo
        wr(7, acc.get("current_indicator", ""))                # Natureza D/C
        wr(8, "")                                              # Centro de custo

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK — exclusão de contas da lista principal
# Roda ANTES do rerun, enquanto st.session_state["viz_editor"] ainda tem os
# checkboxes marcados. Não pode ser lambda nem função local (restrição Streamlit).
# ─────────────────────────────────────────────────────────────────────────────

def _excluir_marcadas_cb():
    """Remove do df_all as contas marcadas com 🗑️ no data_editor da listagem.
    Lê de _viz_edited_df (DataFrame completo salvo no render anterior)
    porque st.session_state["viz_editor"] guarda apenas o delta interno do Streamlit.

    Persiste as exclusões em excluded_accounts para sobreviverem ao rerun do loop de PDF.
    """
    edited = st.session_state.get("_viz_edited_df")
    if edited is None or not isinstance(edited, pd.DataFrame) or "🗑️" not in edited.columns:
        return
    to_del = edited[edited["🗑️"] == True]
    if len(to_del) == 0:
        return

    excl_keys = set(
        zip(to_del["Código"].astype(str), to_del["Empresa"].astype(str))
    )

    # Persiste para sobreviver ao reprocessamento de PDF no próximo rerun
    if "excluded_accounts" not in st.session_state:
        st.session_state["excluded_accounts"] = set()
    st.session_state["excluded_accounts"] = st.session_state["excluded_accounts"] | excl_keys

    # Remove de manual_accounts (contas manuais excluídas não devem ser re-mescladas)
    st.session_state["manual_accounts"] = [
        a for a in st.session_state.get("manual_accounts", [])
        if (str(a.get("code", "")), str(a.get("empresa", ""))) not in excl_keys
    ]

    # Guarda contagem p/ mensagem de confirmação e limpa estado do editor
    st.session_state["_excluidas_count"] = len(to_del)
    st.session_state.pop("viz_editor", None)
    st.session_state.pop("_viz_edited_df", None)


# ─────────────────────────────────────────────────────────────────────────────
# APP STREAMLIT
# ─────────────────────────────────────────────────────────────────────────────

def _load_anthropic_key() -> str:
    """Lê a chave API do .streamlit/secrets.toml ou da variável de ambiente."""
    import os
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        key = ""
    if not key:
        key = os.environ.get("ANTHROPIC_API_KEY", "")
    return key.strip()


# ─────────────────────────────────────────────────────────────────────────────
# CARTA DE RESPONSABILIDADE – CONSTANTES E FUNÇÕES DE GERAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

BHUB_CNPJ_CONTABIL = "43.618.130/0001-03"

PREDEFINED_JUSTIFICATIONS = [
    "",
    "Implantação de saldo via ECD",
    "Constituição de saldo",
    "Validado via Extrato Bancário",
    "Validação via Assinatura de carta de responsabilidade",
    "Saldo constituído via conta de contra-partida",
    "Validado via e-Social",
    "Validado via e-CAC",
    "Validado via FGTS Digital",
    "Validado via PGDAS-D",
    "Validado via QSA",
]

_MESES_PT = {
    "January": "Janeiro", "February": "Fevereiro", "March": "Março",
    "April": "Abril", "May": "Maio", "June": "Junho",
    "July": "Julho", "August": "Agosto", "September": "Setembro",
    "October": "Outubro", "November": "Novembro", "December": "Dezembro",
}


_MESES_PT = [
    "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

def _fmt_cnpj(v: str) -> str:
    """Formata CNPJ para XX.XXX.XXX/XXXX-XX independentemente da entrada."""
    import re as _re
    digits = _re.sub(r'\D', '', str(v).strip())
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"
    return str(v).strip()


def _fmt_date_ptbr(d) -> str:
    """Converte date para '01 de janeiro de 2026'."""
    try:
        return f"{d.day:02d} de {_MESES_PT[d.month]} de {d.year}"
    except Exception:
        return str(d)


def _fmt_currency_carta(v: float) -> str:
    """Formata valor como R$ X.XXX,XX ou -R$ X.XXX,XX."""
    if v < 0:
        return f"-R$ {format_br(abs(v))}"
    return f"R$ {format_br(abs(v))}"


def generate_carta_excel(info: dict, accounts: list) -> io.BytesIO:
    """Gera planilha Excel no formato análise de cliente / carta de responsabilidade."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Análise de Cliente"

    title_font    = Font(name="Arial", bold=True, size=13)
    label_font    = Font(name="Arial", bold=True, size=10)
    normal_font   = Font(name="Arial", size=10)
    header_font   = Font(name="Arial", bold=True, size=9, color="FFFFFF")
    header_fill   = PatternFill("solid", fgColor="000000")
    section_fill  = PatternFill("solid", fgColor="E8E8E8")
    alt_fill      = PatternFill("solid", fgColor="F2F2F2")
    center_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    right_align   = Alignment(horizontal="right",  vertical="center")
    thin          = Side(style="thin")
    border        = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _w(row, col, val, fnt=None, fil=None, aln=None, brd=None):
        c = ws.cell(row=row, column=col, value=val)
        if fnt: c.font       = fnt
        if fil: c.fill       = fil
        if aln: c.alignment  = aln
        if brd: c.border     = brd
        return c

    # ── Cabeçalho geral ──────────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    _w(1, 1, "ANÁLISE DE CLIENTE", title_font, section_fill, center_align)
    ws.row_dimensions[1].height = 26

    row = 3
    meta_fields = [
        ("Cliente", info.get("cliente", "")),
        ("CNPJ",    info.get("cnpj", "")),
        ("I.E.",    info.get("ie", "")),
        ("CCM",     info.get("ccm", "")),
        ("Data do contrato",  info.get("data_contrato", "")),
        ("Plano",   info.get("plano", "Contabilidade as a Service 2.0")),
    ]
    for lbl, val in meta_fields:
        ws.merge_cells(f"A{row}:G{row}")
        _w(row, 1, f"{lbl}: {val}", label_font, aln=left_align)
        row += 1

    # Competência BHub + Link na mesma linha
    ws.merge_cells(f"A{row}:C{row}")
    _w(row, 1, f"Competência BHub: {info.get('competencia_bhub', '')}", label_font, aln=left_align)
    ws.merge_cells(f"D{row}:G{row}")
    _w(row, 4, f"Link Balancete: {info.get('link_balancete', '')}", normal_font, aln=left_align)
    row += 1

    ws.merge_cells(f"A{row}:G{row}")
    _w(row, 1, f"Balancete Implantado: {info.get('data_balancete', '')}", label_font, aln=left_align)
    row += 2

    # ── Título da seção de ajustes ────────────────────────────────────────────
    ws.merge_cells(f"A{row}:G{row}")
    _w(row, 1, "Sugestões de Ajustes", Font(name="Arial", bold=True, size=11), section_fill, center_align)
    ws.row_dimensions[row].height = 20
    row += 2

    # ── Cabeçalho da tabela ───────────────────────────────────────────────────
    date_lbl = info.get("data_balancete", "dd/mm/aaaa")
    col_headers = [
        "Conta de Origem",
        f"Saldo Balancete ({date_lbl})",
        "Valor de Ajuste",
        "Saldo Final/Saldo Suporte",
        "Justificativa",
        "Sugestão de Contra-Partida – BHub",
    ]
    for ci, h in enumerate(col_headers, 1):
        _w(row, ci, h, header_font, header_fill, center_align, border)
    ws.row_dimensions[row].height = 30
    row += 1

    # ── Linhas de contas ──────────────────────────────────────────────────────
    for i, acc in enumerate(accounts):
        row_fill = alt_fill if i % 2 == 0 else None
        saldo       = float(acc.get("saldo", 0) or 0)
        ajuste      = float(acc.get("ajuste", 0) or 0)
        saldo_final = float(acc.get("saldo_final", saldo + ajuste) or saldo + ajuste)
        _w(row, 1, acc.get("conta", ""),          normal_font, row_fill, left_align,  border)
        _w(row, 2, _fmt_currency_carta(saldo),    normal_font, row_fill, right_align, border)
        _w(row, 3, _fmt_currency_carta(ajuste) if ajuste != 0 else "",
                                                  normal_font, row_fill, right_align, border)
        _w(row, 4, _fmt_currency_carta(saldo_final), normal_font, row_fill, right_align, border)
        _w(row, 5, acc.get("justificativa", ""),  normal_font, row_fill, left_align,  border)
        _w(row, 6, acc.get("contra_partida", ""), normal_font, row_fill, left_align,  border)
        ws.row_dimensions[row].height = 16
        row += 1

    # ── Larguras ──────────────────────────────────────────────────────────────
    for ci, w in enumerate([38, 22, 18, 24, 48, 34], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_carta_word(info: dict, accounts: list) -> io.BytesIO:
    """Gera documento Word (.docx) da Carta de Responsabilidade."""
    try:
        from docx import Document as _Doc
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WALIGN
        from docx.enum.table import WD_TABLE_ALIGNMENT as _TALIGN, WD_ALIGN_VERTICAL as _VALIGN
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _OxmlElement
    except ImportError:
        _pip_install("python-docx>=1.0.0")
        from docx import Document as _Doc
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WALIGN
        from docx.enum.table import WD_TABLE_ALIGNMENT as _TALIGN, WD_ALIGN_VERTICAL as _VALIGN
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _OxmlElement

    doc = _Doc()
    for sec in doc.sections:
        sec.top_margin    = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin   = Cm(3.0)
        sec.right_margin  = Cm(2.0)

    def _run(p, text, bold=False, size=11, color=None):
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(size)
        r.font.bold = bold
        if color:
            r.font.color.rgb = RGBColor(*color)
        return r

    def _para(text="", align=None, bold=False, size=11, sa=6, sb=0):
        p = doc.add_paragraph()
        p.alignment = align or _WALIGN.LEFT
        p.paragraph_format.space_after  = Pt(sa)
        p.paragraph_format.space_before = Pt(sb)
        if text:
            _run(p, text, bold=bold, size=size)
        return p

    def _cell_shd(cell, hex_color):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = _OxmlElement("w:shd")
        shd.set(_qn("w:val"),   "clear")
        shd.set(_qn("w:color"), "auto")
        shd.set(_qn("w:fill"),  hex_color)
        tcPr.append(shd)

    # ── Título ────────────────────────────────────────────────────────────────
    _para("Carta de Responsabilidade da Administração",
          align=_WALIGN.CENTER, bold=True, size=16, sa=0)
    _para("sobre os saldos contábeis",
          align=_WALIGN.CENTER, bold=True, size=16, sa=14)

    # ── Destinatário ──────────────────────────────────────────────────────────
    for txt, b in [("À", False), ("BHUB CONTABILIDADE LTDA.", True),
                   (f"CNPJ: {BHUB_CNPJ_CONTABIL}", True),
                   ("Endereço: AV FRANCISCO MATARAZZO, 1500 - Andar 19 Parte – Água Branca", True),
                   ("SAO PAULO, SP – CEP: 05.001-100", True)]:
        _para(txt, bold=b, size=11, sa=0)
    _para("", sa=8)

    _para("Prezados Senhor(a):", bold=False, size=11, sa=10)

    # ── Corpo ─────────────────────────────────────────────────────────────────
    data_bal    = info.get("data_balancete", "")
    competencia = info.get("competencia_bhub", "")
    nome_emp    = info.get("cliente", "")
    cnpj_emp    = info.get("cnpj", "")

    body_p = doc.add_paragraph()
    body_p.paragraph_format.space_after  = Pt(10)
    body_p.paragraph_format.space_before = Pt(0)
    segments = [
        ("Conforme observado durante os procedimentos de transição de contabilidade, na qual a ", False),
        ('BHUB CONTABILIDADE LTDA ("BHub"),', True),
        (f" assumiu a contabilidade da empresa ", False),
        (f"{nome_emp}", True),
        (f", CNPJ: {cnpj_emp},", True),
        (f" a partir do período-base {competencia}, não foram compartilhadas as conciliações contábeis "
         "para as contas listadas abaixo e compreendidas no balancete para o período-base "
         f"{data_bal} fornecido pela antiga contabilidade, e utilizado para implantação dos "
         "saldos iniciais, impossibilitando a confirmação sobre a acuracidade das informações "
         "contábeis recebidas. Além disso, tenho ciência do ajuste a ser efetuado no início de "
         "competência da ", False),
        ('BHUB CONTABILIDADE LTDA ("BHub"),', True),
        (" conforme exposto abaixo:", False),
    ]
    for seg_txt, seg_bold in segments:
        _run(body_p, seg_txt, bold=seg_bold, size=11)

    # ── Tabela de contas ──────────────────────────────────────────────────────
    n_rows = max(len(accounts), 1) + 1
    tbl = doc.add_table(rows=n_rows, cols=4)
    tbl.style = "Table Grid"
    tbl.alignment = _TALIGN.CENTER

    hdr_texts = [
        "Conta de Origem",
        f"Saldo Balancete\n({data_bal})",
        "Valor de Ajuste",
        "Saldo Final/\nSaldo Suporte",
    ]
    for ci, ht in enumerate(hdr_texts):
        cell = tbl.rows[0].cells[ci]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = _WALIGN.CENTER
        cell.vertical_alignment = _VALIGN.CENTER
        _run(p, ht, bold=True, size=9, color=(0xFF, 0xFF, 0xFF))
        _cell_shd(cell, "000000")

    for ri, acc in enumerate(accounts, start=1):
        cells = tbl.rows[ri].cells
        saldo       = float(acc.get("saldo", 0) or 0)
        ajuste      = float(acc.get("ajuste", 0) or 0)
        saldo_final = float(acc.get("saldo_final", saldo + ajuste) or saldo + ajuste)

        def _tc(cell, text, align=_WALIGN.LEFT):
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = align
            cell.vertical_alignment = _VALIGN.CENTER
            _run(p, text, bold=False, size=8)

        _tc(cells[0], acc.get("conta", ""))
        _tc(cells[1], _fmt_currency_carta(saldo),       align=_WALIGN.RIGHT)
        _tc(cells[2], _fmt_currency_carta(ajuste) if ajuste != 0 else "", align=_WALIGN.RIGHT)
        _tc(cells[3], _fmt_currency_carta(saldo_final), align=_WALIGN.RIGHT)

        if ri % 2 == 0:
            for cell in tbl.rows[ri].cells:
                _cell_shd(cell, "F2F2F2")

    # ── Espaço após tabela ────────────────────────────────────────────────────
    _para("", sa=6)

    # ── Declaração ────────────────────────────────────────────────────────────
    _para("Dessa forma, como responsável legal, declaro que:", sa=6)

    for bullet in [
        f"As informações a serem consideradas como saldos finais para o período-base {data_bal}, são as apresentadas; e",
        'Tenho ciência dos impactos trazidos pela BHUB CONTABILIDADE LTDA ("BHub") da não entrega '
        "das documentações listadas acima e estou de acordo com o ajuste supracitado.",
    ]:
        bp = doc.add_paragraph()
        bp.paragraph_format.space_after      = Pt(4)
        bp.paragraph_format.left_indent      = Cm(1.0)
        bp.paragraph_format.first_line_indent = Cm(-0.5)
        _run(bp, "•  " + bullet, bold=False, size=11)

    _para("Também confirmamos que não houve/temos conhecimento de:", sa=4, sb=8)

    for item in [
        "(a)  fraude envolvendo administração ou empregados em cargos de responsabilidade ou confiança;",
        "(b)  fraude envolvendo terceiros que poderiam ter efeito material nas demonstrações contábeis;",
        "(c)  violação ou possíveis violações de leis, normas ou regulamentos cujos efeitos deveriam "
              "ser considerados para divulgação nas demonstrações contábeis, ou mesmo dar origem ao "
              "registro de provisão para contingências passivas.",
    ]:
        ip = doc.add_paragraph()
        ip.paragraph_format.left_indent = Cm(0.5)
        ip.paragraph_format.space_after = Pt(3)
        _run(ip, item, bold=False, size=11)

    # ── Cidade, data e assinatura ─────────────────────────────────────────────
    cidade     = info.get("cidade", "São Paulo")
    data_carta = info.get("data_carta", "")
    _para(f"{cidade}, {data_carta}", align=_WALIGN.CENTER, sa=4, sb=14)
    _para("Atenciosamente,", align=_WALIGN.CENTER, sa=28)
    _para("..............................................................", align=_WALIGN.CENTER, sa=0)
    _para(info.get("representante", "").upper(), align=_WALIGN.CENTER, sa=0)
    _para("Representante Legal", align=_WALIGN.CENTER, sa=0)
    if info.get("cpf", "").strip():
        _para(f"CPF: {info['cpf'].strip()}", align=_WALIGN.CENTER, sa=0)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def generate_carta_pdf(info: dict, accounts: list) -> io.BytesIO:
    """Gera PDF da Carta de Responsabilidade usando reportlab."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib import colors as _rlcolors
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                        Table, TableStyle, HRFlowable)
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    except ImportError:
        _pip_install("reportlab>=4.0.0")
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib import colors as _rlcolors
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                        Table, TableStyle, HRFlowable)
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY

    buf = io.BytesIO()
    doc_pdf = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=2.5*cm, bottomMargin=2.5*cm,
        leftMargin=3*cm, rightMargin=2*cm,
    )

    _hdr_color  = _rlcolors.HexColor("#000000")
    _alt_color  = _rlcolors.HexColor("#F2F2F2")
    _white      = _rlcolors.white
    _grey       = _rlcolors.HexColor("#AAAAAA")

    def _ps(name, fontName="Helvetica", fontSize=11, alignment=TA_LEFT,
            leading=14, spaceAfter=4, spaceBefore=0, bold=False, color=None):
        kwargs = dict(fontName=("Helvetica-Bold" if bold else fontName),
                      fontSize=fontSize, alignment=alignment,
                      leading=leading, spaceAfter=spaceAfter, spaceBefore=spaceBefore)
        if color:
            kwargs["textColor"] = color
        return ParagraphStyle(name, **kwargs)

    s_title   = _ps("title",  bold=True,  fontSize=16, alignment=TA_CENTER, spaceAfter=14)
    s_bold    = _ps("bold",   bold=True,  fontSize=11, spaceAfter=2)
    s_normal  = _ps("norm",   fontSize=11, spaceAfter=4)
    s_body    = _ps("body",   fontSize=11, alignment=TA_JUSTIFY, spaceAfter=10)
    s_center  = _ps("ctr",    fontSize=11, alignment=TA_CENTER, spaceAfter=4)
    s_th      = _ps("th",     bold=True,  fontSize=8.5, alignment=TA_CENTER,
                    spaceAfter=1, color=_white)
    s_td      = _ps("td",     fontSize=8.5, leading=11, spaceAfter=1)
    s_td_r    = _ps("td_r",   fontSize=8.5, leading=11, spaceAfter=1, alignment=TA_RIGHT)
    s_bullet  = ParagraphStyle("blt", fontName="Helvetica", fontSize=11, leading=14,
                               spaceAfter=6, leftIndent=18, firstLineIndent=-12)

    story = []

    story.append(Paragraph("Carta de Responsabilidade da Administração", s_title))
    story.append(Paragraph("sobre os saldos contábeis", s_title))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("À", s_normal))
    story.append(Paragraph("<b>BHUB CONTABILIDADE LTDA.</b>", s_normal))
    story.append(Paragraph(f"<b>CNPJ: {BHUB_CNPJ_CONTABIL}</b>", s_normal))
    story.append(Paragraph("<b>Endereço: AV FRANCISCO MATARAZZO, 1500 - Andar 19 Parte – Água Branca</b>", s_normal))
    story.append(Paragraph("<b>SAO PAULO, SP – CEP: 05.001-100</b>", s_normal))
    story.append(Spacer(1, 0.25*cm))
    story.append(Paragraph("<b>Prezados Senhor(a):</b>", s_normal))
    story.append(Spacer(1, 0.25*cm))

    data_bal    = info.get("data_balancete", "")
    competencia = info.get("competencia_bhub", "")
    nome_emp    = info.get("cliente", "")
    cnpj_emp    = info.get("cnpj", "")

    body_html = (
        'Conforme observado durante os procedimentos de transição de contabilidade, na qual a '
        '<b>BHUB CONTABILIDADE LTDA ("BHub"),</b> assumiu a contabilidade da empresa '
        f'<b>{nome_emp}, CNPJ: {cnpj_emp},</b> a partir do período-base '
        f'{competencia}, não foram compartilhadas as conciliações contábeis para as contas '
        f'listadas abaixo e compreendidas no balancete para o período-base {data_bal} fornecido '
        'pela antiga contabilidade, e utilizado para implantação dos saldos iniciais, '
        'impossibilitando a confirmação sobre a acuracidade das informações contábeis recebidas. '
        'Além disso, tenho ciência do ajuste a ser efetuado no início de competência da '
        '<b>BHUB CONTABILIDADE LTDA ("BHub"),</b> conforme exposto abaixo:'
    )
    story.append(Paragraph(body_html, s_body))
    story.append(Spacer(1, 0.2*cm))

    # ── Tabela de contas ──────────────────────────────────────────────────────
    pw = A4[0] - 5*cm
    col_ws = [pw * 0.40, pw * 0.20, pw * 0.18, pw * 0.22]

    tbl_data = [[
        Paragraph(f"<b>Conta de Origem</b>", s_th),
        Paragraph(f"<b>Saldo Balancete<br/>({data_bal})</b>", s_th),
        Paragraph("<b>Valor de Ajuste</b>", s_th),
        Paragraph("<b>Saldo Final/<br/>Saldo Suporte</b>", s_th),
    ]]
    row_fills_pdf = []
    for i, acc in enumerate(accounts):
        saldo       = float(acc.get("saldo", 0) or 0)
        ajuste      = float(acc.get("ajuste", 0) or 0)
        saldo_final = float(acc.get("saldo_final", saldo + ajuste) or saldo + ajuste)
        tbl_data.append([
            Paragraph(acc.get("conta", ""), s_td),
            Paragraph(_fmt_currency_carta(saldo), s_td_r),
            Paragraph(_fmt_currency_carta(ajuste) if ajuste != 0 else "", s_td_r),
            Paragraph(_fmt_currency_carta(saldo_final), s_td_r),
        ])
        if i % 2 == 0:
            row_fills_pdf.append(("BACKGROUND", (0, i + 1), (-1, i + 1), _alt_color))

    tbl_pdf = Table(tbl_data, colWidths=col_ws, repeatRows=1)
    ts_pdf  = TableStyle([
        ("BACKGROUND",  (0, 0), (-1, 0), _hdr_color),
        ("GRID",        (0, 0), (-1, -1), 0.4, _grey),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",(0, 0), (-1, -1), 4),
    ] + row_fills_pdf)
    tbl_pdf.setStyle(ts_pdf)
    story.append(tbl_pdf)
    story.append(Spacer(1, 0.35*cm))

    story.append(Paragraph("Dessa forma, como responsável legal, declaro que:", s_normal))
    story.append(Paragraph(
        f"•  As informações a serem consideradas como saldos finais para o período-base "
        f"<b>{data_bal}</b>, são as apresentadas; e", s_bullet))
    story.append(Paragraph(
        '•  Tenho ciência dos impactos trazidos pela <b>BHUB CONTABILIDADE LTDA ("BHub")</b> '
        "da não entrega das documentações listadas acima e estou de acordo com o ajuste supracitado.",
        s_bullet))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("Também confirmamos que não houve/temos conhecimento de:", s_normal))
    for item_txt in [
        "(a)  fraude envolvendo administração ou empregados em cargos de responsabilidade ou confiança;",
        "(b)  fraude envolvendo terceiros que poderiam ter efeito material nas demonstrações contábeis;",
        "(c)  violação ou possíveis violações de leis, normas ou regulamentos cujos efeitos deveriam ser "
              "considerados para divulgação nas demonstrações contábeis, ou mesmo dar origem ao registro "
              "de provisão para contingências passivas.",
    ]:
        story.append(Paragraph(item_txt, s_normal))

    story.append(Spacer(1, 0.5*cm))
    cidade     = info.get("cidade", "São Paulo")
    data_carta = info.get("data_carta", "")
    story.append(Paragraph(f"{cidade}, {data_carta}", s_center))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("Atenciosamente,", s_center))
    story.append(Spacer(1, 1.2*cm))
    story.append(Paragraph("..............................................................", s_center))
    story.append(Paragraph(info.get("representante", "").upper(), s_center))
    story.append(Paragraph("Representante Legal", s_center))
    if info.get("cpf", "").strip():
        story.append(Paragraph(f"CPF: {info['cpf'].strip()}", s_center))

    doc_pdf.build(story)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# LANÇAMENTO DE AJUSTE – DOMÍNIO
# ─────────────────────────────────────────────────────────────────────────────

def _parse_conta_code(conta_str: str) -> tuple:
    """Extrai (código, descrição) de 'COD - Descrição' ou retorna ('', texto)."""
    s = str(conta_str).strip()
    import re as _re
    m = _re.match(r"^(\S+)\s*[-–]\s*(.+)$", s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", s


def generate_ajuste_excel(accounts: list, ajuste_date: date, cod_empresa: str) -> io.BytesIO:
    """Gera Excel Domínio para os lançamentos de ajuste da carta (1 lote por conta)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Ajustes"

    date_str = ajuste_date.strftime("%d/%m/%Y")
    ws["A1"] = (
        f"LANÇAMENTOS DE AJUSTE – Empresa: {cod_empresa} | Data: {date_str}"
    )
    ws["A1"].font = Font(bold=True, color="1F4E79", size=11)
    ws.merge_cells("A1:J1")

    headers = [
        "Data", "Cód. Conta Débito", "Cód. Conta Crédito", "Valor",
        "Cód. Histórico", "Complemento Histórico", "Inicia Lote",
        "Código Matriz/Filial", "Centro de Custo Débito", "Centro de Custo Crédito",
    ]
    _hdr(ws, 2, headers)

    col_w = [14, 18, 18, 16, 14, 55, 12, 20, 18, 18]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A3"

    alt = PatternFill("solid", fgColor="EEF2F7")
    row = 3

    for acc in accounts:
        ajuste = float(acc.get("ajuste", 0) or 0)
        if ajuste == 0:
            continue

        code, desc = _parse_conta_code(acc.get("conta", ""))
        if not code:
            continue

        contra_raw = str(acc.get("contra_partida", "")).strip()
        contra, _ = _parse_conta_code(contra_raw)
        abs_val = abs(ajuste)
        cents   = int(round(abs_val * 100))
        hist    = f"AJUSTE DE IMPLANTAÇÃO {date_str} - {desc}"[:55]
        fill1   = alt if row % 2 == 0 else None

        def wr(r, c, v, f=None):
            cell = ws.cell(r, c, v)
            if f:
                cell.fill = f

        if ajuste > 0:
            # Débito na conta → Crédito na contra
            wr(row, 1, date_str, fill1); wr(row, 2, code, fill1); wr(row, 3, "", fill1)
            wr(row, 4, cents, fill1);    wr(row, 5, "", fill1);   wr(row, 6, hist, fill1)
            wr(row, 7, 1, fill1);        wr(row, 8, cod_empresa, fill1)
            wr(row, 9, "", fill1);       wr(row, 10, "", fill1)
            row += 1
            if contra:
                fill2 = alt if row % 2 == 0 else None
                wr(row, 1, date_str, fill2); wr(row, 2, "", fill2); wr(row, 3, contra, fill2)
                wr(row, 4, cents, fill2);    wr(row, 5, "", fill2); wr(row, 6, hist, fill2)
                wr(row, 7, "", fill2);       wr(row, 8, "", fill2)
                wr(row, 9, "", fill2);       wr(row, 10, "", fill2)
                row += 1
        else:
            # Crédito na conta → Débito na contra
            if contra:
                fill1c = alt if row % 2 == 0 else None
                wr(row, 1, date_str, fill1c); wr(row, 2, contra, fill1c); wr(row, 3, "", fill1c)
                wr(row, 4, cents, fill1c);    wr(row, 5, "", fill1c);     wr(row, 6, hist, fill1c)
                wr(row, 7, 1, fill1c);        wr(row, 8, cod_empresa, fill1c)
                wr(row, 9, "", fill1c);       wr(row, 10, "", fill1c)
                row += 1
            fill2 = alt if row % 2 == 0 else None
            wr(row, 1, date_str, fill2); wr(row, 2, "", fill2); wr(row, 3, code, fill2)
            wr(row, 4, cents, fill2);    wr(row, 5, "", fill2); wr(row, 6, hist, fill2)
            lote_v = "" if contra else 1
            emp_v  = "" if contra else cod_empresa
            wr(row, 7, lote_v, fill2); wr(row, 8, emp_v, fill2)
            wr(row, 9, "", fill2);     wr(row, 10, "", fill2)
            row += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_ajuste_txt(accounts: list, ajuste_date: date, cod_empresa: str) -> str:
    """Gera TXT Domínio para os lançamentos de ajuste da carta."""
    lines    = []
    date_str = ajuste_date.strftime("%d/%m/%Y")

    for acc in accounts:
        ajuste = float(acc.get("ajuste", 0) or 0)
        if ajuste == 0:
            continue

        code, desc = _parse_conta_code(acc.get("conta", ""))
        if not code:
            continue

        contra_raw = str(acc.get("contra_partida", "")).strip()
        contra, _ = _parse_conta_code(contra_raw)
        abs_val = abs(ajuste)
        cents   = int(round(abs_val * 100))
        hist    = f"AJUSTE DE IMPLANTAÇÃO {date_str} - {desc}"[:55]

        if ajuste > 0:
            lines.append(f"{date_str};{code};;{cents};;{hist};1;{cod_empresa};;")
            if contra:
                lines.append(f"{date_str};;{contra};{cents};;{hist};;;;")
        else:
            if contra:
                lines.append(f"{date_str};{contra};;{cents};;{hist};1;{cod_empresa};;")
            lines.append(f"{date_str};;{code};{cents};;{hist}{';;;;' if contra else f';1;{cod_empresa};;'}")

    return "\r\n".join(lines)


def main():
    st.title("📝 Implantação de saldo e Gerador de Carta de Responsabilidade – BHub")
    st.caption(
        "Extrai dados de balancetes PDF do contador anterior → "
        "Gera De-Para ECD (Ficha I057) + Planilha e TXT para importação no Domínio"
    )

    # ── Sidebar ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Parâmetros")
        cod_empresa = st.text_input("Código do cliente", value="",
                                    placeholder="Código da empresa no Domínio")
        opening_date = st.date_input("Data dos Lançamentos de Abertura", value=date.today(), format="DD/MM/YYYY")
        contra_code = st.text_input(
            "Conta de Contrapartida (modo simples)",
            value="2875",
            help="Usada apenas no modo 'Simples com conta de contrapartida'. No modo padrão (partidas dobradas) este campo é ignorado.",
        )

        st.divider()
        st.subheader("🔍 Filtro de Contas")
        leaf_mode = st.checkbox(
            "Apenas contas analíticas (sem subcontas)",
            value=True,
            help="Detecta automaticamente as contas folha – sem filhas. Recomendado.",
        )
        min_depth = st.slider(
            "Profundidade mínima da classificação",
            min_value=1, max_value=8, value=4,
            help="Filtra por nível de profundidade (ex.: 4 = inclui 1.1.1.02+). Ativo quando 'Apenas analíticas' estiver desmarcado.",
            disabled=leaf_mode,
        )
        include_no_cls = st.checkbox(
            "Incluir contas sem classificação",
            value=False,
            help="Marque se as contas do balancete não possuem campo de classificação",
        )

        st.divider()
        st.markdown("""
        **📋 Guia rápido:**
        1. Carregue os PDFs dos balancetes
        2. Revise as contas extraídas
        3. Baixe o De-Para e complete o código BHub
        4. Gere a planilha e o TXT para o Domínio
        """)

    # ── Abas ─────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📁  1. Carregar Balancetes",
        "🔀  2. De-Para (ECD I057)",
        "📤  3. Gerar Arquivos Domínio",
        "📝  4. Carta de Responsabilidade",
    ])

    # ════════════════════════════════════════════════════════════════════════
    # TAB 1 – UPLOAD
    # ════════════════════════════════════════════════════════════════════════
    with tab1:
        st.header("Carregar Balancetes do Contador Anterior")
        st.info(
            "Carregue um ou mais PDFs de balancetes. "
            "Cada arquivo pode representar uma empresa diferente."
        )

        if "uploader_key" not in st.session_state:
            st.session_state["uploader_key"] = 0

        pdfs = st.file_uploader(
            "Selecione os arquivos PDF dos balancetes",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"pdf_uploader_{st.session_state['uploader_key']}",
        )

        # Inicializar lista de contas adicionadas manualmente (persiste entre reruns)
        if "manual_accounts" not in st.session_state:
            st.session_state["manual_accounts"] = []
        if "excluded_accounts" not in st.session_state:
            st.session_state["excluded_accounts"] = set()
        if "bhub_confirmed" not in st.session_state:
            st.session_state["bhub_confirmed"] = set()
        if "bhub_cancelled" not in st.session_state:
            st.session_state["bhub_cancelled"] = set()

        if pdfs:
            all_dfs = {}

            for pdf in pdfs:
                # Garante rótulo único mesmo que dois PDFs tenham o mesmo nome de arquivo
                _base_label = pdf.name.rsplit(".", 1)[0]
                company_label = _base_label
                _sfx = 2
                while company_label in all_dfs:
                    company_label = f"{_base_label}_{_sfx}"
                    _sfx += 1

                with st.expander(f"📄 {company_label}", expanded=True):
                    with st.spinner(f"Extraindo dados de {pdf.name}..."):
                        try:
                            records = extract_pdf(pdf)
                            if not records:
                                st.warning("Nenhuma conta encontrada. Verifique o PDF.")
                                continue

                            # Detectar profundidade analítica por grupo antes de anotar
                            depth_summary = get_depth_summary(records)

                            # Anotar tipo A/S em TODOS os registros antes de filtrar
                            records = annotate_tipo(records)
                            df = pd.DataFrame(records)
                            df["grupo"] = df["classification"].apply(get_account_group)

                            # Filtrar
                            has_any_cls = df["classification"].astype(bool).any()
                            if leaf_mode or not has_any_cls:
                                # Sem classificações: annotate_tipo já marcou tudo como "A"
                                if include_no_cls:
                                    mask = (df["tipo"] == "A") | (df["classification"] == "")
                                else:
                                    mask = df["tipo"] == "A"
                            else:
                                if include_no_cls:
                                    mask = df["classification"].apply(
                                        lambda c: cls_depth(c) >= min_depth or not c
                                    )
                                else:
                                    mask = df["classification"].apply(
                                        lambda c: cls_depth(c) >= min_depth
                                    )
                            df_filtered = df[mask].copy()
                            df_filtered["empresa"] = company_label

                            # ── Métricas ────────────────────────────────────
                            total = len(df_filtered)
                            # D=positivo, C=negativo (convenção: ativo devedor, passivo credor)
                            tot_ativo   = signed_total(df_filtered[df_filtered["grupo"] == "Ativo"])
                            tot_passivo = signed_total(df_filtered[df_filtered["grupo"] == "Passivo/PL"])
                            tot_result  = signed_total(df_filtered[df_filtered["grupo"] == "Resultado"])

                            c1, c2, c3, c4 = st.columns(4)
                            c1.metric("Contas analíticas", total)
                            c2.metric("Total Ativo (1)", f"R$ {format_br(tot_ativo)}")
                            c3.metric("Total Passivo/PL (2)", f"R$ {format_br(tot_passivo)}")
                            c4.metric("Total Resultado (3+)", f"R$ {format_br(tot_result)}")

                            # Profundidade analítica detectada por grupo
                            group_labels = {"1": "Ativo", "2": "Passivo/PL"}
                            depth_parts = []
                            for g, d in sorted(depth_summary.items()):
                                label = group_labels.get(g, f"Resultado({g})")
                                depth_parts.append(f"**{label}**: {d} níveis")
                            if depth_parts:
                                st.caption("Profundidade analítica detectada — " + " | ".join(depth_parts))

                            # Verificação: débitos totais = créditos totais
                            tot_d = df_filtered[df_filtered["current_indicator"] == "D"]["current_value"].sum()
                            tot_c = df_filtered[df_filtered["current_indicator"] == "C"]["current_value"].sum()
                            dif = abs(tot_d - tot_c)
                            if dif > 0.01:
                                st.info(
                                    f"ℹ️ Diferença D/C: R$ {format_br(dif)} "
                                    "(normal quando há contas de resultado ou saldos mistos)"
                                )

                            # ── Tabela ───────────────────────────────────────
                            st.dataframe(
                                df_filtered[
                                    ["tipo", "grupo", "code", "classification",
                                     "description", "current_value", "current_indicator"]
                                ].rename(columns={
                                    "tipo": "Tipo",
                                    "grupo": "Grupo",
                                    "code": "Código",
                                    "classification": "Classificação",
                                    "description": "Descrição",
                                    "current_value": "Saldo Atual",
                                    "current_indicator": "D/C",
                                }),
                                use_container_width=True,
                                height=300,
                                column_config={
                                    "Tipo": st.column_config.TextColumn(
                                        "Tipo", width="small",
                                        help="A = Analítica (folha) | S = Sintética (grupo)"
                                    ),
                                    "Grupo": st.column_config.TextColumn("Grupo", width="medium"),
                                    "Saldo Atual": st.column_config.NumberColumn(
                                        "Saldo Atual", format="R$ %.2f"
                                    ),
                                },
                            )

                            # Guardar TODOS os registros (pré-filtro) para busca manual
                            if "all_records" not in st.session_state:
                                st.session_state["all_records"] = {}
                            st.session_state["all_records"][company_label] = records

                            # ── Detecção de estrutura BHub ────────────────────
                            _confirmed_set = st.session_state["bhub_confirmed"]
                            _cancelled_set = st.session_state["bhub_cancelled"]

                            if company_label in _cancelled_set:
                                st.error(
                                    "❌ Balancete cancelado. Remova este arquivo e carregue "
                                    "o balancete correto do contador anterior."
                                )

                            elif company_label not in _confirmed_set:
                                _is_bhub, _bhub_reason = _is_bhub_structure(df_filtered)
                                if _is_bhub:
                                    st.warning(
                                        "⚠️ **Este balancete parece ter a estrutura do plano BHub.**\n\n"
                                        f"Motivo: {_bhub_reason}.\n\n"
                                        "Esta ferramenta é para importar balancetes do **contador anterior** "
                                        "(plano de contas antigo). Carregar um balancete já no formato BHub "
                                        "pode gerar um De-Para incorreto.\n\n"
                                        "**Deseja continuar mesmo assim ou cancelar e carregar outro arquivo?**"
                                    )
                                    _bc1, _bc2 = st.columns(2)
                                    with _bc1:
                                        if st.button(
                                            "✅ Confirmar — usar este balancete",
                                            key=f"bhub_confirm_{company_label}",
                                            type="primary",
                                            use_container_width=True,
                                        ):
                                            st.session_state["bhub_confirmed"].add(company_label)
                                            st.rerun()
                                    with _bc2:
                                        if st.button(
                                            "❌ Cancelar — carregar outro arquivo",
                                            key=f"bhub_cancel_{company_label}",
                                            use_container_width=True,
                                        ):
                                            # Incrementa a chave do uploader → cria widget novo e vazio
                                            st.session_state["uploader_key"] = (
                                                st.session_state.get("uploader_key", 0) + 1
                                            )
                                            # Limpa todo o estado relacionado aos PDFs
                                            for _k in ["df_all", "all_dfs", "all_records",
                                                       "bhub_confirmed", "bhub_cancelled",
                                                       "manual_accounts", "excluded_accounts",
                                                       "ai_bhub_suggestions", "ai_provider_used"]:
                                                st.session_state.pop(_k, None)
                                            st.rerun()
                                else:
                                    # Estrutura normal — adicionar sem confirmação
                                    all_dfs[company_label] = df_filtered
                                    st.success(
                                        f"✅ {total} contas analíticas | "
                                        f"Ativo: R$ {format_br(tot_ativo)} | "
                                        f"Passivo/PL: R$ {format_br(abs(tot_passivo))} | "
                                        f"Resultado: R$ {format_br(tot_result)}"
                                    )

                            else:
                                # Confirmado pelo usuário
                                all_dfs[company_label] = df_filtered
                                st.success(
                                    f"✅ {total} contas analíticas (confirmado pelo usuário) | "
                                    f"Ativo: R$ {format_br(tot_ativo)} | "
                                    f"Passivo/PL: R$ {format_br(abs(tot_passivo))} | "
                                    f"Resultado: R$ {format_br(tot_result)}"
                                )

                        except Exception as e:
                            st.error(f"Erro ao processar {pdf.name}: {e}")

            if all_dfs:
                # Mesclar contas adicionadas manualmente
                for acc in st.session_state.get("manual_accounts", []):
                    emp = acc.get("empresa", "")
                    row_df = pd.DataFrame([acc])
                    if emp in all_dfs:
                        all_dfs[emp] = pd.concat([all_dfs[emp], row_df], ignore_index=True)
                    else:
                        # Empresa digitada manualmente que não veio de PDF
                        all_dfs[emp] = row_df

                df_all = pd.concat(all_dfs.values(), ignore_index=True)

                # Reaplicar exclusões persistidas (sobrevivem ao reprocessamento de PDF)
                excl = st.session_state.get("excluded_accounts", set())
                if excl:
                    def _keep(r):
                        return (str(r.get("code", "")), str(r.get("empresa", ""))) not in excl
                    df_all = df_all[df_all.apply(_keep, axis=1)].reset_index(drop=True)
                    for _emp in list(all_dfs.keys()):
                        all_dfs[_emp] = all_dfs[_emp][
                            all_dfs[_emp].apply(_keep, axis=1)
                        ].reset_index(drop=True)

                st.session_state["df_all"] = df_all
                st.session_state["all_dfs"] = all_dfs

                n_manual = len(st.session_state["manual_accounts"])
                manual_tag = f" (+{n_manual} manual)" if n_manual else ""
                st.success(
                    f"**✅ Total geral: {len(df_all)} contas{manual_tag} | "
                    f"{len(all_dfs)} empresa(s) carregadas**"
                )

            # ── Adicionar Conta Ausente ─────────────────────────────────
            if all_dfs:
                st.divider()
                st.subheader("➕ Adicionar Conta Ausente")
                st.caption(
                    "Use este formulário para incluir manualmente uma conta analítica "
                    "que não foi extraída do PDF. A conta será incorporada ao processo "
                    "automaticamente (De-Para, planilha e TXT do Domínio)."
                )

                # Listar contas já adicionadas – com botão de exclusão por linha
                if st.session_state["manual_accounts"]:
                    st.caption(
                        f"**{len(st.session_state['manual_accounts'])} "
                        "conta(s) adicionadas manualmente** (clique ✕ para remover):"
                    )
                    for i, acc in enumerate(list(st.session_state["manual_accounts"])):
                        c_info, c_del = st.columns([11, 1])
                        with c_info:
                            st.markdown(
                                f"`{acc['code']}` · {acc.get('classification','')} · "
                                f"**{acc.get('description','')}** · "
                                f"R$ {format_br(acc.get('current_value', 0))} "
                                f"{acc.get('current_indicator','')} · _{acc.get('empresa','')}_"
                            )
                        with c_del:
                            if st.button("✕", key=f"del_manual_{i}", help="Remover esta conta"):
                                removed = st.session_state["manual_accounts"].pop(i)
                                # Remove diretamente de df_all e all_dfs
                                _code = str(removed.get("code", ""))
                                _emp  = removed.get("empresa", "")
                                if "df_all" in st.session_state:
                                    dfa = st.session_state["df_all"]
                                    mask = ~(
                                        (dfa["code"].astype(str) == _code) &
                                        (dfa["empresa"] == _emp)
                                    )
                                    st.session_state["df_all"] = dfa[mask].reset_index(drop=True)
                                if "all_dfs" in st.session_state and _emp in st.session_state["all_dfs"]:
                                    dfe = st.session_state["all_dfs"][_emp]
                                    mask = ~(
                                        (dfe["code"].astype(str) == _code) &
                                        (dfe["empresa"] == _emp)
                                    )
                                    st.session_state["all_dfs"][_emp] = dfe[mask].reset_index(drop=True)
                                st.rerun()

                    if st.button("🗑️ Limpar todas as contas manuais", type="secondary"):
                        removed_keys = {
                            (str(a["code"]), a.get("empresa", ""))
                            for a in st.session_state["manual_accounts"]
                        }
                        if "df_all" in st.session_state:
                            dfa = st.session_state["df_all"]
                            st.session_state["df_all"] = dfa[
                                ~dfa.apply(
                                    lambda r: (str(r["code"]), r.get("empresa", "")) in removed_keys,
                                    axis=1,
                                )
                            ].reset_index(drop=True)
                        if "all_dfs" in st.session_state:
                            for _emp, dfe in st.session_state["all_dfs"].items():
                                st.session_state["all_dfs"][_emp] = dfe[
                                    ~dfe.apply(
                                        lambda r: (str(r["code"]), r.get("empresa", "")) in removed_keys,
                                        axis=1,
                                    )
                                ].reset_index(drop=True)
                        st.session_state["manual_accounts"] = []
                        st.rerun()

                # Mensagem de sucesso após confirmação (persiste 1 ciclo via session_state)
                if st.session_state.pop("conta_adicionada_ok", None):
                    la = st.session_state.get("last_added_acc", {})
                    st.success(
                        f"✅ Conta **{la.get('code','')} – {la.get('description','')}** "
                        f"adicionada! Saldo: R$ {format_br(la.get('current_value',0))} "
                        f"{la.get('current_indicator','')}"
                    )

                empresas_disponiveis = list(all_dfs.keys())

                # ── Modo de adição ─────────────────────────────────────────
                _modo_adicao = st.radio(
                    "Como deseja adicionar a conta?",
                    ["🔍 Buscar no PDF", "✏️ Inserir manualmente"],
                    horizontal=True,
                    key="modo_adicao_conta",
                )
                # Limpa resultado de busca ao trocar de modo
                if _modo_adicao == "✏️ Inserir manualmente":
                    st.session_state.pop("found_account", None)

                if _modo_adicao == "🔍 Buscar no PDF":
                    # ── Passo 1-A: Buscar no arquivo ───────────────────────
                    with st.form("form_busca_conta", clear_on_submit=False):
                        st.markdown("**Informe o Código ou a Classificação da conta ausente:**")
                        col_emp, col_busca = st.columns([2, 3])
                        with col_emp:
                            nova_empresa = st.selectbox("Empresa", empresas_disponiveis)
                        with col_busca:
                            busca = st.text_input(
                                "Código ou Classificação",
                                placeholder="10028   ou   4.2.2.05.000013",
                            )
                        buscar_btn = st.form_submit_button(
                            "🔍 Buscar no arquivo", use_container_width=True
                        )

                    if buscar_btn:
                        termo = busca.strip()
                        if not termo:
                            st.error("Informe o código ou a classificação.")
                        else:
                            pool = st.session_state.get("all_records", {}).get(nova_empresa, [])
                            encontrado = next(
                                (r for r in pool
                                 if r.get("code", "") == termo
                                 or r.get("classification", "") == termo),
                                None,
                            )
                            if encontrado is None:
                                st.session_state.pop("found_account", None)
                                st.error(
                                    f"Conta **'{termo}'** não encontrada nos registros extraídos "
                                    f"do PDF de **{nova_empresa}**. "
                                    "Verifique se o PDF foi carregado e se o código/classificação está correto."
                                )
                            else:
                                st.session_state["found_account"] = {
                                    **encontrado, "empresa": nova_empresa
                                }

                else:
                    # ── Passo 1-B: Inserir manualmente ─────────────────────
                    with st.form("form_manual_conta", clear_on_submit=True):
                        st.markdown("**Preencha os dados da conta** _(* obrigatório)_:")
                        _mc1, _mc2, _mc3 = st.columns([2, 1, 1])
                        with _mc1:
                            _nova_emp_m = st.selectbox("Empresa *", empresas_disponiveis, key="sel_emp_manual")
                        with _mc2:
                            _cod_m = st.text_input("Código *", placeholder="ex: 999")
                        with _mc3:
                            _dc_m = st.selectbox("D/C *", ["D — Devedor", "C — Credor"])

                        _mc4, _mc5 = st.columns([3, 1])
                        with _mc4:
                            _desc_m = st.text_input("Descrição *", placeholder="ex: CONTA TRANSITÓRIA ESPECIAL")
                        with _mc5:
                            _val_m = st.number_input("Valor (R$) *", min_value=0.0, step=0.01, format="%.2f")

                        _cls_m = st.text_input(
                            "Classificação (opcional)",
                            placeholder="ex: 1.1.9.01.000001 — deixe vazio se não souber",
                        )
                        _submit_manual = st.form_submit_button(
                            "✅ Adicionar Conta", type="primary", use_container_width=True
                        )

                    if _submit_manual:
                        _cod_m_v   = _cod_m.strip()
                        _desc_m_v  = _desc_m.strip()
                        _cls_m_v   = _cls_m.strip()
                        _dc_final  = "D" if _dc_m.startswith("D") else "C"

                        if not _cod_m_v or not _desc_m_v or _val_m == 0.0:
                            st.error("Código, Descrição e Valor são obrigatórios e Valor deve ser > 0.")
                        else:
                            # Auto-derivar grupo pela classificação ou D/C
                            if _cls_m_v:
                                _grupo_m = get_account_group(_cls_m_v)
                            else:
                                _grupo_m = "Ativo" if _dc_final == "D" else "Passivo/PL"

                            _new_acc_m = {
                                "code":               _cod_m_v,
                                "classification":     _cls_m_v,
                                "description":        _desc_m_v,
                                "debit_mv":           _val_m if _dc_final == "D" else 0.0,
                                "credit_mv":          _val_m if _dc_final == "C" else 0.0,
                                "current_value":      _val_m,
                                "current_indicator":  _dc_final,
                                "tipo":               "A",
                                "grupo":              _grupo_m,
                                "empresa":            _nova_emp_m,
                            }

                            _ja_existe_m = any(
                                a["code"] == _cod_m_v and a["empresa"] == _nova_emp_m
                                for a in st.session_state["manual_accounts"]
                            )
                            if _ja_existe_m:
                                st.warning(f"A conta {_cod_m_v} já foi adicionada para {_nova_emp_m}.")
                            else:
                                st.session_state["manual_accounts"].append(_new_acc_m)
                                _new_row_m = pd.DataFrame([_new_acc_m])
                                if "df_all" in st.session_state:
                                    st.session_state["df_all"] = pd.concat(
                                        [st.session_state["df_all"], _new_row_m], ignore_index=True
                                    )
                                if "all_dfs" in st.session_state:
                                    _adfs = st.session_state["all_dfs"]
                                    if _nova_emp_m in _adfs:
                                        _adfs[_nova_emp_m] = pd.concat(
                                            [_adfs[_nova_emp_m], _new_row_m], ignore_index=True
                                        )
                                    else:
                                        _adfs[_nova_emp_m] = _new_row_m.copy()
                                st.session_state["last_added_acc"]    = _new_acc_m
                                st.session_state["conta_adicionada_ok"] = True
                                st.rerun()

                # ── Passo 2: Confirmar ─────────────────────────────────────
                if "found_account" in st.session_state:
                    fa = st.session_state["found_account"]
                    st.info(
                        f"**Conta encontrada no PDF:**  \n"
                        f"**Código:** {fa.get('code','')}  |  "
                        f"**Classificação:** {fa.get('classification','')}  |  "
                        f"**Descrição:** {fa.get('description','')}  |  "
                        f"**Saldo:** R$ {format_br(fa.get('current_value',0))} "
                        f"{fa.get('current_indicator','')}"
                    )
                    col_ok, col_cancel = st.columns([2, 1])
                    with col_ok:
                        if st.button("✅ Confirmar e Incluir nos Lançamentos", type="primary", use_container_width=True):
                            ja_existe = any(
                                a["code"] == fa["code"] and a["empresa"] == fa["empresa"]
                                for a in st.session_state["manual_accounts"]
                            )
                            if ja_existe:
                                st.warning(f"A conta {fa['code']} já foi adicionada.")
                            else:
                                new_acc = {
                                    "code": fa["code"],
                                    "classification": fa["classification"],
                                    "description": fa["description"],
                                    "debit_mv": fa.get("debit_mv", 0.0),
                                    "credit_mv": fa.get("credit_mv", 0.0),
                                    "current_value": fa["current_value"],
                                    "current_indicator": fa["current_indicator"],
                                    "tipo": "A",
                                    "grupo": get_account_group(fa["classification"]),
                                    "empresa": fa["empresa"],
                                }
                                st.session_state["manual_accounts"].append(new_acc)

                                # Atualiza df_all imediatamente — sem aguardar reprocessamento de PDF
                                new_row_df = pd.DataFrame([new_acc])
                                if "df_all" in st.session_state:
                                    st.session_state["df_all"] = pd.concat(
                                        [st.session_state["df_all"], new_row_df],
                                        ignore_index=True,
                                    )
                                if "all_dfs" in st.session_state:
                                    _emp = new_acc["empresa"]
                                    adfs = st.session_state["all_dfs"]
                                    if _emp in adfs:
                                        adfs[_emp] = pd.concat(
                                            [adfs[_emp], new_row_df], ignore_index=True
                                        )
                                    else:
                                        adfs[_emp] = new_row_df.copy()

                                st.session_state["last_added_acc"] = new_acc
                                st.session_state["conta_adicionada_ok"] = True
                                st.session_state.pop("found_account", None)
                                st.rerun()
                    with col_cancel:
                        if st.button("✖ Cancelar", use_container_width=True):
                            st.session_state.pop("found_account", None)
                            st.rerun()

                # ── Visualização de lançamentos (atualizada) ───────────────
                if st.session_state.get("df_all") is not None and not st.session_state["df_all"].empty:

                    # Mensagem de confirmação após exclusão
                    n_excl = st.session_state.pop("_excluidas_count", 0)
                    if n_excl:
                        st.success(f"✅ {n_excl} conta(s) excluída(s) da lista de lançamentos.")

                    with st.expander("📋 Ver todos os lançamentos (incluindo adições manuais)", expanded=False):
                        # Sempre recalcula a partir do all_dfs local (recém-populado neste ciclo),
                        # evitando que session_state desatualizado (de um rerun com falha) omita empresas.
                        _all_dfs_local = all_dfs if all_dfs else st.session_state.get("all_dfs", {})
                        if _all_dfs_local:
                            df_viz = pd.concat(_all_dfs_local.values(), ignore_index=True)
                        else:
                            df_viz = st.session_state["df_all"]
                        df_viz_nonzero = df_viz[df_viz["current_value"] != 0].copy()
                        df_viz_nonzero["Data"] = opening_date.strftime("%d/%m/%Y")

                        _empresas_viz = df_viz_nonzero["empresa"].unique().tolist()

                        # Expander por empresa quando há múltiplos arquivos
                        if len(_empresas_viz) > 1:
                            for _emp in _empresas_viz:
                                _df_emp = df_viz_nonzero[df_viz_nonzero["empresa"] == _emp]
                                _tot_emp = signed_total(_df_emp)
                                with st.expander(
                                    f"📄 {_emp} — {len(_df_emp)} contas | "
                                    f"Total líq.: R$ {format_br(abs(_tot_emp))}",
                                    expanded=True,
                                ):
                                    st.dataframe(
                                        _df_emp[[
                                            "code", "classification", "description",
                                            "current_value", "current_indicator",
                                        ]].rename(columns={
                                            "code": "Código",
                                            "classification": "Classificação",
                                            "description": "Descrição",
                                            "current_value": "Saldo",
                                            "current_indicator": "D/C",
                                        }).reset_index(drop=True),
                                        use_container_width=True,
                                        height=min(40 * len(_df_emp) + 38, 420),
                                        hide_index=True,
                                    )
                            st.divider()
                            st.caption("Tabela consolidada — marque para excluir contas:")

                        df_display = df_viz_nonzero[[
                            "Data", "empresa", "code", "classification",
                            "description", "current_value", "current_indicator",
                        ]].rename(columns={
                            "Data": "Data",
                            "empresa": "Empresa",
                            "code": "Código",
                            "classification": "Classificação",
                            "description": "Descrição",
                            "current_value": "Saldo",
                            "current_indicator": "D/C",
                        }).reset_index(drop=True).copy()
                        df_display.insert(0, "🗑️", False)

                        df_edited = st.data_editor(
                            df_display,
                            use_container_width=True,
                            height=370,
                            hide_index=True,
                            key="viz_editor",
                            column_config={
                                "🗑️": st.column_config.CheckboxColumn(
                                    "🗑️", width="small",
                                    help="Marque e clique em 'Excluir' para remover da lista",
                                ),
                                "Saldo": st.column_config.NumberColumn("Saldo", format="R$ %.2f"),
                                "Data": st.column_config.TextColumn("Data", width="small"),
                                "Empresa": st.column_config.TextColumn("Empresa", width="medium"),
                                "Código": st.column_config.TextColumn("Código", width="small"),
                                "Classificação": st.column_config.TextColumn("Classificação", width="medium"),
                                "Descrição": st.column_config.TextColumn("Descrição", width="large"),
                                "D/C": st.column_config.TextColumn("D/C", width="small"),
                            },
                            disabled=["Data", "Empresa", "Código", "Classificação", "Descrição", "Saldo", "D/C"],
                        )

                        # Salva o DataFrame completo para o callback ler antes do próximo rerun
                        st.session_state["_viz_edited_df"] = df_edited

                        n_checked = int(df_edited["🗑️"].sum())
                        if n_checked > 0:
                            st.button(
                                f"🗑️ Excluir {n_checked} conta(s) marcada(s)",
                                type="secondary",
                                on_click=_excluir_marcadas_cb,
                                key="btn_excluir_marcadas",
                            )

                        total_signed = signed_total(df_viz_nonzero)
                        total_label = f"R$ {format_br(abs(total_signed))}" if total_signed != 0 else "R$ 0,00 ✅"
                        st.caption(
                            f"{len(df_viz_nonzero)} lançamentos | "
                            f"Total líquido (D−C): {total_label}"
                            + (" ⚠️ diferença" if total_signed != 0 else "")
                        )

    # ════════════════════════════════════════════════════════════════════════
    # TAB 2 – DE-PARA
    # ════════════════════════════════════════════════════════════════════════
    with tab2:
        st.header("De-Para: Plano Antigo → Plano BHub (ECD Ficha I057)")

        if "df_all" not in st.session_state:
            st.info("ℹ️ Carregue os balancetes na aba anterior primeiro.")
        else:
            # Reconstrói df a partir do dict por empresa para garantir que a
            # coluna 'empresa' de cada linha corresponda ao arquivo de origem correto,
            # mesmo que o session_state["df_all"] tenha sido gravado antes de alguma
            # correção de rótulo (ex: dois PDFs com o mesmo nome de arquivo).
            _all_dfs_t2 = st.session_state.get("all_dfs", {})
            if _all_dfs_t2:
                _dfs_labeled = []
                for _emp_key, _df_val in _all_dfs_t2.items():
                    _df_copy = _df_val.copy()
                    _df_copy["empresa"] = _emp_key  # garante rótulo = chave do dict
                    _dfs_labeled.append(_df_copy)
                df = pd.concat(_dfs_labeled, ignore_index=True)
            else:
                df = st.session_state["df_all"]

            # ── Sugestão de Contas BHub ────────────────────────────────────
            st.subheader("🤖 Sugestão de Contas BHub")

            # Lê sugestões IA já calculadas (se houver)
            ai_sugs: dict = st.session_state.get("ai_bhub_suggestions", {})

            # ── Configuração de IA ─────────────────────────────────────────
            with st.expander("⚙️ Configurar IA para sugestão de contas", expanded=not bool(ai_sugs)):
                _PROVIDERS = ["Nenhum (similaridade automática)", "Claude", "ChatGPT", "Gemini"]
                _provider_choice = st.radio(
                    "Provedor de IA",
                    options=_PROVIDERS,
                    index=_PROVIDERS.index(
                        st.session_state.get("ai_provider_choice", "Nenhum (similaridade automática)")
                    ),
                    horizontal=True,
                    key="ai_provider_radio",
                )
                st.session_state["ai_provider_choice"] = _provider_choice

                _ai_key      = ""
                _ai_model    = ""
                _ai_provider = ""
                _ai_label    = _provider_choice

                if _provider_choice.startswith("Nenhum"):
                    st.info(
                        "Sugestão automática por similaridade de texto (descrição, código e classificação). "
                        "Selecione um provedor de IA acima para mapeamento mais preciso."
                    )

                elif _provider_choice == "Claude":
                    _anthropic_key = _load_anthropic_key()
                    # Chave real tem >50 chars e não contém placeholder
                    _key_is_real = bool(_anthropic_key) and len(_anthropic_key) > 50 and "COLOQUE" not in _anthropic_key
                    if _key_is_real:
                        st.info("🔒 Chave Claude carregada via **secrets.toml** / variável de ambiente.")
                        _ai_key = _anthropic_key
                    else:
                        if _anthropic_key and not _key_is_real:
                            st.warning("⚠️ Chave no secrets.toml parece ser um placeholder. Informe a chave real abaixo.")
                        _ai_key = st.text_input(
                            "Chave API Claude (Anthropic)",
                            type="password",
                            key="claude_api_key_input",
                            placeholder="sk-ant-api03-...",
                            help="Obtenha em console.anthropic.com",
                        )
                    _ai_provider = "anthropic"
                    _ai_model    = "claude-3-5-haiku-20241022"

                elif _provider_choice == "ChatGPT":
                    _ai_key = st.text_input(
                        "Chave API OpenAI",
                        type="password",
                        key="openai_api_key_input",
                        value=st.session_state.get("openai_api_key", ""),
                        placeholder="sk-...",
                        help="Obtenha em platform.openai.com",
                    )
                    if _ai_key:
                        st.session_state["openai_api_key"] = _ai_key
                    _ai_model = st.selectbox(
                        "Modelo ChatGPT",
                        options=["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"],
                        index=0,
                        key="openai_model_select",
                    )
                    _ai_provider = "openai"

                elif _provider_choice == "Gemini":
                    _ai_key = st.text_input(
                        "Chave API Gemini (Google)",
                        type="password",
                        key="gemini_api_key_input",
                        value=st.session_state.get("gemini_api_key", ""),
                        placeholder="AIza...",
                        help="Obtenha em aistudio.google.com",
                    )
                    if _ai_key:
                        st.session_state["gemini_api_key"] = _ai_key
                    _gemini_models_list = [
                        "gemini-2.0-flash-lite", "gemini-1.5-flash",
                        "gemini-2.0-flash", "gemini-1.5-pro",
                    ]
                    if _ai_key:
                        try:
                            _dyn = _list_gemini_models(_ai_key)
                            if _dyn:
                                _gemini_models_list = _dyn
                        except Exception:
                            pass
                    _ai_model = st.selectbox(
                        "Modelo Gemini",
                        options=_gemini_models_list,
                        index=0,
                        key="gemini_model_select",
                    )
                    _ai_provider = "gemini"

            _api_key_ok = bool(_ai_key)

            # ── Status e botão de sugestão ─────────────────────────────────
            _col_info, _col_btn = st.columns([4, 1])
            with _col_info:
                if ai_sugs:
                    _n_mapped = sum(1 for v in ai_sugs.values() if v.get("bhub_name"))
                    _src = st.session_state.get("ai_provider_used", "IA")
                    st.success(
                        f"✅ Sugestões por {_src} aplicadas — {_n_mapped} de {len(df)} contas mapeadas. "
                        "Edite os campos `Cód. BHub` conforme necessário."
                    )
                elif _api_key_ok:
                    st.info(
                        f"Clique em **Sugerir com {_ai_label}** para mapear automaticamente as contas. "
                        "A IA usa apenas o plano BHub oficial e sugere contas genéricas "
                        "(Clientes/Fornecedores Nacionais) quando não há match exato."
                    )
                else:
                    n_ativo  = sum(1 for a in BHUB_ACCOUNTS if a["grupo"] == "Ativo")
                    n_pass   = sum(1 for a in BHUB_ACCOUNTS if a["grupo"] == "Passivo/PL")
                    n_result = sum(1 for a in BHUB_ACCOUNTS if a["grupo"] == "Resultado")
                    if _provider_choice.startswith("Nenhum"):
                        st.info(
                            f"Plano BHub: **{len(BHUB_ACCOUNTS)} contas** (Ativo: {n_ativo} | "
                            f"Passivo/PL: {n_pass} | Resultado: {n_result}). "
                            "Sugestões por similaridade automática ativas."
                        )
                    else:
                        st.info(
                            f"Plano BHub: **{len(BHUB_ACCOUNTS)} contas** (Ativo: {n_ativo} | "
                            f"Passivo/PL: {n_pass} | Resultado: {n_result}). "
                            "Informe a chave API acima para ativar sugestões por IA."
                        )
            with _col_btn:
                if ai_sugs:
                    if st.button("🗑️ Limpar IA", key="btn_clear_ai",
                                 help="Remove as sugestões de IA e volta para sugestão automática"):
                        del st.session_state["ai_bhub_suggestions"]
                        st.session_state.pop("ai_provider_used", None)
                        st.rerun()
                elif _api_key_ok:
                    if st.button(f"🤖 Sugerir com {_ai_label}", key="btn_suggest_ai",
                                 type="primary", use_container_width=True):
                        with st.spinner(f"Consultando {_ai_label}... aguarde alguns segundos."):
                            try:
                                _result = suggest_bhub_ai(df, _ai_key, provider=_ai_provider,
                                                          model=_ai_model)
                                st.session_state["ai_bhub_suggestions"] = _result
                                st.session_state["ai_provider_used"] = _ai_label
                                st.rerun()
                            except ImportError as _e:
                                st.error(str(_e))
                            except Exception as _e:
                                st.error(f"Erro ao consultar {_ai_label}: {_e}")

            # Calcular sugestões para todas as contas
            # (usa IA quando disponível, fallback para rule-based)
            sug_rows = []
            for _pos, (_, row) in enumerate(df.iterrows()):
                grupo   = row.get("grupo", "Outros")
                old_cls = str(row.get("classification", ""))
                if _pos in ai_sugs:
                    _ai = ai_sugs[_pos]
                    top = {
                        "code":  _ai.get("bhub_code", ""),
                        "name":  _ai.get("bhub_name", ""),
                        "mask":  _ai.get("bhub_mask", ""),
                        "score": 100.0 if _ai.get("bhub_name") else 0.0,
                    }
                else:
                    sugs = suggest_bhub(str(row["description"]), grupo, old_cls=old_cls)
                    top  = sugs[0] if sugs else {"code": "", "name": "", "mask": "", "score": 0.0}
                sug_rows.append({
                    "Grupo": grupo,
                    "Cód. Antigo": str(row.get("code", "")),
                    "Classif. Antiga": old_cls,
                    "Descrição Antiga": str(row.get("description", "")),
                    "Cód. BHub": top.get("code", ""),
                    "Classif. BHub": top.get("mask", ""),
                    "Nome BHub Sugerido": top.get("name", ""),
                    "Match %": top.get("score", 0.0),
                    "D/C": str(row.get("current_indicator", "")),
                    "Saldo": float(row.get("current_value", 0)),
                    "Empresa": str(row.get("empresa", "")),
                })

            df_sug = pd.DataFrame(sug_rows)

            # ── Aplicar overrides manuais persistidos ─────────────────────
            _ovr_key = "depara_manual_overrides"
            if _ovr_key not in st.session_state:
                st.session_state[_ovr_key] = {}
            for _ridx, _ov in st.session_state[_ovr_key].items():
                if _ridx in df_sug.index:
                    df_sug.loc[_ridx, "Cód. BHub"]          = _ov["code"]
                    df_sug.loc[_ridx, "Classif. BHub"]       = _ov["mask"]
                    df_sug.loc[_ridx, "Nome BHub Sugerido"]  = _ov["name"]
                    df_sug.loc[_ridx, "Match %"]              = 100.0

            # Persiste o De-Para final para uso na Carta de Responsabilidade (Tab 4)
            st.session_state["df_sug_current"] = df_sug.copy()

            _fonte = "IA (Claude)" if ai_sugs else "Automática (texto)"
            st.markdown(
                f"**Fonte das sugestões: {_fonte}.** "
                "Códigos em **:red[vermelho]** são sugestões automáticas. "
                "Use **✏️ Editar mapeamento** abaixo para ajustar."
            )

            _empresas_sug = df_sug["Empresa"].unique().tolist()
            _multi_empresa = len(_empresas_sug) > 1

            # Índices com override manual (para estilo verde)
            _ovr_indices = set(st.session_state[_ovr_key].keys())

            def _style_bhub_col(col):
                """Vermelho = sugestão automática, verde = confirmado manualmente, cinza = vazio."""
                out = []
                for idx, val in col.items():
                    if str(val) == "":
                        out.append("color: #888888")
                    elif idx in _ovr_indices:
                        out.append("color: #1a7f3c; font-weight: bold")
                    else:
                        out.append("color: #cc0000; font-weight: bold")
                return out

            _cols_show = [
                "Nome do Arquivo", "Grupo", "Cód. Antigo", "Classif. Antiga",
                "Descrição Antiga", "D/C", "Saldo",
                "Cód. BHub", "Classif. BHub", "Nome BHub Sugerido", "Match %",
            ]

            if _multi_empresa:
                for _emp_s in _empresas_sug:
                    _df_e = df_sug[df_sug["Empresa"] == _emp_s].rename(
                        columns={"Empresa": "Nome do Arquivo"}
                    )
                    _n_mapeado = (_df_e["Cód. BHub"] != "").sum()
                    with st.expander(
                        f"📄 {_emp_s} — {len(_df_e)} contas | Mapeadas: {_n_mapeado} / {len(_df_e)}",
                        expanded=True,
                    ):
                        st.dataframe(
                            _df_e[_cols_show].style.format({"Saldo": "{:.2f}", "Match %": "{:.2f}"}).apply(_style_bhub_col, subset=["Cód. BHub"]),
                            use_container_width=True,
                            height=min(40 * len(_df_e) + 38, 500),
                            hide_index=True,
                        )
            else:
                _df_single = df_sug.rename(columns={"Empresa": "Nome do Arquivo"})
                st.dataframe(
                    _df_single[_cols_show].style.format({"Saldo": "{:.2f}", "Match %": "{:.2f}"}).apply(_style_bhub_col, subset=["Cód. BHub"]),
                    use_container_width=True,
                    height=400,
                    hide_index=True,
                )

            # ── Painel de edição de mapeamento ────────────────────────────
            with st.expander("✏️ Editar mapeamento de conta", expanded=False):
                # Seletor de qual conta editar
                _acct_opts = [
                    f"{row['Cód. Antigo']} — {str(row['Descrição Antiga'])[:60]}"
                    + (f"  [{row['Empresa']}]" if _multi_empresa else "")
                    for _, row in df_sug.iterrows()
                ]
                _acct_sel = st.selectbox(
                    "Conta a editar:",
                    _acct_opts,
                    key="depara_edit_acct_sel",
                )
                _acct_pos  = _acct_opts.index(_acct_sel)
                _acct_idx  = df_sug.index[_acct_pos]
                _acct_row  = df_sug.loc[_acct_idx]
                _is_manual = _acct_idx in _ovr_indices

                _status_lbl = "confirmado manualmente ✅" if _is_manual else "sugestão automática 🔴"
                st.info(
                    f"**Mapeamento atual ({_status_lbl}):**  \n"
                    f"Código: `{_acct_row['Cód. BHub'] or '—'}` | "
                    f"Classificação: `{_acct_row['Classif. BHub'] or '—'}` | "
                    f"Nome: {_acct_row['Nome BHub Sugerido'] or '—'}"
                )

                st.markdown("**Buscar no plano de contas BHub:**")
                _bcol1, _bcol2 = st.columns([4, 1])
                with _bcol1:
                    _busca_txt = st.text_input(
                        "Termo:",
                        key="depara_busca_txt",
                        placeholder="Ex: banco, honorarios, 4046, 4.1.01...",
                        label_visibility="collapsed",
                    )
                with _bcol2:
                    _busca_modo = st.selectbox(
                        "Por:",
                        ["Descrição", "Código", "Classificação"],
                        key="depara_busca_modo",
                        label_visibility="collapsed",
                    )

                if _busca_txt.strip():
                    _bterm = _busca_txt.strip().upper()
                    if _busca_modo == "Código":
                        _res = [a for a in BHUB_ACCOUNTS if _bterm in str(a["code"]).upper()]
                    elif _busca_modo == "Classificação":
                        _res = [a for a in BHUB_ACCOUNTS if _bterm in str(a["mask"]).upper()]
                    else:
                        _res = [a for a in BHUB_ACCOUNTS if _bterm in str(a["name"]).upper()]
                    _res = _res[:200]

                    if _res:
                        _df_res = pd.DataFrame([
                            {"Código": a["code"], "Classificação": a["mask"],
                             "Nome": a["name"], "Grupo": a["grupo"]}
                            for a in _res
                        ])
                        st.caption(f"{len(_res)} conta(s) encontrada(s):")
                        st.dataframe(
                            _df_res, use_container_width=True,
                            hide_index=True,
                            height=min(40 * len(_df_res) + 38, 280),
                        )
                        _res_opts = [
                            f"{a['code']} | {a['mask']} | {a['name']}"
                            for a in _res
                        ]
                        _res_sel = st.selectbox(
                            "Selecionar conta para aplicar:",
                            _res_opts,
                            key="depara_res_sel",
                        )
                        _conta_nova = _res[_res_opts.index(_res_sel)]

                        if st.button("✅ Aplicar alteração", key="depara_aplicar_btn", type="primary"):
                            st.session_state[_ovr_key][_acct_idx] = {
                                "code": _conta_nova["code"],
                                "mask": _conta_nova["mask"],
                                "name": _conta_nova["name"],
                            }
                            st.rerun()
                    else:
                        st.warning("Nenhuma conta encontrada. Tente outros termos.")
                else:
                    st.caption("Digite um termo para buscar (ex: banco, honorários, 4046).")

                if _is_manual:
                    if st.button("🔄 Restaurar sugestão automática", key="depara_clear_ovr"):
                        st.session_state[_ovr_key].pop(_acct_idx, None)
                        st.rerun()

            n_com_bhub = (df_sug["Cód. BHub"] != "").sum()
            n_sem_bhub = len(df_sug) - n_com_bhub
            c1, c2, c3 = st.columns(3)
            c1.metric("Total contas", len(df_sug))
            c2.metric("Com código BHub", n_com_bhub)
            c3.metric("Pendente (manual)", n_sem_bhub)

            col_a, col_b = st.columns(2)
            with col_a:
                # Montar df para De-Para com sugestões preenchidas
                df_depara_out = df.copy()
                df_depara_out["bhub_code"] = df_sug["Cód. BHub"].values
                df_depara_out["bhub_mask"] = df_sug["Classif. BHub"].values
                buf_sug = generate_depara_excel(df_depara_out)
                st.download_button(
                    label="⬇️  Baixar De-Para com Sugestões (.xlsx)",
                    data=buf_sug,
                    file_name="depara_ECD_ficha_I057.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            with col_b:
                # De-Para vazio (campo BHub em branco)
                buf_vazio = generate_depara_excel(df)
                st.download_button(
                    label="⬇️  Baixar De-Para Vazio (preencher manualmente)",
                    data=buf_vazio,
                    file_name="depara_ECD_ficha_I057_vazio.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            # ── Preview de lançamentos com códigos BHub ────────────────────
            st.divider()
            st.subheader("📋 Prévia dos lançamentos a serem importados")

            df_prev_bhub = df_sug[df_sug["Cód. BHub"].notna() & (df_sug["Cód. BHub"] != "")].copy()
            if df_prev_bhub.empty:
                st.info("Preencha pelo menos um Cód. BHub na tabela acima para visualizar os lançamentos.")
            else:
                data_str = opening_date.strftime("%d/%m/%Y")
                # Lê o modo selecionado no Tab 3 (valor persistido em session_state)
                _modo_val = st.session_state.get("modo_lancto", "Partidas dobradas (1 lançamento único)")
                _modo_prev = "dobradas" if "Partidas" in str(_modo_val) else "simples"

                preview_rows = []
                _first_prev = True
                _prev_emp = None
                for _, r in df_prev_bhub.iterrows():
                    bhub_code = r["Cód. BHub"]
                    ind = r["D/C"]
                    val = r["Saldo"]
                    desc = r["Descrição Antiga"][:50]
                    hist = f"IMPLANTAÇÃO DE SALDO EM {data_str} - {desc}"
                    emp_r = r.get("Empresa", "")
                    # Reseta o flag de lote quando troca de empresa
                    if emp_r != _prev_emp:
                        _first_prev = True
                        _prev_emp = emp_r
                    if _modo_prev == "dobradas":
                        lote_v = 1 if _first_prev else ""
                        _first_prev = False
                        preview_rows.append({
                            "Empresa": emp_r,
                            "Data": data_str,
                            "Débito": bhub_code if ind == "D" else "",
                            "Crédito": bhub_code if ind == "C" else "",
                            "Valor": val,
                            "Histórico": hist,
                            "Lote": lote_v,
                        })
                    else:
                        if ind == "D":
                            preview_rows.append({
                                "Empresa": emp_r,
                                "Data": data_str, "Débito": bhub_code,
                                "Crédito": contra_code, "Valor": val,
                                "Histórico": hist, "Lote": 1,
                            })
                        else:
                            preview_rows.append({
                                "Empresa": emp_r,
                                "Data": data_str, "Débito": contra_code,
                                "Crédito": bhub_code, "Valor": val,
                                "Histórico": hist, "Lote": 1,
                            })

                df_launch_prev = pd.DataFrame(preview_rows)
                modo_label_prev = (
                    "Partidas dobradas (1 lançamento único)"
                    if _modo_prev == "dobradas"
                    else "Lançamentos simples com contrapartida"
                )

                _col_cfg_prev = {"Valor": st.column_config.NumberColumn("Valor", format="R$ %.2f")}
                _empresas_prev_list = df_launch_prev["Empresa"].unique().tolist()
                _multi_prev = len(_empresas_prev_list) > 1

                # Expander por empresa — cada arquivo exibido separadamente
                if _multi_prev:
                    for _emp_pv in _empresas_prev_list:
                        _df_pv = df_launch_prev[df_launch_prev["Empresa"] == _emp_pv]
                        _tot_d_pv = _df_pv.loc[_df_pv["Débito"] != "", "Valor"].sum()
                        _tot_c_pv = _df_pv.loc[_df_pv["Crédito"] != "", "Valor"].sum()
                        _diff_pv  = _tot_d_pv - _tot_c_pv
                        _diff_str_pv = "✅" if abs(_diff_pv) < 0.01 else "⚠️"
                        with st.expander(
                            f"📄 {_emp_pv} — {len(_df_pv)} partidas | "
                            f"Déb: R$ {format_br(_tot_d_pv)} | Cré: R$ {format_br(_tot_c_pv)} {_diff_str_pv}",
                            expanded=True,
                        ):
                            st.dataframe(
                                _df_pv[["Data", "Débito", "Crédito", "Valor", "Histórico", "Lote"]]
                                .reset_index(drop=True),
                                use_container_width=True,
                                height=min(40 * len(_df_pv) + 38, 420),
                                column_config=_col_cfg_prev,
                                hide_index=True,
                            )
                            _diff_lbl_pv = "R$ 0,00 ✅" if abs(_diff_pv) < 0.01 else f"R$ {format_br(abs(_diff_pv))} ⚠️ diferença"
                            st.caption(
                                f"{len(_df_pv)} partidas | {modo_label_prev} | "
                                f"Total Déb: R$ {format_br(_tot_d_pv)} | "
                                f"Total Cré: R$ {format_br(_tot_c_pv)} | "
                                f"Diferença (Déb−Cré): {_diff_lbl_pv}"
                            )
                else:
                    # Arquivo único
                    st.dataframe(
                        df_launch_prev[["Data", "Débito", "Crédito", "Valor", "Histórico", "Lote"]]
                        .reset_index(drop=True),
                        use_container_width=True,
                        height=280,
                        column_config=_col_cfg_prev,
                        hide_index=True,
                    )
                    _tot_deb = df_launch_prev.loc[df_launch_prev["Débito"] != "", "Valor"].sum()
                    _tot_cre = df_launch_prev.loc[df_launch_prev["Crédito"] != "", "Valor"].sum()
                    _diff    = _tot_deb - _tot_cre
                    _diff_str = "R$ 0,00 ✅" if abs(_diff) < 0.01 else f"R$ {format_br(abs(_diff))} ⚠️ diferença"
                    st.caption(
                        f"{len(df_launch_prev)} partidas | {modo_label_prev} | "
                        f"Total Déb: R$ {format_br(_tot_deb)} | "
                        f"Total Cré: R$ {format_br(_tot_cre)} | "
                        f"Diferença (Déb−Cré): {_diff_str}"
                    )

    # ════════════════════════════════════════════════════════════════════════
    # TAB 3 – GERAR ARQUIVOS DOMÍNIO
    # ════════════════════════════════════════════════════════════════════════
    with tab3:
        st.header("Gerar Arquivos para Importação no Domínio")

        if "df_all" not in st.session_state:
            st.info("ℹ️ Carregue os balancetes na aba anterior primeiro.")
        else:
            df_all = st.session_state["df_all"]
            all_dfs = st.session_state.get("all_dfs", {})

            # Selecionar empresa
            empresas = ["TODAS"] + list(all_dfs.keys())
            sel = st.selectbox("Empresa para geração dos arquivos:", empresas)
            df_gen = df_all if sel == "TODAS" else all_dfs[sel]

            # Parâmetros
            st.subheader("⚙️ Parâmetros dos Lançamentos")

            modo_lancto = st.radio(
                "Tipo de lançamento:",
                ["Partidas dobradas (1 lançamento único)", "Simples com conta de contrapartida"],
                index=0,
                horizontal=True,
                key="modo_lancto",
                help=(
                    "**Partidas dobradas**: todas as contas formam um único lançamento contábil "
                    "(sem conta de contrapartida). "
                    "**Simples**: cada conta gera um lançamento separado com conta de contrapartida."
                ),
            )
            modo = "dobradas" if "Partidas" in modo_lancto else "simples"

            c1, c2, c3 = st.columns(3)
            with c1:
                cod = st.text_input("Código da Empresa (Domínio)", value=cod_empresa, key="gen_cod")
            with c2:
                dt = st.date_input("Data de Abertura", value=opening_date, key="gen_dt", format="DD/MM/YYYY")
            with c3:
                if modo == "simples":
                    contra = st.text_input("Conta de Contrapartida", value=contra_code, key="gen_contra")
                else:
                    contra = ""
                    st.info("Partidas dobradas: todas as contas compõem **1 lançamento único** — sem contrapartida.", icon="ℹ️")

            # Filtrar contas com saldo
            df_nonzero = df_gen[df_gen["current_value"] != 0].copy()

            st.subheader("📋 Prévia das Contas a Lançar")
            c1, c2, c3, c4 = st.columns(4)
            tot_d = df_nonzero[df_nonzero["current_indicator"] == "D"]["current_value"].sum()
            tot_c = df_nonzero[df_nonzero["current_indicator"] == "C"]["current_value"].sum()
            c1.metric("Contas com saldo", len(df_nonzero))
            c2.metric("Total Déb.", f"R$ {format_br(tot_d)}")
            c3.metric("Total Cré.", f"R$ {format_br(tot_c)}")
            c4.metric(
                "Balanceado?",
                "✅ Sim" if abs(tot_d - tot_c) < 0.01 else "⚠️ Não",
            )

            df_prev = df_nonzero[["code", "description", "current_value", "current_indicator"]].copy()
            df_prev.insert(0, "Data", dt.strftime("%d/%m/%Y"))
            st.dataframe(
                df_prev.rename(columns={
                    "Data": "Data",
                    "code": "Código",
                    "description": "Descrição",
                    "current_value": "Saldo",
                    "current_indicator": "D/C",
                }),
                use_container_width=True,
                height=250,
                column_config={
                    "Saldo": st.column_config.NumberColumn("Saldo", format="R$ %.2f"),
                },
            )

            st.subheader("📤 Downloads")
            st.markdown("""
            > **Obs. sobre o Valor:** os valores são gerados em **centavos** (R\\$ 1,00 = 100).
            > Verifique a configuração da sua macro no Domínio antes de importar.
            """)

            # Slug do nome da empresa para compor o nome do arquivo
            import re as _re
            if sel == "TODAS":
                _nomes = list(all_dfs.keys())
                _emp_slug = "_".join(_re.sub(r"[^\w\-]", "_", n).strip("_") for n in _nomes)
            else:
                _emp_slug = _re.sub(r"[^\w\-]", "_", sel).strip("_")
            _dt_str = dt.strftime("%d-%m-%Y")
            _base_fname = f"{cod}_SALDO EM {_dt_str}_{_emp_slug}"

            col_a, col_b = st.columns(2)

            with col_a:
                st.markdown("**Planilha Modelo Domínio (Excel)**")
                st.caption("Cole na planilha modelo e rode a macro para gerar o TXT")
                excel_buf = generate_dominio_excel(df_nonzero, cod, dt, contra, modo=modo)
                st.download_button(
                    "⬇️  Baixar planilha_modelo_dominio.xlsx",
                    data=excel_buf,
                    file_name=f"{_base_fname}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            with col_b:
                st.markdown("**Arquivo TXT (direto para Domínio)**")
                st.caption("Arquivo pronto para importar direto no sistema Domínio")
                txt = generate_dominio_txt(df_nonzero, cod, dt, contra, modo=modo)
                st.download_button(
                    "⬇️  Baixar lancamentos_abertura.txt",
                    data=txt.encode("latin-1", errors="replace"),
                    file_name=f"{_base_fname}.txt",
                    mime="text/plain",
                    type="primary",
                    use_container_width=True,
                )

            with st.expander("👁️  Prévia do arquivo TXT (primeiras 30 linhas)"):
                preview = "\n".join(txt.replace("\r\n", "\n").split("\n")[:30])
                st.code(preview, language="text")
                total_lines = len([l for l in txt.split("\r\n") if l.strip()])
                if modo == "dobradas":
                    st.caption(f"Total: {total_lines} linhas | 1 lançamento (partidas dobradas)")
                else:
                    st.caption(f"Total: {total_lines} linhas | {total_lines // 2} lançamentos (simples)")

    # ════════════════════════════════════════════════════════════════════════
    # TAB 4 – CARTA DE RESPONSABILIDADE
    # ════════════════════════════════════════════════════════════════════════
    with tab4:
        st.header("📝 Carta de Responsabilidade sobre Saldos Contábeis")

        # ── Inicializar session_state ─────────────────────────────────────
        if "carta_accounts" not in st.session_state:
            st.session_state["carta_accounts"] = []

        # ── Seção 1 – Informações da empresa ─────────────────────────────
        with st.expander("📋 1. Informações da Empresa e do Contrato", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                carta_cliente = st.text_input(
                    "Razão Social / Nome da Empresa", key="carta_cliente",
                    placeholder="Ex.: ABC Comércio Ltda.")
                carta_cnpj = st.text_input(
                    "CNPJ da Empresa", key="carta_cnpj",
                    placeholder="00.000.000/0000-00")
                carta_ie = st.text_input(
                    "Inscrição Estadual (I.E.)", key="carta_ie",
                    placeholder="Isento ou número")
                carta_ccm = st.text_input(
                    "Inscrição Municipal (CCM)", key="carta_ccm",
                    placeholder="Isento ou número")
            with col2:
                carta_data_contrato = st.text_input(
                    "Data de assinatura do contrato", key="carta_data_contrato",
                    placeholder="dd/mm/aaaa")
                carta_plano = st.text_input(
                    "Plano", key="carta_plano",
                    placeholder="Ex.: Contabilidade as a Service 2.0")
                carta_competencia = st.text_input(
                    "Competência BHub (início)", key="carta_competencia",
                    placeholder="mm/aaaa  ex: 01/2024")
                carta_data_balancete = st.text_input(
                    "Data do Balancete Implantado", key="carta_data_balancete",
                    placeholder="dd/mm/aaaa")
                carta_link = st.text_input(
                    "Link do Balancete (Drive ou similar)",
                    key="carta_link", placeholder="https://...")

        # ── Seção 2 – Informações da carta ───────────────────────────────
        with st.expander("✍️ 2. Informações da Carta e do Assinante", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                carta_cidade = st.text_input(
                    "Cidade", key="carta_cidade", value="São Paulo")
            with col2:
                carta_data_carta = st.date_input(
                    "Data da Carta", key="carta_data_carta",
                    value=None, format="DD/MM/YYYY")
            with col3:
                carta_representante = st.text_input(
                    "Nome do Representante Legal", key="carta_representante",
                    placeholder="Nome completo do responsável")
                carta_cpf = st.text_input(
                    "CPF do Representante (opcional)", key="carta_cpf",
                    placeholder="000.000.000-00")

        # ── Seção 3 – Tabela de contas ───────────────────────────────────
        st.subheader("3. Contas e Saldos para a Carta")

        # Importar do balancete carregado
        col_imp, col_rep_lbl, col_rep, col_rep_manual, col_rep_btn, col_clr_jus, col_add, col_clr = st.columns(
            [2, 1.4, 2, 1.8, 1, 1.2, 1, 1])

        with col_imp:
            if st.button("⬇️ Importar do Balancete", use_container_width=True,
                         help="Importa as contas usando os códigos BHub do De-Para (Aba 2)"):
                all_dfs = st.session_state.get("all_dfs", {})
                df_sug_ref = st.session_state.get("df_sug_current")

                # Monta lookup: código antigo → (código BHub, nome BHub)
                bhub_lookup = {}
                if df_sug_ref is not None and not df_sug_ref.empty:
                    for _, sr in df_sug_ref.iterrows():
                        old_c  = str(sr.get("Cód. Antigo", "")).strip()
                        bhub_c = str(sr.get("Cód. BHub", "")).strip()
                        bhub_n = str(sr.get("Nome BHub Sugerido", "")).strip()
                        if old_c and bhub_c:
                            bhub_lookup[old_c] = (bhub_c, bhub_n)

                rows_imp = []
                for emp_key, df_bal in all_dfs.items():
                    if df_bal is None or df_bal.empty:
                        continue
                    for _, r in df_bal.iterrows():
                        cod  = str(r.get("code", "")).strip()
                        desc = str(r.get("description", "")).strip()
                        val  = float(r.get("current_value", 0) or 0)
                        ind  = str(r.get("current_indicator", "D")).strip()
                        saldo_v = val if ind == "D" else -val

                        # Usa código BHub se disponível no De-Para
                        if cod in bhub_lookup:
                            bhub_c, bhub_n = bhub_lookup[cod]
                            conta_label = f"{bhub_c} - {bhub_n}" if bhub_n else bhub_c
                        else:
                            conta_label = f"{cod} - {desc}" if cod else desc

                        rows_imp.append({
                            "conta":          conta_label,
                            "saldo":          round(saldo_v, 2),
                            "ajuste":         0.0,
                            "saldo_final":    round(saldo_v, 2),
                            "justificativa":  "",
                            "contra_partida": "",
                        })
                if rows_imp:
                    st.session_state["carta_accounts"] = rows_imp
                    suffix = " com códigos BHub do De-Para" if bhub_lookup else " (acesse a Aba 2 para mapear os códigos BHub)"
                    st.success(f"{len(rows_imp)} contas importadas{suffix}.")
                else:
                    st.warning("Nenhum balancete carregado na Aba 1.")

        with col_rep_lbl:
            st.markdown("<div style='padding-top:28px;'>Replicar justificativa:</div>",
                        unsafe_allow_html=True)
        with col_rep:
            jus_replicate = st.selectbox(
                "Justificativa",
                options=PREDEFINED_JUSTIFICATIONS,
                key="carta_jus_replicate",
                label_visibility="collapsed",
            )
        with col_rep_manual:
            jus_manual = st.text_input(
                "Justificativa manual",
                key="carta_jus_manual",
                label_visibility="collapsed",
                placeholder="ou digitar manualmente...",
            )
        with col_rep_btn:
            if st.button("↩️ Para todas", use_container_width=True,
                         help="Replica a justificativa para todas as linhas (campo manual tem prioridade sobre o seletor)"):
                jus_final = jus_manual.strip() if jus_manual.strip() else jus_replicate
                for acc in st.session_state["carta_accounts"]:
                    acc["justificativa"] = jus_final
                st.rerun()

        with col_clr_jus:
            if st.button("🧹 Limpar just.", use_container_width=True,
                         help="Remove a justificativa de todas as linhas"):
                for acc in st.session_state["carta_accounts"]:
                    acc["justificativa"] = ""
                st.rerun()

        with col_add:
            if st.button("➕ Linha", use_container_width=True):
                st.session_state["carta_accounts"].append({
                    "conta": "", "saldo": 0.0, "ajuste": 0.0,
                    "saldo_final": 0.0, "justificativa": "", "contra_partida": "",
                })
                st.rerun()

        with col_clr:
            if st.button("🗑️ Limpar", use_container_width=True):
                st.session_state["carta_accounts"] = []
                st.rerun()

        # Editor da tabela
        if not st.session_state["carta_accounts"]:
            st.session_state["carta_accounts"] = [{
                "conta": "", "saldo": 0.0, "ajuste": 0.0,
                "saldo_final": 0.0, "justificativa": "", "contra_partida": "",
            }]

        df_carta = pd.DataFrame(st.session_state["carta_accounts"])

        edited = st.data_editor(
            df_carta,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "conta": st.column_config.TextColumn(
                    "Conta de Origem",
                    help="Código + Descrição da conta",
                    width="large",
                ),
                "saldo": st.column_config.NumberColumn(
                    "Saldo Balancete",
                    help="Saldo conforme balancete (D=positivo, C=negativo)",
                    format="%.2f",
                    width="medium",
                ),
                "ajuste": st.column_config.NumberColumn(
                    "Valor de Ajuste",
                    help="Valor do ajuste no 1° dia da competência BHub",
                    format="%.2f",
                    width="medium",
                ),
                "saldo_final": st.column_config.NumberColumn(
                    "Saldo Final",
                    help="Calculado automaticamente: Saldo + Ajuste",
                    format="%.2f",
                    width="medium",
                    disabled=True,
                ),
                "justificativa": st.column_config.TextColumn(
                    "Justificativa",
                    help="Digite livremente ou use o seletor + 'Para todas' acima",
                    width="large",
                ),
                "contra_partida": st.column_config.TextColumn(
                    "Contra-Partida BHub",
                    help="Sugestão de conta de contra-partida no plano BHub",
                    width="large",
                ),
            },
            key="carta_editor",
        )

        # Recalcular saldo_final, auto-popular contra-partida e persistir
        edited_list = edited.to_dict("records")
        _sf_changed = False
        for row in edited_list:
            s = float(row.get("saldo", 0) or 0)
            a = float(row.get("ajuste", 0) or 0)
            new_sf = round(s + a, 2)
            if abs(new_sf - float(row.get("saldo_final", 0) or 0)) > 0.001:
                _sf_changed = True
            row["saldo_final"] = new_sf
            # Auto-preenche contra-partida quando há ajuste e campo ainda vazio
            if a != 0 and not str(row.get("contra_partida", "")).strip():
                row["contra_partida"] = "2875 - AJUSTES DE EXERCÍCIOS ANTERIORES"
        st.session_state["carta_accounts"] = edited_list
        # Força rerun para atualizar a coluna Saldo Final (disabled) no editor
        if _sf_changed:
            st.rerun()

        # ── Resumo de totais do Saldo Balancete ──────────────────────────
        total_deb = sum(r["saldo"] for r in edited_list if float(r.get("saldo", 0) or 0) > 0)
        total_cre = sum(abs(float(r.get("saldo", 0) or 0)) for r in edited_list if float(r.get("saldo", 0) or 0) < 0)
        diferenca = total_deb - total_cre

        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
        col_r1.metric("Total Débito – Saldo Balancete", f"R$ {format_br(total_deb)}")
        col_r2.metric("Total Crédito – Saldo Balancete", f"R$ {format_br(total_cre)}")
        col_r3.metric("Diferença", f"R$ {format_br(abs(diferenca))}")
        with col_r4:
            if abs(diferenca) < 0.01:
                st.success("✅ Balanceado")
            else:
                natureza = "D" if diferenca > 0 else "C"
                st.warning(f"⚠️ Diferença de R$ {format_br(abs(diferenca))} ({natureza})")

        # ── Seção 4 – Downloads ──────────────────────────────────────────
        st.subheader("4. Exportar")

        info_carta = {
            "cliente":         st.session_state.get("carta_cliente", ""),
            "cnpj":            _fmt_cnpj(st.session_state.get("carta_cnpj", "")),
            "ie":              st.session_state.get("carta_ie", ""),
            "ccm":             st.session_state.get("carta_ccm", ""),
            "data_contrato":   st.session_state.get("carta_data_contrato", ""),
            "plano":           st.session_state.get("carta_plano", ""),
            "competencia_bhub": st.session_state.get("carta_competencia", ""),
            "data_balancete":  st.session_state.get("carta_data_balancete", ""),
            "link_balancete":  st.session_state.get("carta_link", ""),
            "cidade":          st.session_state.get("carta_cidade", "São Paulo"),
            "data_carta":      _fmt_date_ptbr(st.session_state["carta_data_carta"]) if st.session_state.get("carta_data_carta") else "",
            "representante":   st.session_state.get("carta_representante", ""),
            "cpf":             st.session_state.get("carta_cpf", ""),
        }
        nome_arq = (info_carta["cliente"] or "empresa").replace(" ", "_").replace("/", "-")

        col_dl1, col_dl2, col_dl3 = st.columns(3)

        with col_dl1:
            if st.button("📊 Gerar Excel", use_container_width=True, type="primary"):
                with st.spinner("Gerando planilha..."):
                    try:
                        buf_xl = generate_carta_excel(info_carta, edited_list)
                        st.download_button(
                            "⬇️ Baixar Excel",
                            data=buf_xl,
                            file_name=f"Carta_Responsabilidade_{nome_arq}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar Excel: {e}")

        with col_dl2:
            if st.button("📄 Gerar Word (.docx)", use_container_width=True, type="primary"):
                with st.spinner("Gerando documento Word..."):
                    try:
                        buf_docx = generate_carta_word(info_carta, edited_list)
                        st.download_button(
                            "⬇️ Baixar Word",
                            data=buf_docx,
                            file_name=f"Carta_Responsabilidade_{nome_arq}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar Word: {e}")

        with col_dl3:
            if st.button("📑 Gerar PDF", use_container_width=True, type="primary"):
                with st.spinner("Gerando PDF..."):
                    try:
                        buf_pdf = generate_carta_pdf(info_carta, edited_list)
                        st.download_button(
                            "⬇️ Baixar PDF",
                            data=buf_pdf,
                            file_name=f"Carta_Responsabilidade_{nome_arq}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar PDF: {e}")

        # ── Seção 5 – Lançamento de Ajuste (Domínio) ─────────────────────
        st.divider()
        st.subheader("5. Lançamento de Ajuste – Domínio")
        st.caption(
            "Gera os lançamentos de ajuste da carta no formato da Planilha Modelo Domínio "
            "(Excel + TXT), prontos para importação."
        )

        # Deriva data de ajuste do campo Competência BHub (mm/aaaa → 01/mm/aaaa)
        _competencia_raw = st.session_state.get("carta_competencia", "").strip()
        _default_ajuste_date = opening_date  # fallback sidebar
        try:
            if _competencia_raw:
                _parts = _competencia_raw.replace("-", "/").split("/")
                if len(_parts) == 2:
                    _mes, _ano = int(_parts[0]), int(_parts[1])
                    _default_ajuste_date = date(_ano, _mes, 1)
        except Exception:
            pass

        col_aj_a, col_aj_b = st.columns(2)
        with col_aj_a:
            carta_ajuste_date = st.date_input(
                "Data do lançamento de ajuste",
                value=_default_ajuste_date,
                key="carta_ajuste_date",
                format="DD/MM/YYYY",
                help="Preenchido automaticamente com o 1° dia da Competência BHub informada acima.",
            )
        with col_aj_b:
            carta_ajuste_empresa = st.text_input(
                "Código da Empresa no Domínio",
                value=cod_empresa,
                key="carta_ajuste_empresa",
            )

        contas_com_ajuste = [r for r in edited_list if float(r.get("ajuste", 0) or 0) != 0]
        if not contas_com_ajuste:
            st.info("Nenhuma conta possui Valor de Ajuste preenchido na tabela acima.")
        else:
            st.caption(f"{len(contas_com_ajuste)} conta(s) com ajuste detectada(s).")

            with st.expander("📋 Prévia dos lançamentos de ajuste", expanded=False):
                prev_rows = []
                _date_str = carta_ajuste_date.strftime("%d/%m/%Y")
                _net_pl = 0.0   # impacto líquido em Lucros/Prejuízos Acumulados

                for acc in contas_com_ajuste:
                    ajuste_v = float(acc.get("ajuste", 0) or 0)
                    code, desc = _parse_conta_code(acc.get("conta", ""))
                    contra_raw = str(acc.get("contra_partida", "")).strip() or "2875 - AJUSTES DE EXERCÍCIOS ANTERIORES"
                    contra_c, _ = _parse_conta_code(contra_raw)
                    abs_val = abs(ajuste_v)
                    hist = f"AJUSTE DE IMPLANTAÇÃO {_date_str} - {desc}"[:55]

                    if ajuste_v > 0:
                        # D: conta  |  C: Lucros Acumulados → Lucros aumenta (crédito)
                        deb, cre = code or acc.get("conta", ""), contra_c or contra_raw
                        _net_pl += abs_val    # crédito em Lucros
                    else:
                        # D: Lucros Acumulados  |  C: conta → Lucros diminui (débito)
                        deb, cre = contra_c or contra_raw, code or acc.get("conta", "")
                        _net_pl -= abs_val    # débito em Lucros

                    prev_rows.append({
                        "Data":       _date_str,
                        "Débito":     deb,
                        "Crédito":    cre,
                        "Valor (R$)": f"R$ {format_br(abs_val)}",
                        "Histórico":  hist,
                    })

                df_prev = pd.DataFrame(prev_rows)
                st.dataframe(df_prev, use_container_width=True, hide_index=True)

                # Impacto líquido em Lucros/Prejuízos Acumulados
                st.divider()
                st.markdown("**Impacto líquido no Patrimônio Líquido (Lucros/Prejuízos Acumulados):**")
                cp1, cp2, cp3 = st.columns(3)
                cp1.metric("Qtd. lançamentos", len(prev_rows))
                cp2.metric("Valor total movimentado", f"R$ {format_br(sum(abs(float(r.get('ajuste',0) or 0)) for r in contas_com_ajuste))}")
                if abs(_net_pl) < 0.01:
                    cp3.metric("Saldo líquido em 2875", "R$ 0,00 — Neutro")
                elif _net_pl > 0:
                    cp3.metric("Saldo líquido em 2875", f"+ R$ {format_br(_net_pl)}",
                               delta="Crédito (aumenta PL)", delta_color="normal")
                else:
                    cp3.metric("Saldo líquido em 2875", f"- R$ {format_br(abs(_net_pl))}",
                               delta="Débito (reduz PL)", delta_color="inverse")

        col_aj1, col_aj2 = st.columns(2)

        with col_aj1:
            if st.button("📊 Excel – Lançamento Ajuste", use_container_width=True,
                         disabled=not contas_com_ajuste):
                with st.spinner("Gerando Excel..."):
                    try:
                        buf_aj = generate_ajuste_excel(
                            contas_com_ajuste, carta_ajuste_date, carta_ajuste_empresa
                        )
                        st.download_button(
                            "⬇️ Baixar Excel (Ajuste)",
                            data=buf_aj,
                            file_name=f"Lancamento_Ajuste_{nome_arq}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar Excel: {e}")

        with col_aj2:
            if st.button("📄 TXT – Lançamento Ajuste", use_container_width=True,
                         disabled=not contas_com_ajuste):
                with st.spinner("Gerando TXT..."):
                    try:
                        txt_aj = generate_ajuste_txt(
                            contas_com_ajuste, carta_ajuste_date, carta_ajuste_empresa
                        )
                        st.download_button(
                            "⬇️ Baixar TXT (Ajuste)",
                            data=txt_aj.encode("utf-8"),
                            file_name=f"Lancamento_Ajuste_{nome_arq}.txt",
                            mime="text/plain",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar TXT: {e}")


def generate_dominio_from_depara(
    df_depara: pd.DataFrame,
    cod_empresa: str,
    opening_date: date,
    contra_code: str,
) -> tuple[io.BytesIO, str]:
    """
    Gera arquivos Domínio usando os NOVOS códigos BHub do de-para preenchido.
    O DataFrame deve ter colunas: novo_codigo, saldo, natureza, descricao
    """
    df_work = df_depara.copy()
    df_work["code"] = df_work.get("novo_codigo", df_work.get("code", ""))
    df_work["current_value"] = df_work.get("saldo", df_work.get("current_value", 0))
    df_work["current_indicator"] = df_work.get("natureza", df_work.get("current_indicator", "D"))
    df_work["description"] = df_work.get("descricao_nova", df_work.get("description", ""))

    excel_buf = generate_dominio_excel(df_work, cod_empresa, opening_date, contra_code)
    txt = generate_dominio_txt(df_work, cod_empresa, opening_date, contra_code)
    return excel_buf, txt


if __name__ == "__main__":
    main()
