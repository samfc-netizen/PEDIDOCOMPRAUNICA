import re
import math
import difflib
import json
import unicodedata
from io import BytesIO
from datetime import datetime, date

import pdfplumber
import pandas as pd
import streamlit as st

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

try:
    from google.oauth2 import service_account
    from google.oauth2.credentials import Credentials as UserCredentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
except Exception:
    service_account = None
    UserCredentials = None
    build = None
    MediaIoBaseUpload = None

st.set_page_config(page_title="Análise de Giro e Pedido de Compra", layout="wide", page_icon="📊")

# =========================================================
# CONFIGURAÇÕES DO NEGÓCIO
# =========================================================

LOJAS_MAP = {
    "004": "ADE",
    "006": "GAMA",
    "009": "ÚNICA",
    "012": "SOFNORTE",
    "013": "CEILÂNDIA",
    "014": "SIA",
    "015": "UNAÍ",
    "016": "AG LINDAS",
    "022": "GUARÁ",
    "024": "LUZIÂNIA",
}

CODIGOS_LOJAS = ["004", "006", "012", "013", "014", "015", "016", "022", "024"]
CODIGO_UNICA = "009"
MESES_PADRAO = ["01/2026", "02/2026", "03/2026", "04/2026"]
MESES = MESES_PADRAO.copy()

GOOGLE_DRIVE_ROOT_FOLDER_ID = "1PqWXzphyeU_Q5Wc7UDYeZUzFA4kKZL-H"
GOOGLE_SUBPASTA_PEDIDOS = "Pedidos Editaveis"
GOOGLE_SUBPASTA_FINAIS = "Arquivos Finais"
GOOGLE_PLANILHA_CONTROLE = "Controle de Pedidos - Dauto"
GOOGLE_PLANILHA_CADASTRO = "Cadastro de Produtos - Dauto"
GOOGLE_PEDIDOS_COLUNAS = [
    "id_pedido", "nome_pedido", "fornecedor", "status", "valor",
    "criado_em", "criado_por", "aprovado_em", "aprovado_por",
    "link_pedido", "spreadsheet_id", "link_autcom", "link_fornecedor", "observacao",
]
GOOGLE_ACOMPANHAMENTO_COLUNAS = [
    "data", "mes", "fornecedor", "nome_pedido", "valor",
    "status", "link_pedido", "link_autcom", "link_fornecedor",
]

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def br_to_float(value):
    if value is None:
        return 0.0
    value = str(value).strip()
    if value == "" or value.lower() in ["nan", "none", "-"]:
        return 0.0
    value = value.replace("R$", "").replace(" ", "")
    value = value.replace(".", "").replace(",", ".")
    try:
        return float(value)
    except Exception:
        return 0.0


def numero_planilha_para_float(value):
    """
    Converte números vindos da planilha final, aceitando:
    - 28.12  -> 28.12
    - 28,12  -> 28.12
    - 1.234,56 -> 1234.56
    - 1,234.56 -> 1234.56
    Evita o erro de transformar 28.12 em 2812.
    """
    if value is None:
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            if pd.isna(value):
                return 0.0
            return float(value)
        except Exception:
            pass

    txt = str(value).strip()
    if txt == "" or txt.lower() in ["nan", "none", "-"]:
        return 0.0

    txt = txt.replace("R$", "").replace(" ", "").replace("\xa0", "")

    # Se tem vírgula e ponto, decide pelo último separador como decimal
    if "," in txt and "." in txt:
        if txt.rfind(",") > txt.rfind("."):
            # padrão brasileiro: 1.234,56
            txt = txt.replace(".", "").replace(",", ".")
        else:
            # padrão americano: 1,234.56
            txt = txt.replace(",", "")
    elif "," in txt:
        # padrão brasileiro simples: 28,12
        txt = txt.replace(".", "").replace(",", ".")
    elif "." in txt:
        # Mantém ponto como decimal quando parecer número decimal.
        # Também cobre floats lidos como texto pelo Excel, ex.: 25.940000000000001.
        partes = txt.split(".")
        if len(partes) == 2:
            int_part, dec_part = partes[0], partes[1]
            if len(dec_part) <= 2 or len(dec_part) > 3:
                pass
            elif len(int_part) <= 2:
                pass
            else:
                txt = txt.replace(".", "")
        else:
            # caso venha como milhar: 1.234.567
            txt = txt.replace(".", "")

    try:
        return float(txt)
    except Exception:
        return 0.0


def normalizar_coluna(nome):
    nome = str(nome).strip().upper().replace("\ufeff", "")
    nome = re.sub(r"\s+", " ", nome)
    return nome


def parse_data_br(value):
    if value is None:
        return pd.NaT
    value = str(value).strip()
    if not value or value.lower() in ["nan", "none", "-"]:
        return pd.NaT
    match = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", value)
    if not match:
        return pd.NaT
    data_txt = match.group(1)
    for fmt in ["%d/%m/%Y", "%d/%m/%y"]:
        try:
            return datetime.strptime(data_txt, fmt).date()
        except Exception:
            continue
    return pd.NaT


def format_data_br(value):
    if pd.isna(value) or value is None or value == "":
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    parsed = parse_data_br(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def format_num_br(value, casas=1):
    try:
        value = round(float(value), casas)
        texto = f"{value:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if "," in texto:
            texto = texto.rstrip("0").rstrip(",")
        return texto
    except Exception:
        return value


def format_int_br(value):
    try:
        return f"{int(round(float(value))):,}".replace(",", ".")
    except Exception:
        return value


def format_moeda_br(value):
    try:
        value = float(value)
        return "R$ " + f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def extract_text_from_pdf(uploaded_file):
    text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text(x_tolerance=1, y_tolerance=3)
            if page_text:
                text += page_text + "\n"
    return text


MESES_ABREV_PT = {
    "01": "Jan", "02": "Fev", "03": "Mar", "04": "Abr", "05": "Mai", "06": "Jun",
    "07": "Jul", "08": "Ago", "09": "Set", "10": "Out", "11": "Nov", "12": "Dez",
}


def extrair_meses_giro_pdf(text):
    """
    Puxa os meses diretamente do PDF de Giro.
    Prioriza a ordem real das colunas da tabela, porque alguns relatórios exibem
    "REFERENTE AOS MESES" em ordem decrescente, mas as colunas de giro aparecem em
    ordem crescente no cabeçalho.
    """
    text = str(text or "")

    # Ordem usada para mapear os valores de cada linha do produto.
    # Ex.: o PDF pode dizer "REFERENTE AOS MESES: 06, 05, 04, 03", mas a tabela
    # estar visualmente como "03/2026 04/2026 05/2026 06/2026".
    for raw_line in text.splitlines():
        line = raw_line.strip()
        line_upper = line.upper()
        if "COD" not in line_upper or "MEDIA" not in line_upper:
            continue

        meses_linha = re.findall(r"\b\d{2}/\d{4}\b", line)
        if len(meses_linha) >= 2:
            return meses_linha

    for raw_line in text.splitlines():
        line = raw_line.strip()
        meses_linha = re.findall(r"\b\d{2}/\d{4}\b", line)
        if len(meses_linha) >= 2 and ("ESTOQUE" in line.upper() or "MEDIA" in line.upper() or "MÉDIA" in line.upper()):
            return meses_linha

    padrao_ref = re.search(r"REFERENTE\s+AOS\s+MESES\s*:\s*([^\n]+)", text, flags=re.IGNORECASE)
    if padrao_ref:
        meses = re.findall(r"\b\d{2}/\d{4}\b", padrao_ref.group(1))
        if meses:
            return sorted(meses, key=lambda mes: (int(mes.split("/")[1]), int(mes.split("/")[0])))

    meses = []
    for mes in re.findall(r"\b\d{2}/\d{4}\b", text):
        if mes not in meses:
            meses.append(mes)
    return meses[:4] if meses else MESES_PADRAO.copy()


def label_mes_giro(mes):
    try:
        mm, yyyy = str(mes).split("/")
        return f"{MESES_ABREV_PT.get(mm.zfill(2), mm)}/{str(yyyy)[-2:]}"
    except Exception:
        return str(mes)


def col_giro(prefixo, mes):
    return f"{prefixo} {label_mes_giro(mes)}"


def colunas_pedido_compras(meses_ref=None):
    """
    Ordem oficial da tela Pedido de Compra e do download Pedido Editável.
    Mantém Código Fábrica e Embalagem no final, conforme solicitado.
    """
    meses_ref = meses_ref or MESES
    return [
        "codigo",
        "descricao",
        *[col_giro("Giro Geral", mes) for mes in meses_ref],
        "Média Giro Geral",
        "Estoque Lojas",
        "Estoque Única",
        "Estoque Geral",
        "Saldo em Trânsito/ABERTO",
        "Estoque Final",
        "Estoque Alvo",
        "Sugestão Sistema",
        "Sugestão arredondada",
        "Preço Última Compra",
        "Data Última Compra",
        "PEDIDO Final",
        "Origem Sugestão",
        "Valor Final do Pedido",
        "Embalagem",
        "Código Fábrica",
    ]

# =========================================================
# LEITURA DO PDF DE GIRO DE ESTOQUE
# =========================================================

def _token_numero_giro(valor):
    """
    Identifica números válidos nas colunas numéricas do relatório de giro.
    Evita tratar pedaços da descrição/código de fábrica como giro.
    """
    txt = str(valor or "").strip()
    if txt == "":
        return False
    return bool(re.fullmatch(r"-?\d+(?:\.\d{3})*,\d{1,4}|-?\d+,\d{1,4}|-?\d+", txt))


def _encontrar_unidade_giro(partes, qtd_meses):
    """
    No PDF de giro, algumas descrições contêm a palavra/tipo 'UN' antes da unidade real.
    Exemplo real:
    85582 ... 0.5L UN 1263 T UN 0,00 0,00 0,00 ...
    A unidade correta é o último UN antes da sequência dos meses.

    Esta função procura a unidade que vem imediatamente antes de uma sequência numérica
    com a quantidade de meses do cabeçalho. Assim o sistema não puxa '1263' da descrição
    como se fosse giro de abril.
    """
    candidatos = []

    for i, token in enumerate(partes):
        token_upper = str(token).strip().upper()
        unidade_explicita = token_upper in ["UN", "UND", "UNID", "UNIDADE"]

        # Alguns PDFs colam a unidade no fim do texto, ex.: "...-STUN 2,45 1,00..."
        unidade_colada = (
            not unidade_explicita
            and len(token_upper) > 2
            and token_upper.endswith("UN")
            and not _token_numero_giro(token_upper)
        )

        if not unidade_explicita and not unidade_colada:
            continue

        proximos = partes[i + 1:i + 1 + qtd_meses]
        if len(proximos) < qtd_meses:
            continue

        if all(_token_numero_giro(v) for v in proximos):
            candidatos.append((i, unidade_colada))

    if not candidatos:
        return None, False

    # Usa o último candidato válido, porque a descrição pode conter "UN" antes da unidade real.
    return candidatos[-1]


def parse_linha_giro(line, meses_ref=None):
    """
    Layout esperado:
    COD DESCRICAO DO ITEM UN 04/2026 05/2026 06/2026 MEDIA PREVI.30 ESTOQUE
    SUGESTAO PR.ULT.COMP DT.ULT.COMP PR.VENDA % LUCRO

    Correção importante:
    - Não usa mais o primeiro token "UN" encontrado.
    - Localiza a unidade pela sequência numérica dos meses logo depois dela.
    - Isso evita erro em itens cuja descrição contém "UN" ou códigos numéricos antes
      da unidade real, como o item 85582.
    """
    if not re.match(r"^\d{5}\s+", str(line).strip()):
        return None

    partes = str(line).strip().split()
    codigo = partes[0].zfill(5)

    meses_ref = meses_ref or MESES_PADRAO
    qtd_meses = len(meses_ref)

    un_index, unidade_colada = _encontrar_unidade_giro(partes, qtd_meses)
    if un_index is None:
        return None

    antes_un = partes[1:un_index]
    depois_un = partes[un_index + 1:]

    # Se a unidade veio colada no final do último token da descrição, remove só o "UN".
    if unidade_colada:
        token_sem_un = str(partes[un_index])[:-2].strip()
        if token_sem_un:
            antes_un = partes[1:un_index] + [token_sem_un]

    if len(depois_un) < qtd_meses + 4:
        return None

    # Garantia adicional: os meses precisam ser exatamente a primeira sequência após a unidade.
    if not all(_token_numero_giro(v) for v in depois_un[:qtd_meses]):
        return None

    codigo_fabrica_extraido = ""
    descricao_tokens = list(antes_un)

    if descricao_tokens:
        ultimo_token = descricao_tokens[-1]
        if re.fullmatch(r"\d{5,}", ultimo_token):
            codigo_fabrica_extraido = ultimo_token[-6:] if len(ultimo_token) > 7 else ultimo_token
            descricao_tokens = descricao_tokens[:-1]
        else:
            match_fabrica = re.search(r"(\d{5,})$", ultimo_token)
            if match_fabrica:
                codigo_raw = match_fabrica.group(1)
                codigo_fabrica_extraido = codigo_raw[-6:] if len(codigo_raw) > 7 else codigo_raw
                descricao_tokens[-1] = ultimo_token[:match_fabrica.start(1)].rstrip("- ")

    data_idx = None
    for i, token in enumerate(depois_un):
        if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", str(token).strip()):
            data_idx = i
            break

    dt_ult_compra = pd.NaT
    dt_ult_compra_txt = ""
    pr_ult_compra = 0.0

    if data_idx is not None:
        dt_ult_compra_txt = depois_un[data_idx]
        dt_ult_compra = parse_data_br(dt_ult_compra_txt)
        if data_idx - 1 >= 0:
            pr_ult_compra = br_to_float(depois_un[data_idx - 1])
    else:
        # Após os meses: MEDIA, PREVI.30, ESTOQUE, SUGESTAO, PR.ULT.COMP...
        idx_preco = qtd_meses + 4
        pr_ult_compra = br_to_float(depois_un[idx_preco]) if len(depois_un) > idx_preco else 0.0

    idx_estoque = qtd_meses + 2

    return {
        "codigo": codigo,
        "descricao": " ".join(descricao_tokens).strip(),
        "codigo_fabrica": codigo_fabrica_extraido,
        **{mes: br_to_float(depois_un[i]) for i, mes in enumerate(meses_ref)},
        "estoque": br_to_float(depois_un[idx_estoque]) if len(depois_un) > idx_estoque else 0,
        "pr_ult_compra": pr_ult_compra,
        "dt_ult_compra": dt_ult_compra,
        "dt_ult_compra_txt": dt_ult_compra_txt,
        "codigo_empresa": None,
        "loja": None,
    }


def parse_giro_estoque(text, meses_ref=None):
    meses_ref = meses_ref or extrair_meses_giro_pdf(text)
    registros = []
    empresa_atual = None

    for raw_line in text.splitlines():
        line = raw_line.strip()

        empresa_match = re.search(r"EMPRESA\s*:\s*(\d{3})\s*-", line)
        if empresa_match:
            empresa_atual = empresa_match.group(1)
            continue

        if not empresa_atual or empresa_atual not in LOJAS_MAP:
            continue

        produto = parse_linha_giro(line, meses_ref=meses_ref)
        if produto:
            produto["codigo_empresa"] = empresa_atual
            produto["loja"] = LOJAS_MAP[empresa_atual]
            registros.append(produto)

    return pd.DataFrame(registros)

# =========================================================
# LEITURA DO PDF DE PEDIDOS EM ABERTO / SALDO EM TRÂNSITO
# =========================================================

_NUMERO_BR_RE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d+$|^-?\d+,\d+$|^-?\d+$")


def _eh_numero_br(valor):
    return bool(_NUMERO_BR_RE.match(str(valor).strip()))


def extrair_descricao_pedido_aberto_tokens(tokens, un_index):
    if not tokens or un_index is None or un_index <= 0:
        return ""

    descricao_tokens = []
    primeiro = str(tokens[0]).strip()
    match_primeiro = re.match(r"^\d{5}[-\s]*(.*)$", primeiro)
    if match_primeiro and match_primeiro.group(1).strip():
        descricao_tokens.append(match_primeiro.group(1).strip(" -"))

    descricao_tokens.extend(str(t).strip() for t in tokens[1:un_index] if str(t).strip())
    return " ".join(descricao_tokens).strip()


def encontrar_indice_aberto_no_cabecalho(text):
    """
    No relatório de Pedidos de Compra, após a unidade UN, a sequência numérica é:
    QTDE, TOT.LIT, TOT.KIL, PES.ITE, BAIXADO, ABERTO, VR.UNIT, TOT.IPI, ALQ.IPI, TOT.SUB, TOTAL.

    Portanto, ABERTO é sempre o 6º número depois do UN, índice 5.
    """
    return 5


def parse_linha_pedido_aberto(line, indice_aberto=None):
    """
    Lê uma linha do PDF de pedidos em aberto.

    Regra corrigida:
    - Não usa a posição do cabeçalho inteiro, porque a descrição varia.
    - Localiza a unidade UN.
    - Depois da UN, considera apenas números.
    - Puxa a coluna ABERTO pelo índice fixo 5:
      QTDE=0, TOT.LIT=1, TOT.KIL=2, PES.ITE=3, BAIXADO=4, ABERTO=5, VR.UNIT=6.
    """
    line = line.strip()
    match = re.match(r"^(\d{5})[-\s]", line)
    if not match:
        return None

    codigo = match.group(1).zfill(5)
    partes = line.split()

    un_index = None
    for i, token in enumerate(partes):
        if token.upper() in ["UN", "UND", "UNID", "UNIDADE"]:
            un_index = i
            break

    if un_index is None:
        return None

    descricao = extrair_descricao_pedido_aberto_tokens(partes, un_index)
    valores_numericos = [p for p in partes[un_index + 1:] if _eh_numero_br(p)]
    idx_aberto = 5 if indice_aberto is None else int(indice_aberto)

    if len(valores_numericos) > idx_aberto:
        return {"codigo": codigo, "descricao": descricao, "Saldo em Trânsito/ABERTO": br_to_float(valores_numericos[idx_aberto])}

    return {"codigo": codigo, "descricao": descricao, "Saldo em Trânsito/ABERTO": 0.0}


def parse_pedidos_compra_aberto(text):
    registros = []
    indice_aberto = encontrar_indice_aberto_no_cabecalho(text)

    for raw_line in text.splitlines():
        produto = parse_linha_pedido_aberto(raw_line, indice_aberto=indice_aberto)
        if produto:
            registros.append(produto)

    if not registros:
        return pd.DataFrame(columns=["codigo", "descricao", "Saldo em Trânsito/ABERTO"])

    return pd.DataFrame(registros).groupby("codigo", as_index=False).agg({
        "descricao": "first",
        "Saldo em Trânsito/ABERTO": "sum",
    })


def parse_pedidos_compra_aberto_pdf(uploaded_file):
    """
    Lê o PDF de Pedidos em Aberto / Saldo em Trânsito.

    Correção principal:
    - Para este relatório, a coluna ABERTO é sempre o 6º número após a unidade UN.
      Sequência depois do UN:
      QTDE=0, TOT.LIT=1, TOT.KIL=2, PES.ITE=3, BAIXADO=4, ABERTO=5, VR.UNIT=6.
    - Por isso, a leitura por texto é usada primeiro. Ela é mais segura do que coordenadas,
      porque algumas linhas do PDF vêm com campos colados, ex.: ST000001-WANDA.
    - A leitura por coordenadas fica apenas como fallback.
    """

    # 1) Caminho principal: texto linha a linha, usando a posição fixa do ABERTO após UN.
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    try:
        texto = extract_text_from_pdf(uploaded_file)
        df_texto = parse_pedidos_compra_aberto(texto)
        if df_texto is not None and not df_texto.empty:
            df_texto["codigo"] = df_texto["codigo"].astype(str).str.extract(r"(\d+)", expand=False).fillna("").str.zfill(5)
            df_texto["Saldo em Trânsito/ABERTO"] = pd.to_numeric(df_texto["Saldo em Trânsito/ABERTO"], errors="coerce").fillna(0)
            return df_texto.groupby("codigo", as_index=False).agg({
                "descricao": "first",
                "Saldo em Trânsito/ABERTO": "sum",
            })
    except Exception:
        pass

    # 2) Fallback: tenta por coordenadas caso o texto não tenha retornado itens.
    registros = []

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    try:
        with pdfplumber.open(uploaded_file) as pdf:
            aberto_x = None

            for page in pdf.pages:
                words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False) or []
                if not words:
                    continue

                linhas = {}
                for w in words:
                    top_key = round(float(w.get("top", 0)) / 3) * 3
                    linhas.setdefault(top_key, []).append(w)

                for _, linha_words in sorted(linhas.items()):
                    linha_words = sorted(linha_words, key=lambda w: float(w.get("x0", 0)))
                    textos = [str(w.get("text", "")).strip() for w in linha_words]
                    textos_upper = [t.upper() for t in textos]

                    if (
                        "QTDE" in textos_upper
                        and "BAIXADO" in textos_upper
                        and "ABERTO" in textos_upper
                        and "VR.UNIT" in textos_upper
                    ):
                        idx = textos_upper.index("ABERTO")
                        w = linha_words[idx]
                        aberto_x = (float(w["x0"]) + float(w["x1"])) / 2
                        continue

                    if aberto_x is None or not textos:
                        continue

                    match_codigo = re.match(r"^(\d{5})(?:[-\s]|$)", textos[0])
                    if not match_codigo:
                        continue

                    codigo = match_codigo.group(1).zfill(5)

                    un_index = None
                    for i, token in enumerate(textos):
                        if token.upper() in ["UN", "UND", "UNID", "UNIDADE"]:
                            un_index = i
                            break

                    descricao = extrair_descricao_pedido_aberto_tokens(textos, un_index)

                    # Mesmo no fallback por coordenadas, se houver UN, prioriza a sequência numérica.
                    if un_index is not None:
                        valores_numericos = [p for p in textos[un_index + 1:] if _eh_numero_br(p)]
                        if len(valores_numericos) > 5:
                            registros.append({
                                "codigo": codigo,
                                "descricao": descricao,
                                "Saldo em Trânsito/ABERTO": br_to_float(valores_numericos[5]),
                            })
                            continue

                    candidatos = []
                    for w in linha_words[1:]:
                        txt = str(w.get("text", "")).strip()
                        if not _eh_numero_br(txt):
                            continue
                        cx = (float(w["x0"]) + float(w["x1"])) / 2
                        distancia = abs(cx - aberto_x)
                        candidatos.append((distancia, txt))

                    candidatos = [c for c in candidatos if c[0] <= 22]
                    if candidatos:
                        candidatos.sort(key=lambda x: x[0])
                        registros.append({
                            "codigo": codigo,
                            "descricao": descricao,
                            "Saldo em Trânsito/ABERTO": br_to_float(candidatos[0][1]),
                        })

    except Exception:
        registros = []

    if registros:
        return pd.DataFrame(registros).groupby("codigo", as_index=False).agg({
            "descricao": "first",
            "Saldo em Trânsito/ABERTO": "sum",
        })

    return pd.DataFrame(columns=["codigo", "descricao", "Saldo em Trânsito/ABERTO"])


# =========================================================
# LEITURA DO CSV DE CADASTRO DE PRODUTOS
# =========================================================

@st.cache_data(show_spinner="Lendo cadastro CSV...")
def ler_cadastro_produtos_csv(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()

    def mapear_colunas(colunas):
        colunas_norm = {normalizar_coluna(c): c for c in colunas}

        candidatos_codigo = [
            "CÓDIGO", "CODIGO", "CÓD.ITEM", "COD.ITEM", "CÓD ITEM", "COD ITEM",
            "CÓDIGO ITEM", "CODIGO ITEM",
        ]
        candidatos_descricao = [
            "DESCRIÇÃO DO ITEM", "DESCRICAO DO ITEM", "DESCRIÇÃO", "DESCRICAO",
            "DESC ITEM", "DESCRIÇÃO ITEM", "DESCRICAO ITEM",
        ]
        candidatos_fabrica = [
            "CÓD. FABRICA", "COD. FABRICA", "CÓD. FÁBRICA", "COD. FÁBRICA",
            "CÓDIGO DE FÁBRICA", "CODIGO DE FABRICA", "NOVO CÓDIGO DE FÁBRICA",
            "NOVO CODIGO DE FABRICA", "COD FABRICA", "CÓD FABRICA",
            "CÓDIGO FÁBRICA", "CODIGO FABRICA",
        ]

        def encontrar(candidatos):
            for candidato in candidatos:
                candidato_norm = normalizar_coluna(candidato)
                if candidato_norm in colunas_norm:
                    return colunas_norm[candidato_norm]
            return None

        candidatos_embalagem = [
            "EMBALAGEM", "EMB", "QTD EMBALAGEM", "QUANTIDADE EMBALAGEM",
            "QTDE EMBALAGEM", "QTD. EMBALAGEM", "MULTIPLO", "MÚLTIPLO",
        ]

        return {
            "codigo": encontrar(candidatos_codigo),
            "descricao": encontrar(candidatos_descricao),
            "codigo_fabrica": encontrar(candidatos_fabrica),
            "embalagem": encontrar(candidatos_embalagem),
        }

    tentativas = [
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "latin1"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "latin1"},
        {"sep": "\t", "encoding": "utf-8-sig"},
        {"sep": "\t", "encoding": "latin1"},
    ]

    df = None
    colunas_mapeadas = None
    ultimo_erro = None

    for tentativa in tentativas:
        try:
            uploaded_file.seek(0)
            temp = pd.read_csv(
                uploaded_file,
                sep=tentativa["sep"],
                encoding=tentativa["encoding"],
                dtype=str,
                engine="python",
                on_bad_lines="skip",
            )
            temp.columns = [str(c).strip() for c in temp.columns]
            mapeadas = mapear_colunas(temp.columns)
            if (
                mapeadas["codigo"]
                and mapeadas["descricao"]
                and mapeadas["codigo_fabrica"]
            ):
                df = temp
                colunas_mapeadas = mapeadas
                break
        except Exception as e:
            ultimo_erro = str(e)
            continue

    if df is None or colunas_mapeadas is None:
        st.error("Não consegui ler o CSV de cadastro.")
        st.caption(
            "O CSV pode ter um destes padrões de colunas: CÓDIGO, DESCRIÇÃO DO ITEM, CÓD. FABRICA "
            "ou Cód.Item, Descrição, Novo Código de fábrica."
        )
        if ultimo_erro:
            st.caption(f"Último erro identificado: {ultimo_erro}")
        return pd.DataFrame()

    colunas_selecionadas = [
        colunas_mapeadas["codigo"],
        colunas_mapeadas["descricao"],
        colunas_mapeadas["codigo_fabrica"],
    ]
    novos_nomes = ["codigo", "descricao_cadastro", "codigo_fabrica_cadastro"]

    if colunas_mapeadas.get("embalagem"):
        colunas_selecionadas.append(colunas_mapeadas["embalagem"])
        novos_nomes.append("embalagem")

    cadastro = df[colunas_selecionadas].copy()
    cadastro.columns = novos_nomes

    cadastro["codigo"] = cadastro["codigo"].astype(str).str.extract(r"(\d+)")[0].str.zfill(5)
    cadastro["descricao_cadastro"] = cadastro["descricao_cadastro"].astype(str).str.strip()
    cadastro["codigo_fabrica_cadastro"] = cadastro["codigo_fabrica_cadastro"].astype(str).str.strip()

    if "embalagem" in cadastro.columns:
        cadastro["embalagem"] = cadastro["embalagem"].apply(br_to_float)
        cadastro["embalagem"] = pd.to_numeric(cadastro["embalagem"], errors="coerce").fillna(0).round(0).astype(int)
    else:
        cadastro["embalagem"] = 0

    cadastro = cadastro.dropna(subset=["codigo"])
    cadastro = cadastro[cadastro["codigo"].str.lower() != "nan"]
    cadastro = cadastro.drop_duplicates(subset=["codigo"], keep="first")
    return cadastro


def normalizar_cadastro_produtos_df(df):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    colunas_norm = {normalizar_coluna(c): c for c in df.columns}

    def encontrar(candidatos):
        for candidato in candidatos:
            candidato_norm = normalizar_coluna(candidato)
            if candidato_norm in colunas_norm:
                return colunas_norm[candidato_norm]
        return None

    col_codigo = encontrar([
        "CODIGO", "CÓDIGO", "COD.ITEM", "CÓD.ITEM", "COD ITEM", "CÓD ITEM",
        "CODIGO ITEM", "CÓDIGO ITEM", "codigo",
    ])
    col_descricao = encontrar([
        "DESCRICAO DO ITEM", "DESCRIÇÃO DO ITEM", "DESCRICAO", "DESCRIÇÃO",
        "DESC ITEM", "DESCRICAO ITEM", "DESCRIÇÃO ITEM", "descricao",
    ])
    col_fabrica = encontrar([
        "COD. FABRICA", "CÓD. FABRICA", "COD. FÁBRICA", "CÓD. FÁBRICA",
        "CODIGO DE FABRICA", "CÓDIGO DE FÁBRICA", "NOVO CODIGO DE FABRICA",
        "NOVO CÓDIGO DE FÁBRICA", "COD FABRICA", "CÓD FABRICA",
        "CODIGO FABRICA", "CÓDIGO FÁBRICA", "codigo_fabrica",
    ])
    col_embalagem = encontrar([
        "EMBALAGEM", "EMB", "QTD EMBALAGEM", "QUANTIDADE EMBALAGEM",
        "QTDE EMBALAGEM", "QTD. EMBALAGEM", "MULTIPLO", "MÚLTIPLO", "embalagem",
    ])

    if not col_codigo or not col_descricao or not col_fabrica:
        return pd.DataFrame()

    colunas = [col_codigo, col_descricao, col_fabrica]
    nomes = ["codigo", "descricao_cadastro", "codigo_fabrica_cadastro"]
    if col_embalagem:
        colunas.append(col_embalagem)
        nomes.append("embalagem")

    cadastro = df[colunas].copy()
    cadastro.columns = nomes
    cadastro["codigo"] = cadastro["codigo"].astype(str).str.extract(r"(\d+)")[0].str.zfill(5)
    cadastro["descricao_cadastro"] = cadastro["descricao_cadastro"].astype(str).str.strip()
    cadastro["codigo_fabrica_cadastro"] = cadastro["codigo_fabrica_cadastro"].astype(str).str.strip()

    if "embalagem" in cadastro.columns:
        cadastro["embalagem"] = cadastro["embalagem"].apply(br_to_float)
        cadastro["embalagem"] = pd.to_numeric(cadastro["embalagem"], errors="coerce").fillna(0).round(0).astype(int)
    else:
        cadastro["embalagem"] = 0

    cadastro = cadastro.dropna(subset=["codigo"])
    cadastro = cadastro[cadastro["codigo"].astype(str).str.lower() != "nan"]
    cadastro = cadastro.drop_duplicates(subset=["codigo"], keep="first")
    return cadastro


@st.cache_data(show_spinner="Lendo cadastro do Google Sheets...", ttl=120)
def ler_cadastro_produtos_google_cached(_cache_key):
    _, sheets_service, _, auth_mode = google_get_services()
    recursos = google_get_resources()
    df_raw = google_read_df(sheets_service, recursos["cadastro_id"], "Cadastro")
    return normalizar_cadastro_produtos_df(df_raw)


def ler_cadastro_produtos_google():
    oauth_json = google_oauth_user_json()
    info_json = oauth_json or google_service_account_json()
    if not info_json:
        return pd.DataFrame()
    return ler_cadastro_produtos_google_cached(str(hash(info_json)))


def aplicar_cadastro_dataframe(df_giro, cadastro):
    if cadastro is None or cadastro.empty:
        return df_giro

    df = df_giro.merge(cadastro, on="codigo", how="left")
    df["descricao"] = df["descricao_cadastro"].where(
        df["descricao_cadastro"].notna()
        & (df["descricao_cadastro"].astype(str).str.strip() != "")
        & (df["descricao_cadastro"].astype(str).str.lower() != "nan"),
        df["descricao"],
    )
    df["codigo_fabrica"] = df["codigo_fabrica_cadastro"].where(
        df["codigo_fabrica_cadastro"].notna()
        & (df["codigo_fabrica_cadastro"].astype(str).str.strip() != "")
        & (df["codigo_fabrica_cadastro"].astype(str).str.lower() != "nan"),
        df["codigo_fabrica"],
    )
    if "embalagem" in df.columns:
        df["embalagem"] = pd.to_numeric(df["embalagem"], errors="coerce").fillna(0).round(0).astype(int)
    return df.drop(columns=["descricao_cadastro", "codigo_fabrica_cadastro"], errors="ignore")


def aplicar_cadastro(df_giro, cadastro_csv):
    cadastro = ler_cadastro_produtos_csv(cadastro_csv)
    if cadastro.empty:
        return df_giro

    return aplicar_cadastro_dataframe(df_giro, cadastro)

# =========================================================
# ÚLTIMA COMPRA / PREÇO
# =========================================================

def data_alerta_icon(data_ultima_compra, meses_alerta):
    if pd.isna(data_ultima_compra):
        return ""
    hoje = pd.Timestamp.today().normalize()
    limite = hoje - pd.DateOffset(months=int(meses_alerta))
    return "⚠️" if pd.Timestamp(data_ultima_compra) < limite else ""


def montar_info_compra(df_giro, meses_alerta_sem_compra=3):
    """
    Data de última compra: somente loja 009.
    Preço de última compra: prioriza loja 009; se não houver preço na 009, usa outra unidade com preço.
    """
    if df_giro.empty:
        return pd.DataFrame(columns=["codigo", "Data Última Compra", "Preço Última Compra"])

    df = df_giro.copy()
    df["dt_ult_compra"] = pd.to_datetime(df["dt_ult_compra"], errors="coerce", dayfirst=True)
    df["pr_ult_compra"] = pd.to_numeric(df["pr_ult_compra"], errors="coerce").fillna(0)

    resultados = []
    for codigo, grupo in df.groupby("codigo"):
        g009 = grupo[grupo["codigo_empresa"] == CODIGO_UNICA].copy()

        data_compra = pd.NaT
        data_compra_txt = ""
        if not g009.empty and g009["dt_ult_compra"].notna().any():
            idx = g009["dt_ult_compra"].idxmax()
            data_compra = g009.loc[idx, "dt_ult_compra"]
            raw = str(g009.loc[idx, "dt_ult_compra_txt"] or "").strip()
            data_compra_txt = raw if raw else format_data_br(data_compra)

        preco = 0.0
        precos_009 = g009[g009["pr_ult_compra"] > 0]["pr_ult_compra"] if not g009.empty else pd.Series(dtype=float)
        if not precos_009.empty:
            preco = float(precos_009.iloc[-1])
        else:
            precos_gerais = grupo[grupo["pr_ult_compra"] > 0]["pr_ult_compra"]
            if not precos_gerais.empty:
                preco = float(precos_gerais.iloc[-1])

        icone = data_alerta_icon(data_compra, meses_alerta_sem_compra)
        data_exibicao = f"{icone} {data_compra_txt}".strip() if data_compra_txt else icone

        resultados.append({
            "codigo": codigo,
            "Data Última Compra": data_exibicao,
            "Preço Última Compra": preco,
        })

    return pd.DataFrame(resultados)

# =========================================================
# MONTAGEM DAS TABELAS
# =========================================================

def arredondar_para_embalagem(sugestao, embalagem):
    """
    Arredonda a sugestão para cima, respeitando o múltiplo da embalagem.
    Ex.: sugestão 8 e embalagem 12 => 12; sugestão 20 e embalagem 12 => 24.
    """
    try:
        sugestao = int(math.ceil(float(sugestao or 0)))
        embalagem = int(round(float(embalagem or 0)))

        if sugestao <= 0:
            return 0
        if embalagem <= 0:
            return sugestao

        return int(math.ceil(sugestao / embalagem) * embalagem)
    except Exception:
        try:
            return int(math.ceil(float(sugestao or 0)))
        except Exception:
            return 0


def montar_tabela_consolidada(df_giro, df_transito=None, dias_estoque_alvo=60, meses_alerta_sem_compra=3):
    df_giro = df_giro.copy()

    # Garantias para evitar KeyError quando algum upload não trouxer cadastro/embalagem.
    # A tabela consolidada sempre precisa dessas colunas para o groupby/agg.
    colunas_padrao = {
        "codigo_fabrica": "",
        "embalagem": 0,
        "estoque": 0,
        "codigo_empresa": "",
        "descricao": "",
        "codigo": "",
    }
    for coluna, valor_padrao in colunas_padrao.items():
        if coluna not in df_giro.columns:
            df_giro[coluna] = valor_padrao

    df_lojas = df_giro[df_giro["codigo_empresa"].isin(CODIGOS_LOJAS)].copy()
    df_unica = df_giro[df_giro["codigo_empresa"] == CODIGO_UNICA].copy()

    meses_ref = [m for m in MESES if m in df_giro.columns]
    if not meses_ref:
        meses_ref = [m for m in MESES_PADRAO if m in df_giro.columns]

    agg = {
        "codigo_fabrica": "first",
        "embalagem": "first",
        **{mes: "sum" for mes in meses_ref},
        "estoque": "sum",
    }

    lojas = df_lojas.groupby(["codigo", "descricao"], as_index=False).agg(agg) if not df_lojas.empty else pd.DataFrame(columns=["codigo", "descricao", *agg.keys()])
    unica = df_unica.groupby(["codigo", "descricao"], as_index=False).agg(agg) if not df_unica.empty else pd.DataFrame(columns=["codigo", "descricao", *agg.keys()])

    lojas["Média Giro Lojas"] = lojas[meses_ref].mean(axis=1).round(1) if not lojas.empty and meses_ref else []
    unica["Média Giro Única"] = unica[meses_ref].mean(axis=1).round(1) if not unica.empty and meses_ref else []

    lojas = lojas.rename(columns={
        "codigo_fabrica": "Código Fábrica",
        "embalagem": "Embalagem",
        **{mes: col_giro("Giro Lojas", mes) for mes in meses_ref},
        "estoque": "Estoque Lojas",
    })
    unica = unica.rename(columns={
        "codigo_fabrica": "Código Fábrica Única",
        "embalagem": "Embalagem Única",
        **{mes: col_giro("Giro Única", mes) for mes in meses_ref},
        "estoque": "Estoque Única",
    })

    resumo = pd.merge(lojas, unica, on=["codigo", "descricao"], how="outer").fillna(0)

    for col in ["Código Fábrica", "Código Fábrica Única"]:
        if col not in resumo.columns:
            resumo[col] = ""
        resumo[col] = resumo[col].replace(0, "")

    resumo["Código Fábrica"] = resumo.apply(
        lambda x: x["Código Fábrica"] if x["Código Fábrica"] else x["Código Fábrica Única"], axis=1
    )

    for col in ["Embalagem", "Embalagem Única"]:
        if col not in resumo.columns:
            resumo[col] = 0
        resumo[col] = pd.to_numeric(resumo[col], errors="coerce").fillna(0).round(0).astype(int)

    resumo["Embalagem"] = resumo.apply(
        lambda x: int(x["Embalagem"]) if int(x["Embalagem"]) > 0 else int(x["Embalagem Única"]),
        axis=1,
    )

    colunas_giro_geral = []
    for mes in meses_ref:
        label = label_mes_giro(mes)
        for prefixo in ["Giro Lojas", "Giro Única"]:
            col = f"{prefixo} {label}"
            if col not in resumo.columns:
                resumo[col] = 0
        col_geral = f"Giro Geral {label}"
        resumo[col_geral] = resumo[f"Giro Lojas {label}"] + resumo[f"Giro Única {label}"]
        colunas_giro_geral.append(col_geral)

    resumo["Média Giro Geral"] = resumo[colunas_giro_geral].mean(axis=1).round(1) if colunas_giro_geral else 0

    for col in ["Estoque Lojas", "Estoque Única", "Média Giro Lojas", "Média Giro Única"]:
        if col not in resumo.columns:
            resumo[col] = 0

    resumo["Estoque Atual Geral"] = resumo["Estoque Lojas"] + resumo["Estoque Única"]
    resumo["Estoque Geral"] = resumo["Estoque Atual Geral"]

    if df_transito is not None and not df_transito.empty:
        df_transito = df_transito.copy()
        df_transito["codigo"] = df_transito["codigo"].astype(str).str.extract(r"(\d+)", expand=False).fillna("").str.zfill(5)
        resumo = pd.merge(resumo, df_transito.drop(columns=["descricao"], errors="ignore"), on="codigo", how="left")
    else:
        resumo["Saldo em Trânsito/ABERTO"] = 0

    resumo["Saldo em Trânsito/ABERTO"] = pd.to_numeric(resumo["Saldo em Trânsito/ABERTO"], errors="coerce").fillna(0)
    resumo["Estoque Final"] = resumo["Estoque Atual Geral"] + resumo["Saldo em Trânsito/ABERTO"]
    resumo["Estoque Alvo"] = resumo["Média Giro Geral"] * (dias_estoque_alvo / 30)
    resumo["Sugestão Sistema"] = (resumo["Estoque Alvo"] - resumo["Estoque Final"]).apply(lambda x: max(math.ceil(x), 0)).astype(int)
    resumo["Sugestão arredondada"] = resumo.apply(
        lambda row: arredondar_para_embalagem(row["Sugestão Sistema"], row.get("Embalagem", 0)),
        axis=1,
    ).astype(int)

    info_compra = montar_info_compra(df_giro, meses_alerta_sem_compra)
    resumo = pd.merge(resumo, info_compra, on="codigo", how="left")
    resumo["Preço Última Compra"] = pd.to_numeric(resumo["Preço Última Compra"], errors="coerce").fillna(0)

    resumo = resumo.drop(columns=["Código Fábrica Única", "Embalagem Única"], errors="ignore")
    return resumo.sort_values("descricao").reset_index(drop=True)


def montar_detalhe_produto(df_giro, codigo_produto):
    detalhe = df_giro[df_giro["codigo"] == codigo_produto].copy()
    if detalhe.empty:
        return pd.DataFrame()

    meses_ref = [m for m in MESES if m in detalhe.columns]
    detalhe["Média Giro"] = detalhe[meses_ref].mean(axis=1).round(1) if meses_ref else 0
    rename_meses = {mes: label_mes_giro(mes) for mes in meses_ref}
    detalhe = detalhe.rename(columns={
        "loja": "Unidade",
        "codigo_empresa": "Cód. Empresa",
        **rename_meses,
        "estoque": "Saldo em Estoque",
    })

    colunas = ["Cód. Empresa", "Unidade"] + [label_mes_giro(m) for m in meses_ref] + ["Média Giro", "Saldo em Estoque"]
    return detalhe[colunas].sort_values(["Cód. Empresa", "Unidade"])

# =========================================================
# FORMATAÇÃO / EXPORTAÇÃO
# =========================================================

def filtrar_tabela(df, campos, key):
    busca = st.text_input("Pesquisar", key=key)
    if not busca:
        return df.copy()
    busca = busca.lower()
    filtro = pd.Series(False, index=df.index)
    for campo in campos:
        if campo in df.columns:
            filtro = filtro | df[campo].astype(str).str.lower().str.contains(busca, na=False)
    return df[filtro].copy()


def colorir_colunas_consolidada(col):
    if "Lojas" in col.name:
        return ["background-color: #e8f1ff"] * len(col)
    if "Única" in col.name:
        return ["background-color: #fff1df"] * len(col)
    if "Geral" in col.name:
        return ["background-color: #eaf7ea"] * len(col)
    if "ABERTO" in col.name or "Estoque Final" in col.name:
        return ["background-color: #f3e8ff; font-weight: 600"] * len(col)
    if "Sistema" in col.name or "arredondada" in col.name or "Alvo" in col.name or "PEDIDO Final" in col.name:
        return ["background-color: #ffe8e8"] * len(col)
    return [""] * len(col)


def colorir_colunas_pedido(col):
    if col.name in ["Média Giro Lojas", "Estoque Lojas"]:
        return ["background-color: #e8f1ff"] * len(col)
    if col.name in ["Média Giro Única", "Estoque Única"]:
        return ["background-color: #fff1df"] * len(col)
    if col.name in ["Média Giro Geral", "Estoque Geral"]:
        return ["background-color: #eaf7ea"] * len(col)
    if col.name in ["Saldo em Trânsito/ABERTO", "Estoque Final"]:
        return ["background-color: #f3e8ff; font-weight: 600"] * len(col)
    if col.name in ["Estoque Alvo", "Sugestão Sistema", "Sugestão arredondada", "PEDIDO Final", "Valor Final do Pedido"]:
        return ["background-color: #ffe8e8"] * len(col)
    return [""] * len(col)


def colunas_giro_geral_mensal(df):
    return [c for c in df.columns if str(c).startswith("Giro Geral ")]


def estilos_alerta_giro_fora_curva(row):
    estilos = pd.Series("", index=row.index)
    cols_giro = colunas_giro_geral_mensal(pd.DataFrame(columns=row.index))
    if len(cols_giro) < 2:
        return estilos

    valores = pd.to_numeric(row[cols_giro], errors="coerce").fillna(0)
    positivos = valores[valores > 0]

    colunas_alerta = set()
    if len(positivos) == 1:
        colunas_alerta.add(positivos.index[0])
    elif len(positivos) >= 2:
        for col, valor in valores.items():
            if valor <= 0:
                continue
            outros = valores.drop(index=col)
            media_outros = outros[outros > 0].mean()
            if pd.notna(media_outros) and media_outros > 0 and valor >= media_outros * 2.2:
                colunas_alerta.add(col)

    for col in colunas_alerta:
        estilos[col] = "color: #c2410c; font-weight: 700"
    return estilos

def formatadores_para_tabela(df):
    fmt = {}
    dinheiro = [c for c in df.columns if "Preço" in c or "Valor" in c]
    inteiros = [c for c in df.columns if c in ["Sugestão Sistema", "Sugestão arredondada", "PEDIDO Final", "Embalagem"]]
    for col in df.columns:
        if col in dinheiro:
            fmt[col] = format_moeda_br
        elif col in inteiros:
            fmt[col] = format_int_br
        elif pd.api.types.is_numeric_dtype(df[col]):
            fmt[col] = lambda x: format_num_br(x, 1)
    return fmt


def render_tabela_interativa_colorida(df, height=650):
    styled = df.style.apply(colorir_colunas_consolidada, axis=0).apply(estilos_alerta_giro_fora_curva, axis=1).format(formatadores_para_tabela(df))
    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        height=height,
        column_config={
            "codigo": st.column_config.TextColumn("Código", width="small", pinned=True),
            "descricao": st.column_config.TextColumn("Descrição", width="large", pinned=True),
            "Código Fábrica": st.column_config.TextColumn("Código Fábrica", width="medium"),
        },
    )


def gerar_csv(df):
    return df.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")


# =========================================================
# INTEGRACAO GOOGLE DRIVE / SHEETS
# =========================================================

def google_configurado():
    if build is None or MediaIoBaseUpload is None:
        return False
    try:
        return (
            "google_oauth_user" in st.secrets
            or "GOOGLE_OAUTH_USER_JSON" in st.secrets
            or "google_service_account" in st.secrets
            or "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets
        )
    except Exception:
        return False


def google_mensagem_configuracao():
    if build is None or MediaIoBaseUpload is None:
        return (
            "Instale as dependencias do Google no ambiente: "
            "google-api-python-client, google-auth e google-auth-httplib2."
        )
    return (
        "Configure a autenticação Google em st.secrets. Preferencial: [google_oauth_user] "
        "com client_id, client_secret e refresh_token da conta dona da pasta. "
        "Conta de serviço pode ler/escrever arquivos existentes, mas normalmente não consegue criar "
        "Google Sheets em Meu Drive sem gerar erro de cota/permissão."
    )


def google_service_account_json():
    try:
        if "google_service_account" in st.secrets:
            return json.dumps(dict(st.secrets["google_service_account"]), sort_keys=True)
        if "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets:
            raw = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
            if isinstance(raw, str):
                json.loads(raw)
                return raw
            return json.dumps(dict(raw), sort_keys=True)
    except Exception:
        return ""
    return ""


def google_oauth_user_json():
    """
    Autenticação OAuth da conta dona da pasta, ex.: gdautotintas@gmail.com.
    Use este formato nos Secrets do Streamlit:

    [google_oauth_user]
    client_id = "..."
    client_secret = "..."
    refresh_token = "..."
    token_uri = "https://oauth2.googleapis.com/token"

    Esse modo cria Google Sheets usando a cota/permissão da conta dona,
    evitando os erros storageQuotaExceeded e caller does not have permission
    típicos de conta de serviço em Meu Drive.
    """
    try:
        if "google_oauth_user" in st.secrets:
            return json.dumps(dict(st.secrets["google_oauth_user"]), sort_keys=True)
        if "GOOGLE_OAUTH_USER_JSON" in st.secrets:
            raw = st.secrets["GOOGLE_OAUTH_USER_JSON"]
            if isinstance(raw, str):
                json.loads(raw)
                return raw
            return json.dumps(dict(raw), sort_keys=True)
    except Exception:
        return ""
    return ""


@st.cache_resource(show_spinner=False)
def google_get_services_cached(auth_mode, info_json):
    info = json.loads(info_json)
    scopes = [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]

    if auth_mode == "oauth_user":
        if UserCredentials is None:
            raise RuntimeError("Biblioteca google-auth sem suporte a OAuth user credentials.")
        credentials = UserCredentials(
            token=info.get("token"),
            refresh_token=info.get("refresh_token"),
            token_uri=info.get("token_uri", "https://oauth2.googleapis.com/token"),
            client_id=info.get("client_id"),
            client_secret=info.get("client_secret"),
            scopes=scopes,
        )
        client_email = info.get("user_email") or info.get("email") or "oauth_user"
    else:
        if service_account is None:
            raise RuntimeError("Biblioteca google-auth sem suporte a service_account.")
        credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
        client_email = info.get("client_email", "service_account")

    drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    sheets_service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    return drive_service, sheets_service, client_email, auth_mode


def google_get_services():
    oauth_json = google_oauth_user_json()
    if oauth_json:
        return google_get_services_cached("oauth_user", oauth_json)

    info_json = google_service_account_json()
    if info_json:
        return google_get_services_cached("service_account", info_json)

    raise RuntimeError(google_mensagem_configuracao())

def google_link_pasta(folder_id):
    return f"https://drive.google.com/drive/folders/{folder_id}"


def google_link_planilha(spreadsheet_id):
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"


def google_link_arquivo(file_id):
    return f"https://drive.google.com/file/d/{file_id}/view"


def google_safe_name(nome):
    nome = str(nome or "").strip()
    nome = re.sub(r"[\\/:*?\"<>|]+", "-", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome or f"Pedido {datetime.now().strftime('%Y-%m-%d %H.%M')}"


def google_q_text(value):
    return str(value).replace("\\", "\\\\").replace("'", "\\'")


def google_find_file(drive_service, name, parent_id, mime_type=None):
    partes = [
        f"name = '{google_q_text(name)}'",
        f"'{google_q_text(parent_id)}' in parents",
        "trashed = false",
    ]
    if mime_type:
        partes.append(f"mimeType = '{google_q_text(mime_type)}'")
    result = drive_service.files().list(
        q=" and ".join(partes),
        spaces="drive",
        fields="files(id, name, mimeType, webViewLink)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = result.get("files", [])
    return files[0] if files else None


def google_ensure_folder(drive_service, name, parent_id):
    mime = "application/vnd.google-apps.folder"
    existente = google_find_file(drive_service, name, parent_id, mime)
    if existente:
        return existente["id"]
    criado = drive_service.files().create(
        body={"name": name, "mimeType": mime, "parents": [parent_id]},
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return criado["id"]


def google_get_file_parents(drive_service, file_id):
    try:
        meta = drive_service.files().get(
            fileId=file_id,
            fields="parents",
            supportsAllDrives=True,
        ).execute()
        return ",".join(meta.get("parents", []) or [])
    except Exception:
        return ""


def google_mover_arquivo_para_pasta(drive_service, file_id, parent_id):
    """
    Move o arquivo para a pasta raiz configurada.

    Importante:
    - A planilha é criada pela Google Sheets API.
    - Depois ela é vinculada explicitamente à pasta do Drive informada.
    - Isso evita a criação solta na raiz da conta de serviço.
    """
    parents_atuais = google_get_file_parents(drive_service, file_id)
    drive_service.files().update(
        fileId=file_id,
        addParents=parent_id,
        removeParents=parents_atuais or None,
        fields="id, parents, webViewLink",
        supportsAllDrives=True,
    ).execute()


def google_ensure_spreadsheet(drive_service, sheets_service, name, parent_id):
    """
    Garante uma planilha Google Sheets dentro da pasta informada.

    Correção aplicada:
    - Antes, a planilha era criada diretamente pela Drive API com mimeType de Sheets.
    - Em algumas contas de serviço isso causa erro 403 storageQuotaExceeded.
    - Agora a planilha é criada pela Sheets API e imediatamente movida para a pasta.
    """
    mime = "application/vnd.google-apps.spreadsheet"
    existente = google_find_file(drive_service, name, parent_id, mime)
    if existente:
        return existente["id"]

    try:
        criado = sheets_service.spreadsheets().create(
            body={"properties": {"title": name}},
            fields="spreadsheetId",
        ).execute()
        spreadsheet_id = criado["spreadsheetId"]
        google_mover_arquivo_para_pasta(drive_service, spreadsheet_id, parent_id)
        return spreadsheet_id
    except Exception as e:
        erro = str(e)
        if (
            "storageQuotaExceeded" in erro
            or "Drive storage quota" in erro
            or "The caller does not have permission" in erro
            or "caller does not have permission" in erro
        ):
            raise RuntimeError(
                "O Google bloqueou a criação da planilha com a autenticação atual. "
                "Para criar Google Sheets automaticamente dentro de uma pasta de Meu Drive, "
                "use OAuth da conta dona da pasta em [google_oauth_user] nos Secrets do Streamlit. "
                "Conta de serviço funciona para acessar pastas compartilhadas, mas costuma falhar ao criar "
                "arquivos nativos do Google Sheets em Meu Drive. Erro original: " + erro
            )
        raise


def google_ensure_sheet_tab(sheets_service, spreadsheet_id, sheet_name):
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(title))",
    ).execute()
    existentes = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if sheet_name in existentes:
        return
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
    ).execute()


def google_prepare_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return format_data_br(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return str(value)


def google_df_to_values(df):
    df = df.copy() if df is not None else pd.DataFrame()
    return [list(df.columns)] + [
        [google_prepare_value(row.get(col, "")) for col in df.columns]
        for _, row in df.iterrows()
    ]


def google_write_df(sheets_service, spreadsheet_id, sheet_name, df):
    google_ensure_sheet_tab(sheets_service, spreadsheet_id, sheet_name)
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A:ZZ",
        body={},
    ).execute()
    values = google_df_to_values(df)
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()


def google_read_df(sheets_service, spreadsheet_id, sheet_name=None):
    range_name = f"'{sheet_name}'!A:ZZ" if sheet_name else "A:ZZ"
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name,
    ).execute()
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()
    header = [str(c).strip() for c in values[0]]
    rows = values[1:]
    largura = len(header)
    rows = [row + [""] * (largura - len(row)) if len(row) < largura else row[:largura] for row in rows]
    return pd.DataFrame(rows, columns=header)


def google_append_rows(sheets_service, spreadsheet_id, sheet_name, rows):
    google_ensure_sheet_tab(sheets_service, spreadsheet_id, sheet_name)
    sheets_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_name}'!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


def google_ensure_headers(sheets_service, spreadsheet_id, sheet_name, columns):
    google_ensure_sheet_tab(sheets_service, spreadsheet_id, sheet_name)
    atual = google_read_df(sheets_service, spreadsheet_id, sheet_name)
    if atual.empty and list(atual.columns) == []:
        google_write_df(sheets_service, spreadsheet_id, sheet_name, pd.DataFrame(columns=columns))


@st.cache_data(show_spinner=False, ttl=60)
def google_get_resources_cached(_cache_key):
    drive_service, sheets_service, client_email, auth_mode = google_get_services()
    pedidos_folder_id = google_ensure_folder(drive_service, GOOGLE_SUBPASTA_PEDIDOS, GOOGLE_DRIVE_ROOT_FOLDER_ID)
    finais_folder_id = google_ensure_folder(drive_service, GOOGLE_SUBPASTA_FINAIS, GOOGLE_DRIVE_ROOT_FOLDER_ID)
    controle_id = google_ensure_spreadsheet(drive_service, sheets_service, GOOGLE_PLANILHA_CONTROLE, GOOGLE_DRIVE_ROOT_FOLDER_ID)
    cadastro_id = google_ensure_spreadsheet(drive_service, sheets_service, GOOGLE_PLANILHA_CADASTRO, GOOGLE_DRIVE_ROOT_FOLDER_ID)
    google_ensure_headers(sheets_service, controle_id, "Pedidos", GOOGLE_PEDIDOS_COLUNAS)
    google_ensure_headers(sheets_service, controle_id, "Acompanhamento", GOOGLE_ACOMPANHAMENTO_COLUNAS)
    google_ensure_headers(sheets_service, cadastro_id, "Cadastro", ["codigo", "descricao", "codigo_fabrica", "embalagem"])
    return {
        "client_email": client_email,
        "pedidos_folder_id": pedidos_folder_id,
        "finais_folder_id": finais_folder_id,
        "controle_id": controle_id,
        "cadastro_id": cadastro_id,
        "root_link": google_link_pasta(GOOGLE_DRIVE_ROOT_FOLDER_ID),
        "pedidos_link": google_link_pasta(pedidos_folder_id),
        "finais_link": google_link_pasta(finais_folder_id),
        "controle_link": google_link_planilha(controle_id),
        "cadastro_link": google_link_planilha(cadastro_id),
    }


def google_get_resources():
    oauth_json = google_oauth_user_json()
    info_json = oauth_json or google_service_account_json()
    if not info_json:
        raise RuntimeError(google_mensagem_configuracao())
    return google_get_resources_cached(str(hash(info_json)))


def google_criar_planilha_pedido(nome_pedido, fornecedor, pedido_df, criado_por=""):
    drive_service, sheets_service, _, auth_mode = google_get_services()
    recursos = google_get_resources()
    nome_limpo = google_safe_name(nome_pedido)
    fornecedor_limpo = google_safe_name(fornecedor)
    titulo = f"{datetime.now().strftime('%Y-%m-%d')} - {fornecedor_limpo} - {nome_limpo}"
    spreadsheet_id = google_ensure_spreadsheet(drive_service, sheets_service, titulo, recursos["pedidos_folder_id"])

    df_export = pedido_df.copy()
    if "zx" not in df_export.columns:
        df_export.insert(0, "zx", df_export.get("codigo", ""))
    google_write_df(sheets_service, spreadsheet_id, "Pedido", df_export)
    google_write_df(sheets_service, spreadsheet_id, "Aprovacao", pd.DataFrame([{
        "status": "Em edicao",
        "aprovado_por": "",
        "aprovado_em": "",
        "observacao": "",
    }]))

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    pedido_id = datetime.now().strftime("%Y%m%d%H%M%S")
    link = google_link_planilha(spreadsheet_id)
    valor = totalizar_valor_pedido(df_export)
    google_append_rows(sheets_service, recursos["controle_id"], "Pedidos", [[
        pedido_id, nome_limpo, fornecedor_limpo, "Em edicao", round(float(valor or 0), 2),
        agora, criado_por, "", "", link, spreadsheet_id, "", "", "",
    ]])
    return {"pedido_id": pedido_id, "spreadsheet_id": spreadsheet_id, "link": link, "titulo": titulo}


def google_listar_pedidos():
    _, sheets_service, _, auth_mode = google_get_services()
    recursos = google_get_resources()
    df = google_read_df(sheets_service, recursos["controle_id"], "Pedidos")
    for col in GOOGLE_PEDIDOS_COLUNAS:
        if col not in df.columns:
            df[col] = ""
    return df[GOOGLE_PEDIDOS_COLUNAS]


def google_salvar_pedidos_controle(df):
    _, sheets_service, _, auth_mode = google_get_services()
    recursos = google_get_resources()
    df = df.copy()
    for col in GOOGLE_PEDIDOS_COLUNAS:
        if col not in df.columns:
            df[col] = ""
    google_write_df(sheets_service, recursos["controle_id"], "Pedidos", df[GOOGLE_PEDIDOS_COLUNAS])


def google_atualizar_status_pedido(pedido_id, status, usuario="", observacao="", link_autcom="", link_fornecedor=""):
    df = google_listar_pedidos()
    mask = df["id_pedido"].astype(str) == str(pedido_id)
    if not mask.any():
        raise ValueError("Pedido nao encontrado no controle.")
    idx = df[mask].index[0]
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    df.loc[idx, "status"] = status
    if status.lower().startswith("aprov"):
        df.loc[idx, "aprovado_em"] = agora
        df.loc[idx, "aprovado_por"] = usuario
    if observacao:
        df.loc[idx, "observacao"] = observacao
    if link_autcom:
        df.loc[idx, "link_autcom"] = link_autcom
    if link_fornecedor:
        df.loc[idx, "link_fornecedor"] = link_fornecedor
    google_salvar_pedidos_controle(df)
    return df.loc[idx].to_dict()


def google_ler_pedido_drive(spreadsheet_id):
    _, sheets_service, _, auth_mode = google_get_services()
    df = google_read_df(sheets_service, spreadsheet_id, "Pedido")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def google_upload_bytes(nome_arquivo, dados, mime_type, folder_id):
    drive_service, _, _, auth_mode = google_get_services()
    media = MediaIoBaseUpload(BytesIO(dados), mimetype=mime_type, resumable=False)
    criado = drive_service.files().create(
        body={"name": nome_arquivo, "parents": [folder_id]},
        media_body=media,
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return criado.get("webViewLink") or google_link_arquivo(criado["id"])


def google_registrar_acompanhamento(pedido_info):
    _, sheets_service, _, auth_mode = google_get_services()
    recursos = google_get_resources()
    data_atual = datetime.now().strftime("%d/%m/%Y")
    mes_atual = datetime.now().strftime("%m/%Y")
    google_append_rows(sheets_service, recursos["controle_id"], "Acompanhamento", [[
        data_atual,
        mes_atual,
        pedido_info.get("fornecedor", ""),
        pedido_info.get("nome_pedido", ""),
        pedido_info.get("valor", ""),
        pedido_info.get("status", ""),
        pedido_info.get("link_pedido", ""),
        pedido_info.get("link_autcom", ""),
        pedido_info.get("link_fornecedor", ""),
    ]])


def google_finalizar_pedido(pedido_id, df_tratamento, usuario=""):
    recursos = google_get_resources()
    pedidos = google_listar_pedidos()
    row = pedidos[pedidos["id_pedido"].astype(str) == str(pedido_id)]
    if row.empty:
        raise ValueError("Pedido nao encontrado.")
    row = row.iloc[0].to_dict()
    base_nome = google_safe_name(row.get("nome_pedido") or "pedido")
    autcom_bytes = gerar_excel_autcom_tratamento(df_tratamento)
    fornecedor_bytes = gerar_excel_fornecedor_tratamento(df_tratamento)
    link_autcom = google_upload_bytes(
        f"{base_nome} - importacao autcom.xlsx",
        autcom_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        recursos["finais_folder_id"],
    )
    link_fornecedor = google_upload_bytes(
        f"{base_nome} - envio fornecedor.xlsx",
        fornecedor_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        recursos["finais_folder_id"],
    )
    atualizado = google_atualizar_status_pedido(
        pedido_id,
        "Finalizado",
        usuario=usuario,
        link_autcom=link_autcom,
        link_fornecedor=link_fornecedor,
    )
    google_registrar_acompanhamento(atualizado)
    return link_autcom, link_fornecedor


def gerar_excel_pedido_editavel(df):
    """
    Gera uma planilha Excel editável do pedido.
    Recursos aplicados:
    - Valor Final do Pedido formulado: PEDIDO Final x Preço Última Compra.
    - Painéis congelados para facilitar navegação.
    - Coluna Total Geral do Pedido ao lado do Valor Final, pintada em amarelo.

    Observação: CSV não suporta fórmulas, congelamento de painéis nem pintura de células.
    Por isso este download é gerado em .xlsx.
    """
    if Workbook is None:
        raise RuntimeError("A biblioteca openpyxl não está instalada. Rode: python -m pip install openpyxl")

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    df_export = df.copy()

    colunas = list(df_export.columns)
    if "Valor Final do Pedido" not in colunas:
        df_export["Valor Final do Pedido"] = 0
        colunas = list(df_export.columns)

    # Garante que a coluna Total fique exatamente ao lado de Valor Final do Pedido.
    if "Total Geral do Pedido" in df_export.columns:
        df_export = df_export.drop(columns=["Total Geral do Pedido"])

    pos_valor = list(df_export.columns).index("Valor Final do Pedido")
    colunas = list(df_export.columns)
    colunas.insert(pos_valor + 1, "Total Geral do Pedido")
    df_export["Total Geral do Pedido"] = ""
    df_export = df_export[colunas]

    wb = Workbook()
    ws = wb.active
    ws.title = "Pedido Editável"

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, col_name in enumerate(df_export.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.fill = total_fill if col_name == "Total Geral do Pedido" else header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    idx_pedido = df_export.columns.get_loc("PEDIDO Final") + 1 if "PEDIDO Final" in df_export.columns else None
    idx_preco = df_export.columns.get_loc("Preço Última Compra") + 1 if "Preço Última Compra" in df_export.columns else None
    idx_valor = df_export.columns.get_loc("Valor Final do Pedido") + 1 if "Valor Final do Pedido" in df_export.columns else None
    idx_total = df_export.columns.get_loc("Total Geral do Pedido") + 1

    for row_idx, (_, row) in enumerate(df_export.iterrows(), start=2):
        for col_idx, col_name in enumerate(df_export.columns, start=1):
            if col_name == "Valor Final do Pedido" and idx_pedido and idx_preco:
                pedido_col = get_column_letter(idx_pedido)
                preco_col = get_column_letter(idx_preco)
                value = f"={pedido_col}{row_idx}*{preco_col}{row_idx}"
            elif col_name == "Total Geral do Pedido":
                value = ""
            else:
                value = row.get(col_name, "")

            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

            if col_name == "Total Geral do Pedido":
                cell.fill = total_fill
            if col_name in ["Preço Última Compra", "Valor Final do Pedido", "Total Geral do Pedido"]:
                cell.number_format = 'R$ #,##0.00'
            elif col_name in ["PEDIDO Final", "Sugestão Sistema", "Sugestão arredondada", "Embalagem"]:
                cell.number_format = '0'
            elif isinstance(value, (int, float)):
                cell.number_format = '#,##0.0'

    ultima_linha = max(ws.max_row, 2)
    valor_col = get_column_letter(idx_valor) if idx_valor else None
    total_col = get_column_letter(idx_total)
    if valor_col:
        ws.cell(row=2, column=idx_total, value=f"=SUM({valor_col}2:{valor_col}{ultima_linha})")
        ws.cell(row=2, column=idx_total).fill = total_fill
        ws.cell(row=2, column=idx_total).font = Font(bold=True)
        ws.cell(row=2, column=idx_total).number_format = 'R$ #,##0.00'

    # Congela cabeçalho e as primeiras colunas de identificação.
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(df_export.columns, start=1):
        letter = get_column_letter(col_idx)
        if col_name == "descricao":
            ws.column_dimensions[letter].width = 42
        elif col_name in ["Código Fábrica", "Data Última Compra", "Origem Sugestão"]:
            ws.column_dimensions[letter].width = 18
        elif col_name in ["Valor Final do Pedido", "Total Geral do Pedido", "Preço Última Compra"]:
            ws.column_dimensions[letter].width = 20
        else:
            ws.column_dimensions[letter].width = max(12, min(22, len(str(col_name)) + 2))

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def gerar_copia_fornecedor_csv(df):
    if df is None or df.empty:
        fornecedor = pd.DataFrame(columns=["Código Fábrica", "Descrição", "Quantidade"])
    else:
        fornecedor = df.copy()
        fornecedor["PEDIDO Final"] = pd.to_numeric(fornecedor.get("PEDIDO Final", 0), errors="coerce").fillna(0).round(0).astype(int)
        fornecedor = fornecedor[fornecedor["PEDIDO Final"] > 0].copy()
        for col in ["Código Fábrica", "descricao"]:
            if col not in fornecedor.columns:
                fornecedor[col] = ""
        fornecedor = fornecedor[["Código Fábrica", "descricao", "PEDIDO Final"]].rename(columns={
            "descricao": "Descrição",
            "PEDIDO Final": "Quantidade",
        })
    return fornecedor.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")


def gerar_excel_pedido(df_pedido):
    """
    Excel para importação no Autcom, sem cabeçalho:
    Coluna B = código
    Coluna F = quantidade
    Coluna H = valor unitário
    """
    if Workbook is None:
        raise RuntimeError("A biblioteca openpyxl não está instalada. Rode: python -m pip install openpyxl")

    wb = Workbook()
    ws = wb.active
    ws.title = "Pedido"

    linha_excel = 1
    for _, row in df_pedido.iterrows():
        qtd = int(round(float(row.get("PEDIDO Final", 0) or 0)))
        if qtd <= 0:
            continue
        ws.cell(row=linha_excel, column=2, value=str(row.get("codigo", "")).zfill(5))
        ws.cell(row=linha_excel, column=6, value=qtd)
        ws.cell(row=linha_excel, column=8, value=round(float(str(row.get("Preço Última Compra", 0)).replace(",", "." ) or 0), 2))
        linha_excel += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()



def ler_planilha_tratamento_pedido(uploaded_file):
    """
    Lê a planilha final editável enviada pelo usuário na página Tratamento de Pedido Final.
    Aceita .xlsx, .xls e .csv.
    """
    if uploaded_file is None:
        return pd.DataFrame()

    nome = str(getattr(uploaded_file, "name", "")).lower()

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    if nome.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, dtype=str)

    tentativas = [
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "latin1"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "latin1"},
        {"sep": "\t", "encoding": "utf-8-sig"},
        {"sep": "\t", "encoding": "latin1"},
    ]

    ultimo_erro = None
    for tentativa in tentativas:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(
                uploaded_file,
                sep=tentativa["sep"],
                encoding=tentativa["encoding"],
                dtype=str,
                engine="python",
                on_bad_lines="skip",
            )
        except Exception as e:
            ultimo_erro = str(e)
            continue

    raise RuntimeError(f"Não consegui ler a planilha enviada. Último erro: {ultimo_erro}")


def gerar_excel_autcom_tratamento(df_tratamento):
    """
    Gera o Excel para importação no Autcom a partir da planilha de Tratamento de Pedido Final.
    Sem cabeçalho:
    - Coluna B = código da coluna zx
    - Coluna F = quantidade da coluna PEDIDO Final
    - Coluna H = preço da coluna Preço Última Compra
    """
    if Workbook is None:
        raise RuntimeError("A biblioteca openpyxl não está instalada. Rode: python -m pip install openpyxl")

    df = df_tratamento.copy()
    df.columns = [str(c).strip() for c in df.columns]

    colunas_norm = {normalizar_coluna(c): c for c in df.columns}

    col_codigo = colunas_norm.get("ZX") or colunas_norm.get("CODIGO") or colunas_norm.get("CÓDIGO")
    col_qtd = colunas_norm.get("PEDIDO FINAL")
    col_preco = colunas_norm.get("PREÇO ÚLTIMA COMPRA") or colunas_norm.get("PRECO ULTIMA COMPRA")

    faltantes = []
    if not col_codigo:
        faltantes.append("zx")
    if not col_qtd:
        faltantes.append("PEDIDO Final")
    if not col_preco:
        faltantes.append("Preço Última Compra")

    if faltantes:
        raise ValueError("A planilha enviada não possui as colunas obrigatórias: " + ", ".join(faltantes))

    wb = Workbook()
    ws = wb.active
    ws.title = "Pedido"

    linha_excel = 1
    for _, row in df.iterrows():
        codigo_raw = str(row.get(col_codigo, "")).strip()
        codigo_match = re.search(r"(\d+)", codigo_raw)
        codigo = codigo_match.group(1).zfill(5) if codigo_match else ""

        qtd = br_to_float(row.get(col_qtd, 0))
        preco = numero_planilha_para_float(row.get(col_preco, 0))

        try:
            qtd = int(round(float(qtd)))
        except Exception:
            qtd = 0

        if not codigo or qtd <= 0:
            continue

        ws.cell(row=linha_excel, column=2, value=codigo)
        ws.cell(row=linha_excel, column=6, value=qtd)
        ws.cell(row=linha_excel, column=8, value=round(float(preco or 0), 2))
        ws.cell(row=linha_excel, column=8).number_format = 'R$ #,##0.00'
        linha_excel += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def gerar_excel_fornecedor_tratamento(df_tratamento):
    """
    Gera Excel para envio ao fornecedor a partir da planilha de Tratamento Final:
    A = codigo, B = descricao, C = Codigo de fabrica, D = quantidade.
    """
    if Workbook is None:
        raise RuntimeError("A biblioteca openpyxl não está instalada. Rode: python -m pip install openpyxl")

    df = df_tratamento.copy()
    df.columns = [str(c).strip() for c in df.columns]
    colunas_norm = {normalizar_coluna(c): c for c in df.columns}

    col_codigo = colunas_norm.get("ZX") or colunas_norm.get("CODIGO") or colunas_norm.get("CÓDIGO")
    col_descricao = colunas_norm.get("DESCRICAO") or colunas_norm.get("DESCRIÇÃO") or colunas_norm.get("DESCRICAO DO ITEM") or colunas_norm.get("DESCRIÇÃO DO ITEM")
    col_fabrica = (
        colunas_norm.get("CÓDIGO FÁBRICA") or colunas_norm.get("CODIGO FABRICA") or
        colunas_norm.get("CÓD. FÁBRICA") or colunas_norm.get("COD. FABRICA") or
        colunas_norm.get("CÓDIGO DE FÁBRICA") or colunas_norm.get("CODIGO DE FABRICA")
    )
    col_qtd = colunas_norm.get("PEDIDO FINAL") or colunas_norm.get("QUANTIDADE") or colunas_norm.get("QTD") or colunas_norm.get("QTDE")

    faltantes = []
    if not col_codigo:
        faltantes.append("codigo/zx")
    if not col_descricao:
        faltantes.append("descricao")
    if not col_fabrica:
        faltantes.append("Código Fábrica")
    if not col_qtd:
        faltantes.append("PEDIDO Final/Quantidade")
    if faltantes:
        raise ValueError("A planilha enviada não possui as colunas obrigatórias para fornecedor: " + ", ".join(faltantes))

    wb = Workbook()
    ws = wb.active
    ws.title = "Fornecedor"
    ws.append(["Código", "Descrição", "Código de Fábrica", "Quantidade"])

    for _, row in df.iterrows():
        qtd = br_to_float(row.get(col_qtd, 0))
        try:
            qtd = int(round(float(qtd)))
        except Exception:
            qtd = 0
        if qtd <= 0:
            continue

        codigo_raw = str(row.get(col_codigo, "")).strip()
        codigo_match = re.search(r"(\d+)", codigo_raw)
        codigo = codigo_match.group(1).zfill(5) if codigo_match else codigo_raw

        ws.append([
            codigo,
            str(row.get(col_descricao, "")).strip(),
            str(row.get(col_fabrica, "")).strip(),
            qtd,
        ])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    larguras = {"A": 14, "B": 48, "C": 22, "D": 14}
    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)
    for cell in ws["D"]:
        cell.number_format = '0'

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()



def _coluna_por_candidatos(df, candidatos):
    colunas_norm = {normalizar_coluna(c): c for c in df.columns}
    for candidato in candidatos:
        col = colunas_norm.get(normalizar_coluna(candidato))
        if col:
            return col
    return None


def _normalizar_nome_coluna_flex(nome):
    """Normaliza nome de coluna para reconhecimento flexível."""
    txt = _texto_sem_acentos(nome).upper().strip()
    txt = re.sub(r"[^A-Z0-9]+", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def _coluna_quantidade_flexivel(df):
    """
    Encontra coluna de quantidade mesmo quando o fornecedor usa nomes variados.
    Aceita: QTD, QTDE, QUANT, QUANT., QUANTIDADE, QTD PEDIDA, QTDE SOLICITADA,
    QTD FATURADA, QTD COMPRA, QTE, QTY, QUANTITY, VOL, VOLUME etc.
    Evita confundir com preço, valor total, saldo, código, embalagem ou peso.
    """
    if df is None or df.empty:
        return None

    bloqueios = [
        "PRECO", "PREÇO", "VALOR", "VL", "VLR", "UNIT", "UNITARIO", "UNITÁRIO",
        "TOTAL", "IPI", "SUB", "ST", "COD", "CÓD", "CODIGO", "CÓDIGO",
        "FABRICA", "FÁBRICA", "REFERENCIA", "REFERÊNCIA", "SKU", "DESCR",
        "PESO", "PES", "KIL", "LIT", "LITRO", "EMB", "EMBALAGEM",
        "ESTOQUE", "SALDO", "BAIXADO", "ABERTO", "DATA", "DT",
    ]

    fortes_exatos = {
        "QTD", "QTDE", "QTE", "QTY", "QUANT", "QUANTIDADE", "QUANTITY",
        "QUANTID", "QUANTIDADE PEDIDA", "QTD PEDIDA", "QTDE PEDIDA",
        "QUANTIDADE SOLICITADA", "QTD SOLICITADA", "QTDE SOLICITADA",
        "QUANTIDADE COMPRA", "QTD COMPRA", "QTDE COMPRA",
        "QUANTIDADE DO PEDIDO", "QTD DO PEDIDO", "QTDE DO PEDIDO",
        "QUANTIDADE PEDIDO", "QTD PEDIDO", "QTDE PEDIDO",
        "PEDIDO FINAL", "QUANTIDADE FATURADA", "QTD FATURADA", "QTDE FATURADA",
        "VOLUME", "VOL",
    }

    # 1) Prioriza nomes exatos normalizados.
    for col in df.columns:
        norm = _normalizar_nome_coluna_flex(col)
        if norm in fortes_exatos:
            return col

    # 2) Depois aceita qualquer coluna que contenha uma palavra clara de quantidade,
    # desde que não contenha termos típicos de preço, código, estoque etc.
    padrao_qtd = re.compile(r"(^| )(QTD|QTDE|QTE|QTY|QUANT|QUANTIDADE|QUANTITY|VOLUME|VOL)( |$)")
    for col in df.columns:
        norm = _normalizar_nome_coluna_flex(col)
        if not norm:
            continue
        if padrao_qtd.search(norm) and not any(b in norm for b in bloqueios):
            return col

    # 3) Caso o nome seja algo como 'QTD_PED', 'QTDE-SOLIC', 'QUANT.' já estará
    # normalizado com espaço. Este fallback pega prefixos comuns.
    for col in df.columns:
        norm = _normalizar_nome_coluna_flex(col)
        compact = re.sub(r"[^A-Z0-9]+", "", norm)
        if any(compact.startswith(p) for p in ["QTD", "QTDE", "QTE", "QTY", "QUANT", "QUANTIDADE"]):
            if not any(b in norm for b in bloqueios):
                return col

    return None




def _texto_sem_acentos(txt):
    txt = str(txt or "")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    return txt


def normalizar_codigo_fabrica(valor):
    """
    Normaliza código de fábrica para comparação entre Excel e PDF.
    Exemplos:
    - "I :401" -> "I401"
    - "P-512" -> "P512"
    - " 000123 " -> "000123"
    """
    txt = _texto_sem_acentos(valor).upper().strip()
    txt = re.sub(r"[^A-Z0-9]+", "", txt)
    if txt in ["", "NAN", "NONE", "NULL", "SEM", "SNCODIGO"]:
        return ""
    return txt


def normalizar_descricao_chave(valor):
    txt = _texto_sem_acentos(valor).upper().strip()
    txt = re.sub(r"[^A-Z0-9 ]+", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def _tokens_numericos_linha(txt):
    """Extrai números em padrão BR/US preservando a ordem visual da linha."""
    txt = str(txt or "")
    padrao = r"(?<![A-Z0-9])-?(?:R\$\s*)?\d{1,3}(?:\.\d{3})*(?:,\d+)?(?![A-Z0-9])|(?<![A-Z0-9])-?\d+(?:[\.,]\d+)?(?![A-Z0-9])"
    valores = []
    for m in re.finditer(padrao, txt, flags=re.IGNORECASE):
        bruto = m.group(0).replace("R$", "").strip()
        valores.append((m.start(), bruto, numero_planilha_para_float(bruto)))
    return valores


def _inferir_qtd_preco_total_por_linha(linha, pos_codigo=0):
    """
    Heurística para PDFs de fornecedores com layouts variados.
    Após encontrar o código de fábrica na linha, procura quantidade, preço unitário e total.
    Se houver total, tenta validar pares em que qtd x preço ~= total.
    """
    linha = str(linha or "")
    numeros = [(pos, bruto, val) for pos, bruto, val in _tokens_numericos_linha(linha) if pos >= max(0, pos_codigo)]
    numeros = [(pos, bruto, val) for pos, bruto, val in numeros if pd.notna(val) and float(val) != 0]

    if not numeros:
        return 0.0, 0.0, 0.0

    # Remove números que são parte clara de códigos longos quando aparecem colados ao produto.
    # Mantém quantidades e valores separados por espaço/coluna.
    candidatos = [float(v) for _, _, v in numeros]

    melhor = None
    for i, qtd in enumerate(candidatos):
        if qtd <= 0 or qtd > 100000:
            continue
        for j in range(i + 1, len(candidatos)):
            preco = candidatos[j]
            if preco <= 0:
                continue
            for k in range(j + 1, len(candidatos)):
                total = candidatos[k]
                if total <= 0:
                    continue
                esperado = qtd * preco
                tolerancia = max(0.05, abs(total) * 0.03)
                if abs(esperado - total) <= tolerancia:
                    melhor = (qtd, preco, total)
                    break
            if melhor:
                break
        if melhor:
            break

    if melhor:
        return melhor

    # Fallback: normalmente a primeira medida depois do código é quantidade e a seguinte é preço unitário.
    qtd = candidatos[0] if len(candidatos) >= 1 else 0.0
    preco = candidatos[1] if len(candidatos) >= 2 else 0.0
    total = candidatos[2] if len(candidatos) >= 3 else (qtd * preco if qtd and preco else 0.0)
    return qtd, preco, total


def _extrair_descricao_ao_redor_codigo(linha, codigo):
    linha = str(linha or "").strip()
    if not linha:
        return ""
    cod = str(codigo or "").strip()
    pos = linha.upper().find(cod.upper()) if cod else -1
    if pos >= 0:
        antes = linha[:pos].strip(" -|;:\t")
        depois = linha[pos + len(cod):].strip(" -|;:\t")
        depois = re.sub(r"\s+(?:R\$\s*)?\d[\d\.,]*.*$", "", depois).strip()
        desc = depois or antes
    else:
        desc = re.sub(r"\s+(?:R\$\s*)?\d[\d\.,]*.*$", "", linha).strip()
    return desc[:180]


def extrair_itens_pdf_por_codigos(uploaded_file, codigos_referencia=None):
    """
    Lê PDF de fornecedor em vários modelos.
    Estratégia principal: usa os códigos de fábrica do pedido da Única como âncora.
    Assim, mesmo sem cabeçalho ou com layout diferente, o sistema busca o código no texto/tabela
    e tenta capturar quantidade e preço unitário na mesma linha.
    """
    referencias = {}
    for c in (codigos_referencia or []):
        norm = normalizar_codigo_fabrica(c)
        if norm and len(norm) >= 3:
            referencias[norm] = str(c).strip()

    linhas_texto = []

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            # Linhas extraídas de tabelas, quando o PDF permitir.
            for tabela in (page.extract_tables() or []):
                for linha in tabela:
                    celulas = [str(c or "").strip() for c in (linha or [])]
                    if any(celulas):
                        linhas_texto.append(" | ".join(celulas))

            # Linhas extraídas como texto livre, cobrindo PDFs sem tabela estruturada.
            page_text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for linha in page_text.splitlines():
                linha = linha.strip()
                if linha:
                    linhas_texto.append(linha)

    registros = []
    vistos = set()

    for linha in linhas_texto:
        linha_norm = normalizar_codigo_fabrica(linha)
        encontrados = []

        if referencias:
            for cod_norm, cod_original in referencias.items():
                if cod_norm and cod_norm in linha_norm:
                    encontrados.append((cod_norm, cod_original))
        else:
            # Sem referência, tenta pegar candidatos parecidos com código de fábrica.
            for token in re.findall(r"\b[A-Z]{0,4}\d[A-Z0-9\-\./:]{2,}\b", _texto_sem_acentos(linha).upper()):
                cod_norm = normalizar_codigo_fabrica(token)
                if len(cod_norm) >= 3:
                    encontrados.append((cod_norm, token))

        for cod_norm, cod_original in encontrados:
            pos_raw = _texto_sem_acentos(linha).upper().find(_texto_sem_acentos(cod_original).upper())
            if pos_raw < 0:
                pos_raw = 0
            qtd, preco, total = _inferir_qtd_preco_total_por_linha(linha, pos_codigo=pos_raw + len(str(cod_original)))
            desc = _extrair_descricao_ao_redor_codigo(linha, cod_original)
            chave_visto = (cod_norm, round(qtd, 4), round(preco, 4), round(total, 4), normalizar_descricao_chave(desc))
            if chave_visto in vistos:
                continue
            vistos.add(chave_visto)
            registros.append({
                "Código Fábrica": cod_original,
                "Descrição": desc,
                "Quantidade": qtd,
                "Valor Unitário": preco,
                "Valor Total": total,
                "Linha PDF": linha,
            })

    return pd.DataFrame(registros)



def _particionar_blocos_itens_pdf(texto):
    """
    Divide texto de PDF em blocos de itens quando existe uma linha iniciando com:
    Item Abrev. Unid. Produto...
    Ex.: 1 86007 261 05.66.H003
    """
    linhas = [str(l or "").strip() for l in str(texto or "").splitlines() if str(l or "").strip()]
    blocos = []
    atual = []
    padrao_inicio = re.compile(r"^\d{1,4}\s+\d{3,6}\s+\d{3,6}\s+[A-Z0-9][A-Z0-9\.\-/]*", re.IGNORECASE)

    for linha in linhas:
        if padrao_inicio.match(linha):
            if atual:
                blocos.append(atual)
            atual = [linha]
        elif atual:
            # evita puxar rodapé/cabeçalho grande depois do fim do item
            if re.match(r"^(Página|Pedido:|Cliente:|Item\s+Abrev|Total do Pedido)", linha, flags=re.IGNORECASE):
                if re.match(r"^Total do Pedido", linha, flags=re.IGNORECASE):
                    blocos.append(atual)
                    atual = []
                continue
            atual.append(linha)
    if atual:
        blocos.append(atual)
    return blocos


def _codigo_fabrica_do_bloco_pdf(bloco):
    if not bloco:
        return "", "", ""
    primeira = str(bloco[0]).strip()
    m = re.match(r"^(\d{1,4})\s+(\d{3,6})\s+(\d{3,6})\s+([A-Z0-9][A-Z0-9\.\-/]*)", primeira, flags=re.IGNORECASE)
    if not m:
        return "", "", ""

    item, abrev, unid, produto = m.group(1), m.group(2), m.group(3), m.group(4)
    produto_completo = produto

    # Alguns PDFs quebram o final do código do produto na(s) linha(s) seguinte(s).
    # Ex.: 05.66.H003\n5 => 05.66.H0035; 05.00.M350\n0 => 05.00.M3500.
    for linha in bloco[1:4]:
        token = str(linha).strip()
        if re.fullmatch(r"[A-Z0-9]{1,3}", token, flags=re.IGNORECASE):
            produto_completo += token
            break
        # quando já começou descrição, para
        if re.search(r"[A-ZÁÉÍÓÚÃÕÇ]{3,}", token, flags=re.IGNORECASE) and not re.fullmatch(r"[A-Z0-9\.\-/]+", token, flags=re.IGNORECASE):
            break

    codigo_original = f"{produto_completo}-{unid}"
    return codigo_original, normalizar_codigo_fabrica(codigo_original), unid


def _descricao_do_bloco_pdf(bloco):
    if not bloco:
        return ""
    partes = []
    for linha in bloco[1:]:
        t = str(linha).strip()
        if not t:
            continue
        if re.fullmatch(r"[A-Z0-9]{1,3}", t, flags=re.IGNORECASE):
            continue
        if re.fullmatch(r"\d+(?:[\.,]\d+)?", t):
            continue
        # remove a cauda numérica de quantidade/preço/total da última linha
        t = re.sub(r"\s+\d+(?:[\.,]\d+)?\s+\d+(?:[\.,]\d+)?\s+\d{1,3}(?:\.\d{3})*,\d+\s*$", "", t).strip()
        if re.search(r"[A-ZÁÉÍÓÚÃÕÇ]{3,}", t, flags=re.IGNORECASE):
            partes.append(t)
    return " ".join(partes).strip()[:180]


def _qtd_preco_total_do_bloco_pdf(bloco):
    texto = " ".join(str(x or "").strip() for x in bloco)
    nums = _tokens_numericos_linha(texto)
    valores = [float(v) for _, _, v in nums if pd.notna(v)]
    if len(valores) < 3:
        return 0.0, 0.0, 0.0

    # Nos pedidos de fornecedor, normalmente os 3 últimos números do bloco são:
    # quantidade, preço unitário e preço total. Valida qtd x preço ~= total.
    for i in range(len(valores) - 3, -1, -1):
        qtd, preco, total = valores[i], valores[i + 1], valores[i + 2]
        if qtd > 0 and preco > 0 and total > 0:
            if abs(qtd * preco - total) <= max(0.10, abs(total) * 0.03):
                return qtd, preco, total
    qtd, preco, total = valores[-3], valores[-2], valores[-1]
    return qtd, preco, total


def extrair_itens_pdf_por_blocos(uploaded_file, codigos_referencia=None):
    """
    Parser complementar para PDFs em que as colunas ficam visualmente desalinhadas.
    Exemplo real Sherwin/Lazzuril:
    - Excel: 05.66.H0035-261
    - PDF: linha do produto 05.66.H003 / linha do item com UNID 261 / linha seguinte 5 0
    O parser monta Código de Fábrica como Produto-Unid e valida contra os códigos do Excel.
    """
    referencias = set()
    for c in (codigos_referencia or []):
        norm = normalizar_codigo_fabrica(c)
        if norm:
            referencias.add(norm)

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    textos_paginas = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                textos_paginas.append(page.extract_text(x_tolerance=1, y_tolerance=3) or "")
    except Exception:
        return pd.DataFrame()

    padrao_item = re.compile(r"^(\d{1,4})\s+(\d{3,6})\s+(\d{3,6})\s+(.+)$", flags=re.IGNORECASE)
    padrao_cod_produto = re.compile(r"\b[A-Z0-9]{1,4}(?:\.[A-Z0-9]{1,8}){1,4}[A-Z0-9]*\b", flags=re.IGNORECASE)

    registros = []
    vistos = set()

    for texto in textos_paginas:
        linhas = [str(l or "").strip() for l in str(texto or "").splitlines() if str(l or "").strip()]
        for i, linha in enumerate(linhas):
            m = padrao_item.match(linha)
            if not m:
                continue

            item, abrev, unid, resto = m.group(1), m.group(2), m.group(3), m.group(4)
            nums_linha = _tokens_numericos_linha(linha)
            # item, abrev e unid também entram como número; os três últimos são qtd/preço/total.
            if len(nums_linha) < 6:
                continue
            qtd = float(nums_linha[-3][2])
            preco = float(nums_linha[-2][2])
            total = float(nums_linha[-1][2])
            if qtd <= 0 or preco <= 0:
                continue

            pos_qtd = nums_linha[-3][0]
            prefixo = linha[:pos_qtd].strip()
            prefixo_sem_inicio = re.sub(r"^\d{1,4}\s+\d{3,6}\s+\d{3,6}\s+", "", prefixo).strip()

            prev1 = linhas[i - 1] if i - 1 >= 0 else ""
            prev2 = linhas[i - 2] if i - 2 >= 0 else ""
            next1 = linhas[i + 1] if i + 1 < len(linhas) else ""

            cod_base = ""
            desc_partes = []

            cods_no_prefixo = padrao_cod_produto.findall(prefixo_sem_inicio)
            if cods_no_prefixo:
                cod_base = cods_no_prefixo[-1]
                desc_prefixo = prefixo_sem_inicio.replace(cod_base, " ").strip()
                if desc_prefixo:
                    desc_partes.append(desc_prefixo)
            else:
                # Produto pode estar na linha anterior, junto com litragem e parte da descrição.
                for prev in [prev1, prev2]:
                    cods_prev = padrao_cod_produto.findall(prev)
                    if cods_prev:
                        cod_base = cods_prev[0]
                        desc_prev = prev.replace(cod_base, " ").strip()
                        desc_prev = re.sub(r"^\d+(?:[\.,]\d+)?\s*", "", desc_prev).strip()
                        if desc_prev:
                            desc_partes.append(desc_prev)
                        break
                if prefixo_sem_inicio:
                    desc_partes.append(prefixo_sem_inicio)

            if not cod_base:
                continue

            # Gera variações com complemento quebrado na linha seguinte.
            candidatos_produto = [cod_base]
            next_tokens = str(next1 or "").split()
            if next_tokens:
                primeiro_next = re.sub(r"[^A-Z0-9]", "", _texto_sem_acentos(next_tokens[0]).upper())
                if primeiro_next and len(primeiro_next) <= 3:
                    candidatos_produto.append(cod_base + primeiro_next)

            escolhido_original = ""
            escolhido_norm = ""
            for prod in candidatos_produto:
                original = f"{prod}-{unid}"
                norm = normalizar_codigo_fabrica(original)
                if referencias and norm in referencias:
                    escolhido_original = original
                    escolhido_norm = norm
                    break
            if not escolhido_original:
                original = f"{candidatos_produto[-1]}-{unid}"
                norm = normalizar_codigo_fabrica(original)
                escolhido_original = original
                escolhido_norm = norm

            # Complementa descrição com a linha seguinte quando ela é descrição, removendo quebra de código/litragem.
            if next1:
                prox = str(next1).strip()
                prox_limpo = re.sub(r"^\d{1,3}\s+", "", prox).strip()
                prox_limpo = re.sub(r"^\d+(?:[\.,]\d+)?\s+", "", prox_limpo).strip()
                if re.search(r"[A-ZÁÉÍÓÚÃÕÇ]{3,}", prox_limpo, flags=re.IGNORECASE):
                    desc_partes.append(prox_limpo)

            descricao = " ".join([d for d in desc_partes if d]).strip()
            descricao = re.sub(r"\s+", " ", descricao)[:180]

            chave = (escolhido_norm, round(qtd, 4), round(preco, 4), round(total, 4))
            if chave in vistos:
                continue
            vistos.add(chave)
            registros.append({
                "Código Fábrica": escolhido_original,
                "Descrição": descricao,
                "Quantidade": qtd,
                "Valor Unitário": preco,
                "Valor Total": total,
                "Linha PDF": linha,
            })

    return pd.DataFrame(registros)

def _deduplicar_headers_planilha(headers):
    """Garante nomes únicos de colunas, preservando o texto original quando existir."""
    saida = []
    contagem = {}
    for i, h in enumerate(headers):
        nome = str(h or "").strip()
        if not nome or nome.lower() in ["nan", "none"]:
            nome = f"COLUNA {i + 1}"
        base = nome
        chave = normalizar_coluna(base)
        contagem[chave] = contagem.get(chave, 0) + 1
        if contagem[chave] > 1:
            nome = f"{base}.{contagem[chave]}"
        saida.append(nome)
    return saida


def _pontuar_linha_cabecalho_planilha(valores):
    """
    Pontua uma possível linha de cabeçalho em planilhas de fornecedor.
    Resolve arquivos que vêm com dados do cliente antes da tabela, como:
    CÓDIGO | DESCRIÇÃO | ... | QTDE | LITROS | VL. UNIT. | ... | VL. TOTAL
    """
    textos = [_normalizar_nome_coluna_flex(v) for v in valores]
    compactos = [re.sub(r"[^A-Z0-9]+", "", t) for t in textos]
    linha = " ".join(t for t in textos if t)

    score = 0
    tem_codigo = any(t in ["CODIGO", "COD", "SKU", "REFERENCIA", "REF"] or t.startswith("CODIGO ") or t.startswith("COD ") for t in textos) or any(c in ["CODIGO", "COD", "SKU", "REFERENCIA", "REF"] for c in compactos)
    tem_desc = any("DESCR" in t or t in ["PRODUTO", "ITEM", "NOME"] for t in textos)
    tem_qtd = any(re.search(r"(^| )(QTD|QTDE|QTE|QTY|QUANT|QUANTIDADE)( |$)", t) for t in textos) or any(c in ["QTD", "QTDE", "QTE", "QTY", "QUANT", "QUANTIDADE"] for c in compactos)
    tem_preco = any("PRECO" in t or "UNIT" in t or "VL UNIT" in t or "VR UNIT" in t or "VALOR UNIT" in t for t in textos)
    tem_total = any(("TOTAL" in t and "UNIT" not in t) or "VL TOTAL" in t or "VALOR TOTAL" in t for t in textos)

    score += 3 if tem_codigo else 0
    score += 3 if tem_desc else 0
    score += 4 if tem_qtd else 0
    score += 2 if tem_preco else 0
    score += 2 if tem_total else 0

    # Penaliza linhas de formulário/cadastro do cliente, não tabela de produtos.
    if any(p in linha for p in ["DADOS DO CLIENTE", "CNPJ", "CONDICAO PAGTO", "OBSERVACAO NF", "BAIRRO", "SUB REGIAO"]):
        score -= 4
    return score


def _ajustar_cabecalho_planilha_fornecedor(df_raw):
    """
    Quando o Excel do fornecedor possui cabeçalho fora da primeira linha,
    encontra a linha da tabela e transforma em DataFrame tabular.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    df_scan = df_raw.copy()
    melhor_idx = None
    melhor_score = -999

    limite = min(len(df_scan), 100)
    for idx in range(limite):
        valores = list(df_scan.iloc[idx].values)
        score = _pontuar_linha_cabecalho_planilha(valores)
        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx

    # Exige pelo menos código/descrição/quantidade ou pontuação equivalente.
    if melhor_idx is None or melhor_score < 7:
        return df_raw.copy()

    headers = _deduplicar_headers_planilha(list(df_scan.iloc[melhor_idx].values))
    df = df_scan.iloc[melhor_idx + 1:].copy()
    df.columns = headers

    # Remove linhas totalmente vazias e linhas de totalização.
    df = df.dropna(how="all")
    primeira_col = df.columns[0] if len(df.columns) else None
    if primeira_col:
        primeira_txt = df[primeira_col].astype(str).str.strip().str.upper()
        df = df[~primeira_txt.isin(["", "NAN", "NONE"])]
        df = df[~primeira_txt.str.contains(r"^TOTA(IS|L)?$|^TOTAL", regex=True, na=False)]

    # Remove colunas completamente vazias, mas mantém nomes reconhecidos.
    df = df.dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def ler_planilha_comparativo_fornecedor(uploaded_file):
    """
    Lê Excel/CSV do fornecedor procurando automaticamente a linha real do cabeçalho.
    Importante para modelos em que a tabela começa depois de blocos como DADOS DO CLIENTE.
    """
    if uploaded_file is None:
        return pd.DataFrame()

    nome = str(getattr(uploaded_file, "name", "")).lower()
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    if nome.endswith((".xlsx", ".xls")):
        # header=None permite localizar cabeçalhos que não estão na primeira linha.
        # Para .xls, mantenha xlrd>=2.0.1 no requirements.txt.
        df_raw = pd.read_excel(uploaded_file, dtype=str, header=None)
        return _ajustar_cabecalho_planilha_fornecedor(df_raw)

    # CSV: também pode ter linhas acima do cabeçalho.
    tentativas = [
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "latin1"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "latin1"},
        {"sep": "\t", "encoding": "utf-8-sig"},
        {"sep": "\t", "encoding": "latin1"},
    ]
    ultimo_erro = None
    for tentativa in tentativas:
        try:
            uploaded_file.seek(0)
            df_raw = pd.read_csv(
                uploaded_file,
                sep=tentativa["sep"],
                encoding=tentativa["encoding"],
                dtype=str,
                engine="python",
                on_bad_lines="skip",
                header=None,
            )
            return _ajustar_cabecalho_planilha_fornecedor(df_raw)
        except Exception as e:
            ultimo_erro = str(e)
            continue

    raise RuntimeError(f"Não consegui ler a planilha do fornecedor. Último erro: {ultimo_erro}")


def ler_arquivo_comparativo(uploaded_file, codigos_referencia=None):
    if uploaded_file is None:
        return pd.DataFrame()

    nome = str(getattr(uploaded_file, "name", "")).lower()
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    if nome.endswith(".pdf"):
        df_pdf = extrair_itens_pdf_por_blocos(uploaded_file, codigos_referencia=codigos_referencia)
        if df_pdf.empty:
            df_pdf = extrair_itens_pdf_por_codigos(uploaded_file, codigos_referencia=codigos_referencia)
        if not df_pdf.empty:
            return df_pdf

        # Fallback antigo: tenta tabelas com cabeçalho quando não encontrou códigos no texto.
        linhas = []
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables() or []
                for tabela in tabelas:
                    for linha in tabela:
                        if linha and any(str(c or "").strip() for c in linha):
                            linhas.append([str(c or "").strip() for c in linha])
        if not linhas:
            return pd.DataFrame()

        max_cols = max(len(linha) for linha in linhas)
        linhas = [linha + [""] * (max_cols - len(linha)) for linha in linhas]
        header_idx = 0
        for i, linha in enumerate(linhas[:15]):
            linha_norm = " ".join(normalizar_coluna(c) for c in linha)
            if any(p in linha_norm for p in ["COD", "CÓD", "DESCR", "QTD", "QTDE", "QUANT", "UNIT", "PRECO", "PREÇO"]):
                header_idx = i
                break
        headers = [str(c).strip() or f"COLUNA {i+1}" for i, c in enumerate(linhas[header_idx])]
        return pd.DataFrame(linhas[header_idx + 1:], columns=headers)

    if nome.endswith((".html", ".htm")):
        try:
            uploaded_file.seek(0)
            tabelas = pd.read_html(uploaded_file, dtype=str)
            if tabelas:
                # Usa a maior tabela encontrada no HTML.
                tabelas = sorted(tabelas, key=lambda d: d.shape[0] * max(d.shape[1], 1), reverse=True)
                df_html = tabelas[0].copy()
                df_html.columns = [str(c).strip() for c in df_html.columns]
                return df_html
        except Exception:
            return pd.DataFrame()

    return ler_planilha_comparativo_fornecedor(uploaded_file)


def _opcoes_colunas_mapeamento(df, incluir_vazio=True):
    cols = [str(c) for c in list(df.columns)] if df is not None else []
    return (["-- Não usar --"] + cols) if incluir_vazio else cols


def _primeira_coluna_existente(df, candidatos, permitir_vazio=True):
    if df is None or df.empty:
        return "-- Não usar --" if permitir_vazio else None
    col = _coluna_por_candidatos(df, candidatos)
    if col:
        return str(col)
    return "-- Não usar --" if permitir_vazio else (str(df.columns[0]) if len(df.columns) else None)


def _mapear_colunas_comparativo(df, prefixo, origem):
    st.markdown(f"**Mapeamento do {origem}**")
    op_obrig = _opcoes_colunas_mapeamento(df, incluir_vazio=False)
    op_opc = _opcoes_colunas_mapeamento(df, incluir_vazio=True)

    if not op_obrig:
        st.error(f"Não encontrei colunas no arquivo {origem}.")
        return None

    default_codigo = _primeira_coluna_existente(df, [
        "Código Fábrica", "Codigo Fabrica", "Cód. Fábrica", "Cod. Fabrica",
        "Código de Fábrica", "Codigo de Fabrica", "Código", "Codigo", "Cód.", "Cod.", "Cod",
        "Referência", "Referencia", "Ref", "SKU", "Produto", "Cod Produto", "Código Produto"
    ], permitir_vazio=True)
    default_desc = _primeira_coluna_existente(df, [
        "descricao", "descrição", "descrição do item", "descricao do item", "produto", "item", "nome", "descr", "Linha PDF"
    ], permitir_vazio=True)
    default_qtd = _coluna_quantidade_flexivel(df) or _primeira_coluna_existente(df, [
        "PEDIDO Final", "Quantidade", "Qtd", "Qtde", "Quant", "Qte", "Qty", "QTDE"
    ], permitir_vazio=True)
    default_preco = _primeira_coluna_existente(df, [
        "Preço Última Compra", "Preco Ultima Compra", "Preço", "Preco", "Preço Unitário", "Preco Unitario",
        "Valor Unitário", "Valor Unitario", "Vlr Unit", "Vl Unit", "VL. UNIT.", "VL UNIT", "VR.UNIT", "UNIT.TOT"
    ], permitir_vazio=True)
    default_total = _primeira_coluna_existente(df, [
        "Valor Final do Pedido", "Valor Total", "VL. TOTAL", "VL TOTAL", "Vlr Total", "Total", "Valor"
    ], permitir_vazio=True)

    def idx(opcoes, valor):
        return opcoes.index(valor) if valor in opcoes else 0

    c1, c2 = st.columns(2)
    with c1:
        col_codigo = st.selectbox("Coluna de Código de Fábrica / relacionamento", op_opc, index=idx(op_opc, default_codigo), key=f"{prefixo}_codigo")
        col_desc = st.selectbox("Coluna de descrição", op_opc, index=idx(op_opc, default_desc), key=f"{prefixo}_desc")
    with c2:
        col_qtd = st.selectbox("Coluna de quantidade", op_opc, index=idx(op_opc, default_qtd), key=f"{prefixo}_qtd")
        col_preco = st.selectbox("Coluna de valor unitário", op_opc, index=idx(op_opc, default_preco), key=f"{prefixo}_preco")
        col_total = st.selectbox("Coluna de valor total (opcional)", op_opc, index=idx(op_opc, default_total), key=f"{prefixo}_total")

    def limpar(v):
        return None if v == "-- Não usar --" else v

    return {
        "codigo": limpar(col_codigo),
        "descricao": limpar(col_desc),
        "quantidade": limpar(col_qtd),
        "preco_unitario": limpar(col_preco),
        "valor_total": limpar(col_total),
    }


def normalizar_pedido_comparativo(df, origem, mapa_colunas=None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    mapa_colunas = mapa_colunas or {}

    # Quando houver mapeamento manual, ele manda. O reconhecimento automático fica só como fallback.
    col_fabrica = mapa_colunas.get("codigo") or _coluna_por_candidatos(df, [
        "Código Fábrica", "Codigo Fabrica", "Cód. Fábrica", "Cod. Fabrica",
        "Código de Fábrica", "Codigo de Fabrica", "Cod Fabrica", "Cód Fabrica",
        "Código", "Codigo", "Cód.", "Cod.", "Cod",
        "Referência", "Referencia", "Ref", "Código Fornecedor", "Codigo Fornecedor",
        "Cod Produto", "Código Produto", "Codigo Produto", "SKU", "Part Number", "Linha PDF",
    ])
    col_descricao = mapa_colunas.get("descricao") or _coluna_por_candidatos(df, [
        "descricao", "descrição", "descrição do item", "descricao do item", "produto", "item", "nome", "descr", "linha pdf",
    ])
    col_qtd = mapa_colunas.get("quantidade") or _coluna_quantidade_flexivel(df) or _coluna_por_candidatos(df, [
        "PEDIDO Final", "Quantidade", "Qtd", "Qtde", "Quant", "Quant.", "Qte", "Qty", "Quantity",
        "Quantidade Pedido", "Qtd Pedido", "Qtde Pedido", "Quant Pedido",
        "Quantidade Pedida", "Qtd Pedida", "Qtde Pedida", "Quant Pedida",
        "Quantidade Solicitada", "Qtd Solicitada", "Qtde Solicitada", "Quant Solicitada",
        "Quantidade Compra", "Qtd Compra", "Qtde Compra", "Qtd. Compra",
        "Quantidade Faturada", "Qtd Faturada", "Qtde Faturada",
        "Volume", "Vol",
    ])
    col_preco = mapa_colunas.get("preco_unitario") or _coluna_por_candidatos(df, [
        "Preço Última Compra", "Preco Ultima Compra", "Preço", "Preco", "Preço Unitário", "Preco Unitario",
        "Valor Unitário", "Valor Unitario", "Vlr Unit", "Vl Unit", "VL. UNIT.", "VL UNIT",
        "Vr.Unit", "VR.UNIT", "VR UNIT", "Unitário", "Unitario", "Preço Uni",
        "UNIT.TOT", "UNIT TOT", "Unit Total",
    ])
    col_total = mapa_colunas.get("valor_total") or _coluna_por_candidatos(df, [
        "Valor Final do Pedido", "Valor Total", "VL. TOTAL", "VL TOTAL", "Vlr Total",
        "Total", "Total Geral", "Valor", "Valor Mercadoria",
    ])

    # Quando o arquivo veio de PDF sem cabeçalho perfeito, tenta inferir por linha completa.
    if (not col_qtd or not col_preco) and "Linha PDF" in df.columns:
        linhas_inferidas = df["Linha PDF"].astype(str).apply(_inferir_qtd_preco_total_por_linha)
        df["__qtd_inferida"] = [x[0] for x in linhas_inferidas]
        df["__preco_inferido"] = [x[1] for x in linhas_inferidas]
        df["__total_inferido"] = [x[2] for x in linhas_inferidas]
        col_qtd = col_qtd or "__qtd_inferida"
        col_preco = col_preco or "__preco_inferido"
        col_total = col_total or "__total_inferido"

    faltantes = []
    if not col_fabrica:
        faltantes.append("Código de Fábrica / relacionamento")
    if not col_qtd:
        faltantes.append("quantidade")
    if faltantes:
        raise ValueError(f"O arquivo {origem} precisa de mapeamento para: " + ", ".join(faltantes) + ".")

    out = pd.DataFrame()
    out["origem"] = origem
    out["codigo_fabrica"] = df[col_fabrica].astype(str).str.strip() if col_fabrica else ""
    out["codigo_fabrica_norm"] = out["codigo_fabrica"].apply(normalizar_codigo_fabrica)
    out["descricao"] = df[col_descricao].astype(str).str.strip() if col_descricao else out["codigo_fabrica"]
    out["quantidade"] = df[col_qtd].apply(numero_planilha_para_float)
    out["preco_unitario"] = df[col_preco].apply(numero_planilha_para_float) if col_preco else 0.0
    out["valor_total"] = df[col_total].apply(numero_planilha_para_float) if col_total else 0.0
    out.loc[(out["valor_total"] <= 0) & (out["preco_unitario"] > 0), "valor_total"] = out["quantidade"] * out["preco_unitario"]
    out.loc[(out["preco_unitario"] <= 0) & (out["valor_total"] > 0) & (out["quantidade"] > 0), "preco_unitario"] = out["valor_total"] / out["quantidade"]
    out["descricao_chave"] = out["descricao"].apply(normalizar_descricao_chave)
    out["__linha_origem"] = range(len(out))
    out = out[(out["quantidade"] > 0) | (out["preco_unitario"] > 0) | (out["valor_total"] > 0)].copy()
    return out

def agregar_pedido_comparativo(df):
    if df.empty:
        return df.copy()
    df = df.copy().reset_index(drop=True)
    df["codigo_fabrica"] = df["codigo_fabrica"].fillna("").astype(str).str.strip()
    df["codigo_fabrica_norm"] = df["codigo_fabrica_norm"].fillna("").astype(str).str.strip()

    # Regra principal do comparativo: relacionamento automático SOMENTE por Código de Fábrica.
    # Itens sem código não são aproximados por descrição; ficam separados e podem ser vinculados manualmente.
    df["chave"] = df.apply(
        lambda r: str(r["codigo_fabrica_norm"]) if str(r["codigo_fabrica_norm"]).strip() else f"SEM_CODIGO_{r.name}",
        axis=1,
    )
    agg = df.groupby("chave", as_index=False).agg(
        codigo_fabrica=("codigo_fabrica", "first"),
        codigo_fabrica_norm=("codigo_fabrica_norm", "first"),
        descricao=("descricao", "first"),
        descricao_chave=("descricao_chave", "first"),
        quantidade=("quantidade", "sum"),
        valor_total=("valor_total", "sum"),
    )
    agg["preco_unitario"] = agg.apply(
        lambda r: float(r["valor_total"]) / float(r["quantidade"]) if float(r["quantidade"] or 0) > 0 and float(r["valor_total"] or 0) > 0 else 0,
        axis=1,
    )
    return agg


def codigos_referencia_comparativo(df_unica_raw, mapa_unica=None):
    try:
        unica_norm = normalizar_pedido_comparativo(df_unica_raw, "Única", mapa_unica)
        codigos = unica_norm["codigo_fabrica"].dropna().astype(str).str.strip().tolist()
        return [c for c in codigos if normalizar_codigo_fabrica(c)]
    except Exception:
        return []


def _aplicar_relacionamentos_manuais(unica, fornecedor, relacionamentos):
    relacionamentos = relacionamentos or {}
    if not relacionamentos:
        return fornecedor.copy()

    fornecedor = fornecedor.copy()
    for chave_unica, chave_fornecedor in relacionamentos.items():
        if not chave_unica or not chave_fornecedor:
            continue
        mask_f = fornecedor["chave"].astype(str) == str(chave_fornecedor)
        if mask_f.any():
            fornecedor.loc[mask_f, "chave"] = str(chave_unica)
    return fornecedor


def montar_comparativo_pedidos(df_unica_raw, df_fornecedor_raw, mapa_unica=None, mapa_fornecedor=None, relacionamentos_manuais=None):
    unica_normalizada = normalizar_pedido_comparativo(df_unica_raw, "Única", mapa_unica)

    # Regra do comparativo: itens com PEDIDO FINAL / quantidade da Única igual a zero
    # não entram na base de comparação. Isso evita apontar divergência de itens que
    # foram carregados na planilha, mas não foram efetivamente pedidos.
    if not unica_normalizada.empty and "quantidade" in unica_normalizada.columns:
        unica_normalizada["quantidade"] = pd.to_numeric(unica_normalizada["quantidade"], errors="coerce").fillna(0)
        unica_normalizada = unica_normalizada[unica_normalizada["quantidade"] > 0].copy()

    unica = agregar_pedido_comparativo(unica_normalizada)
    fornecedor = agregar_pedido_comparativo(normalizar_pedido_comparativo(df_fornecedor_raw, "Fornecedor", mapa_fornecedor))
    fornecedor = _aplicar_relacionamentos_manuais(unica, fornecedor, relacionamentos_manuais)

    usados_fornecedor = set()
    linhas = []

    for _, row in unica.iterrows():
        match = None
        metodo = ""
        cod_fab = str(row.get("codigo_fabrica", "")).strip()
        cod_fab_norm = str(row.get("codigo_fabrica_norm", "")).strip()

        # 1) Relacionamento manual salvo pelo usuário.
        candidatos_manual = fornecedor[fornecedor["chave"].astype(str) == str(row.get("chave", ""))]
        candidatos_manual = candidatos_manual[~candidatos_manual.index.isin(usados_fornecedor)]
        if not candidatos_manual.empty:
            match = candidatos_manual.iloc[0]
            usados_fornecedor.add(match.name)
            metodo = "Relacionamento manual"

        # 2) Relacionamento automático SOMENTE por Código de Fábrica normalizado.
        if match is None and cod_fab_norm:
            candidatos = fornecedor[fornecedor["codigo_fabrica_norm"].astype(str).str.strip() == cod_fab_norm]
            candidatos = candidatos[~candidatos.index.isin(usados_fornecedor)]
            if not candidatos.empty:
                match = candidatos.iloc[0]
                usados_fornecedor.add(match.name)
                metodo = "Código de Fábrica"

        # Não usa descrição aproximada/fuzzy para evitar comparação de produtos parecidos.

        if match is None:
            linhas.append({
                "Status": "Não encontrado no fornecedor",
                "Método": "Sem correspondência",
                "Chave Única": row.get("chave", ""),
                "Chave Fornecedor": "",
                "Código Única": cod_fab,
                "Descrição Única": row.get("descricao", ""),
                "Qtd Única": row.get("quantidade", 0),
                "Preço Única": row.get("preco_unitario", 0),
                "Valor Única": row.get("valor_total", 0),
                "Código Fornecedor": "",
                "Descrição Fornecedor": "",
                "Qtd Fornecedor": 0,
                "Preço Fornecedor": 0,
                "Valor Fornecedor": 0,
                "Diferença Qtd": -float(row.get("quantidade", 0) or 0),
                "Diferença Preço": -float(row.get("preco_unitario", 0) or 0),
                "Diferença Preço %": -100.0 if float(row.get("preco_unitario", 0) or 0) > 0 else 0.0,
                "Diferença Valor": -float(row.get("valor_total", 0) or 0),
            })
            continue

        qtd_unica = float(row.get("quantidade", 0) or 0)
        preco_unica = float(row.get("preco_unitario", 0) or 0)
        valor_unica = float(row.get("valor_total", 0) or 0)
        qtd_fornecedor = float(match.get("quantidade", 0) or 0)
        preco_fornecedor = float(match.get("preco_unitario", 0) or 0)
        valor_fornecedor = float(match.get("valor_total", 0) or 0)

        dif_qtd = qtd_fornecedor - qtd_unica
        dif_preco = preco_fornecedor - preco_unica
        dif_preco_pct = (dif_preco / preco_unica * 100) if abs(preco_unica) > 0.000001 else (100.0 if abs(preco_fornecedor) > 0.000001 else 0.0)
        dif_valor = valor_fornecedor - valor_unica
        status = "OK" if abs(dif_qtd) < 0.0001 and abs(dif_preco) < 0.01 else "Divergente"
        linhas.append({
            "Status": status,
            "Método": metodo,
            "Chave Única": row.get("chave", ""),
            "Chave Fornecedor": match.get("chave", ""),
            "Código Única": cod_fab,
            "Descrição Única": row.get("descricao", ""),
            "Qtd Única": row.get("quantidade", 0),
            "Preço Única": row.get("preco_unitario", 0),
            "Valor Única": row.get("valor_total", 0),
            "Código Fornecedor": match.get("codigo_fabrica", ""),
            "Descrição Fornecedor": match.get("descricao", ""),
            "Qtd Fornecedor": match.get("quantidade", 0),
            "Preço Fornecedor": match.get("preco_unitario", 0),
            "Valor Fornecedor": match.get("valor_total", 0),
            "Diferença Qtd": dif_qtd,
            "Diferença Preço": dif_preco,
            "Diferença Preço %": dif_preco_pct,
            "Diferença Valor": dif_valor,
        })

    extras = fornecedor[~fornecedor.index.isin(usados_fornecedor)].copy()
    for _, row in extras.iterrows():
        linhas.append({
            "Status": "Somente fornecedor",
            "Método": "Sem correspondência",
            "Chave Única": "",
            "Chave Fornecedor": row.get("chave", ""),
            "Código Única": "",
            "Descrição Única": "",
            "Qtd Única": 0,
            "Preço Única": 0,
            "Valor Única": 0,
            "Código Fornecedor": row.get("codigo_fabrica", ""),
            "Descrição Fornecedor": row.get("descricao", ""),
            "Qtd Fornecedor": row.get("quantidade", 0),
            "Preço Fornecedor": row.get("preco_unitario", 0),
            "Valor Fornecedor": row.get("valor_total", 0),
            "Diferença Qtd": row.get("quantidade", 0),
            "Diferença Preço": row.get("preco_unitario", 0),
            "Diferença Preço %": 100.0 if float(row.get("preco_unitario", 0) or 0) > 0 else 0.0,
            "Diferença Valor": row.get("valor_total", 0),
        })

    return pd.DataFrame(linhas)


def gerar_relatorio_executivo_comparativo(df_comparativo, limite_itens=80):
    """
    Gera um relatório executivo em Markdown para a conferência do pedido.
    O texto é dividido por tópicos para facilitar copiar/colar no WhatsApp, e-mail ou ata.
    """
    if df_comparativo is None or df_comparativo.empty:
        return """# 📋 RELATÓRIO DE CONFERÊNCIA DO PEDIDO

## ✅ Resumo Geral

- Nenhum item foi encontrado para comparação.

## ✅ Conclusão

Não há divergências a reportar porque o comparativo está vazio.
"""

    df = df_comparativo.copy()
    for col in [
        "Qtd Única", "Qtd Fornecedor", "Preço Única", "Preço Fornecedor",
        "Diferença Qtd", "Diferença Preço", "Diferença Preço %", "Diferença Valor",
        "Valor Única", "Valor Fornecedor",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    total = len(df)
    qtd_ok = int((df["Status"] == "OK").sum()) if "Status" in df.columns else 0
    qtd_div = int((df["Status"] == "Divergente").sum()) if "Status" in df.columns else 0
    qtd_nao_fornecedor = int((df["Status"] == "Não encontrado no fornecedor").sum()) if "Status" in df.columns else 0
    qtd_somente_fornecedor = int((df["Status"] == "Somente fornecedor").sum()) if "Status" in df.columns else 0

    diverg_qtd = df[
        (df["Status"].isin(["Divergente", "Somente fornecedor", "Não encontrado no fornecedor"]))
        & (df["Diferença Qtd"].abs() > 0.0001)
    ].copy()
    diverg_preco = df[
        (df["Status"].isin(["Divergente", "Somente fornecedor", "Não encontrado no fornecedor"]))
        & (df["Diferença Preço"].abs() > 0.01)
    ].copy()

    def nome_item(row):
        codigo = str(row.get("Código Única") or row.get("Código Fornecedor") or "").strip()
        desc = str(row.get("Descrição Única") or row.get("Descrição Fornecedor") or "").strip()
        if codigo and desc:
            return f"{codigo} - {desc}"
        return codigo or desc or "item sem código"

    def sinal_num(valor, casas=1):
        try:
            valor = float(valor or 0)
            sinal = "+" if valor > 0 else ""
            return sinal + str(format_num_br(valor, casas))
        except Exception:
            return str(valor)

    def sinal_pct(valor):
        try:
            valor = float(valor or 0)
            sinal = "+" if valor > 0 else ""
            return sinal + str(format_num_br(valor, 2)) + "%"
        except Exception:
            return str(valor)

    linhas = []
    linhas.append("# 📋 RELATÓRIO DE CONFERÊNCIA DO PEDIDO")
    linhas.append("")
    linhas.append("## ✅ Resumo Geral")
    linhas.append("")
    linhas.append(f"- **Itens comparados:** {format_int_br(total)}")
    linhas.append(f"- **Itens sem divergência:** {format_int_br(qtd_ok)}")
    linhas.append(f"- **Itens com divergência:** {format_int_br(qtd_div)}")
    linhas.append(f"- **Itens no Pedido Única não encontrados no fornecedor:** {format_int_br(qtd_nao_fornecedor)}")
    linhas.append(f"- **Itens enviados pelo fornecedor que não constam no Pedido Única:** {format_int_br(qtd_somente_fornecedor)}")
    linhas.append("")

    linhas.append("## 📦 Divergências de Quantidade")
    linhas.append("")
    if diverg_qtd.empty:
        linhas.append("- Não foram identificadas divergências de quantidade.")
    else:
        linhas.append("Foram identificadas divergências nas quantidades dos seguintes itens:")
        linhas.append("")
        for _, r in diverg_qtd.head(limite_itens).iterrows():
            status = str(r.get("Status", ""))
            item = nome_item(r)
            linhas.append(f"### 🔸 {item}")
            if status == "Somente fornecedor":
                linhas.append("- **Situação:** item consta somente no pedido do fornecedor.")
                linhas.append(f"- **Quantidade no fornecedor:** {format_num_br(r.get('Qtd Fornecedor', 0), 1)}")
                linhas.append("- **Quantidade no Pedido Única:** 0")
                linhas.append("- **Ação recomendada:** verificar se o fornecedor incluiu item indevido ou se houve alteração posterior no pedido.")
            elif status == "Não encontrado no fornecedor":
                linhas.append("- **Situação:** item consta no Pedido Única, mas não foi encontrado no fornecedor.")
                linhas.append(f"- **Quantidade no Pedido Única:** {format_num_br(r.get('Qtd Única', 0), 1)}")
                linhas.append("- **Quantidade no fornecedor:** 0")
                linhas.append("- **Ação recomendada:** confirmar se houve corte, ruptura ou omissão do item pelo fornecedor.")
            else:
                linhas.append(f"- **Quantidade no Pedido Única:** {format_num_br(r.get('Qtd Única', 0), 1)}")
                linhas.append(f"- **Quantidade no fornecedor:** {format_num_br(r.get('Qtd Fornecedor', 0), 1)}")
                linhas.append(f"- **Diferença:** {sinal_num(r.get('Diferença Qtd', 0), 1)} unidade(s)")
            linhas.append("")
        if len(diverg_qtd) > limite_itens:
            linhas.append(f"- Existem mais **{format_int_br(len(diverg_qtd) - limite_itens)}** divergência(s) de quantidade não listadas neste texto.")
            linhas.append("")

    linhas.append("## 💰 Divergências de Preço")
    linhas.append("")
    if diverg_preco.empty:
        linhas.append("- Não foram identificadas divergências de preço unitário.")
    else:
        linhas.append("Foram identificadas divergências nos preços unitários dos seguintes itens:")
        linhas.append("")
        for _, r in diverg_preco.head(limite_itens).iterrows():
            status = str(r.get("Status", ""))
            item = nome_item(r)
            linhas.append(f"### 🔸 {item}")
            if status == "Somente fornecedor":
                linhas.append("- **Situação:** item consta somente no pedido do fornecedor.")
                linhas.append(f"- **Preço fornecedor:** {format_moeda_br(r.get('Preço Fornecedor', 0))}")
                linhas.append("- **Preço Pedido Única:** R$ 0,00")
                linhas.append("- **Ação recomendada:** validar se o item deve entrar no pedido antes da aprovação.")
            elif status == "Não encontrado no fornecedor":
                linhas.append("- **Situação:** item consta somente no Pedido Única.")
                linhas.append(f"- **Preço Pedido Única:** {format_moeda_br(r.get('Preço Única', 0))}")
                linhas.append("- **Preço fornecedor:** R$ 0,00")
                linhas.append("- **Ação recomendada:** confirmar se o fornecedor retirou o item ou se houve falha no arquivo recebido.")
            else:
                linhas.append(f"- **Preço Pedido Única:** {format_moeda_br(r.get('Preço Única', 0))}")
                linhas.append(f"- **Preço fornecedor:** {format_moeda_br(r.get('Preço Fornecedor', 0))}")
                linhas.append(f"- **Diferença unitária:** {format_moeda_br(r.get('Diferença Preço', 0))}")
                linhas.append(f"- **Diferença percentual:** {sinal_pct(r.get('Diferença Preço %', 0))}")
            linhas.append("")
        if len(diverg_preco) > limite_itens:
            linhas.append(f"- Existem mais **{format_int_br(len(diverg_preco) - limite_itens)}** divergência(s) de preço não listadas neste texto.")
            linhas.append("")

    somente_forn = df[df["Status"] == "Somente fornecedor"].copy()
    nao_encontrado = df[df["Status"] == "Não encontrado no fornecedor"].copy()

    linhas.append("## ⚠️ Itens Apenas no Fornecedor")
    linhas.append("")
    if somente_forn.empty:
        linhas.append("- Nenhum item foi encontrado apenas no arquivo do fornecedor.")
    else:
        linhas.append("Os itens abaixo foram enviados pelo fornecedor, porém não constam no Pedido Única considerado para comparação:")
        linhas.append("")
        for _, r in somente_forn.head(limite_itens).iterrows():
            linhas.append(f"- **{nome_item(r)}** | Qtd: **{format_num_br(r.get('Qtd Fornecedor', 0), 1)}** | Preço: **{format_moeda_br(r.get('Preço Fornecedor', 0))}**")
        if len(somente_forn) > limite_itens:
            linhas.append(f"- Mais **{format_int_br(len(somente_forn) - limite_itens)}** item(ns) não listado(s).")
    linhas.append("")

    linhas.append("## ⚠️ Itens do Pedido Única não encontrados no fornecedor")
    linhas.append("")
    if nao_encontrado.empty:
        linhas.append("- Nenhum item do Pedido Única ficou sem correspondência no fornecedor.")
    else:
        linhas.append("Os itens abaixo constam no Pedido Única, mas não foram encontrados no arquivo do fornecedor:")
        linhas.append("")
        for _, r in nao_encontrado.head(limite_itens).iterrows():
            linhas.append(f"- **{nome_item(r)}** | Qtd: **{format_num_br(r.get('Qtd Única', 0), 1)}** | Preço: **{format_moeda_br(r.get('Preço Única', 0))}**")
        if len(nao_encontrado) > limite_itens:
            linhas.append(f"- Mais **{format_int_br(len(nao_encontrado) - limite_itens)}** item(ns) não listado(s).")
    linhas.append("")

    percentual_ok = (qtd_ok / total * 100) if total else 0
    percentual_div = ((total - qtd_ok) / total * 100) if total else 0
    linhas.append("## 📊 Indicadores da Conferência")
    linhas.append("")
    linhas.append(f"- **Percentual sem divergência:** {format_num_br(percentual_ok, 2)}%")
    linhas.append(f"- **Percentual com algum apontamento:** {format_num_br(percentual_div, 2)}%")
    linhas.append(f"- **Total de divergências de quantidade:** {format_int_br(len(diverg_qtd))}")
    linhas.append(f"- **Total de divergências de preço:** {format_int_br(len(diverg_preco))}")
    linhas.append("")

    linhas.append("## ✅ Conclusão")
    linhas.append("")
    if qtd_div == 0 and qtd_nao_fornecedor == 0 and qtd_somente_fornecedor == 0:
        linhas.append("A conferência foi finalizada sem divergências relevantes. O pedido pode seguir para validação final.")
    else:
        linhas.append("A conferência encontrou pontos que precisam ser validados antes da aprovação final do pedido:")
        if len(diverg_qtd) > 0:
            linhas.append(f"- **{format_int_br(len(diverg_qtd))}** divergência(s) de quantidade.")
        if len(diverg_preco) > 0:
            linhas.append(f"- **{format_int_br(len(diverg_preco))}** divergência(s) de preço.")
        if qtd_somente_fornecedor > 0:
            linhas.append(f"- **{format_int_br(qtd_somente_fornecedor)}** item(ns) enviado(s) pelo fornecedor que não constam no Pedido Única.")
        if qtd_nao_fornecedor > 0:
            linhas.append(f"- **{format_int_br(qtd_nao_fornecedor)}** item(ns) do Pedido Única não encontrado(s) no fornecedor.")
        linhas.append("")
        linhas.append("Recomenda-se validar as divergências com o fornecedor antes de aprovar o pedido.")

    return "\n".join(linhas).strip()


def gerar_texto_divergencias_comparativo(df_comparativo):
    """
    Mantida por compatibilidade com o restante do app.
    Agora retorna o relatório executivo completo e um texto focado em preço.
    """
    relatorio = gerar_relatorio_executivo_comparativo(df_comparativo)
    return relatorio, relatorio

def colorir_comparativo_pedidos(row):
    status = str(row.get("Status", ""))
    if status == "OK":
        return ["background-color: #eaf7ea"] * len(row)
    if status == "Divergente":
        return ["background-color: #fff7ed"] * len(row)
    return ["background-color: #ffe8e8; font-weight: 600"] * len(row)


def gerar_excel_comparativo_pedidos(df_comparativo):
    if Workbook is None:
        raise RuntimeError("A biblioteca openpyxl não está instalada. Rode: python -m pip install openpyxl")
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo"
    ws.append(list(df_comparativo.columns))
    for _, row in df_comparativo.iterrows():
        ws.append([row.get(col, "") for col in df_comparativo.columns])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)
    for idx, col in enumerate(df_comparativo.columns, start=1):
        letter = ws.cell(row=1, column=idx).column_letter
        if "Descrição" in col:
            ws.column_dimensions[letter].width = 42
        elif "%" in col:
            ws.column_dimensions[letter].width = 16
            for cell in ws[letter][1:]:
                cell.number_format = '0.00%'
                try:
                    cell.value = float(cell.value or 0) / 100
                except Exception:
                    pass
        elif "Preço" in col or "Valor" in col:
            ws.column_dimensions[letter].width = 16
            for cell in ws[letter][1:]:
                cell.number_format = 'R$ #,##0.00'
        elif "Qtd" in col:
            ws.column_dimensions[letter].width = 13
            for cell in ws[letter][1:]:
                cell.number_format = '#,##0.0'
        else:
            ws.column_dimensions[letter].width = 22
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def render_pagina_comparativo_pedidos():
    st.markdown('<div class="section-title">Comparativo de Pedidos</div>', unsafe_allow_html=True)
    st.caption("Novo modelo: o sistema compara somente pelo Código de Fábrica. A descrição fica apenas para conferência visual e não é usada para vínculo automático, evitando relacionamentos errados.")

    if "relacionamentos_comparativo" not in st.session_state:
        st.session_state["relacionamentos_comparativo"] = {}

    col1, col2 = st.columns(2)
    with col1:
        pedido_unica = st.file_uploader("Planilha do pedido da Única", type=["xlsx", "xls", "csv", "html", "htm"], key="upload_comparativo_unica")
    with col2:
        pedido_fornecedor = st.file_uploader("Pedido do fornecedor", type=["xlsx", "xls", "csv", "pdf", "html", "htm"], key="upload_comparativo_fornecedor")

    if not pedido_unica or not pedido_fornecedor:
        st.info("Envie o pedido da Única e o arquivo do fornecedor para iniciar o comparativo.")
        return

    try:
        df_unica = ler_arquivo_comparativo(pedido_unica)
        if df_unica.empty:
            st.error("Não consegui ler o pedido da Única.")
            return

        st.markdown("### 1. Conferência e mapeamento das colunas")
        with st.expander("Prévia do Pedido Única", expanded=False):
            st.dataframe(df_unica.head(30), use_container_width=True, hide_index=True)

        mapa_unica = _mapear_colunas_comparativo(df_unica, "cmp_unica", "Pedido Única")
        if not mapa_unica:
            return

        codigos_ref = codigos_referencia_comparativo(df_unica, mapa_unica)
        df_fornecedor = ler_arquivo_comparativo(pedido_fornecedor, codigos_referencia=codigos_ref)
        if df_fornecedor.empty:
            st.error("Não consegui ler o arquivo do fornecedor.")
            return

        with st.expander("Prévia do Pedido Fornecedor", expanded=False):
            st.dataframe(df_fornecedor.head(30), use_container_width=True, hide_index=True)

        mapa_fornecedor = _mapear_colunas_comparativo(df_fornecedor, "cmp_fornecedor", "Pedido Fornecedor")
        if not mapa_fornecedor:
            return

        st.markdown("### 2. Relacionamento manual dos itens sem identificação")
        relacionamentos = st.session_state.get("relacionamentos_comparativo", {})
        comparativo_base = montar_comparativo_pedidos(df_unica, df_fornecedor, mapa_unica, mapa_fornecedor, relacionamentos)

        nao_encontrados = comparativo_base[comparativo_base["Status"] == "Não encontrado no fornecedor"].copy()
        somente_fornecedor = comparativo_base[comparativo_base["Status"] == "Somente fornecedor"].copy()

        if not nao_encontrados.empty and not somente_fornecedor.empty:
            with st.expander("Relacionar manualmente itens não encontrados", expanded=True):
                st.caption("O vínculo automático é feito apenas por Código de Fábrica. Use esta tela somente para relacionar manualmente códigos que não bateram entre os arquivos.")
                opcoes_fornecedor = {"-- Não relacionar --": ""}
                for _, r in somente_fornecedor.iterrows():
                    label = f'{str(r.get("Código Fornecedor", "")).strip()} | {str(r.get("Descrição Fornecedor", "")).strip()} | Qtd {format_num_br(r.get("Qtd Fornecedor", 0), 1)} | R$ {format_num_br(r.get("Preço Fornecedor", 0), 2)}'
                    opcoes_fornecedor[label[:250]] = str(r.get("Chave Fornecedor", ""))

                novos_rel = dict(relacionamentos)
                limite_manual = min(len(nao_encontrados), 80)
                for i, (_, r) in enumerate(nao_encontrados.head(limite_manual).iterrows()):
                    st.markdown(f'**Única:** {str(r.get("Código Única", "")).strip()} | {str(r.get("Descrição Única", "")).strip()} | Qtd {format_num_br(r.get("Qtd Única", 0), 1)}')
                    escolha = st.selectbox(
                        "Item correspondente no fornecedor",
                        list(opcoes_fornecedor.keys()),
                        key=f"rel_manual_{i}_{str(r.get('Chave Única', ''))[:20]}",
                    )
                    chave_f = opcoes_fornecedor.get(escolha, "")
                    if chave_f:
                        novos_rel[str(r.get("Chave Única", ""))] = chave_f

                if len(nao_encontrados) > limite_manual:
                    st.info(f"Mostrando os primeiros {limite_manual} itens para relacionamento manual.")

                c_rel1, c_rel2 = st.columns(2)
                if c_rel1.button("Aplicar relacionamentos manuais", type="primary"):
                    st.session_state["relacionamentos_comparativo"] = novos_rel
                    st.rerun()
                if c_rel2.button("Limpar relacionamentos manuais"):
                    st.session_state["relacionamentos_comparativo"] = {}
                    st.rerun()
        else:
            st.info("Não há itens pendentes para relacionamento manual neste momento.")

        st.markdown("### 3. Resultado do comparativo")
        comparativo = montar_comparativo_pedidos(
            df_unica,
            df_fornecedor,
            mapa_unica,
            mapa_fornecedor,
            st.session_state.get("relacionamentos_comparativo", {}),
        )
        st.success(f"Comparativo gerado com {len(comparativo)} linha(s).")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("OK", int((comparativo["Status"] == "OK").sum()))
        c2.metric("Divergentes", int((comparativo["Status"] == "Divergente").sum()))
        c3.metric("Não encontrados", int((comparativo["Status"] == "Não encontrado no fornecedor").sum()))
        c4.metric("Somente fornecedor", int((comparativo["Status"] == "Somente fornecedor").sum()))

        relatorio_executivo = gerar_relatorio_executivo_comparativo(comparativo)
        with st.expander("📋 Relatório executivo da conferência", expanded=True):
            st.markdown(relatorio_executivo)
            st.text_area(
                "Texto pronto para copiar",
                relatorio_executivo,
                height=420,
                key="texto_relatorio_executivo_comparativo",
            )

        termo = st.text_input("Pesquisar no comparativo", key="busca_comparativo_pedidos").strip().lower()
        view = comparativo.copy()
        if termo:
            filtro = pd.Series(False, index=view.index)
            for col in ["Código Única", "Descrição Única", "Código Fornecedor", "Descrição Fornecedor", "Status", "Método"]:
                if col in view.columns:
                    filtro = filtro | view[col].astype(str).str.lower().str.contains(termo, na=False)
            view = view[filtro].copy()

        colunas_ocultar = ["Chave Única", "Chave Fornecedor"]
        view_exibicao = view.drop(columns=colunas_ocultar, errors="ignore")
        st.dataframe(
            view_exibicao.style.apply(colorir_comparativo_pedidos, axis=1).format(formatadores_para_tabela(view_exibicao)),
            use_container_width=True,
            hide_index=True,
            height=620,
        )

        st.download_button(
            "Baixar comparativo em Excel",
            gerar_excel_comparativo_pedidos(comparativo.drop(columns=colunas_ocultar, errors="ignore")),
            "comparativo_pedidos.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(str(e))

def inicializar_pedido_editavel(tabela_resumo):
    colunas_base = colunas_pedido_compras(MESES)

    base = tabela_resumo.copy()
    if "Estoque Geral" not in base.columns and "Estoque Atual Geral" in base.columns:
        base["Estoque Geral"] = base["Estoque Atual Geral"]

    for col in colunas_base:
        if col not in base.columns:
            base[col] = 0 if col not in ["codigo", "descricao", "Código Fábrica", "Data Última Compra"] else ""

    # garante existência das colunas calculadas antes de ordenar
    base["PEDIDO Final"] = pd.to_numeric(base["Sugestão arredondada"], errors="coerce").fillna(0).round(0).astype(int)
    base["Origem Sugestão"] = "Sugestão do sistema"
    base["Valor Final do Pedido"] = base["PEDIDO Final"] * pd.to_numeric(base["Preço Última Compra"], errors="coerce").fillna(0)
    base = base[colunas_base].copy()
    return base

def atualizar_valor_e_origem(df):
    df = df.copy()
    df["PEDIDO Final"] = pd.to_numeric(df.get("PEDIDO Final", 0), errors="coerce").fillna(0).round(0).astype(int)
    df["Sugestão Sistema"] = pd.to_numeric(df.get("Sugestão Sistema", 0), errors="coerce").fillna(0).round(0).astype(int)
    df["Sugestão arredondada"] = pd.to_numeric(df.get("Sugestão arredondada", df["Sugestão Sistema"]), errors="coerce").fillna(0).round(0).astype(int)
    df["Preço Última Compra"] = pd.to_numeric(df.get("Preço Última Compra", 0), errors="coerce").fillna(0)
    df["Valor Final do Pedido"] = df["PEDIDO Final"] * df["Preço Última Compra"]
    df["Origem Sugestão"] = df.apply(
        lambda row: "Sugestão do sistema" if int(row["PEDIDO Final"]) == int(row["Sugestão arredondada"]) else "Alterado pelo usuário",
        axis=1,
    )
    return df




def ajustar_pedido_para_multiplo_embalagem(qtd, embalagem):
    """
    Valida o PEDIDO Final pelo múltiplo da embalagem.
    Se a quantidade não for múltipla, ajusta sempre para o próximo múltiplo acima.
    Ex.: qtd 45 e embalagem 12 => 48.
    """
    try:
        qtd = int(round(float(qtd or 0)))
    except Exception:
        qtd = 0

    try:
        embalagem = int(round(float(embalagem or 0)))
    except Exception:
        embalagem = 0

    if qtd <= 0:
        return 0
    if embalagem <= 1:
        return qtd
    if qtd % embalagem == 0:
        return qtd
    return int(math.ceil(qtd / embalagem) * embalagem)


def validar_pedidos_por_embalagem(df):
    """
    Ajusta todos os pedidos para múltiplos da embalagem e devolve mensagens de alerta.
    """
    df = df.copy()
    mensagens = []

    if "Embalagem" not in df.columns:
        df["Embalagem"] = 0

    df["PEDIDO Final"] = pd.to_numeric(df.get("PEDIDO Final", 0), errors="coerce").fillna(0).round(0).astype(int)
    df["Embalagem"] = pd.to_numeric(df.get("Embalagem", 0), errors="coerce").fillna(0).round(0).astype(int)

    for idx, row in df.iterrows():
        qtd_original = int(row.get("PEDIDO Final", 0) or 0)
        embalagem = int(row.get("Embalagem", 0) or 0)
        qtd_ajustada = ajustar_pedido_para_multiplo_embalagem(qtd_original, embalagem)

        if qtd_original > 0 and embalagem > 1 and qtd_original != qtd_ajustada:
            codigo = str(row.get("codigo", "")).zfill(5)
            descricao = str(row.get("descricao", "")).strip()
            mensagens.append(
                f"Item {codigo} - {descricao}: a embalagem é com {embalagem} unidades. "
                f"O pedido {qtd_original} foi ajustado para {qtd_ajustada}."
            )
            df.at[idx, "PEDIDO Final"] = qtd_ajustada

    return df, mensagens

def totalizar_valor_pedido(df):
    if df.empty:
        return 0.0
    qtd = pd.to_numeric(df.get("PEDIDO Final", 0), errors="coerce").fillna(0)
    preco = pd.to_numeric(df.get("Preço Última Compra", 0), errors="coerce").fillna(0)
    return float((qtd * preco).sum())



# =========================================================
# UI / EXPERIÊNCIA DO USUÁRIO
# =========================================================

APP_NAME = "Análise de Giro e Pedido de Compra"


def aplicar_css_global():
    st.markdown(
        """
        <style>
            :root {
                --primary: #1d4ed8;
                --primary-soft: #eff6ff;
                --bg-soft: #f8fafc;
                --border: #e2e8f0;
                --text-muted: #64748b;
            }

            .main .block-container {
                padding-top: 1.25rem;
                padding-bottom: 2.5rem;
                max-width: 1500px;
            }

            [data-testid="stSidebar"] {
                background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
            }

            [data-testid="stSidebar"] * {
                color: #f8fafc !important;
            }

            [data-testid="stSidebar"] .stRadio label,
            [data-testid="stSidebar"] .stNumberInput label {
                color: #e2e8f0 !important;
            }

            .hero-card {
                background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 52%, #2563eb 100%);
                color: white;
                border-radius: 24px;
                padding: 28px 32px;
                margin-bottom: 22px;
                box-shadow: 0 18px 45px rgba(15, 23, 42, .18);
            }

            .hero-card h1 {
                margin: 0;
                font-size: 34px;
                line-height: 1.1;
                color: white;
                letter-spacing: -0.02em;
            }

            .hero-card p {
                margin: 10px 0 0 0;
                font-size: 15px;
                color: #dbeafe;
            }

            .section-title {
                font-size: 22px;
                font-weight: 800;
                margin: 20px 0 8px 0;
                color: #0f172a;
            }

            .muted {
                color: var(--text-muted);
                font-size: 14px;
            }

            .metric-card {
                background: white;
                border: 1px solid var(--border);
                border-radius: 18px;
                padding: 18px 20px;
                min-height: 118px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, .06);
            }

            .metric-card .label {
                color: #64748b;
                font-size: 13px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: .04em;
            }

            .metric-card .value {
                margin-top: 8px;
                color: #0f172a;
                font-size: 28px;
                font-weight: 850;
                letter-spacing: -0.03em;
            }

            .metric-card .hint {
                margin-top: 5px;
                color: #64748b;
                font-size: 13px;
            }

            .upload-card {
                background: #ffffff;
                border: 1px solid var(--border);
                border-radius: 18px;
                padding: 16px 18px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, .05);
                margin-bottom: 8px;
            }

            .upload-card strong {
                color: #0f172a;
                font-size: 16px;
            }

            .status-ok {
                color: #047857;
                font-weight: 800;
            }

            .status-warn {
                color: #b45309;
                font-weight: 800;
            }

            .download-card {
                background: white;
                border: 1px solid var(--border);
                border-radius: 18px;
                padding: 18px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, .05);
                margin-bottom: 14px;
            }

            div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {
                border-radius: 16px;
                overflow: hidden;
                border: 1px solid #e2e8f0;
            }

            .stButton > button, .stDownloadButton > button {
                border-radius: 12px !important;
                font-weight: 800 !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header(subtitulo="Sistema de apoio à decisão para giro, estoque e compra."):
    st.markdown(
        f"""
        <div class="hero-card">
            <h1>{APP_NAME}</h1>
            <p>{subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric(label, value, hint=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
            <div class="hint">{hint}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_status(titulo, arquivo, obrigatorio=False):
    status = "✓ Arquivo carregado" if arquivo else ("Obrigatório" if obrigatorio else "Opcional")
    classe = "status-ok" if arquivo else "status-warn"
    nome = getattr(arquivo, "name", "") if arquivo else ""
    st.markdown(
        f"""
        <div class="upload-card">
            <strong>{titulo}</strong><br>
            <span class="{classe}">{status}</span><br>
            <span class="muted">{nome}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpis_gerais(tabela_resumo, pedido_df=None):
    pedido_ref = pedido_df.copy() if pedido_df is not None and not pedido_df.empty else inicializar_pedido_editavel(tabela_resumo)
    pedido_ref = atualizar_valor_e_origem(pedido_ref)
    total_produtos = len(tabela_resumo)
    itens_compra = int((pd.to_numeric(pedido_ref.get("PEDIDO Final", 0), errors="coerce").fillna(0) > 0).sum())
    valor_pedido = totalizar_valor_pedido(pedido_ref)
    sem_compra = int(tabela_resumo.get("Data Última Compra", pd.Series(dtype=str)).astype(str).str.contains("⚠️", na=False).sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_metric("Produtos analisados", format_int_br(total_produtos), "Itens processados no giro")
    with c2:
        render_metric("Itens com compra", format_int_br(itens_compra), "Pedido final maior que zero")
    with c3:
        render_metric("Valor do pedido", format_moeda_br(valor_pedido), "Quantidade × última compra")
    with c4:
        render_metric("Alertas sem compra", format_int_br(sem_compra), "Conforme parâmetro definido")


def render_download_card(titulo, descricao):
    st.markdown(
        f"""
        <div class="download-card">
            <strong>{titulo}</strong><br>
            <span class="muted">{descricao}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )



# =========================================================
# MULTIPAGE INDEPENDENTE - RUPTURA POR MARCA
# =========================================================

def _token_numero_br_para_float(token):
    return br_to_float(token)


def _extrair_meses_cabecalho_marca(text):
    meses = []
    match = re.search(r"REFERENTE AOS MESES:\s*([^\n]+)", text, flags=re.IGNORECASE)
    if match:
        meses = re.findall(r"\d{2}/\d{4}", match.group(1))
    if not meses:
        meses = sorted(set(re.findall(r"\b\d{2}/\d{4}\b", text)))[:4]
    return meses[:6] if meses else []


def parse_linha_giro_marca_independente(line, meses_ref):
    """
    Parser independente para o relatório de Giro Geral por Marca.
    Não interfere na lógica atual do dashboard.
    """
    line = str(line).strip()
    if not re.match(r"^\d{5}\s+", line):
        return None

    partes = line.split()
    if len(partes) < 10:
        return None

    codigo = partes[0].zfill(5)
    qtd_meses = len(meses_ref) if meses_ref else 4

    un_index = None
    for i in range(1, len(partes) - (qtd_meses + 3)):
        proximos = partes[i + 1:i + 1 + qtd_meses + 3]
        if len(proximos) >= qtd_meses + 3 and all(_eh_numero_br(t) for t in proximos[:qtd_meses + 3]):
            un_index = i
            break

    if un_index is None:
        # fallback para relatórios que usam UN como unidade principal
        for i, token in enumerate(partes):
            if token.upper() in ["UN", "LT", "GL", "CX", "PC", "MT", "KG", "DC"]:
                proximos = partes[i + 1:i + 1 + qtd_meses + 3]
                if len(proximos) >= qtd_meses + 3 and all(_eh_numero_br(t) for t in proximos[:qtd_meses + 3]):
                    un_index = i
                    break

    if un_index is None:
        return None

    descricao = " ".join(partes[1:un_index]).strip()
    unidade = partes[un_index]
    valores = partes[un_index + 1:]

    if len(valores) < qtd_meses + 3:
        return None

    registro = {
        "codigo": codigo,
        "descricao": descricao,
        "unidade": unidade,
    }

    for idx, mes in enumerate(meses_ref[:qtd_meses]):
        registro[mes] = _token_numero_br_para_float(valores[idx])

    registro["media_pdf"] = _token_numero_br_para_float(valores[qtd_meses])
    registro["previsao_30_pdf"] = _token_numero_br_para_float(valores[qtd_meses + 1])
    registro["estoque"] = _token_numero_br_para_float(valores[qtd_meses + 2])
    return registro


@st.cache_data(show_spinner="Lendo PDF de ruptura por marca...")
def parse_pdf_ruptura_por_marca(bytes_pdf):
    """
    Leitor otimizado para PDF grande.

    O relatório de Giro Geral por Marca pode ter mais de 1.500 páginas.
    pdfplumber é bom para tabelas, mas fica muito lento nesse volume.
    Aqui usamos PyMuPDF primeiro, que extrai texto página a página com muito mais velocidade,
    e só caímos para pdfplumber se PyMuPDF não estiver disponível.
    """
    registros = []
    empresa_atual = None
    marca_cod_atual = ""
    marca_nome_atual = "SEM MARCA"
    meses_ref = []

    def processar_texto_pagina(page_text):
        nonlocal empresa_atual, marca_cod_atual, marca_nome_atual, meses_ref, registros

        if not meses_ref:
            meses_extraidos = _extrair_meses_cabecalho_marca(page_text or "")
            if meses_extraidos:
                meses_ref = meses_extraidos

        if not meses_ref:
            return

        for raw_line in str(page_text or "").splitlines():
            line = raw_line.strip()
            if not line:
                continue

            empresa_match = re.search(r"EMPRESA\s*:\s*(\d{3})\s*-", line, flags=re.IGNORECASE)
            if empresa_match:
                empresa_atual = empresa_match.group(1)
                continue

            marca_match = re.search(r"MARCA\s*:\s*([^\n]+)", line, flags=re.IGNORECASE)
            if marca_match:
                marca_raw = marca_match.group(1).strip()
                marca_partes = re.match(r"([^\-]+)\s*-\s*(.*)", marca_raw)
                if marca_partes:
                    marca_cod_atual = marca_partes.group(1).strip()
                    marca_nome_atual = marca_partes.group(2).strip() or "SEM MARCA"
                else:
                    marca_cod_atual = ""
                    marca_nome_atual = marca_raw or "SEM MARCA"
                continue

            if not empresa_atual or empresa_atual not in LOJAS_MAP:
                continue

            produto = parse_linha_giro_marca_independente(line, meses_ref)
            if produto:
                produto["codigo_empresa"] = empresa_atual
                produto["loja"] = LOJAS_MAP.get(empresa_atual, empresa_atual)
                produto["tipo_unidade"] = "ÚNICA" if empresa_atual == CODIGO_UNICA else "LOJAS DAUTO"
                produto["marca_codigo"] = marca_cod_atual
                produto["marca"] = marca_nome_atual if marca_nome_atual else "SEM MARCA"
                registros.append(produto)

    try:
        import fitz  # PyMuPDF
    except Exception as e:
        raise RuntimeError(
            "Para esta página, instale o PyMuPDF. Rode: pip install pymupdf. "
            "O pdfplumber é lento demais para este relatório com muitas páginas."
        ) from e

    try:
        with fitz.open(stream=bytes_pdf, filetype="pdf") as doc:
            for page in doc:
                processar_texto_pagina(page.get_text("text", sort=True))
    except Exception as e:
        raise RuntimeError(f"Falha ao extrair texto do PDF com PyMuPDF: {e}") from e

    if not meses_ref:
        meses_ref = ["Mês 1", "Mês 2", "Mês 3", "Mês 4"]

    df = pd.DataFrame(registros)
    return df, meses_ref


def classificar_status_ruptura(media_mensal, estoque_geral, dias_cobertura):
    media_mensal = float(media_mensal or 0)
    estoque_geral = float(estoque_geral or 0)
    if media_mensal <= 0:
        return "SEM GIRO"
    if estoque_geral <= 0:
        return "CRÍTICO"
    if dias_cobertura <= 7:
        return "CRÍTICO"
    if dias_cobertura <= 15:
        return "ALTO"
    if dias_cobertura <= 30:
        return "ATENÇÃO"
    return "OK"


def montar_analise_ruptura_por_marca(df_ruptura, meses_ref, df_aberto_ruptura=None):
    if df_ruptura.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = df_ruptura.copy()
    for mes in meses_ref:
        df[mes] = pd.to_numeric(df.get(mes, 0), errors="coerce").fillna(0)
    df["estoque"] = pd.to_numeric(df.get("estoque", 0), errors="coerce").fillna(0)

    agg_dict = {mes: "sum" for mes in meses_ref}
    agg_dict.update({
        "estoque": "sum",
        "marca_codigo": "first",
        "unidade": "first",
    })

    itens = df.groupby(["marca", "codigo", "descricao"], as_index=False).agg(agg_dict)
    itens["Giro Geral"] = itens[meses_ref].sum(axis=1)
    itens["Média Giro Geral"] = itens[meses_ref].mean(axis=1).round(2)
    itens["Estoque Geral"] = itens["estoque"].round(2)

    if df_aberto_ruptura is not None and not df_aberto_ruptura.empty:
        aberto = df_aberto_ruptura.copy()
        aberto["codigo"] = aberto["codigo"].astype(str).str.extract(r"(\d+)")[0].str.zfill(5)
        aberto["Saldo em Trânsito/ABERTO"] = pd.to_numeric(aberto.get("Saldo em Trânsito/ABERTO", 0), errors="coerce").fillna(0)
        aberto = aberto.groupby("codigo", as_index=False)["Saldo em Trânsito/ABERTO"].sum()
        itens = itens.merge(aberto, on="codigo", how="left")
    else:
        itens["Saldo em Trânsito/ABERTO"] = 0

    itens["Saldo em Trânsito/ABERTO"] = pd.to_numeric(itens["Saldo em Trânsito/ABERTO"], errors="coerce").fillna(0).round(2)
    itens["Estoque Considerado"] = (itens["Estoque Geral"] + itens["Saldo em Trânsito/ABERTO"]).round(2)
    itens["Dias de Cobertura"] = itens.apply(
        lambda r: round((float(r["Estoque Considerado"]) / float(r["Média Giro Geral"]) * 30), 1) if float(r["Média Giro Geral"] or 0) > 0 else 9999,
        axis=1,
    )
    itens["Status"] = itens.apply(
        lambda r: classificar_status_ruptura(r["Média Giro Geral"], r["Estoque Considerado"], r["Dias de Cobertura"]),
        axis=1,
    )
    itens["Necessidade 30 dias"] = itens.apply(
        lambda r: max(math.ceil(float(r["Média Giro Geral"] or 0) - float(r["Estoque Considerado"] or 0)), 0),
        axis=1,
    )
    itens["Peso Risco"] = itens["Status"].map({"CRÍTICO": 4, "ALTO": 3, "ATENÇÃO": 2, "OK": 1, "SEM GIRO": 0}).fillna(0)

    resumo = itens.groupby("marca", as_index=False).agg(
        Itens=("codigo", "count"),
        Criticos=("Status", lambda s: int((s == "CRÍTICO").sum())),
        Alto=("Status", lambda s: int((s == "ALTO").sum())),
        Atencao=("Status", lambda s: int((s == "ATENÇÃO").sum())),
        OK=("Status", lambda s: int((s == "OK").sum())),
        Sem_Giro=("Status", lambda s: int((s == "SEM GIRO").sum())),
        Giro_Geral=("Giro Geral", "sum"),
        Media_Giro_Geral=("Média Giro Geral", "sum"),
        Estoque_Geral=("Estoque Geral", "sum"),
        Em_Aberto=("Saldo em Trânsito/ABERTO", "sum"),
        Estoque_Considerado=("Estoque Considerado", "sum"),
        Necessidade_30_dias=("Necessidade 30 dias", "sum"),
        Score_Risco=("Peso Risco", "sum"),
    )
    resumo["% Itens em Risco"] = ((resumo["Criticos"] + resumo["Alto"] + resumo["Atencao"]) / resumo["Itens"].replace(0, pd.NA) * 100).fillna(0).round(1)
    resumo["Dias Cobertura Marca"] = resumo.apply(
        lambda r: round((float(r["Estoque_Considerado"]) / float(r["Media_Giro_Geral"]) * 30), 1) if float(r["Media_Giro_Geral"] or 0) > 0 else 9999,
        axis=1,
    )
    resumo = resumo.sort_values(["Score_Risco", "Criticos", "Alto", "% Itens em Risco"], ascending=[False, False, False, False])

    resumo = resumo.rename(columns={
        "marca": "Marca",
        "Criticos": "Críticos",
        "Atencao": "Atenção",
        "Sem_Giro": "Sem Giro",
        "Giro_Geral": "Giro Geral",
        "Media_Giro_Geral": "Média Giro Geral",
        "Estoque_Geral": "Estoque Geral",
        "Em_Aberto": "Em Aberto",
        "Estoque_Considerado": "Estoque Considerado",
        "Necessidade_30_dias": "Necessidade 30 dias",
        "Score_Risco": "Score Risco",
        "Dias Cobertura Marca": "Dias de Cobertura",
    })

    itens = itens.rename(columns={
        "marca": "Marca",
        "codigo": "Código",
        "descricao": "Descrição",
        "unidade": "UN",
    })
    itens = itens.sort_values(["Peso Risco", "Dias de Cobertura", "Média Giro Geral"], ascending=[False, True, False])
    return resumo.reset_index(drop=True), itens.reset_index(drop=True)


def colorir_status_ruptura(row):
    status = str(row.get("Status", ""))
    if status == "CRÍTICO":
        return ["background-color: #fee2e2; color: #7f1d1d; font-weight: 700"] * len(row)
    if status == "ALTO":
        return ["background-color: #ffedd5; color: #7c2d12; font-weight: 650"] * len(row)
    if status == "ATENÇÃO":
        return ["background-color: #fef9c3; color: #713f12"] * len(row)
    if status == "OK":
        return ["background-color: #dcfce7; color: #14532d"] * len(row)
    return ["background-color: #f1f5f9; color: #475569"] * len(row)


def colorir_resumo_marca(row):
    criticos = int(row.get("Críticos", 0) or 0)
    alto = int(row.get("Alto", 0) or 0)
    atencao = int(row.get("Atenção", 0) or 0)
    if criticos > 0:
        return ["background-color: #fee2e2; color: #7f1d1d; font-weight: 700"] * len(row)
    if alto > 0:
        return ["background-color: #ffedd5; color: #7c2d12; font-weight: 650"] * len(row)
    if atencao > 0:
        return ["background-color: #fef9c3; color: #713f12"] * len(row)
    return [""] * len(row)


def render_pagina_ruptura_por_marca():
    st.markdown('<div class="section-title">🏷️ Ruptura por Marca</div>', unsafe_allow_html=True)
    st.caption(
        "Esta página funciona separada do pedido de compra. Envie aqui o PDF específico de Giro Geral por Marca "
        "e, se houver, o PDF de Pedidos em Aberto. A análise soma lojas Dauto + Única e considera o saldo em aberto no estoque."
    )

    col_pdf_ruptura, col_pdf_aberto = st.columns(2)
    with col_pdf_ruptura:
        pdf_marca = st.file_uploader(
            "PDF - Giro Geral por Marca",
            type=["pdf"],
            key="upload_pdf_ruptura_por_marca_independente",
        )
    with col_pdf_aberto:
        pdf_pedidos_aberto_ruptura = st.file_uploader(
            "PDF - Pedidos em Aberto",
            type=["pdf"],
            key="upload_pdf_pedidos_aberto_ruptura_independente",
        )

    if not pdf_marca:
        st.info("Envie o PDF de Giro Geral por Marca para iniciar esta análise.")
        return

    st.info("Leitura otimizada ativada: o Giro por Marca usa PyMuPDF. O PDF de Pedidos em Aberto, quando enviado, entra como saldo em trânsito no cálculo da ruptura.")

    try:
        bytes_pdf = pdf_marca.getvalue()
        df_ruptura, meses_ref = parse_pdf_ruptura_por_marca(bytes_pdf)
    except Exception as e:
        st.error(f"Não consegui ler o PDF de Ruptura por Marca. Erro: {e}")
        return

    df_aberto_ruptura = pd.DataFrame(columns=["codigo", "Saldo em Trânsito/ABERTO"])
    if pdf_pedidos_aberto_ruptura:
        try:
            with st.spinner("Lendo Pedidos em Aberto para considerar no estoque..."):
                df_aberto_ruptura = parse_pedidos_compra_aberto_pdf(pdf_pedidos_aberto_ruptura)
            st.success(f"Pedidos em aberto lidos: {len(df_aberto_ruptura)} item(ns) com saldo em aberto.")
        except Exception as e:
            st.warning(f"Não consegui ler o PDF de Pedidos em Aberto. A análise seguirá apenas com o estoque atual. Erro: {e}")
            df_aberto_ruptura = pd.DataFrame(columns=["codigo", "Saldo em Trânsito/ABERTO"])

    if df_ruptura.empty:
        st.error("Não consegui extrair os itens do PDF enviado. Verifique se é o relatório de Giro Geral por Marca.")
        return

    resumo_marca, itens_marca = montar_analise_ruptura_por_marca(df_ruptura, meses_ref, df_aberto_ruptura)
    if resumo_marca.empty:
        st.warning("O PDF foi lido, mas não houve dados suficientes para análise.")
        return

    total_itens = int(len(itens_marca))
    total_criticos = int((itens_marca["Status"] == "CRÍTICO").sum())
    total_alto = int((itens_marca["Status"] == "ALTO").sum())
    total_atencao = int((itens_marca["Status"] == "ATENÇÃO").sum())
    total_em_aberto = float(pd.to_numeric(itens_marca.get("Saldo em Trânsito/ABERTO", 0), errors="coerce").fillna(0).sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        render_metric("Itens analisados", format_int_br(total_itens), "Lojas Dauto + Única")
    with c2:
        render_metric("Críticos", format_int_br(total_criticos), "Sem estoque ou até 7 dias")
    with c3:
        render_metric("Alto risco", format_int_br(total_alto), "Até 15 dias")
    with c4:
        render_metric("Atenção", format_int_br(total_atencao), "Até 30 dias")
    with c5:
        render_metric("Em aberto", format_num_br(total_em_aberto, 1), "Somado ao estoque")

    st.markdown("---")
    st.markdown('<div class="section-title">Ranking de marcas por risco</div>', unsafe_allow_html=True)

    busca_marca = st.text_input("Pesquisar marca", key="busca_marca_ruptura")
    resumo_view = resumo_marca.copy()
    if busca_marca:
        resumo_view = resumo_view[resumo_view["Marca"].astype(str).str.lower().str.contains(busca_marca.lower(), na=False)]

    st.dataframe(
        resumo_view.style.apply(colorir_resumo_marca, axis=1).format(formatadores_para_tabela(resumo_view)),
        use_container_width=True,
        hide_index=True,
        height=430,
        column_config={"Marca": st.column_config.TextColumn("Marca", pinned=True, width="large")},
    )

    st.download_button(
        "⬇️ Baixar ranking de marcas em CSV",
        gerar_csv(resumo_marca),
        "ranking_ruptura_por_marca.csv",
        "text/csv",
    )

    st.markdown("---")
    st.markdown('<div class="section-title">Drill por marca</div>', unsafe_allow_html=True)
    marcas = resumo_marca["Marca"].astype(str).tolist()
    marca_selecionada = st.selectbox("Selecione a marca para abrir os produtos", marcas, key="drill_marca_ruptura")

    itens_view = itens_marca[itens_marca["Marca"].astype(str) == str(marca_selecionada)].copy()

    colf1, colf2 = st.columns([1, 1])
    with colf1:
        status_opcoes = ["Todos", "CRÍTICO", "ALTO", "ATENÇÃO", "OK", "SEM GIRO"]
        status_sel = st.selectbox("Filtrar status", status_opcoes, key="status_ruptura_marca")
    with colf2:
        busca_item = st.text_input("Pesquisar produto dentro da marca", key="busca_item_ruptura_marca")

    if status_sel != "Todos":
        itens_view = itens_view[itens_view["Status"] == status_sel]
    if busca_item:
        termo = busca_item.lower()
        itens_view = itens_view[
            itens_view["Código"].astype(str).str.lower().str.contains(termo, na=False)
            | itens_view["Descrição"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    colunas_itens = ["Código", "Descrição", "UN"] + meses_ref + [
        "Giro Geral", "Média Giro Geral", "Estoque Geral", "Saldo em Trânsito/ABERTO", "Estoque Considerado", "Dias de Cobertura", "Necessidade 30 dias", "Status"
    ]
    colunas_itens = [c for c in colunas_itens if c in itens_view.columns]

    st.dataframe(
        itens_view[colunas_itens].style.apply(colorir_status_ruptura, axis=1).format(formatadores_para_tabela(itens_view[colunas_itens])),
        use_container_width=True,
        hide_index=True,
        height=560,
        column_config={
            "Código": st.column_config.TextColumn("Código", pinned=True, width="small"),
            "Descrição": st.column_config.TextColumn("Descrição", pinned=True, width="large"),
        },
    )

    st.download_button(
        "⬇️ Baixar drill da marca em CSV",
        gerar_csv(itens_view[colunas_itens]),
        f"drill_ruptura_{re.sub(r'[^A-Za-z0-9]+', '_', str(marca_selecionada))}.csv",
        "text/csv",
    )

def render_pagina_pedidos_drive():
    st.markdown('<div class="section-title">Pedidos no Google Drive</div>', unsafe_allow_html=True)

    if not google_configurado():
        st.warning(google_mensagem_configuracao())
        st.code(
            """
[google_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "sua-conta@projeto.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
""".strip(),
            language="toml",
        )
        st.stop()

    try:
        recursos = google_get_resources()
        st.success(f"Google Drive conectado: {recursos.get('client_email', '')}")
        c1, c2, c3, c4 = st.columns(4)
        c1.link_button("Pasta raiz", recursos["root_link"])
        c2.link_button("Pedidos editaveis", recursos["pedidos_link"])
        c3.link_button("Arquivos finais", recursos["finais_link"])
        c4.link_button("Controle", recursos["controle_link"])

        pedidos = google_listar_pedidos()
        if pedidos.empty:
            st.info("Ainda nao existem pedidos registrados no Drive.")
            st.stop()

        pedidos_view = pedidos.copy()
        pedidos_view["valor"] = pd.to_numeric(pedidos_view["valor"], errors="coerce").fillna(0)
        st.markdown("### Painel de pedidos")
        st.dataframe(
            pedidos_view[[
                "id_pedido", "nome_pedido", "fornecedor", "status", "valor",
                "criado_em", "criado_por", "aprovado_em", "aprovado_por",
                "link_pedido", "link_autcom", "link_fornecedor",
            ]],
            use_container_width=True,
            hide_index=True,
            height=420,
            column_config={
                "valor": st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                "link_pedido": st.column_config.LinkColumn("Pedido"),
                "link_autcom": st.column_config.LinkColumn("Autcom"),
                "link_fornecedor": st.column_config.LinkColumn("Fornecedor"),
            },
        )

        st.markdown("### Aprovar ou atualizar status")
        opcoes = {
            f"{r.get('id_pedido', '')} | {r.get('fornecedor', '')} | {r.get('nome_pedido', '')} | {r.get('status', '')}": r.get("id_pedido", "")
            for _, r in pedidos.iterrows()
        }
        pedido_label = st.selectbox("Pedido", list(opcoes.keys()), key="drive_pedido_status")
        usuario = st.text_input("Usuario responsavel", value="", key="drive_usuario_status")
        status = st.selectbox("Status", ["Aprovado", "Em edicao", "Reprovado", "Finalizado"], key="drive_status")
        observacao = st.text_input("Observacao", key="drive_observacao_status")

        if st.button("Salvar status no controle", type="primary"):
            atualizado = google_atualizar_status_pedido(opcoes[pedido_label], status, usuario=usuario, observacao=observacao)
            st.success(f"Status atualizado para {atualizado.get('status')}.")
            st.rerun()

        try:
            _, sheets_service, _, auth_mode = google_get_services()
            acomp = google_read_df(sheets_service, recursos["controle_id"], "Acompanhamento")
            if not acomp.empty:
                st.markdown("### Acompanhamento mensal")
                acomp["valor"] = pd.to_numeric(acomp.get("valor", 0), errors="coerce").fillna(0)
                resumo_mes = acomp.groupby(["mes", "fornecedor"], as_index=False)["valor"].sum()
                st.dataframe(
                    resumo_mes,
                    use_container_width=True,
                    hide_index=True,
                    column_config={"valor": st.column_config.NumberColumn("Valor", format="R$ %.2f")},
                )
        except Exception:
            pass
    except Exception as e:
        st.error(str(e))


# =========================================================
# APP STREAMLIT
# =========================================================

aplicar_css_global()
render_header()

st.sidebar.markdown("### 📊 Análise de Giro")
pagina = st.sidebar.radio(
    "Navegação",
    ["📦 Giro Consolidado", "🛒 Pedido de Compra", "📄 Exportações", "📁 Pedidos no Drive", "🏷️ Ruptura por Marca", "⚖️ Comparativo de Pedidos", "⚙️ Tratamento Final"],
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Parâmetros")
dias_estoque_alvo = st.sidebar.number_input(
    "Dias de estoque alvo",
    min_value=1,
    max_value=365,
    value=60,
    step=1,
    help="Define quantos dias de cobertura de estoque o pedido deve considerar.",
)

meses_alerta_sem_compra = st.sidebar.number_input(
    "Alerta sem compra acima de quantos meses?",
    min_value=1,
    max_value=36,
    value=3,
    step=1,
    help="Mostra ⚠️ ao lado da data quando a última compra na loja 009 for mais antiga que este parâmetro.",
)

st.sidebar.caption("Estoque Final = Estoque Atual Geral + Saldo em Trânsito/ABERTO")

if pagina == "🏷️ Ruptura por Marca":
    render_pagina_ruptura_por_marca()
    st.stop()

if pagina == "⚖️ Comparativo de Pedidos":
    render_pagina_comparativo_pedidos()
    st.stop()

if pagina == "📁 Pedidos no Drive":
    render_pagina_pedidos_drive()
    st.stop()


st.markdown('<div class="section-title">Upload dos arquivos</div>', unsafe_allow_html=True)
st.caption("Envie o PDF de Giro para iniciar. Os demais arquivos enriquecem a análise e o pedido final.")
col_upload_1, col_upload_2, col_upload_3 = st.columns(3)
cadastro_google = pd.DataFrame()

with col_upload_1:
    giro_pdf = st.file_uploader("PDF - Giro de Estoque", type=["pdf"], key="upload_giro_pdf")
    render_upload_status("📄 Giro de Estoque", giro_pdf, obrigatorio=True)
with col_upload_2:
    pedidos_pdf = st.file_uploader("PDF - Pedidos em Aberto", type=["pdf"], key="upload_pedidos_pdf")
    render_upload_status("📄 Pedidos em Aberto", pedidos_pdf)
with col_upload_3:
    if google_configurado():
        try:
            cadastro_google = ler_cadastro_produtos_google()
            recursos_google = google_get_resources()
            if not cadastro_google.empty:
                st.success(f"Cadastro lido do Google Sheets: {len(cadastro_google)} item(ns).")
            else:
                st.warning("Cadastro do Google Sheets está vazio ou sem as colunas obrigatórias.")
            st.link_button("Abrir cadastro no Drive", recursos_google["cadastro_link"])
        except Exception as e:
            st.warning(f"Não consegui ler o cadastro do Google Sheets: {e}")
    cadastro_csv = st.file_uploader("CSV - Cadastro de Produtos (fallback)", type=["csv"], key="upload_cadastro_csv")
    render_upload_status("📄 Cadastro de Produtos", cadastro_csv)

if pagina == "⚙️ Tratamento Final":
    st.markdown('<div class="section-title">⚙️ Tratamento de Pedido Final</div>', unsafe_allow_html=True)
    st.caption(
        "Envie a planilha final editável. O sistema vai gerar um Excel para importação no Autcom: "
        "coluna B = zx, coluna F = PEDIDO Final e coluna H = Preço Última Compra."
    )

    if google_configurado():
        st.markdown("### Usar pedido aprovado do Google Drive")
        try:
            pedidos_drive = google_listar_pedidos()
            pedidos_aprovados = pedidos_drive[pedidos_drive["status"].astype(str).str.lower().isin(["aprovado", "em edicao"])].copy()
            if pedidos_aprovados.empty:
                st.info("Não há pedidos aprovados ou em edição no controle do Drive.")
            else:
                opcoes_drive = {
                    f"{r.get('id_pedido', '')} | {r.get('fornecedor', '')} | {r.get('nome_pedido', '')} | {r.get('status', '')}": r.to_dict()
                    for _, r in pedidos_aprovados.iterrows()
                }
                pedido_drive_label = st.selectbox("Pedido do Drive", list(opcoes_drive.keys()), key="tratamento_pedido_drive")
                usuario_finalizacao = st.text_input("Finalizado por", value="", key="tratamento_usuario_drive")
                pedido_drive_info = opcoes_drive[pedido_drive_label]

                if st.button("Ler pedido do Drive e gerar arquivos finais", type="primary"):
                    df_drive = google_ler_pedido_drive(pedido_drive_info["spreadsheet_id"])
                    st.session_state["df_tratamento_drive"] = df_drive
                    st.session_state["pedido_tratamento_drive_id"] = pedido_drive_info["id_pedido"]
                    st.success(f"Pedido lido do Drive: {len(df_drive)} linha(s).")

                df_drive_preview = st.session_state.get("df_tratamento_drive")
                pedido_drive_id = st.session_state.get("pedido_tratamento_drive_id")
                if df_drive_preview is not None and pedido_drive_id:
                    colunas_preview_drive = [c for c in ["zx", "codigo", "descricao", "Código Fábrica", "PEDIDO Final", "Preço Última Compra", "Valor Final do Pedido", "Total Geral do Pedido"] if c in df_drive_preview.columns]
                    st.dataframe(
                        df_drive_preview[colunas_preview_drive].head(50) if colunas_preview_drive else df_drive_preview.head(50),
                        use_container_width=True,
                        hide_index=True,
                        height=320,
                    )
                    if st.button("Salvar arquivos finais no Drive e finalizar pedido"):
                        link_autcom, link_fornecedor = google_finalizar_pedido(
                            pedido_drive_id,
                            df_drive_preview,
                            usuario=usuario_finalizacao,
                        )
                        st.success("Pedido finalizado e arquivos salvos no Drive.")
                        c_autcom, c_forn = st.columns(2)
                        c_autcom.link_button("Abrir arquivo Autcom", link_autcom)
                        c_forn.link_button("Abrir arquivo fornecedor", link_fornecedor)
                        st.rerun()
        except Exception as e:
            st.warning(f"Não consegui usar o fluxo do Google Drive: {e}")

        st.markdown("---")
        st.markdown("### Upload manual")

    planilha_tratamento = st.file_uploader(
        "Planilha do Pedido Final",
        type=["xlsx", "xls", "csv"],
        key="upload_tratamento_pedido_final",
    )

    if not planilha_tratamento:
        st.info("Envie a planilha do pedido final para gerar o arquivo de importação Autcom.")
        st.stop()

    try:
        df_tratamento = ler_planilha_tratamento_pedido(planilha_tratamento)
        df_tratamento.columns = [str(c).strip() for c in df_tratamento.columns]

        st.success(f"Planilha lida com sucesso: {len(df_tratamento)} linha(s).")

        colunas_preview = [c for c in ["zx", "descricao", "Código Fábrica", "PEDIDO Final", "Preço Última Compra", "Valor Final do Pedido", "Total Geral do Pedido"] if c in df_tratamento.columns]
        if colunas_preview:
            st.dataframe(
                df_tratamento[colunas_preview].head(50),
                use_container_width=True,
                hide_index=True,
                height=360,
            )
        else:
            st.dataframe(df_tratamento.head(50), use_container_width=True, hide_index=True, height=360)

        excel_tratamento = gerar_excel_autcom_tratamento(df_tratamento)
        col_dl_autcom, col_dl_fornecedor = st.columns(2)
        with col_dl_autcom:
            st.download_button(
                "⬇️ Baixar pedido tratado para importação no Autcom",
                excel_tratamento,
                "pedido_tratado_importacao_autcom.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        with col_dl_fornecedor:
            st.download_button(
                "⬇️ Baixar pedido para envio ao fornecedor",
                gerar_excel_fornecedor_tratamento(df_tratamento),
                "pedido_envio_fornecedor.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(str(e))

    st.stop()

if not giro_pdf:
    st.info("Envie o PDF de Giro de Estoque para iniciar a análise.")
    st.stop()

with st.spinner("Lendo Giro de Estoque..."):
    texto_giro = extract_text_from_pdf(giro_pdf)
    MESES = extrair_meses_giro_pdf(texto_giro)
    df_giro = parse_giro_estoque(texto_giro, MESES)

if df_giro.empty:
    st.error("Não consegui extrair os dados do Giro de Estoque.")
    st.stop()

if cadastro_google is not None and not cadastro_google.empty:
    df_giro = aplicar_cadastro_dataframe(df_giro, cadastro_google)
else:
    df_giro = aplicar_cadastro(df_giro, cadastro_csv)

df_transito = pd.DataFrame(columns=["codigo", "Saldo em Trânsito/ABERTO"])
if pedidos_pdf:
    with st.spinner("Lendo Pedidos de Compra em Aberto..."):
        df_transito = parse_pedidos_compra_aberto_pdf(pedidos_pdf)

tabela_resumo = montar_tabela_consolidada(
    df_giro,
    df_transito=df_transito,
    dias_estoque_alvo=dias_estoque_alvo,
    meses_alerta_sem_compra=meses_alerta_sem_compra,
)

assinatura_base = (
    tabela_resumo["codigo"].astype(str).str.cat(sep="|")
    + "|fab=" + tabela_resumo.get("Código Fábrica", pd.Series(dtype=str)).astype(str).str.cat(sep="|")
    + "|emb=" + tabela_resumo.get("Embalagem", pd.Series(dtype=str)).astype(str).str.cat(sep="|")
    + f"|dias={dias_estoque_alvo}|alerta={meses_alerta_sem_compra}"
)
if st.session_state.get("assinatura_base_pedido") != assinatura_base:
    st.session_state["pedido_editado"] = inicializar_pedido_editavel(tabela_resumo)
    st.session_state["assinatura_base_pedido"] = assinatura_base

colunas_consolidadas = [
    "codigo", "descricao", "Código Fábrica", "Embalagem",
    *[col_giro("Giro Lojas", mes) for mes in MESES],
    "Média Giro Lojas", "Estoque Lojas",
    *[col_giro("Giro Única", mes) for mes in MESES],
    "Média Giro Única", "Estoque Única",
    *[col_giro("Giro Geral", mes) for mes in MESES],
    "Média Giro Geral", "Estoque Atual Geral", "Estoque Geral", "Saldo em Trânsito/ABERTO", "Estoque Final",
    "Estoque Alvo", "Sugestão Sistema", "Sugestão arredondada", "Data Última Compra", "Preço Última Compra",
]
for col in colunas_consolidadas:
    if col not in tabela_resumo.columns:
        tabela_resumo[col] = 0

render_kpis_gerais(tabela_resumo, st.session_state.get("pedido_editado"))
st.markdown("---")

if pagina == "📦 Giro Consolidado":
    st.markdown('<div class="section-title">📦 Giro Consolidado</div>', unsafe_allow_html=True)
    st.caption(
        "A data da última compra é puxada somente da loja 009. "
        "Quando a data ultrapassa o parâmetro de meses sem compra, aparece o ícone ⚠️ ao lado da data."
    )

    tabela = tabela_resumo[colunas_consolidadas].copy()
    tabela = filtrar_tabela(tabela, ["codigo", "descricao", "Código Fábrica"], "busca_consolidada")
    render_tabela_interativa_colorida(tabela)

    st.download_button(
        "⬇️ Baixar tabela consolidada em CSV",
        gerar_csv(tabela),
        "tabela_consolidada_giro_pedido.csv",
        "text/csv",
    )

    st.markdown("---")
    st.markdown('<div class="section-title">🔎 Drill por produto</div>', unsafe_allow_html=True)
    opcoes_produtos = (
        tabela_resumo["codigo"].astype(str) + " - " + tabela_resumo["descricao"].astype(str)
    ).drop_duplicates().tolist()

    produto_selecionado = st.selectbox(
        "Selecione um item para ver o giro e o saldo em estoque por unidade",
        options=[""] + opcoes_produtos,
        key="produto_drill_consolidada",
    )

    if produto_selecionado:
        codigo_produto = produto_selecionado.split(" - ")[0]
        detalhe = montar_detalhe_produto(df_giro, codigo_produto)
        st.dataframe(
            detalhe.style.format(formatadores_para_tabela(detalhe)),
            use_container_width=True,
            hide_index=True,
            height=360,
            column_config={
                "Cód. Empresa": st.column_config.TextColumn("Cód. Empresa", pinned=True),
                "Unidade": st.column_config.TextColumn("Unidade", pinned=True),
            },
        )

elif pagina == "🛒 Pedido de Compra":
    st.markdown('<div class="section-title">🛒 Pedido de Compra</div>', unsafe_allow_html=True)
    st.caption(
        "Todos os itens aparecem aqui, inclusive os com sugestão zero. "
        "A coluna PEDIDO Final é editável. A coluna Valor Final do Pedido é recalculada por quantidade × preço última compra."
    )

    pedido_base_completo = st.session_state.get("pedido_editado", inicializar_pedido_editavel(tabela_resumo)).copy()
    if "Estoque Geral" not in pedido_base_completo.columns and "Estoque Atual Geral" in pedido_base_completo.columns:
        pedido_base_completo["Estoque Geral"] = pedido_base_completo["Estoque Atual Geral"]
    pedido_base_completo = atualizar_valor_e_origem(pedido_base_completo)

    colunas_sugestao = colunas_pedido_compras(MESES)
    for col in colunas_sugestao:
        if col not in pedido_base_completo.columns:
            pedido_base_completo[col] = 0 if col not in ["codigo", "descricao", "Código Fábrica", "Data Última Compra", "Origem Sugestão"] else ""

    pedido_view = pedido_base_completo[colunas_sugestao].sort_values(["Sugestão Sistema", "descricao"], ascending=[False, True]).copy()
    pedido_view = filtrar_tabela(pedido_view, ["codigo", "descricao", "Código Fábrica"], "busca_sugestao")

    # Recalcula na hora o Valor Final do Pedido quando o usuário altera PEDIDO Final.
    estado_editor = st.session_state.get("editor_pedido_final", {})
    alteracoes_linhas = estado_editor.get("edited_rows", {}) if isinstance(estado_editor, dict) else {}
    if alteracoes_linhas:
        indices_visiveis = list(pedido_view.index)
        for posicao_linha, alteracoes in alteracoes_linhas.items():
            try:
                posicao = int(posicao_linha)
                if posicao < 0 or posicao >= len(indices_visiveis):
                    continue
                indice_real = indices_visiveis[posicao]
                if "PEDIDO Final" in alteracoes:
                    novo_pedido = pd.to_numeric(alteracoes.get("PEDIDO Final"), errors="coerce")
                    if pd.isna(novo_pedido):
                        novo_pedido = 0
                    novo_pedido = int(round(float(novo_pedido)))
                    embalagem_item = int(pd.to_numeric(pedido_view.loc[indice_real].get("Embalagem", 0), errors="coerce") or 0)
                    pedido_validado = ajustar_pedido_para_multiplo_embalagem(novo_pedido, embalagem_item)

                    if novo_pedido > 0 and embalagem_item > 1 and novo_pedido != pedido_validado:
                        descricao_item = str(pedido_view.loc[indice_real].get("descricao", "")).strip()
                        codigo_item = str(pedido_view.loc[indice_real].get("codigo", "")).zfill(5)
                        st.warning(
                            f"Item {codigo_item} - {descricao_item}: a embalagem é com {embalagem_item} unidades. "
                            f"Altere para {pedido_validado}. O sistema ajustou automaticamente para o próximo múltiplo."
                        )

                    pedido_view.loc[indice_real, "PEDIDO Final"] = pedido_validado
                    codigo_alterado = pedido_view.loc[indice_real, "codigo"]
                    mask_base = pedido_base_completo["codigo"].astype(str) == str(codigo_alterado)
                    pedido_base_completo.loc[mask_base, "PEDIDO Final"] = pedido_validado
            except Exception:
                continue

        pedido_base_completo = atualizar_valor_e_origem(pedido_base_completo)
        pedido_view = atualizar_valor_e_origem(pedido_view)
        st.session_state["pedido_editado"] = pedido_base_completo

    pedido_view = pedido_view[colunas_sugestao].copy()
    pedido_style = pedido_view.style.apply(colorir_colunas_pedido, axis=0).apply(estilos_alerta_giro_fora_curva, axis=1).format(formatadores_para_tabela(pedido_view))

    pedido_editado = st.data_editor(
        pedido_style,
        use_container_width=True,
        hide_index=True,
        height=650,
        key="editor_pedido_final",
        disabled=[
            "codigo", "descricao",
            *[col_giro("Giro Geral", mes) for mes in MESES],
            "Média Giro Geral",
            "Estoque Lojas", "Estoque Única", "Estoque Geral",
            "Saldo em Trânsito/ABERTO", "Estoque Final", "Estoque Alvo",
            "Sugestão Sistema", "Sugestão arredondada",
            "Preço Última Compra", "Data Última Compra",
            "Origem Sugestão", "Valor Final do Pedido",
            "Embalagem", "Código Fábrica",
        ],
        column_config={
            "codigo": st.column_config.TextColumn("Código", pinned=True),
            "descricao": st.column_config.TextColumn("Descrição", width="large", pinned=True),
            "Código Fábrica": st.column_config.TextColumn("Código Fábrica", width="medium"),
            "Embalagem": st.column_config.NumberColumn("Embalagem", min_value=0, step=1, format="%d"),
            "Média Giro Lojas": st.column_config.NumberColumn("Média Giro Lojas", format="%.1f"),
            "Estoque Lojas": st.column_config.NumberColumn("Estoque Lojas", format="%.1f"),
            "Média Giro Única": st.column_config.NumberColumn("Média Giro Única", format="%.1f"),
            "Estoque Única": st.column_config.NumberColumn("Estoque Única", format="%.1f"),
            "Média Giro Geral": st.column_config.NumberColumn("Média Giro Geral", format="%.1f"),
            "Estoque Geral": st.column_config.NumberColumn("Estoque Geral", format="%.1f"),
            "Saldo em Trânsito/ABERTO": st.column_config.NumberColumn("Saldo em Trânsito", format="%.1f"),
            "Estoque Final": st.column_config.NumberColumn("Estoque Final", format="%.1f"),
            "Estoque Alvo": st.column_config.NumberColumn("Estoque Alvo", format="%.1f"),
            "Sugestão Sistema": st.column_config.NumberColumn("Sugestão Sistema", format="%d"),
            "Sugestão arredondada": st.column_config.NumberColumn("Sugestão arredondada", format="%d"),
            "Preço Última Compra": st.column_config.NumberColumn("Preço Última Compra", format="R$ %.2f"),
            "PEDIDO Final": st.column_config.NumberColumn("PEDIDO Final", min_value=0, step=1, format="%d"),
            "Valor Final do Pedido": st.column_config.NumberColumn("Valor Final do Pedido", format="R$ %.2f"),
        },
    )

    pedido_editado = pd.DataFrame(pedido_editado)
    pedido_editado, mensagens_validacao = validar_pedidos_por_embalagem(pedido_editado)
    if mensagens_validacao:
        st.warning("Algumas quantidades foram ajustadas para respeitar a embalagem:\n\n" + "\n".join(mensagens_validacao[:10]))
        if len(mensagens_validacao) > 10:
            st.caption(f"Mais {len(mensagens_validacao) - 10} ajuste(s) foram aplicado(s).")

        atualizacoes_validas = pedido_editado[["codigo", "PEDIDO Final"]].copy()
        mapa_validado = atualizacoes_validas.drop_duplicates("codigo", keep="last").set_index("codigo")["PEDIDO Final"]
        pedido_base_completo["PEDIDO Final"] = pedido_base_completo.apply(
            lambda row: int(mapa_validado.loc[row["codigo"]]) if row["codigo"] in mapa_validado.index else int(row["PEDIDO Final"]),
            axis=1,
        )
        st.session_state["pedido_editado"] = atualizar_valor_e_origem(pedido_base_completo)

    pedido_editado = atualizar_valor_e_origem(pedido_editado)
    pedido_editado = pedido_editado[colunas_sugestao]

    valor_editado = totalizar_valor_pedido(pedido_editado)
    st.markdown(
        f"""
        <div style="margin-top: 16px; padding: 18px; border-radius: 14px; background: #f3f6ff; border: 1px solid #d9e2ff;">
            <div style="font-size: 14px; color: #475569;">Valor final do pedido em tela</div>
            <div style="font-size: 30px; font-weight: 700; color: #0f172a;">{format_moeda_br(valor_editado)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("💾 Salvar Pedido", type="primary"):
        base_completa = st.session_state["pedido_editado"].copy()
        if "Estoque Geral" not in base_completa.columns and "Estoque Atual Geral" in base_completa.columns:
            base_completa["Estoque Geral"] = base_completa["Estoque Atual Geral"]
        atualizacoes = pedido_editado[["codigo", "PEDIDO Final", "Embalagem", "descricao"]].copy()
        atualizacoes, mensagens_salvar = validar_pedidos_por_embalagem(atualizacoes)
        if mensagens_salvar:
            st.warning("Antes de salvar, o sistema ajustou quantidades para respeitar a embalagem:\n\n" + "\n".join(mensagens_salvar[:10]))
            if len(mensagens_salvar) > 10:
                st.caption(f"Mais {len(mensagens_salvar) - 10} ajuste(s) foram aplicado(s).")
        atualizacoes["PEDIDO Final"] = pd.to_numeric(atualizacoes["PEDIDO Final"], errors="coerce").fillna(0).round(0).astype(int)
        mapa_qtd = atualizacoes.drop_duplicates("codigo", keep="last").set_index("codigo")["PEDIDO Final"]
        base_completa["PEDIDO Final"] = base_completa.apply(
            lambda row: int(mapa_qtd.loc[row["codigo"]]) if row["codigo"] in mapa_qtd.index else int(row["PEDIDO Final"]),
            axis=1,
        )
        base_completa = atualizar_valor_e_origem(base_completa)
        st.session_state["pedido_editado"] = base_completa
        st.success("Pedido salvo. Vá para a página Exportar Pedido para baixar o Excel e a cópia para fornecedor.")

    try:
        excel_editavel_bytes = gerar_excel_pedido_editavel(pedido_editado)
        st.download_button(
            "⬇️ Baixar pedido editável em Excel",
            excel_editavel_bytes,
            "pedido_editavel.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "⬇️ Baixar pedido editável em CSV",
            gerar_csv(pedido_editado[colunas_pedido_compras(MESES)]),
            "pedido_editavel.csv",
            "text/csv",
        )
    except RuntimeError as e:
        st.error(str(e))

    st.markdown("---")
    st.markdown("### Enviar pedido editável para o Google Drive")
    if google_configurado():
        with st.form("form_exportar_pedido_drive"):
            nome_pedido_drive = st.text_input("Nome do pedido", value=f"Pedido {datetime.now().strftime('%d-%m-%Y')}")
            fornecedor_drive = st.text_input("Fornecedor", value="")
            usuario_drive = st.text_input("Criado por", value="")
            enviar_drive = st.form_submit_button("Criar planilha editável no Drive", type="primary")

        if enviar_drive:
            try:
                pedido_para_drive = st.session_state.get("pedido_editado", pedido_editado).copy()
                pedido_para_drive = atualizar_valor_e_origem(pedido_para_drive)
                pedido_para_drive = pedido_para_drive[colunas_pedido_compras(MESES)]
                resultado_drive = google_criar_planilha_pedido(
                    nome_pedido_drive,
                    fornecedor_drive,
                    pedido_para_drive,
                    criado_por=usuario_drive,
                )
                st.success("Pedido criado no Google Drive.")
                st.link_button("Abrir planilha do pedido", resultado_drive["link"])
            except Exception as e:
                st.error(str(e))
    else:
        st.info(google_mensagem_configuracao())

elif pagina == "📄 Exportações":
    st.markdown('<div class="section-title">📄 Exportações</div>', unsafe_allow_html=True)
    st.caption("O Excel será gerado para importação no Autcom: coluna B = código, coluna F = quantidade, coluna H = valor unitário, sem cabeçalho.")

    pedido_final = st.session_state.get("pedido_editado", inicializar_pedido_editavel(tabela_resumo)).copy()
    pedido_final, mensagens_exportar = validar_pedidos_por_embalagem(pedido_final)
    if mensagens_exportar:
        st.warning("O sistema ajustou quantidades para respeitar a embalagem antes da exportação:\n\n" + "\n".join(mensagens_exportar[:10]))
        if len(mensagens_exportar) > 10:
            st.caption(f"Mais {len(mensagens_exportar) - 10} ajuste(s) foram aplicado(s).")
        st.session_state["pedido_editado"] = pedido_final.copy()
    pedido_final = atualizar_valor_e_origem(pedido_final)
    pedido_final = pedido_final[pedido_final["PEDIDO Final"] > 0].copy().sort_values("descricao")

    valor_final = totalizar_valor_pedido(pedido_final)
    st.markdown(
        f"""
        <div style="margin: 10px 0 18px 0; padding: 18px; border-radius: 14px; background: #f3f6ff; border: 1px solid #d9e2ff;">
            <div style="font-size: 14px; color: #475569;">Valor final do pedido salvo</div>
            <div style="font-size: 30px; font-weight: 700; color: #0f172a;">{format_moeda_br(valor_final)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(
        pedido_final[[
            "codigo", "descricao", "Código Fábrica", "Sugestão Sistema", "Sugestão arredondada", "PEDIDO Final", "Preço Última Compra",
            "Valor Final do Pedido", "Data Última Compra", "Origem Sugestão",
        ]].style.format(formatadores_para_tabela(pedido_final)),
        use_container_width=True,
        hide_index=True,
        height=520,
    )

    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        render_download_card("Excel Autcom", "Arquivo sem cabeçalho: coluna B = código, F = quantidade, H = preço.")
        try:
            excel_bytes = gerar_excel_pedido(pedido_final)
            st.download_button(
                "⬇️ Baixar pedido para importação no Autcom",
                excel_bytes,
                "pedido_importacao_autcom.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        except RuntimeError as e:
            st.error(str(e))

    with col_dl2:
        render_download_card("Cópia para fornecedor", "Lista simples com código de fábrica, descrição e quantidade.")
        st.download_button(
            "⬇️ Baixar cópia CSV para fornecedor",
            gerar_copia_fornecedor_csv(pedido_final),
            "copia_fornecedor.csv",
            "text/csv",
        )
