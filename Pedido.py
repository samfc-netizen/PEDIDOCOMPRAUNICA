import re
import math
from io import BytesIO
from datetime import datetime, date

import pdfplumber
import pandas as pd
import streamlit as st

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

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
        # mantém ponto como decimal quando houver até 2 casas decimais
        partes = txt.split(".")
        if len(partes) == 2 and len(partes[1]) <= 2:
            pass
        else:
            # caso venha como milhar: 1.234 ou 1.234.567
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

def parse_linha_giro(line, meses_ref=None):
    """
    Layout observado:
    COD DESCRICAO CÓD.FABRICA UN 01/2026 02/2026 03/2026 04/2026
    MEDIA DIAS DU ESTOQUE SUGESTAO PR.ULT.COMP DT.ULT.COMP PR.VENDA % LUCRO
    """
    if not re.match(r"^\d{5}\s+", line):
        return None

    partes = line.split()
    codigo = partes[0].zfill(5)

    try:
        un_index = partes.index("UN")
    except ValueError:
        return None

    antes_un = partes[1:un_index]
    depois_un = partes[un_index + 1:]

    meses_ref = meses_ref or MESES_PADRAO
    qtd_meses = len(meses_ref)

    if len(depois_un) < qtd_meses + 5:
        return None

    # Código de fábrica normalmente fica imediatamente antes do UN.
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
        if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", token):
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
        # Posição padrão do PR.ULT.COMP quando não há DT.ULT.COMP na linha.
        pr_ult_compra = br_to_float(depois_un[8]) if len(depois_un) > 8 else 0.0

    return {
        "codigo": codigo,
        "descricao": " ".join(descricao_tokens).strip(),
        "codigo_fabrica": codigo_fabrica_extraido,
        **{mes: br_to_float(depois_un[i]) for i, mes in enumerate(meses_ref)},
        "estoque": br_to_float(depois_un[qtd_meses + 2]) if len(depois_un) > qtd_meses + 2 else 0,
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
    Parser preferencial para o PDF de Pedidos em Aberto usando coordenadas.

    Correção aplicada:
    - Só aceita como cabeçalho a linha real que contém QTDE, BAIXADO, ABERTO e VR.UNIT.
    - Ignora linhas de status como "BAIXADO/ABERTO: ABERTO TOTALMENTE", que antes mudavam a posição da coluna.
    - Nas linhas de produto, lê visualmente a coluna ABERTO, não a coluna VR.UNIT.
    """
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
                    linha_upper = " ".join(textos_upper)

                    # Cabeçalho verdadeiro da tabela. Não confundir com linhas "BAIXADO/ABERTO".
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

                    if aberto_x is None:
                        continue

                    if not textos:
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
                    candidatos = []

                    for w in linha_words[1:]:
                        txt = str(w.get("text", "")).strip()
                        if not _eh_numero_br(txt):
                            continue
                        cx = (float(w["x0"]) + float(w["x1"])) / 2
                        distancia = abs(cx - aberto_x)
                        candidatos.append((distancia, txt))

                    # A coluna ABERTO fica muito próxima do centro do cabeçalho ABERTO.
                    # VR.UNIT fica mais à direita e não deve entrar no filtro.
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

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    texto = extract_text_from_pdf(uploaded_file)
    return parse_pedidos_compra_aberto(texto)


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


def aplicar_cadastro(df_giro, cadastro_csv):
    cadastro = ler_cadastro_produtos_csv(cadastro_csv)
    if cadastro.empty:
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
    return df.drop(columns=["descricao_cadastro", "codigo_fabrica_cadastro"], errors="ignore")

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

# =========================================================
# APP STREAMLIT
# =========================================================

aplicar_css_global()
render_header()

st.sidebar.markdown("### 📊 Análise de Giro")
pagina = st.sidebar.radio(
    "Navegação",
    ["📦 Giro Consolidado", "🛒 Pedido de Compra", "📄 Exportações", "🏷️ Ruptura por Marca", "⚙️ Tratamento Final"],
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


st.markdown('<div class="section-title">Upload dos arquivos</div>', unsafe_allow_html=True)
st.caption("Envie o PDF de Giro para iniciar. Os demais arquivos enriquecem a análise e o pedido final.")
col_upload_1, col_upload_2, col_upload_3 = st.columns(3)

with col_upload_1:
    giro_pdf = st.file_uploader("PDF - Giro de Estoque", type=["pdf"], key="upload_giro_pdf")
    render_upload_status("📄 Giro de Estoque", giro_pdf, obrigatorio=True)
with col_upload_2:
    pedidos_pdf = st.file_uploader("PDF - Pedidos em Aberto", type=["pdf"], key="upload_pedidos_pdf")
    render_upload_status("📄 Pedidos em Aberto", pedidos_pdf)
with col_upload_3:
    cadastro_csv = st.file_uploader("CSV - Cadastro de Produtos", type=["csv"], key="upload_cadastro_csv")
    render_upload_status("📄 Cadastro de Produtos", cadastro_csv)

if pagina == "⚙️ Tratamento Final":
    st.markdown('<div class="section-title">⚙️ Tratamento de Pedido Final</div>', unsafe_allow_html=True)
    st.caption(
        "Envie a planilha final editável. O sistema vai gerar um Excel para importação no Autcom: "
        "coluna B = zx, coluna F = PEDIDO Final e coluna H = Preço Última Compra."
    )

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
        st.download_button(
            "⬇️ Baixar pedido tratado para importação no Autcom",
            excel_tratamento,
            "pedido_tratado_importacao_autcom.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
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
