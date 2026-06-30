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
MESES = ["01/2026", "02/2026", "03/2026", "04/2026"]

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

# =========================================================
# LEITURA DO PDF DE GIRO DE ESTOQUE
# =========================================================

def parse_linha_giro(line):
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

    if len(depois_un) < 9:
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
        "01/2026": br_to_float(depois_un[0]),
        "02/2026": br_to_float(depois_un[1]),
        "03/2026": br_to_float(depois_un[2]),
        "04/2026": br_to_float(depois_un[3]),
        "estoque": br_to_float(depois_un[6]),
        "pr_ult_compra": pr_ult_compra,
        "dt_ult_compra": dt_ult_compra,
        "dt_ult_compra_txt": dt_ult_compra_txt,
        "codigo_empresa": None,
        "loja": None,
    }


def parse_giro_estoque(text):
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

        produto = parse_linha_giro(line)
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

    valores_numericos = [p for p in partes[un_index + 1:] if _eh_numero_br(p)]
    idx_aberto = 5 if indice_aberto is None else int(indice_aberto)

    if len(valores_numericos) > idx_aberto:
        return {"codigo": codigo, "Saldo em Trânsito/ABERTO": br_to_float(valores_numericos[idx_aberto])}

    return {"codigo": codigo, "Saldo em Trânsito/ABERTO": 0.0}


def parse_pedidos_compra_aberto(text):
    registros = []
    indice_aberto = encontrar_indice_aberto_no_cabecalho(text)

    for raw_line in text.splitlines():
        produto = parse_linha_pedido_aberto(raw_line, indice_aberto=indice_aberto)
        if produto:
            registros.append(produto)

    if not registros:
        return pd.DataFrame(columns=["codigo", "Saldo em Trânsito/ABERTO"])

    return pd.DataFrame(registros).groupby("codigo", as_index=False)["Saldo em Trânsito/ABERTO"].sum()


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
                            "Saldo em Trânsito/ABERTO": br_to_float(candidatos[0][1]),
                        })

    except Exception:
        registros = []

    if registros:
        return pd.DataFrame(registros).groupby("codigo", as_index=False)["Saldo em Trânsito/ABERTO"].sum()

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

    agg = {
        "codigo_fabrica": "first",
        "embalagem": "first",
        "01/2026": "sum",
        "02/2026": "sum",
        "03/2026": "sum",
        "04/2026": "sum",
        "estoque": "sum",
    }

    lojas = df_lojas.groupby(["codigo", "descricao"], as_index=False).agg(agg) if not df_lojas.empty else pd.DataFrame(columns=["codigo", "descricao", *agg.keys()])
    unica = df_unica.groupby(["codigo", "descricao"], as_index=False).agg(agg) if not df_unica.empty else pd.DataFrame(columns=["codigo", "descricao", *agg.keys()])

    lojas["Média Giro Lojas"] = lojas[MESES].mean(axis=1).round(1) if not lojas.empty else []
    unica["Média Giro Única"] = unica[MESES].mean(axis=1).round(1) if not unica.empty else []

    lojas = lojas.rename(columns={
        "codigo_fabrica": "Código Fábrica",
        "embalagem": "Embalagem",
        "01/2026": "Giro Lojas Jan/26",
        "02/2026": "Giro Lojas Fev/26",
        "03/2026": "Giro Lojas Mar/26",
        "04/2026": "Giro Lojas Abr/26",
        "estoque": "Estoque Lojas",
    })
    unica = unica.rename(columns={
        "codigo_fabrica": "Código Fábrica Única",
        "embalagem": "Embalagem Única",
        "01/2026": "Giro Única Jan/26",
        "02/2026": "Giro Única Fev/26",
        "03/2026": "Giro Única Mar/26",
        "04/2026": "Giro Única Abr/26",
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

    for mes in ["Jan/26", "Fev/26", "Mar/26", "Abr/26"]:
        for prefixo in ["Giro Lojas", "Giro Única"]:
            col = f"{prefixo} {mes}"
            if col not in resumo.columns:
                resumo[col] = 0
        resumo[f"Giro Geral {mes}"] = resumo[f"Giro Lojas {mes}"] + resumo[f"Giro Única {mes}"]

    resumo["Média Giro Geral"] = resumo[[
        "Giro Geral Jan/26", "Giro Geral Fev/26", "Giro Geral Mar/26", "Giro Geral Abr/26"
    ]].mean(axis=1).round(1)

    for col in ["Estoque Lojas", "Estoque Única", "Média Giro Lojas", "Média Giro Única"]:
        if col not in resumo.columns:
            resumo[col] = 0

    resumo["Estoque Atual Geral"] = resumo["Estoque Lojas"] + resumo["Estoque Única"]
    resumo["Estoque Geral"] = resumo["Estoque Atual Geral"]

    if df_transito is not None and not df_transito.empty:
        resumo = pd.merge(resumo, df_transito, on="codigo", how="left")
    else:
        resumo["Saldo em Trânsito/ABERTO"] = 0

    resumo["Saldo em Trânsito/ABERTO"] = resumo["Saldo em Trânsito/ABERTO"].fillna(0)
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

    detalhe["Média Giro"] = detalhe[MESES].mean(axis=1).round(1)
    detalhe = detalhe.rename(columns={
        "loja": "Unidade",
        "codigo_empresa": "Cód. Empresa",
        "01/2026": "Jan/26",
        "02/2026": "Fev/26",
        "03/2026": "Mar/26",
        "04/2026": "Abr/26",
        "estoque": "Saldo em Estoque",
    })

    return detalhe[[
        "Cód. Empresa", "Unidade", "Jan/26", "Fev/26", "Mar/26", "Abr/26",
        "Média Giro", "Saldo em Estoque",
    ]].sort_values(["Cód. Empresa", "Unidade"])

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
    styled = df.style.apply(colorir_colunas_consolidada, axis=0).format(formatadores_para_tabela(df))
    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        height=height,
        column_config={
            "codigo": st.column_config.TextColumn("Código", width="small", pinned=True),
            "descricao": st.column_config.TextColumn("Descrição", width="large", pinned=True),
            "Código Fábrica": st.column_config.TextColumn("Código Fábrica", width="medium", pinned=True),
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
    ws.freeze_panes = "E2"
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
    colunas_base = [
        "codigo", "descricao", "Código Fábrica", "Embalagem",
        "Média Giro Lojas", "Estoque Lojas",
        "Média Giro Única", "Estoque Única",
        "Média Giro Geral", "Estoque Geral",
        "Saldo em Trânsito/ABERTO", "Estoque Final", "Estoque Alvo",
        "Sugestão Sistema", "Sugestão arredondada", "Preço Última Compra", "Data Última Compra",
    ]

    base = tabela_resumo.copy()
    if "Estoque Geral" not in base.columns and "Estoque Atual Geral" in base.columns:
        base["Estoque Geral"] = base["Estoque Atual Geral"]

    for col in colunas_base:
        if col not in base.columns:
            base[col] = 0 if col not in ["codigo", "descricao", "Código Fábrica", "Data Última Compra"] else ""

    base = base[colunas_base].copy()
    base["PEDIDO Final"] = pd.to_numeric(base["Sugestão arredondada"], errors="coerce").fillna(0).round(0).astype(int)
    base["Origem Sugestão"] = "Sugestão do sistema"
    base["Valor Final do Pedido"] = base["PEDIDO Final"] * pd.to_numeric(base["Preço Última Compra"], errors="coerce").fillna(0)
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
            [data-testid="stSidebar"] .stNumberInput label,
            [data-testid="stSidebar"] .stMarkdown,
            [data-testid="stSidebar"] .stCaption {
                color: #e2e8f0 !important;
            }

            /* Correção de contraste nos parâmetros da sidebar:
               o tema escuro da sidebar não pode deixar textos claros dentro de campos claros. */
            [data-testid="stSidebar"] input,
            [data-testid="stSidebar"] textarea,
            [data-testid="stSidebar"] [contenteditable="true"] {
                color: #0f172a !important;
                background-color: #ffffff !important;
                caret-color: #0f172a !important;
            }

            [data-testid="stSidebar"] div[data-baseweb="input"],
            [data-testid="stSidebar"] div[data-baseweb="base-input"] {
                background-color: #ffffff !important;
                border-color: #cbd5e1 !important;
            }

            [data-testid="stSidebar"] div[data-baseweb="input"] *,
            [data-testid="stSidebar"] div[data-baseweb="base-input"] * {
                color: #0f172a !important;
            }

            [data-testid="stSidebar"] button[aria-label="Increment"],
            [data-testid="stSidebar"] button[aria-label="Decrement"],
            [data-testid="stSidebar"] button[data-testid="stNumberInputStepUp"],
            [data-testid="stSidebar"] button[data-testid="stNumberInputStepDown"] {
                background-color: #f8fafc !important;
                border-color: #cbd5e1 !important;
            }

            [data-testid="stSidebar"] button[aria-label="Increment"] *,
            [data-testid="stSidebar"] button[aria-label="Decrement"] *,
            [data-testid="stSidebar"] button[data-testid="stNumberInputStepUp"] *,
            [data-testid="stSidebar"] button[data-testid="stNumberInputStepDown"] * {
                color: #0f172a !important;
                fill: #0f172a !important;
            }

            [data-testid="stSidebar"] div[role="radiogroup"] label {
                background: rgba(255, 255, 255, 0.06);
                border-radius: 10px;
                padding: 6px 8px;
                margin-bottom: 4px;
            }

            [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
                background: rgba(255, 255, 255, 0.12);
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
# ANÁLISE DE RUPTURA POR MARCA - PDF GIRO GERAL
# =========================================================

_UNIDADES_PDF_GIRO = {"UN", "UND", "UNID", "LT", "L", "GL", "KG", "CX", "PC", "PÇ", "MT", "M", "M2", "M²", "M3", "M³", "RL", "BD", "JG", "PAR", "PT", "SC", "TB", "FR", "KT", "KIT", "DC"}


def _token_numero_br(token):
    return bool(re.fullmatch(r"-?\d{1,3}(?:\.\d{3})*,\d+|-?\d+,\d+|-?\d+(?:\.\d+)?", str(token).strip()))


def _detectar_meses_cabecalho(text):
    """Detecta os quatro meses exibidos na tabela do PDF de Giro Geral."""
    for line in text.splitlines():
        if "COD." in line and "DESCRICAO" in line.upper() and "ESTOQUE" in line.upper():
            meses = re.findall(r"\d{2}/\d{4}", line)
            if len(meses) >= 4:
                return meses[:4]
    meses = re.findall(r"\d{2}/\d{4}", text[:5000])
    meses_unicos = []
    for mes in meses:
        if mes not in meses_unicos:
            meses_unicos.append(mes)
    return meses_unicos[:4] if len(meses_unicos) >= 4 else ["Mês 1", "Mês 2", "Mês 3", "Mês 4"]


def parse_linha_giro_marca(line, meses_detectados):
    """Parser flexível para linha de produto do relatório de giro, preservando marca e unidade."""
    if not re.match(r"^\d{5}\s+", line.strip()):
        return None

    partes = line.split()
    if len(partes) < 12:
        return None

    codigo = partes[0].zfill(5)
    un_index = None

    # A unidade é o token imediatamente antes de uma sequência longa de números da tabela.
    for i in range(2, len(partes) - 8):
        prox = partes[i + 1:i + 8]
        if partes[i].upper() in _UNIDADES_PDF_GIRO and len(prox) >= 7 and all(_token_numero_br(x) for x in prox[:7]):
            un_index = i
            break

    if un_index is None:
        return None

    depois = partes[un_index + 1:]
    if len(depois) < 7:
        return None

    descricao = " ".join(partes[1:un_index]).strip()
    unidade = partes[un_index].upper()

    valores_meses = [br_to_float(v) for v in depois[:4]]
    media_pdf = br_to_float(depois[4]) if len(depois) > 4 else 0.0
    previ30_pdf = br_to_float(depois[5]) if len(depois) > 5 else 0.0
    estoque = br_to_float(depois[6]) if len(depois) > 6 else 0.0

    row = {
        "codigo": codigo,
        "descricao": descricao,
        "unidade": unidade,
        "media_pdf": media_pdf,
        "previ30_pdf": previ30_pdf,
        "estoque": estoque,
    }
    for idx, mes in enumerate(meses_detectados[:4]):
        row[mes] = valores_meses[idx] if idx < len(valores_meses) else 0.0
    return row


@st.cache_data(show_spinner="Analisando PDF por marca...")
def parse_giro_estoque_por_marca_pdf_bytes(pdf_bytes):
    texto = extract_text_from_pdf(BytesIO(pdf_bytes))
    meses_detectados = _detectar_meses_cabecalho(texto)
    registros = []
    empresa_atual = None
    marca_codigo = "SEM CÓDIGO"
    marca_nome = "SEM MARCA"

    for raw_line in texto.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        empresa_match = re.search(r"EMPRESA\s*:\s*(\d{3})\s*-", line)
        if empresa_match:
            empresa_atual = empresa_match.group(1)
            continue

        marca_match = re.search(r"MARCA:\s*(\d+)\s*-\s*(.*)$", line)
        if marca_match:
            marca_codigo = marca_match.group(1).strip()
            marca_nome = marca_match.group(2).strip() or "."
            continue

        if empresa_atual not in LOJAS_MAP:
            continue

        produto = parse_linha_giro_marca(line, meses_detectados)
        if produto:
            produto["codigo_empresa"] = empresa_atual
            produto["loja"] = LOJAS_MAP.get(empresa_atual, empresa_atual)
            produto["tipo_unidade"] = "ÚNICA" if empresa_atual == CODIGO_UNICA else "LOJAS DAUTO"
            produto["marca_codigo"] = marca_codigo
            produto["marca"] = marca_nome
            registros.append(produto)

    return pd.DataFrame(registros), meses_detectados


def classificar_risco_ruptura(cobertura_dias, estoque_geral, media_mensal):
    if media_mensal <= 0:
        return "SEM GIRO"
    if estoque_geral <= 0 or cobertura_dias <= 7:
        return "CRÍTICO"
    if cobertura_dias <= 15:
        return "ALTO"
    if cobertura_dias <= 30:
        return "ATENÇÃO"
    return "OK"


def montar_analise_ruptura_por_marca(df_marca, meses_detectados):
    if df_marca.empty:
        return pd.DataFrame(), pd.DataFrame()

    meses_validos = [m for m in meses_detectados[:4] if m in df_marca.columns]
    chaves = ["marca_codigo", "marca", "codigo", "descricao", "unidade"]

    agg_dict = {m: "sum" for m in meses_validos}
    agg_dict["estoque"] = "sum"
    itens = df_marca.groupby(chaves, as_index=False).agg(agg_dict)

    itens["Giro Total 4 meses"] = itens[meses_validos].sum(axis=1) if meses_validos else 0
    itens["Média Mensal Geral"] = itens[meses_validos].mean(axis=1) if meses_validos else 0
    itens["Previsão 30 dias"] = itens["Média Mensal Geral"]
    itens["Estoque Geral"] = itens["estoque"]
    itens["Cobertura em dias"] = itens.apply(
        lambda r: 9999.0 if r["Média Mensal Geral"] <= 0 else (float(r["Estoque Geral"]) / (float(r["Média Mensal Geral"]) / 30.0)),
        axis=1,
    )
    itens["Falta p/ 30 dias"] = (itens["Previsão 30 dias"] - itens["Estoque Geral"]).clip(lower=0)
    itens["Status Ruptura"] = itens.apply(
        lambda r: classificar_risco_ruptura(r["Cobertura em dias"], r["Estoque Geral"], r["Média Mensal Geral"]),
        axis=1,
    )
    peso = {"CRÍTICO": 100, "ALTO": 60, "ATENÇÃO": 25, "OK": 0, "SEM GIRO": -1}
    itens["Peso Risco"] = itens["Status Ruptura"].map(peso).fillna(0)
    itens["Marca"] = itens["marca_codigo"].astype(str) + " - " + itens["marca"].astype(str)

    risco = itens[itens["Status Ruptura"].isin(["CRÍTICO", "ALTO", "ATENÇÃO"])].copy()
    resumo = itens.groupby(["marca_codigo", "marca", "Marca"], as_index=False).agg(
        **{
            "Itens analisados": ("codigo", "nunique"),
            "Itens em risco": ("Status Ruptura", lambda s: int(s.isin(["CRÍTICO", "ALTO", "ATENÇÃO"]).sum())),
            "Críticos": ("Status Ruptura", lambda s: int((s == "CRÍTICO").sum())),
            "Alto risco": ("Status Ruptura", lambda s: int((s == "ALTO").sum())),
            "Atenção": ("Status Ruptura", lambda s: int((s == "ATENÇÃO").sum())),
            "Giro médio mensal": ("Média Mensal Geral", "sum"),
            "Estoque geral": ("Estoque Geral", "sum"),
            "Falta p/ 30 dias": ("Falta p/ 30 dias", "sum"),
            "Peso Total": ("Peso Risco", "sum"),
        }
    )
    resumo["% itens em risco"] = resumo.apply(lambda r: 0 if r["Itens analisados"] == 0 else r["Itens em risco"] / r["Itens analisados"] * 100, axis=1)
    resumo = resumo.sort_values(["Peso Total", "Críticos", "Alto risco", "Itens em risco", "Falta p/ 30 dias"], ascending=False)

    itens = itens.sort_values(["Peso Risco", "Falta p/ 30 dias", "Média Mensal Geral"], ascending=False)
    return resumo, itens


def _cor_status_ruptura(valor):
    cores = {
        "CRÍTICO": "background-color: #fee2e2; color: #991b1b; font-weight: 800;",
        "ALTO": "background-color: #ffedd5; color: #9a3412; font-weight: 800;",
        "ATENÇÃO": "background-color: #fef9c3; color: #854d0e; font-weight: 800;",
        "OK": "background-color: #dcfce7; color: #166534; font-weight: 800;",
        "SEM GIRO": "background-color: #e5e7eb; color: #374151; font-weight: 800;",
    }
    return cores.get(str(valor), "")


def formatadores_ruptura(df):
    fmt = {}
    for c in df.columns:
        if c in ["Cobertura em dias", "% itens em risco"]:
            fmt[c] = lambda v: "—" if float(v) >= 9999 else format_num_br(v, 1)
        elif c in ["Giro médio mensal", "Estoque geral", "Falta p/ 30 dias", "Média Mensal Geral", "Previsão 30 dias", "Estoque Geral", "Giro Total 4 meses"]:
            fmt[c] = lambda v: format_num_br(v, 1)
        elif c in ["Itens analisados", "Itens em risco", "Críticos", "Alto risco", "Atenção"]:
            fmt[c] = lambda v: format_int_br(v)
    return fmt


def gerar_csv_ruptura(df):
    return df.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")

# =========================================================
# APP STREAMLIT
# =========================================================

aplicar_css_global()
render_header()

st.sidebar.markdown("### 📊 Análise de Giro")
pagina = st.sidebar.radio(
    "Navegação",
    ["📦 Giro Consolidado", "🏷️ Ruptura por Marca", "🛒 Pedido de Compra", "📄 Exportações", "⚙️ Tratamento Final"],
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
    df_giro = parse_giro_estoque(texto_giro)

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
    "Giro Lojas Jan/26", "Giro Lojas Fev/26", "Giro Lojas Mar/26", "Giro Lojas Abr/26",
    "Média Giro Lojas", "Estoque Lojas",
    "Giro Única Jan/26", "Giro Única Fev/26", "Giro Única Mar/26", "Giro Única Abr/26",
    "Média Giro Única", "Estoque Única",
    "Giro Geral Jan/26", "Giro Geral Fev/26", "Giro Geral Mar/26", "Giro Geral Abr/26",
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


elif pagina == "🏷️ Ruptura por Marca":
    st.markdown('<div class="section-title">🏷️ Análise de Ruptura por Marca</div>', unsafe_allow_html=True)
    st.caption(
        "Esta visão soma o estoque e o giro de todas as lojas Dauto (004, 006, 012, 013, 014, 015, 016, 022 e 024) "
        "com a Única (009), calcula a cobertura em dias e organiza as marcas pelo maior risco de ruptura."
    )

    try:
        giro_pdf.seek(0)
    except Exception:
        pass

    df_marca, meses_marca = parse_giro_estoque_por_marca_pdf_bytes(giro_pdf.getvalue())
    if df_marca.empty:
        st.error("Não consegui extrair a análise por marca desse PDF. Confira se o relatório está no layout de Giro de Estoque por Marca.")
        st.stop()

    resumo_marca, itens_marca = montar_analise_ruptura_por_marca(df_marca, meses_marca)
    itens_risco = itens_marca[itens_marca["Status Ruptura"].isin(["CRÍTICO", "ALTO", "ATENÇÃO"])]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_metric("Marcas com risco", format_int_br((resumo_marca["Itens em risco"] > 0).sum()), "Com pelo menos 1 item em risco")
    with c2:
        render_metric("Itens críticos", format_int_br((itens_marca["Status Ruptura"] == "CRÍTICO").sum()), "Estoque zerado ou até 7 dias")
    with c3:
        render_metric("Itens alto risco", format_int_br((itens_marca["Status Ruptura"] == "ALTO").sum()), "Cobertura até 15 dias")
    with c4:
        render_metric("Falta p/ 30 dias", format_num_br(itens_marca["Falta p/ 30 dias"].sum(), 1), "Necessidade teórica geral")

    st.markdown("#### Ranking de marcas por risco de ruptura")
    cols_resumo = ["Marca", "Itens analisados", "Itens em risco", "Críticos", "Alto risco", "Atenção", "% itens em risco", "Giro médio mensal", "Estoque geral", "Falta p/ 30 dias"]
    resumo_view = resumo_marca[cols_resumo].copy()
    resumo_view = resumo_view[resumo_view["Itens em risco"] > 0]

    if resumo_view.empty:
        st.success("Nenhuma marca apresentou item com risco de ruptura pelos critérios atuais.")
    else:
        st.dataframe(
            resumo_view.style.format(formatadores_ruptura(resumo_view)).background_gradient(subset=["Itens em risco", "Críticos", "Alto risco", "Falta p/ 30 dias"]),
            use_container_width=True,
            hide_index=True,
            height=430,
        )

    st.download_button(
        "⬇️ Baixar ranking de marcas em CSV",
        gerar_csv_ruptura(resumo_view if not resumo_view.empty else resumo_marca[cols_resumo]),
        "ranking_ruptura_por_marca.csv",
        "text/csv",
    )

    st.markdown("---")
    st.markdown("#### Drill da marca")
    marcas_opcoes = resumo_marca["Marca"].tolist()
    marca_selecionada = st.selectbox("Selecione uma marca para ver os produtos", marcas_opcoes, key="drill_marca_ruptura")

    somente_risco = st.toggle("Mostrar somente itens em risco", value=True)
    detalhe_marca = itens_marca[itens_marca["Marca"] == marca_selecionada].copy()
    if somente_risco:
        detalhe_marca = detalhe_marca[detalhe_marca["Status Ruptura"].isin(["CRÍTICO", "ALTO", "ATENÇÃO"])]

    cols_itens = ["Status Ruptura", "codigo", "descricao", "unidade"] + [m for m in meses_marca[:4] if m in detalhe_marca.columns] + [
        "Média Mensal Geral", "Estoque Geral", "Cobertura em dias", "Falta p/ 30 dias"
    ]
    for col in cols_itens:
        if col not in detalhe_marca.columns:
            detalhe_marca[col] = 0

    st.dataframe(
        detalhe_marca[cols_itens].style.format(formatadores_ruptura(detalhe_marca)).map(_cor_status_ruptura, subset=["Status Ruptura"]),
        use_container_width=True,
        hide_index=True,
        height=520,
        column_config={
            "Status Ruptura": st.column_config.TextColumn("Status", pinned=True),
            "codigo": st.column_config.TextColumn("Código", pinned=True),
            "descricao": st.column_config.TextColumn("Descrição", pinned=True),
        },
    )

    st.download_button(
        "⬇️ Baixar produtos da marca em CSV",
        gerar_csv_ruptura(detalhe_marca[cols_itens]),
        "drill_produtos_marca_ruptura.csv",
        "text/csv",
    )

    st.markdown("---")
    with st.expander("Critérios usados na classificação"):
        st.markdown(
            "**CRÍTICO:** estoque geral zerado/negativo ou cobertura até 7 dias.  "
            "**ALTO:** cobertura acima de 7 e até 15 dias.  "
            "**ATENÇÃO:** cobertura acima de 15 e até 30 dias.  "
            "**OK:** cobertura acima de 30 dias.  "
            "Itens sem giro ficam fora do risco de ruptura, pois não há consumo médio para projetar falta."
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

    colunas_sugestao = [
        "codigo", "descricao", "Código Fábrica", "Embalagem",
        "Média Giro Lojas", "Estoque Lojas",
        "Média Giro Única", "Estoque Única",
        "Média Giro Geral", "Estoque Geral", "Saldo em Trânsito/ABERTO", "Estoque Final",
        "Estoque Alvo", "Sugestão Sistema", "Sugestão arredondada", "Preço Última Compra", "Data Última Compra",
        "PEDIDO Final", "Origem Sugestão", "Valor Final do Pedido",
    ]
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
    pedido_style = pedido_view.style.apply(colorir_colunas_pedido, axis=0).format(formatadores_para_tabela(pedido_view))

    pedido_editado = st.data_editor(
        pedido_style,
        use_container_width=True,
        hide_index=True,
        height=650,
        key="editor_pedido_final",
        disabled=[
            "codigo", "descricao", "Código Fábrica", "Embalagem",
            "Média Giro Lojas", "Estoque Lojas", "Média Giro Única", "Estoque Única",
            "Média Giro Geral", "Estoque Geral", "Saldo em Trânsito/ABERTO", "Estoque Final",
            "Estoque Alvo", "Sugestão Sistema", "Sugestão arredondada", "Preço Última Compra", "Data Última Compra",
            "Origem Sugestão", "Valor Final do Pedido",
        ],
        column_config={
            "codigo": st.column_config.TextColumn("Código", pinned=True),
            "descricao": st.column_config.TextColumn("Descrição", width="large", pinned=True),
            "Código Fábrica": st.column_config.TextColumn("Código Fábrica", pinned=True),
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
