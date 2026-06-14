import os
import re
import hashlib
from collections import defaultdict

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup

# ==============================
# CONFIGURAÇÕES
# ==============================
ARQUIVO_PLANILHA_PADRAO = "extracao_bruta_pedigree1.xlsx"
ARQUIVO_HTML_BASE = "modelopadraohipotetico_espelhado.html"

st.set_page_config(page_title="Cruzamento Hipotético", layout="wide")
st.title("🐴 Sistema de Pedigree e Cruzamento Hipotético")
st.caption("Novo modelo: preenchimento direto pelos placeholders Item_xx_TextoCompleto e Item_xx_TextoCompleto1")

# ==============================
# GERAÇÕES POR ITEM
# ==============================
ITEMS_GERACAO_1 = {19}
ITEMS_GERACAO_2 = {11, 28}
ITEMS_GERACAO_3 = {7, 15, 24, 32}
ITEMS_GERACAO_4 = {5, 9, 13, 17, 22, 26, 30, 34}
ITEMS_GERACAO_5 = {4, 6, 8, 10, 12, 14, 16, 18, 21, 23, 25, 27, 29, 31, 33, 35}
ITEMS_GERACAO_6 = set(range(36, 52)) | set(range(84, 100))


def geracao_por_item(numero_item: int) -> str:
    if numero_item in ITEMS_GERACAO_1:
        return "1"
    if numero_item in ITEMS_GERACAO_2:
        return "2"
    if numero_item in ITEMS_GERACAO_3:
        return "3"
    if numero_item in ITEMS_GERACAO_4:
        return "4"
    if numero_item in ITEMS_GERACAO_5:
        return "5"
    if numero_item in ITEMS_GERACAO_6:
        return "6"
    if 52 <= numero_item <= 83 or 100 <= numero_item <= 131:
        return "7"
    return ""

# ==============================
# LIMPEZA / NORMALIZAÇÃO
# ==============================

def limpar(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    if valor.lower() in ["nan", "none", "xxx", "não informado", "nao informado", "-"]:
        return ""
    valor = valor.replace("\n", " ").replace("\r", " ")
    valor = re.sub(r"https?://\S+", "", valor)
    valor = re.sub(r"www\.\S+", "", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def normalizar_sbb(sbb):
    sbb = limpar(sbb).upper().replace(" ", "")
    if sbb in ["", "XXX", "-", "NÃOINFORMADO", "NAOINFORMADO", "NAN", "NONE"]:
        return ""
    return sbb


def cor_por_sbb(sbb):
    h = hashlib.md5(sbb.encode("utf-8")).hexdigest()
    return "#" + h[:6]


def extrair_sbb_do_texto(texto):
    texto = limpar(texto)
    # SBBs típicos: B123456, *000532, CHD40125, AD006341 etc.
    m = re.search(r"(\*[0-9]{3,}|[A-Z]{1,4}[0-9]{3,})", texto.upper())
    return m.group(1) if m else ""


def extrair_nome_do_texto(texto, sbb=""):
    texto = limpar(texto)
    if not texto:
        return ""

    # Remove pedaços administrativos do animal principal.
    texto = re.sub(r"\s+-\s*RP:.*$", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\s+RP:.*$", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\s+Nascimento:.*$", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\s+NMGC:.*$", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\s+Pelagem:.*$", "", texto, flags=re.IGNORECASE)

    if sbb:
        # Nome antes de "- SBB" ou antes de " SBB".
        texto = re.sub(rf"\s*-\s*{re.escape(sbb)}.*$", "", texto, flags=re.IGNORECASE)
        texto = re.sub(rf"\s+{re.escape(sbb)}.*$", "", texto, flags=re.IGNORECASE)

    # Se ainda sobrou barra, considera apenas o primeiro animal.
    if " / " in texto:
        texto = texto.split(" / ")[0].strip()

    return limpar(texto).strip(" -/")


def extrair_pelagem_do_texto(texto):
    texto = limpar(texto)
    if not texto:
        return ""

    m = re.search(r"Pelagem:\s*([^/\-]+)$", texto, flags=re.IGNORECASE)
    if m:
        return limpar(m.group(1))

    # Em padrão NOME - SBB / PELAGEM, pega somente o trecho após a última barra.
    # Evita aceitar nomes de outro animal como pelagem quando vier "SBB / OUTRO NOME".
    if " / " in texto:
        cand = limpar(texto.split(" / ")[-1])
        # Pelagens normalmente não têm código SBB e não têm nomes típicos com prefixos/cabanhas.
        if not extrair_sbb_do_texto(cand) and len(cand.split()) <= 5:
            return cand

    return ""


def texto_animal(row, numero_item: int, sufixo: str = ""):
    """Monta texto padronizado usando Nome/SBB/Pelagem e TextoCompleto como apoio."""
    prefixo = f"Item_{numero_item:02d}"

    col_nome = f"{prefixo}_Nome{sufixo}"
    col_sbb = f"{prefixo}_SBB{sufixo}"
    col_pelagem = f"{prefixo}_Pelagem{sufixo}"
    col_texto = f"{prefixo}_TextoCompleto{sufixo}"

    # Algumas planilhas antigas tinham Item_83_TextoCompleto sem sufixo na aba animal2.
    texto_completo = limpar(row.get(col_texto, ""))
    if not texto_completo and sufixo == "1":
        texto_completo = limpar(row.get(f"{prefixo}_TextoCompleto", ""))

    nome = limpar(row.get(col_nome, ""))
    sbb = limpar(row.get(col_sbb, ""))
    pelagem = limpar(row.get(col_pelagem, ""))

    if not sbb:
        sbb = extrair_sbb_do_texto(texto_completo)

    if not nome:
        nome = extrair_nome_do_texto(texto_completo, sbb)

    if not pelagem:
        pelagem = extrair_pelagem_do_texto(texto_completo)

    # Segurança contra nomes que vieram grudados com código/pelagem.
    if sbb and sbb.upper() in nome.upper():
        nome = extrair_nome_do_texto(nome, sbb)

    partes = []
    if nome:
        partes.append(nome)
    if sbb:
        partes.append(f"- {sbb}")
    if pelagem:
        partes.append(f"/ {pelagem}")

    texto_final = " ".join(partes).strip()
    return texto_final, normalizar_sbb(sbb), nome, pelagem


def nome_para_select(row, sufixo=""):
    texto, sbb, nome, _ = texto_animal(row, 19, sufixo)
    if nome and sbb:
        return f"{nome} | {sbb}"
    if texto:
        return texto
    sbb_pesq = limpar(row.get(f"SBB Pesquisado{sufixo}", ""))
    return sbb_pesq or "Animal sem identificação"


def localizar_por_select(df, escolha, sufixo=""):
    sbb = escolha.split("|")[-1].strip() if "|" in escolha else ""
    if sbb:
        col = f"SBB Pesquisado{sufixo}"
        if col in df.columns:
            achou = df[df[col].astype(str).str.strip().str.upper() == sbb.upper()]
            if not achou.empty:
                return achou.iloc[0]

    # fallback pelo texto do select
    for _, row in df.iterrows():
        if nome_para_select(row, sufixo) == escolha:
            return row
    return None

# ==============================
# HTML / PREENCHIMENTO
# ==============================
PLACEHOLDER_RE = re.compile(r"^Item_(\d{2,3})_TextoCompleto(1?)$")


def coletar_placeholders_do_html(soup):
    placeholders = []
    for td in soup.find_all("td"):
        texto = td.get_text(" ", strip=True)
        if PLACEHOLDER_RE.match(texto):
            placeholders.append(texto)
    return placeholders


def montar_mapa_valores(soup, row1, row2):
    mapa = {}
    ocorrencias = defaultdict(list)

    for placeholder in coletar_placeholders_do_html(soup):
        m = PLACEHOLDER_RE.match(placeholder)
        if not m:
            continue

        numero = int(m.group(1))
        sufixo = m.group(2)  # "" = animal1; "1" = animal2
        row = row2 if sufixo == "1" else row1

        texto_final, sbb_norm, nome, pelagem = texto_animal(row, numero, sufixo)
        mapa[placeholder] = {
            "texto": texto_final,
            "sbb": sbb_norm,
            "nome": nome,
            "pelagem": pelagem,
            "geracao": geracao_por_item(numero),
            "item": numero,
        }

        if sbb_norm:
            ocorrencias[sbb_norm].append(mapa[placeholder])

    repetidos = {
        sbb: lista
        for sbb, lista in ocorrencias.items()
        if len(lista) > 1
    }

    return mapa, repetidos


def aplicar_valores_no_html(soup, mapa, repetidos):
    for td in soup.find_all("td"):
        texto_original = td.get_text(" ", strip=True)
        if texto_original not in mapa:
            continue

        info = mapa[texto_original]
        td.clear()

        span = soup.new_tag("span")
        span["style"] = (
            "display:block; font-size:8px; line-height:1.05em; "
            "font-family:Calibri, Arial, sans-serif; font-weight:700;"
        )

        strong = soup.new_tag("strong")
        strong.string = info["texto"]
        span.append(strong)
        td.append(span)

        sbb = info["sbb"]
        if sbb and sbb in repetidos:
            cor = cor_por_sbb(sbb)
            estilo_atual = td.get("style", "")
            td["style"] = f"{estilo_atual}; border-left: 8px solid {cor} !important;"

    style_tag = soup.new_tag("style")
    style_tag.string = """
    body, table, td, th, span, p, div {
        font-family: Calibri, Arial, sans-serif !important;
    }
    td {
        overflow-wrap: anywhere !important;
        vertical-align: middle !important;
    }
    .relatorio-duplicacoes {
        margin-top: 20px;
        font-family: Arial, sans-serif;
        font-size: 13px;
        max-width: 1000px;
    }
    .relatorio-duplicacoes table {
        border-collapse: collapse;
        width: 100%;
    }
    .relatorio-duplicacoes th, .relatorio-duplicacoes td {
        border: 1px solid #999;
        padding: 5px;
        font-size: 12px !important;
        text-align: left;
    }
    .relatorio-duplicacoes th {
        background: #efefef;
    }
    """
    if soup.head:
        soup.head.append(style_tag)


def montar_relatorio_html(soup, repetidos):
    div = soup.new_tag("div")
    div["class"] = "relatorio-duplicacoes"

    h3 = soup.new_tag("h3")
    h3.string = "Relatório de animais repetidos no pedigree"
    div.append(h3)

    if not repetidos:
        p = soup.new_tag("p")
        p.string = "Nenhum animal repetido encontrado."
        div.append(p)
        soup.body.append(div)
        return

    tabela = soup.new_tag("table")
    cab = soup.new_tag("tr")
    for titulo in ["Animal", "SBB", "Quantidade", "Gerações"]:
        th = soup.new_tag("th")
        th.string = titulo
        cab.append(th)
    tabela.append(cab)

    for sbb, lista in sorted(repetidos.items(), key=lambda x: len(x[1]), reverse=True):
        nome = next((x["nome"] for x in lista if x.get("nome")), "")
        geracoes = [x["geracao"] for x in lista if x.get("geracao")]

        tr = soup.new_tag("tr")
        cor = cor_por_sbb(sbb)

        td_nome = soup.new_tag("td")
        td_nome["style"] = f"border-left: 8px solid {cor};"
        td_nome.string = nome
        tr.append(td_nome)

        td_sbb = soup.new_tag("td")
        td_sbb.string = sbb
        tr.append(td_sbb)

        td_qtd = soup.new_tag("td")
        td_qtd.string = str(len(lista))
        tr.append(td_qtd)

        td_ger = soup.new_tag("td")
        td_ger.string = "x".join(geracoes)
        tr.append(td_ger)

        tabela.append(tr)

    div.append(tabela)
    soup.body.append(div)


def gerar_html(row1, row2, caminho_html_base):
    with open(caminho_html_base, "r", encoding="utf-8", errors="ignore") as f:
        soup = BeautifulSoup(f, "html.parser")

    mapa, repetidos = montar_mapa_valores(soup, row1, row2)
    aplicar_valores_no_html(soup, mapa, repetidos)
    montar_relatorio_html(soup, repetidos)

    return str(soup), repetidos

# ==============================
# INTERFACE
# ==============================

st.markdown("Carregue a planilha e selecione o primeiro animal na aba **animal1** e o segundo na aba **animal2**.")

arquivo_planilha = st.file_uploader("Planilha padrão (.xlsx)", type=["xlsx"])
arquivo_html = st.file_uploader("Modelo HTML espelhado", type=["html", "htm"])

if arquivo_planilha is not None:
    excel = pd.ExcelFile(arquivo_planilha)
else:
    if not os.path.exists(ARQUIVO_PLANILHA_PADRAO):
        st.warning(f"Carregue a planilha ou coloque '{ARQUIVO_PLANILHA_PADRAO}' na pasta do app.")
        st.stop()
    excel = pd.ExcelFile(ARQUIVO_PLANILHA_PADRAO)

if "animal1" not in excel.sheet_names or "animal2" not in excel.sheet_names:
    st.error("A planilha precisa conter as abas 'animal1' e 'animal2'.")
    st.stop()

animal1_df = pd.read_excel(excel, sheet_name="animal1", dtype=str)
animal2_df = pd.read_excel(excel, sheet_name="animal2", dtype=str)

if arquivo_html is not None:
    caminho_html_temp = "_modelo_upload_temp.html"
    with open(caminho_html_temp, "wb") as f:
        f.write(arquivo_html.getbuffer())
    caminho_html_base = caminho_html_temp
else:
    if not os.path.exists(ARQUIVO_HTML_BASE):
        st.warning(f"Carregue o HTML ou coloque '{ARQUIVO_HTML_BASE}' na pasta do app.")
        st.stop()
    caminho_html_base = ARQUIVO_HTML_BASE

opcoes1 = [nome_para_select(row, "") for _, row in animal1_df.iterrows()]
opcoes2 = [nome_para_select(row, "1") for _, row in animal2_df.iterrows()]

col1, col2 = st.columns(2)
escolha1 = col1.selectbox("Animal 1", opcoes1)
escolha2 = col2.selectbox("Animal 2", opcoes2)

row1 = localizar_por_select(animal1_df, escolha1, "")
row2 = localizar_por_select(animal2_df, escolha2, "1")

if st.button("Gerar cruzamento hipotético", type="primary"):
    if row1 is None or row2 is None:
        st.error("Não consegui localizar um dos animais selecionados.")
        st.stop()

    html, repetidos = gerar_html(row1, row2, caminho_html_base)

    st.components.v1.html(html, height=800, scrolling=True)
    st.success(f"Repetições encontradas: {len(repetidos)}")

    st.download_button(
        "Baixar HTML do cruzamento",
        data=html,
        file_name="cruzamento_hipotetico.html",
        mime="text/html"
    )
