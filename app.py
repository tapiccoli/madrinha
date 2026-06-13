import os
import re
import hashlib
import unicodedata
from collections import defaultdict

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup

# ==============================
# CONFIGURAÇÕES
# ==============================
ARQUIVO_PLANILHA_PADRAO = "extracao_bruta_pedigree1.xlsx"
ARQUIVO_HTML_BASE = "7gbasehipotetico.html"

st.set_page_config(page_title="Cruzamento Hipotético", layout="wide")
st.title("🐴 Sistema de Pedigree e Cruzamento Hipotético")


# ==============================
# MAPA DOS ITENS DA PLANILHA PARA O MODELO HTML
# ==============================

# Esses itens vêm da página principal de 5 gerações do animal.
MAPA_PRINCIPAL = {
    2: "ANIMAL",
    11: "2M",
    28: "2F",

    7: "3MM",
    15: "3MF",
    24: "3FM",
    32: "3FF",

    5: "4MMM",
    9: "4MMF",
    13: "4MFM",
    17: "4MFF",
    22: "4FMM",
    26: "4FMF",
    30: "4FFM",
    34: "4FFF",

    4: "5MMMM",
    6: "5MMMF",
    8: "5MMFM",
    10: "5MMFF",
    12: "5MFMM",
    14: "5MFMF",
    16: "5MFFM",
    18: "5MFFF",
    21: "5FMMM",
    23: "5FMMF",
    25: "5FMFM",
    27: "5FMFF",
    29: "5FFMM",
    31: "5FFMF",
    33: "5FFFM",
    35: "5FFFF",
}

# Ordem dos 16 itens extras que foram extraídos dos pais e avós.
CAMINHOS_16 = [
    "MMMM", "MMMF", "MMFM", "MMFF",
    "MFMM", "MFMF", "MFFM", "MFFF",
    "FMMM", "FMMF", "FMFM", "FMFF",
    "FFMM", "FFMF", "FFFM", "FFFF",
]

# Blocos extras da planilha:
# Item_36 a 51  = pai do animal principal
# Item_52 a 67  = avô paterno
# Item_68 a 83  = avó paterna
# Item_84 a 99  = mãe
# Item_100 a 115 = avô materno
# Item_116 a 131 = avó materna
BLOCOS_EXTRAS = [
    (36, "6", "M"),
    (52, "7", "MM"),
    (68, "7", "MF"),
    (84, "6", "F"),
    (100, "7", "FM"),
    (116, "7", "FF"),
]


# ==============================
# FUNÇÕES DE LIMPEZA E LEITURA
# ==============================

def limpar(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    if valor.lower() in ["nan", "none", "xxx", "não informado", "nao informado", "-"]:
        return ""
    return valor


def limpar_campo_pedigree(valor):
    """Limpa URLs, quebras de linha e textos inválidos."""
    if pd.isna(valor):
        return ""

    valor = str(valor).strip()

    if valor.lower() in ["nan", "none", "xxx", "não informado", "nao informado", "-"]:
        return ""

    # Remove URLs completas que às vezes vêm da página da ABCCC
    valor = re.sub(r"https?://\S+", "", valor)
    valor = re.sub(r"www\.\S+", "", valor)

    # Remove parâmetros soltos de URL, caso tenham vindo quebrados
    valor = re.sub(r"pesquisa/\S+", "", valor, flags=re.IGNORECASE)
    valor = re.sub(r"sbb_busca=\S+", "", valor, flags=re.IGNORECASE)
    valor = re.sub(r"token=\S+", "", valor, flags=re.IGNORECASE)

    valor = valor.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    valor = re.sub(r"\s+", " ", valor).strip()

    return valor


def normalizar_nome(nome):
    nome = limpar(nome).upper()
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("utf-8")
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def cor_por_nome(nome):
    h = hashlib.md5(nome.encode("utf-8")).hexdigest()
    return "#" + h[:6]


def extrair_primeiro_animal_do_texto(texto):
    """
    Lê Item_xx_TextoCompleto e tenta extrair somente o primeiro animal:
    NOME - SBB / PELAGEM

    Isso evita puxar URL ou dois animais dentro da mesma célula.
    """
    texto = limpar_campo_pedigree(texto)

    if not texto:
        return "", "", ""

    # Padrão de SBB nacional, importado e registros estrangeiros comuns:
    # B014477, CHD01234, UP00123, AP003985, *000532 etc.
    padrao_sbb = r"(?:\*[0-9]{3,}|[A-Z]{1,5}[0-9]{3,})"

    m = re.search(
        rf"^\s*(.*?)\s+-\s+({padrao_sbb})\s*/\s*(.*)$",
        texto,
        flags=re.IGNORECASE,
    )

    if not m:
        return "", "", ""

    nome = limpar_campo_pedigree(m.group(1))
    sbb = limpar_campo_pedigree(m.group(2))
    pelagem = limpar_campo_pedigree(m.group(3))

    # Se a pelagem veio grudada com outro animal, corta antes do próximo padrão "NOME - SBB /".
    # Ex.: "Tordilha Tapada ACULEO NUTRIA II - CHD25321 / Alazã"
    corte = re.search(
        rf"\s+[A-ZÁÉÍÓÚÃÕÂÊÔÇ0-9][A-ZÁÉÍÓÚÃÕÂÊÔÇ0-9\s'\.\-]+\s+-\s+{padrao_sbb}\s*/",
        pelagem,
        flags=re.IGNORECASE,
    )
    if corte:
        pelagem = pelagem[:corte.start()].strip()

    # Segurança extra: se ainda sobrou separador de outro animal, corta.
    if " - " in pelagem:
        pelagem = pelagem.split(" - ")[0].strip()

    return nome, sbb, pelagem


def get_item(row, numero):
    prefixo = f"Item_{numero:02d}"

    nome_col = limpar_campo_pedigree(row.get(f"{prefixo}_Nome", ""))
    sbb_col = limpar_campo_pedigree(row.get(f"{prefixo}_SBB", ""))
    pelagem_col = limpar_campo_pedigree(row.get(f"{prefixo}_Pelagem", ""))
    texto_completo = limpar_campo_pedigree(row.get(f"{prefixo}_TextoCompleto", ""))

    nome_txt, sbb_txt, pelagem_txt = extrair_primeiro_animal_do_texto(texto_completo)

    # Prioridade: TextoCompleto para preencher campos faltantes ou corrigir sujeira.
    nome = nome_txt or nome_col
    sbb = sbb_txt or sbb_col
    pelagem = pelagem_txt or pelagem_col

    # Se o nome veio grudado com registro/pelagem, corta.
    if " - " in nome:
        nome = nome.split(" - ")[0].strip()
    if " / " in nome:
        nome = nome.split(" / ")[0].strip()

    # Se a pelagem ficou longa demais ou com URL/texto estranho, limpa.
    pelagem = limpar_campo_pedigree(pelagem)

    return {
        "item": numero,
        "nome": nome,
        "sbb": sbb,
        "pelagem": pelagem,
        "texto_completo": texto_completo,
    }


def formatar_animal(animal):
    nome = limpar_campo_pedigree(animal.get("nome", ""))
    sbb = limpar_campo_pedigree(animal.get("sbb", ""))
    pelagem = limpar_campo_pedigree(animal.get("pelagem", ""))

    if not nome:
        return ""

    texto = nome

    if sbb:
        texto += f" - {sbb}"

    if pelagem:
        texto += f" / {pelagem}"

    return texto


def nome_linha(row):
    nome = limpar(row.get("Item_02_Nome", ""))
    sbb = limpar(row.get("SBB Pesquisado", ""))
    if nome and sbb:
        return f"{nome} | {sbb}"
    if nome:
        return nome
    return sbb


def buscar_por_opcao(df, opcao):
    sbb = opcao.split("|")[-1].strip()
    achou = df[df["SBB Pesquisado"].astype(str).str.strip() == sbb]
    if achou.empty:
        return None
    return achou.iloc[0]


# ==============================
# MONTAGEM DO MAPA DO PEDIGREE
# ==============================

def montar_mapa_pedigree(row, sufixo=""):
    """
    Retorna um dicionário:
    placeholder do HTML -> animal

    Exemplo:
    ANIMAL -> animal principal
    2M -> pai
    3MM -> avô paterno
    7MMMMMM -> ancestral de 7ª geração

    No cruzamento hipotético, o lado da fêmea usa sufixo '1':
    ANIMAL1, 2M1, 3MM1...
    """
    mapa = {}

    for numero_item, chave in MAPA_PRINCIPAL.items():
        mapa[chave + sufixo] = get_item(row, numero_item)

    for inicio, geracao, prefixo in BLOCOS_EXTRAS:
        for deslocamento, caminho in enumerate(CAMINHOS_16):
            numero_item = inicio + deslocamento
            chave = f"{geracao}{prefixo}{caminho}{sufixo}"
            mapa[chave] = get_item(row, numero_item)

    return mapa


def normalizar_sbb(sbb):
    sbb = str(sbb).strip().upper()
    if sbb in ["", "XXX", "-", "NÃO INFORMADO", "NAO INFORMADO", "NAN", "NONE"]:
        return ""
    return sbb


def cor_por_sbb(sbb):
    h = hashlib.md5(sbb.encode("utf-8")).hexdigest()
    return "#" + h[:6]


def gerar_repeticoes(mapas):
    """
    Identifica repetição pelo SBB, não pelo nome.
    Mesmo nome com SBB diferente = animais diferentes.
    Pelagem nunca entra na conta.
    """
    ocorrencias = defaultdict(list)

    for placeholder, animal in mapas.items():
        sbb = normalizar_sbb(animal.get("sbb", ""))

        if not sbb:
            continue

        ocorrencias[sbb].append({
            "placeholder": placeholder,
            "animal": animal,
            "geracao": placeholder[0] if placeholder and placeholder[0].isdigit() else "1",
        })

    return {
        sbb: dados
        for sbb, dados in ocorrencias.items()
        if len(dados) > 1
    }

def aplicar_valores_no_html(soup, mapa_animais, repetidos, manter_segunda_arvore=True):
    # Se for consulta individual, remove a segunda árvore do modelo hipotético.
    if not manter_segunda_arvore:
        tabelas_diretas = soup.body.find_all("table", recursive=False) if soup.body else []
        if len(tabelas_diretas) >= 2:
            tabelas_diretas[1].decompose()

    for td in soup.find_all("td"):
        texto_original = td.get_text(strip=True)
        if texto_original not in mapa_animais:
            continue

        animal = mapa_animais[texto_original]
        sbb_norm = normalizar_sbb(animal.get("sbb", ""))
        texto_formatado = formatar_animal(animal)

        td.clear()

        span = soup.new_tag("span")
        span["style"] = "display:block; font-size:9px; line-height:1.1em; font-family:Arial, sans-serif;"

        if texto_formatado:
            strong = soup.new_tag("strong")
            strong.string = texto_formatado
            span.append(strong)
        else:
            span.string = ""

        td.append(span)

        # Coloração também é feita por SBB, para não confundir pelagem ou nomes iguais.
        if sbb_norm in repetidos:
            cor = cor_por_sbb(sbb_norm)
            estilo_atual = td.get("style", "")
            td["style"] = f"{estilo_atual}; border-left: 8px solid {cor} !important;"

    style_tag = soup.new_tag("style")
    style_tag.string = """
    body, table, td, th, span, p, div {
        font-size: 9px !important;
        line-height: 1.1em !important;
        font-family: Arial, sans-serif !important;
    }
    td {
        vertical-align: middle !important;
        overflow-wrap: anywhere !important;
    }
    .relatorio-duplicacoes {
        margin-top: 20px;
        font-family: Arial, sans-serif;
        font-size: 13px;
    }
    .relatorio-duplicacoes table {
        border-collapse: collapse;
        width: 100%;
    }
    .relatorio-duplicacoes th, .relatorio-duplicacoes td {
        border: 1px solid #999;
        padding: 5px;
        font-size: 12px !important;
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
    for titulo in ["Animal", "SBB", "Quantidade", "Gerações", "Posições"]:
        th = soup.new_tag("th")
        th.string = titulo
        cab.append(th)
    tabela.append(cab)

    for sbb, ocorrs in sorted(repetidos.items(), key=lambda x: len(x[1]), reverse=True):
        nome_exibicao = ocorrs[0]["animal"].get("nome", "")
        geracoes = [o["geracao"] for o in ocorrs]
        posicoes = [o["placeholder"] for o in ocorrs]

        tr = soup.new_tag("tr")
        cor = cor_por_sbb(sbb)

        td_nome = soup.new_tag("td")
        td_nome["style"] = f"border-left: 8px solid {cor};"
        td_nome.string = nome_exibicao
        tr.append(td_nome)

        td_sbb = soup.new_tag("td")
        td_sbb.string = sbb
        tr.append(td_sbb)

        td_qtd = soup.new_tag("td")
        td_qtd.string = str(len(ocorrs))
        tr.append(td_qtd)

        td_ger = soup.new_tag("td")
        td_ger.string = "x".join(geracoes)
        tr.append(td_ger)

        td_pos = soup.new_tag("td")
        td_pos.string = " / ".join(posicoes)
        tr.append(td_pos)

        tabela.append(tr)

    div.append(tabela)
    soup.body.append(div)

def gerar_html(row_macho, row_femea=None):
    if not os.path.exists(ARQUIVO_HTML_BASE):
        st.error(f"Arquivo base '{ARQUIVO_HTML_BASE}' não encontrado na pasta do app.")
        st.stop()

    with open(ARQUIVO_HTML_BASE, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    if row_femea is None:
        mapa = montar_mapa_pedigree(row_macho, sufixo="")
        repetidos = gerar_repeticoes(mapa)
        aplicar_valores_no_html(soup, mapa, repetidos, manter_segunda_arvore=False)
        montar_relatorio_html(soup, repetidos)
        return str(soup), repetidos

    mapa_macho = montar_mapa_pedigree(row_macho, sufixo="")
    mapa_femea = montar_mapa_pedigree(row_femea, sufixo="1")

    mapa = {}
    mapa.update(mapa_macho)
    mapa.update(mapa_femea)

    repetidos = gerar_repeticoes(mapa)
    aplicar_valores_no_html(soup, mapa, repetidos, manter_segunda_arvore=True)
    montar_relatorio_html(soup, repetidos)
    return str(soup), repetidos


# ==============================
# INTERFACE STREAMLIT
# ==============================

st.markdown("Este protótipo usa o modelo HTML em árvore, não lista em linha.")

arquivo = st.file_uploader("Carregue a planilha de extração bruta", type=["xlsx"])

if arquivo is not None:
    df = pd.read_excel(arquivo, dtype=str)
else:
    if not os.path.exists(ARQUIVO_PLANILHA_PADRAO):
        st.warning(f"Carregue uma planilha ou coloque '{ARQUIVO_PLANILHA_PADRAO}' na pasta do app.")
        st.stop()
    df = pd.read_excel(ARQUIVO_PLANILHA_PADRAO, dtype=str)

if "SBB Pesquisado" not in df.columns:
    st.error("A planilha precisa conter a coluna 'SBB Pesquisado'.")
    st.stop()

opcoes = [nome_linha(row) for _, row in df.iterrows()]

modo = st.sidebar.radio("Modo", ["Consultar pedigree individual", "Cruzamento hipotético"])

if modo == "Consultar pedigree individual":
    escolha = st.selectbox("Selecione o animal", opcoes)
    row = buscar_por_opcao(df, escolha)

    if st.button("Gerar pedigree") and row is not None:
        html, repetidos = gerar_html(row)
        st.components.v1.html(html, height=720, scrolling=True)
        st.download_button("Baixar HTML", data=html, file_name="pedigree_individual.html", mime="text/html")
        st.success(f"Repetições encontradas: {len(repetidos)}")

else:
    col1, col2 = st.columns(2)
    macho = col1.selectbox("Selecione o macho", opcoes)
    femea = col2.selectbox("Selecione a fêmea", opcoes)

    row_macho = buscar_por_opcao(df, macho)
    row_femea = buscar_por_opcao(df, femea)

    if st.button("Gerar cruzamento hipotético") and row_macho is not None and row_femea is not None:
        html, repetidos = gerar_html(row_macho, row_femea)
        st.components.v1.html(html, height=720, scrolling=True)
        st.download_button("Baixar HTML", data=html, file_name="cruzamento_hipotetico.html", mime="text/html")
        st.success(f"Repetições encontradas: {len(repetidos)}")
