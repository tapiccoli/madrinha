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
    """Limpeza simples para campos vazios."""
    if pd.isna(valor):
        return ""
    valor = str(valor).strip()
    if valor.lower() in ["nan", "none", "xxx", "não informado", "nao informado", "-"]:
        return ""
    return valor


# Pelagens comuns para não confundir pelagem com nome de animal
PELAGENS_CONHECIDAS = {
    "ALAZA", "ALAZÃ", "ALAZAO", "ALAZÃO",
    "BAIA", "BAIO", "BRAGADA", "COLORADA", "COLORADO",
    "DOURADILHA", "GATEADA", "LOBUNA", "MOURA", "PICAÇA", "PICACA",
    "PRETA", "PRETO", "ROSILHA", "RUANA", "TOBIANA", "TOSTADA",
    "TORDILHA", "ZAINA",
    "TORDILHA TAPADA", "TORDILHA LOBUNA", "TORDILHA NEGRA",
    "GATEADA ROSILHA", "GATEADA BRAGADA", "GATEADA RUIVA BRAGADA",
    "ROSILHA MOURA", "ROSILHA COLORADA", "ROSILHA MOURA TAPADA",
    "TOSTADA ESCURA", "TOSTADA ALAZÃ", "TOSTADA ALAZA",
    "COLORADA REQUEIMADA", "TORDILHA VINAGRE", "ALAZÃ BRAGADA",
    "GATEADA ROSILHA TAPADA", "ROSILHA COLORADA TAPADA",
    "ZAINA COLORADA", "PICAÇA", "PICAÇA NEGRA", "PICAÇA COLORADA",
    "ZAINA NEGRA", "BAIA BRANCA", "GATEADA ESCURA", "GATEADA RUIVA",
    "TORDILHA BRAGADA", "TORDILHA VINAGRE BRAGADA", "TOSTADA REQUEIMADA",
    "TOSTADA REQUEIMADA BRAGADA", "MOURA TAPADA", "ROSILHA TAPADA",
}

REGISTRO_RE = re.compile(r"(\*\d{4,6}|[A-Z]{1,4}\d{4,6})", re.IGNORECASE)
URL_RE = re.compile(r"https?://\S+|www\.\S+", re.IGNORECASE)


def limpar_campo(valor):
    """Remove URL, quebras e lixo básico sem tentar interpretar o conteúdo."""
    valor = limpar(valor)
    if not valor:
        return ""
    valor = URL_RE.sub("", valor)
    valor = valor.replace("\n", " ").replace("\r", " ").replace("\xa0", " ")
    valor = re.sub(r"\s+", " ", valor).strip()
    valor = valor.strip(" -/")
    return valor


def normalizar_nome(nome):
    nome = limpar_campo(nome).upper()
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("utf-8")
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome


def parece_pelagem(texto):
    txt = normalizar_nome(texto)
    if not txt:
        return False
    if txt in {normalizar_nome(p) for p in PELAGENS_CONHECIDAS}:
        return True
    # Quando o campo começa com uma pelagem e depois vem outro animal, ex.: "Tordilha Tapada - TREN TREN..."
    for pel in PELAGENS_CONHECIDAS:
        pel_norm = normalizar_nome(pel)
        if txt == pel_norm or txt.startswith(pel_norm + " -"):
            return True
    return False


def extrair_sbb(texto):
    texto = limpar_campo(texto).upper()
    if not texto:
        return ""
    m = REGISTRO_RE.search(texto)
    if not m:
        return ""
    return m.group(1).upper()


def extrair_animal_de_texto(texto):
    """
    Extrai apenas UM animal de textos bagunçados.
    Entrada aceita:
      NOME - SBB / PELAGEM
      - NOME - SBB / / OUTRO ANIMAL
      PELAGEM - NOME - SBB / PELAGEM / OUTRO ANIMAL
      NOME - SBB / URL
    Saída: nome, sbb, pelagem.
    """
    texto = limpar_campo(texto)
    if not texto:
        return {"nome": "", "sbb": "", "pelagem": ""}

    texto = URL_RE.sub("", texto)
    texto = texto.strip(" -/")

    m = REGISTRO_RE.search(texto)
    if not m:
        return {"nome": "", "sbb": "", "pelagem": ""}

    sbb = m.group(1).upper()

    antes = texto[:m.start()].strip(" -/")
    depois = texto[m.end():].strip()

    # O nome é o último bloco antes do SBB que não seja pelagem.
    partes_antes = [p.strip(" -/") for p in re.split(r"\s+-\s+", antes) if p.strip(" -/")]
    nome = ""
    for parte in reversed(partes_antes):
        if not parece_pelagem(parte):
            nome = parte
            break
    if not nome and partes_antes:
        nome = partes_antes[-1]

    # Pelagem é apenas o primeiro conteúdo real depois da primeira barra.
    # Se depois da barra vier nome de outro animal, não usa.
    pelagem = ""
    if "/" in depois:
        pedacos = [p.strip(" -/") for p in depois.split("/")]
        for pedaco in pedacos:
            if not pedaco:
                continue
            if URL_RE.search(pedaco):
                continue
            if REGISTRO_RE.search(pedaco):
                continue
            if " - " in pedaco:
                continue
            if parece_pelagem(pedaco):
                pelagem = pedaco
                break
    else:
        # Casos raros: "NOME - SBB PELAGEM". Só usa se parecer pelagem conhecida.
        resto = depois.strip(" -/")
        if parece_pelagem(resto):
            pelagem = resto

    if pelagem and not parece_pelagem(pelagem):
        pelagem = ""

    return {
        "nome": limpar_campo(nome),
        "sbb": sbb,
        "pelagem": limpar_campo(pelagem),
    }


def limpar_pelagem(valor):
    """
    Mantém somente pelagens reais.
    Isso evita que a última coluna use o nome do próximo animal como se fosse pelagem.
    Exemplo ruim que passa a ser bloqueado:
      TREN TREN ARREBOL - *000413 / ACULEO NUTRIA II
    """
    valor = limpar_campo(valor)
    if not valor:
        return ""
    if URL_RE.search(valor):
        return ""
    if REGISTRO_RE.search(valor):
        return ""
    if " - " in valor:
        return ""
    if "/" in valor:
        valor = valor.split("/")[0].strip()
    if valor.upper() in ["XXX", "NÃO INFORMADO", "NAO INFORMADO"]:
        return ""

    # Regra principal: só aceita se parecer pelagem conhecida.
    # Nomes de animais como LICORERA, MAPOLA, ACULEO NUTRIA II, BT RESTINGA etc. são descartados.
    if not parece_pelagem(valor):
        return ""

    return valor.strip(" -/")


def cor_por_nome(nome):
    h = hashlib.md5(nome.encode("utf-8")).hexdigest()
    return "#" + h[:6]


def get_item(row, numero):
    prefixo = f"Item_{numero:02d}"

    bruto_nome = row.get(f"{prefixo}_Nome", "")
    bruto_sbb = row.get(f"{prefixo}_SBB", "")
    bruto_pelagem = row.get(f"{prefixo}_Pelagem", "")
    bruto_texto = row.get(f"{prefixo}_TextoCompleto", "")

    # Primeiro tenta interpretar os campos que podem vir completos/bagunçados.
    candidatos_parse = [
        extrair_animal_de_texto(bruto_texto),
        extrair_animal_de_texto(bruto_sbb),
        extrair_animal_de_texto(bruto_nome),
    ]

    nome = limpar_campo(bruto_nome)
    sbb = extrair_sbb(bruto_sbb)
    pelagem = limpar_pelagem(bruto_pelagem)

    # Se o nome veio como pelagem ou vazio, busca o nome nos textos completos.
    if not nome or parece_pelagem(nome) or REGISTRO_RE.search(nome):
        for cand in candidatos_parse:
            if cand["nome"]:
                nome = cand["nome"]
                break

    # Se o campo nome veio como "NOME - SBB / PELAGEM", separa corretamente.
    parse_nome = extrair_animal_de_texto(bruto_nome)
    if parse_nome["nome"] and parse_nome["sbb"]:
        nome = parse_nome["nome"]
        if not sbb:
            sbb = parse_nome["sbb"]
        if not pelagem:
            pelagem = parse_nome["pelagem"]

    # SBB sempre deve ser só o registro, nunca o texto todo.
    if not sbb:
        for cand in candidatos_parse:
            if cand["sbb"]:
                sbb = cand["sbb"]
                break

    # Pelagem: usa a coluna própria; se estiver vazia, usa o texto completo.
    if not pelagem:
        for cand in candidatos_parse:
            if cand["pelagem"]:
                pelagem = cand["pelagem"]
                break

    # Segurança final: nome não pode conter URL, SBB nem outro animal grudado.
    nome = limpar_campo(nome)
    if " - " in nome and extrair_sbb(nome):
        nome = extrair_animal_de_texto(nome)["nome"]
    if "/" in nome:
        nome = nome.split("/")[0].strip()
    nome = nome.strip(" -/")

    sbb = normalizar_sbb(sbb)
    pelagem = limpar_pelagem(pelagem)

    return {
        "item": numero,
        "nome": nome,
        "sbb": sbb,
        "pelagem": pelagem,
    }


def formatar_animal(animal):
    nome = limpar_campo(animal.get("nome", ""))
    sbb = normalizar_sbb(animal.get("sbb", ""))
    pelagem = limpar_pelagem(animal.get("pelagem", ""))

    if not nome and not sbb:
        return ""

    texto = nome if nome else "Sem nome"
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
    sbb = limpar_campo(sbb).upper()
    if sbb in ["", "XXX", "-", "NÃO INFORMADO", "NAO INFORMADO", "NAN", "NONE"]:
        return ""
    m = REGISTRO_RE.search(sbb)
    if not m:
        return ""
    return m.group(1).upper()


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
