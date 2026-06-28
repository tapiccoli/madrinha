import os
import time
from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
CAMINHO_SBB = os.path.join(PASTA_SCRIPT, "sbb.txt")
CAMINHO_SAIDA = os.path.join(PASTA_SCRIPT, "extracao_bruta_pedigree.xlsx")
CAMINHO_LOG = os.path.join(PASTA_SCRIPT, "log_erros.txt")

URL_PESQUISA = "https://www.cavalocrioulo.org.br/pesquisa/pesquisas.php"

# Estes são os 16 itens que serão extraídos da página do PAI e da MÃE.
ITENS_DESEJADOS = [4, 6, 8, 10, 12, 14, 16, 18, 21, 23, 25, 27, 29, 31, 33, 35]

# Quantidade final fixa da planilha:
# Item_01 até Item_35 = animal principal
# Item_36 até Item_51 = pai
# Item_52 até Item_67 = mãe
TOTAL_ITENS_FINAIS = 67


# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def registrar_erro(texto):
    """Registra erros em um TXT separado, sem criar colunas extras na planilha."""
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(CAMINHO_LOG, "a", encoding="utf-8") as f:
        f.write(f"[{data_hora}] {texto}\n")


def valor_ou_xxx(valor):
    """Garante que nenhum campo fique vazio na planilha."""
    if valor is None:
        return "xxx"

    valor = str(valor).strip()

    if not valor:
        return "xxx"

    if valor.lower() in ["nan", "none", "não informado", "nao informado", "-"]:
        return "xxx"

    return valor


def criar_animal_xxx():
    return {
        "nome": "xxx",
        "sbb": "xxx",
        "pelagem": "xxx",
        "texto_completo": "xxx"
    }


def preencher_item(linha, numero_item, animal):
    """Preenche um item específico da planilha com Nome, SBB, Pelagem e TextoCompleto."""
    linha[f"Item_{numero_item:02d}_Nome"] = valor_ou_xxx(animal.get("nome"))
    linha[f"Item_{numero_item:02d}_SBB"] = valor_ou_xxx(animal.get("sbb"))
    linha[f"Item_{numero_item:02d}_Pelagem"] = valor_ou_xxx(animal.get("pelagem"))
    linha[f"Item_{numero_item:02d}_TextoCompleto"] = valor_ou_xxx(animal.get("texto_completo"))


def preencher_bloco_xxx(linha, item_destino_inicial, quantidade=16):
    """
    Preenche um bloco inteiro com xxx.
    Usado quando pai ou mãe não existem, não abrem ou dão erro.
    """
    for numero_item in range(item_destino_inicial, item_destino_inicial + quantidade):
        preencher_item(linha, numero_item, criar_animal_xxx())


def garantir_estrutura_fixa(linha):
    """
    Garante que a linha tenha sempre Item_01 até Item_67.
    Isso evita célula em branco e evita a criação de colunas desalinhadas.
    """
    for i in range(1, TOTAL_ITENS_FINAIS + 1):
        for campo in ["Nome", "SBB", "Pelagem", "TextoCompleto"]:
            coluna = f"Item_{i:02d}_{campo}"
            if coluna not in linha or valor_ou_xxx(linha[coluna]) == "xxx":
                linha[coluna] = "xxx"


def colunas_ordenadas():
    """Define a ordem fixa das colunas do Excel."""
    colunas = ["SBB Pesquisado"]

    for i in range(1, TOTAL_ITENS_FINAIS + 1):
        colunas.extend([
            f"Item_{i:02d}_Nome",
            f"Item_{i:02d}_SBB",
            f"Item_{i:02d}_Pelagem",
            f"Item_{i:02d}_TextoCompleto",
        ])

    return colunas


# ============================================================
# EXTRAÇÃO DO HTML
# ============================================================

def extrair_animais_pagina_print(html):
    soup = BeautifulSoup(html, "html.parser")
    animais = []

    for forte in soup.find_all("strong"):
        nome = valor_ou_xxx(forte.get_text(" ", strip=True))

        pai = forte.parent
        texto = valor_ou_xxx(pai.get_text(" ", strip=True))

        links = pai.find_all("a")
        sbb = links[0].get_text(strip=True) if links else "xxx"
        sbb = valor_ou_xxx(sbb)

        pelagem = "xxx"
        if "/" in texto:
            partes = texto.split("/")
            if len(partes) > 1:
                pelagem = valor_ou_xxx(partes[-1].strip())

        animais.append({
            "nome": nome,
            "sbb": sbb,
            "pelagem": pelagem,
            "texto_completo": texto
        })

    return animais


def abrir_animal_por_sbb(driver, sbb):
    driver.get(URL_PESQUISA)
    time.sleep(0.5)

    campo_sbb = WebDriverWait(driver, 8).until(
        EC.presence_of_element_located((By.ID, "sbb"))
    )

    campo_sbb.clear()
    campo_sbb.send_keys(sbb)
    campo_sbb.send_keys(Keys.RETURN)

    time.sleep(1)


def abrir_pagina_imprimir_5g(driver):
    wait = WebDriverWait(driver, 8)

    link_5g = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, '//a[contains(@href, "quinta_geracao.php")]')
        )
    )

    url_5g = link_5g.get_attribute("href")
    driver.get(url_5g)

    link_imprimir = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//a[contains(@href, "quinta_geracao.php") and contains(@href, "print=true")]'
            )
        )
    )

    url_imprimir = link_imprimir.get_attribute("href")
    driver.get(url_imprimir)

    wait.until(lambda navegador: "print=true" in navegador.current_url)
    time.sleep(0.5)

    return driver.page_source, url_imprimir


# ============================================================
# MONTAGEM DA LINHA DA PLANILHA
# ============================================================

def montar_linha_excel(sbb_pesquisado, animais):
    """
    Extrai somente os 35 primeiros itens do animal principal.
    Mesmo se vierem mais de 35 animais no HTML, ignora o excedente.
    Mesmo se vierem menos de 35, completa com xxx.
    """
    linha = {
        "SBB Pesquisado": valor_ou_xxx(sbb_pesquisado)
    }

    for i in range(1, 36):
        if i <= len(animais):
            animal = animais[i - 1]
        else:
            animal = criar_animal_xxx()

        preencher_item(linha, i, animal)

    return linha


def extrair_itens_de_outro_sbb(driver, linha, animais_origem, item_origem, item_destino_inicial, nome_bloco):
    """
    Abre o SBB de um animal de origem e extrai somente os ITENS_DESEJADOS.
    Pai: item_origem 11 -> grava Item_36 até Item_51
    Mãe: item_origem 28 -> grava Item_52 até Item_67
    """
    try:
        indice_origem = item_origem - 1

        if indice_origem >= len(animais_origem):
            preencher_bloco_xxx(linha, item_destino_inicial)
            registrar_erro(
                f"{nome_bloco} | Item de origem {item_origem} não existe na extração principal."
            )
            return

        sbb_origem = valor_ou_xxx(animais_origem[indice_origem].get("sbb"))

        print(f"Abrindo {nome_bloco}: Item {item_origem} | SBB {sbb_origem}")

        if sbb_origem == "xxx":
            preencher_bloco_xxx(linha, item_destino_inicial)
            registrar_erro(
                f"{nome_bloco} | SBB vazio no Item {item_origem} do animal principal."
            )
            return

        abrir_animal_por_sbb(driver, sbb_origem)

        html_print, url_print = abrir_pagina_imprimir_5g(driver)

        print(f"URL extraída {nome_bloco}: {url_print}")

        animais_extraidos = extrair_animais_pagina_print(html_print)

        proximo_item = item_destino_inicial

        for numero_item in ITENS_DESEJADOS:
            indice_python = numero_item - 1

            if indice_python < len(animais_extraidos):
                animal = animais_extraidos[indice_python]
            else:
                animal = criar_animal_xxx()
                registrar_erro(
                    f"{nome_bloco} | SBB {sbb_origem} | Item {numero_item} não encontrado."
                )

            preencher_item(linha, proximo_item, animal)
            proximo_item += 1

    except Exception as e:
        preencher_bloco_xxx(linha, item_destino_inicial)
        registrar_erro(
            f"{nome_bloco} | Erro ao extrair bloco iniciado em Item_{item_destino_inicial:02d}: {str(e)}"
        )
        print(f"Erro em {nome_bloco}: {e}")


def criar_linha_erro_animal_principal(sbb, erro):
    """
    Se o animal principal não abrir, cria a linha completa com Item_01 até Item_67 = xxx.
    """
    linha = {
        "SBB Pesquisado": valor_ou_xxx(sbb)
    }

    for i in range(1, TOTAL_ITENS_FINAIS + 1):
        preencher_item(linha, i, criar_animal_xxx())

    registrar_erro(f"SBB PRINCIPAL {sbb} | {str(erro)}")

    return linha


def salvar_excel_com_nome_disponivel(df, caminho_base):
    nome_base, extensao = os.path.splitext(caminho_base)
    contador = 1

    while True:
        try:
            if contador == 1:
                caminho_final = caminho_base
            else:
                caminho_final = f"{nome_base}{contador}{extensao}"

            df.to_excel(caminho_final, index=False)
            print(f"Planilha salva em: {caminho_final}")
            return caminho_final

        except PermissionError:
            contador += 1


# ============================================================
# EXECUÇÃO
# ============================================================

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)

try:
    with open(CAMINHO_SBB, "r", encoding="utf-8") as file:
        lista_sbb = [linha.strip() for linha in file.readlines() if linha.strip()]

    dados_finais = []

    for sbb in lista_sbb:
        print(f"Processando SBB principal: {sbb}")

        try:
            abrir_animal_por_sbb(driver, sbb)

            html_print, url_imprimir = abrir_pagina_imprimir_5g(driver)

            animais = extrair_animais_pagina_print(html_print)
            animais_principal = animais.copy()

            linha = montar_linha_excel(sbb, animais)

            # Pai:
            # abre o SBB do Item 11 do animal principal
            # grava os 16 ITENS_DESEJADOS em Item_36 até Item_51
            extrair_itens_de_outro_sbb(
                driver=driver,
                linha=linha,
                animais_origem=animais_principal,
                item_origem=11,
                item_destino_inicial=36,
                nome_bloco="Pai"
            )

            # Mãe:
            # abre o SBB do Item 28 do animal principal
            # grava os 16 ITENS_DESEJADOS em Item_52 até Item_67
            extrair_itens_de_outro_sbb(
                driver=driver,
                linha=linha,
                animais_origem=animais_principal,
                item_origem=28,
                item_destino_inicial=52,
                nome_bloco="Mae"
            )

            garantir_estrutura_fixa(linha)
            dados_finais.append(linha)

            print(f"OK: {sbb}")

        except Exception as e:
            print(f"Erro no SBB principal {sbb}: {e}")
            dados_finais.append(criar_linha_erro_animal_principal(sbb, e))

    df = pd.DataFrame(dados_finais)

    # Mantém somente as colunas fixas esperadas.
    # Isso elimina qualquer coluna acidental, como Erro, URL, Total, Item_68 etc.
    colunas = colunas_ordenadas()
    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = "xxx"

    df = df[colunas]

    arquivo_salvo = salvar_excel_com_nome_disponivel(df, CAMINHO_SAIDA)

    print("Extração finalizada.")
    print(f"Planilha salva em: {arquivo_salvo}")
    print(f"Log de erros salvo em: {CAMINHO_LOG}")

finally:
    driver.quit()
