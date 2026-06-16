import os
import time
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
CAMINHO_SBB = os.path.join(PASTA_SCRIPT, "sbb.txt")
CAMINHO_SAIDA = os.path.join(PASTA_SCRIPT, "extracao_bruta_pedigree.xlsx")

URL_PESQUISA = "https://www.cavalocrioulo.org.br/pesquisa/pesquisas.php"

ITENS_DESEJADOS = [4, 6, 8, 10, 12, 14, 16, 18, 21, 23, 25, 27, 29, 31, 33, 35]

def extrair_animais_pagina_print(html):
    soup = BeautifulSoup(html, "html.parser")
    animais = []

    for forte in soup.find_all("strong"):
        nome = forte.get_text(" ", strip=True)

        # mantém item vazio para não quebrar a contagem
        if not nome:
            nome = "xxx"

        pai = forte.parent
        texto = pai.get_text(" ", strip=True)

        links = pai.find_all("a")
        sbb = links[0].get_text(strip=True) if links else ""

        if not sbb:
            sbb = "xxx"

        pelagem = ""
        if "/" in texto:
            partes = texto.split("/")
            if len(partes) > 1:
                pelagem = partes[-1].strip()

        if not pelagem:
            pelagem = "xxx"

        animais.append({
            "nome": nome,
            "sbb": sbb,
            "pelagem": pelagem,
            "texto_completo": texto if texto else "xxx"
        })

    return animais

def abrir_animal_por_sbb(driver, sbb):
    driver.get(URL_PESQUISA)
    time.sleep(2)

    campo_sbb = driver.find_element(By.ID, "sbb")
    campo_sbb.clear()
    campo_sbb.send_keys(sbb)
    campo_sbb.send_keys(Keys.RETURN)

    time.sleep(4)

def abrir_pagina_imprimir_5g(driver):

    wait = WebDriverWait(driver, 15)

    # encontra botão 5ª geração
    link_5g = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, '//a[contains(@href, "quinta_geracao.php")]')
        )
    )

    url_5g = link_5g.get_attribute("href")

    driver.get(url_5g)

    # espera botão imprimir aparecer
    wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//a[contains(@href, "quinta_geracao.php") and contains(@href, "print=true")]'
            )
        )
    )

    link_imprimir = driver.find_element(
        By.XPATH,
        '//a[contains(@href, "quinta_geracao.php") and contains(@href, "print=true")]'
    )

    url_imprimir = link_imprimir.get_attribute("href")

    driver.get(url_imprimir)

    # espera carregar página correta
    wait.until(
        lambda navegador: "print=true" in navegador.current_url
    )

    time.sleep(2)

    html = driver.page_source

    return html, url_imprimir

def montar_linha_excel(sbb_pesquisado, url_imprimir, animais):
    linha = {
        "SBB Pesquisado": sbb_pesquisado,
        "URL Impressao": url_imprimir,
        "Total Itens Extraidos Inicial": len(animais)
    }

    for i, animal in enumerate(animais, start=1):
        linha[f"Item_{i:02d}_Nome"] = animal["nome"]
        linha[f"Item_{i:02d}_SBB"] = animal["sbb"]
        linha[f"Item_{i:02d}_Pelagem"] = animal["pelagem"]
        linha[f"Item_{i:02d}_TextoCompleto"] = animal["texto_completo"]

    return linha


def extrair_itens_de_outro_sbb(driver, linha, animais_origem, item_origem, item_destino_inicial, nome_bloco):
    try:
        indice_origem = item_origem - 1
        sbb_origem = animais_origem[indice_origem]["sbb"]

        print(f"Abrindo {nome_bloco}: Item {item_origem} | SBB {sbb_origem}")

        if not sbb_origem:
            linha[f"Erro {nome_bloco}"] = "SBB vazio"
            return

        abrir_animal_por_sbb(driver, sbb_origem)

        html_print, url_print = abrir_pagina_imprimir_5g(driver)

        print(f"URL extraída: {url_print}")

        animais_extraidos = extrair_animais_pagina_print(html_print)

        proximo_item = item_destino_inicial

        for numero_item in ITENS_DESEJADOS:
            indice_python = numero_item - 1

            if indice_python < len(animais_extraidos):
                animal = animais_extraidos[indice_python]

                linha[f"Item_{proximo_item:02d}_Nome"] = animal["nome"]
                linha[f"Item_{proximo_item:02d}_SBB"] = animal["sbb"]
                linha[f"Item_{proximo_item:02d}_Pelagem"] = animal["pelagem"]
                linha[f"Item_{proximo_item:02d}_TextoCompleto"] = animal["texto_completo"]

            proximo_item += 1

        linha[f"URL {nome_bloco}"] = url_print

    except Exception as e:
        linha[f"Erro {nome_bloco}"] = str(e)
        print(f"Erro em {nome_bloco}: {e}")


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)

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

        linha = montar_linha_excel(sbb, url_imprimir, animais)

        # Item 11: pai do animal principal
        extrair_itens_de_outro_sbb(
            driver=driver,
            linha=linha,
            animais_origem=animais_principal,
            item_origem=11,
            item_destino_inicial=36,
            nome_bloco="Item 11"
        )

        # Item 28: mãe do animal principal
        extrair_itens_de_outro_sbb(
            driver=driver,
            linha=linha,
            animais_origem=animais_principal,
            item_origem=28,
            item_destino_inicial=84,
            nome_bloco="Item 28"
        )

    
        dados_finais.append(linha)

        print(f"OK: {sbb}")

    except Exception as e:
        print(f"Erro no SBB principal {sbb}: {e}")

        dados_finais.append({
            "SBB Pesquisado": sbb,
            "Erro": str(e)
        })

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

df_animal1 = pd.DataFrame(dados_finais)

# cria cópia para animal2
df_animal2 = df_animal1.copy()

# renomeia as colunas da aba animal2
novas_colunas = {}

for coluna in df_animal2.columns:

    if coluna.startswith("Item_"):
        novas_colunas[coluna] = coluna + "1"
    else:
        novas_colunas[coluna] = coluna

df_animal2.rename(columns=novas_colunas, inplace=True)

nome_base, extensao = os.path.splitext(CAMINHO_SAIDA)
contador = 1

while True:

    try:

        if contador == 1:
            caminho_final = CAMINHO_SAIDA
        else:
            caminho_final = f"{nome_base}{contador}{extensao}"

        with pd.ExcelWriter(
            caminho_final,
            engine="openpyxl"
        ) as writer:

            df_animal1.to_excel(
                writer,
                sheet_name="animal1",
                index=False
            )

            df_animal2.to_excel(
                writer,
                sheet_name="animal2",
                index=False
            )

        print(f"Planilha salva em: {caminho_final}")
        break

    except PermissionError:
        contador += 1

print("Extração finalizada.")
print(f"Planilha salva em: {caminho_final}")

driver.quit()