from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import re
from time import sleep
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

def raspar_best_buy():
    # Configura o navegador
    driver = webdriver.Chrome()
    driver.get('https://www.bestbuy.com/?intl=nosplash')

    # Pesquisa por "macbook"
    pesquisar = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Type to search. Navigate forward to hear suggestions']"))
    )
    pesquisar.click()
    pesquisar.send_keys('macbook')

    pyautogui.press('enter')
    sleep(5)

    # Verifica se o arquivo Excel já existe
    try:
        wb = load_workbook('produtos_macbook_bestbuy.xlsx')
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = "Produtos"
        ws.append(["Nome", "Preço"])  # Cabeçalhos da planilha

    while True:
        # Encontrar todos os containers de produtos (li com classe sku-item)
        items = driver.find_elements(By.CSS_SELECTOR, "li.sku-item")

        # Iterar sobre os itens encontrados
        for item in items:
            # Capturar o nome do produto
            try:
                nome = item.find_element(By.CSS_SELECTOR, "h4.sku-title a").text
            except:
                nome = 'Nome não encontrado'

            # Capturar o preço do produto
            preco = 'Preço não encontrado'
            try:
                elements = item.find_elements(By.CSS_SELECTOR, "span[aria-hidden='true']")
                for element in elements:
                    texto = element.text
                    if re.match(r"^\$\d{1,3}(,\d{3})*(\.\d{2})?$", texto):
                        preco = texto
                        break
            except:
                pass

            # Adicionar os dados na planilha
            ws.append([nome, preco])

        # Salvar os dados no arquivo Excel após cada página
        wb.save('produtos_macbook_bestbuy.xlsx')

        # Verifica se há próxima página
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, "a.sku-list-page-next")
            if "aria-disabled" in next_button.get_attribute("class"):
                print("Não há mais páginas.")
                break
            else:
                next_button.click()
                sleep(5)  # Esperar a próxima página carregar
        except:
            print("Não encontrou o botão Next.")
            break

    # Ajustar largura das colunas
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)  # Obter a letra da coluna
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    # Salvar os dados no arquivo Excel ao final
    wb.save('produtos_macbook_bestbuy.xlsx')

    # Fechar o navegador
    driver.quit()

# Chama a função para raspar os dados
raspar_best_buy()
