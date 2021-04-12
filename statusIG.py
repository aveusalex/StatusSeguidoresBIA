from selenium import webdriver
from time import sleep
from openpyxl import load_workbook
from datetime import datetime


def calcula_aumento(a: int, d: int):
    resultado = (a - d) / d

    return "{:.2f}".format(resultado)


def salva_ultimo(ultimo_seguidores, ultimo_data, diretorio_salvamento):
    with open(diretorio_salvamento, "w") as save:
        # ordem: seguidor, data.
        save.writelines([ultimo_seguidores + "\n", ultimo_data])


def carrega_ultimo_salvo(diretorio_salvamento):
    with open(diretorio_salvamento) as save:
        ultimos = save.readlines()

    return ultimos


def adicionaseguidor(diretorio_salvamento, diretorio_planilha):
    arquivo_excel = load_workbook(diretorio_planilha)
    planilha1 = arquivo_excel.active
    ultimos = carrega_ultimo_salvo(diretorio_salvamento)
    ultima_linha_seguidores = str(int(ultimos[0]) + 1)
    ultima_linha_data = str(int(ultimos[1]) + 1)
    seguidores = extrai_seguidores()
    data_atual = f"{datetime.now().date().day}/{datetime.now().date().month}/{datetime.now().date().year}"

    # escreve a quantidade de seguidores
    planilha1[f"B{ultima_linha_seguidores}"] = seguidores

    # escreve a data atual
    planilha1[f"A{ultima_linha_data}"] = data_atual

    # salva as linhas atuais
    salva_ultimo(ultima_linha_seguidores, ultima_linha_data, diretorio_salvamento)

    # calcula aumento
    seguidor_antes = planilha1[f"B{str(int(ultima_linha_data) - 1)}"]
    planilha1[f"D{ultima_linha_data}"] = calcula_aumento(int(seguidores), int(seguidor_antes.value))

    # calcula data
    data_antes = datetime.strptime(planilha1[f"A{str(int(ultima_linha_data) - 1)}"].value, "%d/%m/%Y")
    aux = datetime.strptime(data_atual, "%d/%m/%Y")
    qtd_dias = abs((aux - data_antes).days)
    planilha1[f"F{ultima_linha_data}"] = qtd_dias

    arquivo_excel.save(diretorio_planilha)


def acessa_site():
    driver = webdriver.Firefox(executable_path="C:\\geckodriver\\geckodriver.exe")
    driver.get("https://www.instagram.com/biaufg/")
    sleep(5)

    return driver


def extrai_seguidores():
    driver = acessa_site()
    seguidores = driver.find_elements_by_xpath("//span[@class='g47SY ']")
    seguidores = seguidores[1].text
    driver.quit()

    return seguidores


def main(diretorio_salvamento="C:\\Users\\alexv\\Documents\\relatoriobia-saves\\relt.txt",
         diretorio_planilha="C:\\Users\\alexv\\Documents\\RelatorioBiaufg.xlsx"):

    adicionaseguidor(diretorio_salvamento, diretorio_planilha)


if __name__ == '__main__':
    main()
