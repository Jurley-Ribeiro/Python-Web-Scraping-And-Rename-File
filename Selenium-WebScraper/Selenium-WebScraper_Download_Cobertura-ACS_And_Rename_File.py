import os, fnmatch
# shutil
import shutil
from unicodedata import normalize

#selenium
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import time
import os.path

driver = webdriver.Chrome(executable_path="C:\\Users\\ses83971874053\\PycharmProjects\\TesteSelenium\\chromedriver.exe")

# Cobertura de Agentes Comunitários de Saúde
wait = WebDriverWait(driver, 3)
try:
    driver.get("https://egestorab.saude.gov.br/paginas/acessoPublico/relatorios/relHistoricoCoberturaACS.xhtml")
    access_consulta_button = driver.find_element_by_id("j_idt58:tipoConsulta:1").click()
    access_uf_button = driver.find_elements_by_id("j_idt58:unidGeo:4")[0].click()
    driver.implicitly_wait(3)
    dropdown_regiao = driver.find_element_by_xpath("//*[@id='j_idt58:regiao']/option[7]")
    dropdown_regiao.click()
    dropdown_uf = driver.find_elements_by_xpath("//*[@id ='j_idt58:estados']/option[4]")[0].click()
    time.sleep(3)
    dropdown_mun = driver.find_elements_by_xpath("//*[@id='j_idt58:municipios']/option[2]")[0].click()
    dropdownCompetencia = driver.find_elements_by_xpath("//*[@id='j_idt58:compInicio']/option[2]")[0].click()
    download_button = driver.find_elements_by_xpath("//*[@id='j_idt58']/div[1]/div/div/div/div/div[3]/label[2]/i")[0].click()
    access_uf_button = driver.find_elements_by_id("j_idt58:unidGeo:4")[0].click()
    download_button = driver.find_elements_by_xpath("//*[@id='j_idt58']/div[1]/div/div/div/div/div[3]/label[2]/i")[0].click()
    export_button = driver.find_elements_by_id("j_idt58:exportar")[0].click()
    time.sleep(5)
finally:
    print("Sucesso")
    driver.quit()


def moveTo(from_path, to_path, pattern):
    # copy new file to process' folder
    for root, dirs, files in os.walk(from_path):
        for filename in fnmatch.filter(files, pattern):
            xls_file = os.path.join(root, filename)
            shutil.move(xls_file, to_path)
            break
        break


moveTo("C:\\Users\\ses83971874053\\Downloads", "C:\\Users\\ses83971874053\\Desktop\\Atencao Basica",
       "Cobertura-ACS*.xls")

# Rename file to AAAAMM
path_dir = "C:\\Users\\ses83971874053\\Desktop\\Atencao Basica"
path_files = os.listdir(path_dir)
pattern = "Cobertura-ACS*.xls"
for files in path_files:
    if fnmatch.fnmatch(files, pattern):
        s = files
        print(s)
        separa = (s.split("SUL"))
        a, b = (s.split("SUL"))
        month = (b[1:4])
        year = (b[-8:-4])
        print("Month = " + month)
        print("Year = " + year)
        if month == 'Jan':
            m = '01'
        elif month == 'Fev':
            m = '02'
        elif month == 'Mar':
            m = '03'
        elif month == 'Abr':
            m = '04'
        elif month == 'Jun':
            m = '06'
        elif month == 'Jul':
            m = '07'
        elif month == 'Ago':
            m = '08'
        elif month == 'Set':
            m = '09'
        elif month == 'Out':
            m = '10'
        elif month == 'Nov':
            m = '11'
        elif month == 'Dez':
            m = '12'
        else:
            m = 'nenhum'
        rename_file = path_dir + "\\" + (year + m) + ".xls"
        src = os.path.join(path_dir, s)
        os.rename(src, rename_file)
        print("Arquivo renomeado com sucesso:\n AAAAMM - ", year + m)
