from customtkinter import *
from PIL import Image
from tkinter import messagebox
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time
import pyautogui
import os
from selenium.webdriver.common.keys import Keys
import datetime
from datetime import datetime,timedelta
from selenium.common.exceptions import TimeoutException
import os
import shutil   
import webbrowser

grupo_whasapp = 'Abastecimento de Lojas'
link_painel_demanda_abs_new = r'link aqui'
folder_status_olpn = r'C:\Users\2960006959\Desktop\projetos\PY---Bookmark-Posting-Tool-B.P.T---CB\diretorio_status_olpn'
link_folder_sharepoint = r'https://viavarejo.sharepoint.com/:f:/s/grupo_de_dados_wave/EkLSnUXubD9BmGsJaEm9vn0BxWMaidjLc_8rjYPe_5QY2w?e=k4vlJZ'

def automacao_principal():
    def clear_folder_status_olpn(folder):
        if os.path.exists(folder):
            for arquivo in os.listdir(folder):
                arquivo_caminho = os.path.join(folder, arquivo)
                try:
                    if os.path.isfile(arquivo_caminho) or os.path.islink(arquivo_caminho):
                        os.unlink(arquivo_caminho)
                    elif os.path.isdir(arquivo_caminho):
                        shutil.rmtree(arquivo_caminho)
                except Exception as e:
                    print(f'Erro ao deletar o arquivo {arquivo_caminho}: {e}')
        else:
            os.makedirs(folder)
    clear_folder_status_olpn(folder_status_olpn)
    print('DIRETORIO DO STATUS OLPN FOI LIMPO!')

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory":folder_status_olpn,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option('prefs', prefs)
    print('CHROME CONFIGURADO')

    ibm_cognos = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    ibm_cognos.get('https://viavp-sci.sce.manh.com/bi/?perspective=home')
    ibm_cognos.maximize_window()

    try:
        WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.presence_of_element_located((By.XPATH,'//*[@id="content"]/div[2]/div[1]'))
        )
        print('Elemento de localização(Título) encontrado')
    except TimeoutException as e:
        messagebox.showinfo('Elemento de localização','Elemento de localização(Título) não encontrado.', e)

    try:
        azure = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="downshift-0-toggle-button"]'))
        )
        time.sleep(2)
        azure.click()
        time.sleep(1)
        pyautogui.press('down')
        pyautogui.press('enter')
        print('Elemento de localização(Azure) selecionado')
    except TimeoutException as e:
        messagebox.showinfo('Elemento de localização', 'Elemento de localização(azure) não encontrado', e)     

    try:
        email = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="i0116"]'))
        )
        time.sleep(1)
        email.send_keys('vinicius.barbosa@viavarejo.com.br')
        time.sleep(1)
        pyautogui.press('enter')
        print('Elemento de localização(email) preechido')
    except TimeoutException as e:
        messagebox.showinfo('Elemento de localização', 'Elemento de localização(email) não encontrado', e)

    try:
        senha = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="i0118"]')))
        time.sleep(1)
        senha.send_keys('via@1660')
        time.sleep(1)
        pyautogui.press('enter')
        print('Elemento de localização(senha) foi preenchido')
    except TimeoutException as e:
        messagebox.showinfo('ELemento de localização','Elemento de localização(senha) não encontrada',e)

    try:
        WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.presence_of_element_located((By.XPATH,'//*[@id="lightbox"]/div[3]/div/div[2]/div/div[1]'))
        )
        pyautogui.press('enter')
        print('Elemtno de localização(confirmação de login) aceita')
    except TimeoutError as e:
        messagebox.showinfo('ELemento de localização','Elemento de localização(confirmação de login) não encontrada',e)

    try:
        WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.presence_of_element_located((By.XPATH,'//*[@id="com.ibm.bi.glass.common.viewSwitcher"]'))
        )
        print('Elemeno de localização(tela principal) encontrado')
    except TimeoutException as e:
        messagebox.showinfo('ELemento de localização','Elemento de localização(tela principal) não encontrada',e)

    ibm_cognos.get('https://viavp-sci.sce.manh.com/bi/?perspective=authoring&id=i79E326D8D72B45F795E0897FCE0606F6&objRef=i79E326D8D72B45F795E0897FCE0606F6&action=run&format=CSV&cmPropStr=%7B%22id%22%3A%22i79E326D8D72B45F795E0897FCE0606F6%22%2C%22type%22%3A%22report%22%2C%22defaultName%22%3A%223.11%20-%20Status%20Wave%20%2B%20oLPN%22%2C%22permissions%22%3A%5B%22execute%22%2C%22read%22%2C%22traverse%22%5D%7D')

    try:
        frame = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.frame_to_be_available_and_switch_to_it((By.XPATH,'//*[@id="rsIFrameManager_1"]'))
        )
        print('Elemento de verificação(iframe) encontrado')
        print('Conectaco ao iframe')
    except TimeoutException as e:
        messagebox.showinfo('ELemento de localização','Elemento de localização(iframe) não encontrado',e)

    try:
        opcao_1200 = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="dv17_ValueComboBox"]'))
        )
        selecionador = Select(opcao_1200)
        selecionador.select_by_value('1200')
        print('Elemento de localização(1200) selecionado')
    except TimeoutException as e:
        messagebox.showinfo('ELemento de localização','Elemento de localização(1200) não encontrado',e)

    data_inicio = datetime.now()
    weekday = data_inicio.weekday()

    if weekday in [1,2,3,4,5]:
        data_final = data_inicio - timedelta(days=1)
    elif weekday == 6:
        data_final = data_inicio - timedelta(days=2)
    elif weekday == 0:
        data_final = data_inicio - timedelta(days=3)

    data_inicio_tratada = data_inicio.strftime('%d/%m/%Y')
    data_final_tratada = data_final.strftime('%d/%m/%Y')

    try:
        campo_data_inicio = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="dv57__tblDateTextBox__txtInput"]'))
        )
        campo_data_inicio.clear()
        time.sleep(1)
        campo_data_inicio.send_keys(data_inicio_tratada)
        print('Elemento de localização(Campo data início) localizado')
        print('Campo início preenchido')
    except TimeoutException as e:
        messagebox.showinfo('Elemento de localização','Elemento de localização(Campo data início) não localizado')

    try:
        campo_data_final = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="dv65__tblDateTextBox__txtInput"]'))
        )
        campo_data_final.clear()
        time.sleep(1)
        campo_data_final.send_keys(data_final_tratada)
        print('Elemento de localização(Campo data final) localizado')
        print('Campo final preenchido')
    except TimeoutException as e:
        messagebox.showinfo('Elemento de localização',' Elemento de localização(Campo data final) não localizado')

    try:
        expedicoes = WebDriverWait(ibm_cognos,30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="dv73_HyperLink_0"]'))
        )
        expedicoes.click()
        print('Elemento de localização(todos) localizado')
        print('Campo todos selecionado')
    except TimeoutException as e:
        messagebox.showinfo('Elemento de localização','Elemento de localização(todos) não localizado')

    try:
        download = WebDriverWait(ibm_cognos, 30).until(
            expected_conditions.element_to_be_clickable((By.XPATH,'//*[@id="dv133"]'))
        )
        download.click()
        print('Elemento de localização(concluir) localizado')
        print('Campo concluir selecionado')
    except TimeoutException as e :
        messagebox.showinfo('Elemento de localização','Elemento de localização(concluir) não localizado')

    def espera_do_download(diretorio, timeout=300):
        aguardando_tempo = 0
        while aguardando_tempo < timeout:
            arquivos = os.listdir(diretorio)
            csv_files = [arquivo for arquivo in arquivos if arquivo.endswith('.csv')]
            if csv_files:
                print('Downloand do arquivo do tipo CSV completo')
                return csv_files[0]
            time.sleep(2)
            aguardando_tempo +=2
        print('Nenhum arquibo do tipo CSV encontrado')
        return None
    
    if espera_do_download(folder_status_olpn):
        arquivo_baixado = espera_do_download(folder_status_olpn)
        if arquivo_baixado:
            caminho_arquivo_baixado  = os.path.join(folder_status_olpn, arquivo_baixado)
            new_name = os.path.join(folder_status_olpn, 'status_olpn.csv')

            try:
                os.rename(caminho_arquivo_baixado, new_name)
                print(f'Arquivo renomeado para {new_name}')
            except Exception as e:
                print(f'Erro ao renomear o arquivo {e}')
    else:
        print('O download não foi concluído no tempo estimado')
        messagebox.showerror('Status download','O download não foi concluído dentro do tempo estipulado. 300 segundos')
        ibm_cognos.quit()

    ibm_cognos.quit()
    
    
    time.sleep(30)

    ibm_cognos.quit()



automacao_principal()