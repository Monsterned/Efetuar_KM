import pandas as pd
from openpyxl import load_workbook
import os
import numpy as np
import pyautogui
import ctypes
from datetime import datetime
import pyautogui
from selenium.webdriver.common.action_chains import ActionChains
import shutil
from datetime import datetime, timedelta

caminho = os.getcwd() 
caminho_sistema = caminho.replace("C", "T", 1)

def click_image(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_info_manifesto(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.moveTo(center_x, center_y)  # Movendo o cursor para a posição da imagem               
                pyautogui.moveRel(30, 0)  # Movendo o cursor para cima
                pyautogui.click()  # Clicando no local da imagem
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)


def confirmação(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    not_found_count = 0
    max_attempts = 3
    while not_found_count < max_attempts:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Imagem foi encontrada na tela.")
                not_found_count = 0  # Reseta o contador se a imagem for encontrada
            else:
                not_found_count += 1
                print(f"Imagem não encontrada na tela. Tentativa {not_found_count} de {max_attempts}.")
        except Exception as e:
            print(f"Erro ao tentar localizar a imagem: {e}")
            not_found_count += 1
        if not_found_count < max_attempts:
            pyautogui.sleep(1)
        else:
            print("Imagem não encontrada após 3 tentativas. Saindo da função.")
            break

def verificar_campo(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Campo encontrado.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def check_caps_lock():
    return ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0

def status_manifesto(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Campo encontrado.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")

        image_name2 = 'status_baixado.png'
        image_path2 = os.path.join(current_dir, caminho_imagem, image_name2)
        try:
            print("Tentando achar a segunda imagem.")
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                print("Campo encontrado.")
                click_image('botao_frota.png')
                pyautogui.press("alt")
                pyautogui.press("alt")
                pyautogui.press("right")
                for i in range(2): 
                    pyautogui.press("down")
                pyautogui.press("right")
                for i in range(14): 
                    pyautogui.press("down")
                pyautogui.press("enter")
                pyautogui.sleep(2)
                click_image('incluir.png')
                verificar_campo('status_tela_cancelamento.png')
                click_image('CTRC.png')
                for i in range(10): 
                    pyautogui.press("up")
                #click_image('cancelar_baixa_manifesto.png')
                pyautogui.press("down")  
                pyautogui.press("enter")   
                pyautogui.sleep(2)
                click_info_manifesto('filial_cancelamento.png')
                pyautogui.press("backspace")
                pyautogui.write(str(filial))
                pyautogui.sleep(1)
                click_info_manifesto('serie_cancelamento.png')
                pyautogui.press("backspace")
                pyautogui.write(str(serie))
                pyautogui.sleep(1)
                click_info_manifesto('manifesto_cancelamento.png')
                for i in range(10): 
                    pyautogui.press("backspace")
                pyautogui.write(str(manifesto))
                pyautogui.sleep(1)
                for i in range(2):
                    pyautogui.press("tab")
                pyautogui.press("1")
                pyautogui.sleep(1)
                pyautogui.press("tab")
                pyautogui.write("correção de km")
                pyautogui.sleep(1)
                click_image('salvar_cancelamento.png')
                click_image('sim_cancelamento.png')
                click_image('ok_cancelamento.png')
                click_image('voltar.png')
                pyautogui.sleep(2)
                click_image('incluir.png')
                pyautogui.sleep(2)
                click_info_manifesto('filial.png')
                pyautogui.press("backspace")
                pyautogui.write(str(filial))
                pyautogui.sleep(1)
                click_info_manifesto('serie.png')
                pyautogui.press("backspace")
                pyautogui.write(str(serie))
                pyautogui.sleep(1)
                click_info_manifesto('campo_manifesto.png')
                for i in range(10): 
                    pyautogui.press("backspace")
                pyautogui.write(str(manifesto))
                pyautogui.press("tab")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")

        pyautogui.sleep(1)

def check_caps_lock():
    return ctypes.windll.user32.GetKeyState(0x14) & 0xffff != 0

def alt_press(key):
    pyautogui.keyDown('alt')
    pyautogui.press(key)
    pyautogui.keyUp('alt')



# # #LOGIN
# if check_caps_lock():
#     pyautogui.press("capslock")  # Desativa o CAPS LOCK se estiver ativado
# pyautogui.keyDown('win')
# pyautogui.press("m")
# pyautogui.keyUp('win')
# #click_image('logo_rodopar_areatrabalho.png')#PC ESCRITORIO
# click_image('logo_rodopar_areatrabalho_resumido.png')#PC ESCRITORIO
# #click_image('logo_rodopar_areatrabalho.png')#PC CASA
# pyautogui.click()
# click_image('conectar_rodopar.png')
# #click_image('conectar_rodopar1.png')
# click_image('senha_rodopar_1.png')
# #click_image('senha_rodopar_2.png')
# pyautogui.write("17@mudar")
# click_image('ok_primeiro_login.png')
# #click_image('ok_primeiro_login2.png')
# click_image('sim_primeiro_login.png')
# #click_image('sim_primeiro_login2.png')
# click_image('segundo_login.png')    
# pyautogui.sleep(1)
# pyautogui.write("anascimento")
# pyautogui.press("tab")
# pyautogui.write("990607")
# for i in range(2): 
#     pyautogui.press("enter")
# click_image('filial_1.png')
# pyautogui.press("enter")

click_image('botao_frota.png')
pyautogui.press("alt")
pyautogui.press("alt")
pyautogui.press("right")
for i in range(2): 
    pyautogui.press("down")
pyautogui.press("right")
for i in range(6): 
    pyautogui.press("down")
pyautogui.press("right")
pyautogui.press("enter")

Planilha_km = pd.read_excel("BASE_KM.xlsx")
#print(Planilha_km)

numero_linhas = len(Planilha_km)
#print(numero_linhas)

for i, linha in enumerate(Planilha_km.index):
    filial = Planilha_km.loc[linha, "FILIAL"]
    serie = Planilha_km.loc[linha, "SERIE"]
    manifesto = Planilha_km.loc[linha, "MANIFESTO"]
    km = Planilha_km.loc[linha, "KM"]
    agora = datetime.now()
    um_minuto_atras = agora - timedelta(minutes=1)
    agora_formatado = agora.strftime("%d/%m/%Y %H:%M")
    um_minuto_atras_formatado = um_minuto_atras.strftime("%d/%m/%Y%H:%M")
    falta = numero_linhas - i
    print(f'filial:{filial} serie:{serie} manifesto:{manifesto} km:{km} falta:{falta}')

    click_image('incluir.png')
    pyautogui.sleep(2)
    click_info_manifesto('filial.png')
    pyautogui.press("backspace")
    pyautogui.write(str(filial))
    pyautogui.sleep(1)
    click_info_manifesto('serie.png')
    pyautogui.press("backspace")
    pyautogui.write(str(serie))
    pyautogui.sleep(1)
    click_info_manifesto('campo_manifesto.png')
    for i in range(10): 
        pyautogui.press("backspace")
    pyautogui.write(str(manifesto))
    pyautogui.press("tab")
    status_manifesto('status_cadastrado.png')
    click_image('campo_km.png')
    for i in range(10): 
        pyautogui.press("backspace")
    pyautogui.write(str(km))
    click_image('salvar_km.png')
    click_image('ok.png')
    pyautogui.sleep(2)
    click_image('botao_recalculo.png')
    click_image('sim_recalculo.png')
    pyautogui.sleep(3)
    click_image('baixar_manifesto.png')
    click_image('tela_baixar_data.png')
    pyautogui.write(str(um_minuto_atras_formatado))
    pyautogui.sleep(1)
    click_image('ok2.png')
    pyautogui.sleep(1)
    click_image('ok.png')
    pyautogui.sleep(1)
    click_image('nao.png')
    pyautogui.sleep(2)
    click_image('encerrar_sefaz.png')
    click_image('campo_encerramento_numero_manifesto.png')
    pyautogui.write(str(manifesto))
    click_image('botao_encerrar.png')
    pyautogui.sleep(4)
    #verificar_campo('encerrado.png')
    click_image('fechar_tela_encerramento.png')
    pyautogui.sleep(2)

click_image('voltar.png')
# alt_press("f4")
# click_image('fechar_rodopar.png')