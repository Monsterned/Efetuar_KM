import pandas as pd
from openpyxl import load_workbook
import os
import numpy as np
import pyautogui
import ctypes
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys          
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
import pyautogui
from selenium.webdriver.common.action_chains import ActionChains
import shutil
from datetime import datetime, timedelta


def status_manifesto(image_path,image_path3,image_path4, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    image_path4 = os.path.join(current_dir, caminho_imagem, image_path4)
    nao_encontrado = False
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                print("Campo encontrado cadastrado.")
                aviso_ativo = aviso_vencimento('curso_vencido.png')
                return aviso_ativo,nao_encontrado
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position3:
                print("Campo encontrado emitido.")
                aviso_ativo = aviso_vencimento('curso_vencido.png')
                return aviso_ativo,nao_encontrado
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        image_name2 = 'status_baixado.png'
        image_path2 = os.path.join(current_dir, caminho_imagem, image_name2)
        try:
            position4 = pyautogui.locateOnScreen(image_path4, confidence=confidence)
            if position4:
                print("Manifesto nao localizado.")
                nao_encontrado = True
                caminho_do_arquivo = 'BASE_KM.xlsx'
                nome_da_aba = 'Planilha1'
                wb = load_workbook(caminho_do_arquivo)
                ws = wb[nome_da_aba]
                coluna_ost = 'E'  
                if linha > ws.max_row:
                    ws[coluna_ost + str(linha_especifica)] = 'NAO LOCALIZADO'
                else:
                    ws[coluna_ost + str(linha_especifica)] = 'NAO LOCALIZADO'
                wb.save(caminho_do_arquivo)
                wb.close()
                return nao_encontrado
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        try:
            print("Tentando achar a segunda imagem.")
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                print("Campo de baixado encontrado.")
                aviso_ativo = aviso_vencimento('curso_vencido.png')
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
                for i in range(2):
                    pyautogui.press("enter")
                pyautogui.sleep(1)
                click_image('janela.png')
                click_image('tela_cancelamento.png')
                click_image('confirmacao_tela_cancelamento.png')
                click_image('incluir.png')
                pyautogui.sleep(2)
                verificar_campo_cancelamento('status_tela_cancelamento.png')
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
                pyautogui.write("correcao de km")
                pyautogui.sleep(1)
                click_image('salvar_cancelamento.png')
                click_image('sim_cancelamento.png')
                click_image('ok_cancelamento.png')
                click_image('voltar.png')
                print('cancelado')
                pyautogui.sleep(2)
                click_image('incluir.png')
                pyautogui.sleep(2)
                novo_lançamento('campo_numero_manifesto.png')
                pyautogui.sleep(1)
                click_info_manifesto('filial.png')
                pyautogui.press("backspace")
                pyautogui.write(str(filial))
                pyautogui.sleep(1)
                pyautogui.press("tab")
                click_info_manifesto('serie.png')
                pyautogui.press("backspace")
                pyautogui.write(str(serie))
                pyautogui.sleep(1)
                pyautogui.press("tab")
                click_info_manifesto('campo_manifesto.png')
                for i in range(10): 
                    pyautogui.press("backspace")
                for i in range(10): 
                    pyautogui.press("del")
                pyautogui.sleep(0.5)
                print(manifesto)
                pyautogui.write(str(manifesto))
                pyautogui.sleep(1)
                pyautogui.press("tab")
                status_manifesto('status_cadastrado.png','status_emitido.png')
                return aviso_ativo,nao_encontrado
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")

        pyautogui.sleep(1)

nao_encontrado, aviso_ativo = status_manifesto('status_cadastrado.png','status_emitido.png','campo_numero_manifesto.png')