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
        image_path2 = os.path.join(current_dir, caminho_imagem, "query_timeout_expered.png") 
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                print("Imagem de query_timeout_expered foi encontrada na tela.")
                click_image('ok.png')
        except Exception as e:
            print("Aguardando...")
        pyautogui.sleep(1)

# click_image('botao_faturamento.png')
# click_image('botao_faturamento_movimentacao.png')
# click_image('botao_faturamento_movimentacao_manifesto.png')
# click_image('botao_faturamento_movimentacao_manifesto_emissao.png')

click_image('botao_faturamento.png')
click_image('botao_faturamento_movimentacao.png')
click_image('botao_faturamento_movimentacao_cancelamento.png')

