import time
import pyautogui
from time import sleep
import pyperclip
import os
import subprocess
import pyperclip
import pyautogui
from PIL import Image
import numpy as np
import pytesseract
from dotenv import load_dotenv
from storagefunctions import (pressingKey,make_window_visible)

def OTPCERRADA():
    numero_extraido = button_Telf = VentanaTelef = PantallaTelefonia = None 

    while button_Telf is None:
        button_Telf = pyautogui.locateOnScreen('C:/AproTelefonica/assets/BotonTelefonia.png', grayscale = True,confidence=0.9)   
    print("CRM OTH Save Incident button is present!")
    button_Telf_x,button_Telf_y = pyautogui.center(button_Telf)
    pyautogui.click(button_Telf_x, button_Telf_y)   
    sleep(1)
    while VentanaTelef is None:
        print('Buscando ventana de decicion de telefonia')
        VentanaTelef = pyautogui.locateOnScreen("C:/AproTelefonica/assets/VentanaTelef.png", grayscale=True, confidence=0.9)
        sleep(0.5)
    pyautogui.press('enter')
    print('Ventana decision visible')
    sleep(1)
    while PantallaTelefonia is None:
        print('Buscando Pantalla Telefono')  
        PantallaTelefonia = pyautogui.locateOnScreen("C:/AproTelefonica/assets/pantallatelefonia.png", grayscale=True, confidence=0.9)
        sleep(0.5)
        if PantallaTelefonia is not None:
            print('Ventana telefonia is visible')
            break
    sleep(1)
    pyautogui.doubleClick(92,431)
    sleep(1)
    print('apertura de ventana tbcl')
    sleep(2)
    # Configurar la ruta al ejecutable de Tesseract OCR
    pytesseract.pytesseract.tesseract_cmd = r'C:\AproTelefonica/Tesseract-OCR/tesseract.exe'

    # Definir las coordenadas de la zona seleccionada
    x1, y1 = 248, 250
    x2, y2 = 311, 263

    # Tomar una captura de pantalla de la zona seleccionada
    screenshot = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))

    # Guardar la captura de pantalla como imagen
    screenshot.save("screenshot.png")

    # Utilizar OCR para extraer texto de la imagen
    texto_extraido = pytesseract.image_to_string(screenshot)

    # Imprimir el texto extraído
    print("Texto extraído de la imagen:", texto_extraido)
    sleep(0.5)
    numero_extraido = texto_extraido[3:]  # Eliminar los tres primeros dígitos
    # Imprimir el número sin los tres primeros dígitos
    print("Número sin los tres primeros dígitos:", numero_extraido)
    sleep(1)
    return numero_extraido
    
def OTPABIERTA():
    edit_asunto_Telf = edit_informe_Telf = numero_extraido = button_Telf = VentanaTelef = PantallaTelefonia = None 

    while button_Telf is None:
        button_Telf = pyautogui.locateOnScreen('C:/AproTelefonica/assets/BotonTelefonia.png', grayscale = True,confidence=0.9)   
    print("CRM OTH Save Incident button is present!")
    button_Telf_x,button_Telf_y = pyautogui.center(button_Telf)
    pyautogui.click(button_Telf_x, button_Telf_y)   
    sleep(1)
    while VentanaTelef is None:
        print('Buscando ventana de decicion de telefonia')
        VentanaTelef = pyautogui.locateOnScreen("C:/AproTelefonica/assets/VentanaTelef.png", grayscale=True, confidence=0.9)
        sleep(0.5)
    pyautogui.press('enter')
    print('Ventana decision visible')
    sleep(1)
    while PantallaTelefonia is None:
        print('Buscando Pantalla Telefono')  
        PantallaTelefonia = pyautogui.locateOnScreen("C:/AproTelefonica/assets/pantallatelefonia.png", grayscale=True, confidence=0.9)
        sleep(0.5)
        if PantallaTelefonia is not None:
            print('Ventana telefonia is visible')
            break
    sleep(1)
    pyautogui.doubleClick(386,579)
    print('Novedad Encontrada')
    sleep(0.5)
    while edit_asunto_Telf is None:
        edit_asunto_Telf = pyautogui.locateOnScreen('C:/AproTelefonica/assets/Editar_asunto_tel.png', grayscale = True,confidence=0.9)   
    print("boton oprimido editar asunto de telefono!")
    edit_asunto_Telf_x,edit_asunto_Telf_y = pyautogui.center(edit_asunto_Telf)
    pyautogui.click(edit_asunto_Telf_x, edit_asunto_Telf_y)   
    sleep(2)
    # /////////buscar boton editar informe telefoni///////////
    while edit_informe_Telf is None:
        edit_informe_Telf = pyautogui.locateOnScreen('C:/AproTelefonica/assets/edit_informe_tel.png', grayscale = True,confidence=0.9)   
    print("boton oprimido info de telefonia!")
    edit_informe_Telf_x,edit_informe_Telf_y = pyautogui.center(edit_informe_Telf)
    pyautogui.click(edit_informe_Telf_x, edit_informe_Telf_y)   
    sleep(3)

    pyautogui.click(298,148)
    print("numeracion publica visible")
    sleep(1)
    
    # Configurar la ruta al ejecutable de Tesseract OCR
    pytesseract.pytesseract.tesseract_cmd = r'C:\AproTelefonica/Tesseract-OCR/tesseract.exe'

    # Definir las coordenadas de la zona seleccionada
    x1, y1 = 216, 497
    x2, y2 = 280, 509

    # Tomar una captura de pantalla de la zona seleccionada
    screenshot = pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1))

    # Guardar la captura de pantalla como imagen
    screenshot.save("screenshot.png")

    # Utilizar OCR para extraer texto de la imagen
    texto_extraido = pytesseract.image_to_string(screenshot)

    # Imprimir el texto extraído
    print("Texto extraído de la imagen:", texto_extraido)
    sleep(0.5)
    numero_extraido = texto_extraido[3:]  # Eliminar los tres primeros dígitos
    # Imprimir el número sin los tres primeros dígitos
    print("Número sin los tres primeros dígitos:", numero_extraido)
    sleep(1)
    return numero_extraido