import time
import pyautogui
from time import sleep
import pyperclip
from storagefunctions import (pressingKey)
from ProcesesOTPCer import (OTPCERRADA,OTPABIERTA)
import numpy as np
import re
import pydirectinput

def searchAproTelefono(incidentId):
    
    numeros_str = numeros = num_proceso_int = Num_Compara_int = numero_telef_en_pantalla = numero_cabesera = OTH_Configuracion = Num_Compara = num_proceso = crm_dashboard = OTH_Estado_Confi1= OTH_Estado_Confi2 = crm_warning_message = crm_ot_blocked_message = mod_consulta_popup = None 
    crmAttempts = 0
    
    #/////////////////////////////////// CRM DASHBOARD ///////////////////////////////////////////    
    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/AproTelefonica/assets/crm_dashboard.png', grayscale = True,confidence=0.85)
        sleep(0.5)
        if crmAttempts == 5:
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].minimize()
            pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
            print("Estoy dentro de los 5 intentos para ver el Dashboard del CRM")
            crmAttempts = 0
        crmAttempts = crmAttempts + 1
                
    print("CRM Dashboard GUI is present!")   
    pyautogui.click(pyautogui.center(crm_dashboard))    
    
    sleep(0.5)
    pressingKey('f2')
    while mod_consulta_popup is None:        
        mod_consulta_popup = pyautogui.locateOnScreen('C:/AproTelefonica/assets/mod_consulta_popup.png', grayscale = True,confidence=0.85)            
        pressingKey('f2')
    print("mod_consulta_popup field is present and detected on GUI screen!")
    sleep(0.5)
    pyautogui.write(incidentId)
    sleep(0.5)
    pressingKey('enter')
    
    #/////////////////////////////////// FASE DE INGRESO ///////////////////////////////////////////
    
    while crm_ot_blocked_message is None and crm_warning_message is None and crmAttempts < 10:
        crm_warning_message = pyautogui.locateOnScreen('C:/AproTelefonica/assets/mensaje_advertencia.png', grayscale = True,confidence=0.9)   
        crm_ot_blocked_message = pyautogui.locateOnScreen('C:/AproTelefonica/assets/crm_ot_blocked_message.png', grayscale = True,confidence=0.9)   
        sleep(1)
        if crmAttempts == 8:
            print("Estoy dentro de los 8 intentos de medio seg para esperar algun pop up de OT bloqueada inesperado en el CRM")                        
        crmAttempts += 1
        print(crmAttempts)            
            
    # Maximize CRM window 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].maximize()
    print("CRM Edit view Window was maximized!")    
    sleep(1)
    pyautogui.doubleClick(1218,469) 
    time.sleep(1)
    pyautogui.click(679,517)
    time.sleep(1)
    pyautogui.moveTo(1218,469)
    sleep(0.5)
    Intentos = 0
    while OTH_Configuracion is None:
        print("buscando OTH_Configuraciob in screen")
        OTH_Configuracion = pyautogui.locateOnScreen('C:/AproTelefonica/assets/OTH_Configuracion.png', grayscale = True,confidence=0.95)  
        Intentos += 1                              
        if OTH_Configuracion is not None:
            print("OTH_Configracion is present!")
            pyautogui.moveTo(pyautogui.center(OTH_Configuracion))            
            pyautogui.doubleClick()
            sleep(1) 
            break
        if Intentos == 10:
            print("OTH_Configuracion no se encuebtra")
            sleep(0.5)
            return 14
    # Maximize CRM OTH window 
    pyautogui.getWindowsWithTitle("Tarea hija")[0].maximize()
    print("CRM OTH CONFIGURACION view Window was maximized!")

    #//////////PRUEBA NUMERO DETALLE\\\\\\\\\\\\\\\
    # mover a detalles parte de abjo
    pyautogui.click(523,696)
    # Coordenadas del punto en el que se hará clic inicialmente
    click_x = 515
    click_y = 637
    pyautogui.moveTo(523,696)
    pyautogui.moveRel(523,696)  

    # Establecer el número máximo de intentos
    max_Intestonum = 100
    Intestonum = 0
    # Bucle para buscar la imagen
    while Intestonum < max_Intestonum:
        print("Intento:", Intestonum)
        # Hacer clic en la parte definida de la pantalla
        pyautogui.click(click_x, click_y)
        #time.sleep(0.5)  # Esperar un poco después del clic

        # Buscar la imagen que contiene la palabra "numero"
        numero_cabesera = pyautogui.locateOnScreen("C:/AproTelefonica/assets/numero_cabecera.png", grayscale=True, confidence=0.97)
        numero_telef_en_pantalla = pyautogui.locateOnScreen("C:/AproTelefonica/assets/numero_con_tilde.png", grayscale=True, confidence=0.97)

        if numero_cabesera is not None:
            # Si se encuentra la imagen, realizar un doble clic en el centro
            pyautogui.moveTo(pyautogui.center(numero_cabesera))
            pyautogui.doubleClick()
            sleep(0.5)
            # Presionar las teclas Control, Shift y Down al mismo tiempo
            pydirectinput.keyDown('ctrl')
            pydirectinput.keyDown('shift')
            pydirectinput.keyDown('down')
            #   Esperar un segundo para mantener las teclas presionadas
            time.sleep(1)
            # Liberar las teclas Control, Shift y Down
            pydirectinput.keyUp('ctrl')
            pydirectinput.keyUp('shift')
            pydirectinput.keyUp('down')
            sleep(0.5)
            pyautogui.hotkey('ctrl','c')
            sleep(1)
            # Obtener el texto del portapapeles
            Num_Compara = pyperclip.paste()
            sleep(0.5)
            numeros = re.findall(r'\d+', Num_Compara)
            # Convertir la lista de números a una cadena
            numeros_str = ', '.join(numeros)
            # Imprimir la cadena de números
            print('Número encontrado:', numeros_str)
            sleep(1)
            print("Número encontrado! Almacenado")
             # Cierre ventana OTH Configuracion
            pyautogui.getWindowsWithTitle("Tarea hija")[0].close()        
            sleep(2)
            break  # Salir del bucle una vez que se encuentra y guarda la variable volatil

        elif numero_telef_en_pantalla is not None:
            # Si se encuentra la imagen, realizar un doble clic en el centro
            pyautogui.moveTo(pyautogui.center(numero_telef_en_pantalla))
            pyautogui.doubleClick()
            sleep(0.5)
            # Presionar las teclas Control, Shift y Down al mismo tiempo
            pydirectinput.keyDown('ctrl')
            pydirectinput.keyDown('shift')
            pydirectinput.keyDown('down')
            #   Esperar un segundo para mantener las teclas presionadas
            time.sleep(1)
            # Liberar las teclas Control, Shift y Down
            pydirectinput.keyUp('ctrl')
            pydirectinput.keyUp('shift')
            pydirectinput.keyUp('down')
            sleep(0.5)
            pyautogui.hotkey('ctrl','c')
            sleep(1)
            # Obtener el texto del portapapeles
            Num_Compara = pyperclip.paste()
            sleep(0.5)
            numeros = re.findall(r'\d+', Num_Compara)
            # Convertir la lista de números a una cadena
            numeros_str = ', '.join(numeros)
            # Imprimir la cadena de números
            print('Número encontrado:', numeros_str)
            sleep(1)
            print("Número encontrado! Almacenado")
             # Cierre ventana OTH Configuracion
            pyautogui.getWindowsWithTitle("Tarea hija")[0].close()        
            sleep(2)
            break 
        # Hacer scroll hacia abajo de forma rápida
        pyautogui.scroll(100)  # El valor negativo indica desplazamiento hacia abajo rápido
        # Incrementar el contador de intentos
        Intestonum += 1 
        if Intestonum == max_Intestonum:
            print("No se pudo encontrar el número después de", max_Intestonum, "intentos.")
            print('Numero no encontrado revisar de forma manual')
            return 13
    sleep (1)

    # Bucle para buscar otros estados
    while OTH_Estado_Confi1 is None or OTH_Estado_Confi2 is None:
        print("Buscando Estado en la pantalla...")
        OTH_Estado_Confi2 = pyautogui.locateOnScreen('C:/AproTelefonica/assets/Estado_Cerrada.png', grayscale=True, confidence=0.97)
        OTH_Estado_Confi1 = pyautogui.locateOnScreen('C:/AproTelefonica/assets/Estado_Cancelado.png', grayscale=True, confidence=0.97)
        
        if OTH_Estado_Confi2 is not None:
            print('Incio de proceso para OTP estado Cerrado')
            sleep(0.5)
            num_proceso = OTPCERRADA()
            sleep(0.5)
            break
        elif OTH_Estado_Confi1 is not None:
            print("Ot cancelada")
            sleep(0.5)
            return 11
        else:
            print('Incio de proceso para OTP estado abierta')
            sleep(0.5)
            num_proceso = OTPABIERTA()
            print('Finalizado proceso para OTP estado Cerrado')
            pyautogui.click(990,10)
            sleep(0.5)
            break
        
    # Realizar un proceso si se encuentra cualquiera de los otros estados
    sleep(0.5)
    print("Estado encntrado, numeros almacenados")
    pyautogui.click(1119,21)
    sleep(1)
    # Close Edit Incident View 
    pyautogui.getWindowsWithTitle("Ordenes de Trabajo v8")[0].close()        
    sleep(1)

    num_proceso_int = int(num_proceso)
    Num_Compara_int = int(numeros_str)
    print("Numero encontrado", Num_Compara_int)
    print("Numero Extraido", num_proceso_int)
    sleep(1)
    if np.array_equal(Num_Compara_int, num_proceso_int):
        print("Número correcto")
        sleep(0.5)
        return 0
    else:
        print("Número no conside")
        sleep(0.5)
        return 12