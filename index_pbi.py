

from openpyxl.worksheet.table import Table
from typing import Container, final
import shutil
from openpyxl.worksheet.table import TableStyleInfo
from pandas.core.frame import DataFrame
from selenium import webdriver
import selenium
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys
from time import process_time, sleep
from selenium.webdriver.common.action_chains import ActionChains
import os
from datetime import date
from datetime import datetime
from datetime import timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import openpyxl as xl
from openpyxl import load_workbook
import warnings
from selenium.webdriver.chrome.options import Options
from subprocess import call
import win32com.client
from win32com.client import Dispatch
import ctypes

from variables import *


user32 = ctypes.windll.user32
screensize = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
ancho = user32.GetSystemMetrics(0)
largo = user32.GetSystemMetrics(1)
print("-------------------")
print("La resolucion de tu monitor es de")
print(ancho)
print(largo)
print("--------------------")
chrome_options_saa = Options()
chrome_options_saa.add_argument("--headless")
chrome_options_saa.add_argument("--no-sandbox")
chrome_options_saa.add_argument("--disable-gpu")
chrome_options_saa.add_argument(f"--window-size={ancho},{largo}")

class Actualiza_Reporte:
    def __init__(self , segundo=0):
        ## Ejecutar en primer o segundo plano
        if segundo == 0:
            self.driver = webdriver.Chrome(executable_path=ruta_usuario_drivers)

        if segundo== 1:
            self.driver = webdriver.Chrome( chrome_options=chrome_options_saa , executable_path=r"C:\driver chrome\chromedriver")
    

    # Busca un elemento por XPATH, si usar WebdriverWait
    def Encuentra_y_click(self , xpath , tipo=0):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                #driver= webdriver.Chrome(executable_path=ruta)
                i = i +1
                elemento = self.driver.find_element_by_xpath(xpath)
                elemento.click()
            except:
                self.errores(tipo)            
            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a=False
                    exit()
    
    ## Encuentra un elemento y escribe, sin usar WebdriverWait
    def Encuentra_y_escribe(self, xpath, texto, tipo=0):
        j=0
        b=True
        while b==True:
            j=j
            try:
                j=j+1
                elemento = self.driver.find_element_by_xpath(xpath)
                elemento.send_keys(texto)

            except Exception as e:
                print(type(e).__name__)
                self.errores(tipo)
            else:
                break
            finally:
                print(f"Intento {j}")
                if j==5:
                    b=False
                    exit()
    

    ## Uno de los mas utilizados. Utiliza el Webdriver para esperar hasta que el elemento este en pantalla y luego le da click
    ## Reg ruta = Ruta donde esta el archivo para ingresar el registro de ERRORES
    ## File = Nombre del archivo donde se lleva el registro de Errores
    def espera_y_click(self , xpath , tiempo=10):
        k=0
        c= True
        while c==True:
            k=k
            try:
                k = k +1
                elemento= WebDriverWait(self.driver, tiempo).until(EC.presence_of_element_located((By.XPATH , xpath)))
                elemento.click()

            except Exception as e:
                print(type(e).__name__)
                self.driver.refresh()
            else:
                break
            finally:
                print(f"Intento {k}")
                if k==5:
                    c=False
                    self.driver.quit()
                    exit()

    
    def espera_y_escribe(self, xpath, texto,  des = "Sin descripcion" , tiempo = 10):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i = i +1
                elemento= WebDriverWait(self.driver, tiempo).until(EC.presence_of_element_located((By.XPATH , xpath)))
                elemento.send_keys(texto)
            except Exception as e:
                print(type(e).__name__)
                print(f"MENSAJE Error en la funcion espera y escribe con el xpath: {xpath}. {des}")
                self.driver.refresh()

            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a=False
                    self.driver.quit()
                    exit()


    def espera_y_click_temporizado(self , xpath , tiempo, tipo=0):
        k=0
        c= True
        while c==True:
            k=k
            try:
                k = k +1
                elemento= WebDriverWait(self.driver,tiempo).until(EC.presence_of_element_located((By.XPATH , xpath)))
                elemento.click()
            except:
                self.errores(tipo)
            else:
                break
            finally:
                print(f"Intento {k}")
                if k==5:
                    c=False
                    exit()

    def espera_y_escribe_temporizado(self, xpath, texto, tiempo, tipo=0):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i = i +1
                elemento= WebDriverWait(self.driver, tiempo).until(EC.presence_of_element_located((By.XPATH , xpath)))
                elemento.send_keys(texto)
            except:
                self.errores(0)
            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a=False
                    exit()

    def esperar_y_moverse(self , xpath, tiempo = 10 ):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i = i + a
                elemento = WebDriverWait(self.driver , tiempo).until(EC.presence_of_element_located((By.XPATH , xpath)))
                ActionChains(self.driver).move_to_element(elemento).perform()
            except Exception as e:
                self.driver.refresh()
            else:
                break
            finally:
                print(f"Inteto {i}")
                if i==5:
                    a = False
                    exit()

    def esperar_y_moverse_clase(self , clase , tiempo = 10 ):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i = i + a
                elemento = WebDriverWait(self.driver , tiempo).until(EC.presence_of_element_located((By.CLASS_NAME , clase)))
                ActionChains(self.driver).move_to_element(elemento).perform()
            except Exception as e:
                self.driver.refresh()
                
            else:
                break
            finally:
                print(f"Inteto {i}")
                if i==5:
                    a = False
                    exit()

    def esperar_y_moverse_text(self , text, tiempo = 10 ):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i = i + a
                elemento = WebDriverWait(self.driver , tiempo).until(EC.presence_of_element_located((By.LINK_TEXT, text)))
                ActionChains(self.driver).move_to_element(elemento).perform()
            except Exception as e:
                self.driver.refresh()
                
            else:
                break
            finally:
                print(f"Inteto {i}")
                if i==5:
                    a = False
                    exit()

    def esperar_moverse_y_click(self, xpath , tiempo = 10):
        i = 0
        a = True
        while a==True:
            i= i 
            try:
                i = i +1
                elemento = WebDriverWait(self.driver , tiempo).until(EC.element_located_to_be_selected((By.XPATH , xpath)))
                ActionChains(self.driver).move_to_element(elemento).click().perform()
            except Exception as e:
                print(type(e).__name__)
                self.driver.refresh()

            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a = False
                    exit()
    
    def esperar_moverse_y_dobleclick(self , xpath , tiempo = 10 , tipo=0):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i= i +1
                elemento = WebDriverWait(self.driver, tiempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
                ActionChains(self.driver).move_to_element(elemento).double_click().perform()
            except:
                self.errores(tipo)
            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a= False
                    exit()

    def cambio_pestana_1(self , tiempo = 3 , tipo = 0):
        i=0 
        a = True
        while a==True:
            i = i
            try:
                i = i +1 
                self.driver.implicitly_wait(tiempo)
                # self.driver.switch_to_window(self.driver.window_handles[1])
                self.driver.switch_to_window(self.driver.window_handles[1])
            except:
                self.driver.refresh()
            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a= False
                    exit()

    def cambia_pestana_base(self , tiempo=3 , tipo = 0):
        i = 0
        a = True
        while a == True:
            i = i
            try:
                i = i +1
                self.driver.implicitly_wait(tiempo)
                self.driver.switch_to_window(self.driver.window_handles[0])
            except:
                self.errores(tipo)
            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a = False
                    exit()
    
    def cambio_iframe(self, iframe , tiempo = 3 , tipo =0 ):
        i = 0
        a = True
        while a==True:
            i = i
            try:
                i = i + 1 
                self.driver.implicitly_wait(tiempo)
                self.driver.switch_to.frame(iframe)
            except:
                self.errores(tipo)
            else:
                break
            finally:
                print(f"Intento {i}")
                if i==5:
                    a = False
                    exit()
    
    def cambio_url(self, url, tiempo = 2 , tipo = 0):
        i = 0
        a = True
        while a == True:
            i = i 
            try:
                i = i +1

                self.driver

                self.driver.implicitly_wait(tiempo)
                self.driver.get(url)
            except:
                self.errores(0)
            else:
                break
            finally:
                print(f"Intento {i}")
                if i ==5:
                    a = False
                    self.driver.quit()
                    exit()
        
    def refrescar(self , tiempo  = 2 , tipo = 0):
        i = 0
        a = True
        while a==True:
            i = i 
            try:
                i = i +1
                self.driver.implicitly_wait(tiempo)
                self.driver.refresh()
            except:
                self.errores(tipo)
            else:
                break
            finally:
                print(f"Intento {i}")
                if i ==5:
                    a = False
                    self.driver.quit()
                    exit()

    def espera_y_borra(self, xpath , espacios = 50 , tiempo= 10 , tipo = 0):
        i = 0
        a = True
        while a == True:
            i = i 
            try:
                i = i +1 
                elemento = WebDriverWait(self.driver , tiempo).until(EC.presence_of_element_located((By.XPATH , xpath)))
                ## Iterador para borrar
                j = 0
                while j < espacios:
                    elemento.send_keys(Keys.BACK_SPACE)
                    j = j +1 
            except:
                elemento2= WebDriverWait(self.driver,tiempo).until(EC.presence_of_element_located((By.XPATH , '//*[@id="bienvenida-step"]/div[3]/ul/li[3]/a')))
                elemento2.click()
            else:
                break
            finally:
                print(f"Intento {i}")
                if i == 5:
                    a = False
                    self.driver.quit()
                    exit()

    def click_f(self , tiempo = 0.5):
        sleep(tiempo)
        ActionChains(self.driver).click().perform()

    def doble_click_f(self , tiempo = 0.5):
        sleep(tiempo)
        ActionChains(self.driver).double_click().perform()

    ## Cierra navegador 
    def cierra(self , tiempo = 4):
        self.driver.implicitly_wait(tiempo)
        self.driver.quit()

    def espera_y_cierra(self , tiempo = 3):
        sleep(tiempo)
        self.driver.quit()

    ## Esta funcion da tiempo a cargar a la pagina, cuando te cambias de url
    def espera_que_cargue(self , xpath, tiempo = 30):
        WebDriverWait(self.driver,tiempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
        print(f"El elemento {xpath} esta cargado")
    
    def espera_que_cargue_text(self , texto, tiempo = 30):
        WebDriverWait(self.driver,tiempo).until(EC.presence_of_element_located((By.LINK_TEXT, texto)))
        print(f"El elemento {texto} esta cargado")