#Recursos
import os.path
import time
import warnings
import traceback
import pickle
import pandas as pd

from datetime import datetime
from pickle import FALSE, TRUE
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


#Ignorar DeprecationWarning
warnings.filterwarnings("ignore", category=DeprecationWarning) 

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly','https://www.googleapis.com/auth/spreadsheets']


creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            '/Users/keilamarin/Documents/Reserve/Bot_BOD/creds.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

#API services
gmail = build('gmail', 'v1', credentials=creds)
sheets = build('sheets','v4', credentials=creds)

def bot():
    #Call de chromedriver
    CHROMEDRIVER_PATH = '/usr/local/bin/chromedriver'
    WINDOW_SIZE = "1920,1080"
    
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument("--test-type")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-first-run")
    chrome_options.add_argument("--no-default-browser-check")
    chrome_options.add_argument("--ignore-certificate-errors")
    
    driver = webdriver.Chrome(executable_path=CHROMEDRIVER_PATH,chrome_options=chrome_options)

    #Definir ActionChains
    action = ActionChains(driver)

    #Inicio de sesion DESK
    driver.get("https://desk.reserve.org/login")
    

    #Credenciales
    usuario = "vastoperador4@gmail.com"
    clave = "Roma1234"

    #Login//
    #Usuario
    try:
        usertxt = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div/div/div/div/div/form/div[1]/div/div/input'))
        )
        usertxt.send_keys(usuario)
    except:
        driver.quit()

    #Clave
    passtxt = driver.find_element_by_xpath('//*[@id="root"]/div/main/div/div/div/div/div/form/div[2]/div/div/input')
    passtxt.send_keys(clave)

    continuar = driver.find_element_by_xpath('//*[@id="root"]/div/main/div/div/div/div/div/form/div[3]/button')
    continuar.click()
    
    #Ir a VOTC
    try:
        panel = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//a[@href="/votc_board"]'))
        )
        panel.click()
        print("Sesión iniciada en Desk")
    except:
        driver.quit
        print("No se logró iniciar sesión en Desk, reiniciar bot")
        
    #Inicio de sesion BOD
    time.sleep(3)
    driver.execute_script("window.open('https://web.bancadigitalbod.com/nblee6/f/ext/Login/index.xhtml', 'new window')")
    driver.switch_to.window(driver.window_handles[1])

    #Credenciales
    spreadsheet_id = '1BVnD6shewsbUDbhsxoNHa78ndSb9Djb0Tog8lsC8tQc'

    credenciales = sheets.spreadsheets().values().get(
        spreadsheetId = spreadsheet_id,
        majorDimension = 'COLUMNS',
        range = 'Credenciales!B4:B7'
    ).execute()

    tipo = credenciales['values'][0][0]
    id = credenciales['values'][0][1]
    usuario = credenciales['values'][0][2]
    clave = credenciales['values'][0][3]

    def bod_user():
        #Login//
        #Identificación
        try:
            tiposelect = WebDriverWait(driver,20).until(
                EC.presence_of_element_located((By.XPATH,'//select[@id="form:selectNumIdCli"]'))
            )
            tiposelect.send_keys(tipo)
        except:
            driver.quit()

        idtxt = driver.find_element_by_xpath('//input[@id="form:txtNumIdCli"]')
        idtxt.send_keys(id)

        continuar1 = driver.find_element_by_xpath('//a[@id="form:validarLogin"]')
        continuar1.click()

        #Usuario
        try:
            usertxt = WebDriverWait(driver,20).until(
                EC.presence_of_element_located((By.XPATH,'//input[@id="form:nombre"]'))
            )
            usertxt.send_keys(usuario)
        except:
            driver.quit()

        continuar2 = driver.find_element_by_xpath('//a[@id="form:siguiente"]')
        continuar2.click()
    bod_user()

    #Error ocasional
    try:
        ocassion = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="formError:dialogSesionActiveButtonNo"]'))
        )
        ocassion.click()
        bod_user()
    except:
        driver.quit

    #Preguntas de seguridad
    preguntas = sheets.spreadsheets().values().get(
        spreadsheetId = spreadsheet_id,
        majorDimension = 'ROWS',
        range = 'Credenciales!A10:B15'
    ).execute()

    #Respuestas
    try:
        respuesta = WebDriverWait(driver,20).until(
            EC.presence_of_element_located((By.XPATH,'//input[@id="j_idt13:j_idt14:form-psec:j_idt17:0:respuestaPregunta"]'))
        )
        pregunta = driver.find_element_by_xpath('//*[@id="j_idt13:j_idt14:form-psec"]/div[1]/p[1]').text
        
        if preguntas['values'][0][0] in str(pregunta):
            respuesta.send_keys(preguntas['values'][0][1])
        elif preguntas['values'][1][0] in str(pregunta):
            respuesta.send_keys(preguntas['values'][1][1])
        elif preguntas['values'][2][0] in str(pregunta):
            respuesta.send_keys(preguntas['values'][2][1])
        elif preguntas['values'][3][0] in str(pregunta):
            respuesta.send_keys(preguntas['values'][3][1])
        elif preguntas['values'][4][0] in str(pregunta):
            respuesta.send_keys(preguntas['values'][4][1])
        else:
            respuesta.send_keys(preguntas['values'][5][1])
    except:
        driver.quit
        print("Inicio de sesion fallido, reiniciar bot")
    try:
        respuesta2 = WebDriverWait(driver,5).until(
            EC.presence_of_element_located((By.XPATH,'//input[@id="j_idt13:j_idt14:form-psec:j_idt17:1:respuestaPregunta"]'))
        )
        pregunta2 = driver.find_element_by_xpath('//*[@id="j_idt13:j_idt14:form-psec"]/div[2]/p[1]').text
        
        if preguntas['values'][0][0] in str(pregunta2):
            respuesta2.send_keys(preguntas['values'][0][1])
        elif preguntas['values'][1][0] in str(pregunta2):
            respuesta2.send_keys(preguntas['values'][1][1])
        elif preguntas['values'][2][0] in str(pregunta2):
            respuesta2.send_keys(preguntas['values'][2][1])
        elif preguntas['values'][3][0] in str(pregunta2):
            respuesta2.send_keys(preguntas['values'][3][1])
        elif preguntas['values'][4][0] in str(pregunta2):
            respuesta2.send_keys(preguntas['values'][4][1])
        else:
            respuesta2.send_keys(preguntas['values'][5][1])
    except:
        continuar3 = driver.find_element_by_xpath('//a[@id="j_idt13:j_idt14:form-psec:siguiente"]')
        continuar3.click()

    continuar3 = driver.find_element_by_xpath('//a[@id="j_idt13:j_idt14:form-psec:siguiente"]')
    continuar3.click()

    continuar4 = driver.find_element_by_xpath('//*[@id="j_idt12:form"]/div/div[2]/a[2]')
    continuar4.click()

    #Clave
    try:
        passtxt = WebDriverWait(driver,20).until(
            EC.presence_of_element_located((By.XPATH,'//input[@id="form:contrasena"]'))
        )
        passtxt.send_keys(clave)
    except:
        driver.quit()
        print("Inicio de sesion fallido, reiniciar bot")

    continuar5 = driver.find_element_by_xpath('//*[@id="form:siguiente"]')
    continuar5.click()
    print("Iniciada sesion en banco")
    
    #Busqueda de transacciones tomadas
    driver.switch_to.window(driver.window_handles[0])
    print("Buscando transacciones")

    while TRUE:
        
        #Transacciones tomadas
        filas = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
        num_filas = len(filas)
        print("Transacciones en cola", num_filas/2)
        
        if num_filas > 0:

            #Empezar a medir tiempo de operacion
            begin_time = datetime.now()

            #Busqueda de transacciones tomadas
            time.sleep(1)

            #Contar el numero de transacciones
            filas = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
            num_filas = len(filas)
            ultima = num_filas - 1

            #Recoger datos de transacción más antigua BOD
            drop_down = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima)+']/td[8]/button').click()

            #Volver a contar transacciones
            filas = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
                
            #Datos de primera fila
            tipo = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima)+']/td[2]/div').text
            banco = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima)+']/td[3]').text
            print(tipo,banco)
            
            #Procesar Salida
            if banco == 'BOD' and tipo == 'Sell':

                try:
                    #Datos de segunda fila
                    trans_id = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[1]/div/p').text
                    nombre_reserve = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[2]/div/p').text
                    usuario_reserve = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[3]/div/p').text
                    rsv = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[4]/div/p').text
                    tasa = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[2]/div[4]/div/p').text
                    telefono = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[3]/div/div/p').text
                    beneficiario = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div[1]/div/p').text
                    persona = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div[2]/div/p').text
                    identificacion = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div[3]/div/p').text
                    tipo_cuenta = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div[4]/div/p').text
                    num_cuenta = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div[5]/div/p').text
                    monto1 = round(float(rsv)*float(tasa),2)
                    confirmacion_salida = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/input').get_attribute("class")

                    print(tipo, banco, trans_id, nombre_reserve, tipo_cuenta, num_cuenta)
                except:
                    driver.quit()

                #Ir a transferencias
                driver.switch_to.window(driver.window_handles[1])
                    
                #Verificar inactividad en banco
                try:
                    driver.switch_to.window(driver.window_handles[1])
                    keepalive = WebDriverWait(driver,7).until(
                            EC.visibility_of_element_located((By.XPATH,'//*[@href="javascript:keepAliveSession()"]'))
                    )
                    action.move_to_element
                    keepalive.click()
                except:
                    pass

                try:
                    menu = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="supercontenedor"]/div[1]'))
                    )
                    menu.click()
                except:
                        driver.quit

                transfers = driver.find_element_by_xpath('//*[@id="ico-menu-12"]')
                transfers.click()

                transfers2 = driver.find_element_by_xpath('//*[@id="submenu-ppal-12"]/ul/a[1]/li')
                transfers2.click()

                #Agregar un nuevo beneficiario
                try:
                    add = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm"]/div[6]/a'))
                    )
                    add.click()
                except:
                    driver.quit

                for t in range(2):
                    error = None
                    if t == 2:
                        s = 1
                    else:
                        s = 20
                    try:
                        otp = WebDriverWait(driver,5).until(
                            EC.presence_of_element_located((By.XPATH,'//*[@class="ancho100 zonaSeguraInput"]'))
                        )
                        time.sleep(s)
                        # Call the Gmail API
                        results = gmail.users().labels().list(userId='me').execute()
                        labels = results.get('labels', [])

                        #Get Messages
                        results = gmail.users().messages().list(userId='me', labelIds=['INBOX']).execute()
                        messages = results.get('messages', [])

                        message_count = 1
                        for message in messages[:message_count]:
                            msg = gmail.users().messages().get(userId='me', id=message['id']).execute()
                            email = str(msg['snippet']).split()
                            pre_codigo = email[9]
                            codigo = pre_codigo.split(",")
                            codigo2 = str(codigo[0])

                        otp.send_keys(codigo2)

                        continuar_otp = driver.find_element_by_xpath('//a[@class="bod-button buttonCodigosOtp button-continuar"]')
                        continuar_otp.click()

                        benf_tipo = WebDriverWait(driver,10).until(
                            EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:select"]'))
                        )
                        t = t+1
                    except TimeoutException as e:
                        error = e
                    if error is None:
                        break
                if error is not None:
                    driver.quit

                benf_tipo = WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:select"]'))
                )
                benf_tipo.send_keys(persona)
                time.sleep(1)
                benf_id = WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:numeroIdentificacion"]'))
                )
                benf_id.send_keys(int(identificacion))
                time.sleep(1)
                benf_name =  WebDriverWait(driver,10).until(
                    EC.visibility_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:nombreCompleto"]'))
                )
                benf_name.send_keys(identificacion)

                benf_alias =  WebDriverWait(driver,10).until(
                    EC.visibility_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:alias"]'))
                )
                benf_alias.send_keys(identificacion)

                benf_cuenta =  WebDriverWait(driver,10).until(
                    EC.visibility_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:numeroProducto"]'))
                )
                benf_cuenta.send_keys(num_cuenta)

                continuar_sin_guardar =  WebDriverWait(driver,10).until(
                    EC.visibility_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:continuarSinGuardar"]'))
                )
                continuar_sin_guardar.click()

                #Error ocasional
                try:
                    invalido = WebDriverWait(driver,3).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm"]/div[5]/div[4]/div[2]/span'))
                        )
                    benf_tipo = driver.find_element_by_xpath('//*[@id="TransferenciasForm:select"]')
                    benf_tipo.send_keys(persona)

                    benf_id = driver.find_element_by_xpath('//*[@id="TransferenciasForm:numeroIdentificacion"]')
                    benf_id.send_keys(identificacion)

                    benf_name = driver.find_element_by_xpath('//*[@id="TransferenciasForm:nombreCompleto"]')
                    benf_name.send_keys(identificacion)

                    benf_alias = driver.find_element_by_xpath('//*[@id="TransferenciasForm:alias"]')
                    benf_alias.send_keys(identificacion)

                    benf_cuenta = driver.find_element_by_xpath('//*[@id="TransferenciasForm:numeroProducto"]')
                    benf_cuenta.send_keys(num_cuenta)

                    continuar_sin_guardar = driver.find_element_by_xpath('//*[@id="TransferenciasForm:continuarSinGuardar"]')
                    continuar_sin_guardar.click()
                except:
                    print('Todo bien')

                #Completar proceso de transferencia
                try:
                    monto = WebDriverWait(driver,20).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:monto"]'))
                    )
                    monto.send_keys(monto1)

                    continuar6 =  WebDriverWait(driver,20).until(
                            EC.presence_of_element_located((By.XPATH,'//a[@class="ui-commandlink ui-widget bod-button button-activar"]'))
                    )
                    continuar6.click()

                    select_cuenta =  WebDriverWait(driver,20).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:dtOrigen_data"]/tr'))
                    )
                    select_cuenta.click()
                    
                    select_concepto =  WebDriverWait(driver,20).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:concepto"]'))
                    )
                    select_concepto.send_keys('pago')

                    continuar7 =  WebDriverWait(driver,20).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:botonConcepto"]'))
                    )
                    continuar7.click()

                    continuar8 = WebDriverWait(driver,20).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="LotesForm:botonTrasferir"]'))
                    )
                    continuar8.click()

                    continuar9 = WebDriverWait(driver,20).until(
                        EC.visibility_of_element_located((By.XPATH,'//*[@id="LotesForm:ejecucion"]'))
                    )
                    continuar9.click()
                except:
                    driver.quit

                #Copiar confirmación de Salida
                confirmacion = WebDriverWait(driver,20).until(
                        EC.visibility_of_element_located((By.XPATH,'//td[@id="t2"][6]'))
                ).text

                #Volver a DESK
                driver.switch_to.window(driver.window_handles[0])

                confirmacion_salida = driver.find_element_by_xpath("//input[@class='MuiInputBase-input MuiOutlinedInput-input MuiInputBase-inputMarginDense MuiOutlinedInput-inputMarginDense']")
                confirmacion_salida.send_keys(confirmacion)
                
                time.sleep(1)
                
                #Manejar Salida
                finalizar_salida = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,"//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained']"))
                )
                action.move_to_element(finalizar_salida)
                finalizar_salida.click()

                estatus = 'COMPLETADA'

                #Pegar tiempo de operación
                print(datetime.now()-begin_time, tipo)

            #Entradas
            elif banco == 'BOD' and tipo == 'Buy':

                try:
                    trans_id = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[1]/div/p').text
                    nombre_reserve = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[2]/div/p').text
                    usuario_reserve = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[3]/div/p').text
                    rsv = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[1]/div[4]/div/p').text
                    tasa = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[2]/div[4]/div/p').text
                    telefono = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[3]/div/div/p').text
                    remitente = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[4]/div/div/div/p').text
                    beneficiario = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[1]/div/p').text
                    persona = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[2]/div/p').text
                    identificacion = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[3]/div/p').text
                    tipo_cuenta = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[4]/div/p').text
                    num_cuenta = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[5]/div/div[5]/div/p').text
                    monto = round(float(rsv)*float(tasa),2)
                    confirmacion = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[1]/div[1]/div[6]/div/div/div/p').text
                    print(tipo, banco, trans_id, nombre_reserve, confirmacion,)
                except:
                    driver.quit
                
                #Ir a Consultas
                driver.switch_to.window(driver.window_handles[1])
                try:
                    menu = WebDriverWait(driver,20).until(
                        EC.presence_of_element_located((By.XPATH,'//*[@id="supercontenedor"]/div[1]'))
                    )
                    menu.click()
                except:
                    driver.quit
            
                consultas = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="ico-menu-3"]'))
                )
                consultas.click()

                consultas2 = WebDriverWait(driver,20).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="submenu-ppal-3"]/ul/a[1]/li'))
                )
                consultas2.click()

                #Buscar numero de referencia (intentar 2 veces)
                for t in range(2):
                    error = None
                    try:
                        rango_fechas = WebDriverWait(driver,20).until(
                            EC.presence_of_element_located((By.XPATH,'//*[@id="date-range1"]'))
                        )
                        rango_fechas.click()

                        hoy = WebDriverWait(driver,20).until(
                            EC.presence_of_element_located((By.XPATH,'//*[@class="day toMonth  valid real-today"]'))
                        )
                        action.move_to_element(hoy).click().perform()
                        action.move_to_element(hoy).click().perform()


                        lookup = WebDriverWait(driver,20).until(
                            EC.presence_of_element_located((By.XPATH,"//input[@placeholder='Buscar movimientos']"))
                            )
                        lookup.send_keys(confirmacion)

                        match = WebDriverWait(driver,20).until(
                            EC.presence_of_element_located((By.XPATH,'//*[@id="formMovimientos:tablaMovimientos:0:detalle_movimiento"]/div[3]/p'))
                            ).text

                        monto3 = WebDriverWait(driver,20).until(
                            EC.presence_of_element_located((By.XPATH,'//*[@id="formMovimientos:tablaMovimientos:0:detalle_movimiento"]/div[5]/p'))
                            ).text
                        monto4 = monto3.split(" ")
                        monto5 = monto4[1]
                        monto6 = float(monto5.replace(',','.'))

                        print(match, "/", confirmacion, "__", monto, "/", monto6) 
                        t = t+1
                    except TimeoutException as e:
                        error = e
                    if error is None:
                        break
                if error is not None:
                    #Volver a DESK
                    driver.switch_to.window(driver.window_handles[0])
                    
                    #Contar el numero de transacciones
                    filas = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
                    num_filas = len(filas)
                    ultima = num_filas - 1

                    #Cancelar transacción
                    accion_votc = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[2]/div/div/div/div')
                    accion_votc.click()
                    time.sleep(2)
                    item_votc = driver.find_element_by_xpath('//*[@id="menu-"]/div[3]/ul/li[2]/span')
                    action.move_to_element(item_votc)
                    item_votc.click()
                    time.sleep(2)
                    cancelar = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[2]/div/button')
                    cancelar.click()

                    estatus = 'CANCELADA'
                        

                #Manejar entrada
                try:
                    if str(confirmacion) == str(match):
                        
                        #Volver a DESK
                        driver.switch_to.window(driver.window_handles[0])
                        
                        #Contar el numero de transacciones
                        filas = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
                        num_filas = len(filas)
                        ultima = num_filas - 1

                        if monto6 != monto:
                            time.sleep(2)
                            discrepancia = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div[2]/div[1]/button')
                            discrepancia.click()
                            time.sleep(2)
                            correccion = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div/input')
                            correccion.send_keys(monto6)
                            time.sleep(2)
                            interna = WebDriverWait(driver,20).until(
                                EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div[1]/button'))
                            )
                            interna.click()
                            time.sleep(2)
                            confirmacion_interna = WebDriverWait(driver,20).until(
                                EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr[4]/td/div/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div/input'))
                            )
                            confirmacion_interna.send_keys(confirmacion)
                            time.sleep(2)
                            completar_entrada = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div[2]/div[3]/button')
                            action.move_to_element(completar_entrada)
                            completar_entrada.click
                            time.sleep(2)
                            monto = monto6
                            estatus = 'COMPLETADA'
                        else:
                            interna = WebDriverWait(driver,20).until(
                                EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div[1]/button'))
                            )
                            interna.click()

                            confirmacion_interna = WebDriverWait(driver,20).until(
                                EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr[4]/td/div/div/div/div/div[2]/div[1]/div[2]/div[2]/div/div/input'))
                            )
                            confirmacion_interna.send_keys(confirmacion)

                            completar_entrada = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[1]/div[2]/div[3]/button')
                            action.move_to_element(completar_entrada)
                            completar_entrada.click()

                            estatus = 'COMPLETADA'
                    else:
                        #Volver a DESK
                        driver.switch_to.window(driver.window_handles[0])
                        
                        #Contar el numero de transacciones
                        filas = driver.find_elements_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr')
                        num_filas = len(filas)
                        ultima = num_filas - 1

                        #Cancelar transacción
                        accion_votc = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[2]/div/div/div/div')
                        accion_votc.click()

                        item_votc = driver.find_element_by_xpath('//*[@id="menu-"]/div[3]/ul/li[2]/span')
                        action.move_to_element(item_votc).click().perform()

                        cancelar = driver.find_element_by_xpath('//*[@id="root"]/div/main/div[1]/div[2]/div/div[3]/table/tbody/tr['+str(ultima+1)+']/td/div/div/div/div/div[2]/div[2]/div/button')
                        cancelar.click()

                        estatus = 'CANCELADA'
                except:
                    pass

                #Pegar tiempo de operación
                print(datetime.now()-begin_time, tipo)
            
            #Exportar datos de operación a sheets
            lista_datos = ((str(tipo), 
                str(trans_id),
                str(nombre_reserve),
                str(usuario_reserve),
                str(rsv),
                str(tasa),
                str(monto),
                str(telefono),
                str(beneficiario),
                str(persona),
                str(identificacion),
                str(tipo_cuenta),
                str(num_cuenta),
                str(confirmacion),
                str(estatus),
                str(datetime.now()),
                str(datetime.now()-begin_time)),
                ()
            )

            cuerpo = {
                'majorDimension' : 'ROWS',
                'values' : lista_datos
            }

            #Conseguir la primera fila vacia en sheets
            filas_sheets = sheets.spreadsheets().values().get(
                spreadsheetId = spreadsheet_id,
                majorDimension = 'ROWS',
                range = 'Log Operaciones!A1:A'
            ).execute()
            last_row = len(filas_sheets)+1

            #Exportar datos
            response = sheets.spreadsheets().values().update(
                spreadsheetId = spreadsheet_id,
                valueInputOption = 'USER_ENTERED',
                range = 'Log Operaciones!A'+str(last_row),
                body = cuerpo
            )
            response.execute()

        else:
            try:
                driver.switch_to.window(driver.window_handles[1])
                keepalive = WebDriverWait(driver,7).until(
                        EC.visibility_of_element_located((By.XPATH,'//*[@href="javascript:keepAliveSession()"]'))
                )
                action.move_to_element
                keepalive.click()
            except:
                driver.switch_to.window(driver.window_handles[0])
        continue

bot()
