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
from selenium.common.exceptions import StaleElementReferenceException
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
    chrome_options.add_argument("--start-maximazed")
    
    driver = webdriver.Chrome(executable_path=CHROMEDRIVER_PATH,chrome_options=chrome_options)
    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)

    #Definir ActionChains
    action = ActionChains(driver)
    
    #Bienvenida
    print("Hola operardor VAST, bienvenido a tu BOT BOD, en breves momentos empezaré a operar.")
       
    #Inicio de sesion DESK
    driver.get('https://web.bancadigitalbod.com/nblee6/f/ext/Login/index.xhtml')
    

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
    print("Sesión iniciada en el banco")
    

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

    print("Entrando a transferencias")

    #Agregar un nuevo beneficiario
    try:
        add = WebDriverWait(driver,20).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:j_idt341"]'))
        )
        action.move_to_element(add)
        add.click()
        add.send_keys("solimilas")
    except:
        driver.quit

    time.sleep(7)
    soli = WebDriverWait(driver,20).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:dtTerceros_data"]/tr'))
    )
    soli.click()

    for t in range(2):
        error = None
        if t == 2:
            s = 1
        else:
            s = 20
        try:
            otp = WebDriverWait(driver,3).until(
                EC.presence_of_element_located((By.XPATH,'//*[@class="ancho100 zonaSeguraInput"]'))
            )
            print("Esperando por clave especial")
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
            print("Clave especial recibida:", codigo2)
            continuar_otp = driver.find_element_by_xpath('//a[@class="bod-button buttonCodigosOtp button-continuar"]')
            continuar_otp.click()

            t = t+1
        except TimeoutException as e:
            error = e
        if error is None:
            break
    if error is not None:
        driver.quit

    print("Realizando transferencia")

    origen = WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:dtProductoTercero_data"]/tr'))
    )
    origen.click()

    #Completar proceso de transferencia

    monto = WebDriverWait(driver,20).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="TransferenciasForm:monto"]'))
    )
    monto.send_keys("1")

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
        EC.presence_of_element_located((By.XPATH,'//*[@id="LotesForm:ejecucion"]'))
    )
    continuar9.click()
    print("Transferencia realizada con éxito")

        

    #Copiar confirmación de Salida
    freno = WebDriverWait(driver,30).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="formResultado"]/div[2]/a'))
    )
    i = 0
    confirmacion = WebDriverWait(driver,20).until(
        EC.visibility_of_all_elements_located((By.XPATH,'//td[@id="t1"]'))
    )
    for n in confirmacion:
        titulo = n.text
        i = i + 1
        print(i, titulo)
    i = 0
    confirmacion2 = WebDriverWait(driver,20).until(
        EC.visibility_of_all_elements_located((By.XPATH,'//td[@id="t2"]'))
    )
    for m in confirmacion2:
        titulo2 = m.text
        i = i + 1
        print(i, titulo2)
bot()
