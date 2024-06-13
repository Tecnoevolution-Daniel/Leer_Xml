import imaplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email
from email.header import decode_header
from imapclient import IMAPClient
import time
import os
import zipfile
from pathlib import Path
import shutil
import xml.etree.ElementTree as ET
import json
import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import threading
import xmltodict

# Configuración del servidor SMTP y credenciales
smtp_server = 'smtp.office365.com'
smtp_port = 587
username = 'pruebatomy@outlook.com'
password = 'Colombia2023++'

# Configuración del servidor IMAP
imap_server = 'outlook.office365.com'
imap_port = 993

def crear_carpeta(ruta):
    try:
        os.makedirs(ruta)
        print(f"Carpeta '{ruta}' creada exitosamente.")
    except OSError as e:
        print(f"Error al crear la carpeta '{ruta}': {e}")

def mover_archivo(origen, destino):
    try:
        shutil.move(origen, destino)
        print(f"Archivo '{origen}' movido a '{destino}' exitosamente.")
    except Exception as e:
        print(f"Error al mover el archivo: {e}")

# Función para leer el contenido de un archivo XML
def leer_xml(ruta_archivo):
    with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
        contenido = archivo.read()
    return contenido

# Función para extraer el contenido de la etiqueta <sts:QRCode>
def extraer_contenido_qrcode(xml):
    print("Procesando XML para extraer QRCode...")
    soup = BeautifulSoup(xml, 'xml')  # Utilizamos el parser 'xml'
    
    # Buscar en <cbc:Description> con contenido CDATA
    description_tag = soup.find('cbc:Description')
    if description_tag and description_tag.string:
        # Analizar el contenido de CDATA como XML
        contenido_cdata = BeautifulSoup(description_tag.string, 'xml')
        qrcode_tag = contenido_cdata.find('sts:QRCode')
        if qrcode_tag:
            # Buscar la URL dentro del contenido del QRCode
            qrcode_text = qrcode_tag.text
            url_match = re.search(r'https?://[^\s]+', qrcode_text)
            if url_match:
                return url_match.group(0)
    
    # Buscar directamente en el XML para cualquier etiqueta <sts:QRCode>
    qrcode_tags = soup.find_all('sts:QRCode')
    for qrcode_tag in qrcode_tags:
        if qrcode_tag:
            qrcode_text = qrcode_tag.text
            url_match = re.search(r'https?://[^\s]+', qrcode_text)
            if url_match:
                return url_match.group(0)
    
    print("No se encontró la etiqueta <sts:QRCode> en el XML.")
    return None

# Función para manejar nuevos correos
def handle_new_emails():
    try:
        # Conexión al servidor IMAP
        with IMAPClient(imap_server, imap_port, use_uid=True) as imap:
            imap.login(username, password)
            imap.select_folder('INBOX')
            
            # Buscar correos no leídos
            messages = imap.search(['UNSEEN'])
            if messages:
                print("¡Tienes nuevos correos en tu bandeja de entrada!")
                for msgid, data in imap.fetch(messages, ['RFC822']).items():
                    email_message = email.message_from_bytes(data[b'RFC822'])
                    subject = str(email_message.get('Subject'))
                    
                    # Verificar si hay archivos adjuntos
                    if email_message.is_multipart():
                        for part in email_message.walk():
                            content_type = part.get_content_type()
                            if 'zip' in content_type:
                                print("¡Se encontró un archivo adjunto .zip!")
                               
                                filename = part.get_filename()

                                folder_path = 'Archivos/temp/'
                                folder_path_pdf = 'Archivos/pdf/'
                                folder_path_xml = 'Archivos/xml/'

                                if filename:
                                            
                                    attachment_path = os.path.join(folder_path, filename)
                                
                                    with open(attachment_path, 'wb') as f:
                                        f.write(part.get_payload(decode=True))

                                    filename_sin_ext = Path(filename).stem
                                    
                                    nueva_carpeta = os.path.join(folder_path, filename_sin_ext)
                                    crear_carpeta(nueva_carpeta)

                                    with zipfile.ZipFile(attachment_path, 'r') as zip_ref:
                                        zip_ref.extractall(nueva_carpeta) 
                                        print("¡Archivo ZIP descomprimido con éxito!")

                                    # Listar archivos descomprimidos
                                    files = os.listdir(nueva_carpeta)
                                    pdf_files = [file for file in files if file.endswith('.pdf')]
                                    xml_files = [file for file in files if file.endswith('.xml')]

                                    # Verificar si el ZIP contiene ambos archivos PDF y XML
                                    if len(pdf_files) > 0 and len(xml_files) > 0:
                                        print("El archivo ZIP contiene ambos archivos PDF y XML.")
                                        for file in files:
                                            full_file_path = os.path.join(nueva_carpeta, file)
                                            if file.endswith('.pdf'):
                                                pdf_destino = os.path.join(folder_path_pdf, file)
                                                mover_archivo(full_file_path, pdf_destino)
                                            elif file.endswith('.xml'):
                                                xml_destino = os.path.join(folder_path_xml, file)
                                                mover_archivo(full_file_path, xml_destino)
                                                
                                                # Procesar el XML
                                                procesar_xml(xml_destino)
                                    else:
                                        print("El archivo ZIP no contiene ambos archivos PDF y XML. Intentando extraer URL de QRCode...")
                                        # Si no tiene ambos archivos, intentar procesar con Selenium
                                        for file in files:
                                            if file.endswith('.xml'):
                                                xml_path = os.path.join(nueva_carpeta, file)
                                                contenido_qrcode = extraer_contenido_qrcode(leer_xml(xml_path))
                                                if contenido_qrcode:
                                                    print(f"Contenido de QRCode: {contenido_qrcode}")
                                                    abrir_pagina_qrcode(contenido_qrcode)
                                                else:
                                                    print("No se encontró la etiqueta <sts:QRCode> en el archivo XML.")
                                         
                                # Cerrar la conexión
                                #imap.logout()
                                
    except Exception as e:
        print(f"Error: {e}")

def procesar_xml(xml_destino):
    try:
        with open(xml_destino, 'r', encoding='utf-8') as archivo:
            text = archivo.read()

        tree = ET.ElementTree(ET.fromstring(text))
        root = tree.getroot()

        namespaces = {
            "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
            "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
            "sts": "dian:gov:co:facturaelectronica:Structures-2-1",
        }

        # NOMBRE
        empresa_elem = root.find('.//cbc:RegistrationName', namespaces)
        Empresa = empresa_elem.text if empresa_elem is not None else ""

        # NIT
        nit_elem = root.find('.//cbc:CompanyID', namespaces)
        nit = nit_elem.text if nit_elem is not None else ""

        # FACTURA
        factura_elem = root.find('.//cbc:ParentDocumentID', namespaces)
        if factura_elem is None:
            factura_elem = root.find('.//cbc:ID', namespaces)
        Factura = factura_elem.text if factura_elem is not None else ""

        # Procesar XML con xmltodict
        with open(xml_destino, 'r', encoding='utf-8') as fd:
            doc = xmltodict.parse(fd.read())

        # Intentar extraer la descripción, teniendo en cuenta ambas posibles estructuras
        cbc_description = None
        try:
            cbc_description = doc['AttachedDocument']['cac:Attachment']['cac:ExternalReference']['cbc:Description']
        except KeyError:
            try:
                cbc_description = doc['Invoice']['cac:AdditionalDocumentReference']['cac:Attachment']['cac:ExternalReference']['cbc:Description']
            except KeyError:
                pass

        # Función para extraer valores usando regex
        def extraer_valor(pattern, text):
            match = re.search(pattern, text)
            return match.group(1).strip() if match else ""

        # Extraer IVA y Valor de IVA de la descripción si está disponible, sino del XML directamente
        iva, iva_p, total = None, None, None

        if cbc_description:
            iva = extraer_valor(r'<cbc:Percent>(.*?)</cbc:Percent>', cbc_description)
            iva_p = extraer_valor(r'ValIva:\s*([\d,]+)', cbc_description)
            total = extraer_valor(r'ValTolFac:\s*([\d,]+)', cbc_description)

        # Si no se encuentra en la descripción, extraer del XML directamente
        if not iva:
            tax_amount_elem = root.find('.//cac:TaxTotal/cbc:TaxAmount', namespaces)
            iva = tax_amount_elem.text if tax_amount_elem is not None else ""

        if not iva_p:
            taxable_amount_elem = root.find('.//cac:TaxTotal/cbc:TaxableAmount', namespaces)
            iva_p = taxable_amount_elem.text if taxable_amount_elem is not None else ""

        if not total:
            payable_amount_elem = root.find('.//cac:LegalMonetaryTotal/cbc:PayableAmount', namespaces)
            total = payable_amount_elem.text if payable_amount_elem is not None else ""

        # También podemos usar expresiones regulares directamente en el texto XML si algunos valores aún no se han encontrado
        if not iva:
            iva_regex = re.search(r'<cbc:Percent>([\d,.]+)</cbc:Percent>', text)
            iva = iva_regex.group(1) if iva_regex else iva

        if not iva_p:
            iva_p_regex = re.search(r'<cbc:TaxableAmount.*?>([\d,.]+)</cbc:TaxableAmount>', text)
            iva_p = iva_p_regex.group(1) if iva_p_regex else iva_p

        if not total:
            total_regex = re.search(r'<cbc:PayableAmount.*?>([\d,.]+)</cbc:PayableAmount>', text)
            total = total_regex.group(1) if total_regex else total

        lista = {
            'Empresa': Empresa,
            'NIT': nit,
            'Factura': Factura,
            'IVA': iva,
            'IVA_Precio': iva_p,
            'Total': total
        }

        print("\nInformación extraída del XML:")
        print(f"Empresa: {Empresa}")
        print(f"NIT: {nit}")
        print(f"Factura: {Factura}")
        print(f"IVA: {iva}")
        print(f"Valor de IVA: {iva_p}")
        print(f"Total: {total}")
        print("\n")
        
    except Exception as e:
        print(f"Error al procesar el XML: {e}")

def abrir_pagina_qrcode(qrcode_url):
    try:
        # Validar la URL
        if not re.match(r'^https?://', qrcode_url):
            raise ValueError(f"URL inválida: {qrcode_url}")

        # Inicializa el controlador de Chrome
        service = Service(ChromeDriverManager().install())
        options = Options()
        options.add_argument("--start-maximized")  # Abrir navegador maximizado
        driver = webdriver.Chrome(service=service, options=options)

        # Abre la URL que contiene el QRCode
        driver.get(qrcode_url)

        # Espera explícita para el elemento específico
        xpath = '//*[@id="html-gdoc"]/div[3]/div/div[1]/div[3]/p/a'
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()

        # Pausar para ver la página abierta (puedes ajustar el tiempo)
        time.sleep(10)

        # Cerrar el navegador
        driver.quit()

    except Exception as e:
        print(f"Error al abrir la página del QRCode: {e}")

while True:
    hilo1 = threading.Thread(target=handle_new_emails, args=())
    hilo1.start()
    # Esperar un tiempo antes de verificar nuevamente (por ejemplo, cada 60 segundos)
    time.sleep(20)
