import imaplib
import email
from email.header import decode_header
from imapclient import IMAPClient
import os
import zipfile
from pathlib import Path
import shutil
import xml.etree.ElementTree as ET
import xmltodict
import re
import json 
import time  # Importar la biblioteca time para introducir pausas

# Configuración del servidor SMTP y credenciales
smtp_server = 'smtp.office365.com'
smtp_port = 587
username = 'pruebatomy@outlook.com'
password = 'OMGtecnoevolution2024#'

# Configuración del servidor IMAP
imap_server = 'outlook.office365.com'
imap_port = 993

# Función para crear carpeta
def crear_carpeta(ruta):
    try:
        if not os.path.exists(ruta):
            os.makedirs(ruta)
            print(f"Carpeta '{ruta}' creada exitosamente.")
        else:
            print(f"La carpeta '{ruta}' ya existe.")
    except OSError as e:
        print(f"Error al crear la carpeta '{ruta}': {e}")

# Función para mover los archivos xml y pdf
def mover_archivo(origen, destino):
    try:
        shutil.move(origen, destino)
        print(f"Archivo '{origen}' movido a '{destino}' exitosamente.")
    except Exception as e:
        print(f"Error al mover el archivo: {e}")


def mostrar(arreglo):
    # Carpeta donde se guardará el archivo JSON
    carpeta = 'Archivos/json/'

    # Crear la carpeta si no existe
    if not os.path.exists(carpeta):
        os.makedirs(carpeta)
        print(f"Carpeta '{carpeta}' creada exitosamente.")

    # Ruta completa del archivo JSON
    ruta_json = os.path.join(carpeta, 'archivo.json')

    try:
        # Convertir el arreglo a formato JSON
        json_string = json.dumps(arreglo, indent=4, ensure_ascii=False)

        # Escribir el JSON en el archivo
        with open(ruta_json, 'w', encoding='utf-8') as archivo:
            archivo.write(json_string)

        print(f"Se ha creado '{ruta_json}' con éxito!")
    
    except Exception as e:
        print(f"Error al escribir el archivo JSON: {e}")


# Función para manejar nuevos correos
def handle_new_emails():
    try:
        # Conexión al servidor IMAP
        with IMAPClient(imap_server, imap_port, use_uid=True) as imap:
            imap.login(username, password)
            imap.select_folder('INBOX')

            while True:  # Bucle infinito para estar siempre escuchando
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

                                    folder_path = Path('Archivos/temp/')
                                    folder_path_pdf = Path('Archivos/pdf/')
                                    folder_path_xml = Path('Archivos/xml/')

                                    if filename:
                                        attachment_path = folder_path / filename

                                        with open(attachment_path, 'wb') as f:
                                            f.write(part.get_payload(decode=True))

                                        filename_sin_ext = attachment_path.stem

                                        nueva_carpeta = folder_path / filename_sin_ext
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
                                                full_file_path = nueva_carpeta / file
                                                if file.endswith('.pdf'):
                                                    pdf_destino = folder_path_pdf / file
                                                    mover_archivo(full_file_path, pdf_destino)
                                                elif file.endswith('.xml'):
                                                    xml_destino = folder_path_xml / file
                                                    mover_archivo(full_file_path, xml_destino)
                                                    # Procesar el XML
                                                    
                                                    resultado = procesar_xml(xml_destino)
                                                    mostrar(resultado)
                                        else:
                                            print("El archivo ZIP no contiene ambos archivos PDF y XML. Intentando extraer URL de QRCode...")
                                            
                else:
                    print("No hay nuevos correos.")
                
                # Esperar antes de revisar nuevamente
                time.sleep(20)  # Pausa de 20 segundos antes de la siguiente verificación

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

        # Función para extraer valores usando regex
        def extraer_valor(pattern, text):
            match = re.search(pattern, text)
            return match.group(1).strip() if match else None

        # Obtener información directamente del XML
        def obtener_valor(xpath, namespaces):
            elem = root.find(xpath, namespaces)
            return elem.text if elem is not None else None

        # Obtener el nombre de la empresa
        Empresa = obtener_valor('.//cbc:RegistrationName', namespaces)

        # Obtener el NIT de la empresa
        nit = obtener_valor('.//cbc:CompanyID', namespaces)

        # Obtener el número de factura
        factura_elem = root.find('.//cbc:ParentDocumentID', namespaces)
        Factura = factura_elem.text if factura_elem is not None else obtener_valor('.//cbc:ID', namespaces)
    

        # Procesar XML con xmltodict para facilitar la manipulación de los datos
        doc = xmltodict.parse(text)

        # Intentar extraer la descripción, teniendo en cuenta ambas posibles estructuras
        cbc_description = None
        try:
            cbc_description = doc['AttachedDocument']['cac:Attachment']['cac:ExternalReference']['cbc:Description']
        except KeyError:
            try:
                cbc_description = doc['Invoice']['cac:AdditionalDocumentReference']['cac:Attachment']['cac:ExternalReference']['cbc:Description']
            except KeyError:
                pass

        # Extraer IVA y Valor de IVA de la descripción si está disponible, sino del XML directamente
        iva = iva2 = iva_p1 = valor_iva2 = total = impuesto_iva = impuesto_iva2 = totalimpuesto = subtotal = None

        if cbc_description:
            # Extraer todos los valores de <cbc:Percent>
            iva_values = re.findall(r'<cbc:Percent>(.*?)</cbc:Percent>', cbc_description)
            if len(iva_values) > 0:
                iva = iva_values[0].strip()
            if len(iva_values) > 1:
                iva2 = iva_values[1].strip()

            # Extraer Valor de IVA y Total
            iva_p1 = extraer_valor(r'ValIva:\s*([\d,.]+)', cbc_description)
            total = extraer_valor(r'ValTolFac:\s*([\d,.]+)', cbc_description)
            valor_iva2 = extraer_valor(r'ValIva2:\s*([\d,.]+)', cbc_description)

        # Si no se encuentra en la descripción, extraer del XML directamente
        if not iva:
            iva_elems = root.findall('.//cac:TaxTotal/cac:TaxSubtotal/cbc:Percent', namespaces)
            if len(iva_elems) > 0:
                iva = iva_elems[0].text.strip()
            if len(iva_elems) > 1:
                iva2 = iva_elems[1].text.strip()


        # Obtener el iva de la otra rama si aún no se ha encontrado
        if not iva:
            iva = obtener_valor('.//cbc:Percent', namespaces)
            #print("Este es el porcentaje IVA: ", iva)

        if not iva_p1:
            iva_p1 = obtener_valor('.//cac:TaxTotal/cbc:TaxableAmount', namespaces)

        if not total:
            total = obtener_valor('.//cac:LegalMonetaryTotal/cbc:PayableAmount', namespaces)

        if not iva_p1:
            iva_p_regex = re.search(r'<cbc:TaxableAmount.*?>([\d,.]+)</cbc:TaxableAmount>', text)
            iva_p1 = iva_p_regex.group(1) if iva_p_regex else iva_p1

        if not total:
            total_regex = re.search(r'<cbc:PayableAmount.*?>([\d,.]+)</cbc:PayableAmount>', text)
            total = total_regex.group(1) if total_regex else total

        if not valor_iva2:
            valor_iva2_regex = re.findall(r'<cbc:TaxableAmount.*?>([\d,.]+)</cbc:TaxableAmount>', text)
            if len(valor_iva2_regex) > 1:
                valor_iva2 = valor_iva2_regex[1].strip()

        # Extraer impuestoIva, impuestoIva2 y totalimpuesto usando expresiones regulares
        impuesto_iva_regex = re.findall(r'<cbc:TaxAmount.*?>([\d,.]+)</cbc:TaxAmount>', text)
        if len(impuesto_iva_regex) > 0:
            totalimpuesto = impuesto_iva_regex[0].strip()
            impuesto_iva = totalimpuesto

        if len(impuesto_iva_regex) > 1:
            impuesto_iva2 = impuesto_iva_regex[1].strip()

        # Extraer el subtotal
        subtotal_match = re.search(r'<cbc:LineExtensionAmount.*?>([\d,.]+)</cbc:LineExtensionAmount>', text)
        if subtotal_match:
            subtotal = subtotal_match.group(1).strip()

        # Si no hay un segundo IVA o si IVA y IVA2 son iguales, asignar null a Valor_IVA2 e Impuesto_IVA2
        if not iva2 or (iva and iva2 and iva == iva2):
            iva2 = None
            valor_iva2 = None
            impuesto_iva2 = None

        # Crear un diccionario con la información extraída del XML
        lista = {
            'Empresa': Empresa,
            'NIT': nit,
            'Factura': Factura,
            'IVA': iva if iva else None,
            'IVA2': iva2 if iva2 else None,
            'Valor_IVA': iva_p1 if iva_p1 else None,
            'Valor_IVA2': valor_iva2 if valor_iva2 else None,
            'Impuesto_IVA': impuesto_iva if impuesto_iva else None,
            'Impuesto_IVA2': impuesto_iva2 if impuesto_iva2 else None,
            'Total_Impuesto': totalimpuesto if totalimpuesto else None,
            'Subtotal': subtotal if subtotal else None,
            'Total': total if total else None
        }

        # Retornar el diccionario con la información procesada
        return lista

    except Exception as e:
        # Manejar cualquier error que ocurra durante el procesamiento del XML
        print(f"Error al procesar el XML: {e}")
        return None


handle_new_emails() 

