import xml.etree.ElementTree as ET
import re
import xmltodict

nombre = '9'
XML = f'Facturas/XML/{nombre}.xml'

with open(XML, 'r', encoding='utf-8') as archivo:
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
with open(XML, 'r', encoding='utf-8') as fd:
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
if not iva or not iva_p or not total:
    iva_regex = re.search(r'<cbc:TaxAmount.*?>([\d,.]+)</cbc:TaxAmount>', text)
    if iva_regex and not iva:
        iva = iva_regex.group(1)
    
    taxable_amount_regex = re.search(r'<cbc:TaxableAmount.*?>([\d,.]+)</cbc:TaxableAmount>', text)
    if taxable_amount_regex and not iva_p:
        iva_p = taxable_amount_regex.group(1)
    
    payable_amount_regex = re.search(r'<cbc:PayableAmount.*?>([\d,.]+)</cbc:PayableAmount>', text)
    if payable_amount_regex and not total:
        total = payable_amount_regex.group(1)

lista = {
    'Empresa': Empresa,
    'NIT': nit,
    'Factura': Factura,
    'iva': iva,
    'iva_precio': iva_p,
    'total': total
}

print("\n")
print('Empresa: ', Empresa)
print('NIT: ', nit)
print('Factura: ', Factura)
print('IVA: ', iva, "%")
print('Valor de IVA: ', iva_p)
print('Total: ', total)
print("\n")
