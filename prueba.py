import fitz 
import re
import pandas

with fitz.open("dian.pdf") as doc:

    text = ""
    for page in doc:
        text += page.get_text()
        

text = text.replace('�','')

#print(text)

lines = text.split("\n")

print(lines)

#Factura
factura = lines[31]
factura = factura.split(":")
factura = factura[1]
factura = factura.replace(" ","")

print("\nFactura: ",factura)

#Proveedor
print("\nNombre proveedor: ",lines[55])
print("\n")
#Tipo documento
id = ""
print(lines[56])
doc = lines[56]
doc = doc.split(":")
doc = doc[1]
doc = doc.replace(" ","")

num = lines[58]
num = num.split(":")
num = num[1]
num = num.replace(" ","")
num = num.replace(" ","")

print("\n")

id = doc + ": " + num
print(id)
print("\n")


