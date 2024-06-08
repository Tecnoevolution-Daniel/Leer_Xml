import fitz
import re
import pandas

with fitz.open("Facturas/PDF/dian2.pdf") as doc:

    text = ""
    for page in doc:
        text += page.get_text()
        

text = text.replace('�','')
lines = text.split("\n")


print(lines)


#Factura
index_rv = [1 if re.match(r'^Número de Factura',line) else 0 for line in lines]
print(index_rv)

index_rv = [i for i, s in enumerate(index_rv) if s==1 in index_rv]
#print(index_rv)


print(lines[index_rv[0]])


#Proveedor
index_rv = [1 if re.match(r'^Nombre Comercial',line) else 0 for line in lines]
index_rv = [i for i, s in enumerate(index_rv) if s==1 in index_rv]

print(lines[index_rv[0]])



#Documento

index_rv = [1 if re.match(r'^Nit del Emisor',line) else 0 for line in lines]

index_rv = [i for i, s in enumerate(index_rv) if s==1 in index_rv]

print(lines[index_rv[0]])


#Porcentaje IVA

index_rv = [1 if re.match(r'\d+\.\d+,\d+ \d+\.00',line) else 0 for line in lines]

index_rv = [i for i, s in enumerate(index_rv) if s==1 in index_rv]

'''
print(index_rv)

for i in index_rv:
    print(lines[i])
'''

porcentaje = lines[index_rv[0]].split(" ")

print("Porcentaje de IVA: ",porcentaje[1],"%")


#Total factura

index_rv = [1 if re.match(r'^Total factura',line) else 0 for line in lines]

index_rv = [i for i, s in enumerate(index_rv) if s==1 in index_rv]

total = lines[index_rv[0]+1]

print("Total Factura: ",total)