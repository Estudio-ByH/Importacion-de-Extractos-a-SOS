import pandas as pd

Empresa = input("Ingrese el nombre de Directorio donde se encuentran los Extractos del Banco ICBC: ")

# INGRESAR NOMBRE DE DIRECTORIO DONDE ESTAN LOS EXTRACTOS

Enero = pd.read_excel(Empresa+"/ICBC/01-22.xlsx", header=1)
Febrero = pd.read_excel(Empresa+"/ICBC/02-22.xlsx", header=1)
Marzo = pd.read_excel(Empresa+"/ICBC/03-22.xlsx", header=1)
Abril = pd.read_excel(Empresa+"/ICBC/04-22.xlsx", header=1)
Mayo = pd.read_excel(Empresa+"/ICBC/05-22.xlsx", header=1)
Junio = pd.read_excel(Empresa+"/ICBC/06-22.xlsx", header=1)
Julio = pd.read_excel(Empresa+"/ICBC/07-22.xlsx", header=1)
Agosto = pd.read_excel(Empresa+"/ICBC/08-22.xlsx", header=1)
Septiembre = pd.read_excel(Empresa+"/ICBC/09-22.xlsx", header=1)
Octubre = pd.read_excel(Empresa+"/ICBC/10-22.xlsx", header=1)


Extractos = pd.concat([Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre], axis= 0)

del Enero
del Febrero
del Marzo
del Abril
del Mayo
del Junio
del Julio
del Agosto
del Septiembre
del Octubre

Extractos = Extractos.drop(columns= Extractos.columns[8:10])
Extractos = Extractos.drop(columns= Extractos.columns[1])
Extractos = Extractos.drop(columns= Extractos.columns[5])
Extractos = Extractos.fillna(0)

# ELIMINAR DUPLICADOS PARA COPIAR EN DESCRIPCION EXTRACTO SOS-CONTADOR
Descripcion_extracto = Extractos["Concepto"].astype(str).drop_duplicates()

#ARMADO FORMATO IMPORTACIÓN

Primera_fila = {"Fecha": "01/01/2022", "Descripcion": 0, "Importe": 0, "Saldo": 0}


Importador = Primera_fila
Importador = Extractos.drop(columns= Extractos.columns[2:6])


Importador["Fecha contable"] = pd.to_datetime(Importador["Fecha contable"], format = "%d/%m/%Y")
Importador["Fecha contable"] = Importador["Fecha contable"].astype(str)
Importador["Año"] = Importador["Fecha contable"].str[0:4]
Importador["Mes"] = Importador["Fecha contable"].str[5:7]
Importador["Dia"] = Importador["Fecha contable"].str[8:10]
Importador["Fecha"] = Importador["Dia"].astype(str) + "/" + Importador["Mes"].astype(str) +"/"+ Importador["Año"].astype(str)
Importador["Descripcion"] = Importador["Concepto"]
Importador["Importe"] = Extractos["Debito en $"].astype(float) + Extractos["Credito en $"].astype(float)
Importador["Saldo"] = Extractos["Saldo en $"]

Importador = Importador.drop(["Año","Mes","Dia","Fecha contable","Concepto"], axis=1)

Importador = Importador.append(Primera_fila, ignore_index= True)

Importador["Fecha"] = pd.to_datetime(Importador["Fecha"], format = "%d/%m/%Y")
Importador = Importador.sort_values("Fecha")

Importador["Fecha"] = Importador["Fecha"].astype(str)
Importador["Año"] = Importador["Fecha"].str[0:4]
Importador["Mes"] = Importador["Fecha"].str[5:7]
Importador["Dia"] = Importador["Fecha"].str[8:10]
Importador["Fecha"] = Importador["Dia"].astype(str) + "/" + Importador["Mes"].astype(str) +"/"+ Importador["Año"].astype(str)
Importador = Importador.drop(["Año","Mes","Dia"], axis=1)

del Primera_fila

Importador.to_excel(Empresa+'/ICBC '+Empresa+' Importar Extracto Bancario.xlsx', sheet_name='Extracto a SOS-Contador', index= False)
Descripcion_extracto.to_excel(Empresa+"/ICBC Descripciones "+Empresa+".xlsx", sheet_name="Descripciones", index= False)
Extractos.to_excel(Empresa+"/ICBC Consolidado "+Empresa+".xlsx", sheet_name="ICBC año completo", index= False)

del Empresa

print("Archivo de Importación de Extractos de Banco ICBC generado con éxito")





