import pandas as pd


# INGRESAR NOMBRE DE DIRECTORIO DONDE ESTAN LOS EXTRACTOS
Empresa = "FAX SRL"
# input("Ingrese el nombre de Directorio donde se encuentran los Extractos del Banco CORRIENTES: ").upper()



Enero = pd.read_excel(Empresa+"/CORRIENTES/01-22.xlsx", header=0)
Febrero = pd.read_excel(Empresa+"/CORRIENTES/02-22.xlsx", header=0)
Marzo = pd.read_excel(Empresa+"/CORRIENTES/03-22.xlsx", header=0)
Abril = pd.read_excel(Empresa+"/CORRIENTES/04-22.xlsx", header=0)
Mayo = pd.read_excel(Empresa+"/CORRIENTES/05-22.xlsx", header=0)
Junio = pd.read_excel(Empresa+"/CORRIENTES/06-22.xlsx", header=0)
Julio = pd.read_excel(Empresa+"/CORRIENTES/07-22.xlsx", header=0)
Agosto = pd.read_excel(Empresa+"/CORRIENTES/08-22.xlsx", header=0)
Septiembre = pd.read_excel(Empresa+"/CORRIENTES/09-22.xlsx", header=0)
Octubre = pd.read_excel(Empresa+"/CORRIENTES/10-22.xlsx", header=0)
Noviembre = pd.read_excel(Empresa+"/CORRIENTES/11-22.xlsx", header=0)

Extractos = pd.concat([Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, 
                       Agosto, Septiembre, Octubre, Noviembre], axis= 0)

# CONCATENAR EJERCICIO COMPLETO
Extractos = Extractos.fillna(0)
# Extractos = Extractos.drop(columns= Extractos.columns[1:5])
# Extractos = Extractos.drop(columns= Extractos.columns[3:6])
Extractos["IMPORTE"] = - Extractos["DEBITOS"] + Extractos["CREDITOS"]
Extractos = Extractos[Extractos["IMPORTE"] != 0]

Extractos = Extractos[Extractos["FECHA"] != 0]
# DATETIME FORMATO OK!!
Extractos["FECHA"] = pd.to_datetime(pd.Series(Extractos["FECHA"]))
Extractos["FECHA"] = Extractos["FECHA"].dt.strftime("%d/%m/%Y")

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
del Noviembre

Extractos["Saldo"] = Extractos["SALDO"]
Extractos["NUM. CHEQUE"] = Extractos["REFERENCIA"]
Extractos = Extractos.drop(["DEBITOS", "CREDITOS", "REFERENCIA", "SALDO", "CONCEPTO"], axis = 1)

Extractos = Extractos.rename(columns={"FECHA" : "Fecha","DETALLE" : "Concepto","IMPORTE" : "Importe"})

# ELIMINAR DUPLICADOS PARA COPIAR EN DESCRIPCION EXTRACTO SOS-CONTADOR
Descripcion_extracto = Extractos["Concepto"].astype(str).drop_duplicates()

Primera_fila = {"Fecha": "01/01/2022", "Concepto": 0, "Importe": 0, "Saldo": 0}

Importador = Extractos
Importador = Importador.drop(["NUM. CHEQUE"], axis = 1)
Importador = Importador.append(Primera_fila, ignore_index= True)
Importador["Fecha"] = pd.to_datetime(Importador["Fecha"], format = "%d/%m/%Y")
Importador = Importador.sort_values("Fecha")
Importador["Fecha"] = pd.to_datetime(pd.Series(Importador["Fecha"]))
Importador["Fecha"] = Importador["Fecha"].dt.strftime("%d/%m/%Y")

del Primera_fila

Importador.to_excel(Empresa+'/CORRIENTES '+Empresa+' Importar Extracto Bancario.xlsx', sheet_name='Extracto a SOS-Contador', index= False)
Descripcion_extracto.to_excel(Empresa+"/CORRIENTES Descripciones "+Empresa+".xlsx", sheet_name="Descripciones", index= False)
Extractos.to_excel(Empresa+"/CORRIENTES Consolidado "+Empresa+".xlsx", sheet_name="CORRIENTES año completo", index= False)


del Empresa

print("Archivo de Importación de Extractos de Banco CORRIENTES generado con éxito")


