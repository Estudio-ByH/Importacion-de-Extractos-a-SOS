import pandas as pd


# INGRESAR NOMBRE DE DIRECTORIO DONDE ESTAN LOS EXTRACTOS


Enero = pd.read_excel("TRANSPLANTE MISIONES SRL/MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Febrero = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Marzo = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Abril = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Mayo = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Junio = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Julio = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Agosto = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Septiembre = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Octubre = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)
Noviembre = pd.read_excel("MACRO/TRANSFERENCIAS/01-22.xls", header=7)

Transferencias = pd.concat([Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre], axis= 0)

# CONCATENAR EJERCICIO COMPLETO
Transferencias = Transferencias.fillna(0)


# DATETIME FORMATO OK!!
Transferencias["Fecha"] = pd.to_datetime(pd.Series(Transferencias["Fecha"]))
Transferencias["Fecha"] = Transferencias["Fecha"].dt.strftime("%d/%m/%Y")

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

# ELIMINAR DUPLICADOS PARA COPIAR EN DESCRIPCION EXTRACTO SOS-CONTADOR
Descripcion_extracto = Extractos["Concepto"].astype(str).drop_duplicates()

Primera_fila = {"Fecha": "01/01/2022", "Concepto": 0, "Importe": 0, "Saldo": 0}

Importador = Extractos
Importador = Importador.append(Primera_fila, ignore_index= True)
Importador["Fecha"] = pd.to_datetime(Importador["Fecha"], format = "%d/%m/%Y")
Importador = Importador.sort_values("Fecha")
Importador["Fecha"] = pd.to_datetime(pd.Series(Importador["Fecha"]))
Importador["Fecha"] = Importador["Fecha"].dt.strftime("%d/%m/%Y")

del Primera_fila

Importador.to_excel(Empresa+'/MACRO '+Empresa+' Importar Extracto Bancario.xlsx', sheet_name='Extracto a SOS-Contador', index= False)
Descripcion_extracto.to_excel(Empresa+"/MACRO Descripciones "+Empresa+".xlsx", sheet_name="Descripciones", index= False)
Extractos.to_excel(Empresa+"/MACRO Consolidado "+Empresa+".xlsx", sheet_name="MACRO año completo", index= False)


del Empresa

print("Archivo de Importación de Extractos de Banco MACRO generado con éxito")


