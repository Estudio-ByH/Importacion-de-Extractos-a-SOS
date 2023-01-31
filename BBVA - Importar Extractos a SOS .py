import pandas as pd


# INGRESAR NOMBRE DE DIRECTORIO DONDE ESTAN LOS EXTRACTOS
Empresa = input("Ingrese el nombre de Directorio donde se encuentran los Extractos del Banco BBVA: ").upper()

Enero = pd.read_excel(Empresa+"/BBVA/01-22.xlsx", header=0)
Febrero = pd.read_excel(Empresa+"/BBVA/02-22.xlsx", header=0)
Marzo = pd.read_excel(Empresa+"/BBVA/03-22.xlsx", header=0)
Abril = pd.read_excel(Empresa+"/BBVA/04-22.xlsx", header=0)
Mayo = pd.read_excel(Empresa+"/BBVA/05-22.xlsx", header=0)
Junio = pd.read_excel(Empresa+"/BBVA/06-22.xlsx", header=0)
Julio = pd.read_excel(Empresa+"/BBVA/07-22.xlsx", header=0)
Agosto = pd.read_excel(Empresa+"/BBVA/08-22.xlsx", header=0)
Septiembre = pd.read_excel(Empresa+"/BBVA/09-22.xlsx", header=0)
Octubre = pd.read_excel(Empresa+"/BBVA/10-22.xlsx", header=0)
Noviembre = pd.read_excel(Empresa+"/BBVA/11-22.xlsx", header=0)
Diciembre = pd.read_excel(Empresa+"/BBVA/12-22.xlsx", header=0)

Enero = Enero.fillna("")
Enero = Enero[Enero["SALDO"] != ""]
Enero["SALDO"] = Enero["SALDO"].str.replace("SALDO", "")
Enero["SALDO"] = Enero["SALDO"].str.replace(".", "")
Enero["SALDO"] = Enero["SALDO"].str.replace(",", ".")
Enero["DÉBITO"] = Enero["DÉBITO"].str.replace("DÉBITO", "")
Enero["DÉBITO"] = Enero["DÉBITO"].str.replace(".", "")
Enero["DÉBITO"] = Enero["DÉBITO"].str.replace(",", ".")
Enero["CRÉDITO"] = Enero["CRÉDITO"].str.replace("CRÉDITO", "")
Enero["CRÉDITO"] = Enero["CRÉDITO"].str.replace(".", "")
Enero["CRÉDITO"] = Enero["CRÉDITO"].str.replace(",", ".")
Enero["SALDO"] = pd.to_numeric(Enero["SALDO"].astype(float))
Enero["DÉBITO"] = pd.to_numeric(Enero["DÉBITO"]).astype(float)
Enero["CRÉDITO"] = pd.to_numeric(Enero["CRÉDITO"]).astype(float)
Enero = Enero.rename(columns={"Unnamed: 3": "CHEQUE"})    

Febrero = Febrero.fillna("")
Febrero = Febrero[Febrero["SALDO"] != ""]
Febrero["SALDO"] = Febrero["SALDO"].str.replace("SALDO", "")
Febrero["SALDO"] = Febrero["SALDO"].str.replace(".", "")
Febrero["SALDO"] = Febrero["SALDO"].str.replace(",", ".")
Febrero["DÉBITO"] = Febrero["DÉBITO"].str.replace("DÉBITO", "")
Febrero["DÉBITO"] = Febrero["DÉBITO"].str.replace(".", "")
Febrero["DÉBITO"] = Febrero["DÉBITO"].str.replace(",", ".")
Febrero["CRÉDITO"] = Febrero["CRÉDITO"].str.replace("CRÉDITO", "")
Febrero["CRÉDITO"] = Febrero["CRÉDITO"].str.replace(".", "")
Febrero["CRÉDITO"] = Febrero["CRÉDITO"].str.replace(",", ".")
Febrero["SALDO"] = pd.to_numeric(Febrero["SALDO"].astype(float))
Febrero["DÉBITO"] = pd.to_numeric(Febrero["DÉBITO"]).astype(float)
Febrero["CRÉDITO"] = pd.to_numeric(Febrero["CRÉDITO"]).astype(float)
Febrero = Febrero.rename(columns={"Unnamed: 3": "CHEQUE"})    

Marzo = Marzo.fillna("")
Marzo = Marzo[Marzo["SALDO"] != ""]
Marzo["SALDO"] = Marzo["SALDO"].str.replace("SALDO", "")
Marzo["SALDO"] = Marzo["SALDO"].str.replace(".", "")
Marzo["SALDO"] = Marzo["SALDO"].str.replace(",", ".")
Marzo["DÉBITO"] = Marzo["DÉBITO"].str.replace("DÉBITO", "")
Marzo["DÉBITO"] = Marzo["DÉBITO"].str.replace(".", "")
Marzo["DÉBITO"] = Marzo["DÉBITO"].str.replace(",", ".")
Marzo["CRÉDITO"] = Marzo["CRÉDITO"].str.replace("CRÉDITO", "")
Marzo["CRÉDITO"] = Marzo["CRÉDITO"].str.replace(".", "")
Marzo["CRÉDITO"] = Marzo["CRÉDITO"].str.replace(",", ".")
Marzo["SALDO"] = pd.to_numeric(Marzo["SALDO"].astype(float))
Marzo["DÉBITO"] = pd.to_numeric(Marzo["DÉBITO"]).astype(float)
Marzo["CRÉDITO"] = pd.to_numeric(Marzo["CRÉDITO"]).astype(float)
Marzo = Marzo.rename(columns={"Unnamed: 3": "CHEQUE"})    

Abril = Abril.fillna("")
Abril = Abril[Abril["SALDO"] != ""]
Abril["SALDO"] = Abril["SALDO"].str.replace("SALDO", "")
Abril["SALDO"] = Abril["SALDO"].str.replace(".", "")
Abril["SALDO"] = Abril["SALDO"].str.replace(",", ".")
Abril["DÉBITO"] = Abril["DÉBITO"].str.replace("DÉBITO", "")
Abril["DÉBITO"] = Abril["DÉBITO"].str.replace(".", "")
Abril["DÉBITO"] = Abril["DÉBITO"].str.replace(",", ".")
Abril["CRÉDITO"] = Abril["CRÉDITO"].str.replace("CRÉDITO", "")
Abril["CRÉDITO"] = Abril["CRÉDITO"].str.replace(".", "")
Abril["CRÉDITO"] = Abril["CRÉDITO"].str.replace(",", ".")
Abril["SALDO"] = pd.to_numeric(Abril["SALDO"].astype(float))
Abril["DÉBITO"] = pd.to_numeric(Abril["DÉBITO"]).astype(float)
Abril["CRÉDITO"] = pd.to_numeric(Abril["CRÉDITO"]).astype(float)
Abril = Abril.rename(columns={"Unnamed: 3": "CHEQUE"})    

Mayo = Mayo.fillna("")
Mayo = Mayo[Mayo["SALDO"] != ""]
Mayo["SALDO"] = Mayo["SALDO"].str.replace("SALDO", "")
Mayo["SALDO"] = Mayo["SALDO"].str.replace(".", "")
Mayo["SALDO"] = Mayo["SALDO"].str.replace(",", ".")
Mayo["DÉBITO"] = Mayo["DÉBITO"].str.replace("DÉBITO", "")
Mayo["DÉBITO"] = Mayo["DÉBITO"].str.replace(".", "")
Mayo["DÉBITO"] = Mayo["DÉBITO"].str.replace(",", ".")
Mayo["CRÉDITO"] = Mayo["CRÉDITO"].str.replace("CRÉDITO", "")
Mayo["CRÉDITO"] = Mayo["CRÉDITO"].str.replace(".", "")
Mayo["CRÉDITO"] = Mayo["CRÉDITO"].str.replace(",", ".")
Mayo["SALDO"] = pd.to_numeric(Mayo["SALDO"].astype(float))
Mayo["DÉBITO"] = pd.to_numeric(Mayo["DÉBITO"]).astype(float)
Mayo["CRÉDITO"] = pd.to_numeric(Mayo["CRÉDITO"]).astype(float)
Mayo = Mayo.rename(columns={"Unnamed: 3": "CHEQUE"})    

Junio = Junio.fillna("")
Junio = Junio[Junio["SALDO"] != ""]
Junio["SALDO"] = Junio["SALDO"].str.replace("SALDO", "")
Junio["SALDO"] = Junio["SALDO"].str.replace(".", "")
Junio["SALDO"] = Junio["SALDO"].str.replace(",", ".")
Junio["DÉBITO"] = Junio["DÉBITO"].str.replace("DÉBITO", "")
Junio["DÉBITO"] = Junio["DÉBITO"].str.replace(".", "")
Junio["DÉBITO"] = Junio["DÉBITO"].str.replace(",", ".")
Junio["CRÉDITO"] = Junio["CRÉDITO"].str.replace("CRÉDITO", "")
Junio["CRÉDITO"] = Junio["CRÉDITO"].str.replace(".", "")
Junio["CRÉDITO"] = Junio["CRÉDITO"].str.replace(",", ".")
Junio["SALDO"] = pd.to_numeric(Junio["SALDO"].astype(float))
Junio["DÉBITO"] = pd.to_numeric(Junio["DÉBITO"]).astype(float)
Junio["CRÉDITO"] = pd.to_numeric(Junio["CRÉDITO"]).astype(float)
Junio = Junio.rename(columns={"Unnamed: 3": "CHEQUE"})    

Julio = Julio.fillna("")
Julio = Julio[Julio["SALDO"] != ""]
Julio["SALDO"] = Julio["SALDO"].str.replace("SALDO", "")
Julio["SALDO"] = Julio["SALDO"].str.replace(".", "")
Julio["SALDO"] = Julio["SALDO"].str.replace(",", ".")
Julio["DÉBITO"] = Julio["DÉBITO"].str.replace("DÉBITO", "")
Julio["DÉBITO"] = Julio["DÉBITO"].str.replace(".", "")
Julio["DÉBITO"] = Julio["DÉBITO"].str.replace(",", ".")
Julio["CRÉDITO"] = Julio["CRÉDITO"].str.replace("CRÉDITO", "")
Julio["CRÉDITO"] = Julio["CRÉDITO"].str.replace(".", "")
Julio["CRÉDITO"] = Julio["CRÉDITO"].str.replace(",", ".")
Julio["SALDO"] = pd.to_numeric(Julio["SALDO"].astype(float))
Julio["DÉBITO"] = pd.to_numeric(Julio["DÉBITO"]).astype(float)
Julio["CRÉDITO"] = pd.to_numeric(Julio["CRÉDITO"]).astype(float)
Julio = Julio.rename(columns={"Unnamed: 3": "CHEQUE"})    

Agosto = Agosto.fillna("")
Agosto = Agosto[Agosto["SALDO"] != ""]
Agosto["SALDO"] = Agosto["SALDO"].str.replace("SALDO", "")
Agosto["SALDO"] = Agosto["SALDO"].str.replace(".", "")
Agosto["SALDO"] = Agosto["SALDO"].str.replace(",", ".")
Agosto["DÉBITO"] = Agosto["DÉBITO"].str.replace("DÉBITO", "")
Agosto["DÉBITO"] = Agosto["DÉBITO"].str.replace(".", "")
Agosto["DÉBITO"] = Agosto["DÉBITO"].str.replace(",", ".")
Agosto["CRÉDITO"] = Agosto["CRÉDITO"].str.replace("CRÉDITO", "")
Agosto["CRÉDITO"] = Agosto["CRÉDITO"].str.replace(".", "")
Agosto["CRÉDITO"] = Agosto["CRÉDITO"].str.replace(",", ".")
Agosto["SALDO"] = pd.to_numeric(Agosto["SALDO"].astype(float))
Agosto["DÉBITO"] = pd.to_numeric(Agosto["DÉBITO"]).astype(float)
Agosto["CRÉDITO"] = pd.to_numeric(Agosto["CRÉDITO"]).astype(float)
Agosto = Agosto.rename(columns={"Unnamed: 3": "CHEQUE"})    

Septiembre = Septiembre.fillna("")
Septiembre = Septiembre[Septiembre["SALDO"] != ""]
Septiembre["SALDO"] = Septiembre["SALDO"].str.replace("SALDO", "")
Septiembre["SALDO"] = Septiembre["SALDO"].str.replace(".", "")
Septiembre["SALDO"] = Septiembre["SALDO"].str.replace(",", ".")
Septiembre["DÉBITO"] = Septiembre["DÉBITO"].str.replace("DÉBITO", "")
Septiembre["DÉBITO"] = Septiembre["DÉBITO"].str.replace(".", "")
Septiembre["DÉBITO"] = Septiembre["DÉBITO"].str.replace(",", ".")
Septiembre["CRÉDITO"] = Septiembre["CRÉDITO"].str.replace("CRÉDITO", "")
Septiembre["CRÉDITO"] = Septiembre["CRÉDITO"].str.replace(".", "")
Septiembre["CRÉDITO"] = Septiembre["CRÉDITO"].str.replace(",", ".")
Septiembre["SALDO"] = pd.to_numeric(Septiembre["SALDO"].astype(float))
Septiembre["DÉBITO"] = pd.to_numeric(Septiembre["DÉBITO"]).astype(float)
Septiembre["CRÉDITO"] = pd.to_numeric(Septiembre["CRÉDITO"]).astype(float)
Septiembre = Septiembre.rename(columns={"Unnamed: 3": "CHEQUE"})    

Octubre = Octubre.fillna("")
Octubre = Octubre[Octubre["SALDO"] != ""]
Octubre["SALDO"] = Octubre["SALDO"].str.replace("SALDO", "")
Octubre["SALDO"] = Octubre["SALDO"].str.replace(".", "")
Octubre["SALDO"] = Octubre["SALDO"].str.replace(",", ".")
Octubre["DÉBITO"] = Octubre["DÉBITO"].str.replace("DÉBITO", "")
Octubre["DÉBITO"] = Octubre["DÉBITO"].str.replace(".", "")
Octubre["DÉBITO"] = Octubre["DÉBITO"].str.replace(",", ".")
Octubre["CRÉDITO"] = Octubre["CRÉDITO"].str.replace("CRÉDITO", "")
Octubre["CRÉDITO"] = Octubre["CRÉDITO"].str.replace(".", "")
Octubre["CRÉDITO"] = Octubre["CRÉDITO"].str.replace(",", ".")
Octubre["SALDO"] = pd.to_numeric(Octubre["SALDO"].astype(float))
Octubre["DÉBITO"] = pd.to_numeric(Octubre["DÉBITO"]).astype(float)
Octubre["CRÉDITO"] = pd.to_numeric(Octubre["CRÉDITO"]).astype(float)
Octubre = Octubre.rename(columns={"Unnamed: 3": "CHEQUE"})

Noviembre = Noviembre.fillna("")
Noviembre = Noviembre[Noviembre["SALDO"] != ""]
Noviembre["SALDO"] = Noviembre["SALDO"].str.replace("SALDO", "")
Noviembre["SALDO"] = Noviembre["SALDO"].str.replace(".", "")
Noviembre["SALDO"] = Noviembre["SALDO"].str.replace(",", ".")
Noviembre["DÉBITO"] = Noviembre["DÉBITO"].str.replace("DÉBITO", "")
Noviembre["DÉBITO"] = Noviembre["DÉBITO"].str.replace(".", "")
Noviembre["DÉBITO"] = Noviembre["DÉBITO"].str.replace(",", ".")
Noviembre["CRÉDITO"] = Noviembre["CRÉDITO"].str.replace("CRÉDITO", "")
Noviembre["CRÉDITO"] = Noviembre["CRÉDITO"].str.replace(".", "")
Noviembre["CRÉDITO"] = Noviembre["CRÉDITO"].str.replace(",", ".")
Noviembre["SALDO"] = pd.to_numeric(Noviembre["SALDO"].astype(float))
Noviembre["DÉBITO"] = pd.to_numeric(Noviembre["DÉBITO"]).astype(float)
Noviembre["CRÉDITO"] = pd.to_numeric(Noviembre["CRÉDITO"]).astype(float)
Noviembre = Noviembre.rename(columns={"Unnamed: 3": "CHEQUE"})  

Diciembre = Diciembre.fillna("")
Diciembre = Diciembre[Diciembre["SALDO"] != ""]
Diciembre["SALDO"] = Diciembre["SALDO"].str.replace("SALDO", "")
Diciembre["SALDO"] = Diciembre["SALDO"].str.replace(".", "")
Diciembre["SALDO"] = Diciembre["SALDO"].str.replace(",", ".")
Diciembre["DÉBITO"] = Diciembre["DÉBITO"].str.replace("DÉBITO", "")
Diciembre["DÉBITO"] = Diciembre["DÉBITO"].str.replace(".", "")
Diciembre["DÉBITO"] = Diciembre["DÉBITO"].str.replace(",", ".")
Diciembre["CRÉDITO"] = Diciembre["CRÉDITO"].str.replace("CRÉDITO", "")
Diciembre["CRÉDITO"] = Diciembre["CRÉDITO"].str.replace(".", "")
Diciembre["CRÉDITO"] = Diciembre["CRÉDITO"].str.replace(",", ".")
Diciembre["SALDO"] = pd.to_numeric(Diciembre["SALDO"].astype(float))
Diciembre["DÉBITO"] = pd.to_numeric(Diciembre["DÉBITO"]).astype(float)
Diciembre["CRÉDITO"] = pd.to_numeric(Diciembre["CRÉDITO"]).astype(float)
Diciembre = Diciembre.rename(columns={"Unnamed: 3": "CHEQUE"})      

Extractos = pd.concat([Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre], axis= 0)

# CONCATENAR EJERCICIO COMPLETO
Extractos = Extractos.fillna(0)
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/01", "/01/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/02", "/02/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/03", "/03/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/04", "/04/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/05", "/05/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/06", "/06/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/07", "/07/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/08", "/08/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/09", "/09/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/10", "/10/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/11", "/11/2022")
Extractos["FECHA"] = Extractos["FECHA"].str.replace("/12", "/12/2022")
Extractos["IMPORTE"] = Extractos["DÉBITO"] + Extractos["CRÉDITO"]

# Extractos = Extractos.drop(columns= Extractos.columns[1:5])
# Extractos = Extractos.drop(columns= Extractos.columns[3:6])
Extractos = Extractos.fillna(0)
Extractos = Extractos[Extractos["IMPORTE"] != 0]
Extractos = Extractos.drop(["ORIGEN","DÉBITO","CRÉDITO","SALDO"], axis=1)

Veps = Extractos[Extractos["CONCEPTO"] == "PAGOS AFIP"]

# DATETIME FORMATO OK!!
# Extractos["FECHA"] = pd.to_datetime(pd.Series(Extractos["FECHA"]))
# Extractos["FECHA"] = Extractos["FECHA"].dt.strftime("%d/%m/%Y")

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
del Diciembre

# ELIMINAR DUPLICADOS PARA COPIAR EN DESCRIPCION EXTRACTO SOS-CONTADOR

Descripcion_extracto = Extractos["CONCEPTO"].astype(str).drop_duplicates()

Primera_fila = {"FECHA": "01/01/2022", "CONCEPTO": 0, "IMPORTE": 0, "SALDO": 0}

Importador = Extractos
Importador = Importador.append(Primera_fila, ignore_index= True)
Importador["FECHA"] = pd.to_datetime(Importador["FECHA"], format = "%d/%m/%Y")
Importador = Importador.sort_values("FECHA")
Importador["FECHA"] = pd.to_datetime(pd.Series(Importador["FECHA"]))
Importador["FECHA"] = Importador["FECHA"].dt.strftime("%d/%m/%Y")



del Primera_fila

Importador.to_excel(Empresa+'/BBVA '+Empresa+' Importar Extracto Bancario.xlsx', sheet_name='Extracto a SOS-Contador', index= False)
Descripcion_extracto.to_excel(Empresa+"/BBVA Descripciones "+Empresa+".xlsx", sheet_name="Descripciones", index= False)
Extractos.to_excel(Empresa+"/BBVA Consolidado "+Empresa+".xlsx", sheet_name="BBVA año completo", index= False)
Veps.to_excel(Empresa+"/BBVA VEPs "+Empresa+".xlsx", sheet_name="VEPs Control", index= False)

del Empresa

print("Archivo de Importación de Extractos de Banco BBVA generado con éxito")


