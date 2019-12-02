# --------------------------------------------------------	
#   MODELO SUBASTA 	DE RECONFIGURACIÓN 2019 
#   DESARROLLADO DE USO LIBRE 
#   Oscar Carreño - Rightside SAS - 2019 - www.rightside.app
# --------------------------------------------------------
import pandas as pd
from pandasql import sqldf
import datetime
from os import getcwd
from openpyxl import load_workbook


xlFile = getcwd() + r"\subastaRECONF.xlsm"            
xlFile1 = getcwd() + r"\subastaRECONF_SALIDAS.xlsx"            

try:
    myfile = open(xlFile1, "a+") # 
except IOError:
    print ()
    print ("Por favor cierre el Archivo ",xlFile1)
    exit()

book1 = load_workbook(xlFile)
sh = book1["EJECUTAR"]

PMCC = sh['M10'].value
Qsubastada = sh['M12'].value
tolerancia = sh['M23'].value
toleranciaABS = sh['M25'].value
tiempoLimite = sh['M27'].value
optimizador = sh['M21'].value 

xl_reconf = pd.ExcelFile(xlFile)

#LECTURA TODAS LAS HOJAS DE ENTRADA DE EXCEL;
ofertas = xl_reconf.parse('ofertas').set_index(['agente','planta']) 

#LECTURA DE INDICES;
plantas  = sqldf("SELECT distinct(planta) FROM ofertas;", locals())

