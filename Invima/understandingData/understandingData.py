import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime
from subprocess import run

#Count time of execution of program
from time import perf_counter

class convertData():

    def __init__(self, path):
        self.path_data = path

    def createVbs(self):
        #write vbscript to file
        vbscript="""if WScript.Arguments.Count < 3 Then
            WScript.Echo "Please specify the source and the destination files. Usage: ExcelToCsv <xls/xlsx source file> <csv destination file> <worksheet number (starts at 1)>"
            Wscript.Quit
        End If

        csv_format = 6

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
        dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
        worksheet_number = CInt(WScript.Arguments.Item(2))

        Dim oExcel
        Set oExcel = CreateObject("Excel.Application")

        Dim oBook
        Set oBook = oExcel.Workbooks.Open(src_file)
        oBook.Worksheets(worksheet_number).Activate

        oBook.SaveAs dest_file, csv_format

        oBook.Close False
        oExcel.Quit""";

        f = open(self.path_data / 'ExcelToCsv.vbs','w')
        f.write(str(vbscript.encode(encoding='utf-8').decode(encoding='utf-8')))
        f.close()
        
        # convert each sheet to csv and then read it using read_csv
    def convertExcelToCsv(self, names_origin, names_sheets, convert=1): 

        # name_origin = nombre tabla
        # name_sheets = nombre hojas en tabla
        #if we have already the csv converted, we can chose doesnt convert with parameter convert=1 to yes, =0 to no


        df={}
        if(convert==1):
            for name_origin, name_sheets in zip(names_origin, names_sheets):
                i=1
                excel= str(self.path_data / "excel_data/{}".format(name_origin))
                for sheet in name_sheets:
                    path_sheet = "csv_data/" + sheet + ".csv"
                    csv = str(self.path_data / path_sheet)
                    path_vbs = str(self.path_data / 'ExcelToCsv.vbs')
                    run(['cscript.exe', path_vbs, excel, csv, str(i)])
                    df[sheet]=pd.read_csv(csv, encoding="ISO-8859-1")
                    i+=1
            return(df)
        else:
            for name_origin, name_sheets in zip(names_origin, names_sheets):
                i=1
                for sheet in name_sheets:
                    path_sheet = "csv_data/" + sheet + ".csv"
                    csv = str(self.path_data / path_sheet)
                    df[sheet]=pd.read_csv(csv, encoding="ISO-8859-1")
            return(df)

# Start the stopwatch / counter 
t1_start = perf_counter()


#Principal path where is the proyect data
data_path = Path('C:/Users/andre/OneDrive/Heinsohn Proyects/Invima BI/data')

#tables information, names of tables and name of sheets in tables
tables = {'names': ['Tablas2.xlsx', 'Tablas.xlsx', 'DataPacientes.xlsx'],
         'names_sheets':[['FASES', 'H_PROTOCOLOS'],
                         ['DOCUMENTOS', 'COMITE_ETICA', 'SUMINISTROS_IMPORTADOS', 'CONSENTIMIENTO_INFORMADO', 'ENMIENDAS', 'MANUAL_INVESTIGADOR'],
                         ['PACIENTES', 'INSTITUCIONES', 'INVESTIGADORES', 'SUSPENDER_INVESTIGACION', 'DISPOSITIVOS_BIOMEDICOS', 'MEDICAMENTOS', 'PROTOCOLOS']]}



#use class convert to prepare convertion from xlsx to csv
convert = convertData(data_path)

#Create the VBS file that convert files
convert.createVbs()

#Create a DataFrame that contains all the tables converted before
df = convert.convertExcelToCsv(tables['names'],
                 tables['names_sheets'], convert=0)
#assigne corresponding dataframes to a variable
fases = df['FASES']
h_protocolos = df['H_PROTOCOLOS']
documentos = df['DOCUMENTOS']
pacientes = df['PACIENTES']
instituciones = df['INSTITUCIONES']
investigadores = df['INVESTIGADORES']
suspender_investigacion = df['SUSPENDER_INVESTIGACION']
protocolos = df['PROTOCOLOS']

#you can explore the dataframe to know better about informacion, this space is to do that:
##############################################################################################################
#Lets see some satistical values, select with number 1 to 7 the dataframe
#yo can select the dataFrame you want to see, fases, h_protocolos, documentos, etc

#protocolos.info()

#plot values in a DataFrame specific, in 'protocolos' you can plot 'CODIGOINVIMA' and 'ESTADO' for example



#protocolos[protocolos['ESTADO']=='APROBADO']['VERSION'].hist()
#plt.show()

###############################################################################################################
#finished the area to exploration the DataFrame
# Stop the stopwatch / counter 
t1_stop = perf_counter() 

print("time: {}".format(t1_stop-t1_start))