import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt

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
        
    def convertExcelToCsv(self, name_origin, number_sheets, name_sheets): 

        # convert each sheet to csv and then read it using read_csv
        df={}
        i=1
        from subprocess import call
        excel= str(self.path_data / "excel_data/{}".format(name_origin))
        for sheet in name_sheets:
            name_sheet = "csv_data/" + sheet + ".csv"
            csv = str(self.path_data / name_sheet)
            path_vbs = str(self.path_data / 'ExcelToCsv.vbs')
            call(['cscript.exe', path_vbs, excel, csv, str(i)])
            df[sheet]=pd.read_csv(csv)
            i+=1
        return(df)





# Start the stopwatch / counter 
t1_start = perf_counter()

data_path = Path('C:/Users/andre/OneDrive/Heinsohn Proyects/Invima BI/data')
principal_tables = {'name': 'Tablas2.xlsx', 'number_sheets': 2, 'name_sheets':['FASES', 'H_PROTOCOLOS']}
documents_tables = {'name': 'Tablas.xlsx', 'number_sheets': 6, 'name_sheets':['DOCUMENTOS', 'COMITE_ETICA', 'SUMINISTROS_IMPORTADOS', 'CONSENTIMIENTO_INFORMADO', 'ENMIENDAS', 'MANUAL_INVESTIGADOR']}
protocols_tables = {'name': 'DataPacientes.xlsx', 'number_sheets': 7, 'name_sheets':['PACIENTES', 'INSTITUCIONES', 'INVESTIGADORES', 'SUSPENDER_INVESTIGACION', 'DISPOSITIVOS_BIOMEDICOS', 'MEDICAMENTOS', 'PROTOCOLOS']}

convert = convertData(data_path)
convert.createVbs()
df = convert.convertExcelToCsv(principal_tables['name'],
                 principal_tables['number_sheets'],
                 principal_tables['name_sheets'])

fases = pd.read_csv(data_path / 'csv_data/FASES.csv') 

h_protocolos = pd.read_csv(data_path / 'csv_data/H_PROTOCOLOS.csv')

#documentos = pd.read_excel(data_path / documents_tables, sheet_name=0, read_only=True)

#pacientes = pd.read_excel(data_path / protocols_tables, sheet_name=0, read_only=True)

#instituciones = pd.read_excel(data_path / protocols_tables, sheet_name=1, read_only=True)

#investigadores = pd.read_excel(data_path / protocols_tables, sheet_name=2, read_only=True)

#suspender_investigacion = pd.read_excel(data_path / protocols_tables, sheet_name=3, read_only=True)

#protocolos = pd.read_excel(data_path / protocols_tables, sheet_name=6, read_only=True)


##adding new columns that make useful data

#documentos['TIEMPO_ESTUDIO_DOCUMENTOS'] = documentos.FECHA_RESPUESTA - documentos.FECHA_PRESENTACION

# Stop the stopwatch / counter 
t1_stop = perf_counter() 

print("time: {}".format(t1_stop-t1_start))