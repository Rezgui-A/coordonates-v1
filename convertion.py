import pandas as pd
import xlsxwriter
import configparser

#Config
config = configparser.ConfigParser()
config.read('config.cfg')
DATA_EXCEL= config.get('EXCEL','file')
EXCEL_SHEETNAME= config.get('EXCEL','sheet')

#This Version uses Negative Values instead of letters (N , S , W ,E)

#DATA_EXCEL = 'exemple.xlsx' #Excel File
#EXCEL_SHEETNAME = "Sheet1" #Excel Sheet Name

#Data load
DATA = pd.read_excel(DATA_EXCEL)

def DDtoDMS(dd): #DD Decimal Degrees , DMS Degrees Minutes , seconds
    return str(int(float(dd))) +"°" +str(abs(float(dd)%1*60)) +"'" +str(abs(float(dd)%1*60%1*60))+ '"'


def DMStoDD(DMS):
    return str(float(DMS.split("°")[0]) + float(DMS.split("°")[1].split("'")[0])/60+(float(DMS.split("°")[1].split("'")[1].split('"')[0])/3600))

def syntaxCorrection(DATA):
    DATA.fillna('*',inplace = True)
    for i in DATA.index:
        if(str(DATA.loc[i, "X (DD)"]).find(" ") !=-1): #Removing Random spaces
            DATA.loc[i, "X (DD)"] = str(DATA.loc[i, "X (DD)"]).replace(" ", '')
        if(str(DATA.loc[i, "Y (DD)"]).find(" ") !=-1):
            DATA.loc[i, "Y (DD)"] = str(DATA.loc[i, "Y (DD)"]).replace(" ", '')
        if(str(DATA.loc[i, "X (DMS)"]).find(" ") !=-1):
            DATA.loc[i, "X (DMS)"] = str(DATA.loc[i, "X (DMS)"]).replace(" ", '')
        if(str(DATA.loc[i, "Y (DMS)"]).find(" ") !=-1):
            DATA.loc[i, "Y (DMS)"] = str(DATA.loc[i, "Y (DMS)"]).replace(" ", '')
         
        if(str(DATA.loc[i, "X (DMS)"]).find("`") !=-1): #Changing Symbols
            DATA.loc[i, "X (DMS)"] = str(DATA.loc[i, "X (DMS)"]).replace('`', "'")
        if(str(DATA.loc[i, "Y (DMS)"]).find("`") !=-1):
            DATA.loc[i, "Y (DMS)"] = str(DATA.loc[i, "Y (DMS)"]).replace('`', "'")
        if(str(DATA.loc[i, "Y (DMS)"]).find("’") !=-1):
            DATA.loc[i, "Y (DMS)"] = str(DATA.loc[i, "Y (DMS)"]).replace("’", "'")
        if(str(DATA.loc[i, "X (DMS)"]).find("’") !=-1):
            DATA.loc[i, "X (DMS)"] = str(DATA.loc[i, "X (DMS)"]).replace("’", "'")
        
        
        if(str(DATA.loc[i, "X (DD)"]) =="") : #Fillna Bug fix
            DATA.loc[i, "X (DD)"] = "*"
        if(str(DATA.loc[i, "Y (DD)"]) ==""):
            DATA.loc[i, "Y (DD)"] = "*"
        if(str(DATA.loc[i, "X (DMS)"]) ==""):
            DATA.loc[i, "X (DMS)"] = "*"
        if(str(DATA.loc[i, "Y (DMS)"]) ==""):
            DATA.loc[i, "Y (DMS)"] = "*"
            
    #DATA.fillna('*',inplace = True)
def Convertion(DATA):

    for x in DATA.index: 
        if(DATA.loc[x,"X (DD)"] == "*" ):
            if(DATA.loc[x,"X (DMS)"] != "*"):
                DATA.loc[x,"X (DD)"] = DMStoDD(str(DATA.loc[x,"X (DMS)"]))
            else : 
                print("Error converting Row with index = ",x, "X (DMS) and X(DD) missing")
        if(DATA.loc[x,"Y (DD)"] == "*" ):
            if(DATA.loc[x,"Y (DMS)"] != "*"):
                DATA.loc[x,"Y (DD)"] = DMStoDD(str(DATA.loc[x,"Y (DMS)"]))
            else : 
                print("Error converting Row with index = ",x,"Y(DMS) and Y(DD) missing")
        if (DATA.loc[x,"X (DMS)"] == "*" ):
            if(DATA.loc[x,"X (DD)"] != "*"):
                DATA.loc[x,"X (DMS)"] = DDtoDMS(DATA.loc[x,"X (DD)"])
        if (DATA.loc[x,"Y (DMS)"] == "*" ):
            if(DATA.loc[x,"Y (DD)"] != "*"):
                DATA.loc[x,"Y (DMS)"] = DDtoDMS(DATA.loc[x,"Y (DD)"])

def ExtractingDATA(DATA):
    DATA.sample(4)
    writer = pd.ExcelWriter(DATA_EXCEL, engine='xlsxwriter')
    DATA.to_excel(writer, sheet_name=EXCEL_SHEETNAME,
            startrow=0, index=False)
    #DATA.to_excel(writer, sheet_name=EXCEL_SHEETNAME)
    workbook  = writer.book
    workbook.close()
    worksheet = writer.sheets[EXCEL_SHEETNAME]

syntaxCorrection(DATA)
Convertion(DATA)
ExtractingDATA(DATA)
print(DATA)