from cmath import nan
import win32com.client as win32
import win32com.client
from win32com.client import combrowse
import pandas as pd
import os
import sys

class Opium:
    def __init__(file,name:str):
        file.name = name
        app  = win32.Dispatch('Opium.Application')
        opium  = app.Documents
        file.doc=opium.Open(file.name)
        #print (file.name)
    def __del__(file):
        os.system('TASKKILL /F /IM Opium.exe')
        #print('__del__ was called')
    def file_name(file) -> str:
        print("file name:" + file.name)
    def CRV(file) -> pd.DataFrame:
        doc =file.doc
        NullValue = doc.Empty()

        CRV_List = [[]]
        try:
            CRVNamesList = list(doc.CRVNamesList())
            #print('получить список имен загруженных кривых:\n',CRVNamesList)
        except:
            CRVNamesList = []
        

        if  len(CRVNamesList) > 0:
            for CRVName in CRVNamesList:
                CRV_List.append([CRVName,list(doc.CRVGet(CRVName))])

            df = pd.DataFrame(CRV_List,columns =['CRVName','CRVNamesList'])
            df= df[df['CRVName'].notna()]
            df=df.set_index('CRVName')
            df=df.explode('CRVNamesList')

            df["MD"]=df['CRVNamesList'].apply(lambda x: x[0]  )
            df["Value"]=df['CRVNamesList'].apply(lambda x: x[1]  )
            df=df.drop('CRVNamesList',axis=1)
            df['Value']=df['Value'].apply(lambda x : nan if NullValue == x else x)

            df=df.reset_index()
            df= df[df['Value'].notna()]
            df= df[df['CRVName'].notna()]
            df= df[df['MD'].notna()]
            df= df[df['MD']>0]
            df = df.pivot(index='MD', columns='CRVName',values='Value')
            #df.to_excel('uu.xlsx')

            #print(df.head())
            #print(df.info())
        else:
            df = pd.DataFrame()
        return(df) 
    def GRN(file) -> pd.DataFrame:
        doc =file.doc
        NullValue = doc.Empty()

        
        CRV_List = [[]]
        try:
            CRVNamesList = list(doc.GRNNamesList())
            #print('получить список имен загруженных кривых:\n',CRVNamesList)
        except:
            CRVNamesList = []
        

        if  len(CRVNamesList) > 0:
            for CRVName in CRVNamesList:
                CRV_List.append([CRVName,list(doc.GRNGet(CRVName))])

            df = pd.DataFrame(CRV_List,columns =['CRVName','CRVNamesList'])
            df= df[df['CRVName'].notna()]
            df=df.set_index('CRVName')
            df=df.explode('CRVNamesList')

            df["MD"]=df['CRVNamesList'].apply(lambda x: x[0]  )
            df["Value"]=df['CRVNamesList'].apply(lambda x: x[1]  )
            df=df.drop('CRVNamesList',axis=1)
            df['Value']=df['Value'].apply(lambda x : nan if NullValue == x else x)

            df=df.reset_index()
            #df= df[df['Value'].notna()]
            #df['Value']=df['Value'].ffill()
            df= df[df['CRVName'].notna()]
            df= df[df['MD'].notna()]
            df= df[df['MD']>0]
            df = df.pivot(index='MD', columns='CRVName',values='Value')
            #df.to_excel('uu.xlsx')

            #print(df.head())
            #print(df.info())
        else:
            df = pd.DataFrame()
        return(df)
    def PLS(file) -> pd.DataFrame:
        doc =file.doc
        NullValue = doc.Empty()

        CRV_List = [[]]
        try:
            CRVNamesList = list(doc.PLSNamesList())
            #print('получить список имен загруженных кривых:\n',CRVNamesList)
        except:
            CRVNamesList = []
        

        if  len(CRVNamesList) > 0:
            for CRVName in CRVNamesList:
                CRV_List.append([CRVName,list(doc.PLSGet(CRVName))])

            df = pd.DataFrame(CRV_List,columns =['CRVName','CRVNamesList'])
            df= df[df['CRVName'].notna()]
            df=df.set_index('CRVName')
            df=df.explode('CRVNamesList')

            df["MD"]=df['CRVNamesList'].apply(lambda x: x[0]  )
            df["Value"]=df['CRVNamesList'].apply(lambda x: x[1]  )
            df=df.drop('CRVNamesList',axis=1)
            df['Value']=df['Value'].apply(lambda x : nan if NullValue == x else x)

            df=df.reset_index()
            #df= df[df['Value'].notna()]
            #df['Value']=df['Value'].ffill()
            df= df[df['CRVName'].notna()]
            df= df[df['MD'].notna()]
            df= df[df['MD']>0]
            df = df.pivot(index='MD', columns='CRVName',values='Value')
            #df.to_excel('uu.xlsx')

            #print(df.head())
            #print(df.info())
        else:
            df = pd.DataFrame()
        return(df)
    def PLT(file) -> pd.DataFrame:
        doc =file.doc
        NullValue = doc.Empty()
        CRV_List = [[]]
        try:
            CRVNamesList = list(doc.PLTNamesList())
            #print('получить список имен загруженных кривых:\n',CRVNamesList)
        except:
            CRVNamesList = []
        print(len(CRVNamesList))

        if  len(CRVNamesList) > 0:
            for CRVName in CRVNamesList:
                CRV_List.append([CRVName,list(doc.PLTGet(CRVName))])

            df = pd.DataFrame(CRV_List,columns =['CRVName','CRVNamesList'])
            df= df[df['CRVName'].notna()]
            df=df.set_index('CRVName')
            df=df.explode('CRVNamesList')

            df["MD"]=df['CRVNamesList'].apply(lambda x: x[0]  )
            df["Value"]=df['CRVNamesList'].apply(lambda x: x[1]  )
            df=df.drop('CRVNamesList',axis=1)
            df['Value']=df['Value'].apply(lambda x : nan if NullValue == x else x)

            df=df.reset_index()
            #df= df[df['Value'].notna()]
            df['Value']=df['Value'].ffill()
            df= df[df['CRVName'].notna()]
            df= df[df['MD'].notna()]
            df= df[df['MD']>0]
            df = df.pivot(index='MD', columns='CRVName',values='Value')
            #df.to_excel('uu.xlsx')

            #print(df.head())
            #print(df.info())
        else:
            df = pd.DataFrame()
        return(df)
    #конвертує pln в ексель
    def printXls(file):
        name = file.name.strip('.pln')
        with pd.ExcelWriter(name+'.xlsx') as writer:
            Opium.CRV(file).to_excel(writer, sheet_name='CRV - Криві')
            Opium.GRN(file).to_excel(writer, sheet_name='GRN - Набір границь')
            Opium.PLS(file).to_excel(writer, sheet_name='PLS - відліки')
            Opium.PLT(file).to_excel(writer, sheet_name='PLT')

#Ott=Opium('C:/Users/dmytro.lishchynskiy/Desktop/py/test.pln')

#print(Ott.CRV().head())
#print(Ott.GRN().head())
#print(Ott.PLS().head())
#print(Ott.PLT().head())

def main() -> int:
    #Ott.printXls()
    files = os.listdir()
    files = list(filter(lambda x: '.pln' in x ,files))
    
    for file in files:
        longName=os.path.join(os.getcwd(), file)
        print(longName)
        Opium(longName).printXls()
    #input()
    return 0

if __name__ == '__main__':
    sys.exit(main())  



