#!/usr/bin/env python3

import argparse
import sys
from pathlib import Path
import pandas as pd
import numpy as np
import gspread
import requests
import openpyxl
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from datetime import datetime, timedelta
from collections import Counter
num_rows = 0
max = 0
path = "C:\\Users\\vwalk\\Angry_bot\\Angry_bot\\"
total = 0
class member:
    def __init__(self,name,dn,AP,AAP,DP,GS,c,spec):
        self.name = name
        self.dn = dn
        self.AP = AP
        self.AAP = AAP
        self.DP = DP
        self.GS = GS
        self.c = c
        self.spec = spec
def auth(sheetname):
    global max
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = Credentials.from_service_account_file(
    "credentials.json", scopes=SCOPES
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(
    "1Gggr3CyEMCZcjImj7JGRAEvZZ0TJc3OIpbgTUlvyha4")
    worksheet = sheet.worksheet(sheetname)
    df = pd.DataFrame(worksheet.get_all_records())
    df.to_excel("NSguildresponses.xlsx", index=False)
    
def upload(df,sheet,max):
    creds = Credentials.from_service_account_file(
    "credentials.json",
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
    "1Gggr3CyEMCZcjImj7JGRAEvZZ0TJc3OIpbgTUlvyha4"
)
    worksheet = spreadsheet.worksheet(sheet)  
    df_upload = df.copy()
    df_upload = df_upload.fillna("")

    for col in df_upload.columns:
        if pd.api.types.is_datetime64_any_dtype(df_upload[col]):
            df_upload[col] = df_upload[col].dt.strftime("%Y-%m-%d %H:%M:%S")
    
    worksheet.clear()
    worksheet.update(
    [df_upload.columns.tolist()] +
    df_upload.values.tolist()
)   
    worksheet.resize(rows=len(df_upload) + 1)
    


def downloadsheet():
    url = "https://docs.google.com/spreadsheets/d/1Gggr3CyEMCZcjImj7JGRAEvZZ0TJc3OIpbgTUlvyha4/export?format=xlsx&gid=1765336975"
    res = requests.get(url)
    with open("NS2.xlsx", "wb") as f:
        f.write(res.content)
        
def getColumn(name):
    xls = pd.read_excel(path+"NSguildresponses.xlsx", sheet_name="Form Responses 1")
    st = xls[name].dropna().tolist()
    return st

def getColumnNumber(name):
    xls = pd.read_excel(path+"NSguildresponses.xlsx", sheet_name="Form Responses 1")
    gearscoreName = xls[name]
    AP = pd.to_numeric(gearscoreName, errors="coerce").dropna().tolist()
    return AP

def getRow(row):
    xls = pd.read_excel(path+"NSguildresponses.xlsx", sheet_name="Form Responses 1")
    rown = xls.iloc[row].dropna().tolist()
    num_rows = len(xls)
    return rown

def parse_dt(s):
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unknown date format: {s}")

def latestfilter(listforrow):
    now = datetime.now()
    cutoff = now - timedelta(days=300)
    filtered = []
    for dt, *rest in listforrow:
        dt_parsed = pd.to_datetime(dt, errors="coerce")
        if pd.notna(dt_parsed) and dt_parsed >= cutoff:
            filtered.append([dt_parsed, *rest])
    return filtered

def removedoublename(newfilter,num):
    latest = {}
    for row in newfilter:
        name = row[num]
        if isinstance(row[0], str):
            dt = parse_dt(row[0])
        else:
            dt = row[0]
        normalized_row = (dt, *row[1:])
        if name not in latest or dt > latest[name][0]:
            latest[name] = normalized_row
    return list(latest.values())
        
def sortbyABC(num,choice):
    xls = pd.read_excel(path+"NSguildresponses.xlsx", sheet_name="Sheet1")
    global max
    max = len(xls)
    listforrow = []
    global GS
    GS = 0
    num_rows = len(xls)
    for i in range(0,num_rows):
        rown = xls.iloc[i].dropna().tolist()
        listforrow.append(rown)
    listforrow = sortbyGs(0,"NSguildresponses")
    listforrow = removedoublename(listforrow,1)
    listforrow = removedoublename(listforrow,2)
    listforrow.sort(key=lambda r: str(r[num]).lower(), reverse=False)
    xls = pd.DataFrame(listforrow)
    upload(xls,choice,max)
    xls.to_excel("holy.xlsx", index=False)

def sortbydate():
    xls = pd.read_excel(path+"NSguildresponses.xlsx", sheet_name="Sheet1")
    listforrow = []
    max = 0
    global GS
    GS = 0
    num_rows = len(xls)
    for i in range(0,num_rows):
        rown = xls.iloc[i].dropna().tolist()
        listforrow.append(rown)
    listforrow = removedoublename(listforrow,1)
    listforrow = removedoublename(listforrow,2)
    listforrow.sort(key=lambda r: r[0], reverse=True)
    xls = pd.DataFrame(listforrow)
    upload(xls,"Response",max)
    xls.to_excel("holy.xlsx", index=False)

def sortbyGs(sort,sheet):
    xls = pd.read_excel(path+sheet+".xlsx", sheet_name="Sheet1")
    listforrow = []
    global GS
    GS = 0
    num_rows = len(xls)
    for i in range(0,num_rows):
        rown = xls.iloc[i].dropna().tolist()
        if isinstance(rown[4], (int,float,np.integer,np.floating)):
            if isinstance(rown[5], (int,float,np.integer,np.floating)):
                if isinstance(rown[6], (int,float,np.integer,np.floating)):
                    GS = getGS(rown[4],rown[5],rown[6])
        if isinstance(rown[4], (int,float,np.integer,np.floating)):
            if GS < 894:
                rown.append(GS)
            listforrow.append(rown)
    last_index = len(listforrow[0]) - 1
    listforrow = [r for r in listforrow if len(r) > last_index]
    if sort == 1:
        newfilter = sorted(listforrow,key=lambda x: x[11],reverse=True)
        uf = removedoublename(newfilter,1)
        uf = removedoublename(newfilter,2)
        xls = pd.DataFrame(uf)
        upload(xls,"SortedGS",0)
        xls.to_excel("holy.xlsx", index=False)
    if sort == 0:
        return listforrow
   

def averageGS(row):
    xls = pd.read_excel(path+"holy.xlsx", sheet_name="Sheet1")
    num_rows = len(xls)
    listnum = []
    for i in range(0,num_rows):
        rown = xls.iloc[i].dropna().tolist()
        number = rown[11]
        listnum.append(int(number))
    total = sum(listnum) / num_rows
    xls.loc[row,"Average"] = total
    upload(xls,"SortedGS",0)
    xls.to_excel("holy.xlsx", index=False)
    return total

def quickfind(name):
    auth("SortedFamilyname")
    xls = pd.read_excel(path+"NSguildresponses.xlsx", sheet_name="Sheet1")
    num_rows = len(xls)
    results = []
    for i in range(0,num_rows):
        rown = xls.iloc[i].dropna().tolist()
        vname = rown[2]
        dname = rown[1]
        if vname.lower() == name.lower() or dname.lower() == name.lower():
            return rown
    return "no results"

def counter(num,row,out):
    xls = pd.read_excel(path+"holy.xlsx", sheet_name="Sheet1")
    num_rows = len(xls)
    listnum = []
    for i in range(0,num_rows):
        rown = xls.iloc[i].dropna().tolist()
        number = rown[num]
        listnum.append(number.replace(" ",""))
    words = " ".join(listnum).split()
    total = Counter(words)
    row = row
    for cls, total in total.items():
        xls.loc[row,"Class"] = cls
        xls.loc[row,"Count"] = total
        row+=1
    upload(xls,out,0)
    xls.to_excel("holy.xlsx", index=False)
    return Counter(words)
    
def getGS(ap,aap,dp):
    if(ap > aap):
        return ap + dp
    elif(aap > ap):
        return aap + dp
    else:
        return ap + dp
    

def sortmain():
    #name = getColumn("Discord Name")
    auth("Response")
    #downloadsheet()
    sortbydate()
    sortbyGs(1,"NSguildresponses")
    averageGS(1)
    sortbyABC(2,"SortedFamilyname")
    sortbyABC(3,"Sortedbyclass")
    counter(3,1,"Sortedbyclass")
    #counter(7,14,"SortedGS")
    


    
if __name__ == "__main__":
    sortmain()
