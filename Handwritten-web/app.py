# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 20:29:14 2020

@author: Manomay
"""
import os
import html5lib
import win32com.client
import pythoncom
import re
import pandas as pd
#from selenium import webdriver
#from selenium.webdriver.common.action_chains import ActionChains
#from selenium.webdriver.common.by import By
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as ec
from time import sleep
#from selenium.common.exceptions import TimeoutException
import dateutil
import datetime
import sys
import numpy as np
import logging
import pyodbc
import dateparser
from flask import Flask, render_template
import win32com.client
import datetime
from flask_material import Material
from flask import Flask, render_template, url_for, request, redirect, Response, session, flash, send_from_directory, \
    send_file, abort, jsonify, logging, make_response
    
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

app = Flask(__name__,static_url_path="", static_folder=resource_path('static'), template_folder=resource_path("templates"))
Material(app)
print(resource_path('static'))
print(resource_path("templates"))



@app.route('/')
def index():
    return render_template("index.html")

@app.route("/dashboard", methods=['POST',"GET"])
def dashpage():
    cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.2.150;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay')
    cursor = cnxn.cursor()
    query = """select *  from SP_PolicyExtraction"""
    #sql_query = pd.read_sql_query(query,cnxn)
    table = cursor.execute(query)
    for i in table:
        print(i[0])
    
    ID = []
    Date =[]
    Origin=[]
    Doc_Name =[]
    Meta_Data_Extracted=[]
    PolicyNumber=[]
    Name_Insured=[]
    TypeOf_Doc=[]
    Success_Rate=[]
    Status=[]
    Date_Completed = []
    Action = []
    file_path=[]
    file_rename=[]
    query1 = "select count(*) from SP_PolicyExtraction"
    table = cursor.execute(query)
    for i in table:
        ID.append(i[0])
        Date.append(i[1])
        #Name.append(i[4])
        #Date_Time_Submitted.append(i[2])
        Origin.append(i[2])
        Doc_Name.append(i[3])
        Meta_Data_Extracted.append(i[4])
        PolicyNumber.append(i[5])
        Name_Insured.append(i[6])
        TypeOf_Doc.append(i[7])
        Success_Rate.append(i[8])
        Status.append(i[9])
        Date_Completed.append(i[10])
        Action.append(i[11])
        file_path.append(i[12])
        file_rename.append(i[13])
    print(file_rename,"#############################################################")
    count = cursor.execute(query1)
    count = [int(i[0]) for i in cursor.fetchall()]
    cp = count[0]
    return render_template("dashboard.html",ID=ID,Date=Date,Origin=Origin,file_path=file_path,file_rename=file_rename,Doc_Name=Doc_Name,Meta_Data_Extracted=Meta_Data_Extracted,PolicyNumber=PolicyNumber,Name_Insured=Name_Insured,TypeOf_Doc=TypeOf_Doc,Success_Rate=Success_Rate,Status=Status,Action=Action,Date_Completed=Date_Completed,cp=cp)


@app.route("/images", methods=['POST','GET','OPTIONS'])
def images():
    print("its image path")
    filepath= request.form.get("filepath")
    print(filepath)
    idfile= request.form.get("idfile")
    print(idfile)
    rate= request.form.get("rate")
    print(idfile)
    polnum= request.form.get("polnum")
    print(filepath)
    nain= request.form.get("nain")
    print(filepath)
    tydo= request.form.get("tydo")
    print(filepath)
    filerename = request.form.get("filerename")
    print(filerename)
    t = os.path.split(filepath)
    print(t)
    filepath1 = filepath.replace("/","\\")
    filelist = os.listdir("static/")
    print(filelist)
    imageList = os.listdir('static/'+t[-1])
    imagelist = ['../static/'+t[-1]+"/"+image for image in imageList]
    print(imageList)
    #return render_template("home.html")
    return render_template("home.html", imagelist=imagelist,filepath=filepath,rate=rate,filerename=filerename,polnum=polnum,nain=nain,tydo=tydo,idfile=idfile)
@app.route("/docrename", methods=['POST','GET','OPTIONS'])
def docrename():
    print("Its doc rename dun")
    polnum= request.form.get("polnum")
    nain= request.form.get("nain")
    tydo= request.form.get("tydo")
    rate= request.form.get("rate")
    filerename = request.form.get("filerename")
    print(filerename,"+++++++++++++++++++++++++++")
    if tydo=="Renewal Supplemental Application " +"Citizens Assumption Policies." or tydo=="Renewal Supplemental Application " +"Citizens Assumption Policies":
        sf ="RSACAP"
    elif tydo== "Renewal Supplemental Application":
        sf = "RSA"
    elif tydo=="COMMON POLICY CHANGE ENDROSEMENT":
        sf="CPCE"
    elif tydo=="ACKNOWLEDGEMENT OF CONSENT TO RATE":
        sf="ACR"
    elif tydo=="INSURANCE COVERAGE NOTIFICATION(S) ":
        sf="ICN"
    elif tydo=="COMMERCIAL PROPERTY POLICY DECLARATIONS":
        sf="CPPD"
    else:
        sf="--"
    met = polnum+"_"+nain+"_"+sf
    
    idfile= request.form.get("idfile")
    filepath= request.form.get("filepath")
    t = os.path.split(filerename)
    print(t)
    filerenamed = t[0]+"\\"+met+".pdf"
    os.rename(filerename,filerenamed)
    if met.count("--")==1:
            rate = "70%"
    elif met.count("--")==2:
        rate = "50%"
    elif met.count("--")==3:
        rate = "0%"
    else:
        rate ="100%"
    print(filepath,"its idfile================================================================================================================")
    cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.2.150;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay')
    cursor = cnxn.cursor()
    sql = "UPDATE SP_PolicyExtraction SET PolicyNumber = ?,Name_Insured=?,TypeOf_Doc=?,Success_Rate=?,Meta_Data_Extracted=?,file_rename=? WHERE ID = ?"
    val = (polnum,nain,tydo,rate,met,filerenamed,idfile)
    cursor.execute(sql, val)
    cursor.commit()
    return redirect(url_for('dashpage'))
@app.route("/home", methods=['POST','GET','OPTIONS'])
def home():
    return render_template("home.html")
@app.route("/dochistory", methods=['POST','GET','OPTIONS'])
def dochistory():
    return render_template("dochistory.html")
@app.route("/login", methods=['POST','GET','OPTIONS'])
def login():
    return render_template("index.html")
if __name__ == '__main__':
    app.run(debug=True)
