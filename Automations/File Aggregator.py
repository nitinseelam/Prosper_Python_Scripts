from jira import JIRA
import pyodbc
import sqlalchemy
import getpass
import datetime
import pandas as pd
import numpy as np
import os
import time
import email
import HTMLParser
#import weasyprint
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.parser import Parser
import mimetypes
import smtplib
import xlsxwriter
import argparse
from os.path import expanduser
import json
import pdfkit
import xml
import html
import re
from lxml.html import fromstring, tostring

engine = sqlalchemy.create_engine("mssql+pyodbc://"+db_username+":"+db_password+"@dbrpt/Sandbox?driver=SQL+Server+Native+Client+11.0")
cnxn=engine.connect()

#cnx = pyodbc.connect('Trusted_connection= yes;DRIVER={SQL Server};SERVER=dbrpt;DATABASE=CircleOne')
now = datetime.datetime.now().strftime("%Y-%m-%d")
now

query = """
; Select * from Sandbox.dbo.DebtSaleMediaPull
"""
df = pd.read_sql_query(query, cnxn)
cnxn.close()

path = '//c1.prod/shares/Data/SQL/Nitin/Debt Sale Files/'

def terms_pdf(name, loanid, html, path):
    html = unicode(html, errors='ignore')
    pdfkit.from_string(html, path + str(loanid) + '.' + str(name) + '.' + 'Terms Of Use.pdf')
    
def br_pdf(name, loanid, html, path):
    html = unicode(html, errors='ignore')
    pdfkit.from_string(html, path + str(loanid) + '.' + str(name) + '.' + 'Borrower Registration Agreement.pdf')
    
def pn_pdf(name, loanid, html, path):
    html = unicode(html, errors='ignore')
    pdfkit.from_string(html, path + str(loanid) + '.' + str(name) + '.' + 'Borrower Promissory Note.pdf')
    
def tila_pdf(name, loanid, html,path):
    tila_html = unicode(html, errors='ignore')
    pdfkit.from_string(tila_html, path + str(loanid) + '.' + str(name) + '.' + 'Loan Truth in Lending Disclosure.pdf')
    
def fn_pdf(name, loanid, html,path):
    final_notice = HTMLParser.HTMLParser().unescape(html)
    start = final_notice.find('<font ')
    end = final_notice.find('/font>', start)+6
    fn = final_notice[start:end]
    fn = fn.replace("\r","")
    fn = fn.replace("\n","")
    fn = fn.replace("=","")
    fn_html = '<html><body>' + fn + '</body></html>'
    #html = unicode(html, errors='ignore')
    pdfkit.from_string(fn_html, path + str(loanid) + '.' + str(name) + '.' + 'Final Notice.pdf')
    
def co_pdf(name, loanid, html, path):
    
# Change Of Ownership Document is not encoded in DB.
    pdfkit.from_string(html, path + str(loanid) + '.' + str(name) + '.' + 'Ownership.pdf')

error_log = []

for i in range(0, len(df)):
    folder = path

    # Terms Of Use
    try:
        if df['Terms_Of_Use'][i] != 'None':
            terms_pdf(df['LastName'][i],df['LoanID'][i],df['Terms_Of_Use'][i], folder)   
        else:
            print('Missing Terms Of Use')
            error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Terms Of Use'
                 ,'Reason': 'Missing Document in Data'})
    except:
        print('Failed to convert Terms Of Use for:' + str(df['LoanID'][i]))
        error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Terms Of Use'
                 ,'Reason': 'Failed to Convert to PDF'})
                 

    #Borrower Registeration
    try:
        if df['Borrower_Registeration'][i] != None:
            br_pdf(df['LastName'][i],df['LoanID'][i],df['Borrower_Registeration'][i], folder)
         
        else:
            print('Missing Borrower Registration Form')
            error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Borrower Registration'
                 ,'Reason': 'Missing Document in Data'})
    except:
        print('Failed to convert Borrower Reigstration for:' + str(df['LoanID'][i]))
        error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Borrower Registration'
                 ,'Reason': 'Failed to Convert to PDF'})

        
    #Promissory Note
    try:
        if df['Promissory_Note'][i] != None:
            pn_pdf(df['LastName'][i],df['LoanID'][i],df['Promissory_Note'][i], folder)
        else:
            print('Missing Promissory Note')
            error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Promissory Note'
                 ,'Reason': 'Missing Document in Data'})
            
    except:
        print('Failed to convert Promissory Note for:' + str(df['LoanID'][i]))
        error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Promissory Note'
                 ,'Reason': 'Failed to Convert to PDF'})
 
    #TILA
    try:
        
        if df['TILA'][i] != None:
            tila_pdf(df['LastName'][i],df['LoanID'][i],df['TILA'][i], folder)
            
        else:
            print('Missing TILA Document')
            error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'TILA'
                 ,'Reason': 'Missing Document in Data'})
            
    except:
        print('Failed to convert TILA for:' + str(df['LoanID'][i]))
        error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'TILA'
                 ,'Reason': 'Failed to Convert to PDF'})
 
    try:
        
        if df['Final_Notice'][i] != None:
            fn_pdf(df['LastName'][i],df['LoanID'][i],df['Final_Notice'][i], folder)
        else:
            print('Missing Final Notice Document')
            error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Final Notice'
                 ,'Reason': 'Missing Document in Data'})
    except:
        print('Failed to convert Final Notice for:' + str(df['LoanID'][i]))
        error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Final Notice'
                 ,'Reason': 'Failed to Convert to PDF'})

    #Change Of Ownership
 
    try:
        
        if df['Change_Of_Ownership'][i] != None:
            co_pdf(df['LastName'][i],df['LoanID'][i],df['Change_Of_Ownership'][i], folder)
        else:
            print('Missing Change Of Ownership Document')
            error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Change Of Ownership'
                 ,'Reason': 'Missing Document in Data'})
    except:
        print('Failed to convert Change Of Ownership for:' + str(df['LoanID'][i]))
        error_log.append({
                  'LoanID': str(df['LoanID'][i])
                 ,'UserID': str(df['UserId'][i])
                 ,'Document': 'Change Of Ownership'
                 ,'Reason': 'Failed to Convert to PDF'})
                 

error_log
cnxn=engine.connect()
error_df = pd.DataFrame(error_log)
error_df.to_sql('DebtSaleMediaErrorLog_BK032019', cnxn, if_exists='append', index= False)
cnxn.close()

