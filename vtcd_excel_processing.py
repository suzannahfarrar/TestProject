# -*- coding: utf-8 -*-
"""
Created on Tue May 14 11:22:27 2019

@author: 1579394
"""

from datetime import datetime
import pandas as pd
from pandas import ExcelWriter
import os
import sys

class ReportDownload:
    
    # Declaring country variable
    countries = []
    countries_CORE_TEAM = []
    countries_IN = []
    todays_date = datetime.now
    
    # Initializing variables within constructor
    def __init__(self):
        self.countries = ["AE", "AO", "BH", "CI", "CM", "DF", "GM", "HK", "IQ", "JO", "MU", "NG", "OM", "QA", "SG", "SL", "SP", "UK", "US", "ZW"]
        self.countries_CORE_TEAM = ["BW", "KE", "TZ", "UG", "ZM", "ZA","GH"]
        self.countries_IN = ["IN"]
        self.todays_date = datetime.now().strftime("%Y%m%d")
        
    def downloadReport(self, report_name):
        sourcePath = "C:/Suzannah/Report_03/HK.xls"
        if os.path.isfile(sourcePath) == False:
            print("No files")
        #for country in self.countries:
          #  sourcePath = "";
           # if report_name == "2": 
               
                #sourcePath = "/root/Trade DTPRGBO reports/GSSC/" + country + "/Daily/" + self.todays_date + "/" + country + "_D(17) Guarantee Outstanding Claims.xls"
                #if os.path.isfile(sourcePath) == False:
                 #   sourcePath = "/root/Trade DTPRGBO reports/GSSC/" + country + "/Daily/" + self.todays_date + "/" + country + "_(17) Guarantee Outstanding Claims.xls"
                #if os.path.isfile(sourcePath) == False:
                 #   sourcePath = "/root/Trade DTPRGBO reports/GSSC/" + country + "/Daily/" + self.todays_date + "/" + self.todays_date + "/" + country + "_D(17) Guarantee Outstanding Claims.xls"
            
        try:
            excel_report = pd.read_excel(sourcePath, sheetname=0)
            excel_report = excel_report.drop(0)
            headers = excel_report.iloc[0]
            excel_report = pd.DataFrame(excel_report.values[1:], columns = headers)
            
            #print(excel_report)
            
            for x in range(0, len(excel_report)):
                if excel_report.at[x,'CLAIM_NO'] != 11 and excel_report.at[x,'CLAIM_NO'] != 1 and excel_report.at[x,'CLAIM_NO'] != 2:
                    excel_report = excel_report.drop(excel_report.at[x,5])
            
            #print(excel_report) 
            
            with ExcelWriter("C:/Suzannah/Report_03/Hello.xlsx") as writer:
                excel_report.to_excel(writer)
                writer.save()
                
        except Exception:
            print("Oops, we've encountered an error - ", sys.exc_info()[0])

obj = ReportDownload()
obj.downloadReport("2")


    