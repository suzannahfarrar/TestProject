# -*- coding: utf-8 -*-
"""
Created on Tue May 14 11:22:27 2019

@author: 1579394
"""

from datetime import datetime
import pandas as pd
import os

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
        for country in self.countries:
            sourcePath = "";
            if report_name == "2":            
                sourcePath = "/root/Trade DTPRGBO reports/GSSC/" + country + "/Daily/" + self.todays_date + "/" + country + "_D(17) Guarantee Outstanding Claims.xls"
                if os.path.isfile(sourcePath) == False:
                    sourcePath = "/root/Trade DTPRGBO reports/GSSC/" + country + "/Daily/" + self.todays_date + "/" + country + "_(17) Guarantee Outstanding Claims.xls"
                if os.path.isfile(sourcePath) == False:
                    sourcePath = "/root/Trade DTPRGBO reports/GSSC/" + country + "/Daily/" + self.todays_date + "/" + self.todays_date + "/" + country + "_D(17) Guarantee Outstanding Claims.xls"
            
            try:
                file = open(sourcePath, 'r')
                excel_report = pd.read_excel(file)
                excel_report.head()
                
            except Exception:
                print(Exception + "")

obj = ReportDownload()
    