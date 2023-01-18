# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 11:46:28 2022

@author: 60208787
"""

import datetime
import pandas as pd
import os
import time
import csv



class MPC_results:
    def __init__(self, folder_path, results_path = None, machine = None, date = None, datetime = None, measurement_type = None, beam_energy = None, results = None, passed = True):
        
        self.path = os.path.abspath(folder_path)
        self.file_name = self.path.split('\\')[-1]
        

        # From ReadMPC
        self.folder_path = folder_path
        self.results_path = results_path
        self.machine = machine
        self.date = date
        self.datetime = datetime
        self.measurement_type = measurement_type
        self.beam_energy = beam_energy
        self.results = results
        self.passed = passed        
        self.machineSN = {
            "2361": "Coast: 2361",
            "2362": "Outback: 2362",
            "2972": "LA317: 2972",
            "1733": "LA414: 1733",
            "1182": "LA512: 1182",
        }
        
        self.process_folder()
        # self.check_for_results()
        if self.check_for_results:
            self.read_results()
        
    @property
    def check_for_results(self):
        self.results_path = f"{self.folder_path}/Results.csv"

        datetime_now = datetime.datetime.now()
        result_datetime = self.datetime
        difference = datetime_now - result_datetime

        if difference.total_seconds() < 25 * 60:
            wait_time = 30  # seconds
            max_retries = 40  # maximum wait time of wait_time * max_retries (s)
        else:
            wait_time = 0  # seconds
            max_retries = 1  # maximum wait time of wait_time * max_retries (s)

        count = 0
        results_found = False
        while count < max_retries:
            try:
                with open(self.results_path, 'r') as f:
                    count = max_retries
                    results_found = True
            except IOError:
                time.sleep(wait_time)
                count += 1

        return results_found
    
    def process_folder(self):
        mpc_folder_path = self.path.split("\\")[-1]
        if "NDS-WKS-SN" in mpc_folder_path:
            folder_property = mpc_folder_path.split("-")

            machine_sn = folder_property[2].split("SN")[-1]
            self.machine = self.machineSN[machine_sn]

            year = folder_property[3]
            month = folder_property[4]
            day = folder_property[5]
            hour = folder_property[6]
            minute = folder_property[7]
            second = folder_property[8]

            self.date = "{}-{}-{} {}:{}:{}".format(year,month,day,hour,minute,second)
            self.datetime = datetime.datetime(int(year),int(month),int(day),int(hour),int(minute),int(second))

            if "Template" in mpc_folder_path:
                self.measurement_type = folder_property[10].split("Template")[0]     ## Beam Check/Geometry Check/Enhanced Couch/MLC
                self.beam_energy = folder_property[10].split("Template")[-1]
            elif self.measurement_type == "GeometryCheck":
                self.beam_energy = self.beam_energy + "MVkV"
            else:
                self.measurement_type = folder_property[11]+"Check"
                self.beam_energy = folder_property[10]
    
    
    def read_results(self):
        results_path = self.results_path
        results = {}
        with open(results_path) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            i = 0
            for row in csv_reader:
                if i > 0:
                    test_name = (row[0].split("/")[-1]).split(" [")[0]
                    if "MLCLeaf" in test_name:
                        test_name = (row[0].split("/")[-2])[-1]+"-"+test_name
                    elif "MLCBacklashLeaf" in test_name:
                        test_name = (row[0].split("/")[-2])[-1]+"-"+test_name
                    elif "EnhancedCouch" in row[0].split("/")[0]:
                        test_name = "EnhancedCouch"+(row[0].split("/")[-1]).split(" [")[0]
                        
                    results[test_name] = float(row[1])
                    if row[3] == "Failed":
                        self.passed = False
                i = i + 1
        self.results = results
    
        
    
    def write_MPC_to_MyQAFolder(self, MyQAFolder):
    
        # Give list of MPC folders for single day
        # Generate MPC results for each check for each day
        
        xlsx_path = f"{MyQAFolder}/Results_SN{self.machine[-4:]}_{self.datetime.strftime('%Y%m%d-%H_%M')}.xlsx"
        MPC_df = pd.DataFrame(columns=['Value'])
        MPC_df.loc['Reference Date'] = self.datetime
        MPC_df = pd.concat([MPC_df, pd.DataFrame.from_dict(self.results,orient='index',columns=['Value'])])
        
        if not os.path.isfile(xlsx_path):
            # Use writer mode if file doesn't exist
            with pd.ExcelWriter(xlsx_path, engine='openpyxl',mode='w') as writer:        
                MPC_df.to_excel(writer,sheet_name=self.beam_energy)
        else:
            # Use append mode if file does exist
    
            with pd.ExcelWriter(xlsx_path, engine='openpyxl',mode='a',if_sheet_exists='replace') as writer:
                MPC_df.to_excel(writer,sheet_name=self.beam_energy)

class LogResults:
    def __init__(self,logfile_path, logfile = None):
        self.logfile_path = logfile_path
        self.logfile = logfile

    def open_log_r(self):
        try:
            self.logfile = open(self.logfile_path, 'r')
        except IOError:
            self.open_log_a()
            self.close_log()
            self.logfile = open(self.logfile_path, 'r')

    def open_log_a(self):
        self.logfile = open(self.logfile_path, 'a+')

    def write_to_log(self,new_result):
        self.logfile.write(new_result+"\n")

    def close_log(self):
        self.logfile.close()

    def check_if_previously_processed(self,new_result):
        self.open_log_r()
        loglist = self.logfile.readlines()
        found = False

        for line in loglist:
            if new_result in line:
                found = True
                # print('New Result found')

        self.close_log()
        return found

    def add_processed_folder_to_log(self,new_result):
        self.open_log_a()
        self.write_to_log(new_result)
        self.close_log()

