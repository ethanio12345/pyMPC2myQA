# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 11:46:28 2022

@author: 60208787
"""

import datetime
import numpy as np
import pandas as pd
import os
import time
import csv



class MPC_results:
    def __init__(self, folder_path, results_path = None, machine = None, date = None, datetime = None, measurement_type = None, beam_energy = None, results = None, passed = True):
        
        self.path = os.path.abspath(folder_path)
        self.file_name = os.path.basename(self.path)
        

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

        if self.check_for_results:
            self.read_results()
        
    @property
    def check_for_results(self):
        self.results_path = f"{self.folder_path}/Results.csv"

        try:
            with open(self.results_path, 'r') as f:
                results_found = True
        except IOError:
            pass

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
                self.measurement_type = folder_property[10].split("Template")[0]     ## Beam Check/Geometry Check/Enhanced Couch or in TB2.7 EnhancedMLC or TB 3.0 CollimationDevices Check. 
                self.beam_energy = folder_property[10].split("Template")[-1]
            elif self.measurement_type == "GeometryCheck":
                self.beam_energy = self.beam_energy + "MVkV"
            else:
                self.measurement_type = folder_property[11]+"Check"
                self.beam_energy = folder_property[10]
    
    
    def read_results(self):
        results_path = self.results_path
        results = {}

        with open(results_path, newline="\n") as csv_file:
            
            next(csv_file)
            lines = csv_file.readlines()

            for row in lines:
                
                total_name = row.split(',')[0].split(" [")[0].split("/")
                test_name = total_name[-1]
                
                if "MLCLeaf" in test_name:
                    test_name = f"{total_name[-2][-1]}-{test_name}"
                elif "MLCBacklashLeaf" in test_name:
                    test_name = f"{total_name[-2][-1]}-{test_name}"

                if "CollimationDevicesGroup" == total_name[0]:
                    if len(total_name) == 3:
                        test_name = total_name[-1]
                    elif len(total_name) == 4:
                        test_name = f"{total_name[-2][-1]}-{total_name[-1]}"
                    elif len(total_name) == 5:
                        test_name = f"Pos{total_name[-3][-1]}-Bank{total_name[-2][-1]}-{total_name[-1]}"
                elif "EnhancedCouchGroup" == total_name[0]:
                        test_name = f"EnhancedCouch{total_name[-1]}"

                results[test_name] = float(row.split(',')[1])
                if row.split(',')[-1] == "Failed":
                    self.passed = False
        self.results = results
    
        
    
    def write_MPC_to_MyQAFolder(self, MyQAFolder):
    
        # Give list of MPC folders for single day
        # Generate MPC results for each check for each day
        
        xlsx_path = os.path.join(MyQAFolder,f"Results_SN{self.machine[-4:]}_{self.datetime.strftime('%Y%m%d-%H_%M')}.xlsx")
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



if __name__ == '__main__':
    mpc = MPC_results(r'V:\LA4\TDS\H191733\MPCChecks\NDS-WKS-SN1733-2023-03-24-07-15-09-0010-CollimationDevicesCheckTemplate6x')
    mpc.process_folder()
    mpc.read_results()
    print("ready and waiting")
    # mpc.write_MPC_to_MyQAFolder(r'V:\01 Physics Clinical QA\05 LA4')