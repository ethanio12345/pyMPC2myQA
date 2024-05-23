# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 11:46:28 2022

@author: 60208787
"""

import datetime
import pandas as pd
import os, csv
from pathlib import Path



class MPC_results:
    def __init__(self, folder_path, results_path = None, machine = None, date = None, datetime = None, measurement_type = None, beam_energy = None, passed = None):
        
        # From ReadMPC
        self.folder_path = str(Path(folder_path).resolve())
        self.has_results_file = False

        if not Path(f"{self.folder_path}/Results.csv").is_file():
            raise FileNotFoundError
        else:
            self.has_results_file = True

        
        self.machine = machine
        self.date = date
        self.datetime = datetime
        
        self.measurement_type = measurement_type
        self.beam_energy = beam_energy
        self.passed = passed        
        self.machineSN = {
            "2361": "Coast: 2361",
            "2362": "Outback: 2362",
            "2972": "LA317: 2972",
            "1733": "LA414: 1733",
            "1182": "LA512: 1182",            
            "6406": "LA224: 6406",
        }
        
        self.process_folder()

    
    def process_folder(self):
        mpc_folder_path = self.folder_path.split("\\")[-1]
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

        with open(f"{self.folder_path}/Results.csv", 'r') as f:
            csv_data = csv.reader(f,delimiter=",")
            next(csv_data)
            results = {}

            for row in csv_data:
                
                total_name = row[0].split(" [")[0].split("/")
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

                results[test_name] = float(row[1])
                if row[-1] == "Failed":
                    self.passed = False
        return results

        
    
    def write_MPC_to_MyQAFolder(self, MPC_results, MyQAFolder):
    
        # Give list of MPC folders for single day
        # Generate MPC results for each check for each day
        
        xlsx_path = os.path.join(MyQAFolder,f"Results_SN{self.machine[-4:]}_{self.datetime.strftime('%Y%m%d-%H_%M')}.xlsx")
        MPC_df = pd.DataFrame(columns=['Value'])
        MPC_df.loc['Reference Date'] = self.datetime
        MPC_df = pd.concat([MPC_df, pd.DataFrame.from_dict(MPC_results,orient='index',columns=['Value'])])
        if not os.path.isfile(xlsx_path):
            # Use writer mode if file doesn't exist
            with pd.ExcelWriter(xlsx_path, engine='openpyxl',mode='w') as writer:        
                MPC_df.to_excel(writer,sheet_name=self.beam_energy)
        else:
            # Use append mode if file does exist
    
            with pd.ExcelWriter(xlsx_path, engine='openpyxl',mode='a',if_sheet_exists='replace') as writer:
                MPC_df.to_excel(writer,sheet_name=self.beam_energy)



if __name__ == '__main__':
    # mpc = MPC_results(r'//weshonc-afs0001/va_transfer/LA5/TDS/H191182/MPCChecks\NDS-WKS-SN1182-2024-05-13-15-03-46-0003-GeometryCheckTemplate6xMVkV')
    mpc = MPC_results(r'V:\LA5\TDS\H191182\MPCChecks\NDS-WKS-SN1182-2024-05-13-15-03-46-0001-BeamCheckTemplate10xFFF')
    results = mpc.read_results()
    MPC_df = pd.DataFrame(columns=['Value'])
    MPC_df.loc['Reference Date'] = "Test"
    MPC_df = pd.concat([MPC_df, pd.DataFrame.from_dict(results,orient='index',columns=['Value'])])
    print(MPC_df) # debugging
    print(results) # debugging
    print("ready and waiting")
    # mpc.write_MPC_to_MyQAFolder(r'V:\01 Physics Clinical QA\05 LA4')