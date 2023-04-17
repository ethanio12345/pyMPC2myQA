# -*- coding: utf-8 -*-
"""
__created__ = 20220721
__authour__ = 60208787
__ver__ = 0.0.2
"""

import modules.classy as classy

import glob, os, yaml
from tqdm import tqdm
import pandas as pd
import openpyxl
import numpy as np


def make_df_from_template(template_openpyxl, sheetname):
    template_df = pd.DataFrame(data = template_openpyxl[sheetname].values).set_index(0)
    template_df.columns = template_df.iloc[0]
    
    return template_df[1:]

def processing_MPC_folders(config):
    
    # Log file assumed to live at top level of folder
    logfile = f"{config['parent_path']}/logfile_mpc_processed.txt"
    log = classy.LogResults(logfile)
    
    # Log file above is good for individual checks, but list here is faster for checking if files exist
    with open(logfile, 'r') as f:
        loglist = f.read().splitlines()
    
    for machine in config['machine_paths']: 
        
        machine_path = config['root_va_transfer_path'] + machine
        machine_name = machine.split('/')[1]
        #Search and sort for all folders under each machine path
        list_of_MPC_folders = sorted([f.path for f in os.scandir(machine_path) 
                              if f.is_dir() 
                              if 'NDS-WKS-SN' in f.path
                              if f.path not in loglist],
                              reverse=True)
        
        # Only using this as indicator, not as tracking failed files
        failed_folders_count = 0 
        for i in tqdm(list_of_MPC_folders):
            
            ### Try to generate an object for each folder...if fail (like no results csv),
            ### then add the folder to a failed count
            try:
                #Custom MPC module for these objects
                MPC_obj = classy.MPC_results(i)
                MPC_obj.write_MPC_to_MyQAFolder(f"{config['root_results_path']}/{config['number_in_results_path']} {machine_name}/MPC/Raw")
                with open(logfile, 'r') as f:
                    loglist = f.write(i+"\n")                             
                # log.add_processed_folder_to_log(i)
            except:
                failed_folders_count += 1
            
        print(f"{failed_folders_count} failed in {len(list_of_MPC_folders)} from {machine_name} as no results CSV...check manually \n")


def processing_results_files(config):
    
    
    # Log file assumed to live at top level of folder
    logfile = f"{config['parent_path']}/logfile_myQA_processed.txt"
    log = classy.LogResults(logfile)
    
    # Log file above is good for individual checks, but list here is faster
    with open(logfile, 'r') as f:
        loglist = f.read().splitlines()
        

    for machine in config['machines']:
        
        print(f"Processing {machine} myQA results")
        
        # There is a '{machine}' within results_folder_path variable, hence machine=machine to apply
        list_of_results_files = sorted([x for x in glob.glob(f"{config['root_results_path']}/{config['number_in_results_path']} {machine}/MPC/Raw/Results_*.xlsx")
                                        if x not in loglist],
                                        reverse=True)
        
        # Read in template xltx for MyQA and reuse here
        template = openpyxl.load_workbook(f"{config['parent_path']}/Results/Template.xltx")
        
        for file in tqdm(list_of_results_files):
            file_to_write_to = file.replace('\\','/').replace('/Raw','') 
            print(file)
            print(file_to_write_to)
            results_file = openpyxl.load_workbook(file)
                # Append mode (assumes file exists already) and replace sheet with new values
            with pd.ExcelWriter(file_to_write_to) as writer:
                    
                    # For each sheet *actually* in the MPC results file
                    for sheet in results_file.sheetnames:

                        #After first sheet in loop, take ref date to write in
                        ref_date = results_file[sheet].cell(2,2).value
                        
                        # Case handing of actual sheets names
                        # Could do this in a ResultsFile object to keep main tidy, but eh
                        if sheet == '6xMVkVEnhancedCouch' and '6xMVkV' in results_file.sheetnames:
                            
                            #Create new df from existing workbook using function to keep more tidy
                            template_df = make_df_from_template(template, '6xMVkV')
                            
                            #As book is generator not list, need to force past the first item/val to extract vals in loop
                            try:
                                values_6x = pd.DataFrame(results_file['6xMVkV'].values)
                            except KeyError:
                                print(f'{sheet} sheet not found')
                                break
                            values_6x = values_6x.set_index(values_6x.columns[0])
                            values_6x.columns = values_6x.iloc[0]
                            values_6x = values_6x[1:]

                            
                            #As book is generator not list, need to force past the first item/val to extract vals in loop
                            try:
                                values_6xext = pd.DataFrame(results_file['6xMVkVEnhancedCouch'].values)
                            except KeyError:
                                print(f'{sheet} sheet not found')
                                break
                            values_6xext = values_6xext.set_index(values_6xext.columns[0])
                            values_6xext.columns = values_6xext.iloc[0]
                            values_6xext = values_6xext[1:]
                            
                            for item in template_df.index:
                                if 'Enhanced' in item:
                                    try:
                                        template_df.loc[item] = values_6xext.loc[item]
                                    except KeyError:
                                        print("Item doesn't exist in sheet")
                                        
                                else:
                                    try:
                                        template_df.loc[item] = values_6x.loc[item]
                                    except KeyError:
                                        print("Item doesn't exist in sheet")
                                        

                            template_df.to_excel(writer,sheet_name='6xMVkV')
                            results_file.remove(results_file['6xMVkVEnhancedCouch'])
                

                         
                        elif sheet == '6xMVkVEnhancedCouch' and '6xMVkV' not in results_file.sheetnames:
                        
                            template_df = make_df_from_template(template, '6xMVkV').loc[:,:'Value'].dropna()
        
                            try:
                                values = pd.DataFrame(results_file['6xMVkVEnhancedCouch'].values)
                            except KeyError:
                                print(f'{sheet} sheet not found')
                                break                        
                            values = values.set_index(values.columns[0])
                            values.columns = values.iloc[0]
                            values = values[1:]
                            values = values.loc[:,:'Value'].dropna()

                            
                            for item in template_df.index:
                                try:
                                    template_df.loc[item] = values.loc[item]
                                except KeyError:
                                    print('Some item either isn\'t in the template or the MPC sheet')
                                
                                
                            template_df.to_excel(writer,sheet_name=sheet)
                        
                            template_df.to_excel(writer,sheet_name='6xMVkV')
                            results_file.remove(results_file['6xMVkVEnhancedCouch'])
                    
                        elif sheet == '6x':
                             ### THIS IS THE ENHANCED MLC FILES
                            template_df = make_df_from_template(template, '6x_MLC').loc[:,:'Value'].dropna()

        
                            try:
                                values = pd.DataFrame(results_file['6x'].values)
                            except KeyError:
                                print(f'{sheet} sheet not found')
                                break                            
                            values = values.set_index(values.columns[0])
                            values.columns = values.iloc[0]
                            values = values[1:]
                            values = values.loc[:,:'Value'].dropna()

                            
                            for item in template_df.index:
                                try:
                                    template_df.loc[item] = values.loc[item]
                                except KeyError:
                                    print('Some item either isn\'t in the template or the MPC sheet')
                                
                            template_df.to_excel(writer,sheet_name=sheet)

                            results_file.remove(results_file['6x'])
                            
                        elif sheet == '6x_MLC' and '6x' in results_file.sheetnames:
                            
                            results_file.remove(results_file['6x'])
                            
                        elif sheet in template.sheetnames:
                            template_df = make_df_from_template(template, sheet).loc[:,:'Value'].dropna()
        
                            try:
                                values = pd.DataFrame(results_file['sheet'].values)
                            except KeyError:
                                print(f'{sheet} sheet not found')
                                break                            
                            values = values.set_index(values.columns[0])
                            values.columns = values.iloc[0]
                            values = values[1:]
                            values = values.loc[:,:'Value'].dropna()

                            for item in template_df.index:
                                try:
                                    template_df.loc[item] = values.loc[item]
                                except KeyError:
                                    print('Some item either isn\'t in the template or the MPC sheet')
                                
                            template_df.to_excel(writer,sheet_name=sheet)
        
                        else:
                            print('Unknown sheet name')
                            pass


                    log.add_processed_folder_to_log(file)


        

            
if __name__ == '__main__':
    
    #Assumes config is in same folder as main script
    try:
        with open('config.yaml','r') as stream:
            config = yaml.safe_load(stream)
    except:
        pass
    print('\n Starting checks of MPC Folders \n')

    processing_MPC_folders(config)
    
    # print('\n Starting processing of data for myQA \n')
    # processing_results_files(config)



