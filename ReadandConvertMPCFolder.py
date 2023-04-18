# -*- coding: utf-8 -*-
"""
__created__ = 20220721
__authour__ = 60208787
__ver__ = 0.0.2
"""

import modules.classy as classy

import logging
import glob, os, yaml
from tqdm import tqdm
import pandas as pd
import openpyxl
import numpy as np


def makde_df_from_openpyxl(openpyxl_file, sheetname):
    try:
        values = pd.DataFrame(openpyxl_file[sheetname].values)
    except KeyError:
        print(f'{sheetname} sheet not found')
    else:   
        values = values.set_index(values.columns[0])
        values.columns = values.iloc[0]
        values = values[1:].loc[:,:'Value'].dropna()
        return values

def processing_MPC_folders(config):
    
    # Log file assumed to live at top level of folder
    logfile = "logfile_mpc_processed.txt" # f"{config['parent_path']}/
    
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
        with open(logfile, 'a') as f:
            for i in tqdm(list_of_MPC_folders, smoothing=0.2):
                
                ### Try to generate an object for each folder...if fail (like no results csv),
                ### then add the folder to a failed count
                try:
                    #Custom MPC module for these objects
                    MPC_obj = classy.MPC_results(i)
                    MPC_obj.write_MPC_to_MyQAFolder(f"{config['root_results_path']}/{config['number_in_results_path']} {machine_name}/MPC/Raw")
                except:
                    failed_folders_count += 1
                    mylogger.error(f"{i} failed to read MPC")
                else:
                    f.write(f"{i}\n")     
            
        print(f"{failed_folders_count} failed in {len(list_of_MPC_folders)} from {machine_name} as no results CSV...check manually \n")


def processing_results_files(config):
    
    
    # Log file assumed to live at top level of folder
    logfile = "logfile_myQA_processed.txt" # f"{config['parent_path']}/


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
        with open(logfile, 'a') as f:
            for file in tqdm(list_of_results_files, smoothing=0):

                mylogger.info(file)
                file_to_write_to = file.replace('\\','/').replace('/Raw','/Results') 
                
                results_file = openpyxl.load_workbook(file)

                check_names = set(results_file.sheetnames)
                check_temp = set(template.sheetnames)

                with pd.ExcelWriter(file_to_write_to) as writer:
                    # For each sheet *actually* in the MPC results file
                    for sheet in results_file.sheetnames:


                        ref_date = results_file[results_file.sheetnames[0]].cell(2,2).value

                        
                        # Case handing of actual sheets names
                        # Could do this in a ResultsFile object to keep main tidy, but eh
                        if sheet == '6xMVkVEnhancedCouch' and '6xMVkV' in results_file.sheetnames:
                            
                            #Create new df from existing workbook using function to keep more tidy
                            template_df = makde_df_from_openpyxl(template, '6xMVkV')
                            
                            #As book is generator not list, need to force past the first item/val to extract vals in loop
                            try:
                                values_6x = makde_df_from_openpyxl(results_file, '6xMVkV')
                                
                                #As book is generator not list, need to force past the first item/val to extract vals in loop
                                values_6xext = makde_df_from_openpyxl(results_file, '6xMVkVEnhancedCouch')
                            except AttributeError:
                                mylogger.error('Seems to not have generated values df...break')
                                break
                            else:
                                for item in template_df.index:
                                    if 'Enhanced' in item:
                                        try:
                                            template_df.loc[item] = values_6xext.loc[item]
                                        except KeyError:
                                            mylogger.warning(f"{item} doesn't exist in sheet")
                                            
                                    else:
                                        try:
                                            template_df.loc[item] = values_6x.loc[item]
                                        except KeyError:
                                            mylogger.warning(f"{item} doesn't exist in sheet")
                                            
                                check_names.add('6xMVkV')
                                template_df.to_excel(writer,sheet_name='6xMVkV')
                                # results_file.remove(results_file['6xMVkVEnhancedCouch'])
                

                            
                        elif sheet == '6xMVkVEnhancedCouch' and '6xMVkV' not in results_file.sheetnames:
                        
                            template_df = makde_df_from_openpyxl(template, '6xMVkV')
        

                            try:
                                values = makde_df_from_openpyxl(results_file, '6xMVkVEnhancedCouch')
                                
                            except AttributeError:
                                mylogger.error('Seems to not have generated values df...break')
                                break
                            else:
                                for item in template_df.index:
                                        try:
                                            template_df.loc[item] = values.loc[item]
                                        except KeyError:
                                            mylogger.warning(f"{item} doesn't exist in sheet")
                                check_names.add('6xMVkV')
                                template_df.to_excel(writer,sheet_name="6xMVkV")

                        elif sheet == '6x':
                                ### THIS IS THE ENHANCED MLC FILES
                            template_df = makde_df_from_openpyxl(template, '6x_MLC')

        
                            try:
                                values = makde_df_from_openpyxl(results_file, sheet)
                                
                            except AttributeError:
                                mylogger.error('Seems to not have generated values df...break')
                                break
                            else:
                                for item in template_df.index:
                                        try:
                                            template_df.loc[item] = values.loc[item]
                                        except KeyError:
                                            mylogger.warning(f"{item} doesn't exist in sheet")
                                check_names.add("6x_MLC")

                                template_df.to_excel(writer,sheet_name="6x_MLC")
                            

                            
                        elif sheet in template.sheetnames:
                            template_df = makde_df_from_openpyxl(template, sheet)
        
                            try:
                                values = makde_df_from_openpyxl(results_file, sheet)
                                
                            except AttributeError:
                                mylogger.error('Seems to not have generated values df...break')

                                break
                            else:
                                for item in template_df.index:
                                        try:
                                            template_df.loc[item] = values.loc[item]
                                        except KeyError:
                                            mylogger.warning(f"{item} doesn't exist in sheet")
                                template_df.to_excel(writer,sheet_name=sheet)
                    else:
                        mylogger.warning(f"{sheet} not in template")
                        pass
                    
                    temp_not_in_result_file = check_temp-check_names

                    for sheet in temp_not_in_result_file:
                        template_df = makde_df_from_openpyxl(template, sheet)
                        template_df.to_excel(writer,sheet_name=sheet)


                f.write(f"{file}\n") 
                # log.add_processed_folder_to_log(file)

        
class MyFilter(object):
    def __init__(self, level):
        self.__level = level

    def filter(self, logRecord):
        return logRecord.levelno == self.__level
            
if __name__ == '__main__':
        
    #Assumes config is in same folder as main script
    try:
        with open('config.yaml','r') as stream:
            config = yaml.safe_load(stream)
    except:
        pass

    # Set up logger at top level of script
    mylogger = logging.getLogger('mylogger')

    # Handler two is for writing out error messages files
    handler2 = logging.FileHandler("dev.log",mode='w')
    #SEND THROUGH EVERYTHING ABOVE INFO HERE
    handler2.setLevel(logging.INFO)
    handler2.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
    mylogger.addHandler(handler2)

    #SET GLOBAL LEVEL TO INFO
    mylogger.setLevel(logging.INFO)

    print('\n Starting checks of MPC Folders \n')
    processing_MPC_folders(config)
    
    print('\n Starting processing of data for myQA \n')
    processing_results_files(config)