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
                MPC_obj.write_MPC_to_MyQAFolder(f"{config['root_results_path']}/{config['number_in_results_path']} {machine_name}/MPC")
                                                
                log.add_processed_folder_to_log(i)
            except:
                failed_folders_count += 1
            
        print(f"{failed_folders_count} failed in {len(list_of_MPC_folders)} as no results CSV...check manually \n")


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
        list_of_results_files = sorted([x for x in glob.glob(f"{config['root_results_path']}/{config['number_in_results_path']} {machine}/MPC/Results_*.xlsx")
                                        if x not in loglist],
                                        reverse=True)
        
        # Read in template xltx for MyQA and reuse here
        template = openpyxl.load_workbook(f"{config['parent_path']}/Results/Template.xltx")
        
        for file in tqdm(list_of_results_files):
            try:
                # Append mode (assumes file exists already) and replace sheet with new values
                with pd.ExcelWriter(file, engine='openpyxl',mode='a',if_sheet_exists='replace') as writer:
                    # For each sheet *actually* in the results file
                    for sheet in writer.book.sheetnames:
                        
                        #After fine sheet in loop, take ref date to write in
                        ref_date = writer.book[sheet].cell(2,2).value
                        
                        # Case handing of actual sheets names
                        # Could do this in a ResultsFile object to keep main tidy, but eh
                        if sheet == '6xMVkVEnhancedCouch' and '6xMVkV' in writer.book.sheetnames:
                            
                            #Create new df from existing workbook using function to keep more tidy
                            template_df = make_df_from_template(template, '6xMVkV')
                            
                            #As book is generator not list, need to force past the first item/val to extract vals in loop
                            values_6x = pd.DataFrame(writer.book['6xMVkV'].values)
                            # values_6x.set_index(values_6x.columns[0])
                            values_6x.columns = values_6x.iloc[0]
                            values_6x = values_6x.drop(0)
                            

                            
                            #As book is generator not list, need to force past the first item/val to extract vals in loop
                            values_6xext = pd.DataFrame(writer.book['6xMVkVEnhancedCouch'].values)
                            # values_6xext.set_index(values_6xext.columns[0])
                            values_6xext.columns = values_6xext.iloc[0]
                            values_6xext = values_6xext.drop(0)
                            
                            for i,vals in template_df.values:
                                if 'Enhanced' in vals[0]:
                                    template_df.loc[i,vals[0]] = values_6xext[i,vals[0]]
                                else:
                                    template_df.loc[i,vals[0]] = values_6x[i,vals[0]]

                            template_df.to_excel(writer,sheet_name='6xMVkV')
                            writer.book.remove(writer.book['6xMVkVEnhancedCouch'])
                

                         
                        # elif sheet == '6xMVkVEnhancedCouch' and '6xMVkV' not in writer.book.sheetnames:
                        
        
                        #     template_df = make_df_from_template(template, '6xMVkV')
                    
                        #     #As book is generator not list, need to force past the first item/val to extract vals in loop
                        #     itervals = iter(writer.book['6xMVkVEnhancedCouch'].values)
                        #     next(itervals)
                            
                        #     for item,val in itervals:
                        #         template_df.loc[item] = val
                        
                        #     template_df.to_excel(writer,sheet_name='6xMVkV')
                        #     writer.book.remove(writer.book['6xMVkVEnhancedCouch'])
                    
                        # elif sheet == '6x':
        
                        #     template_df = make_df_from_template(template, '6x_MLC')

                        #     #As book is generator not list, need to force past the first item/val to extract vals in loop
                        #     itervals = iter(writer.book[sheet].values)
                        #     next(itervals)
                            
                        #     for item,val in itervals:
                        #         template_df.loc[item] = val
                                
                        #     template_df.to_excel(writer,sheet_name='6x_MLC')
                        #     writer.book.remove(writer.book['6x'])
                            
                        # elif sheet == '6x_MLC' and '6x' in writer.book.sheetnames:
                            
                        #     template_df = make_df_from_template(template, '6x')
        
                        #     #As book is generator not list, need to force past the first item/val to extract vals in loop
                        #     itervals = iter(writer.book[sheet].values)
                        #     next(itervals)
                            
                        #     for item,val in itervals:
                        #         template_df.loc[item] = val
                                
                        #     template_df.to_excel(writer,sheet_name='6x_MLC')
                        #     writer.book.remove(writer.book['6x'])
                            
                        # elif sheet in template.sheetnames:
                            template_df = make_df_from_template(template, sheet)
        
        
                            #As book is generator not list, need to force past the first item/val to extract vals in loop
                            itervals = iter(writer.book[sheet].values)
                            next(itervals)
                            
                            for item,val in itervals:
                                template_df.loc[item] = val
                                
                            template_df.to_excel(writer,sheet_name=sheet)
        
                        else:
                            pass
                            # print('Unknown sheet name')
                        
                    #Pre-process file to ensure all relevant tabs are there...don't know if separation necessary
        
                    for sheet in template.sheetnames:
                        
                        if sheet not in writer.book.sheetnames:
                            
                            template_df = make_df_from_template(template, sheet)
                            template_df.loc['Reference Date'] = ref_date                            
                            template_df.to_excel(writer,sheet_name=sheet)
                    # Finished processing results file and if everything ok by this stage, add file to log
                    log.add_processed_folder_to_log(file)
            except:
                print('Failed')

        

            
if __name__ == '__main__':
    
    #Assumes config is in same folder as main script
    try:
        with open('config.yaml','r') as stream:
            config = yaml.safe_load(stream)
    except:
        pass
    print('\n Starting checks of MPC Folders \n')

    # processing_MPC_folders(config)
    
    print('\n Starting processing of data for myQA \n')
    processing_results_files(config)



