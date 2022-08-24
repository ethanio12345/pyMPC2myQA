# -*- coding: utf-8 -*-

__created__ = 20220721
__authour__ = 60208787

import modules.classy as classy

import glob
import os
from tqdm import tqdm
import pandas as pd
import openpyxl


# This here for if working on local desktop (minimise network traffic)
local_parent_path: "C:/Users/60208787/Downloads/pyMPC2MyQA"

# True parent path for where to work on/save results to
parent_path: "P:/4. Software/Code Development/Python Projects/pyMPC2myQA"

# Points to VA_TRANSFER network location
root_va_transfer_path: "//weshonc-afs0001/va_transfer"
old_root_va_transfer_path: "//10.8.232.180/va_transfer"               

def processing_MPC_folders():
    
  
    myQAFolder = parent_path + '/Results'
    
    # Points to VA_TRANSFER network location
    root_va_transfer_path = '//weshonc-afs0001/va_transfer'
    old_root_va_transfer_path = '//10.8.232.180/va_transfer'

    machine_paths = ["/LA5/TDS/H191182/MPCChecks",
                     "/CST/TDS/H192361/MPCChecks",
                     "/OBK/TDS/H192362/MPCChecks",
                     "/LA3/TDS/H192972/MPCChecks",
                     "/LA4/TDS/H191733/MPCChecks",
                     ]
  
    ########## THIS TO BE UPDATED WITH NEW SERVER SYSTEMS ####################
    machine_folder_paths = [root_va_transfer_path + x for x in machine_paths]
    
    logfile = parent_path +'/logfile_mpc_processed.txt'
    log = classy.LogResults(logfile)
    
    # Log file above is good for individual checks, but list here is faster
    with open(logfile, 'r') as f:
        loglist = f.read().splitlines()
    
    for machine in machine_folder_paths: 
        
        #Search and sort for all folders under each machine path
        list_of_MPC_folders = sorted([f.path for f in os.scandir(machine) 
                              if f.is_dir() 
                              if 'NDS-WKS-SN' in f.path 
                              if f.path not in loglist],
                              reverse=True)
        
        # Only using this as indicator, not as tracking failed files
        failed_folders_count = 0 
        for i in tqdm(list_of_MPC_folders):
            
            ### Try to generate an object for each folder...if fail (like no results),
            ### then add the folder to a failed count
            try:
                MPC_obj = classy.MPC_results(i)
                MPC_obj.write_MPC_to_MyQAFolder(myQAFolder)
                log.add_processed_folder_to_log(i)
            except:
                failed_folders_count += 1
            
        print('{} Failed in {} as no results CSV...check manually \n'.format(failed_folders_count,machine))

def make_df_from_template(template_openpyxl, sheetname):
    template_df = pd.DataFrame(data = template_openpyxl[sheetname].values).set_index(0)
    template_df.columns = template_df.iloc[0]
    
    return template_df[1:]

def processing_results_files():
    
    myQAFolder = parent_path + '\Results'
    
    logfile = parent_path+'\logfile_myQA_processed.txt'
    log = classy.LogResults(logfile)
    
    # Log file above is good for individual checks, but list here is faster
    with open(logfile, 'r') as f:
        loglist = f.read().splitlines()
    
    #Search and sort for all results file in Results Folder
    list_of_results_files = sorted([x for x in glob.glob(myQAFolder+"/Results_*.xlsx") 
                                    if x not in loglist],
                                   reverse=True)
    
    # Read in template file for MyQA template and reuse for template here
    template = openpyxl.load_workbook(parent_path + '/Results/Template.xltx')
    
    for file in tqdm(list_of_results_files):
        
        # Append mode and replace sheet as it should already exist
        with pd.ExcelWriter(file, engine='openpyxl',mode='a',if_sheet_exists='replace') as writer:
            
            # For each sheet *actually* in the results file
            for sheet in writer.book.sheetnames:
                
                #After loop, will take ref date from last sheet
                ref_date = writer.book[sheet].cell(2,2).value
                
                # Case handing of actual file
                if sheet == '6xMVkVEnhancedCouch' and '6xMVkV' in writer.book.sheetnames:
                    
                    #Create new df from existing workbook using 
                    # not doing this as function as 
                    template_df = make_df_from_template(template, '6xMVkV')
                    
                    #Itervals as 
                    itervals = iter(writer.book['6xMVkV'].values)
                    next(itervals)
                    
                    for item,val in itervals:
                        template_df.loc[item] = val
                    
                    itervals = iter(writer.book['6xMVkVEnhancedCouch'].values)
                    next(itervals)
                    
                    for item,val in itervals:
                        if 'Enhanced' in item:
                            template_df.loc[item] = val

                    template_df.to_excel(writer,sheet_name='6xMVkV')
                    writer.book.remove(writer.book['6xMVkVEnhancedCouch'])
        
                        # print('Enhanced Couch sheet renamed \n')
                 
                elif sheet == '6xMVkVEnhancedCouch' and '6xMVkV' not in writer.book.sheetnames:
                

                    template_df = make_df_from_template(template, '6xMVkV')
            
                    itervals = iter(writer.book['6xMVkVEnhancedCouch'].values)
                    next(itervals)
                    
                    for item,val in itervals:
                        template_df.loc[item] = val
                
                    template_df.to_excel(writer,sheet_name='6xMVkV')
                    writer.book.remove(writer.book['6xMVkVEnhancedCouch'])

                        # print('Enhanced Couch sheet renamed \n')
            
                elif sheet == '6x':

                    template_df = make_df_from_template(template, '6x_MLC')
                    itervals = iter(writer.book[sheet].values)
                    next(itervals)
                    
                    for item,val in itervals:
                        template_df.loc[item] = val
                        
                    template_df.to_excel(writer,sheet_name='6x_MLC')
                    writer.book.remove(writer.book['6x'])
                    
                elif sheet == '6x_MLC' and '6x' in writer.book.sheetnames:
                    
                    template_df = make_df_from_template(template, '6x')

                    itervals = iter(writer.book[sheet].values)
                    next(itervals)
                    
                    for item,val in itervals:
                        template_df.loc[item] = val
                        
                    template_df.to_excel(writer,sheet_name='6x_MLC')
                    writer.book.remove(writer.book['6x'])

                    # print('MLC sheet renamed \n')
                    
                elif sheet in template.sheetnames:
                    template_df = make_df_from_template(template, sheet)


                    itervals = iter(writer.book[sheet].values)
                    next(itervals)
                    
                    for item,val in itervals:
                        template_df.loc[item] = val
                        
                    template_df.to_excel(writer,sheet_name=sheet)

                        
                    # print('MLC sheet renamed \n')
                else:
                    print('Unknown sheet name')
                
            #Pre-process file to ensure all relevant tabs are there...don't know if separation necessary

            for sheet in template.sheetnames:
                if sheet not in writer.book.sheetnames:
                    
                    template_df = make_df_from_template(template, sheet)

                    
                    template_df.loc['Reference Date'] = ref_date
                    
                    template_df.to_excel(writer,sheet_name=sheet)
            
            log.add_processed_folder_to_log(file)
      

        

            
if __name__ == '__main__':
    
    try:
        with open('config.yaml','r') as stream:
            config = yaml.safe_load(stream)
    print('\n Starting checks of MPC Folders \n')

    processing_MPC_folders(config)
    
    print('\n Starting processing of data for myQA \n')
    processing_results_files(config)



