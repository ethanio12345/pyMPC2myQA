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

def make_df_from_template(template_openpyxl, sheetname):
    template_df = pd.DataFrame(data = template_openpyxl[sheetname].values).set_index(0)
    template_df.columns = template_df.iloc[0]
    
    return template_df[1:]

def processing_MPC_folders(config):
    
    
    logfile = '/'.join([config['parent_path'], 'logfile_mpc_processed.txt'])
    log = classy.LogResults(logfile)
    
    # Log file above is good for individual checks, but list here is faster
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
            
            ### Try to generate an object for each folder...if fail (like no results),
            ### then add the folder to a failed count
            try:
                MPC_obj = classy.MPC_results(i)
                MPC_obj.write_MPC_to_MyQAFolder('/'.join([config['root_results_path'],
                                                          '/{} {}'.format(config['number_in_results_path'], machine_name),
                                                          '/MPC']
                                                         )
                                                )
                log.add_processed_folder_to_log(i)
            except:
                failed_folders_count += 1
            
        print('{} Failed in {} as no results CSV...check manually \n'.format(failed_folders_count,machine))


def processing_results_files(config):
    
    
    
    logfile = config['parent_path']+'/logfile_myQA_processed.txt'
    
    log = classy.LogResults(logfile)
    
    # Log file above is good for individual checks, but list here is faster
    with open(logfile, 'r') as f:
        loglist = f.read().splitlines()
        

    
    #Search and sort for all results file in Results Folder
    for machine in config['machines']:
        
        print("Assessing {} myQA results".format(machine))
        
        # There is a {machine} within results_folder_path, so machine=machine to apply
        list_of_results_files = sorted([x for x in glob.glob('/'.join([config['results_folder_path'].format(machine=machine),
                                                                       'Results_*.xlsx']))
                                        if x not in loglist],
                                       reverse=True)
        
        # Read in template file for MyQA template and reuse for template here
        template = openpyxl.load_workbook(config['parent_path'] + '/Results/Template.xltx')
        try:
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
        except:
            print('Failed')

        

            
if __name__ == '__main__':
    
    try:
        with open('config.yaml','r') as stream:
            config = yaml.safe_load(stream)
    except:
        pass
    print('\n Starting checks of MPC Folders \n')

    processing_MPC_folders(config)
    
    print('\n Starting processing of data for myQA \n')
    processing_results_files(config)



