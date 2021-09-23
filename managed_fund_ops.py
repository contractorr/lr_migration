##############################################################################################################
import log 
logger = log.get_logger('root')
import glob
import os 
import mig_functions as mig

# Destination of key files
os.chdir('C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrate These')  
input_files = glob.glob('[!~]*.xlsx')         
dest_directory = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrated'
##############################################################################################################

# Start logging
if os.path.exists(f'Managed.log'): 
    os.remove(f'Managed.log')

log.update_handler(logger,'Managed')

for input_file in input_files:
    # Add the file name to the dest_directory and log_file_directory
    name = input_file.split('@')[1]
    dest_file = dest_directory + f'/Managed Fund Operations - {name} - RC Migrated.xlsx'
    print(f'File: {input_file}')   
    mig.migrate_managed_data(input_file, dest_file, env='UAT')
    #mig.migrate_managed_data(input_file, dest_file,True)



    

    
    
    