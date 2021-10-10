##############################################################################################################
import log 
logger = log.get_logger('root')
import glob
import os 
import mig_functions as mig
import pandas as pd

# Destination of key files
os.chdir(r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrate These')
input_files = glob.glob('[!~]*.xlsx')              
dest_directory = r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrated'
# Get FX Rates
fx_rates_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/2 UAT Import Files/01 FX Rates.xlsx'
fx_rates = pd.read_excel(fx_rates_file, index_col=None,header=2, usecols='C:Q')[1:]
fx_rates.rename({'Rates.Destination Currency': 'dst_ccy'
                ,'Rates.Source Currency': 'src_ccy'
                ,'Rates.Reference date': 'date'
                ,'Rates.Rate': 'rate'
                ,'Rates.Description': 'description'}
                    , axis='columns', inplace=True)
##############################################################################################################

# Replace backslashes in dest_directory
dest_directory = dest_directory.replace('\\','/')

for input_file in input_files:

    # Extract the file name and format it 
    if 'LRI_' in input_file:
        file_name = input_file.split("LRI_")[1].split("@")[0]
        file_name = file_name.replace('shares','Shares')
        prefix = file_name.split(' ')[0]
    else:
        file_name = 'Y Shares'   
        prefix = 'Y' 

    # Add the file name to the dest_directory and log_file_directory
    dest_file = dest_directory + f'/{prefix}2. Fund Operations - RC Migrated.xlsx'

    # Start logging
    if os.path.exists(f'{file_name}.log'): 
        os.remove(f'{file_name}.log')

    log.update_handler(logger,file_name)
    logger.info(f'Started logging for {file_name}')
    print(f'File: {file_name}')       

    mig.migrate_investee_data(input_file, dest_file,file_name,fx_rates=fx_rates,env='UAT')
    


    

