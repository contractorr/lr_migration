##############################################################################################################
import glob
import os 
import pandas as pd

# Destination of key files
os.chdir(r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrated')
input_files = glob.glob('*.xlsx')            
##############################################################################################################

total = 0
for input_file in input_files:
    # Add the file name to the dest_directory and log_file_directory
    if input_file == 'C:\\Users\\$AHp2. Fund Operations - RC Migrated.xlsx':
        pass
    else:    
        df = pd.read_excel(input_file,skiprows=3)
        count = max(df.count())
        total += count
        print(f'{input_file}: {count} {total}')

print(f'Total: {total}')