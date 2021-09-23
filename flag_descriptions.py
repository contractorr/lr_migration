##############################################################################################################
import log 
logger = log.get_logger('root')
import glob
import os 
import mig_functions as mig


# Destination of key files
os.chdir(r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrate - All Files')
#os.chdir(r'C:\Users\RajContractor\OneDrive - IT-Venture Ltd\Documents\Temp')
#os.chdir(r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Test')
input_files = glob.glob('*.xlsx')            
dest_directory = r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrated'
log.update_handler(logger,'Descriptions')
##############################################################################################################

# Replace backslashes in dest_directory
# dest_directory = dest_directory.replace('\\','/')
# dest_file = dest_directory + f'/DescriptionFlag.xlsx' 
# wb = oxl.Workbook()
# ws = wb.active
all_rows = []

# dst_row_num_fees = 2
# dst_row_num_roc = 2
# dst_row_num_cg = 2

# ws['A1'].font = oxl.styles.Font(size=15)
# ws['B1'].font = oxl.styles.Font(size=15)
# ws['C1'].font = oxl.styles.Font(size=15)
# ws['D1'].font = oxl.styles.Font(size=15)

# ws[f'A1'] = 'Share'
# ws[f'B1'] = 'Investee Fund'
# ws[f'C1'] = 'Date'
# ws[f'D1'] = 'Description'

def compile():
    for input_file in input_files:
        #if input_file == 'LRI_O shares@2021 Q1.xlsx':
        print(f'{input_file}-----------------------------------------------------------------')
        # Extract the file name and format it 
        if 'LRI_' in input_file:
            file_name = input_file.split("LRI_")[1].split("@")[0]
            file_name = file_name.replace('shares','Shares')
            prefix = file_name.split(' ')[0]
        else:
            file_name = 'Y Shares'   
            prefix = 'Y' 

        dst_rows = mig.compile_data(input_file, file_name)
        all_rows.append(dst_rows)
    
    return all_rows

# for row in all_rows:
#     for index, item in row.iterrows():
#         if isinstance(item['description'],int):
#             pass 
#         elif len(item['description']) > 254:
#             print(f"{item['fund']}: row {index}")
#         else:
#             print(len(item['description']))

    

    

