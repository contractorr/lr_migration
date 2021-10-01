##############################################################################################################
import glob
import os 
import mig_functions as mig
import pandas as pd
import datetime as dt
import log 
logger = log.get_logger('root')
log.update_handler(logger,'Bank Ops')

dest_directory = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrated/Bank Ops'
#dest_file = dest_directory + f'/Bank Operations.xlsx'  
    
##############################################################################################################


# Investee Fund Ops
print(f'Working On: Investee Fund Ops')
os.chdir('C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/2 UAT Import Files/Investee Fund Ops')
investee_files = glob.glob('*.xlsx')  
concat_fund_ops = pd.DataFrame()
dest_file = dest_directory + f'/Investee Fund Ops.xlsx'  

for input_file in investee_files:
    # Add the file name to the dest_directory and log_file_directory
    name = input_file.split('Fund Operations')[0] 
    print(f'\tFile: {input_file}')
    fund_ops = pd.read_excel(input_file,index_col=None,skiprows=[0,2,3])  
    concat_fund_ops = pd.concat([fund_ops,concat_fund_ops])
    

fund_ops_post_starting_balance = concat_fund_ops[pd.to_datetime(concat_fund_ops['SETTLEMENTDATE1']) >= dt.datetime(2021,1,1)]
mig.migrate_investee_fund_op_bank_ops(fund_ops_post_starting_balance, dest_file)
    
# Managed Fund Ops
print(f'Working On: Managed Fund Ops')   
managed_fund_op_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/2 UAT Import Files/Managed Fund Ops/Managed Fund Operations - RC Migrated.xlsx'
dest_file = dest_directory + f'/Managed Fund Ops.xlsx'

mig.migrate_managed_fund_op_bank_ops(managed_fund_op_file, dest_file)

# Cash Transfers    
print(f'Working On: Cash Transfers')   
dest_file = dest_directory + f'/Cash Transfers.xlsx'
mig.migrate_cash_transfer_bank_ops(dest_file)

# Fees
print(f'Working On: Managed Fund Fees') 
dest_file = dest_directory + f'/Managed Fund Fees.xlsx'
#mig.migrate_managed_fee_bank_ops(dest_file)

# Fees and Incomes
print(f'Working On: Fees and Incomes') 
dest_file = dest_directory + f'/Fees and Incomes.xlsx'
mig.migrate_fee_and_income_bank_ops(dest_file)


