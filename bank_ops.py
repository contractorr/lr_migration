##############################################################################################################
import glob
import os 
import mig_functions as mig
import pandas as pd
import datetime as dt


dest_directory = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrated'

    
##############################################################################################################
# Destination of investee fund op files
os.chdir('C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/Investee Fund Ops')
investee_files = glob.glob('*.xlsx')  

# Investee Fund Ops
print(f'Working On: Investee Fund Ops')
concat_fund_ops = pd.DataFrame()
dest_file_fo = dest_directory + f'/Bank Operations - Investee Fund Ops - RC Migrated.xlsx'  
for input_file in investee_files:
    # Add the file name to the dest_directory and log_file_directory
    name = input_file.split('Fund Operations')[0] 
    print(f'\tFile: {input_file}')
    fund_ops = pd.read_excel(input_file,index_col=None,skiprows=[0,2,3])  
    concat_fund_ops = pd.concat([fund_ops,concat_fund_ops])
    

fund_ops_post_starting_balance = concat_fund_ops[pd.to_datetime(concat_fund_ops['SETTLEMENTDATE1']) >= dt.datetime(2021,1,1)]
mig.migrate_investee_fund_op_bank_ops(fund_ops_post_starting_balance, dest_file_fo)
    
# Managed Fund Ops
print(f'Working On: Managed Fund Ops')   
managed_fund_op_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/Managed Fund Ops/Managed Fund Operations - RC Migrated.xlsx'
dest_file_fo = dest_directory + f'/Bank Operations - Managed Fund Ops - RC Migrated.xlsx'
mig.migrate_managed_fund_op_bank_ops(managed_fund_op_file, dest_file_fo)

# Cash Transfers    
print(f'Working On: Cash Transfers')   
dest_file_ct = dest_directory + f'/Bank Operations - Cash Transfers - RC Migrated.xlsx'
mig.migrate_cash_transfer_bank_ops(dest_file_ct)

# Fees
print(f'Working On: Fees') 
dest_file_fees = dest_directory + f'/Bank Operations - Managed Fund Fees - RC Migrated.xlsx'
mig.migrate_managed_fee_bank_ops(dest_file_fees)
