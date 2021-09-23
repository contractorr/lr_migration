import pandas as pd
import mig_functions as mig
from fuzzywuzzy import process

# Read in Compilation data
comp_file_lr = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Data Migration/Investee Fund Operation Compilation Files/Returned by LR/Compilation_LR.xlsx'
comp_fees = pd.read_excel(comp_file_lr, index_col=None, sheet_name='Fees (2)')
comp_fees['investee_fund'] = comp_fees['Investee Fund']
# All rows with non-zero capital gains from return of capital/fees sheet should have been copied to the Capital Gains sheet, so we don't need to read in the return of capital sheet
comp_cap_gains = pd.read_excel(comp_file_lr, index_col=None, sheet_name='Capital Gains (2)')
comp_cap_gains['investee_fund'] = comp_cap_gains['Investee Fund']

# Read in official fund names
fund_import_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/UAT Import Files/05 Funds.xlsx'

# Replace the shares in our compilation data
prev_share = ''
name_changes = pd.DataFrame()
for fee in comp_fees.itertuples():
    share = fee.Share
    if share in ['AGp Shares','AHp Shares','AHs Shares','AIp Shares','AIs Shares']:
        share = share.replace('s ',' ')
        share = share.replace('p ',' ')
        comp_fees.loc[fee.Index,'Share'] = share
    if share != prev_share:        
        print(share)
        fund_import = pd.read_excel(fund_import_file, index_col=None, header=2, usecols='D:Q',sheet_name=share.split(' ' )[0])[2:] # First fund is the management fund 
        fund_import = fund_import['Fund'].to_list()
        prev_share = share
    orig_fund = fee.investee_fund
    replaced_fund = mig.replace_fund_name(share.split(' ')[0],orig_fund)
    fund = process.extractOne(replaced_fund,fund_import)[0]
    fund = fund.strip()
    comp_fees.loc[fee.Index,'Investee Fund'] = fund
    if fund != orig_fund:
        change = {'share': share, 'old': orig_fund, 'replaced': replaced_fund, 'new':fund}
        name_changes = name_changes.append(change, ignore_index=True)

prev_share = ''
for cg in comp_cap_gains.itertuples():
    share = cg.Share
    if share in ['AGp Shares','AHp Shares','AHs Shares','AIp Shares','AIs Shares']:
        share = share.replace('s ',' ')
        share = share.replace('p ',' ')
        comp_cap_gains.loc[cg.Index,'Share'] = share
    if share != prev_share:      
        print(share)
        fund_import = pd.read_excel(fund_import_file, index_col=None, header=2, usecols='D:Q',sheet_name=share.split(' ' )[0])[2:] # First fund is the management fund 
        fund_import = fund_import['Fund'].to_list()
        prev_share = share
    orig_fund = cg.investee_fund
    replaced_fund = mig.replace_fund_name(share.split(' ')[0],orig_fund)
    fund = process.extractOne(replaced_fund,fund_import)[0]
    fund = fund.strip()
    comp_cap_gains.loc[cg.Index,'Investee Fund'] = fund
    if fund != orig_fund:
        change = {'share': share, 'old': orig_fund, 'replaced': replaced_fund, 'new':fund}
        name_changes = name_changes.append(change, ignore_index=True)

name_changes = name_changes.drop_duplicates()
name_changes[['share','old','replaced','new']]
# Drop duplicates
comp_fees.drop_duplicates(inplace=True)
comp_cap_gains.drop_duplicates(inplace=True)

# Add investee_fund column
comp_fees['investee_fund'] = comp_fees['Investee Fund']
comp_cap_gains['investee_fund'] = comp_cap_gains['Investee Fund']

for fee in comp_fees.itertuples():
    if fee.old_description != fee.old_description:
        comp_fees.loc[fee.Index,'old_description'] = fee.Description

# Save our file
with pd.ExcelWriter('C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Data Migration/Investee Fund Operation Compilation Files/Returned by LR/Compilation_LR_Clean.xlsx') as writer:
    comp_fees.to_excel(writer,sheet_name='Fees',index=False,columns=['Share','Investee Fund','Date','old_description','Description','fee_amount','roc_amount','cg_amount','Mapping','investee_fund','SplitInd']) 
    comp_cap_gains.to_excel(writer,sheet_name='Capital Gains',index=False,columns=['Share','Investee Fund','Date','Description','fee_amount','roc_amount','cg_amount','Mapping','Recallable/Non Recallable?','investee_fund','Fund Op','SplitInd'])

