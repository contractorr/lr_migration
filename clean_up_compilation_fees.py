import pandas as pd
import glob
import os 
import mig_functions as mig
import openpyxl as oxl
import pandas as pd
from fuzzywuzzy import process


# Make a dataframe containing all investee fund ops:
os.chdir(r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrate - All Files')
input_files = glob.glob('[!~]*.xlsx')         
input_files.pop()
input_files.remove('LRI_AH shares@2021 Q1 - duplicate of AHs.xlsx')
fund_import_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/04 Funds.xlsx'
def fix_fund_name(df_mapping):
    df_mapping['fund'] = df_mapping['Investee Fund'].copy()
    for row in df_mapping.itertuples():
        share = row.Share
        share = share.split(' ')[0] 
        fund_import = pd.read_excel(fund_import_file, index_col=None, header=2, usecols='D:Q',sheet_name=share)[2:] # First fund is the management fund 
        fund_import = fund_import['Fund'].to_list()
        fund_name_compare = mig.replace_fund_name(share,row.fund)
        fund_name = process.extractOne(fund_name_compare,fund_import)[0]
        df_mapping.loc[row.Index,'fund'] = fund_name
    
    return df_mapping

all_rows = pd.DataFrame()
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
    all_rows = all_rows.append(dst_rows,ignore_index=True)


# Analyse fees from compilation file 
comp_file_lr = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Data Migration/Investee Fund Operation Compilation Files/Returned by LR/Compilation_LR_Clean.xlsx'
comp_fees = pd.read_excel(comp_file_lr, index_col=None, sheet_name='Fees')
comp_fees['investee_fund'] = comp_fees['Investee Fund']
comp_fees_sum = comp_fees[comp_fees['SplitInd'] != 0].groupby(['Share','old_description','Date']).sum()
comp_fees_sum.reset_index(inplace=True)
for fee in comp_fees_sum.itertuples():
    desc = fee.old_description
    desc = desc.replace(u'\xa0', u' ')
    comp_fees_sum.loc[fee.Index,'old_description'] = desc

# Find all fees that we haven't been able to narrow down:     
for fee in comp_fees_sum.itertuples():
    rel = all_rows[(all_rows['investor']==fee.Share + ' - LR')&(all_rows['description']==fee.old_description)&(all_rows['date']==fee.Date)]
    if len(rel) == 1:
        pass
    else:
        # Print all fees where luca has changed the original description because he's an idiot
        # We'll need to replace the description with what's in the source data for these rows within Compilation_LR (C:\Users\RajContractor\IT-Venture Ltd\Lion River - Documents\Data Migration\Investee Fund Operation Compilation Files\Returned by LR)
        print('------------------------------------------') 
        print(f'Not found: {fee}')
        rel = all_rows[(all_rows['investor']==fee.Share + ' - LR')&(all_rows['date']==fee.Date)]
        print(f"Source desc found: {rel[['fund_name','description','fees_fund_ccy','fees_fund_ccy_inside_commitment']]}")
        print('------------------------------------------') 

comp_fees_sum['inside_commitment'] = None
for fee in comp_fees_sum.itertuples():
    rel = all_rows[(all_rows['investor']==fee.Share + ' - LR')&(all_rows['description']==fee.old_description)&(all_rows['date']==fee.Date)]
    if rel['fees_fund_ccy'].iloc[0] == round(fee.fee_amount,2):
        #print(f"Outside: {fee}")
        comp_fees_sum.loc[fee.Index,'inside_commitment'] = 'No'
    elif rel['fees_fund_ccy_inside_commitment'].iloc[0] == round(fee.fee_amount,2):
        #print(f"Inside: {fee}")
        comp_fees_sum.loc[fee.Index,'inside_commitment'] = 'Yes'
    else:
        print('------------------------------------------') 
        print(f"Unclassified: {fee}")
        print(f"Source Data: {rel[['fund_name','description','fees_fund_ccy','fees_fund_ccy_inside_commitment']]}")
        print('------------------------------------------') 
    

# ------------------------------------------
# Unclassified: Pandas(Index=0, Share='M Shares', old_description='Administrative costs/RoC', Date=Timestamp('2016-01-29 00:00:00'), fee_amount=12244850.0, roc_amount=2244850, cg_amount=0.0, SplitInd=-3, inside_commitment=None)
#             fund_name                   description             fees_fund_ccy         fees_fund_ccy_inside_commitment
# 6287  ITV - Advantage Partners IV-S  Administrative costs/RoC      2244850.0                       10000000.0
# ------------------------------------------
# Unclassified: Pandas(Index=9, Share='M Shares', old_description='Mgmt fees Q1+Q2 2017 + fees', Date=Timestamp('2017-02-08 00:00:00'), fee_amount=164369.69, roc_amount=0, cg_amount=0.0, SplitInd=15, inside_commitment=None)
#             fund_name                   description             fees_fund_ccy         fees_fund_ccy_inside_commitment
# 6599  ITV - Ambienta II           Mgmt fees Q1+Q2 2017 + fees         0.0                        169409.69
# ------------------------------------------

comp_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Data Migration/Investee Fund Operation Compilation Files/Returned by LR/Compilation_LR.xlsx'
comp_fees = pd.read_excel(comp_file, index_col=None, sheet_name='Fees (2)')
comp_fees['investee_fund'] = comp_fees['Investee Fund']
comp_fees_split = comp_fees[comp_fees['SplitInd'] != 0].copy()
for fee in comp_fees_split.itertuples():
    desc = fee.old_description
    if desc == desc:
        desc = desc.replace(u'\xa0', u' ')
    else:
        desc = fee.Description
    comp_fees_split.loc[fee.Index,'old_description'] = desc

for fee in comp_fees_split.itertuples():
    rel = comp_fees_sum[(comp_fees_sum['Share']==fee.Share)&(comp_fees_sum['old_description']==fee.old_description)&(comp_fees_sum['Date']==fee.Date)]
    if len(rel) == 1:
        comp_fees_split.loc[fee.Index,'inside_commitment'] = rel['inside_commitment'].iloc[0]
    else:
        print('ERROR')


comp_fees_split.to_excel('Split Fees 2.xlsx',columns=['Share','Investee Fund','Date','old_description','Description','fee_amount','roc_amount','cg_amount','Mapping','inside_commitment'],index=None)

# Output rows to template file:
import mig_functions as mig
import openpyxl as oxl
import pandas as pd
input_file = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrate - All Files/LRI_M shares@2021 Q1.xlsx'
file_name = 'M Shares' 
dst_rows = mig.compile_data(input_file, file_name)
dst_file = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrated/LRI_AHp shares migrated.xlsx'      
template_file = 'C:/Users/RajContractor/OneDrive - IT-Venture Ltd/Documents/Temp/N2. Fund Operations - Test Template.xlsx'
wb = oxl.load_workbook(template_file)
dst_active_sheet = wb.active
dst_row_num = 5 # Ignore headers in the template 
sheet_count = 1 # Count the sheets
###############################
descriptions_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Data Migration/Investee Fund Operation Compilation Files/Compilation - All Fees, Return of Capital and Capital Gains.xlsx'
desc_xl = pd.ExcelFile(descriptions_file)
for sheet in desc_xl.book.worksheets:
    if sheet.title == 'Fees':
        desc_fees = desc_xl.parse(sheet.title)
    # All capital gains within the return of capital sheet are classified as Realised Gain/Loss so we don't need to read that sheet in.
    elif sheet.title == 'Capital Gains':
        desc_cap_gains = desc_xl.parse(sheet.title)
dst_rows = mig.categorise_fees_cap_gains(dst_rows,desc_fees,desc_cap_gains)
###############################
for i, dst_row in dst_rows.iterrows():
    dst_row_num = mig.insert_row(dst_row                                                        # pass in the row we want to insert                                                
                                ,dst_row_num                                                    # this is just the current row count 
                                ,dst_active_sheet                                               # the active sheet we're editing
                                )

wb.save(dst_file)


dst_rows[(dst_rows['investor']=='M Shares - LR')&(dst_rows['description']=='Administrative costs/RoC')].loc[:,['fees_fund_ccy','fees_fund_ccy_inside_commitment','roc_fund_ccy','redraw_fund_ccy']]