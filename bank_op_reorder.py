##############################################################################################################
import openpyxl as oxl
import pandas as pd

template_file = 'C:/Users/RajContractor/OneDrive - IT-Venture Ltd/Documents/Temp/BankOp_Template.xlsx'
input_file = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrate These/BankOps_Raw.xlsx'
dest_file = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrated/BankOps.xlsx'
##############################################################################################################

data = pd.read_excel(input_file,index_col=None)
wb = oxl.load_workbook(template_file)
dst_active_sheet = wb.active
dst_row_num = 5 # Ignore headers in the template 

for row in data.itertuples():
    dst_active_sheet[f'A{dst_row_num}'] = ''                                    # -------------------------------------- BankOperation
    dst_active_sheet[f'B{dst_row_num}'] = ''                                    # -------------------------------------- PaymentAllocation
    dst_active_sheet[f'C{dst_row_num}'] = row.CLOSEDATE                     # -------------------------------------- Bank Op Date CLOSEDATE1
    dst_active_sheet[f'D{dst_row_num}'] = row.CURRENCY11                    # -------------------------------------- Payment currency CURRENCY11
    dst_active_sheet[f'E{dst_row_num}'] = row.OPTYPE1                       # -------------------------------------- Bank Op Type OPTYPE1
    dst_active_sheet[f'F{dst_row_num}'] = row.LINKEDENTITY                  # -------------------------------------- Managed Fund Name LINKEDENTITY
    dst_active_sheet[f'G{dst_row_num}'] = 'FUND'                            # -------------------------------------- Entity Type XX_ENTITYCLASS
    dst_active_sheet[f'H{dst_row_num}'] = row.FUND5                         # -------------------------------------- Entity Name in Bank Op FUND5
    dst_active_sheet[f'I{dst_row_num}'] = row.COUNTERPARTY_TYPE             # -------------------------------------- Counterparty Type XX_COUNTERPARTYCLASS
    dst_active_sheet[f'J{dst_row_num}'] = row.COUNTERPARTY_FUND             # -------------------------------------- Counterparty Fund.Fund FUND_CF
    dst_active_sheet[f'K{dst_row_num}'] = row.ENTITY_BANKACCOUNT            # -------------------------------------- Bank Account BANKACCOUNTC1
    dst_active_sheet[f'L{dst_row_num}'] = row.COUNTERPARTY_BANKACCOUNT      # -------------------------------------- Bank Account ACCOUNTCODE2
    dst_active_sheet[f'M{dst_row_num}'] = row.AMOUNT                            # -------------------------------------- Amount (Bank) AMOUNT1
    dst_active_sheet[f'N{dst_row_num}'] = row.AMOUNT                            # -------------------------------------- Amount (Counterparty) AMOUNTC1
    dst_active_sheet[f'O{dst_row_num}'] = row.AMOUNT                            # -------------------------------------- Amount (Payment) AMOUNT21 -- the main one
    dst_active_sheet[f'P{dst_row_num}'] = row.AMOUNT                            # -------------------------------------- Amount (Entity) AMOUNT31 
    #(AMOUNTCB1 - VCBANKACCTOP.AMOUNTCB - Counterparty Bank Currency)          
    dst_active_sheet[f'Q{dst_row_num}'] = 'TRUE'                            # -------------------------------------- Draft DRAFT1
    dst_active_sheet[f'R{dst_row_num}'] = 'FALSE'                           # -------------------------------------- Locked XX_LOCKED
    dst_active_sheet[f'S{dst_row_num}'] = row.CLOSEDATE                     # -------------------------------------- CLOSEDATE22
    dst_active_sheet[f'T{dst_row_num}'] = row.INDEXOP                       # -------------------------------------- INDEXOP22
    dst_active_sheet[f'U{dst_row_num}'] = row.fund_op_type                  # -------------------------------------- Fund Op Type OTYPE22

    dst_row_num += 1

wb.save(dest_file)