import pandas as pd
import openpyxl as oxl

        
input_file = r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrated\Compilation_2.xlsx'
input_file = input_file.replace('\\','/')
dest_directory = r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrated'
dest_directory = dest_directory.replace('\\','/')
dest_file = dest_directory + f'/Categorised_Fund_Ops.xlsx' 
##############################################################################################################

# Replace backslashes in dest_directory
dest_directory = dest_directory.replace('\\','/')
dest_file = dest_directory + f'/Compilation_2.xlsx' 
#input_file = 'C:/Users/RajContractor/Documents/Python Files/Dev/LR Migration/Migrated/Compilation_2.xlsx'
wb = oxl.Workbook()

def split_files(input_file):
    
    xl = pd.ExcelFile(input_file)
    sheets = xl.book.worksheets
    for sheet in sheets:
        # Create sheet and add columns  
        sh = wb.create_sheet(sheet.title)
        sh[f'A1'] = 'Share'
        sh[f'B1'] = 'Investee Fund'
        sh[f'C1'] = 'Date'
        sh[f'D1'] = 'Description'
        sh[f'E1'] = 'Fee Amount'
        sh[f'F1'] = 'Return of Cap Amount'
        sh[f'G1'] = 'Capital Gains Amount'
        sh[f'H1'] = 'Dist: Investment'
        sh[f'I1'] = 'Dist: Realized Gain / (Loss)'
        sh[f'J1'] = 'Dist: Dividends'
        sh[f'K1'] = 'Dist: Interests'
        sh[f'L1'] = 'Dist: Other Income'
        sh[f'M1'] = 'Dist: Withholding Tax'
        sh[f'N1'] = 'Dist: Carry'
        sh[f'O1'] = 'Dist: Subsequent Close Interest'
        sh[f'P1'] = 'Of which Redrawable amount'
        sh[f'Q1'] = 'Index'
        sh[f'R1'] = 'Investor'
        sh[f'S1'] = 'Fund Operation'
        count = 2
        data = xl.parse(sheet=sheet.title)
        for i, row in data.iterrows():
            if 'withholding' in row['Description'].lower() and row['Return of Cap Amount'] == 0 and row['Fee Amount'] > 0:
                data.loc[i,'Dist: Withholding Tax'] = row['Fee Amount']
                data.loc[i,'Fee Amount'] = 0
            elif 'withholding' in row['Description'].lower() and row['Return of Cap Amount'] >0 and row['Fee Amount'] == 0:
                data.loc[i,'Dist: Withholding Tax'] = row['Return of Cap Amount']
                data.loc[i,'Return of Cap Amount'] = 0

        for i, row in data.iterrows():
            sh[f'A{count}'] = row['Share']
            sh[f'B{count}'] = row['Investee Fund']
            sh[f'C{count}'] = row['Date']
            sh[f'D{count}'] = row['Description']
            sh[f'E{count}'] = row['Fee Amount']
            sh[f'F{count}'] = row['Return of Cap Amount']
            sh[f'G{count}'] = row['Capital Gains Amount']
            sh[f'H{count}'] = row['Dist: Investment']
            sh[f'I{count}'] = row['Dist: Realized Gain / (Loss)']
            sh[f'J{count}'] = row['Dist: Dividends']
            sh[f'K{count}'] = row['Dist: Interests']
            sh[f'L{count}'] = row['Dist: Other Income']
            sh[f'M{count}'] = row['Dist: Withholding Tax']
            sh[f'N{count}'] = row['Dist: Carry']
            sh[f'O{count}'] = row['Dist: Subsequent Close Interest']
            sh[f'P{count}'] = row['Dist: Of which Redrawable amount']
            sh[f'Q{count}'] = row['Unnamed: 16']
            sh[f'R{count}'] = row['Unnamed: 17']
            sh[f'S{count}'] = row['Unnamed: 18']
            count = count + 1


wb.save(dest_file)   