##############################################################################################################
import pandas as pd
import re

# Destination of key files          
dest_directory = r'C:\Users\RajContractor\Documents\Python Files\Dev\LR Migration\Migrate These'
##############################################################################################################

# Replace backslashes in dest_directory
dest_directory = dest_directory.replace('\\','/')


def split_files(input_file):
    
    xl = pd.ExcelFile(input_file)

    # Loop through sheets
    share = ' ' 
    sheets = xl.book.worksheets
    sheets.reverse()
    files = {}
    for sheet in sheets:
        if sheet.sheet_state == 'hidden':
            # Ignore hidden sheets
            pass 
        else:
            if re.match('a[a-zA-Z] share[\d\D]*',(sheet.title).lower()) or (sheet.title == 'Z Shares'):
                share = sheet.title
                share = share.replace('Share','shares')
                share = re.match('[\D]*shares',share).group()
                files[share] = []
                dest_file = dest_directory + f'/LRI_{share}@2021 Q1.xlsx'
                if 'writer' in globals():
                    writer.close()
                writer = pd.ExcelWriter(dest_file)
            elif share == ' ':
                # No share found yet
                pass 
            else:                
                sheet_name = sheet.title
                df = xl.parse(sheet_name)
                if df.keys()[0] == 'Name' and 'Unnamed' not in df.keys()[1]:
                    # We've found a fund
                    fund_name = df.keys()[1].strip()
                    files[share].append(fund_name)
                    df.to_excel(writer, sheet_name = sheet_name, index=False)
                    writer.save()
    writer.close()
    return files    

def split_files_rhone():
    input_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/LR Reports/Investee Funds/JV Rhone Performance (B-C-D- F-H-I-J-K-L-P-Q-S-T-X-Y) @2021 Q1.xlsx'
    xl = pd.ExcelFile(input_file)

    # Loop through sheets
    share = ' ' 
    sheets = xl.book.worksheets
    files = {}
    for sheet in sheets:
        print(sheet.title)
        m = re.fullmatch('[A-Z] Inv',sheet.title)
        if sheet.sheet_state == 'hidden':
            # Ignore hidden sheets
            pass 
        elif m:
            share = sheet.title
            share = share.replace('Inv','shares')
            files[share] = []
            dest_file = dest_directory + f'/LRI_{share}@2021 Q1.xlsx'
            if 'writer' in globals():
                writer.close()
            writer = pd.ExcelWriter(dest_file)         
            df = xl.parse(sheet.title)
            if df.keys()[0] == 'Name' and 'Unnamed' not in df.keys()[1]:
                # We've found a fund
                fund_name = df.keys()[1].strip()
                files[share].append(fund_name)
                df.to_excel(writer, sheet_name = fund_name, index=False)
                writer.save()
    writer.close()
    return files   
