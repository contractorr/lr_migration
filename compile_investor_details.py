from fuzzywuzzy import fuzz
import pandas as pd

comm_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/LR Data Files/Managed Funds/MASTERFILE LRI Performance_@31.03.2021 - YES PEX - 2.xlsx'
comm = pd.read_excel(comm_file,sheet_name='CAS by shareholders',header=None)

# Get the relevant rows of data from the file 
r = [i for i in range(6,41)]
r.insert(0,3)
commitment = comm.loc[r,:].set_index(0)

# Get all the investors 
investors = comm.loc[1,:].dropna()
investors = investors.apply(lambda x: x.split('-')[0])
investors = investors.iloc[1:-1]

# Replace investor names with their official names
# Get the official list of investors
investors_import_file = 'C:/Users/RajContractor/IT-Venture Ltd/Lion River - Documents/Import Files/ITV Import Files/06 Investors.xlsx'
investors_import = pd.read_excel(investors_import_file, index_col=None, header=2, usecols='C:H',sheet_name='Investors Import')[1:] 
# Replace any old names with the new ones
print("1. Commitments Masterfile: Replace all old names with the new name based on the Investors import file")
old_names = investors_import['Description'].dropna()
# Aachenmunchener Vericherung AG and Generali Versicherung AG (Germany) got merged into Generali Deutschland Versicherung AG. The former was treated as a 'name change' and ignored. Maybe the same should be done for Generali Versicherung AG (Germany). 
old_names.loc[14] = old_names.loc[14].split('\n')[1]
old_names = old_names.apply(lambda x: x.split('\"')[1])
old_names = old_names.apply(lambda x: x.strip())
for index, name in old_names.iteritems():
	for row in investors.iteritems():
        	if row[1].strip() == name:
                    print(f"\t{investors.loc[row[0]]} --> {investors_import.loc[index,'Investor']}")
                    investors.loc[row[0]] = investors_import.loc[index,'Investor']

# Replace all the names with the best match from the list of investors we've imported into the system
print("2. Commitments Masterfile: Replace all the names with the official names based on our Investors import file")
official_investors = investors_import['Investor'].to_list()
official_investors_pd = pd.DataFrame()
official_investors_pd['investor'] = investors_import['Investor']
for row in investors.iteritems():
    investor_instance = row[1]
    if investor_instance.strip() == 'Generali Poisťovňa, a. s.':
        investor_instance = 'ITV - Generali Poist’ovňa a.s.'
    elif investor_instance.strip() == 'AachenMunchener Lebensversicherung AG':
        investor_instance = 'Generali Deutschland Lebensversicherung AG'
    ratio = official_investors_pd['investor'].apply(lambda x: fuzz.token_sort_ratio(x,investor_instance))
    correct_investor = official_investors_pd.loc[ratio.idxmax(),'investor']
    investors.loc[row[0]] = correct_investor
    if ratio.max() < 101:
        print(f'\t{row[1]} --> {correct_investor} ({ratio.max()}% match)')

# Store the commitments for each investor in a dictionary 
comm_dict = {}
for investor in investors.iteritems():
	comm_dict[investor[1]] = commitment.loc[:,[investor[0],investor[0]+4]].dropna()
	comm_dict[investor[1]].columns = comm_dict[investor[1]].iloc[0]
	comm_dict[investor[1]] = comm_dict[investor[1]].iloc[1:]

nominals_file = 'C:/Users/RajContractor/Documents/Lion River/LR Reports/Managed Funds/Nominals_New_Shares.xlsx'
nominals = pd.read_excel(nominals_file)
print("3. Nominals file: Replace all old names with the new name based on the Investors import file")
for index, name in old_names.iteritems():
	for row in nominals.itertuples():        
		if row.investor.strip() == name:
			print(f"\t{nominals.loc[row.Index,'investor']} --> {investors_import.loc[index,'Investor']}")
			nominals.loc[row.Index,'investor'] = investors_import.loc[index,'Investor']
			

official_investors = investors_import['Investor'].to_list()
transformations = pd.DataFrame()
print("4. Nominals file: Replace all the names with the official names based on our Investors import file")
for row in nominals.itertuples():
    investor_instance = row.investor
    if investor_instance.strip() == 'Generali Poisťovňa, a. s.':
        investor_instance = 'Generali Poist’ovňa a.s.'
    elif investor_instance.strip() == 'Generali España SA':
        investor_instance = 'Generali España S.A. de Seguros y Reaseguros'
    elif investor_instance.strip() == 'Generali Česká pojišťovna a.s.':
        investor_instance = 'Generali Česká pojišt’ovna a.s.'
    elif investor_instance.strip() == 'AachenMunchener Lebensversicherung AG':
        investor_instance = 'Generali Deutschland Lebensversicherung AG'
    elif investor_instance.strip() == 'AachenMunchener Versicherung AG': # This is a transfer with description of 'Name change' in the nominals file. We will ignore this.
        investor_instance = 'Generali Deutschland Versicherung AG'
    # elif investor_instance.strip() == 'Generali Versicherung AG (Germany)':
    #     investor_instance = 'Generali Deutschland Versicherung AG'
    ratio = official_investors_pd['investor'].apply(lambda x: fuzz.token_sort_ratio(x,investor_instance))
    correct_investor = official_investors_pd.loc[ratio.idxmax(),'investor']
    nominals.loc[row[0],'investor'] = correct_investor
    if ratio.max() < 96:
        transformations = transformations.append({'% match': ratio.max(), 'new_name':correct_investor, 'old_name': row[1]}, ignore_index=True)
        #print(f'\t{row[1]} --> {correct_investor} ({ratio.max()}% match)')

transformations = transformations.drop_duplicates()
transformations = transformations.sort_values(by=['% match','new_name'])
print(transformations) 

print("5. Add the commitment for each investor within each share to our nominals dataframe")
for investor in official_investors:
    if investor in ['ITV - Lion River II N.V.']:
        #print(investor)
        comm = comm_dict[investor].iloc[0,1]
        share = 'A Shares'
        nominals.loc[(nominals['share'] == share)&(nominals['investor']==investor),'effective_commitment'] = comm
    elif investor in comm_dict.keys():
        #print(investor)
        comm_vals = comm_dict[investor]
        for i, comm in comm_vals.iterrows():
            share = comm['Class of Shares'] + ' Shares'
            #print(f"{share}: {comm['Commitment']}")
            nominals.loc[(nominals['share'] == share)&(nominals['investor']==investor),'effective_commitment'] = int(comm['Commitment'])
    else:
        print('\t No commitment found for ',investor.replace('ITV - ',''))

# The original investors are the ones that were present at the start of a fund but were since transferred out. They will need a 'commitment' row. 
original_investors = pd.DataFrame(['Generali levensverzekering maatschappij N.V.'
                     ,'Assicurazioni Generali S.p.A.'
                     ,'Corporate World Opportunities Ltd'
                     ,'Generali Belgium N.V.'
                     ,'Generali Pojišťovna a.s.'
                     ,'Generali Versicherung AG (Germany)'] # Aachenmunchener Vericherung AG and Generali Versicherung AG (Germany) got merged into Generali Deutschland Versicherung AG. The former was treated as a 'name change' and ignored. Maybe the same should be done for Generali Versicherung AG (Germany). 
                     ,columns=['original_investor'])

# Replace any old names with the new ones - there shouldn't be any changes here
for index, name in old_names.iteritems():
	for original_investor in original_investors.itertuples():
        	if original_investor[1].strip() == name:
                    print(f"\tWARNING --------------------- {original_investors.loc[row[0]]} --> {investors_import.loc[index,'Investor']}")
                    original_investors.loc[row[0]] = investors_import.loc[index,'Investor']

# Replace all the names with the best match from the list of investors we've imported into the system
print("6. Original Investors dataframe: Replace all the names with the official names based on our Investors import file")
for original_investor in original_investors.itertuples():
    investor_instance = original_investor[1]
    if investor_instance.strip() == 'Generali Poisťovňa, a. s.':
        investor_instance = 'ITV - Generali Poist’ovňa a.s.'
    elif investor_instance.strip() == 'Generali Versicherung AG (Germany)':
        investor_instance = 'ITV - Generali Versicherung AG (Germany)'
    ratio = official_investors_pd['investor'].apply(lambda x: fuzz.token_sort_ratio(x,investor_instance))
    correct_investor = official_investors_pd.loc[ratio.idxmax(),'investor']
    original_investors.loc[original_investor[0]] = correct_investor
    if ratio.max() < 101:
        print(f'\t{original_investor[1]} --> {correct_investor} ({ratio.max()}% match)')

print("7. Nominals dataframe: Calculate the commitment for all investors")
for nominal in nominals.itertuples():
    if nominal.transfer_ind == 1 and nominal.investor in list(original_investors['original_investor']):
        # Original investors where the commitment should be populated. Add up the current commitment from all new investors (transfer > 0) and attribute this to the original investors
        new_investor_commitments = nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)]['effective_commitment'].sum()
        nominals.loc[nominal.Index,'commitment'] = new_investor_commitments
        # We have neglected shares that contain more than 1 original invesor 
        if nominal.share == 'A Shares':
            # This is a bit different because Lion River II N.V. had shares issued to it after all the other investors, and had shares transferred to it as well 
            LR = nominals[(nominals['transfer']>0)&(nominals['transfer_ind'] ==0)&(nominals['share'] == nominal.share)]
            comm_before_transfer = LR['effective_commitment']*(LR['issue']/(LR['transfer']+LR['issue']))
            nominals.loc[nominal.Index,'commitment'] = new_investor_commitments - comm_before_transfer.values[0]
            nominals.loc[LR.index[0],'commitment'] = comm_before_transfer.values[0]
        elif nominal.share in ['Z Shares','AA Shares','AB Shares','AC Shares','AD Shares','AE Shares','AF Shares']:
            # This is a bit different and could change in the future! Generali Versicherung AG (Germany) is an original investor who transfers to Generali Deutschland Versicherung AG
            if 'Germany' in nominal.investor:
                # Add up the shares issued to Generali Versicherung AG (Germany) and Generali Deutschland Versicherung AG
                tot_issue = nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&(nominals['investor'].str.contains('Deutschland'))]['issue'].values[0] + nominals[(nominals['share']==nominal.share)&(nominals['transfer']<0)&(nominals['investor'].str.contains('Germany'))]['issue'].values[0]                
                # Set the commitment to the effective commitment of Generali Deutschland Versicherung AG
                nominals.loc[nominal.Index,'commitment'] = (nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&(nominals['investor'].str.contains('Deutschland'))]['effective_commitment'].values[0] * nominal.issue/tot_issue)
            else:
                # Add up the shares issued to Generali Versicherung AG (Germany) and Generali Deutschland Versicherung AG
                tot_issue = nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&((nominals['investor'].str.contains('Česká'))|(nominals['investor'].str.contains('Ceska')))]['issue'].values[0] + nominals[(nominals['share']==nominal.share)&(nominals['transfer']<0)&(~nominals['investor'].str.contains('Germany'))]['issue'].values[0]                
                # Set the commitment to the effective commitment of Generali Deutschland Versicherung AG
                # Set the commitment to the effective commitment of Generali Česká pojišt’ovna a.s.
                nominals.loc[nominal.Index,'commitment'] = (nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&(~nominals['investor'].str.contains('Deutschland'))]['effective_commitment'].values[0] * nominal.issue/tot_issue)
    elif nominal.transfer_ind == 1 and nominal.investor not in list(original_investors['original_investor']):
        if nominal.share in ['Z Shares','AA Shares','AB Shares','AC Shares','AD Shares','AE Shares','AF Shares']:
            if 'Deutschland' in nominal.investor:
                tot_issue = nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&(nominals['investor'].str.contains('Deutschland'))]['issue'].values[0] + nominals[(nominals['share']==nominal.share)&(nominals['transfer']<0)&(nominals['investor'].str.contains('Germany'))]['issue'].values[0]
                nominals.loc[nominal.Index,'commitment'] = nominal.effective_commitment * nominal.issue/tot_issue
            else:
                tot_issue = nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&((nominals['investor'].str.contains('Česká'))|(nominals['investor'].str.contains('Ceska')))]['issue'].values[0] + nominals[(nominals['share']==nominal.share)&(nominals['transfer']<0)&(~nominals['investor'].str.contains('Germany'))]['issue'].values[0]
                nominals.loc[nominal.Index,'commitment'] = (nominals[(nominals['share']==nominal.share)&(nominals['transfer']>0)&(~nominals['investor'].str.contains('Deutschland'))]['effective_commitment'].values[0] * nominal.issue/tot_issue)
        else:
            all_rows_for_investor = nominals[(nominals['share']==nominal.share)& (nominals['investor'] == nominal.investor)]
            if (all_rows_for_investor['transfer'].sum() + all_rows_for_investor['issue'].sum()) > 0 and all_rows_for_investor['issue'].sum() > 0:
                # The commitment recorded in the file for any investors who are new and still around should match the effective commitment
                ratio = all_rows_for_investor['issue'].sum()/(all_rows_for_investor['transfer'].sum() + all_rows_for_investor['issue'].sum())
                nominals.loc[nominal.Index,'commitment'] = nominals.loc[nominal.Index,'effective_commitment'] * ratio




# For any rows with transfer_ind = 0, the commitment will be unpopulated (unless it's already populated in the Nominals_New_Shares file). So populate it. 
nominals.loc[nominals['transfer_ind']==0,'commitment'] = nominals.loc[nominals['transfer_ind']==0,'effective_commitment']

nominals = nominals.drop(columns=['Unnamed: 10','BAWAG'])
nominals.loc[nominals['commitment'].isna(),'commitment'] = 0

for nominal in nominals.itertuples():
    if nominal.transfer == nominal.transfer:
        nominals.loc[nominal.Index,'transfer_ind'] = 1
    else:
        nominals.loc[nominal.Index,'transfer_ind'] = 0

nominals = nominals.astype({'commitment':'int64','transfer_ind': 'int64'})
nominals.to_excel('Investor_Details.xlsx',index=None)


# Any original investors (transfer = 1, issue == -transfer) need to have the commitment of the investors they transferred to
# AH shares > ITV - S.C. Generali Romania Asigurare Reasigurare S.A.




"""
    if 'Česká' in investor_instance:
        investor_instance = 'Generali Česká pojišt’ovna a.s.'
    elif 'Generali Poisťovňa, a. s.' in investor_instance:
        investor_instance = 'Generali Poist’ovňa a.s.'
    elif 'Generali Versicherung AG (Germany)' in investor_instance:
        investor_instance = 'Generali Deutschland Versicherung AG'
    elif 'Generali Versicherung AG (Austria)' in investor_instance:
        investor_instance = 'Generali Versicherung AG'
    elif 'Generali CEE Holding N.V.' in investor_instance:
        investor_instance = 'Generali CEE Holding B.V.'

    # ---------------------2---------------------
    # Read in commitments and drop empty and aggregate values 
    commitments_file = 'C:/Users/RajContractor/Documents/Lion River/LR Reports/Managed Funds/Commitment.xlsx'

    commitments = pd.read_excel(commitments_file,index_col=None,header=7) # 3 columns - 'Class of shares/shareholder, % Ownership, Commitment

    # Remove rows that are a) unpopulated, b) 100 
    drop_index = commitments[commitments['% Ownership'].isin([1.000000])|np.isnan(commitments['% Ownership'])].index 
    commitments.drop(drop_index, inplace=True)    

    # split up the title so we can tell what shares each row corresponds to
    # The column contains things like 'M shares  - Generali Italia S.p.A. (Non-Life)' 
    shares = []
    holders = []

    for row in commitments['Class of shares/shareholder']:

        shareholder = row.split(' - ')
        share = shareholder[0] # share class goes in here, e.g. 'M shares' 
        holder = ''

        for i in range(len(shareholder)):
            if i == 0:
                share = shareholder[i]
            elif i == 1:
                holder = shareholder[i] # shareholder goes in here 
            else:
                holder = holder + ' - ' + shareholder[i] # so that 'Non-Life' is restored

        share = share.replace('shares','Shares')
        shares.append(share)
        holders.append(holder)

    commitments['share'] = shares
    commitments['holder'] = holders
    commitments['share'] = commitments['share'].str.strip()

    # This shouldn't be necessary
    drop_index = commitments[commitments['holder'] == ''].index   
    commitments.drop(drop_index , inplace=True)

    # Merge the (life) and (non-life) amounts 
    # Use a new 'com' array to merge these amounts. This will be all lowercase  
    # Use a new 'inv' array to store the investor's name. This will be as above but not converted to lowercase. 
    commitments['combine'] = commitments['holder'].str.lower()
    combine = []
    investors = []

    for i in range(len(commitments['combine'])):
        com = commitments['combine'].iloc[i].replace('(non-life)','').strip() 
        com = com.replace('(life)','').strip()
        inv = commitments['holder'].iloc[i].replace('(non-life)','').strip() 
        inv = inv.replace('(life)','').strip()
        inv = inv.replace('(Life)','').strip()
        inv = inv.replace('(Non-Life)','').strip()
        inv = inv.replace('(Non-life)','').strip()
        combine.append(com)d
        investors.append(inv)

    # Store the above arrays in new 'combine' and 'investors columns 
    commitments['combine'] = combine
    commitments['investor'] = investors

    # ---------------------3---------------------
    # Create a new investor_details dataframe that will store the combined % ownership, commitment, issue and tranfer for each investor/share   
    investor_details = pd.DataFrame()
    investor_details = commitments.groupby(['investor','share']).sum()
    investor_details.reset_index(inplace=True)
    investor_details.sort_values(by=['share','investor'],inplace=True, ignore_index=True)
    
    # ---------------------4---------------------
    # Add Nominals 
    nominal_file = "C:/Users/RajContractor/Documents/Lion River/LR Reports/Managed Funds/Nominals.xlsx"
    xl_nom = pd.ExcelFile(nominal_file)
    nominal_rows = xl_nom.parse()
    #nominal_rows.sort_values(by=['Share','Investor'],inplace=True,ignore_index=True)  

    investor_details = pd.merge(nominal_rows[['share','investor','issue_date','issue','transfer_ind','transfer_date','transfer']]
                               ,investor_details
                               ,on=['investor','share']) 
    #"""