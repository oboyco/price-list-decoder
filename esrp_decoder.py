import pandas as pd


# Open the txt file
with open('03-12-2018 PMD102 A-type.txt', 'r') as fp:
    lines = fp.read().splitlines()  # read each line in a list

part_number = [x[9:19].strip(' ') for x in lines]
description = [x[25:51] for x in lines]
price = [float(x[54:64].replace('**********', '0'))/100 for x in lines]
weight = [int(x[79:88]) for x in lines]
cmc = [x[110:113] for x in lines]
dg = [x[115:117] for x in lines]
price_multiplied = []

df = pd.DataFrame([part_number, description, weight, cmc, dg, price, price_multiplied]).T
df.columns = ['Part number', 'Description', 'Weight, g', 'CMC', 'Discount group'
    , 'ESRP', 'ESRP (multiplied)']
df = df[~df['Description'].astype(str).str.startswith('CANCEL ')]
df = df.sort_values(by='ESRP', ascending=False)


 #create new dataframe with ArtNo created from part of the Description
df2 = df.copy()[['Description']]
df2.columns = ['Part number']
df2['Part number'] = df2['Part number'].str.split(n=1).str[0]

 #merge price from the first dataframe
df2 = pd.merge(df2, df[['Part number', 'ESRP']], how='left', on='Part number')

 #create a new column 'Price (multiplied)' and fill NANs from original 'Price' column
df['ESRP (multiplied)'] = df2['ESRP'].values
df['ESRP (multiplied)'] = df['ESRP (multiplied)'].fillna(df['ESRP'])
df = df.sort_values(by='ESRP (multiplied)', ascending=False)

df.to_excel('output.xlsx', sheet_name='ESRP', index=False)
