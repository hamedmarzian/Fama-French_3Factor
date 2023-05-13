import pandas as pd
import math

#                              ******before you start, please read the 'Read Me' file******

# Input the name you want to be selected for the final file ('your input' + '_final')
file_name = input("What should be the name of your final file?!\n")
name = file_name + '_final' + '.xlsx'

# Read the original Excel file into a DataFrame
df = pd.read_excel('Path of the original Excel file\File name.xlsx', sheet_name = 'your Excel sheet name')

# Create a new Excel file with the above input + _final
writer = pd.ExcelWriter(name, engine ='openpyxl')

# Write the original DataFrame to the original sheet in the new Excel file
df.to_excel(writer, sheet_name ='raw', index = False)

# Create a new sheet named 'Size' in the new Excel file
# Companies are divided into two categories based on size (B for Big and S for Small)
median_tir = df['Market Value tir'].median()
df_size = pd.DataFrame({'Company': df['Company'], 'Market Value tir': df['Market Value tir']})
df_size['size'] = df_size['Market Value tir'].apply(lambda x: 'B' if x >= median_tir else 'S')

# Write the DataFrame to the 'Size' sheet in the new Excel file
df_size.to_excel(writer, sheet_name ='Size', index = False)

# Create a new sheet named 'BM' in the new Excel file
df_bm = pd.DataFrame({'Company': df['Company'], 'Book Value': df['Book Value'], 'Market Value esfand': df['Market Value esfand']})

# Calculate B/M
df_bm['Market Value Million'] = df_bm['Market Value esfand'] / 1000000
df_bm['B/M'] = df_bm['Book Value'] / df_bm['Market Value Million']
bm_cutoff1 = df_bm['B/M'].quantile(0.3)
bm_cutoff2 = df_bm['B/M'].quantile(0.7)

# Companies are divided into three categories based on B/M (H for High, M for Medium and L for Low)
df_bm.loc[df_bm['B/M'] >= bm_cutoff2, 'BM'] = 'H'
df_bm.loc[(df_bm['B/M'] < bm_cutoff2) & (df_bm['B/M'] >= bm_cutoff1), 'BM'] = 'M'
df_bm.loc[df_bm['B/M'] < bm_cutoff1, 'BM'] = 'L'

# Calculate the number of companies in each group
num_H = int(len(df_bm) * 0.3)
num_M = int(len(df_bm) * 0.4)
num_L = len(df_bm) - num_H - num_M

# If there are any remaining companies, assign them to the 'M' group
if num_L > 0:
    df_bm.loc[df_bm['BM'].isna(), 'BM'] = 'M'

# Write the DataFrame to the 'BM' sheet in the new Excel file
df_bm.to_excel(writer, sheet_name = 'BM', index = False)

# Create a new sheet named 'F&F' in the new Excel file
df_ff = pd.DataFrame({'Company': df['Company'], 'size': df_size['size'], 'BM': df_bm['BM']})

# Add a new column to the DataFrame with the group number for each row
df_ff['Category'] = df_ff['size'].astype(str) + (df_ff['BM']).astype(str)

# Add columns '1M' to '12M' from the 'your Excel sheet name' to this sheet
for i in range(1, 13):
    df_ff[str(i) + 'M'] = df[str(i) + 'M']

# Calculate RM for each month
df_ff.loc['RM'] = {'Company': 'RM'}
for i in range(1, 13):
    df_ff.loc['RM', str(i) + 'M'] = df_ff[str(i) + 'M'].mean()

# Write the DataFrame to the 'F&F' sheet in the new Excel file
df_ff.to_excel(writer, sheet_name = 'F&F', index = False)

# Create a new sheet named 'mean-6' in the new Excel file
df_mean6 = pd.DataFrame(columns = ['Category'] + [str(i) + 'M' for i in range(1, 13)])

# Define a function to calculate the mean values for each category
def calculate_means(df, category):
    means = []
    for i in range(1, 13):
        means.append(df.loc[df['Category'] == category, str(i) + 'M'].mean())
    return means

# Calculate the mean values for each category and add them to the DataFrame
categories = ['BH', 'BM', 'BL', 'SH', 'SM', 'SL']
for category in categories:
    means = calculate_means(df_ff, category)
    df_mean6.loc[len(df_mean6)] = [category] + means

# Add two rows for SMB and HML
df_mean6.loc[len(df_mean6)] = ['SMB'] + [''] * 12
df_mean6.loc[len(df_mean6)] = ['HML'] + [''] * 12

# Calculate SMB and HML for each month
for i in range(1, 13):
    SMB = (df_mean6.loc[df_mean6['Category'].isin(['SL', 'SM', 'SH']), str(i) + 'M'].mean() 
           - df_mean6.loc[df_mean6['Category'].isin(['BL', 'BM', 'BH']), str(i) + 'M'].mean())
    HML = (df_mean6.loc[df_mean6['Category'].isin(['SH', 'BH']), str(i) + 'M'].mean() 
           - df_mean6.loc[df_mean6['Category'].isin(['SL', 'BL']), str(i) + 'M'].mean())
    df_mean6.loc[df_mean6['Category'] == 'SMB', str(i) + 'M'] = SMB
    df_mean6.loc[df_mean6['Category'] == 'HML', str(i) + 'M'] = HML

# Write the DataFrame to the 'mean-6' sheet in the new Excel file
df_mean6.to_excel(writer, sheet_name = 'mean-6', index = False)

# Create a new sheet named 'Size_5' in the new Excel file
df_size_5 = pd.DataFrame({'Company': df['Company'], 'Market Value tir': df['Market Value tir']})

# Add columns '1M' to '12M' from the 'your Excel sheet name' to this sheet
for i in range(1, 13):
    df_size_5[str(i) + 'M'] = df[str(i) + 'M']

# Companies are divided into five categories based on size (Q1 to Q5)
df_size_5.sort_values('Market Value tir', ascending=False, inplace=True)
df_size_5['Size_5'] = pd.qcut(df_size_5['Market Value tir'], 5, labels=['Q1', 'Q2', 'Q3', 'Q4', 'Q5'])

# Write the DataFrame to the 'Size_5' sheet in the new Excel file
df_size_5.to_excel(writer, sheet_name = 'Size_5', index = False)

# Create a new sheet named 'BM_5' in the new Excel file
df_BM_5 = pd.DataFrame({'Company': df['Company'], 'B/M': df_bm['B/M']})

for i in range(1, 13):
    df_BM_5[str(i) + 'M'] = df[str(i) + 'M']

# Companies are divided into five categories based on B/M (Q1 to Q5)
df_BM_5.sort_values('B/M', ascending = False, inplace = True)
df_BM_5['BM_5'] = pd.qcut(df_BM_5['B/M'], 5, labels = ['Q1', 'Q2', 'Q3', 'Q4', 'Q5'])

# Write the DataFrame to the 'BM_5' sheet in the new Excel file
df_BM_5.to_excel(writer, sheet_name = 'BM_5', index = False)

# Create a new sheet named '25_portf' in the new Excel file
df_25_portf = pd.DataFrame({'Company': df['Company'], 'Size_5': df_size_5['Size_5'], 'BM_5': df_BM_5['BM_5']})

# Add a new column to the DataFrame with the group number for each row
df_25_portf['Group'] = df_25_portf['Size_5'].astype(str) + df_25_portf['BM_5'].astype(str)

for i in range(1, 13):
    df_25_portf[str(i) + 'M'] = df[str(i) + 'M']

# Sort the DataFrame by group number and then by company name
df_25_portf.sort_values(['Group', 'Company'], inplace = True)

# Write the DataFrame to the '25_portf' sheet in the new Excel file
df_25_portf.to_excel(writer, sheet_name = '25_portf', index = False)

# Create a new sheet named 'mean-25' in the new Excel file
df_mean25 = pd.DataFrame(columns = ['Group'] + [str(i) + 'M' for i in range(1, 13)])

# Define a function to calculate the mean values for each Group
def calculate_means_Group(df, group):
    meanss = []
    for i in range(1, 13):
        meanss.append(df.loc[df['Group'] == group, str(i) + 'M'].mean())
    return meanss

# Calculate the mean values for each category and add them to the DataFrame
groups = ['Q1Q1', 'Q1Q2', 'Q1Q3', 'Q1Q4', 'Q1Q5', 'Q2Q1', 'Q2Q2', 'Q2Q3', 'Q2Q4', 'Q2Q5', 'Q3Q1', 'Q3Q2', 'Q3Q3', 'Q3Q4', 'Q3Q5', 'Q4Q1', 'Q4Q2', 'Q4Q3', 'Q4Q4', 'Q4Q5', 'Q5Q1', 'Q5Q2', 'Q5Q3', 'Q5Q4', 'Q5Q5']
for group in groups:
    meanss = calculate_means_Group(df_25_portf, group)
    df_mean25.loc[len(df_mean25)] = [group] + meanss

# Write the DataFrame to the 'mean-6' sheet in the new Excel file
df_mean25.to_excel(writer, sheet_name = 'mean-25', index = False)

# save
writer.save()
