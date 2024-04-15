import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def navigate_to_sheet(excel_file, sheet_name):
    # Read the Excel file
    xls = pd.ExcelFile(excel_file)
        
    # Read the sheet
    df = pd.read_excel(xls, sheet_name)

    #drop
    #df.columns = df.columns.str.strip()
    #df.rows = df.rows.str.strip()
    #print (df.columns)
    df = df.drop(df.index[[0,2,3]])
    df = df.drop(df.columns[[0,1,3,4]], axis=1)
    #df.replace('', pd.NA, inplace=True)
    #df.fillna(0, inplace=True)
    df_subset = df.iloc[:, 1:]  # Exclude first two columns
    df_subset = df_subset.iloc[1:, :]  # Exclude first two rows
    df_subset = df_subset.apply(pd.to_numeric, errors='coerce')
    df['Row Sum'] = df_subset.sum(axis=1, skipna=True)
    df.insert(1, 'Row Sum', df.pop('Row Sum'))
    
    #save
    df.to_excel("1test.xlsx", index=False)

    unique_values = df.iloc[1:, 0].unique()  # Get unique values from the first column
    for value in unique_values:
        # Filter the DataFrame for the current value
        subset_df = df[df.iloc[:, 0] == value]
        
        # Concatenate the second row (header) with the subset DataFrame
        subset_df = pd.concat([df.iloc[[0]], subset_df])
        subset_df = subset_df.dropna(axis=1)
        #subset_df = subset_df.groupby(subset_df.columns, axis=1).sum()
        sum = subset_df.iloc[:,1]
        print(sum)

        #subset_df_normalized = subset_df.iloc[:, 1:-1].div(subset_df.iloc[:, 1], axis=0) * 100
        #print(subset_df_normalized)
        # Print the subset DataFrame for inspection
        """print(f"Rows for value '{value}':")
        print(subset_df)"""

excel_file = "1.xlsx"
#direct inventory contributions sheet 
sheet_name = "Direct impact contributions"
df = navigate_to_sheet(excel_file, sheet_name)
#data_frame = navigate_to_sheet(excel_file, sheet_name)
