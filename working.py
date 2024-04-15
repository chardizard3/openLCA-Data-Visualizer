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
    df = df.drop(df.index[[0,4]])
    df = df.drop(df.columns[[0,1,3,5,6,7,8,9,10]], axis=1)
     # Calculate sum of column 2
    sum = df.iloc[:, 1].sum()
    
    #df.loc['Sum'] = ['Sum', sum]

    df['Percentage'] = df['Unnamed: 4'] / sum * 100
    #save
    df.to_excel("2test.xlsx", index=False)

    # Plot the pie chart
    plt.figure(figsize=(10, 7))
    patches, texts, _ = plt.pie(df['Percentage'], startangle=140, autopct='%1.1f%%')
    plt.legend(df['Unnamed: 2'], loc="center left", bbox_to_anchor=(0.8,0.5))
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.title('Pie Chart of Percentages by Impact')
    plt.show()

excel_file = "1.xlsx"
sheet_name = "Impacts"
df = navigate_to_sheet(excel_file, sheet_name)
