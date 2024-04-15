import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import filedialog

def navigate_to_sheet():
    excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])

    sheet_name = "Impacts"
    # Read the Excel file
    xls = pd.ExcelFile(excel_file)

    # Read the sheet
    df = pd.read_excel(xls, sheet_name)

    #data cleaning
    df = df.drop(df.index[[0, 4]])
    df = df.drop(df.columns[[0, 1, 3, 5, 6, 7, 8, 9, 10]], axis=1)
    sum_value = df.iloc[:, 1].sum()
    df['Percentage'] = df['Unnamed: 4'] / sum_value * 100

    #plot
    fig, ax = plt.subplots(figsize=(10, 7))
    patches, texts, _ = ax.pie(df['Percentage'], startangle=140, autopct='%1.1f%%')
    ax.legend(df['Unnamed: 2'], loc="upper left", bbox_to_anchor=(0.7,1))
    ax.axis('equal')
    ax.set_title('Pie Chart of Percentages by Impact')

    #gui for plot
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas.get_tk_widget().grid(row=1, column=0, columnspan=3)

#gui
#main
root = tk.Tk()
root.title("OpenLCA Impacts Data Visualization")

#buttons
browse_button = tk.Button(root, text="Browse File", command=navigate_to_sheet)
browse_button.grid(row=0, column=0, padx=5, pady=5)
quit_button = tk.Button(root, text="Quit", command=root.quit)
quit_button.grid(row=0, column=1, padx=5, pady=5)

root.mainloop()
