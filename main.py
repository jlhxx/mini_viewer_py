#basic tkinter-based GUI application

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from scipy import stats

#read excel files into pandas df
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        global df
        df = pd.read_excel(file_path)
        sort_column_menu['values'] = df.columns.tolist()  # Populate the sort dropdown
        filter_column_menu['values'] = df.columns.tolist()  # Populate the filter dropdown
        stat_test_column1_menu['values'] = df.columns.tolist()  # Populate t-test dropdown
        stat_test_column2_menu['values'] = df.columns.tolist()  # Populate t-test dropdown
        display_data(df)

#show df in tkinter treeview
def display_data(df):
    for widget in frame.winfo_children():
        widget.destroy()  # Clear previous data

    table = ttk.Treeview(frame)
    table['columns'] = list(df.columns)
    table['show'] = 'headings'

    for col in table['columns']:
        table.heading(col, text=col)

    for row in df.to_numpy():
        table.insert("", "end", values=row)

    table.pack(fill=tk.BOTH, expand=True)

#sort data + filter
def sort_data():
    selected_column = sort_column_var.get()
    ascending_order = sort_order_var.get()
    if selected_column:
        sorted_df = df.sort_values(by=selected_column, ascending=(ascending_order == 'Ascending'))
        display_data(sorted_df)

def filter_data():
    selected_column = filter_column_var.get()
    filter_value = filter_value_var.get()
    if selected_column and filter_value:
        filtered_df = df[df[selected_column].astype(str) == filter_value]
        display_data(filtered_df)

#perform basic ttest - display pop up with p value
def perform_t_test():
    column1 = stat_test_column1_var.get()
    column2 = stat_test_column2_var.get()
    
    if column1 and column2:
        data1 = df[column1].dropna()
        data2 = df[column2].dropna()
        
        if len(data1) > 1 and len(data2) > 1:
            t_stat, p_value = stats.ttest_ind(data1, data2)
            messagebox.showinfo("T-Test Result", f"T-statistic: {t_stat}\nP-value: {p_value}")
        else:
            messagebox.showwarning("Data Error", "Not enough data to perform t-test.")
    else:
        messagebox.showwarning("Input Error", "Please select both columns to perform the t-test.")

def save_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df.to_excel(file_path, index=False)

#main application window
root = tk.Tk()
root.title("Excel Data Manipulation")

#upload/save buttons
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file)
upload_button.pack()

save_button = tk.Button(root, text="Save File", command=save_file)
save_button.pack()

#UI sorting
sort_column_var = tk.StringVar()
sort_order_var = tk.StringVar(value='Ascending')

sort_column_label = tk.Label(root, text="Sort by:")
sort_column_label.pack()
sort_column_menu = ttk.Combobox(root, textvariable=sort_column_var)
sort_column_menu.pack()

sort_ascending_radio = tk.Radiobutton(root, text="Ascending", variable=sort_order_var, value='Ascending')
sort_descending_radio = tk.Radiobutton(root, text="Descending", variable=sort_order_var, value='Descending')
sort_ascending_radio.pack()
sort_descending_radio.pack()

sort_button = tk.Button(root, text="Sort Data", command=sort_data)
sort_button.pack()

#UI filter
filter_column_var = tk.StringVar()
filter_value_var = tk.StringVar()

filter_column_label = tk.Label(root, text="Filter by:")
filter_column_label.pack()
filter_column_menu = ttk.Combobox(root, textvariable=filter_column_var)
filter_column_menu.pack()

filter_value_label = tk.Label(root, text="Filter value:")
filter_value_label.pack()
filter_value_entry = tk.Entry(root, textvariable=filter_value_var)
filter_value_entry.pack()

filter_button = tk.Button(root, text="Filter Data", command=filter_data)
filter_button.pack()

#stats UI
stat_test_column1_var = tk.StringVar()
stat_test_column2_var = tk.StringVar()

stat_test_label = tk.Label(root, text="Perform T-Test:")
stat_test_label.pack()

stat_test_column1_label = tk.Label(root, text="Select Column 1:")
stat_test_column1_label.pack()
stat_test_column1_menu = ttk.Combobox(root, textvariable=stat_test_column1_var)
stat_test_column1_menu.pack()

stat_test_column2_label = tk.Label(root, text="Select Column 2:")
stat_test_column2_label.pack()
stat_test_column2_menu = ttk.Combobox(root, textvariable=stat_test_column2_var)
stat_test_column2_menu.pack()

t_test_button = tk.Button(root, text="Perform T-Test", command=perform_t_test)
t_test_button.pack()

#display data
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

#run main loop - tkinter event loop
root.mainloop()
