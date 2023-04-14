import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext
import mysql.connector
import io

data = None

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        load_data(file_path)

def load_data(file_path):
    global df
    df = pd.read_excel(file_path)
    file_name_label.config(text=file_path.split('/')[-1])
    update_checkboxes()

def update_checkboxes():
    for col in df.columns:
        if col not in column_vars:
            var = tk.BooleanVar(value=True)
            column_vars[col] = var
            tk.Checkbutton(column_frame, text=col, variable=var, onvalue=True, offvalue=False).pack(anchor='w')
    update_sort_and_aggregation_column_options()

def update_sort_and_aggregation_column_options():
    sort_column_dropdown['values'] = list(df.columns)
    aggregation_column_dropdown['values'] = list(df.columns)

def transform_data():
    selected_columns = [col for col in column_vars if column_vars[col].get()]
    if not selected_columns:
        messagebox.showerror("Error", "No columns selected")
        return

    transformed_df = df[selected_columns].copy()
    agg_function = aggregation_var.get()
    agg_column = aggregation_column_var.get()
    sort_column = sort_column_var.get()
    sort_order = sort_order_var.get()
    group_data = group_data_var.get()

    if group_data and agg_function != 'None' and agg_column:
        groupby_columns = [col for col in selected_columns if col != agg_column]
        if agg_function == 'Sum':
            transformed_df = transformed_df.groupby(groupby_columns).agg({agg_column: 'sum'}).reset_index()
        elif agg_function == 'Max':
            transformed_df = transformed_df.groupby(groupby_columns).agg({agg_column: 'max'}).reset_index()
        elif agg_function == 'Min':
            transformed_df = transformed_df.groupby(groupby_columns).agg({agg_column: 'min'}).reset_index()

    if sort_column:
        transformed_df = transformed_df.sort_values(by=sort_column, ascending=(sort_order == 'Ascending'))

    save_data(transformed_df)

def save_data(data):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        data.to_excel(file_path, index=False)
        messagebox.showinfo("Success", "File saved successfully")

def connect_to_mysql():
    try:
        conn = mysql.connector.connect(
            host=mysql_host_var.get(),
            user=mysql_user_var.get(),
            password=mysql_password_var.get(),
            database=mysql_database_var.get()
        )
        return conn
    except mysql.connector.Error as err:
        messagebox.showerror("Error", str(err))
        return None

def load_tables():
    conn = connect_to_mysql()
    if conn:
        cursor = conn.cursor()
        cursor.execute("SHOW TABLES")
        tables = cursor.fetchall()
        cursor.close()
        conn.close()
        update_table_checkboxes(tables)
    for idx, table in enumerate(tables):
        var = tk.BooleanVar(value=False)
        table_vars[table] = var
        checkbox = tk.Checkbutton(table_frame, text=table[0], variable=var, onvalue=True, offvalue=False, command=lambda table=table: table_selected(table))
        checkbox.grid(row=idx, column=0, sticky='w')

def table_selected(table):
    if table_vars[table].get():
        mysql_query_entry.delete('1.0', tk.END)
        mysql_query_entry.insert(tk.END, f"SELECT * FROM {table[0]}")

def update_table_checkboxes(tables):
    for table in tables:
        if table not in table_vars:
            var = tk.BooleanVar(value=False)
            table_vars[table] = var
            tk.Checkbutton(table_frame, text=table[0], variable=var, onvalue=True, offvalue=False).pack(anchor='w')

def execute_query():
    global data
    conn = connect_to_mysql()
    if conn:
        query = mysql_query_entry.get("1.0", tk.END).strip()
        if query:
            try:
                data = pd.read_sql_query(query, conn)
                print(data)
                display_data(data.head(10))
            except Exception as e:
                messagebox.showerror("Error", str(e))
        conn.close()

def display_data(data):
    # Clear the previous data in the treeview
    for row in data_tree.get_children():
        data_tree.delete(row)

    # Clear the previous columns in the treeview
    data_tree["columns"] = ()

    # Set the new columns and column names
    columns = data.columns
    data_tree["columns"] = tuple(columns)
    for col in columns:
        data_tree.heading(col, text=col)
        data_tree.column(col, anchor='center', stretch=True)

    # Insert new data into the treeview
    for index, row in data.iterrows():
        data_tree.insert("", "end", values=tuple(row))

def export_to_excel():
    global data
    if data is not None:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All Files", "*.*")])
        if file_path:
            try:
                data.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Data exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Warning", "No data to export")

root = tk.Tk()
root.title("Data Transformer")

column_vars = {}

notebook = ttk.Notebook(root)
excel_tab = ttk.Frame(notebook)
mysql_tab = ttk.Frame(notebook)
notebook.add(excel_tab, text="Excel")
notebook.add(mysql_tab, text="MySQL")
notebook.pack(expand=True, fill='both')

file_name_label = tk.Label(excel_tab, text="No file selected")
file_name_label.pack()

column_frame = tk.LabelFrame(excel_tab, text="Columns", padx=10, pady=10)
column_frame.pack(padx=10, pady=10, fill='both', expand=True)

aggregation_var = tk.StringVar()
aggregation_label = tk.Label(excel_tab, text="Aggregation:")
aggregation_label.pack()
aggregation_dropdown = ttk.Combobox(excel_tab, textvariable=aggregation_var, values=["None", "Sum", "Max", "Min"], state="readonly")
aggregation_dropdown.current(0)
aggregation_dropdown.pack()

aggregation_column_var = tk.StringVar()
aggregation_column_label = tk.Label(excel_tab, text="Column to aggregate:")
aggregation_column_label.pack()
aggregation_column_dropdown = ttk.Combobox(excel_tab, textvariable=aggregation_column_var, state="readonly")
aggregation_column_dropdown.pack()

sort_column_var = tk.StringVar()
sort_column_label = tk.Label(excel_tab, text="Sort by column:")
sort_column_label.pack()
sort_column_dropdown = ttk.Combobox(excel_tab, textvariable=sort_column_var, state="readonly")
sort_column_dropdown.pack()

sort_order_var = tk.StringVar()
sort_order_label = tk.Label(excel_tab, text="Sort order:")
sort_order_label.pack()
sort_order_dropdown = ttk.Combobox(excel_tab, textvariable=sort_order_var, values=["Ascending", "Descending"], state="readonly")
sort_order_dropdown.current(0)
sort_order_dropdown.pack()

group_data_var = tk.BooleanVar(value=True)
group_data_checkbox = tk.Checkbutton(excel_tab, text="Group data", variable=group_data_var, onvalue=True, offvalue=False)
group_data_checkbox.pack(anchor='w')

browse_button = tk.Button(excel_tab, text="Browse", command=browse_file)
browse_button.pack(pady=(10, 5))

transform_button = tk.Button(excel_tab, text="Transform", command=transform_data)
transform_button.pack(pady=5)

# MySQL Tab
table_vars = {}

mysql_host_var = tk.StringVar()
mysql_host_label = tk.Label(mysql_tab, text="Host:")
mysql_host_label.grid(row=0, column=0, padx=(10, 5), pady=(10, 5), sticky='e')
mysql_host_entry = tk.Entry(mysql_tab, textvariable=mysql_host_var)
mysql_host_entry.grid(row=0, column=1, padx=(5, 10), pady=(10, 5), sticky='w')

mysql_user_var = tk.StringVar()
mysql_user_label = tk.Label(mysql_tab, text="User:")
mysql_user_label.grid(row=1, column=0, padx=(10, 5), pady=5, sticky='e')
mysql_user_entry = tk.Entry(mysql_tab, textvariable=mysql_user_var)
mysql_user_entry.grid(row=1, column=1, padx=(5, 10), pady=5, sticky='w')

mysql_password_var = tk.StringVar()
mysql_password_label = tk.Label(mysql_tab, text="Password:")
mysql_password_label.grid(row=2, column=0, padx=(10, 5), pady=5, sticky='e')
mysql_password_entry = tk.Entry(mysql_tab, textvariable=mysql_password_var, show="*")
mysql_password_entry.grid(row=2, column=1, padx=(5, 10), pady=5, sticky='w')

mysql_database_var = tk.StringVar()
mysql_database_label = tk.Label(mysql_tab, text="Database:")
mysql_database_label.grid(row=3, column=0, padx=(10, 5), pady=5, sticky='e')
mysql_database_entry = tk.Entry(mysql_tab, textvariable=mysql_database_var)
mysql_database_entry.grid(row=3, column=1, padx=(5, 10), pady=5, sticky='w')

mysql_query_var = tk.StringVar()
mysql_query_label = tk.Label(mysql_tab, text="Query:")
mysql_query_label.grid(row=4, column=0, padx=(10, 5), pady=(10, 5), sticky='e')
mysql_query_entry = scrolledtext.ScrolledText(mysql_tab, height=4, width=40)
mysql_query_entry.grid(row=4, column=1, padx=(5, 10), pady=(10, 5), sticky='w')

load_button = tk.Button(mysql_tab, text="Load", command=execute_query)
load_button.grid(row=6, column=1, pady=(10, 5), sticky='e')

table_frame = tk.LabelFrame(mysql_tab, text="Tables", padx=10, pady=10)
table_frame.grid(row=0, column=2, rowspan=5, padx=10, pady=10, sticky='nsew')

mysql_browse_button = tk.Button(mysql_tab, text="Browse", command=load_tables)
mysql_browse_button.grid(row=5, column=1, pady=(10, 5), sticky='e')

data_tree = ttk.Treeview(mysql_tab, show="headings")
data_tree.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

download_button = tk.Button(mysql_tab, text="Download", command=export_to_excel)
download_button.grid(row=8, column=1, pady=(10, 5), sticky='e')

root.mainloop()