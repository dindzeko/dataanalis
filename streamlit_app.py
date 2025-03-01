import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.dataframes = {}
        self.current_df = None

        self.setup_ui()

    def setup_ui(self):
        # Upload Frame
        self.upload_frame = ttk.LabelFrame(self.root, text="Upload Excel Files")
        self.upload_frame.pack(padx=10, pady=10, fill="x")

        ttk.Button(self.upload_frame, text="Upload Files", command=self.upload_files).pack(pady=5)
        self.tables_listbox = tk.Listbox(self.upload_frame)
        self.tables_listbox.pack(padx=5, pady=5, fill="x")

        # Join Frame
        self.join_frame = ttk.LabelFrame(self.root, text="Join Tables")
        self.join_frame.pack(padx=10, pady=10, fill="x")

        self.table1_var = tk.StringVar()
        self.table2_var = tk.StringVar()
        self.join_type_var = tk.StringVar(value="INNER")
        self.table1_col_var = tk.StringVar()
        self.table2_col_var = tk.StringVar()

        ttk.Label(self.join_frame, text="Table 1:").grid(row=0, column=0)
        ttk.Combobox(self.join_frame, textvariable=self.table1_var).grid(row=0, column=1)
        ttk.Label(self.join_frame, text="Column:").grid(row=0, column=2)
        ttk.Combobox(self.join_frame, textvariable=self.table1_col_var).grid(row=0, column=3)

        ttk.Label(self.join_frame, text="Table 2:").grid(row=1, column=0)
        ttk.Combobox(self.join_frame, textvariable=self.table2_var).grid(row=1, column=1)
        ttk.Label(self.join_frame, text="Column:").grid(row=1, column=2)
        ttk.Combobox(self.join_frame, textvariable=self.table2_col_var).grid(row=1, column=3)

        ttk.Label(self.join_frame, text="Join Type:").grid(row=2, column=0)
        ttk.Combobox(self.join_frame, textvariable=self.join_type_var, 
                    values=["INNER", "LEFT", "RIGHT"]).grid(row=2, column=1)
        ttk.Button(self.join_frame, text="Perform Join", command=self.perform_join).grid(row=2, column=2)

        # Filter Frame
        self.filter_frame = ttk.LabelFrame(self.root, text="Filter Data")
        self.filter_frame.pack(padx=10, pady=10, fill="x")

        self.filter_col_var = tk.StringVar()
        self.filter_op_var = tk.StringVar(value="=")
        self.filter_value_var = tk.StringVar()

        ttk.Label(self.filter_frame, text="Column:").grid(row=0, column=0)
        ttk.Combobox(self.filter_frame, textvariable=self.filter_col_var).grid(row=0, column=1)
        ttk.Label(self.filter_frame, text="Operator:").grid(row=0, column=2)
        ttk.Combobox(self.filter_frame, textvariable=self.filter_op_var, 
                    values=["=", ">", "<", ">=", "<=", "<>", "BETWEEN", "LIKE", "IN"]).grid(row=0, column=3)
        ttk.Label(self.filter_frame, text="Value:").grid(row=0, column=4)
        ttk.Entry(self.filter_frame, textvariable=self.filter_value_var).grid(row=0, column=5)
        ttk.Button(self.filter_frame, text="Apply Filter", command=self.apply_filter).grid(row=0, column=6)

        # Sort Frame
        self.sort_frame = ttk.LabelFrame(self.root, text="Sort Data")
        self.sort_frame.pack(padx=10, pady=10, fill="x")

        self.sort_col_var = tk.StringVar()
        self.sort_order_var = tk.StringVar(value="ASC")

        ttk.Label(self.sort_frame, text="Column:").grid(row=0, column=0)
        ttk.Combobox(self.sort_frame, textvariable=self.sort_col_var).grid(row=0, column=1)
        ttk.Label(self.sort_frame, text="Order:").grid(row=0, column=2)
        ttk.Combobox(self.sort_frame, textvariable=self.sort_order_var, 
                    values=["ASC", "DESC"]).grid(row=0, column=3)
        ttk.Button(self.sort_frame, text="Apply Sort", command=self.apply_sort).grid(row=0, column=4)

        # Display Frame
        self.display_frame = ttk.Frame(self.root)
        self.display_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.tree = ttk.Treeview(self.display_frame)
        self.tree.pack(side="left", fill="both", expand=True)
        ttk.Scrollbar(self.display_frame, orient="vertical", command=self.tree.yview).pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        ttk.Button(self.root, text="Export to Excel", command=self.export_data).pack(pady=10)

        # Trace variables
        self.table1_var.trace('w', self.update_table1_columns)
        self.table2_var.trace('w', self.update_table2_columns)

    def upload_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        for file in files:
            try:
                excel = pd.ExcelFile(file)
                for sheet in excel.sheet_names:
                    df = excel.parse(sheet)
                    key = f"{file.split('/')[-1]}_{sheet}"
                    self.dataframes[key] = df
                self.update_tables_listbox()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def update_tables_listbox(self):
        self.tables_listbox.delete(0, tk.END)
        tables = list(self.dataframes.keys())
        for table in tables:
            self.tables_listbox.insert(tk.END, table)
        self.update_comboboxes(tables)

    def update_comboboxes(self, tables):
        for widget in [self.table1_var, self.table2_var, self.filter_col_var, self.sort_col_var]:
            widget.set('')
        self.table1_combo['values'] = tables
        self.table2_combo['values'] = tables

    def update_table1_columns(self, *args):
        table = self.table1_var.get()
        if table in self.dataframes:
            cols = list(self.dataframes[table].columns)
            self.table1_col_combo['values'] = cols
            self.table1_col_var.set(cols[0] if cols else '')

    def update_table2_columns(self, *args):
        table = self.table2_var.get()
        if table in self.dataframes:
            cols = list(self.dataframes[table].columns)
            self.table2_col_combo['values'] = cols
            self.table2_col_var.set(cols[0] if cols else '')

    def perform_join(self):
        try:
            df1 = self.dataframes[self.table1_var.get()]
            df2 = self.dataframes[self.table2_var.get()]
            merged = pd.merge(df1, df2, 
                            left_on=self.table1_col_var.get(), 
                            right_on=self.table2_col_var.get(),
                            how=self.join_type_var.get().lower())
            self.current_df = merged
            self.display_dataframe()
        except Exception as e:
            messagebox.showerror("Join Error", str(e))

    def apply_filter(self):
        if self.current_df is None:
            return
        try:
            col = self.filter_col_var.get()
            op = self.filter_op_var.get()
            val = self.filter_value_var.get()
            
            if op == "BETWEEN":
                low, high = map(float, val.split(','))
                mask = self.current_df[col].between(low, high)
            elif op == "LIKE":
                mask = self.current_df[col].str.contains(val.replace('%', '.*'))
            elif op == "IN":
                mask = self.current_df[col].isin(val.split(','))
            else:
                mask = eval(f"self.current_df[col] {op} {val}")
            
            self.current_df = self.current_df[mask]
            self.display_dataframe()
        except Exception as e:
            messagebox.showerror("Filter Error", str(e))

    def apply_sort(self):
        if self.current_df is None:
            return
        try:
            col = self.sort_col_var.get()
            ascending = self.sort_order_var.get() == "ASC"
            self.current_df = self.current_df.sort_values(col, ascending=ascending)
            self.display_dataframe()
        except Exception as e:
            messagebox.showerror("Sort Error", str(e))

    def display_dataframe(self):
        self.tree.delete(*self.tree.get_children())
        if self.current_df is not None:
            self.tree["columns"] = list(self.current_df.columns)
            for col in self.current_df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)
            for _, row in self.current_df.iterrows():
                self.tree.insert("", "end", values=list(row))

    def export_data(self):
        if self.current_df is None:
            return
        file = filedialog.asksaveasfilename(defaultextension=".xlsx")
        try:
            self.current_df.to_excel(file, index=False)
            messagebox.showinfo("Success", "Data exported successfully!")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
