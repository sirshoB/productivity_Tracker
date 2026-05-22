import tkinter as tk
from tkinter import filedialog, messagebox, ttk
 
import pandas as pd
from openpyxl import Workbook
 
 
class AppUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Productivity Tracker")
        self.root.geometry("1200x800")
 
        self.input_df = None
        self.result_df = None
        self.table = None
 
        self._build_ui()
 
   
    def _build_ui(self):
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        tk.Button(
            button_frame,
            text="Select 1st Excel",
            command=self.select_first_excel
            ).pack(side="left", padx=5)
        tk.Button(
            button_frame,
            text="Select 2nd Excel and Start",
            command=self.select_second_excel_and_start
            ).pack(side="left", padx=5)
        #print("2nd excel read success")
        tk.Button(
            self.root,
            text="Download Excel",
            command=self.export_table_to_excel
            ).pack(pady=5)
 
 
 
   
    def select_first_excel(self):
        excel_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
            )
        if not excel_path:
            return
        self.input_df = pd.read_excel(excel_path, sheet_name="Sheet1")
        self.input_df.columns = self.input_df.columns.str.strip()
        self.input_df["Emp Id"] = (self.input_df["Emp Id"].astype(str).str.replace(".0", "", regex=False).str.strip())
 
        messagebox.showinfo("Success", "1st Excel selected successfully.")
    def select_second_excel_and_start(self):
        excel_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
            )
        if not excel_path:
            return
        self.second_df = pd.read_excel(excel_path,sheet_name="DATA")
        self.second_df.columns = self.second_df.columns.str.strip()
        #print(self.second_df.column.)
        self.second_df["Emp Id"]=(self.second_df["Emp Id"].astype(str).str.replace(".0","",regex=False).str.strip())
        #print(self.second_df.columns)
        if self.input_df is None:
            messagebox.showwarning("Missing File", "Please select 1st Excel first.")
            return
        self.calculate_productivity()
        self.show_table()
 
 
    
    def calculate_productivity(self):
 
        
 
        self.input_df.columns = [
        "Date",
        "Emp Id",
        "Emp Name",
        "Supervisor",
        "Primary Queue",
        "Primary Productivity Target",
        "Primary Assigned for the day",
        "Primary Assignment Completed",
        "Primary Deficit",
        "Secondary Queue",
        "Secondary Productivity Target",
        "Secondary Assigned for the day",
        "Secondary Assignment Completed",
        "Secondary Deficit",
        "Reason of not meeting Numbers",
        "Hours Audited",
        "Adjusted Productivity Hours",
        "Existing Total Productivity Hours",
        "Productivity % for the day",
        "Supervisor Agreement"
    ]
 
        df = self.input_df
 
 
 
        df["Primary Productivity Target"] = pd.to_numeric(df["Primary Productivity Target"], errors="coerce").fillna(0)
        df["Primary Assigned for the day"] = pd.to_numeric(df["Primary Assigned for the day"], errors="coerce").fillna(0)
 
        df["Hourly Target"] = df["Primary Productivity Target"] / 8
 
        df["Productivity Hours"] = df.apply(
            lambda row: 0 if row["Hourly Target"] == 0
            else row["Primary Assigned for the day"] / row["Hourly Target"],
            axis=1
        )
 
        df["Secondary Productivity Target"] = pd.to_numeric(df["Secondary Productivity Target"], errors="coerce").fillna(0)
        df["Secondary Assigned for the day"] = pd.to_numeric(df["Secondary Assigned for the day"], errors="coerce").fillna(0)
 
        df["Secondary Hourly Target"] = df["Secondary Productivity Target"] / 8
 
        df["Secondary Productivity Hours"] = df.apply(
            lambda row: 0 if row["Secondary Hourly Target"] == 0
            else row["Secondary Assigned for the day"] / row["Secondary Hourly Target"],
            axis=1
        )
        print("--- 2nd file Minute column dtype:", self.second_df["Minute"].dtype)
        print("--- Minute sample values:", self.second_df["Minute"].dropna().head(3).tolist())
        df["Adjusted Productivity Hours"] = df.apply(self.calculate_adjusted_prod, axis=1)
        df["Total Productivity Hours"] = (
            df["Productivity Hours"] + df["Secondary Productivity Hours"] + df["Adjusted Productivity Hours"]
        )

        for col in ["Hourly Target", "Productivity Hours", "Secondary Hourly Target",
                    "Secondary Productivity Hours", "Adjusted Productivity Hours", "Total Productivity Hours"]:
            df[col] = df[col].round(2)
 
        self.result_df = df[
            [
                "Emp Id",
                "Emp Name",
                "Supervisor",
                "Primary Queue",
                "Primary Productivity Target",
                "Primary Assigned for the day",
                "Hourly Target",
                "Productivity Hours",
                "Secondary Queue",
                "Secondary Productivity Target",
                "Secondary Assigned for the day",
                "Secondary Hourly Target",
                "Secondary Productivity Hours",
                "Adjusted Productivity Hours",
                "Total Productivity Hours",
               
            ]
        ]
    
    # def calculate_adjusted_prod(self, row):
    #     sequential_df = self.second_df.copy()
    #     emp_id = str(row["Emp Id"]).replace(".0", "").strip()
    #     work_date = pd.to_datetime(row["Date"],errors="coerce").date()
    #     sequential_df["Emp Id"] = (sequential_df["Emp Id"].astype(str).str.replace(".0", "", regex=False).str.strip())
    #     sequential_df["Date"] = pd.to_datetime(sequential_df["Date"],errors="coerce").dt.date
    #     sequential_df["Category1"] = (sequential_df["Category1"].astype(str).str.strip().str.lower())
    #     sequential_df["Reason (Ops Tracker)"] = (sequential_df["Reason (Ops Tracker)"].astype(str).str.strip().str.lower())
    #     sequential_df["Minute"] = pd.to_numeric(sequential_df["Minute"],errors="coerce").fillna(0)
    #     selected_categories = ["aux", "idle time"]
    #     selected_reasons = ["coaching","quality","special projects","system down","team meeting","training","calibration"]
    #     filtered_df = sequential_df[(sequential_df["Emp Id"] == emp_id) & (sequential_df["Date"] == work_date)
    #                                 & (sequential_df["Category1"].isin(selected_categories))
    #                                 & (sequential_df["Reason (Ops Tracker)"].isin(selected_reasons))]
    #     total_minutes = filtered_df["Minute"].sum()
    #     return total_minutes / 60
    def calculate_adjusted_prod(self, row):
        sequential_df = self.second_df.copy()
        emp_id = str(row["Emp Id"]).replace(".0", "").strip()
        work_date = pd.to_datetime(row["Date"], errors="coerce").date()
        sequential_df["Emp Id"] = (sequential_df["Emp Id"].astype(str).str.replace(".0", "", regex=False).str.strip())
        sequential_df["Date"] = pd.to_datetime(sequential_df["Date"], errors="coerce").dt.date
        sequential_df["Category1"] = (sequential_df["Category1"].astype(str).str.strip().str.lower())
        sequential_df["Reason (Ops Tracker)"] = (sequential_df["Reason (Ops Tracker)"].astype(str).str.strip().str.lower())

        minute_col = sequential_df["Minute"]
        if pd.api.types.is_timedelta64_dtype(minute_col):
            sequential_df["Minute"] = minute_col.dt.total_seconds() / 60
        elif pd.api.types.is_datetime64_any_dtype(minute_col):
            sequential_df["Minute"] = minute_col.dt.hour * 60 + minute_col.dt.minute + minute_col.dt.second / 60
        else:
            sequential_df["Minute"] = pd.to_numeric(minute_col, errors="coerce").fillna(0)

        selected_categories = ["aux", "idle time"]
        selected_reasons = ["coaching", "quality", "special projects", "system down", "team meeting", "training", "calibration"]
        filtered_df = sequential_df[
            (sequential_df["Emp Id"] == emp_id) &
            (sequential_df["Date"] == work_date) &
            (sequential_df["Category1"].isin(selected_categories)) &
            (sequential_df["Reason (Ops Tracker)"].isin(selected_reasons))
        ]
        total_minutes = filtered_df["Minute"].sum()
        print(f"EMP ID: {emp_id} | Work date: {work_date} | Matched rows: {len(filtered_df)} | Total minutes: {total_minutes:.2f}")
        if len(filtered_df) == 0:
            emp_rows = sequential_df[sequential_df["Emp Id"] == emp_id]
            if emp_rows.empty:
                print("  Emp Id NOT FOUND in 2nd file at all.")
            else:
                print("  Sample dates in 2nd file:", emp_rows["Date"].unique()[:3].tolist())
                print("  Sample Category1:", emp_rows["Category1"].unique()[:5].tolist())
                print("  Sample Reason:", emp_rows["Reason (Ops Tracker)"].unique()[:5].tolist())
        return total_minutes / 60
    
 
 
        
   
 
    def show_table(self):
        if self.table:
            self.table.destroy()
 
        columns = ["Sl No"] + list(self.result_df.columns)
 
        table_frame = tk.Frame(self.root)
        table_frame.pack(padx=20, pady=10, fill="both", expand=True)
 
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal")
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical")
 
        self.table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=25,
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set
        )
 
        x_scroll.config(command=self.table.xview)
        y_scroll.config(command=self.table.yview)
 
        self.table.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
 
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
 
        for col in columns:
            self.table.heading(col, text=col)
            self.table.column(col, width=180, anchor="center", stretch=False)
 
        for serial_no, (_, row) in enumerate(self.result_df.iterrows(), start=1):
            self.table.insert("", "end", values=[serial_no] + list(row))
 
 
    def export_table_to_excel(self):
        if self.result_df is None:
            messagebox.showwarning("No Data", "Please calculate productivity first.")
            return
 
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Productivity Report"
        )
 
        if not file_path:
            return
 
        self.result_df.to_excel(file_path, index=False)
 
        messagebox.showinfo("Success", "Productivity report exported successfully!")
 
 
if __name__ == "__main__":
    root = tk.Tk()
    app = AppUI(root)
    root.mainloop()
 
 