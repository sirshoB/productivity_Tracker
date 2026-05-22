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
        tk.Button(
            self.root,
            text="Select Excel & Calculate Productivity",
            command=self.start
        ).pack(pady=10)

        tk.Button(
            self.root,
            text="Download Excel",
            command=self.export_table_to_excel
        ).pack(pady=5)

    def start(self):
        excel_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not excel_path:
            return

        self.input_df = pd.read_excel(excel_path, sheet_name="Sheet1")
        self.input_df = pd.read_excel(excel_path)
        self.input_df.columns = self.input_df.columns.str.strip()
        #print(self.input_df.columns.tolist())

        self.calculate_productivity()
        self.show_table()

    # def calculate_productivity(self):
    #     df = self.input_df.copy()

    #     #self.input_df = pd.read_excel(excel_path)
        
    #     #self.input_df.columns = [
    #         #"Date",
    #         #"Emp ID",
    #         #"Emp Name",
    #         #"Supervisor",
    #         #"Primary Queue",
    #         #"Primary Productivity Target",
    #         #"Primary Assigned for the day",
    #         #"Primary Assignment Completed",
    #         #"Secondary Queue",
    #         #"Secondary Productivity Target",
    #         #"Secondary Assigned for the day",
    #         #"Secondary Assignment Completed"
    #     #]

    #     print(df.columns)

    #     self.input_df.columns = [
    #     "Date",
    #     "Emp ID",
    #     "Emp Name",
    #     "Supervisor",
    #     "Primary Queue",
    #     "Primary Productivity Target",
    #     "Primary Assigned for the day",
    #     "Primary Assignment Completed",
    #     "Secondary Queue",
    #     "Secondary Productivity Target",
    #     "Secondary Assigned for the day",
    #     "Secondary Assignment Completed",
    #     "Col13",
    #     "Col14",
    #     "Col15",
    #     "Col16",
    #     "Col17",
    #     "Col18",
    #     "Col19",
    #     "Col20",
    #     ]
    #     print(self.input_df.columns)

        
    #     self.input_df["Hourly Target"] = self.input_df["Primary Productivity Target"] / 8

    #     self.input_df["Primary Assigned for the day"] = pd.to_numeric(self.input_df["Primary Assigned for the day"], errors = 'coerce')
        
    #     self.input_df["Productivity Hours"] = self.input_df.apply(
    #         lambda row: 0
    #         if row["Hourly Target"] == 0
    #         else float(row["Primary Assigned for the day"]) / float(row["Hourly Target"]),
    #         axis=1
    #     )

    #     self.input_df["Secondary Productivity Target"] = pd.to_numeric(self.input_df["Secondary Productivity Target"],errors="coerce").fillna(0)
    #     self.input_df["Secondary Hourly Target"] = self.input_df["Secondary Productivity Target"] / 8

    #     self.input_df["Secondary Assigned for the day"] = pd.to_numeric(self.input_df["Secondary Assigned for the day"],errors="coerce").fillna(0)

    #     self.input_df["Secondary Productivity Hours"] = self.input_df.apply(
    #         lambda row: 0
    #         if row["Secondary Hourly Target"] == 0
    #         else float(row["Secondary Assigned for the day"]) / float(row["Secondary Hourly Target"]),
    #         axis=1
    #     )

    #     #Final total productivity hours
    #     self.input_df["Total Productivity Hours"] = (self.input_df["Productivity Hours"] + self.input_df["Secondary Productivity Hours"])

    #     self.result_df = self.input_df[
    #         [
    #             "Emp Name",
    #             "Supervisor",
    #             "Primary Productivity Target",
    #             "Primary Assigned for the day",
    #             "Hourly Target",
    #             "Productivity Hours",
    #             "Secondary Hourly Target",
    #             "Secondary Productivity Target",
    #             "Secondary Productivity Hours",
    #             "Total Productivity Hours"
    #         ]
    #     ]
    def calculate_productivity(self):
        self.input_df.columns = [
        "Date",
        "Emp ID",
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

        df["Total Productivity Hours"] = (
            df["Productivity Hours"] + df["Adjusted Productivity Hours"]
        )

        self.result_df = df[
            [
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

    # def show_table(self):
    #     if self.table:
    #         self.table.destroy()

    #     #columns = list(self.result_df.columns)
    #     columns = ["Sl No"] + list(self.result_df.columns)

    #     self.table = ttk.Treeview(
    #         self.root,
    #         columns=columns,
    #         show="headings",
    #         height=25
    #     )

    #     for col in columns:
    #         self.table.heading(col, text=col)
    #         self.table.column(col, width=180, anchor="center")

    #     for index, row in self.result_df.iterrows():
    #         self.table.insert("", "end", values=[index + 1] + list(row))

    #     self.table.pack(padx=20, pady=10, fill="both", expand=True)

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