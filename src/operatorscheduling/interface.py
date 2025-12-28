import csv
import sys
import os
import subprocess
import tkinter as tk
import tksheet
from tkinter import ttk
import customtkinter as ctk

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        
        self.schedule_file = ctk.StringVar(value="No Schedule File selected")
        self.wg_schedule_file= ctk.StringVar(value="No Working Group File selected")
        self.selected_operator: str = "ALL"
        self.selected_date: str = ""
        self.send_meeting: bool = False
        self.selected_row = None
        self.selected_column = None

        self.geometry("800x600")
        self.title("Working schedule reminder maker")

        self.schedule_select_button = ctk.CTkButton(self, text="Select Schedule File", command=lambda: self.select_file("schedule_file"))
        self.schedule_select_button.pack(padx=5, pady=30)

        self.schedule_label = ctk.CTkLabel(self, textvariable=self.schedule_file, text_color="white")
        self.schedule_label.pack(padx=10, pady=5)

        self.working_group_select_button = ctk.CTkButton(self, text="Select Working Group File", command=lambda: self.select_file("wg_schedule_file"))
        self.working_group_select_button.pack(padx=5, pady=40)

        self.working_group_schedule_label = ctk.CTkLabel(self, textvariable=self.wg_schedule_file, text_color="white")
        self.working_group_schedule_label.pack(padx=10, pady=5)

        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = None

        self.run_button = ctk.CTkButton(self, text="Run", command=self.run_program)
        self.run_button.pack()

    def run_program(self):
        subprocess.run([
                sys.executable,
                "src/operatorscheduling/main.py",
                "--input",
                self.schedule_file.get(),
                "--service",
                self.wg_schedule_file.get(),
                ])

    def on_selection(self, event=None):
        selected_rows = self.sheet.get_selected_rows()
        selected_columns = self.sheet.get_selected_columns()

        if selected_rows:
            self.selected_row = self.rows[selected_rows.pop()]
            self.selected_column = None
            print(f"Row selected: {self.selected_row}")
            print(f"Column selected: {self.selected_column}")
        if selected_columns:
            self.selected_row = None
            self.selected_column = self.headers[selected_columns.pop()]
            print(f"Column selected: {self.selected_column}")
            print(f"Row selected: {self.selected_row}")

    def select_file(self, attr):
        print(f"File {attr} selection")
        cl_attr = getattr(self, attr)
        file_path = ctk.filedialog.askopenfilename()
        if not file_path:
            return

        if file_path:
            file_name = os.path.basename(file_path)
            cl_attr.set(file_name)

        self.load_csv(file_path)

    def load_csv(self, file_path):

        if self.tree:
            self.tree.destroy()

        with open(file_path, newline="", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            self.headers = next(reader)
            self.rows = list(reader)
        
        self.sheet = tksheet.Sheet(self.table_frame)
        self.sheet.pack(fill="both", expand=True)

        self.sheet.headers(self.headers)
        self.sheet.set_sheet_data(self.rows)

        self.sheet.enable_bindings(
            "single_select",
            "row_select",
            "column_select",
            "arrowkeys",
            "column_width_resize",
            "rc_delete_column",
            "delete"
            )
        self.sheet.bind("<<SheetSelect>>", self.on_selection)
        


if __name__ == "__main__":
    app = App()
    app.mainloop()




