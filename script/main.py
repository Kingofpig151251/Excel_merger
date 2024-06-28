import json
import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


class Model:
    def __init__(self):
        self.folder_path = ""
        self.reference_file_path = ""
        self.selected_columns = []

    def load_reference_file(self):
        try:
            df = pd.read_excel(self.reference_file_path)
            return df.columns
        except Exception as e:
            raise e

    def merge_files(self):
        dataframes = []

        for root, dirs, files in os.walk(self.folder_path):
            for file in files:
                if file.endswith((".xlsx", ".xls")):
                    try:
                        df = pd.read_excel(os.path.join(root, file))
                        df.columns = [col.replace(" ", "") for col in df.columns]
                        dataframes.append(df)
                    except Exception as e:
                        print(f"Error processing {file}: {str(e)}")

        if dataframes:
            from functools import reduce

            merged_df = reduce(
                lambda left, right: pd.merge(
                    left, right, on=self.selected_columns, how="outer"
                ),
                dataframes,
            )
        else:
            merged_df = None

        return merged_df


class View(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.language = "en"
        self.load_translations()
        self.title(self.translate("Title"))
        self.geometry("400x400")
        if getattr(sys, "frozen", False):
            script_dir = sys._MEIPASS
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "Merger.ico")
        self.iconbitmap(icon_path)
        self.widgets = {}
        self.create_widgets()

    def load_translations(self):
        if getattr(sys, "frozen", False):
            script_dir = sys._MEIPASS
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        languages_path = os.path.join(script_dir, "languages.json")
        with open(languages_path, "r", encoding="utf-8") as file:
            self.translations = json.load(file)

    def translate(self, text):
        return self.translations[text][self.language]

    def create_widgets(self):
        self.widgets["select_folder_button"] = tk.Button(
            self,
            text=self.translate("Select Folder"),
            command=self.controller.select_folder,
        )
        self.widgets["select_folder_button"].pack(pady=5)

        self.widgets["select_folder_label"] = tk.Label(self, text="")
        self.widgets["select_folder_label"].pack(pady=5)

        self.widgets["select_reference_button"] = tk.Button(
            self,
            text=self.translate("Select Reference File"),
            command=self.controller.select_reference_file,
        )
        self.widgets["select_reference_button"].pack(pady=5)

        self.widgets["select_reference_label"] = tk.Label(self, text="")
        self.widgets["select_reference_label"].pack(pady=5)

        self.columns_listbox = tk.Listbox(self, selectmode=tk.MULTIPLE)
        self.columns_listbox.pack(pady=5)

        self.widgets["merge_files_button"] = tk.Button(
            self,
            text=self.translate("Merge Files"),
            command=self.controller.merge_files,
        )
        self.widgets["merge_files_button"].pack(pady=5)

        buttons_frame = tk.Frame(self)
        buttons_frame.pack(pady=5)

        self.widgets["help_button"] = tk.Button(
            buttons_frame, text=self.translate("Help"), command=self.show_help
        )
        self.widgets["help_button"].grid(row=0, column=0, padx=5)

        self.widgets["toggle_language_button"] = tk.Button(
            buttons_frame,
            text=self.translate("Toggle Language"),
            command=self.toggle_language,
        )
        self.widgets["toggle_language_button"].grid(row=0, column=1, padx=5)

    def toggle_language(self):
        self.language = "zh" if self.language == "en" else "en"
        self.update_ui()

    def update_ui(self):
        self.title(self.translate("Title"))
        self.widgets["select_folder_button"]["text"] = self.translate("Select Folder")
        self.widgets["select_reference_button"]["text"] = self.translate(
            "Select Reference File"
        )
        self.widgets["merge_files_button"]["text"] = self.translate("Merge Files")
        self.widgets["help_button"]["text"] = self.translate("Help")
        self.widgets["toggle_language_button"]["text"] = self.translate(
            "Toggle Language"
        )

    def update_columns_listbox(self, columns):
        self.columns_listbox.delete(0, tk.END)
        for column in columns:
            self.columns_listbox.insert(tk.END, column)

    def show_help(self):
        help_text = self.translate("Help Text")
        messagebox.showinfo(self.translate("How to use Excel Merger:"), help_text)

    def update_folder_label(self, select_folder):
        self.widgets["select_folder_label"]["text"] += select_folder

    def update_reference_label(self, reference_file):
        self.widgets["select_reference_label"]["text"] += reference_file


class Controller:
    def __init__(self):
        self.model = Model()
        self.view = View(self)

    def select_folder(self):
        self.model.folder_path = filedialog.askdirectory()
        self.view.update_folder_label(os.path.basename(self.model.folder_path))

    def select_reference_file(self):
        self.model.reference_file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        self.view.update_reference_label(
            os.path.basename(self.model.reference_file_path)
        )
        try:
            columns = self.model.load_reference_file()
            self.view.update_columns_listbox(columns)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def merge_files(self):
        self.model.selected_columns = [
            self.view.columns_listbox.get(i)
            for i in self.view.columns_listbox.curselection()
        ]
        if not self.model.selected_columns:
            messagebox.showerror("Error", "No columns selected")
            return
        merged_df = self.model.merge_files()
        if merged_df is not None:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
            )
            if save_path:
                merged_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", "Files merged successfully")
                os.startfile(os.path.dirname(save_path))
        else:
            messagebox.showerror("Error", "Failed to merge files or no files to merge")


if __name__ == "__main__":
    controller = Controller()
    controller.view.mainloop()
