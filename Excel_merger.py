import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

texts_chinese = {
    "title": "Excel合併工具",
    "instruction_text": """使用說明：
1. 選擇資源文件夾：包含所有需要合併的Excel文件。
2. 選擇輸出文件夾：合併後的Excel文件將被保存在此文件夾。
3. 選擇參考Excel：選擇一個Excel文件以選擇合併的鍵（列名）。
4. 開始合併：點擊後開始合併過程，完成後會顯示提示。""",
    "switch_language": "English",
    "select_input_folder": "選擇資源文件夾",
    "select_output_folder": "選擇輸出文件夾",
    "select_reference_file": "選擇參考Excel",
    "entry_label": "輸出文件名：",
    "start_merge": "開始合併",
    "column_selection_dialog_title": "選擇欄名",
    "column_selection_instruction": "請選擇要作為合併鍵的列名：",
    "column_selection_confirm_button": "確認",
    "warning_all_fields_title": "警告",
    "warning_all_fields_message": "請確保所有欄位都已正確填寫！",
    "merge_complete_title": "完成",
    "merge_complete_message": "Excel文件已合併完成！",
}

texts_english = {
    "title": "Excel Merger",
    "instruction_text": """Instructions:
1. Select input folder: Contains all Excel files to be merged.
2. Select output folder: The merged Excel file will be saved in this folder.
3. Select reference Excel: Select an Excel file to choose the merge keys (column names).
4. Start merge: Click to start the merge process, a prompt will be shown when completed.""",
    "switch_language": "中文",
    "select_input_folder": "Select input folder",
    "select_output_folder": "Select output folder",
    "select_reference_file": "Select reference Excel",
    "entry_label": "Output file name:",
    "start_merge": "Start merge",
    "column_selection_dialog_title": "Select columns",
    "column_selection_instruction": "Please select the column names to be used as merge keys:",
    "column_selection_confirm_button": "Confirm",
    "warning_all_fields_title": "Warning",
    "warning_all_fields_message": "Please make sure all fields are correctly filled in!",
    "merge_complete_title": "Complete",
    "merge_complete_message": "Excel files have been merged!",
}


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.current_language = "Chinese"
        self.texts = texts_chinese
        self.setup_ui()
        self.update_texts()

    def setup_ui(self):
        self.root.geometry("550x450")
        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.merge_keys_var = tk.StringVar()
        self.custom_output_name_var = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        self.create_language_switch_button()
        self.create_instructions_label()
        self.create_folder_selection_buttons()
        self.create_output_name_label()
        self.create_output_name_entry()
        self.create_start_merge_button()

    def create_language_switch_button(self):
        self.switch_language_button = tk.Button(
            self.root,
            command=self.toggle_language,
            width=10,
            pady=5,
        )
        self.switch_language_button.pack(anchor="ne")

    def create_instructions_label(self):
        tk.Label(
            self.root,
            justify="left",
            padx=10,
            pady=10,
        ).pack()

    def create_folder_selection_buttons(self):
        self.create_button("select_input_folder", self.select_input_folder)
        tk.Label(
            self.root,
            textvariable=self.input_path_var,
            width=50,
            pady=5,
            wraplength=500,
        ).pack()
        self.create_button("select_output_folder", self.select_output_folder)
        tk.Label(
            self.root,
            textvariable=self.output_path_var,
            width=50,
            pady=5,
            wraplength=500,
        ).pack()
        self.create_button("select_reference_file", self.select_reference_file)
        tk.Label(
            self.root,
            textvariable=self.merge_keys_var,
            width=50,
            pady=5,
            wraplength=500,
        ).pack()

    def create_output_name_label(self):
        tk.Label(self.root, text=self.texts["entry_label"]).pack()

    def create_output_name_entry(self):
        tk.Label(self.root, width=50).pack(pady=0)
        tk.Entry(self.root, textvariable=self.custom_output_name_var, width=30).pack(
            pady=(0, 10)
        )

    def create_start_merge_button(self):
        self.create_button("start_merge", self.start_merge)

    def create_button(self, text_key, command):
        tk.Button(
            self.root,
            text=self.texts.get(text_key, ""),
            command=command,
            width=20,
            pady=5,
        ).pack()

    def select_input_folder(self):
        self.input_path_var.set(filedialog.askdirectory())

    def select_output_folder(self):
        self.output_path_var.set(filedialog.askdirectory())

    def select_reference_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            df = pd.read_excel(file_path)
            self.show_column_selection_dialog(df.columns.tolist())

    def show_column_selection_dialog(self, columns):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.texts["column_selection_dialog_title"])
        dialog.geometry("400x400+100+100")

        instruction_label = tk.Label(
            dialog, text=self.texts["column_selection_instruction"], pady=10
        )
        instruction_label.pack()

        container = tk.Frame(dialog)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", on_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        var_dict = {}
        for column in columns:
            var = tk.BooleanVar()
            tk.Checkbutton(scrollable_frame, text=column, variable=var).pack(anchor="w")
            var_dict[column] = var

        confirm_button = tk.Button(
            dialog,
            text=self.texts["column_selection_confirm_button"],
            command=lambda: self.on_column_selection_confirmed(var_dict, dialog),
        )
        confirm_button.pack(pady=10)

    def on_column_selection_confirmed(self, var_dict, dialog):
        selected_columns = [column for column, var in var_dict.items() if var.get()]
        self.merge_keys_var.set(",".join(selected_columns))
        dialog.destroy()

    def start_merge(self):
        input_folder = self.input_path_var.get()
        output_folder = self.output_path_var.get()
        merge_keys_str = self.merge_keys_var.get()
        merge_keys = [key.strip() for key in merge_keys_str.split(",")]
        if input_folder and output_folder and merge_keys:
            self.combine_excel_files(input_folder, output_folder, merge_keys)
        else:
            messagebox.showwarning(
                self.texts["warning_all_fields_title"],
                self.texts["warning_all_fields_message"],
            )

    def combine_excel_files(self, root_folder, output_folder, merge_keys):
        custom_output_name = self.custom_output_name_var.get().strip()
        if not custom_output_name.endswith(".xlsx"):
            custom_output_name += ".xlsx"
        output_file_name = (
            custom_output_name if custom_output_name else "merged_Excel.xlsx"
        )
        output_path = os.path.join(output_folder, output_file_name)
        combined_df = pd.DataFrame()
        for root, dirs, files in os.walk(root_folder):
            for file in files:
                if file.endswith((".xlsx", ".xls")):
                    file_path = os.path.join(root, file)
                    df = pd.read_excel(file_path)
                    df.columns = df.columns.str.strip()
                    if combined_df.empty:
                        combined_df = df
                    else:
                        combined_df = pd.merge(
                            combined_df, df, on=merge_keys, how="outer"
                        )
        combined_df.to_excel(output_path, index=False)
        messagebox.showinfo(
            self.texts["merge_complete_title"], self.texts["merge_complete_message"]
        )
        os.startfile(output_folder)

    def toggle_language(self):
        if self.current_language == "Chinese":
            self.current_language = "English"
            self.texts = texts_english
        else:
            self.current_language = "Chinese"
            self.texts = texts_chinese
        self.update_texts()

    def update_texts(self):
        self.root.title(self.texts["title"])
        self.root.children["!button"].config(text=self.texts["switch_language"])
        self.root.children["!button2"].config(text=self.texts["select_input_folder"])
        self.root.children["!button3"].config(text=self.texts["select_output_folder"])
        self.root.children["!button4"].config(text=self.texts["select_reference_file"])
        self.root.children["!button5"].config(text=self.texts["start_merge"])

        self.root.children["!label"].config(text=self.texts["instruction_text"])
        self.root.children["!label2"].config(textvariable=self.input_path_var)
        self.root.children["!label3"].config(textvariable=self.output_path_var)
        self.root.children["!label4"].config(textvariable=self.merge_keys_var)
        self.root.children["!label5"].config(text=self.texts["entry_label"])


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
