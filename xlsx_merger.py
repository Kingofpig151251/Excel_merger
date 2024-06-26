import pandas as pd
import os

def combine_excel_files(root_folder):
    combined_df = pd.DataFrame()
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                df = pd.read_excel(file_path)
                if combined_df.empty:
                    combined_df = df
                else:
                    # 更改欄位名稱, 例如on=['class', 'name']，合併時以 class 和 name 欄位為準
                    combined_df = pd.merge(combined_df, df, on=['class', 'name'], how='outer')
    return combined_df

# 設定資料夾路徑
folder_path = 'C:\\Users\\King\\Documents\\Excel_merger\\Excel_resource'

# 呼叫函數並將結果保存到新的 Excel 檔案
combined_df = combine_excel_files(folder_path)
combined_df.to_excel('merged_Excel.xlsx', index=False)