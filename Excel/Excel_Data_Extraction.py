import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def extract_data_from_excels(
    input_dir, 
    output_file, 
    target_col, 
    search_str, 
    match_type="partial"
):
    results = []
    
    for file in Path(input_dir).glob("*.xlsx"):
        try:
            xls = pd.ExcelFile(file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, dtype=str)
                if df.empty:
                    continue
                
                # 列指定
                if target_col.isdigit():  # 数字なら列番号
                    col_idx = int(target_col)
                    if col_idx >= df.shape[1]:
                        continue
                    col = df.iloc[:, col_idx]
                else:  # 列名なら名前で検索
                    if target_col not in df.columns:
                        continue
                    col = df[target_col]
                
                # 検索条件
                if match_type == "exact":
                    mask = col == search_str
                else:  # 部分一致
                    mask = col.str.contains(search_str, na=False)
                
                matched = df[mask]
                if not matched.empty:
                    for _, row in matched.iterrows():
                        results.append(
                            [file.name, sheet_name] + row.tolist()
                        )
        except Exception as e:
            print(f"Error processing {file}: {e}")
    
    if results:
        max_len = max(len(r) for r in results)
        results = [r + [""]*(max_len-len(r)) for r in results]
        
        columns = ["元ファイル名", "元シート名"] + [f"Col{i}" for i in range(1, max_len-1)]
        result_df = pd.DataFrame(results, columns=columns)
        result_df.to_excel(output_file, index=False)
        return True
    else:
        return False


# ----------------- Tkinter GUI -----------------
def run_gui():
    def browse_folder():
        folder = filedialog.askdirectory()
        if folder:
            input_dir_var.set(folder)
    
    def run_extract():
        input_dir = input_dir_var.get()
        output_file = output_file_var.get()
        target_col = target_col_var.get()
        search_str = search_str_var.get()
        match_type = match_type_var.get()
        
        if not input_dir or not Path(input_dir).exists():
            messagebox.showerror("エラー", "入力フォルダを選択してください")
            return
        if not output_file:
            messagebox.showerror("エラー", "出力ファイル名を入力してください")
            return
        if not target_col:
            messagebox.showerror("エラー", "検索対象の列を入力してください")
            return
        if not search_str:
            messagebox.showerror("エラー", "検索文字列を入力してください")
            return
        
        success = extract_data_from_excels(input_dir, output_file, target_col, search_str, match_type)
        if success:
            messagebox.showinfo("完了", f"抽出結果を {output_file} に保存しました")
        else:
            messagebox.showinfo("結果なし", "一致するデータは見つかりませんでした")
    
    root = tk.Tk()
    root.title("Excel データ抽出ツール")
    root.geometry("500x300")
    
    # 入力フォルダ
    tk.Label(root, text="入力フォルダ:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    input_dir_var = tk.StringVar()
    tk.Entry(root, textvariable=input_dir_var, width=40).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="参照", command=browse_folder).grid(row=0, column=2, padx=5, pady=5)
    
    # 出力ファイル
    tk.Label(root, text="出力ファイル名:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    output_file_var = tk.StringVar(value="まとめ結果.xlsx")
    tk.Entry(root, textvariable=output_file_var, width=40).grid(row=1, column=1, padx=5, pady=5)
    
    # 検索列
    tk.Label(root, text="検索対象列 (列名 or 列番号):").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    target_col_var = tk.StringVar()
    tk.Entry(root, textvariable=target_col_var, width=20).grid(row=2, column=1, padx=5, pady=5, sticky="w")
    
    # 検索文字列
    tk.Label(root, text="検索文字列:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
    search_str_var = tk.StringVar()
    tk.Entry(root, textvariable=search_str_var, width=20).grid(row=3, column=1, padx=5, pady=5, sticky="w")
    
    # 一致方法
    tk.Label(root, text="一致方法:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
    match_type_var = tk.StringVar(value="partial")
    ttk.Combobox(root, textvariable=match_type_var, values=["partial", "exact"], width=10).grid(row=4, column=1, padx=5, pady=5, sticky="w")
    
    # 実行ボタン
    tk.Button(root, text="実行", command=run_extract, bg="lightblue").grid(row=5, column=1, pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    run_gui()
