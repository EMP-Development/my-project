import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 列記号 → 列番号
def col_letter_to_index(col_letter):
    col_letter = col_letter.upper()
    index = 0
    for i, c in enumerate(reversed(col_letter)):
        index += (ord(c) - ord('A') + 1) * (26 ** i)
    return index - 1

# Excel抽出処理
def extract_data_from_excels(input_dir, output_file, target_col_letter, search_list, match_type="partial"):
    results = []
    col_idx = col_letter_to_index(target_col_letter)

    for file in Path(input_dir).glob("*.xlsx"):
        try:
            xls = pd.ExcelFile(file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, dtype=str)
                if df.empty or col_idx >= df.shape[1]:
                    continue
                col = df.iloc[:, col_idx].astype(str).str.strip()

                mask = pd.Series([False]*len(df))
                for s in search_list:
                    if not s:
                        continue
                    if match_type == "exact":
                        mask |= (col == s)
                    else:
                        mask |= col.str.contains(s, na=False)
                matched = df[mask]
                if not matched.empty:
                    for _, row in matched.iterrows():
                        results.append([file.name, sheet_name] + row.tolist())
        except Exception as e:
            print(f"Error processing {file}: {e}")

    if results:
        max_len = max(len(r) for r in results)
        results = [r + [""]*(max_len-len(r)) for r in results]
        columns = ["元ファイル名", "元シート名"] + [f"Col{i}" for i in range(1, max_len-1)]
        result_df = pd.DataFrame(results, columns=columns)
        result_df.to_excel(output_file, index=False)
        return True
    return False

# ----------------- Tkinter GUI -----------------
def run_gui():
    def browse_folder():
        folder = filedialog.askdirectory()
        if folder:
            input_dir_var.set(folder)

    def add_search_item():
        s = search_entry.get().strip()
        if s and s not in search_listbox.get(0, tk.END):
            search_listbox.insert(tk.END, s)
        search_entry.delete(0, tk.END)

    def remove_search_item():
        sel = search_listbox.curselection()
        for i in reversed(sel):
            search_listbox.delete(i)

    def run_extract():
        input_dir = input_dir_var.get()
        target_col = target_col_var.get().strip().upper()
        search_items = list(search_listbox.get(0, tk.END))
        match_type_jp = match_type_var.get()
        match_type = "partial" if match_type_jp == "部分一致" else "exact"

        output_file_name = output_file_var.get().strip()
        output_file_path = str(Path(input_dir) / output_file_name)  # 入力フォルダ内に作成

        if not input_dir:
            messagebox.showerror("エラー", "入力フォルダを指定してください")
            return
        if not output_file_name:
            messagebox.showerror("エラー", "出力ファイル名を入力してください")
            return
        if not target_col:
            messagebox.showerror("エラー", "検索対象列を入力してください")
            return
        if not search_items:
            messagebox.showerror("エラー", "検索文字列を追加してください")
            return

        success = extract_data_from_excels(input_dir, output_file_path, target_col, search_items, match_type)
        if success:
            messagebox.showinfo("完了", f"抽出結果を {output_file_path} に保存しました")
        else:
            messagebox.showinfo("結果なし", "一致するデータは見つかりませんでした")

    root = tk.Tk()
    root.title("Excel データ抽出ツール")
    root.geometry("550x400")

    # 入力フォルダ
    tk.Label(root, text="入力フォルダ:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    input_dir_var = tk.StringVar(value=str(Path.cwd()))
    tk.Entry(root, textvariable=input_dir_var, width=35).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="参照", command=browse_folder).grid(row=0, column=2, padx=5, pady=5)

    # 出力ファイル
    tk.Label(root, text="出力ファイル名:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    output_file_var = tk.StringVar(value="まとめ結果.xlsx")
    tk.Entry(root, textvariable=output_file_var, width=40).grid(row=1, column=1, padx=5, pady=5, columnspan=2)

    # 検索列
    tk.Label(root, text="検索対象列 (列記号 A/B/C …):").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    target_col_var = tk.StringVar(value="B")
    tk.Entry(root, textvariable=target_col_var, width=10).grid(row=2, column=1, padx=5, pady=5, sticky="w")

    # 検索文字列
    tk.Label(root, text="検索文字列:").grid(row=3, column=0, sticky="ne", padx=5, pady=5)
    search_entry = tk.Entry(root, width=30)
    search_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
    tk.Button(root, text="追加", command=add_search_item).grid(row=3, column=2, padx=5, pady=5)

    search_listbox = tk.Listbox(root, height=6, width=30)
    search_listbox.grid(row=4, column=1, padx=5, pady=5, sticky="w")
    tk.Button(root, text="削除", command=remove_search_item).grid(row=4, column=2, padx=5, pady=5)

    # 一致方法
    tk.Label(root, text="一致方法:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
    match_type_var = tk.StringVar(value="部分一致")
    ttk.Combobox(root, textvariable=match_type_var, values=["部分一致", "完全一致"], width=10, state="readonly").grid(row=5, column=1, padx=5, pady=5, sticky="w")

    # 実行ボタン
    tk.Button(root, text="実行", command=run_extract, bg="lightblue", width=15).grid(row=6, column=1, pady=20)

    root.mainloop()

if __name__ == "__main__":
    run_gui()
