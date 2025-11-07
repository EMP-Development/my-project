import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


# 列記号(A/B/...) → 列番号(0始まり)
def col_letter_to_index(col_letter):
    col_letter = col_letter.upper()
    index = 0
    for i, c in enumerate(reversed(col_letter)):
        index += (ord(c) - ord('A') + 1) * (26 ** i)
    return index - 1


# シート名をExcelの制約に合わせて調整
def sanitize_sheet_name(name, existing_names):
    # 禁止文字を置換
    invalid_chars = [":", "\\", "/", "?", "*", "[", "]"]
    for ch in invalid_chars:
        name = name.replace(ch, "_")
    # 長さ制限
    name = name[:31]
    # 重複回避
    original = name
    i = 1
    while name in existing_names or name.strip() == "":
        suffix = f"_{i}"
        name = (original[:31 - len(suffix)] + suffix)
        i += 1
    existing_names.add(name)
    return name


# Excel抽出処理
def extract_data_from_excels(input_dir, output_file, target_col_letter, search_list, match_type="partial", separate_sheets=False):
    if separate_sheets:
        results = {s: [] for s in search_list}
    else:
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

                for s in search_list:
                    if match_type == "exact":
                        mask = (col == s)
                    else:
                        mask = col.str.contains(s, na=False)
                    matched = df[mask]
                    if not matched.empty:
                        for _, row in matched.iterrows():
                            row_data = [file.name, sheet_name] + row.tolist()
                            if separate_sheets:
                                results[s].append(row_data)
                            else:
                                results.append(row_data)
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # 出力処理
    if separate_sheets:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            existing_names = set()
            for s, rows in results.items():
                if rows:
                    max_len = max(len(r) for r in rows)
                    rows = [r + [""] * (max_len - len(r)) for r in rows]
                    columns = ["元ファイル名", "元シート名"] + [f"Col{i}" for i in range(1, max_len - 1)]
                    df_out = pd.DataFrame(rows, columns=columns)
                    sheet_name = sanitize_sheet_name(str(s), existing_names)
                    df_out.to_excel(writer, sheet_name=sheet_name, index=False)
        return any(len(v) > 0 for v in results.values())
    else:
        if results:
            max_len = max(len(r) for r in results)
            results = [r + [""] * (max_len - len(r)) for r in results]
            columns = ["元ファイル名", "元シート名"] + [f"Col{i}" for i in range(1, max_len - 1)]
            pd.DataFrame(results, columns=columns).to_excel(output_file, index=False)
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
        output_file_path = str(Path(input_dir) / output_file_name)

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

        success = extract_data_from_excels(
            input_dir, output_file_path, target_col, search_items,
            match_type, separate_sheets_var.get()
        )
        if success:
            messagebox.showinfo("完了", f"抽出結果を {output_file_path} に保存しました")
        else:
            messagebox.showinfo("結果なし", "一致するデータは見つかりませんでした")

    root = tk.Tk()
    root.title("Excel データ抽出ツール")
    root.geometry("600x450")

    # 入力フォルダ
    tk.Label(root, text="入力フォルダ:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    input_dir_var = tk.StringVar(value=str(Path.cwd()))
    tk.Entry(root, textvariable=input_dir_var, width=40).grid(row=0, column=1, padx=5, pady=5)
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

    search_listbox = tk.Listbox(root, height=6, width=40)
    search_listbox.grid(row=4, column=1, padx=5, pady=5, sticky="w")
    tk.Button(root, text="削除", command=remove_search_item).grid(row=4, column=2, padx=5, pady=5)

    # 出力方法（チェックボックス）
    separate_sheets_var = tk.BooleanVar(value=False)
    tk.Checkbutton(root, text="検索結果を別シートで出力する", variable=separate_sheets_var).grid(row=5, column=1, sticky="w", padx=5, pady=5)

    # 一致方法
    tk.Label(root, text="一致方法:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
    match_type_var = tk.StringVar(value="部分一致")
    ttk.Combobox(root, textvariable=match_type_var, values=["部分一致", "完全一致"], width=10, state="readonly").grid(row=6, column=1, padx=5, pady=5, sticky="w")

    # 実行ボタン
    tk.Button(root, text="実行", command=run_extract, bg="lightblue", width=20).grid(row=7, column=1, pady=20)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
