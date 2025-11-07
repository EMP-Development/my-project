import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

class ExcelMapperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel データ抽出・転記ツール")
        self.root.geometry("800x600")

        self.source_file = tk.StringVar()
        self.dest_file = tk.StringVar()
        self.key_column = tk.StringVar(value='A')
        self.overwrite = tk.BooleanVar(value=True)

        self.mappings = []  # (source_col, dest_col)

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        # --- ファイル選択 ---
        ttk.Label(frame, text="コピー元ファイル:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.source_file, width=60).grid(row=0, column=1)
        ttk.Button(frame, text="参照", command=self.select_source_file).grid(row=0, column=2)

        ttk.Label(frame, text="コピー先ファイル:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.dest_file, width=60).grid(row=1, column=1)
        ttk.Button(frame, text="参照", command=self.select_dest_file).grid(row=1, column=2)

        # --- キー列設定 ---
        ttk.Label(frame, text="キー列（社員番号）:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.key_column, width=10).grid(row=2, column=1, sticky="w")

        # --- 上書き設定 ---
        ttk.Label(frame, text="既存データの扱い:").grid(row=3, column=0, sticky="w")
        ttk.Radiobutton(frame, text="上書きする", variable=self.overwrite, value=True).grid(row=3, column=1, sticky="w")
        ttk.Radiobutton(frame, text="上書きしない", variable=self.overwrite, value=False).grid(row=3, column=1, sticky="e")

        # --- マッピング設定 ---
        ttk.Label(frame, text="転記設定（元列 → 先列）:").grid(row=4, column=0, sticky="w")

        self.mapping_frame = ttk.Frame(frame)
        self.mapping_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=5)
        self.mapping_frame.columnconfigure(1, weight=1)

        ttk.Button(frame, text="設定を追加", command=self.add_mapping_row).grid(row=6, column=1, sticky="e")
        ttk.Button(frame, text="設定を削除", command=self.remove_mapping_row).grid(row=6, column=2, sticky="w")

        # --- 実行ボタン ---
        ttk.Button(frame, text="転記実行", command=self.execute_mapping).grid(row=7, column=1, pady=20)

        self.add_mapping_row()

    def select_source_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.source_file.set(path)

    def select_dest_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.dest_file.set(path)

    def add_mapping_row(self):
        row = len(self.mappings)
        src_col = tk.StringVar()
        dst_col = tk.StringVar()
        ttk.Entry(self.mapping_frame, textvariable=src_col, width=10).grid(row=row, column=0, padx=5)
        ttk.Label(self.mapping_frame, text="→").grid(row=row, column=1)
        ttk.Entry(self.mapping_frame, textvariable=dst_col, width=10).grid(row=row, column=2, padx=5)
        self.mappings.append((src_col, dst_col))

    def remove_mapping_row(self):
        if self.mappings:
            src_col, dst_col = self.mappings.pop()
            for widget in self.mapping_frame.grid_slaves(row=len(self.mappings)):
                widget.destroy()

    def execute_mapping(self):
        try:
            src_path = self.source_file.get()
            dst_path = self.dest_file.get()
            key_col = self.key_column.get()
            overwrite = self.overwrite.get()

            if not src_path or not dst_path:
                messagebox.showerror("エラー", "ファイルを選択してください。")
                return

            src_df = pd.read_excel(src_path, sheet_name=0)
            dst_df = pd.read_excel(dst_path, sheet_name=0)

            for src_col, dst_col in self.mappings:
                s = src_col.get().strip()
                d = dst_col.get().strip()
                if s and d:
                    merged = dst_df.set_index(key_col).combine_first(
                        src_df.set_index(key_col)
                    )
                    for idx, row in src_df.iterrows():
                        key_val = row[key_col]
                        if key_val in dst_df[key_col].values:
                            if overwrite or pd.isna(dst_df.loc[dst_df[key_col]==key_val, d].values[0]):
                                val = row.get(s, None)
                                dst_df.loc[dst_df[key_col]==key_val, d] = val

            dst_df.to_excel(dst_path, index=False)
            messagebox.showinfo("完了", "転記が完了しました。")

        except Exception as e:
            messagebox.showerror("エラー", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMapperApp(root)
    root.mainloop()
