import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd


class ExcelMapperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel データ抽出・転記ツール")
        self.root.geometry("1100x750")

        # ファイル情報
        self.src_file = tk.StringVar()
        self.src_sheet = tk.StringVar()
        self.dst_file = tk.StringVar()
        self.dst_sheet = tk.StringVar()

        # 開始行設定
        self.src_start_row = tk.IntVar(value=1)
        self.dst_start_row = tk.IntVar(value=1)

        # 除外列設定
        self.skip_columns = []  # [(tk.StringVar), ...]

        # 転記設定（元列→先列）
        self.mappings = []  # [(src_col, dst_col, frame)]

        self.key_col = tk.StringVar(value="A")
        self.overwrite = tk.BooleanVar(value=True)

        self.build_ui()

    def build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # === 元データ ===
        src_frame = ttk.LabelFrame(main_frame, text="元データ", padding=10)
        src_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        src_frame.columnconfigure(1, weight=1)

        ttk.Label(src_frame, text="ファイル:").grid(row=0, column=0, sticky="w")
        ttk.Entry(src_frame, textvariable=self.src_file).grid(row=0, column=1, sticky="ew")
        ttk.Button(src_frame, text="参照", command=self.select_src_file).grid(row=0, column=2, padx=5)

        ttk.Label(src_frame, text="シート:").grid(row=1, column=0, sticky="w")
        self.src_sheet_combo = ttk.Combobox(src_frame, textvariable=self.src_sheet, state="readonly")
        self.src_sheet_combo.grid(row=1, column=1, sticky="ew")

        ttk.Label(src_frame, text="データ開始行:").grid(row=2, column=0, sticky="w")
        ttk.Entry(src_frame, textvariable=self.src_start_row, width=5).grid(row=2, column=1, sticky="w")

        ttk.Label(src_frame, text="参照しない列:").grid(row=3, column=0, sticky="w")
        self.skip_frame = ttk.Frame(src_frame)
        self.skip_frame.grid(row=3, column=1, sticky="ew")
        ttk.Button(src_frame, text="追加", command=self.add_skip_column).grid(row=3, column=2)
        ttk.Button(src_frame, text="削除", command=self.remove_skip_column).grid(row=3, column=3)
        self.add_skip_column()

        # === 転記先 ===
        dst_frame = ttk.LabelFrame(main_frame, text="転記先", padding=10)
        dst_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        dst_frame.columnconfigure(1, weight=1)

        ttk.Label(dst_frame, text="ファイル:").grid(row=0, column=0, sticky="w")
        ttk.Entry(dst_frame, textvariable=self.dst_file).grid(row=0, column=1, sticky="ew")
        ttk.Button(dst_frame, text="参照", command=self.select_dst_file).grid(row=0, column=2, padx=5)

        ttk.Label(dst_frame, text="シート:").grid(row=1, column=0, sticky="w")
        self.dst_sheet_combo = ttk.Combobox(dst_frame, textvariable=self.dst_sheet, state="readonly")
        self.dst_sheet_combo.grid(row=1, column=1, sticky="ew")

        ttk.Label(dst_frame, text="データ開始行:").grid(row=2, column=0, sticky="w")
        ttk.Entry(dst_frame, textvariable=self.dst_start_row, width=5).grid(row=2, column=1, sticky="w")

        # === 転記設定 ===
        map_frame = ttk.LabelFrame(main_frame, text="転記設定（元列 → 先列）", padding=10)
        map_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=10)
        map_frame.columnconfigure(0, weight=1)
        map_frame.rowconfigure(0, weight=1)

        self.mapping_frame = ttk.Frame(map_frame)
        self.mapping_frame.grid(row=0, column=0, sticky="nsew")
        self.mapping_frame.columnconfigure(0, weight=1)

        ttk.Button(map_frame, text="設定を追加", command=self.add_mapping_row).grid(row=1, column=0, sticky="e", pady=5)
        ttk.Button(map_frame, text="最後の設定を削除", command=self.remove_mapping_row).grid(row=1, column=0, sticky="w", padx=5)

        self.add_mapping_row()

        # === 下部コントロール ===
        ctrl_frame = ttk.Frame(main_frame)
        ctrl_frame.grid(row=4, column=0, columnspan=2, pady=10)
        ttk.Radiobutton(ctrl_frame, text="上書きする", variable=self.overwrite, value=True).pack(side="left", padx=5)
        ttk.Radiobutton(ctrl_frame, text="上書きしない", variable=self.overwrite, value=False).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="転記実行", command=self.execute_mapping).pack(side="right", padx=10)

    # === ファイル選択 ===
    def select_src_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.src_file.set(path)
            self.load_sheet_names(path, self.src_sheet_combo, self.src_sheet)

    def select_dst_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.dst_file.set(path)
            self.load_sheet_names(path, self.dst_sheet_combo, self.dst_sheet)

    def load_sheet_names(self, path, combo, var):
        try:
            xl = pd.ExcelFile(path)
            combo["values"] = xl.sheet_names
            if xl.sheet_names:
                var.set(xl.sheet_names[0])
        except Exception as e:
            messagebox.showerror("エラー", f"シート名取得に失敗しました:\n{e}")

    # === 除外列設定 ===
    def add_skip_column(self):
        var = tk.StringVar()
        idx = len(self.skip_columns)
        entry = ttk.Entry(self.skip_frame, textvariable=var, width=8)
        entry.grid(row=0, column=idx, padx=3)
        self.skip_columns.append(var)

    def remove_skip_column(self):
        if self.skip_columns:
            var = self.skip_columns.pop()
            for w in self.skip_frame.grid_slaves():
                if str(w.cget("textvariable")) == str(var):
                    w.destroy()
                    break

    # === 転記設定 ===
    def add_mapping_row(self):
        frame = ttk.Frame(self.mapping_frame)
        frame.pack(fill="x", pady=2)

        src_col = tk.StringVar()
        dst_col = tk.StringVar()

        ttk.Entry(frame, textvariable=src_col, width=10).pack(side="left", padx=5)
        ttk.Label(frame, text="→").pack(side="left")
        ttk.Entry(frame, textvariable=dst_col, width=10).pack(side="left", padx=5)

        # 個別削除ボタン
        ttk.Button(frame, text="削除", command=lambda f=frame: self.delete_specific_mapping(f)).pack(side="right", padx=5)

        self.mappings.append((src_col, dst_col, frame))

    def remove_mapping_row(self):
        if self.mappings:
            src_col, dst_col, frame = self.mappings.pop()
            frame.destroy()

    def delete_specific_mapping(self, target_frame):
        """特定の1行だけ削除"""
        for i, (_, _, frame) in enumerate(self.mappings):
            if frame == target_frame:
                frame.destroy()
                del self.mappings[i]
                break

    # === 実行処理（確認） ===
    def execute_mapping(self):
        try:
            src_path = self.src_file.get()
            dst_path = self.dst_file.get()
            src_sheet = self.src_sheet.get()
            dst_sheet = self.dst_sheet.get()

            if not all([src_path, dst_path, src_sheet, dst_sheet]):
                messagebox.showerror("エラー", "ファイルとシートをすべて選択してください。")
                return

            mapping_info = [(m[0].get(), m[1].get()) for m in self.mappings if m[0].get() and m[1].get()]
            skip_info = [s.get() for s in self.skip_columns if s.get()]

            info = (
                f"✅ 実行準備OK\n\n"
                f"元ファイル: {src_path}\n"
                f"元シート: {src_sheet}\n"
                f"開始行: {self.src_start_row.get()}\n"
                f"除外列: {skip_info}\n\n"
                f"先ファイル: {dst_path}\n"
                f"先シート: {dst_sheet}\n"
                f"開始行: {self.dst_start_row.get()}\n\n"
                f"転記設定:\n"
                + "\n".join([f"  {s} → {d}" for s, d in mapping_info])
            )
            messagebox.showinfo("設定確認", info)

        except Exception as e:
            messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMapperApp(root)
    root.mainloop()
