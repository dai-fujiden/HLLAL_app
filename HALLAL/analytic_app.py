import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

# ASCIIアートの生成
# 'HALLAL' をかっこいいフォントで表示
ascii_art = """
  _   _    _    _     _    _    _
 | | | |  / \  | |   | |  / \  | |
 | |_| | / _ \ | |   | | / _ \ | |
 |  _  |/ ___ \| |___| |/ ___ \| |
 |_| |_/_/   \_\_____/_/_/   \_\_|
"""

# ロードされたDataFrameを保存するためのグローバル辞書
loaded_dataframes = {}
# 現在表示されているDataFrameのファイルパスとシート名を追跡するための変数
current_dataframe_path = None
current_dataframe_sheet = None

# グローバルリスト：Treeviewのルートとなるディレクトリのパスを保持
# これにより、Treeviewがクリアされても元のディレクトリ構造を再構築できる
global_root_directories = []

# 現在の作業ディレクトリ (CWD) を保持するグローバル変数
# アプリケーションのルートパスとして使用される
global_current_working_directory = os.path.dirname(os.path.abspath(__file__))

# CWD履歴 (メモリ上に保持)
global_cwd_history = [global_current_working_directory]


# --- Tkinter Style 設定 ---
def configure_styles():
    style = ttk.Style()
    # 全体的なテーマ設定
    style.theme_use('clam') # 'clam'はモダンな見た目を提供します

    # フレームのスタイル
    style.configure('TFrame', background='#F0F2F5', borderwidth=0)
    style.configure('White.TFrame', background='#FFFFFF', borderwidth=1, relief='solid')
    style.configure('LightGray.TFrame', background='#E0E0E0', borderwidth=1, relief='solid')
    style.configure('Accent.TFrame', background='#E6F2FF', borderwidth=0) # アクセントカラーのフレーム

    # ラベルのスタイル
    style.configure('TLabel', background='#F0F2F5', foreground='#333333', font=('Arial', 10))
    style.configure('Title.TLabel', font=('Arial', 24, 'bold'), foreground='#2C3E50', background='#F0F2F5') # より濃い青系のタイトル
    style.configure('Header.TLabel', font=('Arial', 16, 'bold'), foreground='#34495E', background='#E0E0E0') # より濃いヘッダー
    style.configure('SubHeader.TLabel', font=('Arial', 12, 'bold'), foreground='#333333', background='#F0F2F5')
    style.configure('CurrentFile.TLabel', font=('Arial', 16, 'bold'), foreground='#2C3E50', background='#FFFFFF')
    style.configure('Note.TLabel', font=('Arial', 9), foreground='#7F8C8D', background='#FFFFFF') # 落ち着いたグレー

    # ボタンのスタイル
    style.configure('TButton', font=('Arial', 12, 'bold'), foreground='#FFFFFF', background='#3498DB', relief='flat', borderwidth=0, padding=8) # 青色のボタン
    style.map('TButton',
              background=[('active', '#2980B9')],
              foreground=[('active', '#FFFFFF')])
    style.configure('Green.TButton', background='#2ECC71') # 緑色のボタン
    style.map('Green.TButton',
              background=[('active', '#27AE60')])
    style.configure('Gray.TButton', background='#95A5A6') # グレーのボタン
    style.map('Gray.TButton',
              background=[('active', '#7F8C8D')])

    # Treeviewのスタイル
    style.configure('Treeview', background='#FFFFFF', foreground='#333333', fieldbackground='#FFFFFF', borderwidth=0, relief='flat', rowheight=25)
    style.map('Treeview', background=[('selected', '#3498DB')]) # 選択時の背景色を青に
    style.configure('Treeview.Heading', font=('Arial', 10, 'bold'), background='#BDC3C7', foreground='#2C3E50', padding=5) # ヘッダーを明るいグレーに

    # スクロールバーのスタイル
    style.configure('TScrollbar', troughcolor='#ECF0F1', background='#BDC3C7', borderwidth=0, relief='flat')
    style.map('TScrollbar', background=[('active', '#95A5A6')])

    # TEntryのスタイル
    style.configure('TEntry', fieldbackground='#FFFFFF', foreground='#333333', borderwidth=1, relief='solid', padding=5)
    # TComboboxのスタイル
    style.configure('TCombobox', fieldbackground='#FFFFFF', foreground='#333333', selectbackground='#3498DB', selectforeground='#FFFFFF', borderwidth=1, relief='solid')
    style.map('TCombobox',
              fieldbackground=[('readonly', '#FFFFFF')],
              selectbackground=[('readonly', '#3498DB')],
              selectforeground=[('readonly', '#FFFFFF')])

    # TCheckbutton, TRadiobuttonのスタイル
    style.configure('TCheckbutton', background='#E0E0E0', foreground='#333333', indicatorcolor='#3498DB', padding=5)
    style.map('TCheckbutton',
              background=[('active', '#D0D0D0')],
              foreground=[('active', '#333333')])
    style.configure('TRadiobutton', background='#E0E0E0', foreground='#333333', indicatorcolor='#3498DB', padding=5)
    style.map('TRadiobutton',
              background=[('active', '#D0D0D0')],
              foreground=[('active', '#333333')])


# --- パス操作ヘルパー関数 ---
def get_relative_path(absolute_path, base_path):
    """
    絶対パスを基準パスからの相対パスに変換する。
    相対パスにできない場合は絶対パスを返す。
    """
    try:
        return os.path.relpath(absolute_path, base_path)
    except ValueError:
        return absolute_path # 異なるドライブなど、相対パスにできない場合

def get_absolute_path(relative_path, base_path):
    """
    相対パスを基準パスからの絶対パスに変換する。
    """
    return os.path.abspath(os.path.join(base_path, relative_path))


# --- データ処理ヘルパー関数 ---
def get_supported_files_in_directory(directory):
    """指定されたディレクトリ内のサポートされているファイルをリストアップする。"""
    files = []
    try:
        for item in os.listdir(directory):
            path = os.path.join(directory, item)
            if os.path.isfile(path):
                if item.lower().endswith(('.csv', '.h5', '.hdf', '.xlsx', '.xls')):
                    files.append(path)
    except PermissionError:
        messagebox.showwarning("アクセス拒否", f"ディレクトリ '{directory}' へのアクセスが拒否されました。")
    except Exception as e:
        messagebox.showerror("エラー", f"ディレクトリ '{directory}' の読み込み中にエラーが発生しました: {e}")
    return files

def prompt_for_excel_sheet(parent_window, file_path):
    """Excelファイルの場合、シート名を選択するダイアログを表示する。"""
    try:
        excel_sheets = pd.ExcelFile(file_path).sheet_names
    except Exception as e:
        messagebox.showerror("エラー", f"Excelシートの読み込み中にエラーが発生しました: {e}")
        return None

    if not excel_sheets:
        messagebox.showwarning("警告", "このExcelファイルにはシートが見つかりません。")
        return None

    sheet_dialog = tk.Toplevel(parent_window)
    sheet_dialog.title("Excelシートを選択")
    sheet_dialog.geometry("300x200")
    sheet_dialog.grab_set() # モーダルにする
    sheet_dialog.transient(parent_window)
    sheet_dialog.configure(bg="#F0F2F5")

    ttk.Label(sheet_dialog, text="シートを選択してください:", style='SubHeader.TLabel', background="#F0F2F5").pack(pady=10)

    sheet_listbox = tk.Listbox(sheet_dialog, bg="#FFFFFF", fg="#333333", selectbackground="#3498DB", selectforeground="#FFFFFF", font=("Courier", 10), height=5, relief="flat", bd=1)
    for sheet in excel_sheets:
        sheet_listbox.insert(tk.END, sheet)
    sheet_listbox.pack(padx=10, pady=5, fill="both", expand=True)

    selected_sheet = None
    def on_sheet_select_dialog():
        nonlocal selected_sheet
        if sheet_listbox.curselection():
            selected_sheet = sheet_listbox.get(sheet_listbox.curselection())
            sheet_dialog.destroy()
        else:
            messagebox.showwarning("警告", "シートを選択してください。")

    select_button = ttk.Button(
        sheet_dialog,
        text="選択",
        command=on_sheet_select_dialog,
        style='TButton',
        cursor="hand2"
    )
    select_button.pack(pady=10)

    parent_window.wait_window(sheet_dialog)
    return selected_sheet

def load_and_display_dataframe(file_path, sheet_name=None, dataframe_text_widget=None, current_file_label_widget=None, 
                               start_row_entry=None, end_row_entry=None, start_col_entry=None, end_col_entry=None,
                               row_label_entry=None, col_label_entry=None):
    """
    指定されたファイルをデータフレームとしてロードし、右パネルに表示する。
    """
    global current_dataframe_path, current_dataframe_sheet
    df_key = file_path
    if sheet_name:
        df_key += f"_{sheet_name}"

    if df_key in loaded_dataframes:
        df = loaded_dataframes[df_key]
    else:
        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.lower().endswith(('.h5', '.hdf')):
                df = pd.read_hdf(file_path)
            elif file_path.lower().endswith(('.xlsx', '.xls')):
                if not sheet_name:
                    messagebox.showerror("エラー", "Excelファイルにはシート名の指定が必要です。")
                    return
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            loaded_dataframes[df_key] = df
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました: {e}")
            if dataframe_text_widget:
                dataframe_text_widget.config(state="normal")
                dataframe_text_widget.delete("1.0", tk.END)
                dataframe_text_widget.insert(tk.END, f"エラー: ファイルをロードできませんでした。\n{e}")
                dataframe_text_widget.config(state="disabled")
            if current_file_label_widget:
                current_file_label_widget.config(text="エラー: ファイルロード")
            return

    current_dataframe_path = file_path
    current_dataframe_sheet = sheet_name

    current_file_display_name = os.path.basename(file_path)
    if sheet_name:
        current_file_display_name += f" (シート: {sheet_name})"
    if current_file_label_widget:
        current_file_label_widget.config(text=f"現在のファイル: {current_file_display_name}")

    # エントリーをクリアし、デフォルトで全範囲を表示
    if start_row_entry: start_row_entry.delete(0, tk.END)
    if end_row_entry: end_row_entry.delete(0, tk.END)
    if start_col_entry: start_col_entry.delete(0, tk.END)
    if end_col_entry: end_col_entry.delete(0, tk.END)
    if row_label_entry: row_label_entry.delete(0, tk.END)
    if col_label_entry: col_label_entry.delete(0, tk.END)

    # デフォルトで最初の20行と全列を表示
    display_df = df.iloc[:20, :]
    if dataframe_text_widget:
        dataframe_text_widget.config(state="normal")
        dataframe_text_widget.delete("1.0", tk.END)
        dataframe_text_widget.insert(tk.END, display_df.to_string())
        dataframe_text_widget.config(state="disabled")

def display_dataframe_content(dataframe_text_widget, current_file_label_widget, 
                              start_row_entry, end_row_entry, start_col_entry, end_col_entry,
                              row_label_entry, col_label_entry):
    """
    入力された行/列の範囲またはラベルに基づいてデータフレームを表示する。
    """
    global current_dataframe_path, current_dataframe_sheet
    if current_dataframe_path is None:
        messagebox.showwarning("警告", "表示するファイルが選択されていません。")
        return

    df_key = current_dataframe_path
    if current_dataframe_sheet:
        df_key += f"_{current_dataframe_sheet}"

    if df_key not in loaded_dataframes:
        messagebox.showerror("エラー", "データフレームがロードされていません。")
        return

    df = loaded_dataframes[df_key]

    try:
        # 行の指定方法を決定
        row_selection = slice(None) # デフォルトは全行
        row_label_input = row_label_entry.get().strip()
        start_row_idx_input = start_row_entry.get().strip()
        end_row_idx_input = end_row_entry.get().strip()

        if row_label_input:
            # 行ラベルによる検索 (locを使用)
            if row_label_input in df.index:
                row_selection = row_label_input
            else:
                messagebox.showwarning("警告", f"指定された行ラベル '{row_label_input}' は見つかりませんでした。")
                return
        elif start_row_idx_input or end_row_idx_input:
            # 行インデックスによるスライス (ilocを使用)
            start_row = int(start_row_idx_input) if start_row_idx_input else 0
            end_row = int(end_row_idx_input) if end_row_idx_input else None
            row_selection = slice(start_row, end_row)
            if start_row < 0:
                messagebox.showerror("エラー", "開始行は0以上の数値を入力してください。")
                return
            if end_row is not None and end_row < start_row:
                messagebox.showerror("エラー", "終了行は開始行以上の数値を入力してください。")
                return

        # 列の指定方法を決定
        col_selection = slice(None) # デフォルトは全列
        col_label_input = col_label_entry.get().strip()
        start_col_idx_input = start_col_entry.get().strip()
        end_col_idx_input = end_col_entry.get().strip()

        if col_label_input:
            # 列ラベルによる検索 (locを使用)
            if col_label_input in df.columns:
                col_selection = col_label_input
            else:
                messagebox.showwarning("警告", f"指定された列ラベル '{col_label_input}' は見つかりませんでした。")
                return
        elif start_col_idx_input or end_col_idx_input:
            # 列インデックスによるスライス (ilocを使用)
            start_col = int(start_col_idx_input) if start_col_idx_input else 0
            end_col = int(end_col_idx_input) if end_col_idx_input else None
            col_selection = slice(start_col, end_col)

    except ValueError as e:
        messagebox.showerror("エラー", f"入力値が無効です: {e}\n行/列は数値インデックスまたはラベルを入力してください。")
        return
    except KeyError as e: # locでラベルが見つからなかった場合
        messagebox.showerror("エラー", f"指定されたラベルが見つかりません: {e}")
        return
    except Exception as e:
        messagebox.showerror("エラー", f"データフレームの表示中に予期せぬエラーが発生しました: {e}")
        return

    display_df = None
    try:
        # 行と列の選択を組み合わせる
        if isinstance(row_selection, slice) and isinstance(col_selection, slice):
            display_df = df.iloc[row_selection, col_selection]
        elif isinstance(row_selection, str) and isinstance(col_selection, slice):
            # 特定の行ラベルと列スライス
            display_df = df.loc[[row_selection], col_selection] # 行ラベルはリストで渡す
        elif isinstance(row_selection, slice) and isinstance(col_selection, str):
            # 行スライスと特定の列ラベル
            display_df = df.loc[row_selection, [col_selection]] # 列ラベルはリストで渡す
        elif isinstance(row_selection, str) and isinstance(col_selection, str):
            # 特定の行ラベルと列ラベル
            display_df = df.loc[[row_selection], [col_selection]]
        else:
            messagebox.showerror("エラー", "行と列の選択の組み合わせが不正です。")
            return

    except IndexError as e:
        messagebox.showerror("エラー", f"指定された行または列の範囲がデータフレームの範囲外です: {e}")
        return
    except Exception as e:
        messagebox.showerror("エラー", f"データフレームの表示中に予期せぬエラーが発生しました: {e}")
        return

    dataframe_text_widget.config(state="normal")
    dataframe_text_widget.delete("1.0", tk.END)
    dataframe_text_widget.insert(tk.END, display_df.to_string())
    dataframe_text_widget.config(state="disabled")

# --- ページ定義 ---
def show_file_processing_page(initial_directory_paths=None):
    """
    ファイルリストとデータフレーム表示機能を持つページを表示する関数。
    複数の初期ディレクトリパスをリストとして受け取る。
    """
    file_processing_page = tk.Tk()
    file_processing_page.title(f"データ分析 - ファイル処理")
    file_processing_page.geometry("1200x800")
    file_processing_page.configure(bg="#F0F2F5") # 明るい背景

    # スタイル設定を適用
    configure_styles()

    # グリッドレイアウトを設定
    file_processing_page.columnconfigure(0, weight=1) # 左パネル（ファイルツリー）
    file_processing_page.columnconfigure(1, weight=3) # 右パネル（データフレーム表示）
    file_processing_page.rowconfigure(0, weight=0) # 検索バーとフィルター行
    file_processing_page.rowconfigure(1, weight=1) # メインコンテンツ行
    file_processing_page.rowconfigure(2, weight=0) # 戻るボタン行

    # --- 検索バー、ディレクトリ追加、フィルターオプション ---
    top_controls_frame = ttk.Frame(file_processing_page, style='LightGray.TFrame')
    top_controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
    top_controls_frame.columnconfigure(0, weight=1) # 検索エントリー
    top_controls_frame.columnconfigure(1, weight=0) # 検索ボタン
    top_controls_frame.columnconfigure(2, weight=0) # ディレクトリ追加ボタン
    top_controls_frame.columnconfigure(3, weight=0) # 拡張子フィルター
    top_controls_frame.columnconfigure(4, weight=0) # 検索オプション

    # 検索エントリーとボタン
    search_entry = ttk.Entry(top_controls_frame, width=50, style='TEntry')
    search_entry.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    search_entry.bind("<Return>", lambda event: filter_treeview())

    search_button = ttk.Button(
        top_controls_frame,
        text="検索",
        command=lambda: filter_treeview(),
        style='TButton',
        cursor="hand2"
    )
    search_button.grid(row=0, column=1, sticky="e", padx=5, pady=5)

    # ディレクトリ追加ボタン
    def add_new_directory_action():
        """
        新しいディレクトリを選択し、グローバルリストに追加してTreeviewを更新する。
        これにより、複数のディレクトリを順次追加して管理できる。
        """
        new_dir = filedialog.askdirectory()
        if new_dir:
            if new_dir not in global_root_directories:
                global_root_directories.append(new_dir)
                filter_treeview() # 新しいディレクトリを追加したらツリーを再構築
            else:
                messagebox.showinfo("情報", "このディレクトリは既にリストに追加されています。")

    add_dir_button = ttk.Button(
        top_controls_frame,
        text="ディレクトリを追加",
        command=add_new_directory_action,
        style='Green.TButton',
        cursor="hand2"
    )
    add_dir_button.grid(row=0, column=2, sticky="e", padx=5, pady=5)

    # 拡張子フィルターチェックボックス
    extension_filter_frame = ttk.Frame(top_controls_frame, style='LightGray.TFrame')
    extension_filter_frame.grid(row=0, column=3, sticky="ew", padx=10, pady=5)

    csv_var = tk.BooleanVar(value=True)
    h5_var = tk.BooleanVar(value=True)
    excel_var = tk.BooleanVar(value=True)

    ttk.Checkbutton(extension_filter_frame, text="CSV", variable=csv_var, command=lambda: filter_treeview(), style='TCheckbutton').pack(side="left", padx=2)
    ttk.Checkbutton(extension_filter_frame, text="H5", variable=h5_var, command=lambda: filter_treeview(), style='TCheckbutton').pack(side="left", padx=2)
    ttk.Checkbutton(extension_filter_frame, text="Excel", variable=excel_var, command=lambda: filter_treeview(), style='TCheckbutton').pack(side="left", padx=2)

    # 検索オプション（範囲とタイプ）
    search_options_frame = ttk.Frame(top_controls_frame, style='LightGray.TFrame')
    search_options_frame.grid(row=0, column=4, sticky="ew", padx=10, pady=5)

    search_scope_var = tk.StringVar(value="all") # all, directories, files
    search_type_var = tk.StringVar(value="partial") # partial, full

    ttk.Label(search_options_frame, text="範囲:", style='TLabel').pack(side="left", padx=5)
    ttk.Radiobutton(search_options_frame, text="全て", variable=search_scope_var, value="all", command=lambda: filter_treeview(), style='TRadiobutton').pack(side="left")
    ttk.Radiobutton(search_options_frame, text="Dir", variable=search_scope_var, value="directories", command=lambda: filter_treeview(), style='TRadiobutton').pack(side="left")
    ttk.Radiobutton(search_options_frame, text="File", variable=search_scope_var, value="files", command=lambda: filter_treeview(), style='TRadiobutton').pack(side="left")

    ttk.Label(search_options_frame, text="タイプ:", style='TLabel').pack(side="left", padx=10)
    ttk.Radiobutton(search_options_frame, text="部分", variable=search_type_var, value="partial", command=lambda: filter_treeview(), style='TRadiobutton').pack(side="left")
    ttk.Radiobutton(search_options_frame, text="完全", variable=search_type_var, value="full", command=lambda: filter_treeview(), style='TRadiobutton').pack(side="left")


    # --- 左パネル: ファイルツリー ---
    left_frame = ttk.Frame(file_processing_page, style='White.TFrame')
    left_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
    left_frame.columnconfigure(0, weight=1)
    left_frame.rowconfigure(0, weight=1)

    file_tree = ttk.Treeview(left_frame, style='Treeview')
    file_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

    tree_scrollbar_y = ttk.Scrollbar(left_frame, orient="vertical", command=file_tree.yview, style='TScrollbar')
    tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
    file_tree.configure(yscrollcommand=tree_scrollbar_y.set)

    tree_scrollbar_x = ttk.Scrollbar(left_frame, orient="horizontal", command=file_tree.xview, style='TScrollbar')
    tree_scrollbar_x.grid(row=1, column=0, sticky="ew")
    file_tree.configure(xscrollcommand=tree_scrollbar_x.set)

    def add_files_to_treeview(tree, current_dir, parent_iid, search_term="", active_extensions=None, search_scope="all", search_type="partial"):
        """Treeviewにファイルとサブディレクトリを再帰的に追加する。"""
        if active_extensions is None:
            active_extensions = {'.csv', '.h5', '.hdf', '.xlsx', '.xls'} # デフォルトで全て有効

        try:
            items = sorted(os.listdir(current_dir), key=lambda s: (not os.path.isdir(os.path.join(current_dir, s)), s.lower()))
            for item_name in items:
                path = os.path.join(current_dir, item_name)
                
                is_dir = os.path.isdir(path)
                is_file = os.path.isfile(path)
                
                match_name = False
                if search_type == "partial":
                    match_name = search_term in item_name.lower()
                else: # full
                    match_name = search_term == item_name.lower()

                if is_dir:
                    if (search_scope == "all" or search_scope == "directories") and (match_name or not search_term):
                        # ディレクトリ名を相対パスで表示
                        display_name = get_relative_path(path, global_current_working_directory)
                        dir_iid = tree.insert(parent_iid, "end", text=display_name, open=False, tags=("directory",))
                        add_files_to_treeview(tree, path, dir_iid, search_term, active_extensions, search_scope, search_type)
                elif is_file:
                    file_ext = os.path.splitext(item_name)[1].lower()
                    if file_ext in active_extensions:
                        if (search_scope == "all" or search_scope == "files") and (match_name or not search_term):
                            # ファイル名を相対パスで表示（ただし、valuesには絶対パスを格納）
                            display_name = get_relative_path(path, global_current_working_directory)
                            tree.insert(parent_iid, "end", text=display_name, values=(path,), tags=("file",))
        except PermissionError:
            print(f"Permission denied: {current_dir}")
        except Exception as e:
            print(f"Error listing directory {current_dir}: {e}")

    def filter_treeview():
        """検索条件とフィルターに基づいてTreeviewを再構築する。"""
        search_term = search_entry.get().lower()
        
        active_extensions = set()
        if csv_var.get(): active_extensions.update(['.csv'])
        if h5_var.get(): active_extensions.update(['.h5', '.hdf'])
        if excel_var.get(): active_extensions.update(['.xlsx', '.xls'])

        search_scope_val = search_scope_var.get()
        search_type_val = search_type_var.get()

        # Treeviewの現在の内容をクリア
        for item in file_tree.get_children():
            file_tree.delete(item)

        # グローバルに保存されているルートディレクトリパスからTreeviewを再構築
        for root_path in global_root_directories:
            # ルートパス自体が検索条件に一致するかをチェック（ディレクトリ検索の場合）
            root_name = os.path.basename(root_path)
            root_match = False
            if search_type_val == "partial":
                root_match = search_term in root_name.lower()
            else: # full
                root_match = search_term == root_name.lower()

            # ディレクトリ検索の場合、ルートパスが一致しないならスキップ
            if search_scope_val == "directories" and not (root_match or not search_term):
                continue 

            # ルートディレクトリ名を相対パスで表示
            display_root_name = get_relative_path(root_path, global_current_working_directory)
            root_item_id = file_tree.insert("", "end", text=display_root_name, open=True, tags=("directory",))
            add_files_to_treeview(file_tree, root_path, root_item_id, search_term, active_extensions, search_scope_val, search_type_val)

    # 初期ディレクトリの追加 (初回起動時のみ)
    if initial_directory_paths: # リストとして受け取る
        for path in initial_directory_paths:
            if path and path not in global_root_directories:
                global_root_directories.append(path)
    
    # Treeviewを初期表示または再フィルタリング
    filter_treeview()

    # --- 右パネル: データフレーム表示と操作 ---
    right_frame = ttk.Frame(file_processing_page, style='White.TFrame')
    right_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
    right_frame.rowconfigure(0, weight=0) # 現在のファイルラベル
    right_frame.rowconfigure(1, weight=1) # データフレームテキスト表示
    right_frame.rowconfigure(2, weight=0) # スクロールバー
    right_frame.rowconfigure(3, weight=0) # 範囲入力フレーム
    right_frame.rowconfigure(4, weight=0) # 注意書き
    right_frame.columnconfigure(0, weight=1)

    current_file_label = ttk.Label(
        right_frame,
        text="ファイルが選択されていません",
        style='CurrentFile.TLabel'
    )
    current_file_label.grid(row=0, column=0, sticky="ew", pady=10)

    dataframe_text = tk.Text(
        right_frame,
        wrap="none",
        bg="#F8F8F8", # データ表示エリアは少し暗く
        fg="#333333",
        font=("Courier", 10),
        relief="flat", bd=0,
        state="disabled"
    )
    dataframe_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

    dataframe_scrollbar_y = ttk.Scrollbar(right_frame, orient="vertical", command=dataframe_text.yview, style='TScrollbar')
    dataframe_scrollbar_y.grid(row=1, column=1, sticky="ns")
    dataframe_text.config(yscrollcommand=dataframe_scrollbar_y.set)

    dataframe_scrollbar_x = ttk.Scrollbar(right_frame, orient="horizontal", command=dataframe_text.xview, style='TScrollbar')
    dataframe_scrollbar_x.grid(row=2, column=0, sticky="ew")
    dataframe_text.config(xscrollcommand=dataframe_scrollbar_x.set)

    # 行/列範囲入力フレーム
    range_frame = ttk.Frame(right_frame, style='White.TFrame')
    range_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
    
    # 行の範囲とラベル
    ttk.Label(range_frame, text="行(idx):", style='TLabel').grid(row=0, column=0, sticky="w", padx=5)
    start_row_entry = ttk.Entry(range_frame, width=5, style='TEntry')
    start_row_entry.grid(row=0, column=1, sticky="ew", padx=2)
    ttk.Label(range_frame, text=":", style='TLabel').grid(row=0, column=2, sticky="w")
    end_row_entry = ttk.Entry(range_frame, width=5, style='TEntry')
    end_row_entry.grid(row=0, column=3, sticky="ew", padx=2)

    ttk.Label(range_frame, text="行(Lbl):", style='TLabel').grid(row=1, column=0, sticky="w", padx=5, pady=2)
    row_label_entry = ttk.Entry(range_frame, width=15, style='TEntry')
    row_label_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=2, pady=2)

    # 列の範囲とラベル
    ttk.Label(range_frame, text="列(idx):", style='TLabel').grid(row=0, column=4, sticky="w", padx=10)
    start_col_entry = ttk.Entry(range_frame, width=5, style='TEntry')
    start_col_entry.grid(row=0, column=5, sticky="ew", padx=2)
    ttk.Label(range_frame, text=":", style='TLabel').grid(row=0, column=6, sticky="w")
    end_col_entry = ttk.Entry(range_frame, width=5, style='TEntry')
    end_col_entry.grid(row=0, column=7, sticky="ew", padx=2)

    ttk.Label(range_frame, text="列(Lbl):", style='TLabel').grid(row=1, column=4, sticky="w", padx=10, pady=2)
    col_label_entry = ttk.Entry(range_frame, width=15, style='TEntry')
    col_label_entry.grid(row=1, column=5, columnspan=3, sticky="ew", padx=2, pady=2)

    display_button = ttk.Button(
        range_frame,
        text="表示",
        command=lambda: display_dataframe_content(dataframe_text, current_file_label, 
                                                 start_row_entry, end_row_entry, 
                                                 start_col_entry, end_col_entry,
                                                 row_label_entry, col_label_entry),
        style='TButton',
        cursor="hand2"
    )
    display_button.grid(row=0, column=8, rowspan=2, sticky="nsew", padx=10, pady=5) # 2行にまたがる

    # DataFrameの行/列選択に関する注意書き
    ttk.Label(
        right_frame,
        text="注: 行/列のドラッグ選択はTextウィジェットでは直接サポートされていません。\n数値インデックスまたはラベル、または '開始:終了' 形式で範囲指定して表示してください。",
        style='Note.TLabel'
    ).grid(row=4, column=0, sticky="ew", pady=5)

    def on_tree_select(event):
        """Treeviewでアイテムが選択されたときのイベントハンドラ。"""
        selected_item_id = file_tree.selection()
        if not selected_item_id:
            return
        
        item_tags = file_tree.item(selected_item_id, "tags")
        
        if "file" in item_tags:
            file_path = file_tree.item(selected_item_id, "values")[0] # valuesには絶対パスが格納されている
            sheet_name = None
            if file_path.lower().endswith(('.xlsx', '.xls')):
                sheet_name = prompt_for_excel_sheet(file_processing_page, file_path)
                if sheet_name is None:
                    file_tree.selection_remove(selected_item_id)
                    return
            load_and_display_dataframe(file_path, sheet_name, dataframe_text, current_file_label, 
                                       start_row_entry, end_row_entry, start_col_entry, end_col_entry,
                                       row_label_entry, col_label_entry)
        elif "directory" in item_tags:
            if file_tree.item(selected_item_id, "open"):
                file_tree.item(selected_item_id, open=False)
            else:
                file_tree.item(selected_item_id, open=True)

    file_tree.bind("<<TreeviewSelect>>", on_tree_select)

    # --- 戻るボタン ---
    back_button = ttk.Button(
        file_processing_page,
        text="戻る",
        command=lambda: [file_processing_page.destroy(), create_start_page()],
        style='Gray.TButton',
        cursor="hand2"
    )
    back_button.grid(row=2, column=0, columnspan=2, pady=10)

    file_processing_page.mainloop()


def show_directory_selection_page(current_working_directory):
    """
    ディレクトリ選択ページを表示する関数。
    現在の作業ディレクトリを引数として受け取る。
    """
    directory_selection_page = tk.Tk()
    directory_selection_page.title("データ分析 - ディレクトリを選択")
    directory_selection_page.geometry("800x550") # 高さを少し増やす
    directory_selection_page.configure(bg="#F0F2F5") # 明るいグレーの背景

    # スタイル設定を適用
    configure_styles() 

    selected_directories_list = [] # 選択されたディレクトリパスを保持するリスト (絶対パス)

    ttk.Label(
        directory_selection_page,
        text="作業するディレクトリを選択してください",
        style='Title.TLabel'
    ).pack(pady=30)

    # 選択されたディレクトリを表示するリストボックス
    selected_dirs_frame = ttk.Frame(directory_selection_page, style='White.TFrame')
    selected_dirs_frame.pack(pady=10, padx=20, fill="both", expand=True)
    selected_dirs_frame.columnconfigure(0, weight=1)
    selected_dirs_frame.rowconfigure(0, weight=1)

    selected_dirs_listbox = tk.Listbox(
        selected_dirs_frame,
        selectmode="extended", # 複数選択を可能にする (Ctrl/Shiftクリック)
        exportselection=False, # 選択が他のウィジェットに影響しないようにする
        bg="#FFFFFF", fg="#333333", selectbackground="#3498DB", selectforeground="#FFFFFF",
        font=("Arial", 10), relief="solid", bd=1
    )
    selected_dirs_listbox.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

    dirs_scrollbar_y = ttk.Scrollbar(selected_dirs_frame, orient="vertical", command=selected_dirs_listbox.yview, style='TScrollbar')
    dirs_scrollbar_y.grid(row=0, column=1, sticky="ns")
    selected_dirs_listbox.config(yscrollcommand=dirs_scrollbar_y.set)

    def choose_directory_action():
        """
        ディレクトリ選択ダイアログを開き、選択されたパスをリストボックスに追加する。
        リストボックスには相対パスで表示される。
        """
        directory = filedialog.askdirectory()
        if directory:
            if directory not in selected_directories_list:
                selected_directories_list.append(directory)
                # リストボックスには相対パスで表示
                display_path = get_relative_path(directory, current_working_directory)
                selected_dirs_listbox.insert(tk.END, display_path)
                next_button.config(state="normal") # ディレクトリが追加されたらNextボタンを有効にする
            else:
                messagebox.showinfo("情報", "このディレクトリは既にリストに追加されています。")

    def remove_selected_directories():
        """
        リストボックスで選択されたディレクトリを削除する。
        """
        selected_indices = selected_dirs_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "削除するディレクトリを選択してください。")
            return
        
        # 逆順に削除することで、インデックスのずれを防ぐ
        for index in reversed(selected_indices):
            selected_directories_list.pop(index) # 内部の絶対パスリストから削除
            selected_dirs_listbox.delete(index) # 表示リストから削除
        
        if not selected_directories_list:
            next_button.config(state="disabled") # リストが空になったらNextボタンを無効にする


    button_frame = ttk.Frame(directory_selection_page, style='TFrame')
    button_frame.pack(pady=15)

    choose_dir_button = ttk.Button(
        button_frame,
        text="ディレクトリを参照",
        command=choose_directory_action,
        style='TButton',
        cursor="hand2"
    )
    choose_dir_button.pack(side="left", padx=10)

    remove_dir_button = ttk.Button(
        button_frame,
        text="選択したディレクトリを削除",
        command=remove_selected_directories,
        style='Gray.TButton',
        cursor="hand2"
    )
    remove_dir_button.pack(side="left", padx=10)

    def proceed_to_analysis():
        """
        選択されたディレクトリでファイル処理ページに進む。
        """
        if selected_directories_list:
            directory_selection_page.destroy()
            # ファイル処理ページには絶対パスのリストを渡す
            show_file_processing_page(selected_directories_list) 
        else:
            messagebox.showwarning("警告", "ディレクトリが選択されていません。")

    next_button = ttk.Button(
        directory_selection_page,
        text="次へ",
        command=proceed_to_analysis,
        style='Green.TButton',
        cursor="hand2",
        state="disabled" # 最初は無効にする
    )
    next_button.pack(pady=20)

    back_to_start_button = ttk.Button(
        directory_selection_page,
        text="スタートに戻る",
        command=lambda: [directory_selection_page.destroy(), create_start_page()],
        style='Gray.TButton',
        cursor="hand2"
    )
    back_to_start_button.pack(pady=10)

    directory_selection_page.mainloop()


def create_start_page():
    """
    アプリの開始ページを作成・表示する関数。
    """
    start_page = tk.Tk()
    start_page.title("データ分析 - حَلَّلَ")
    start_page.geometry("700x450")
    start_page.configure(bg="#F0F2F5") # 明るい背景

    # スタイル設定を適用 (Tk()インスタンス作成直後)
    configure_styles() 

    # 現在の作業ディレクトリ (CWD) の設定
    global global_current_working_directory
    current_cwd_var = tk.StringVar(value=global_current_working_directory)

    # CWD履歴の管理
    cwd_history_var = tk.StringVar()
    
    # CWD選択フレーム
    cwd_frame = ttk.Frame(start_page, style='LightGray.TFrame')
    cwd_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
    cwd_frame.columnconfigure(1, weight=1) # Entryが広がるように

    ttk.Label(cwd_frame, text="現在の作業ディレクトリ:", style='SubHeader.TLabel', background='#E0E0E0').grid(row=0, column=0, sticky="w", padx=5, pady=5)
    cwd_entry = ttk.Entry(cwd_frame, textvariable=current_cwd_var, width=60, style='TEntry')
    cwd_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

    def browse_cwd():
        new_cwd = filedialog.askdirectory()
        if new_cwd:
            current_cwd_var.set(new_cwd)
            if new_cwd not in global_cwd_history:
                global_cwd_history.append(new_cwd)
                cwd_history_combobox['values'] = global_cwd_history # コンボボックスを更新

    browse_cwd_button = ttk.Button(
        cwd_frame,
        text="参照",
        command=browse_cwd,
        style='TButton',
        cursor="hand2"
    )
    browse_cwd_button.grid(row=0, column=2, sticky="e", padx=5, pady=5)

    # CWD履歴コンボボックス
    ttk.Label(cwd_frame, text="履歴から選択:", style='SubHeader.TLabel', background='#E0E0E0').grid(row=1, column=0, sticky="w", padx=5, pady=5)
    cwd_history_combobox = ttk.Combobox(
        cwd_frame,
        textvariable=cwd_history_var,
        values=global_cwd_history,
        state="readonly", # 読み取り専用
        width=50,
        style='TCombobox'
    )
    cwd_history_combobox.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    
    def on_history_select(event):
        selected_history_path = cwd_history_var.get()
        current_cwd_var.set(selected_history_path)

    cwd_history_combobox.bind("<<ComboboxSelected>>", on_history_select)


    HALLAL_label = ttk.Label(
        start_page,
        text=ascii_art,
        font=("Courier", 18, "bold"),
        justify="center",
        foreground="#007bff", # 青色で強調
        background="#F0F2F5"
    )
    HALLAL_label.grid(row=1, column=0, sticky="nsew", padx=20, pady=10) # CWDフレームの下に配置

    hallal_arabic_label = ttk.Label(
        start_page,
        text="حَلَّلَ",
        font=("Arial", 60, "bold"),
        justify="center",
        foreground="#333333", # 濃いテキスト色
        background="#F0F2F5"
    )
    hallal_arabic_label.grid(row=3, column=0, sticky="nsew", padx=20, pady=10) # ボタンの下に配置

    def start_badaea():
        global global_current_working_directory
        global_current_working_directory = current_cwd_var.get() # CWDを更新
        if global_current_working_directory not in global_cwd_history:
            global_cwd_history.append(global_current_working_directory) # 新しいCWDを履歴に追加
        
        start_page.destroy()
        show_directory_selection_page(global_current_working_directory) # CWDを渡す

    start_button = ttk.Button(
        start_page,
        text="開始",
        command=start_badaea,
        style='Green.TButton',
        cursor="hand2"
    )
    start_button.grid(row=2, column=0, sticky="nsew", padx=50, pady=20) # HALLAL_labelとアラビア語ラベルの間に配置

    start_page.columnconfigure(0, weight=1)
    start_page.rowconfigure(0, weight=0) # CWDフレームは固定
    start_page.rowconfigure(1, weight=3) # HALLAL ASCIIアート
    start_page.rowconfigure(2, weight=1) # 開始ボタン
    start_page.rowconfigure(3, weight=2) # アラビア語ラベル

    start_page.mainloop()

# アプリケーションの開始
if __name__ == "__main__":
    create_start_page()
