import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import pandas as pd
import os
import numpy as np # For calculations
import matplotlib.pyplot as plt # For plotting
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk # For embedding plot
from matplotlib.colors import Normalize # For colormaps
from matplotlib import cm # For colormaps

# ASCIIアートの生成
# 'HALLAL' をかっこいいフォントで表示
ASCII_ART_HALLAL = """
  _   _    _    _     _    _    _
 | | | |  / \  | |   | |  / \  | |
 | |_| | / _ \ | |   | | / _ \ | |
 |  _  |/ ___ \| |___| |/ ___ \| |
 |_| |_/_/   \_\_____/_/_/   \_\_|
"""

# --- グローバル変数 ---
# ロードされたDataFrameを保存するためのグローバル辞書
loaded_dataframes = {}
# 現在表示されているDataFrameのファイルパスとシート名を追跡するための変数
current_dataframe_path = None
current_dataframe_sheet = None

# ユーザーが作成した変数を保存するためのグローバル辞書
# 例: {'var_name': {'value': pandas.Series/ndarray, 'source_file': 'filename', 'source_column': 'col_name', 'source_sheet': 'sheet_name'}}
global_variables = {}

# Treeviewのルートとなるディレクトリのパスを保持するグローバルリスト
global_root_directories = []

# 現在の作業ディレクトリ (CWD) を保持するグローバル変数
global_current_working_directory = os.path.dirname(os.path.abspath(__file__))

# CWD履歴 (メモリ上に保持)
global_cwd_history = [global_current_working_directory]


# --- Tkinter Style 設定 ---
def configure_styles():
    """
    Tkinter ttkウィジェットのスタイルを設定する。
    """
    style = ttk.Style()
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
    style.configure('Red.TButton', background='#E74C3C') # 赤色のボタン
    style.map('Red.TButton',
              background=[('active', '#C0392B')])


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
        return absolute_path

def get_absolute_path(relative_path, base_path):
    """
    相対パスを基準パスからの絶対パスに変換する。
    """
    return os.path.abspath(os.path.join(base_path, relative_path))

def generate_variable_name(column_name, file_path, sheet_name=None, depth=3):
    """
    列名、ファイル名、ディレクトリパスから変数名を生成する。
    例: .../aaa/bbb/ccc.h5 の x -> x_ccc_bbb_aaa
    """
    parts = []
    
    # 1. 列名
    parts.append(str(column_name))

    # 2. ファイル名 (拡張子なし)
    filename_without_ext = os.path.splitext(os.path.basename(file_path))[0]
    parts.append(filename_without_ext)

    # 3. シート名 (Excelの場合のみ)
    if sheet_name:
        parts.append(sheet_name)

    # 4. 親ディレクトリを遡る
    current_dir = os.path.dirname(file_path)
    for _ in range(depth):
        if current_dir and current_dir != os.path.dirname(current_dir): # ルートディレクトリに到達したら停止
            dir_name = os.path.basename(current_dir)
            if dir_name: # 空の文字列でないことを確認
                parts.append(dir_name)
            current_dir = os.path.dirname(current_dir)
        else:
            break
    
    # 不要な文字を置換または削除
    cleaned_parts = []
    for part in parts:
        cleaned_part = ''.join(char for char in part if char.isalnum() or char == '_')
        if cleaned_part:
            cleaned_parts.append(cleaned_part)

    return '_'.join(cleaned_parts)


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
    sheet_dialog.grab_set()
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
                               row_label_entry=None, col_label_entry=None, filter_expression_entry=None):
    """
    指定されたファイルをデータフレームとしてロードし、右パネルに表示する。
    """
    global current_dataframe_path, current_dataframe_sheet, loaded_dataframes
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
    if filter_expression_entry: filter_expression_entry.delete(0, tk.END) # Clear filter expression

    # デフォルトで最初の20行と全列を表示
    display_df = df.iloc[:20, :]
    if dataframe_text_widget:
        dataframe_text_widget.config(state="normal")
        dataframe_text_widget.delete("1.0", tk.END)
        dataframe_text_widget.insert(tk.END, display_df.to_string())
        dataframe_text_widget.config(state="disabled")

def display_dataframe_content(dataframe_text_widget, current_file_label_widget, 
                              start_row_entry, end_row_entry, start_col_entry, end_col_entry,
                              row_label_entry, col_label_entry, filter_expression_entry):
    """
    入力された行/列の範囲またはラベル、およびフィルタ式に基づいてデータフレームを表示する。
    """
    global current_dataframe_path, current_dataframe_sheet, loaded_dataframes
    if current_dataframe_path is None:
        messagebox.showwarning("警告", "表示するファイルが選択されていません。")
        return

    df_key = current_dataframe_path
    if current_dataframe_sheet:
        df_key += f"_{current_dataframe_sheet}"

    if df_key not in loaded_dataframes:
        messagebox.showerror("エラー", "データフレームがロードされていません。")
        return

    df = loaded_dataframes[df_key].copy() # フィルタリングのためにコピーを作成

    # フィルタ式の適用
    filter_expr = filter_expression_entry.get().strip()
    if filter_expr:
        try:
            df = df.query(filter_expr)
            if df.empty:
                messagebox.showinfo("情報", "フィルタリングの結果、データがありません。")
                dataframe_text_widget.config(state="normal")
                dataframe_text_widget.delete("1.0", tk.END)
                dataframe_text_widget.insert(tk.END, "フィルタリングの結果、データがありません。")
                dataframe_text_widget.config(state="disabled")
                return
        except Exception as e:
            messagebox.showerror("フィルタエラー", f"フィルタ式の適用中にエラーが発生しました: {e}\n式を確認してください。")
            return

    try:
        # 行の指定方法を決定
        row_selection = slice(None) # デフォルトは全行
        row_label_input = row_label_entry.get().strip()
        start_row_idx_input = start_row_entry.get().strip()
        end_row_idx_input = end_row_entry.get().strip()

        if row_label_input:
            if row_label_input in df.index:
                row_selection = row_label_input
            else:
                messagebox.showwarning("警告", f"指定された行ラベル '{row_label_input}' は見つかりませんでした。")
                return
        elif start_row_idx_input or end_row_idx_input:
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
            if col_label_input in df.columns:
                col_selection = col_label_input
            else:
                messagebox.showwarning("警告", f"指定された列ラベル '{col_label_input}' は見つかりませんでした。")
                return
        elif start_col_idx_input or end_col_idx_input:
            start_col = int(start_col_idx_input) if start_col_idx_input else 0
            end_col = int(end_col_idx_input) if end_col_idx_input else None
            col_selection = slice(start_col, end_col)

    except ValueError as e:
        messagebox.showerror("エラー", f"入力値が無効です: {e}\n行/列は数値インデックスまたはラベルを入力してください。")
        return
    except KeyError as e:
        messagebox.showerror("エラー", f"指定されたラベルが見つかりません: {e}")
        return
    except Exception as e:
        messagebox.showerror("エラー", f"データフレームの表示中に予期せぬエラーが発生しました: {e}")
        return

    display_df = None
    try:
        if isinstance(row_selection, slice) and isinstance(col_selection, slice):
            display_df = df.iloc[row_selection, col_selection]
        elif isinstance(row_selection, str) and isinstance(col_selection, slice):
            display_df = df.loc[[row_selection], col_selection]
        elif isinstance(row_selection, slice) and isinstance(col_selection, str):
            display_df = df.loc[row_selection, [col_selection]]
        elif isinstance(row_selection, str) and isinstance(col_selection, str):
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

def embed_variables_dialog(parent_window, df_to_embed, file_path, sheet_name, variable_listbox_widget):
    """
    データフレームの列を変数に組み込むためのカスタムダイアログを表示する。
    """
    global global_variables

    embed_dialog = tk.Toplevel(parent_window)
    embed_dialog.title(f"変数に組み込む: {os.path.basename(file_path)}")
    embed_dialog.geometry("600x500")
    embed_dialog.grab_set()
    embed_dialog.transient(parent_window)
    embed_dialog.configure(bg="#F0F2F5")

    embed_dialog.columnconfigure(0, weight=1)
    embed_dialog.rowconfigure(1, weight=1) # Treeviewが広がるように

    ttk.Label(embed_dialog, text="組み込む変数を選択・編集:", style='SubHeader.TLabel', background="#F0F2F5").pack(pady=10)

    # Treeviewで列候補を表示
    columns = ('include', 'original_column', 'default_name', 'new_name')
    var_tree = ttk.Treeview(embed_dialog, columns=columns, show='headings', selectmode='none')
    var_tree.pack(fill="both", expand=True, padx=10, pady=5)

    var_tree.heading('include', text='組み込む')
    var_tree.heading('original_column', text='元の列')
    var_tree.heading('default_name', text='デフォルト名')
    var_tree.heading('new_name', text='新しい変数名')

    var_tree.column('include', width=60, anchor='center')
    var_tree.column('original_column', width=150)
    var_tree.column('default_name', width=150)
    var_tree.column('new_name', width=150)

    # チェックボックスの状態を管理する辞書
    checkbox_vars = {}
    # 新しい変数名のEntryウィジェットを管理する辞書
    entry_widgets = {}

    def toggle_checkbox(item_id):
        current_value = checkbox_vars[item_id].get()
        checkbox_vars[item_id].set(not current_value)
        # Treeviewの表示を更新 (チェックマークの有無)
        var_tree.item(item_id, values=(
            "✅" if checkbox_vars[item_id].get() else "",
            var_tree.item(item_id, 'values')[1],
            var_tree.item(item_id, 'values')[2],
            entry_widgets[item_id].get()
        ))

    def on_entry_change(item_id, new_name_var):
        # Entryの変更がTreeviewの表示に反映されるようにする
        var_tree.item(item_id, values=(
            "✅" if checkbox_vars[item_id].get() else "",
            var_tree.item(item_id, 'values')[1],
            var_tree.item(item_id, 'values')[2],
            new_name_var.get()
        ))

    for col_name in df_to_embed.columns:
        default_var_name = generate_variable_name(col_name, file_path, sheet_name)
        
        item_id = var_tree.insert('', 'end', values=("✅", col_name, default_var_name, default_var_name))
        
        # チェックボックス用のBooleanVar
        checkbox_vars[item_id] = tk.BooleanVar(value=True)
        # Entry用のStringVar
        new_name_var = tk.StringVar(value=default_var_name)
        new_name_var.trace_add("write", lambda name, index, mode, item_id=item_id, new_name_var=new_name_var: on_entry_change(item_id, new_name_var))

        # EntryウィジェットをTreeviewのセルに配置
        entry_widgets[item_id] = ttk.Entry(var_tree, textvariable=new_name_var, style='TEntry')
        var_tree.window_create(item_id, column='new_name', window=entry_widgets[item_id])

        # チェックボックス用のウィジェットをTreeviewのセルに配置
        cb = ttk.Checkbutton(var_tree, variable=checkbox_vars[item_id], command=lambda item_id=item_id: toggle_checkbox(item_id), style='TCheckbutton')
        var_tree.window_create(item_id, column='include', window=cb)


    def process_selected_variables():
#        nonlocal global_variables # グローバル変数を更新
        processed_count = 0
        for item_id in var_tree.get_children():
            if checkbox_vars[item_id].get(): # チェックボックスがONの場合
                original_col = var_tree.item(item_id, 'values')[1]
                var_name = entry_widgets[item_id].get().strip()

                if not var_name:
                    messagebox.showwarning("警告", f"'{original_col}' の変数名が空です。スキップします。")
                    continue

                if var_name in global_variables:
                    if not messagebox.askyesno("警告", f"変数名 '{var_name}' は既に存在します。上書きしますか？"):
                        continue

                global_variables[var_name] = {
                    'value': df_to_embed[original_col],
                    'source_file': os.path.basename(file_path),
                    'source_sheet': sheet_name,
                    'source_column': original_col
                }
                processed_count += 1
        
        if processed_count > 0:
            messagebox.showinfo("情報", f"{processed_count}個の変数を組み込みました。")
            update_variable_list(variable_listbox_widget) # メイン画面の変数リストを更新
            embed_dialog.destroy()
        else:
            messagebox.showinfo("情報", "選択された変数は組み込まれませんでした。")

    embed_button = ttk.Button(
        embed_dialog,
        text="変数に組み込む",
        command=process_selected_variables,
        style='Green.TButton',
        cursor="hand2"
    )
    embed_button.pack(pady=10)

    embed_dialog.wait_window(embed_dialog)


def update_variable_list(variable_listbox_widget):
    """グローバル変数リストを更新して表示する。"""
    variable_listbox_widget.delete(0, tk.END)
    for var_name, var_info in global_variables.items():
        source_info = f"({var_info['source_file']}"
        if var_info['source_sheet']:
            source_info += f" - {var_info['source_sheet']}"
        source_info += f", Col: {var_info['source_column']})"
        variable_listbox_widget.insert(tk.END, f"{var_name} {source_info}")

def embed_multiple_variables_from_selection(parent_window, file_tree_widget, 
                                            start_row_entry, end_row_entry, start_col_entry, end_col_entry,
                                            row_label_entry, col_label_entry, variable_listbox_widget):
    """
    Treeviewで選択された複数のファイルから、指定範囲のデータを変数に一括で組み込む。
    """
    global loaded_dataframes, global_variables

    selected_item_ids = file_tree_widget.selection()
    if not selected_item_ids:
        messagebox.showwarning("警告", "変数を組み込むファイルを選択してください。")
        return

    # フィルタリング/スライス条件を事前に取得
    row_label_input = row_label_entry.get().strip()
    start_row_idx_input = start_row_entry.get().strip()
    end_row_idx_input = end_row_entry.get().strip()
    col_label_input = col_label_entry.get().strip()
    start_col_idx_input = start_col_entry.get().strip()
    end_col_idx_input = end_col_entry.get().strip()

    processed_files_count = 0
    processed_vars_count = 0

    for item_id in selected_item_ids:
        item_tags = file_tree_widget.item(item_id, "tags")
        if "file" in item_tags:
            file_path = file_tree_widget.item(item_id, "values")[0]
            sheet_name = None
            if file_path.lower().endswith(('.xlsx', '.xls')):
                sheet_name = prompt_for_excel_sheet(parent_window, file_path)
                if sheet_name is None:
                    continue # ユーザーがシート選択をキャンセルした場合
            
            df_key = file_path
            if sheet_name:
                df_key += f"_{sheet_name}"

            df = loaded_dataframes.get(df_key)
            if df is None:
                try:
                    if file_path.lower().endswith('.csv'):
                        df = pd.read_csv(file_path)
                    elif file_path.lower().endswith(('.h5', '.hdf')):
                        df = pd.read_hdf(file_path)
                    elif file_path.lower().endswith(('.xlsx', '.xls')):
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                    loaded_dataframes[df_key] = df
                except Exception as e:
                    messagebox.showerror("エラー", f"ファイル '{os.path.basename(file_path)}' の読み込み中にエラーが発生しました: {e}")
                    continue

            try:
                # 行の選択を適用
                row_selection = slice(None)
                if row_label_input: row_selection = row_label_input
                elif start_row_idx_input or end_row_idx_input:
                    start_row = int(start_row_idx_input) if start_row_idx_input else 0
                    end_row = int(end_row_idx_input) if end_row_idx_input else None
                    row_selection = slice(start_row, end_row)

                # 列の選択を適用
                col_selection = slice(None)
                if col_label_input: col_selection = col_label_input
                elif start_col_idx_input or end_col_idx_input:
                    start_col = int(start_col_idx_input) if start_col_idx_input else 0
                    end_col = int(end_col_idx_input) if end_col_idx_input else None
                    col_selection = slice(start_col, end_col)

                # 実際にスライスを適用
                if isinstance(row_selection, slice) and isinstance(col_selection, slice):
                    df_slice = df.iloc[row_selection, col_selection]
                elif isinstance(row_selection, str) and isinstance(col_selection, slice):
                    df_slice = df.loc[[row_selection], col_selection]
                elif isinstance(row_selection, slice) and isinstance(col_selection, str):
                    df_slice = df.loc[row_selection, [col_selection]]
                elif isinstance(row_selection, str) and isinstance(col_selection, str):
                    df_slice = df.loc[[row_selection], [col_selection]]
                else:
                    messagebox.showerror("エラー", "行と列の選択の組み合わせが不正です。")
                    continue

                if df_slice.empty:
                    messagebox.showwarning("警告", f"ファイル '{os.path.basename(file_path)}' の指定範囲でデータが見つかりませんでした。")
                    continue

                source_filename_base = os.path.splitext(os.path.basename(file_path))[0]

                # 各列を変数に組み込むためのダイアログ (一括処理のため、ここでは自動命名のみ)
                # ユーザーに確認ダイアログを表示し、一括で組み込むか尋ねる
                if not messagebox.askyesno("一括組み込み確認", 
                                            f"ファイル '{os.path.basename(file_path)}' から {len(df_slice.columns)} 列を変数に組み込みますか？\n（変数名は自動生成されます）"):
                    continue

                for col_name in df_slice.columns:
                    var_name = generate_variable_name(col_name, file_path, sheet_name)
                    
                    if var_name in global_variables:
                        if not messagebox.askyesno("警告", f"変数名 '{var_name}' は既に存在します。上書きしますか？"):
                            continue

                    global_variables[var_name] = {
                        'value': df_slice[col_name],
                        'source_file': os.path.basename(file_path),
                        'source_sheet': sheet_name,
                        'source_column': col_name
                    }
                    processed_vars_count += 1
                processed_files_count += 1

            except Exception as e:
                messagebox.showerror("エラー", f"ファイル '{os.path.basename(file_path)}' のデータ処理中にエラーが発生しました: {e}")
                continue
    
    if processed_vars_count > 0:
        update_variable_list(variable_listbox_widget)
        messagebox.showinfo("情報", f"{processed_files_count}個のファイルから合計{processed_vars_count}個の変数を組み込みました。")
    else:
        messagebox.showinfo("情報", "選択されたファイルから変数は組み込まれませんでした。")

# --- プロット機能 ---
# プロットのレイヤーを管理するリスト
# 各要素は辞書で、プロットタイプ、変数、スタイル設定、axes情報などを含む
plot_layers = []
current_figure = None
current_canvas = None
current_toolbar = None

def show_plot_page(parent_window):
    """
    プロット機能を提供する新しいToplevelウィンドウを表示する。
    """
    plot_window = tk.Toplevel(parent_window)
    plot_window.title("Plotting Tool")
    plot_window.geometry("1200x800")
    plot_window.configure(bg="#F0F2F5")
    plot_window.grab_set()

    plot_window.columnconfigure(0, weight=1) # コントロールパネル
    plot_window.columnconfigure(1, weight=3) # プロット表示エリア
    plot_window.rowconfigure(0, weight=1)

    # --- コントロールパネル (左側) ---
    control_frame = ttk.Frame(plot_window, style='White.TFrame')
    control_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    control_frame.columnconfigure(0, weight=1)
    control_frame.rowconfigure(0, weight=0) # Plot Settings Header
    control_frame.rowconfigure(1, weight=1) # Plot Type & Variable Selection
    control_frame.rowconfigure(2, weight=1) # Plot List
    control_frame.rowconfigure(3, weight=0) # Generate/Update Button

    ttk.Label(control_frame, text="Plot Settings", style='Header.TLabel', background="#FFFFFF").pack(pady=10)

    # プロット追加用コントロール
    add_plot_controls_frame = ttk.Frame(control_frame, style='White.TFrame')
    add_plot_controls_frame.pack(fill="x", padx=5, pady=5)
    add_plot_controls_frame.columnconfigure(0, weight=1)

    ttk.Label(add_plot_controls_frame, text="Plot Type:", style='SubHeader.TLabel', background="#FFFFFF").pack(pady=5)
    plot_type_var = tk.StringVar(value="scatter (2D)")
    plot_types = [
        "scatter (2D)", "plot (2D)", "hist (2D)", "contour (2D)", "contourf (2D)", 
        "fill_between (2D)", "tricontourf (2D)", "streamplot (2D)", "quiver (2D)",
        "scatter (3D)", "plot (3D)", "quiver (3D)"
    ]
    plot_type_combobox = ttk.Combobox(add_plot_controls_frame, textvariable=plot_type_var, values=plot_types, state="readonly", style='TCombobox')
    plot_type_combobox.pack(pady=5)

    ttk.Label(add_plot_controls_frame, text="Variables:", style='SubHeader.TLabel', background="#FFFFFF").pack(pady=5)
    
    # すべての変数コンボボックスとラベルを一度作成
    # X Variable
    x_var_label = ttk.Label(add_plot_controls_frame, text="X Variable:", style='TLabel', background="#FFFFFF")
    x_var_combobox = ttk.Combobox(add_plot_controls_frame, values=list(global_variables.keys()), state="readonly", style='TCombobox')
    # Y Variable
    y_var_label = ttk.Label(add_plot_controls_frame, text="Y Variable:", style='TLabel', background="#FFFFFF")
    y_var_combobox = ttk.Combobox(add_plot_controls_frame, values=list(global_variables.keys()), state="readonly", style='TCombobox')
    # Z Variable (for 3D)
    z_var_label = ttk.Label(add_plot_controls_frame, text="Z Variable (for 3D):", style='TLabel', background="#FFFFFF")
    z_var_combobox = ttk.Combobox(add_plot_controls_frame, values=list(global_variables.keys()), state="readonly", style='TCombobox')
    # U Variable (for quiver/streamplot)
    u_var_label = ttk.Label(add_plot_controls_frame, text="U Variable:", style='TLabel', background="#FFFFFF")
    u_var_combobox = ttk.Combobox(add_plot_controls_frame, values=list(global_variables.keys()), state="readonly", style='TCombobox')
    # V Variable (for quiver/streamplot)
    v_var_label = ttk.Label(add_plot_controls_frame, text="V Variable:", style='TLabel', background="#FFFFFF")
    v_var_combobox = ttk.Combobox(add_plot_controls_frame, values=list(global_variables.keys()), state="readonly", style='TCombobox')


    # Dynamic variable input based on plot type
    def update_variable_inputs(*args):
        # すべてのウィジェットを非表示にする
        for widget in [x_var_label, x_var_combobox, y_var_label, y_var_combobox, 
                       z_var_label, z_var_combobox, u_var_label, u_var_combobox, 
                       v_var_label, v_var_combobox]:
            widget.pack_forget()

        ptype = plot_type_var.get()
        if ptype in ["scatter (2D)", "plot (2D)", "fill_between (2D)"]:
            x_var_label.pack()
            x_var_combobox.pack()
            y_var_label.pack()
            y_var_combobox.pack()
        elif ptype == "hist (2D)":
            x_var_label.config(text="Data Variable:")
            x_var_label.pack()
            x_var_combobox.pack()
        elif ptype in ["contour (2D)", "contourf (2D)", "tricontourf (2D)"]:
            x_var_label.pack()
            x_var_combobox.pack()
            y_var_label.pack()
            y_var_combobox.pack()
            z_var_label.config(text="Z Variable:")
            z_var_label.pack()
            z_var_combobox.pack()
        elif ptype == "streamplot (2D)" or ptype == "quiver (2D)":
            x_var_label.pack()
            x_var_combobox.pack()
            y_var_label.pack()
            y_var_combobox.pack()
            u_var_label.pack()
            u_var_combobox.pack()
            v_var_label.pack()
            v_var_combobox.pack()
        elif ptype in ["scatter (3D)", "plot (3D)"]:
            x_var_label.pack()
            x_var_combobox.pack()
            y_var_label.pack()
            y_var_combobox.pack()
            z_var_label.pack()
            z_var_combobox.pack()
        elif ptype == "quiver (3D)":
            x_var_label.pack()
            x_var_combobox.pack()
            y_var_label.pack()
            y_var_combobox.pack()
            z_var_label.pack()
            z_var_combobox.pack()
            u_var_label.pack()
            u_var_combobox.pack()
            v_var_label.pack()
            v_var_combobox.pack()
            
        # Update combobox values in case global_variables changed
        current_var_keys = list(global_variables.keys())
        x_var_combobox['values'] = current_var_keys
        y_var_combobox['values'] = current_var_keys
        z_var_combobox['values'] = current_var_keys
        u_var_combobox['values'] = current_var_keys
        v_var_combobox['values'] = current_var_keys

    plot_type_var.trace_add("write", update_variable_inputs)
    update_variable_inputs() # 初期表示

    add_plot_button = ttk.Button(
        add_plot_controls_frame,
        text="プロット追加",
        command=lambda: add_plot_layer(plot_type_var.get(), x_var_combobox.get(), y_var_combobox.get(), 
                                       z_var_combobox.get(), u_var_combobox.get(), v_var_combobox.get()),
        style='Green.TButton',
        cursor="hand2"
    )
    add_plot_button.pack(pady=10)

    # プロットリスト
    ttk.Label(control_frame, text="Current Plots:", style='SubHeader.TLabel', background="#FFFFFF").pack(pady=5)
    plot_list_tree = ttk.Treeview(control_frame, columns=('type', 'vars'), show='headings')
    plot_list_tree.heading('type', text='タイプ')
    plot_list_tree.heading('vars', text='変数')
    plot_list_tree.column('type', width=100)
    plot_list_tree.column('vars', width=200)
    plot_list_tree.pack(fill="both", expand=True, padx=5, pady=5)

    def remove_plot_layer(item_id):
        idx_to_remove = -1
        for i, layer in enumerate(plot_layers):
            if layer['id'] == item_id:
                idx_to_remove = i
                break
        if idx_to_remove != -1:
            plot_layers.pop(idx_to_remove)
            plot_list_tree.delete(item_id)
            redraw_plot_figure() # プロットを再描画

    # プロットリストのアイテムがクリックされたら削除ボタンを表示
    def on_plot_list_select(event):
        selected_item_id = plot_list_tree.focus()
        if selected_item_id:
            if messagebox.askyesno("確認", "このプロットを削除しますか？"):
                remove_plot_layer(selected_item_id)
    plot_list_tree.bind("<<TreeviewSelect>>", on_plot_list_select)


    # プロット更新ボタン
    update_plot_button = ttk.Button(
        control_frame,
        text="プロットを更新",
        command=lambda: redraw_plot_figure(),
        style='TButton',
        cursor="hand2"
    )
    update_plot_button.pack(pady=10)

    # --- プロット表示エリア (右側) ---
    plot_area_frame = ttk.Frame(plot_window, style='White.TFrame')
    plot_area_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
    plot_area_frame.columnconfigure(0, weight=1)
    plot_area_frame.rowconfigure(0, weight=1)

    fig, ax = plt.subplots(figsize=(8, 6), dpi=100) # 初期FigureとAxes
    canvas = FigureCanvasTkAgg(fig, master=plot_area_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    toolbar = NavigationToolbar2Tk(canvas, plot_area_frame)
    toolbar.update()
    canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    global current_figure, current_canvas, current_toolbar
    current_figure = fig
    current_canvas = canvas
    current_toolbar = toolbar

    def add_plot_layer(plot_type, x_var_name, y_var_name, z_var_name=None, u_var_name=None, v_var_name=None):
        """
        プロットレイヤーを追加し、リストを更新してプロットを再描画する。
        """
        layer_id = f"plot_{len(plot_layers)}_{plot_type.replace(' ', '_')}_{x_var_name}_{y_var_name}"
        # 重複チェック (簡易版)
        for layer in plot_layers:
            if (layer['type'] == plot_type and layer['x_var'] == x_var_name and 
                layer['y_var'] == y_var_name and layer['z_var'] == z_var_name and
                layer['u_var'] == u_var_name and layer['v_var'] == v_var_name):
                messagebox.showwarning("警告", "同じプロットが既に存在します。")
                return

        plot_info = {
            'id': layer_id,
            'type': plot_type,
            'x_var': x_var_name,
            'y_var': y_var_name,
            'z_var': z_var_name,
            'u_var': u_var_name,
            'v_var': v_var_name,
            'style': {} # ここに将来的にスタイルオプションを追加
        }
        plot_layers.append(plot_info)
        
        var_display = f"X:{x_var_name}, Y:{y_var_name}"
        if z_var_name and '3D' in plot_type: var_display += f", Z:{z_var_name}"
        if u_var_name and v_var_name: var_display += f", U:{u_var_name}, V:{v_var_name}"
        
        plot_list_tree.insert('', 'end', iid=layer_id, values=(plot_type, var_display))
        redraw_plot_figure()


    def redraw_plot_figure():
        """
        現在のplot_layersリストに基づいてプロット全体を再描画する。
        """
        global current_figure, current_canvas, current_toolbar
        
        if current_figure:
            plt.close(current_figure) # 既存のFigureを閉じる

        fig = plt.Figure(figsize=(8, 6), dpi=100)
        
        # Axesの作成と管理
        ax = None
        has_3d_plot = any('3D' in layer['type'] for layer in plot_layers)
        if has_3d_plot:
            ax = fig.add_subplot(111, projection='3d')
        else:
            ax = fig.add_subplot(111)
        
        # 軸ラベル、タイトル、目盛りなどの設定は、各プロットレイヤーの描画後にまとめて行うか、
        # より詳細なGUIコントロールを追加する必要がある。
        # ここでは、各レイヤーが軸ラベルを設定し、最後のレイヤーの設定が適用される。

        for layer in plot_layers:
            ptype = layer['type']
            x_var_name = layer['x_var']
            y_var_name = layer['y_var']
            z_var_name = layer['z_var']
            u_var_name = layer['u_var']
            v_var_name = layer['v_var']

            try:
                x_data = global_variables.get(x_var_name, {}).get('value')
                y_data = global_variables.get(y_var_name, {}).get('value')
                z_data = global_variables.get(z_var_name, {}).get('value')
                u_data = global_variables.get(u_var_name, {}).get('value')
                v_data = global_variables.get(v_var_name, {}).get('value')

                # データがSeriesの場合、NumPy配列に変換 (特にcontourfなどで必要)
                if isinstance(x_data, pd.Series): x_data = x_data.to_numpy()
                if isinstance(y_data, pd.Series): y_data = y_data.to_numpy()
                if isinstance(z_data, pd.Series): z_data = z_data.to_numpy()
                if isinstance(u_data, pd.Series): u_data = u_data.to_numpy()
                if isinstance(v_data, pd.Series): v_data = v_data.to_numpy()

                # データがNoneの場合はスキップ
                if x_data is None or y_data is None:
                    messagebox.showwarning("警告", f"プロット '{layer['id']}' のデータが不足しています。")
                    continue

                # プロットタイプに応じた描画
                if ptype == "scatter (2D)":
                    ax.scatter(x_data, y_data, label=layer['id'])
                elif ptype == "plot (2D)":
                    ax.plot(x_data, y_data, label=layer['id'])
                elif ptype == "hist (2D)":
                    ax.hist(x_data, bins=30, label=layer['id'])
                elif ptype == "fill_between (2D)":
                    ax.fill_between(x_data, y_data, color='skyblue', alpha=0.4, label=layer['id'])
                elif ptype == "contour (2D)":
                    if z_data is None: raise ValueError("Z data required for contour.")
                    X, Y = np.meshgrid(np.unique(x_data), np.unique(y_data))
                    Z = plt.mlab.griddata(x_data, y_data, z_data, X, Y, interp='linear')
                    ax.contour(X, Y, Z, label=layer['id'])
                elif ptype == "contourf (2D)":
                    if z_data is None: raise ValueError("Z data required for contourf.")
                    X, Y = np.meshgrid(np.unique(x_data), np.unique(y_data))
                    Z = plt.mlab.griddata(x_data, y_data, z_data, X, Y, interp='linear')
                    ax.contourf(X, Y, Z, label=layer['id'])
                elif ptype == "tricontourf (2D)":
                    if z_data is None: raise ValueError("Z data required for tricontourf.")
                    ax.tricontourf(x_data, y_data, z_data, label=layer['id'])
                elif ptype == "streamplot (2D)":
                    if u_data is None or v_data is None: raise ValueError("U and V data required for streamplot.")
                    X, Y = np.meshgrid(np.unique(x_data), np.unique(y_data))
                    U = plt.mlab.griddata(x_data, y_data, u_data, X, Y, interp='linear')
                    V = plt.mlab.griddata(x_data, y_data, v_data, X, Y, interp='linear')
                    ax.streamplot(X, Y, U, V, label=layer['id'])
                elif ptype == "quiver (2D)":
                    if u_data is None or v_data is None: raise ValueError("U and V data required for quiver (2D).")
                    ax.quiver(x_data, y_data, u_data, v_data, label=layer['id'])
                elif ptype == "scatter (3D)":
                    if z_data is None: raise ValueError("Z data required for 3D scatter.")
                    ax.scatter(x_data, y_data, z_data, label=layer['id'])
                elif ptype == "plot (3D)":
                    if z_data is None: raise ValueError("Z data required for 3D plot.")
                    ax.plot(x_data, y_data, z_data, label=layer['id'])
                elif ptype == "quiver (3D)":
                    if z_data is None or u_data is None or v_data is None: raise ValueError("Z, U, V data required for 3D quiver.")
                    ax.quiver(x_data, y_data, z_data, u_data, v_data, np.zeros_like(u_data), label=layer['id']) # Z direction is often 0 for 3D quiver unless specified
                else:
                    messagebox.showwarning("警告", f"プロットタイプ '{ptype}' は未実装です。")
                    continue

            except KeyError as e:
                messagebox.showerror("エラー", f"プロット '{layer['id']}' の変数 '{e}' が見つからないか、データがありません。")
                continue
            except ValueError as e:
                messagebox.showerror("エラー", f"プロット '{layer['id']}' のデータ形式が不正です: {e}")
                continue
            except Exception as e:
                messagebox.showerror("プロットエラー", f"プロット '{layer['id']}' の生成中にエラーが発生しました: {e}")
                continue
        
        # 凡例の表示
        if plot_layers: # プロットがある場合のみ凡例を表示
            ax.legend()
        
        # Figureを更新
        current_canvas.figure = fig # 既存のCanvasに新しいFigureをセット
        current_canvas.draw() # 再描画

        if current_toolbar:
            current_toolbar.destroy() # 既存のツールバーを破棄
        current_toolbar = NavigationToolbar2Tk(current_canvas, plot_area_frame)
        current_toolbar.update()
        current_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        current_figure = fig # グローバル変数に新しいFigureを保存


def show_calculation_page(parent_window):
    """
    計算機能を提供する新しいToplevelウィンドウを表示する。
    """
    calc_window = tk.Toplevel(parent_window)
    calc_window.title("Calculation Tool")
    calc_window.geometry("800x600")
    calc_window.configure(bg="#F0F2F5")
    calc_window.grab_set()

    calc_window.columnconfigure(0, weight=1)
    calc_window.rowconfigure(0, weight=1) # 変数リスト
    calc_window.rowconfigure(1, weight=2) # コード入力
    calc_window.rowconfigure(2, weight=0) # ボタン
    calc_window.rowconfigure(3, weight=1) # 出力

    # --- 変数リスト (左上) ---
    var_list_frame = ttk.Frame(calc_window, style='White.TFrame')
    var_list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    var_list_frame.columnconfigure(0, weight=1)
    var_list_frame.rowconfigure(0, weight=1)
    
    ttk.Label(var_list_frame, text="利用可能な変数:", style='SubHeader.TLabel', background="#FFFFFF").pack(pady=5)
    calc_var_listbox = tk.Listbox(var_list_frame, bg="#FFFFFF", fg="#333333", font=("Courier", 10), relief="flat")
    calc_var_listbox.pack(fill="both", expand=True, padx=5, pady=5)

    def refresh_calc_var_list():
        calc_var_listbox.delete(0, tk.END)
        for var_name, var_info in global_variables.items():
            source_info = f"({var_info['source_column']} from {var_info['source_file']})"
            calc_var_listbox.insert(tk.END, f"{var_name} {source_info}")
    
    refresh_calc_var_list()

    # --- コード入力エリア (中央) ---
    code_input_frame = ttk.Frame(calc_window, style='White.TFrame')
    code_input_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
    code_input_frame.columnconfigure(0, weight=1)
    code_input_frame.rowconfigure(0, weight=1)

    ttk.Label(code_input_frame, text="Pythonコード:", style='SubHeader.TLabel', background="#FFFFFF").pack(pady=5)
    code_text = tk.Text(code_input_frame, wrap="word", bg="#F8F8F8", fg="#333333", font=("Courier", 10), relief="solid", bd=1)
    code_text.pack(fill="both", expand=True, padx=5, pady=5)
    code_text.insert(tk.END, "# 例: numpy (np), pandas (pd) が利用可能です\n# result = my_variable_from_csv + np.array([1,2,3])\n# print(result.head() if hasattr(result, 'head') else result)\n")

    # --- 実行ボタンと変数追加ボタン (中央下) ---
    button_frame = ttk.Frame(calc_window, style='TFrame')
    button_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=1)

    # exec_globalsとexec_localsを保持するためのリスト (スコープをまたいでアクセス可能にするため)
    # [globals_dict, locals_dict]
    last_exec_scope = [{}, {}] 

    output_text = None # 前方参照のためにNoneで初期化

    def execute_code():
        nonlocal last_exec_scope # outer scopeのlast_exec_scopeを変更可能にする
        code = code_text.get("1.0", tk.END).strip()
        if not code:
            return

        # 実行環境を準備
        exec_globals = {
            'pd': pd,
            'np': np,
            'messagebox': messagebox,
            'os': os,
            'loaded_dataframes': loaded_dataframes,
            **{name: var_info['value'] for name, var_info in global_variables.items()} # ユーザー変数を展開
        }
        exec_locals = {} # ローカル変数は空

        # 出力のリダイレクト
        old_stdout = os.sys.stdout
        redirected_output_buffer = []
        class StdoutRedirector:
            def write(self, s):
                redirected_output_buffer.append(s)
                output_text.config(state="normal")
                output_text.insert(tk.END, s)
                output_text.see(tk.END)
                output_text.config(state="disabled")

            def flush(self): # flushメソッドも定義
                pass

        os.sys.stdout = StdoutRedirector()

        output_text.config(state="normal")
        output_text.delete("1.0", tk.END)
        redirected_output_buffer.clear()

        try:
            exec(code, exec_globals, exec_locals)
            last_exec_scope[0] = exec_globals # 実行後のスコープを保存
            last_exec_scope[1] = exec_locals
        except Exception as e:
            output_text.config(state="normal")
            output_text.insert(tk.END, f"エラー:\n{e}\n")
            output_text.config(state="disabled")
        finally:
            os.sys.stdout = old_stdout # stdoutを元に戻す

    def add_calculated_variable():
        var_name = simpledialog.askstring("新しい変数名", "計算結果の変数名を指定してください。\n（例: コードで `my_result = ...` とした場合、ここに `my_result` と入力）", parent=calc_window)
        if var_name:
            # last_exec_scopeから結果を取得
            val = None
            if var_name in last_exec_scope[1]: # 実行されたコードのローカル変数にあるか優先的にチェック
                val = last_exec_scope[1][var_name]
            elif var_name in last_exec_scope[0]: # グローバル変数にあるかチェック
                val = last_exec_scope[0][var_name]
            else:
                messagebox.showwarning("警告", f"指定された変数名 '{var_name}' は、最後に実行されたコードのスコープ内で見つかりませんでした。")
                return

            # pandas Series/DataFrame/numpy arrayを想定
            if not hasattr(val, '__len__') and not isinstance(val, (int, float, bool)): # スカラー値も許容するが、ここでは配列/Seriesを想定
                messagebox.showwarning("警告", "指定された変数は配列、Series、またはDataFrameではありません。")
                return

            if var_name in global_variables:
                if not messagebox.askyesno("警告", f"変数名 '{var_name}' は既に存在します。上書きしますか？"):
                    return
            
            # 変数情報を保存
            global_variables[var_name] = {
                'value': val,
                'source_file': 'Calculation',
                'source_sheet': None,
                'source_column': var_name
            }
            # 計算後、ファイル処理ページに戻ったときに変数リストを再描画する
            # メッセージボックスで通知し、ユーザーが戻ったときに更新されることを期待
            messagebox.showinfo("成功", f"変数 '{var_name}' が追加されました。ファイル処理ページの変数リストを更新してください。")
            
            # 実際には、以下のように呼び出し元（file_processing_page）のvariable_listboxを更新する
            # parent_window.event_generate("<<RefreshVariableList>>")
            # そして、file_processing_pageでこのイベントをバインドする
            # 今回は、show_file_processing_pageに戻ったときに自動更新されるようにする

    execute_button = ttk.Button(
        button_frame,
        text="実行",
        command=execute_code,
        style='Green.TButton',
        cursor="hand2"
    )
    execute_button.pack(side="left", fill="x", expand=True, padx=5)

    add_var_button = ttk.Button(
        button_frame,
        text="新しい変数として追加",
        command=add_calculated_variable,
        style='TButton',
        cursor="hand2"
    )
    add_var_button.pack(side="left", fill="x", expand=True, padx=5)

    # --- 出力エリア (右下) ---
    output_frame = ttk.Frame(calc_window, style='White.TFrame')
    output_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)
    output_frame.columnconfigure(0, weight=1)
    output_frame.rowconfigure(0, weight=1)

    ttk.Label(output_frame, text="出力:", style='SubHeader.TLabel', background="#FFFFFF").pack(pady=5)
    output_text = tk.Text(output_frame, wrap="word", bg="#F8F8F8", fg="#333333", font=("Courier", 10), relief="solid", bd=1, state="disabled")
    output_text.pack(fill="both", expand=True, padx=5, pady=5)


def show_file_processing_page(initial_directory_paths=None):
    """
    ファイルリストとデータフレーム表示機能を持つページを表示する関数。
    複数の初期ディレクトリパスをリストとして受け取る。
    """
    file_processing_page = tk.Tk()
    file_processing_page.title(f"データ分析 - ファイル処理")
    file_processing_page.geometry("1200x800")
    file_processing_page.configure(bg="#F0F2F5")

    configure_styles()

    file_processing_page.columnconfigure(0, weight=1) # 左パネル（ファイルツリーと変数リスト）
    file_processing_page.columnconfigure(1, weight=3) # 右パネル（データフレーム表示と操作）
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


    # --- 左パネル: ファイルツリーと変数リスト ---
    left_panel_frame = ttk.Frame(file_processing_page, style='White.TFrame')
    left_panel_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
    left_panel_frame.columnconfigure(0, weight=1)
    left_panel_frame.rowconfigure(0, weight=3) # ファイルツリーに多めにスペース
    left_panel_frame.rowconfigure(1, weight=1) # 変数リストに少なめにスペース

    ttk.Label(left_panel_frame, text="ファイルツリー", style='Header.TLabel', background="#FFFFFF").grid(row=0, column=0, sticky="nw", padx=5, pady=5)
    file_tree = ttk.Treeview(left_panel_frame, style='Treeview')
    file_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=(30,5))

    tree_scrollbar_y = ttk.Scrollbar(left_panel_frame, orient="vertical", command=file_tree.yview, style='TScrollbar')
    tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
    file_tree.configure(yscrollcommand=tree_scrollbar_y.set)

    tree_scrollbar_x = ttk.Scrollbar(left_panel_frame, orient="horizontal", command=file_tree.xview, style='TScrollbar')
    tree_scrollbar_x.grid(row=1, column=0, sticky="ew")
    file_tree.configure(xscrollcommand=tree_scrollbar_x.set)

    # 変数リストの表示
    ttk.Label(left_panel_frame, text="変数", style='Header.TLabel', background="#FFFFFF").grid(row=1, column=0, sticky="nw", padx=5, pady=(10,5))
    variable_listbox = tk.Listbox(
        left_panel_frame,
        bg="#F8F8F8", fg="#333333", selectbackground="#3498DB", selectforeground="#FFFFFF",
        font=("Courier", 9), relief="solid", bd=1
    )
    variable_listbox.grid(row=1, column=0, sticky="nsew", padx=5, pady=(30,5))
    
    # 変数リストスクロールバー
    var_list_scrollbar_y = ttk.Scrollbar(left_panel_frame, orient="vertical", command=variable_listbox.yview, style='TScrollbar')
    var_list_scrollbar_y.grid(row=1, column=1, sticky="ns", pady=(30,5))
    variable_listbox.config(yscrollcommand=var_list_scrollbar_y.set)

    def add_files_to_treeview(tree, current_dir, parent_iid, search_term="", active_extensions=None, search_scope="all", search_type="partial"):
        """Treeviewにファイルとサブディレクトリを再帰的に追加する。"""
        if active_extensions is None:
            active_extensions = {'.csv', '.h5', '.hdf', '.xlsx', '.xls'}

        try:
            items = sorted(os.listdir(current_dir), key=lambda s: (not os.path.isdir(os.path.join(current_dir, s)), s.lower()))
            for item_name in items:
                path = os.path.join(current_dir, item_name)
                
                is_dir = os.path.isdir(path)
                is_file = os.path.isfile(path)
                
                match_name = False
                if search_type == "partial":
                    match_name = search_term in item_name.lower()
                else:
                    match_name = search_term == item_name.lower()

                if is_dir:
                    if (search_scope == "all" or search_scope == "directories") and (match_name or not search_term):
                        display_name = get_relative_path(path, global_current_working_directory)
                        dir_iid = tree.insert(parent_iid, "end", text=display_name, open=False, tags=("directory",))
                        add_files_to_treeview(tree, path, dir_iid, search_term, active_extensions, search_scope, search_type)
                elif is_file:
                    file_ext = os.path.splitext(item_name)[1].lower()
                    if file_ext in active_extensions:
                        if (search_scope == "all" or search_scope == "files") and (match_name or not search_term):
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

        # Treeviewの現在の内容をクリアする前に、選択をクリア
        file_tree.selection_remove(file_tree.selection())
        for item in file_tree.get_children():
            file_tree.delete(item)

        for root_path in global_root_directories:
            root_name = os.path.basename(root_path)
            root_match = False
            if search_type_val == "partial":
                root_match = search_term in root_name.lower()
            else:
                root_match = search_term == root_name.lower()

            if search_scope_val == "directories" and not (root_match or not search_term):
                continue 

            display_root_name = get_relative_path(root_path, global_current_working_directory)
            root_item_id = file_tree.insert("", "end", text=display_root_name, open=True, tags=("directory",))
            add_files_to_treeview(file_tree, root_path, root_item_id, search_term, active_extensions, search_scope_val, search_type_val)

    # 初期ディレクトリの追加 (初回起動時のみ)
    if initial_directory_paths:
        for path in initial_directory_paths:
            if path and path not in global_root_directories:
                global_root_directories.append(path)
    
    filter_treeview()
    update_variable_list(variable_listbox)

    # --- 右パネル: データフレーム表示と操作 ---
    right_frame = ttk.Frame(file_processing_page, style='White.TFrame')
    right_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
    right_frame.rowconfigure(0, weight=0) # 現在のファイルラベル
    right_frame.rowconfigure(1, weight=1) # データフレームテキスト表示
    right_frame.rowconfigure(2, weight=0) # スクロールバー
    right_frame.rowconfigure(3, weight=0) # 範囲入力フレーム
    right_frame.rowconfigure(4, weight=0) # フィルタ式入力
    right_frame.rowconfigure(5, weight=0) # 注意書き
    right_frame.rowconfigure(6, weight=0) # 新しい機能ボタン行
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
        bg="#F8F8F8",
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
                                                 row_label_entry, col_label_entry, filter_expression_entry),
        style='TButton',
        cursor="hand2"
    )
    display_button.grid(row=0, column=8, rowspan=2, sticky="nsew", padx=10, pady=5) # 2行にまたがる

    # フィルタ式入力
    filter_frame = ttk.Frame(right_frame, style='White.TFrame')
    filter_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
    filter_frame.columnconfigure(1, weight=1)
    ttk.Label(filter_frame, text="フィルタ式 (Pandas query):", style='TLabel').grid(row=0, column=0, sticky="w", padx=5)
    filter_expression_entry = ttk.Entry(filter_frame, width=50, style='TEntry')
    filter_expression_entry.grid(row=0, column=1, sticky="ew", padx=5)
    ttk.Label(filter_frame, text="例: `col_A > 10 and col_B == 'value'`", style='Note.TLabel').grid(row=1, column=1, sticky="w", padx=5)


    # DataFrameの行/列選択に関する注意書き
    ttk.Label(
        right_frame,
        text="注: 行/列のドラッグ選択はTextウィジェットでは直接サポートされていません。\n数値インデックスまたはラベル、または '開始:終了' 形式で範囲指定して表示してください。",
        style='Note.TLabel'
    ).grid(row=5, column=0, sticky="ew", pady=5)

    # 新しい機能ボタン
    feature_buttons_frame = ttk.Frame(right_frame, style='White.TFrame')
    feature_buttons_frame.grid(row=6, column=0, sticky="ew", padx=10, pady=10)
    feature_buttons_frame.columnconfigure(0, weight=1)
    feature_buttons_frame.columnconfigure(1, weight=1)
    feature_buttons_frame.columnconfigure(2, weight=1)

    embed_var_button = ttk.Button(
        feature_buttons_frame,
        text="変数に組み込む",
        command=lambda: embed_variables_dialog(file_processing_page, loaded_dataframes.get(current_dataframe_path + (f"_{current_dataframe_sheet}" if current_dataframe_sheet else "")), current_dataframe_path, current_dataframe_sheet, variable_listbox),
        style='TButton',
        cursor="hand2"
    )
    embed_var_button.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

    embed_multi_var_button = ttk.Button(
        feature_buttons_frame,
        text="複数ファイルを変数に組み込む",
        command=lambda: embed_multiple_variables_from_selection(file_processing_page, file_tree,
                                                                 start_row_entry, end_row_entry, 
                                                                 start_col_entry, end_col_entry,
                                                                 row_label_entry, col_label_entry, variable_listbox),
        style='TButton',
        cursor="hand2"
    )
    embed_multi_var_button.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

    plot_button = ttk.Button(
        feature_buttons_frame,
        text="プロット",
        command=lambda: show_plot_page(file_processing_page),
        style='Green.TButton',
        cursor="hand2"
    )
    plot_button.grid(row=0, column=2, sticky="ew", padx=5, pady=5)

    calculate_button = ttk.Button(
        feature_buttons_frame,
        text="演算",
        command=lambda: show_calculation_page(file_processing_page),
        style='Green.TButton',
        cursor="hand2"
    )
    calculate_button.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)


    def on_tree_select(event):
        """Treeviewでアイテムが選択されたときのイベントハンドラ。"""
        selected_item_id = file_tree.selection()
        if not selected_item_id:
            return
        
        # 複数選択されている場合は、最初のアイテムのみを処理
        # これは単一ファイル表示の目的のためであり、複数ファイル操作は専用ボタンを使用
        if isinstance(selected_item_id, tuple) and len(selected_item_id) > 1:
            selected_item_id = selected_item_id[0] 

        item_tags = file_tree.item(selected_item_id, "tags")
        
        if "file" in item_tags:
            file_path = file_tree.item(selected_item_id, "values")[0]
            sheet_name = None
            if file_path.lower().endswith(('.xlsx', '.xls')):
                sheet_name = prompt_for_excel_sheet(file_processing_page, file_path)
                if sheet_name is None:
                    file_tree.selection_remove(selected_item_id)
                    return
            load_and_display_dataframe(file_path, sheet_name, dataframe_text, current_file_label, 
                                       start_row_entry, end_row_entry, start_col_entry, end_col_entry,
                                       row_label_entry, col_label_entry, filter_expression_entry)
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
    directory_selection_page.geometry("800x550")
    directory_selection_page.configure(bg="#F0F2F5")

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
        exportselection=False,
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
                display_path = get_relative_path(directory, current_working_directory)
                selected_dirs_listbox.insert(tk.END, display_path)
                next_button.config(state="normal")
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
        
        for index in reversed(selected_indices):
            selected_directories_list.pop(index)
            selected_dirs_listbox.delete(index)
        
        if not selected_directories_list:
            next_button.config(state="disabled")


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
            show_file_processing_page(selected_directories_list)
        else:
            messagebox.showwarning("警告", "ディレクトリが選択されていません。")

    next_button = ttk.Button(
        directory_selection_page,
        text="次へ",
        command=proceed_to_analysis,
        style='Green.TButton',
        cursor="hand2",
        state="disabled"
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
    start_page.configure(bg="#F0F2F5")

    configure_styles() 

    global global_current_working_directory
    current_cwd_var = tk.StringVar(value=global_current_working_directory)

    cwd_history_var = tk.StringVar()
    
    cwd_frame = ttk.Frame(start_page, style='LightGray.TFrame')
    cwd_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
    cwd_frame.columnconfigure(1, weight=1)

    ttk.Label(cwd_frame, text="現在の作業ディレクトリ:", style='SubHeader.TLabel', background='#E0E0E0').grid(row=0, column=0, sticky="w", padx=5, pady=5)
    cwd_entry = ttk.Entry(cwd_frame, textvariable=current_cwd_var, width=60, style='TEntry')
    cwd_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

    def browse_cwd():
        new_cwd = filedialog.askdirectory()
        if new_cwd:
            current_cwd_var.set(new_cwd)
            if new_cwd not in global_cwd_history:
                global_cwd_history.append(new_cwd)
                cwd_history_combobox['values'] = global_cwd_history

    browse_cwd_button = ttk.Button(
        cwd_frame,
        text="参照",
        command=browse_cwd,
        style='TButton',
        cursor="hand2"
    )
    browse_cwd_button.grid(row=0, column=2, sticky="e", padx=5, pady=5)

    ttk.Label(cwd_frame, text="履歴から選択:", style='SubHeader.TLabel', background='#E0E0E0').grid(row=1, column=0, sticky="w", padx=5, pady=5)
    cwd_history_combobox = ttk.Combobox(
        cwd_frame,
        textvariable=cwd_history_var,
        values=global_cwd_history,
        state="readonly",
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
        text=ASCII_ART_HALLAL,
        font=("Courier", 18, "bold"),
        justify="center",
        foreground="#007bff",
        background="#F0F2F5"
    )
    HALLAL_label.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

    hallal_arabic_label = ttk.Label(
        start_page,
        text="حَلَّلَ",
        font=("Arial", 60, "bold"),
        justify="center",
        foreground="#333333",
        background="#F0F2F5"
    )
    hallal_arabic_label.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)

    def start_badaea():
        global global_current_working_directory
        global_current_working_directory = current_cwd_var.get()
        if global_current_working_directory not in global_cwd_history:
            global_cwd_history.append(global_current_working_directory)
        
        start_page.destroy()
        show_directory_selection_page(global_current_working_directory)

    start_button = ttk.Button(
        start_page,
        text="開始",
        command=start_badaea,
        style='Green.TButton',
        cursor="hand2"
    )
    start_button.grid(row=2, column=0, sticky="nsew", padx=50, pady=20)

    start_page.columnconfigure(0, weight=1)
    start_page.rowconfigure(0, weight=0)
    start_page.rowconfigure(1, weight=3)
    start_page.rowconfigure(2, weight=1)
    start_page.rowconfigure(3, weight=2)

    start_page.mainloop()

# アプリケーションの開始
if __name__ == "__main__":
    create_start_page()
