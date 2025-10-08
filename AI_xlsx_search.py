import os
import pandas as pd

# 検索対象のディレクトリを指定します
# ローカルPCのパスはColabから直接アクセスできません。
# ExcelファイルをColabにアップロードし、そのパスを指定してください。
search_directory = r"C:\yoshida\検索用バッチ" # 例: "/content/uploaded_excel_files"

# 指定されたディレクトリにあるすべてのファイルを取得します
try:
    files = os.listdir(search_directory)
except FileNotFoundError:
    print(f"エラー: 指定されたディレクトリ '{search_directory}' が見つかりません。")
    files = []

# Excelファイルのみをフィルタリングします (.xlsx および .xls 拡張子)
excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]

# 検索したい文字列を定義します (必要に応じて変更してください)
search_term = "入出庫ボディワーク"

print(f"以下のExcelファイルを検索します: {excel_files}")

# Excelファイルを繰り返し処理し、検索語を検索します
for excel_file in excel_files:
    file_path = os.path.join(search_directory, excel_file)
    print(f"\nファイル '{file_path}' を処理中...")
    try:
        # Excelファイルの全てのシート名を読み込みます
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names

        if not sheet_names:
            print(f"  ファイル '{excel_file}' にシートが見つかりませんでした。スキップします。")
            continue

        found_in_file = False
        # 各シートを繰り返し処理します
        for sheet_name in sheet_names:
            print(f"  シート '{sheet_name}' を処理中...")
            try:
                # シートをDataFrameとして読み込みます
                df = xl.parse(sheet_name)

                # DataFrameが空でないことを確認します
                if not df.empty:
                    found_in_sheet = False
                    # すべての列を検索語について検索します
                    for column in df.columns:
                        # 列のデータ型を文字列に変換し、検索語が含まれているか確認します
                        # na=False は NaN 値を無視します
                        if df[column].astype(str).str.contains(search_term, na=False).any():
                            print(f"    '{search_term}' がシート '{sheet_name}' の列 '{column}' で見つかりました。")
                            found_in_file = True
                            found_in_sheet = True

                    if not found_in_sheet:
                         print(f"    シート '{sheet_name}' に '{search_term}' は見つかりませんでした。")
                else:
                    print(f"    シート '{sheet_name}' は空です。スキップします。")

            except Exception as e:
                print(f"  シート '{sheet_name}' の処理中にエラーが発生しました: {e}")

        if not found_in_file:
            print(f"  ファイル '{excel_file}' のどのシートにも '{search_term}' は見つかりませんでした。")


    except FileNotFoundError:
        print(f"  エラー: ファイル '{excel_file}' が見つかりません。")
    except Exception as e:
        print(f"  ファイル '{excel_file}' の処理中にエラーが発生しました: {e}")

print("\n検索が完了しました。")
