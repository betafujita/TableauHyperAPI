import os
import shutil
import sys
import zipfile
from pathlib import Path
from tableauhyperapi import HyperProcess, Telemetry, Connection, CreateMode

def extract_hyper_from_twbx(twbx_path, extract_dir="extracted_hyper"):
    """
    .twbx ファイル(ZIP形式)から .hyper ファイルをすべて抽出する
    """
    twbx_path = Path(twbx_path)
    extract_dir = Path(extract_dir)
    hyper_files = []
    
    with zipfile.ZipFile(twbx_path, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            if file_info.filename.endswith('.hyper'):
                filename = os.path.basename(file_info.filename)
                target_path = extract_dir / filename
                
                extract_dir.mkdir(parents=True, exist_ok=True)
                
                with zip_ref.open(file_info) as source, open(target_path, "wb") as target:
                    shutil.copyfileobj(source, target)
                
                print(f"抽出成功: {filename}")
                hyper_files.append(target_path)
                
    return hyper_files

def analyze_hyper_structure(hyper_path):
    """
    Hyper API を使用して、.hyper ファイルの構造（カラム・行数・プレビュー）を解析する
    """
    print(f"\n--- {os.path.basename(hyper_path)} の構造解析結果 ---")
    
    with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
        with Connection(endpoint=hyper.endpoint, database=str(hyper_path), create_mode=CreateMode.NONE) as connection:
            schema_names = connection.catalog.get_schema_names()
            
            for schema in schema_names:
                print(f"\n[Schema: {schema}]")
                table_names = connection.catalog.get_table_names(schema=schema)
                
                if not table_names:
                    print("  (テーブルが存在しません)")
                    continue
                    
                for table in table_names:
                    print(f"\n  ■ Table: {table.name}")
                    
                    # 1. カラム（列名とデータ型）の取得
                    print("    【カラム一覧】")
                    table_definition = connection.catalog.get_table_definition(name=table)
                    column_names = []
                    for column in table_definition.columns:
                        print(f"      - {column.name.unescaped}: {column.type}")
                        column_names.append(column.name.unescaped)
                    
                    # 2. 総行数のカウント (SQLの実行)
                    # テーブル名はエスケープが必要な場合があるため、引数から文字列を生成
                    count_query = f"SELECT COUNT(*) FROM {table}"
                    row_count = connection.execute_scalar_query(query=count_query)
                    print(f"    【総行数】\n      {row_count:,} 行")
                    
                    # 3. 先頭3行のデータプレビュー (SQLの実行)
                    if row_count > 0:
                        print("    【データプレビュー (先頭3行)】")
                        preview_query = f"SELECT * FROM {table} LIMIT 300"
                        
                        # execute_query で結果セットを取得
                        with connection.execute_query(query=preview_query) as result:
                            # ヘッダー（列名）の表示
                            print(f"      │ {' │ '.join(column_names[:5])} ...")
                            print(f"      ├─{'─' * 40}")
                            
                            # 行データの表示 (最大5列に制限して見やすく表示)
                            for row in result:
                                row_str = [str(val) for val in row[:5]]
                                print(f"      │ {' │ '.join(row_str)} ...")
                    else:
                        print("    【データプレビュー】\n      データが空です。")



# ==========================================
# 実行メイン処理
# ==========================================
if __name__ == "__main__":
    # 引数の数をチェック (スクリプト名 + twbxパス の計2つが必要)
    if len(sys.argv) < 2:
        print("エラー: 解析対象の .twbx ファイルのパスを指定してください。")
        print("使い方: python script.py <対象ファイルのパス>")
        sys.exit(1)
        
    # 第1引数からファイルパスを取得
    twbx_file_path = Path(sys.argv[1])
    
    # ファイルの存在チェック
    if not twbx_file_path.exists():
        print(f"エラー: 指定されたファイルが存在しません: {twbx_file_path}")
        sys.exit(1)
        
    # 拡張子のチェック
    if twbx_file_path.suffix.lower() != '.twbx':
        print(f"エラー: 拡張子が .twbx ではありません: {twbx_file_path.suffix}")
        sys.exit(1)

    print(f"解析を開始します: {twbx_file_path.name}")
    
    # 1. twbx から hyper ファイルの抽出
    extracted_hypers = extract_hyper_from_twbx(twbx_file_path)
    
    # 2. 抽出された各 hyper ファイルの構造解析
    if extracted_hypers:
        for hyper_file in extracted_hypers:
            analyze_hyper_structure(hyper_file)
    else:
        print("\n.hyper ファイルが検出されませんでした。")
