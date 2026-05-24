import sys
from pathlib import Path
import polars as pl
from tableauhyperapi import HyperProcess, Telemetry, Connection, CreateMode, TableDefinition, SqlType, Inserter

def infer_and_convert_schema(csv_path):
    """
    CSVを読み込み、最初の10行を基準に数値型（Int/Float）を優先判定し、
    Polarsのデータフレームとして返す
    """
    # 1. 最初の10行だけを「すべて文字列（String）」として読み込む
    preview_df = pl.read_csv(csv_path, n_rows=10, infer_schema_length=0)
    
    schema_overrides = {}
    
    for col in preview_df.columns:
        sample_values = preview_df[col].to_list()
        
        is_int = True
        is_float = True
        
        for val in sample_values:
            if val is None or str(val).strip() == "":
                continue
                
            val_str = str(val).strip()
            
            # 整数(Int)の判定
            if is_int:
                if not (val_str.isdigit() or (val_str.startswith('-') and val_str[1:].isdigit())):
                    is_int = False
            
            # 浮動小数点数(Float)の判定
            if is_float:
                try:
                    float(val_str)
                except ValueError:
                    is_float = False
                    
        # 数値型を優先して型を確定（デフォルトはテキスト型）
        if is_int:
            schema_overrides[col] = pl.Int64
        elif is_float:
            schema_overrides[col] = pl.Float64
        else:
            schema_overrides[col] = pl.String

    # 2. 確定したスキーマ型を適用して、CSV全体を本読み込み
    df = pl.read_csv(csv_path, schema_overrides=schema_overrides)
    return df

def map_polars_type_to_hyper(pl_type):
    """
    Polarsのデータ型を、Tableau Hyper APIのSqlTypeにマッピングする
    """
    if pl_type == pl.Int64 or pl_type == pl.Int32:
        return SqlType.big_int()
    elif pl_type == pl.Float64 or pl_type == pl.Float32:
        return SqlType.double()
    elif pl_type == pl.Boolean:
        return SqlType.boolean()
    elif pl_type == pl.Date:
        return SqlType.date()
    elif pl_type == pl.Datetime:
        return SqlType.timestamp()
    else:
        return SqlType.text()

def process_and_save_to_hyper_direct(csv_path, output_hyper_path):
    """
    Polarsでデータを読み込み、加工処理を挟んでHyper APIで直接.hyperを出力する
    """
    print(f"1. CSVの型判定と読み込みを開始します: {csv_path.name}")
    df = infer_and_convert_schema(csv_path)
    
    # ---------------------------------------------------------------
    # 【加工処理ステップの余地】
    # 将来的にフィルタリング、列の追加、型変換、結合などの処理をここに記述できます。
    # ---------------------------------------------------------------
    print("2. 加工処理ステップを実行中... (必要に応じてここにコードを追加)")
    # 例: df = df.with_columns(pl.col("日付列").str.to_date("%Y/%m/%d"))
    # ---------------------------------------------------------------
    
    # 最終的なデータ型の対応表を画面に表示する処理
    print("\n==============================================")
    print("【最終データ型確認】Hyperファイルに書き込まれるスキーマ")
    print("==============================================")
    print(f"{'列名':<20} | {'Polarsでの型':<15} -> {'Hyperでの型'}")
    print("-" * 60)
    
    # CSVのファイル名をテーブル名として定義
    table_name = csv_path.stem
    table_def = TableDefinition(table_name=table_name)
    
    for col_name, pl_type in df.schema.items():
        hyper_type = map_polars_type_to_hyper(pl_type)
        table_def.add_column(name=col_name, type=hyper_type)
        
        # 画面に見やすく表示
        print(f"{col_name:<20} | {str(pl_type):<15} -> {str(hyper_type)}")
    print("==============================================\n")
    
    print("3. Hyper APIを使用して直接 .hyper ファイルを生成中...")
    
    # Hyperファイルの起動とデータベース作成
    with HyperProcess(telemetry=Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
        with Connection(endpoint=hyper.endpoint, database=str(output_hyper_path), create_mode=CreateMode.CREATE_AND_REPLACE) as connection:
            
            # データベース内にテーブルを作成
            connection.catalog.create_table(table_definition=table_def)
            
            # 修正: 'table_definition' ではなく 'table' を使用します
            with Inserter(connection, table_def) as inserter:

                inserter.add_rows(df.iter_rows())
                inserter.execute()
                
    print(f"成功: Hyperファイルが生成されました -> {output_hyper_path.resolve()}")

# ==========================================
# 実行メイン処理
# ==========================================
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("エラー: 読み込む .csv ファイルのパスを指定してください。")
        print("使い方: python script.py <CSVファイルのパス>")
        sys.exit(1)
        
    # sys.argv の後ろに「[1]」を記述して、1番目の引数（ファイルパス）を取り出します
    input_csv = Path(sys.argv[1])
    if not input_csv.exists() or input_csv.suffix.lower() != '.csv':
        print(f"エラー: 有効なCSVファイルが見つかりません -> {input_csv}")
        sys.exit(1)
        
    output_hyper = input_csv.with_suffix('.hyper')
    process_and_save_to_hyper_direct(input_csv, output_hyper)
