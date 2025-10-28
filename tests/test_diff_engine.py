"""
Diff Engine のテスト
"""
import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from core.excel_reader import ExcelReader
from core.diff_engine import DiffEngine


def test_diff_engine():
    """Diff Engine の機能をテスト"""
    
    print("=" * 70)
    print("Diff Engine テスト開始")
    print("=" * 70)
    
    tests_dir = Path(__file__).parent
    old_file = tests_dir / 'sample_old.xlsx'
    new_file = tests_dir / 'sample_new.xlsx'
    sheet_name = '社員リスト'
    
    # データ読み込み
    print("\n【Step 1】データ読み込み")
    old_reader = ExcelReader(old_file)
    new_reader = ExcelReader(new_file)
    
    old_df, error = old_reader.read_sheet(sheet_name)
    new_df, error = new_reader.read_sheet(sheet_name)
    
    if old_df is None or new_df is None:
        print(f"❌ エラー: {error}")
        return
    
    print(f"  旧データ: {len(old_df)} 行, {len(old_df.columns)} 列")
    print(f"  新データ: {len(new_df)} 行, {len(new_df.columns)} 列")
    
    # Diff Engine 実行
    print("\n【Step 2】Diff Engine 実行")
    engine = DiffEngine(old_df, new_df)
    results = engine.compare()
    
    print(f"  差分数: {len(results)} 件")
    
    # 結果を種類別にカウント
    added_count = sum(1 for r in results if r.change_type == 'added')
    deleted_count = sum(1 for r in results if r.change_type == 'deleted')
    changed_count = sum(1 for r in results if r.change_type == 'changed')
    
    print(f"\n  📊 内訳:")
    print(f"     - 追加 (Added):   {added_count} 行")
    print(f"     - 削除 (Deleted): {deleted_count} 行")
    print(f"     - 変更 (Changed): {changed_count} 行")
    
    # 詳細表示
    print("\n【Step 3】差分詳細")
    print("-" * 70)
    
    for i, result in enumerate(results, 1):
        print(f"\n{i}. [{result.change_type.upper()}] Row {result.row_index}")
        
        if result.change_type == 'deleted':
            print(f"   旧データ:")
            for col, val in result.old_data.items():
                print(f"     {col}: {val}")
        
        elif result.change_type == 'added':
            print(f"   新データ:")
            for col, val in result.new_data.items():
                print(f"     {col}: {val}")
        
        elif result.change_type == 'changed':
            print(f"   変更された列: {', '.join(result.changed_columns)}")
            for col in result.changed_columns:
                old_val = result.old_data[col]
                new_val = result.new_data[col]
                print(f"     {col}: {old_val} → {new_val}")
    
    print("\n" + "=" * 70)
    print("テスト完了")
    print("=" * 70)


if __name__ == '__main__':
    test_diff_engine()
