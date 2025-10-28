"""
Excel Writer のテスト
"""
import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from core.excel_reader import ExcelReader
from core.diff_engine import DiffEngine
from core.excel_writer import ExcelWriter


def test_excel_writer():
    """Excel Writer の機能をテスト"""
    
    print("=" * 70)
    print("Excel Writer テスト開始")
    print("=" * 70)
    
    tests_dir = Path(__file__).parent
    old_file = tests_dir / 'sample_old.xlsx'
    new_file = tests_dir / 'sample_new.xlsx'
    sheet_name = '社員リスト'
    
    # Step 1: データ読み込み
    print("\n【Step 1】データ読み込み")
    old_reader = ExcelReader(old_file)
    new_reader = ExcelReader(new_file)
    
    old_df, _ = old_reader.read_sheet(sheet_name)
    new_df, _ = new_reader.read_sheet(sheet_name)
    
    print(f"  ✅ 読み込み完了")
    
    # Step 2: Diff Engine 実行
    print("\n【Step 2】差分検出")
    engine = DiffEngine(old_df, new_df)
    results = engine.compare()
    
    print(f"  差分数: {len(results)} 件")
    print(f"    - 追加: {sum(1 for r in results if r.change_type == 'added')} 行")
    print(f"    - 削除: {sum(1 for r in results if r.change_type == 'deleted')} 行")
    print(f"    - 変更: {sum(1 for r in results if r.change_type == 'changed')} 行")
    
    # Step 3: Excel 出力
    print("\n【Step 3】Excel ファイル出力")
    writer = ExcelWriter(tests_dir, sheet_name)
    columns = new_df.columns.tolist()
    
    output_path = writer.write_diff_results(columns, results)
    
    print(f"  ✅ 出力完了: {output_path.name}")
    print(f"  📂 場所: {output_path}")
    
    # 確認
    print("\n【Step 4】ファイル確認")
    if output_path.exists():
        file_size = output_path.stat().st_size / 1024  # KB
        print(f"  ✅ ファイル存在確認 OK")
        print(f"  📊 ファイルサイズ: {file_size:.2f} KB")
    else:
        print(f"  ❌ ファイルが見つかりません")
    
    print("\n" + "=" * 70)
    print("テスト完了")
    print("=" * 70)
    print(f"\n💡 出力ファイルを開いて確認してください:")
    print(f"   {output_path}")


if __name__ == '__main__':
    test_excel_writer()
