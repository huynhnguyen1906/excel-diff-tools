"""
Excel Reader のテスト
"""
import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from core.excel_reader import ExcelReader


def test_excel_reader():
    """Excel Reader の機能をテスト"""
    
    print("=" * 60)
    print("Excel Reader テスト開始")
    print("=" * 60)
    
    tests_dir = Path(__file__).parent
    old_file = tests_dir / 'sample_old.xlsx'
    new_file = tests_dir / 'sample_new.xlsx'
    
    # Test 1: ファイル検証
    print("\n【Test 1】ファイル検証")
    reader = ExcelReader(old_file)
    is_valid, error = reader.validate_file()
    print(f"  結果: {'✅ OK' if is_valid else '❌ NG'}")
    if error:
        print(f"  エラー: {error}")
    
    # Test 2: シート名取得
    print("\n【Test 2】シート名取得")
    sheet_names = reader.get_sheet_names()
    print(f"  シート数: {len(sheet_names)}")
    print(f"  シート名: {sheet_names}")
    
    # Test 3: シート検証（存在するシート）
    print("\n【Test 3】シート検証（存在するシート）")
    is_valid, error = reader.validate_sheet('社員リスト')
    print(f"  結果: {'✅ OK' if is_valid else '❌ NG'}")
    if error:
        print(f"  エラー: {error}")
    
    # Test 4: シート検証（存在しないシート）
    print("\n【Test 4】シート検証（存在しないシート）")
    is_valid, error = reader.validate_sheet('存在しないシート')
    print(f"  結果: {'❌ NG' if not is_valid else '✅ OK（本来はNGであるべき）'}")
    if error:
        print(f"  エラー: {error}")
    
    # Test 5: データ読み込み
    print("\n【Test 5】データ読み込み")
    df, error = reader.read_sheet('社員リスト')
    if df is not None:
        print(f"  結果: ✅ OK")
        print(f"  行数: {len(df)}")
        print(f"  列数: {len(df.columns)}")
        print(f"  列名: {df.columns.tolist()}")
        print(f"\n  先頭3行:")
        print(df.head(3).to_string(index=False))
    else:
        print(f"  結果: ❌ NG")
        print(f"  エラー: {error}")
    
    # Test 6: シート情報取得
    print("\n【Test 6】シート情報取得")
    info, error = reader.get_sheet_info('社員リスト')
    if info:
        print(f"  結果: ✅ OK")
        print(f"  行数: {info['rows']}")
        print(f"  列数: {info['columns']}")
        print(f"  列名: {info['column_names']}")
    else:
        print(f"  結果: ❌ NG")
        print(f"  エラー: {error}")
    
    # Test 7: 新ファイルも同様にテスト
    print("\n【Test 7】新ファイルの読み込み")
    reader_new = ExcelReader(new_file)
    df_new, error = reader_new.read_sheet('社員リスト')
    if df_new is not None:
        print(f"  結果: ✅ OK")
        print(f"  行数: {len(df_new)}")
        print(f"\n  先頭3行:")
        print(df_new.head(3).to_string(index=False))
    else:
        print(f"  結果: ❌ NG")
        print(f"  エラー: {error}")
    
    print("\n" + "=" * 60)
    print("テスト完了")
    print("=" * 60)


if __name__ == '__main__':
    test_excel_reader()
