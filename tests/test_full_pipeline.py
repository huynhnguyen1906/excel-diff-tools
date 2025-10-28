"""
完全な処理パイプラインのテスト
GUI なしで Excel Reader → Diff Engine → Excel Writer の流れをテスト
"""
from pathlib import Path
import sys

# プロジェクトルートを sys.path に追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from core.excel_reader import ExcelReader
from core.data_normalizer import DataNormalizer
from core.diff_engine import DiffEngine
from core.excel_writer import ExcelWriter


def test_full_pipeline():
    """完全なパイプラインをテスト"""
    print("=" * 70)
    print("完全パイプラインテスト開始")
    print("=" * 70)
    print()
    
    # ファイルパス
    test_dir = Path(__file__).parent
    old_file = test_dir / "sample_old.xlsx"
    new_file = test_dir / "sample_new.xlsx"
    sheet_name = "社員リスト"
    
    # Step 1: ファイル検証
    print("【Step 1】ファイル検証")
    old_reader = ExcelReader(old_file)
    new_reader = ExcelReader(new_file)
    
    # 旧ファイル検証
    is_valid, error_msg = old_reader.validate_file()
    if not is_valid:
        print(f"  ❌ 旧ファイルエラー: {error_msg}")
        return False
    
    is_valid, error_msg = old_reader.validate_sheet(sheet_name)
    if not is_valid:
        print(f"  ❌ 旧ファイルシートエラー: {error_msg}")
        return False
    
    # 新ファイル検証
    is_valid, error_msg = new_reader.validate_file()
    if not is_valid:
        print(f"  ❌ 新ファイルエラー: {error_msg}")
        return False
    
    is_valid, error_msg = new_reader.validate_sheet(sheet_name)
    if not is_valid:
        print(f"  ❌ 新ファイルシートエラー: {error_msg}")
        return False
    
    print("  ✅ ファイル検証 OK")
    print()
    
    # Step 2: データ読み込み
    print("【Step 2】データ読み込み")
    old_df, error_msg = old_reader.read_sheet(sheet_name)
    if old_df is None:
        print(f"  ❌ 旧ファイル読み込みエラー: {error_msg}")
        return False
    
    new_df, error_msg = new_reader.read_sheet(sheet_name)
    if new_df is None:
        print(f"  ❌ 新ファイル読み込みエラー: {error_msg}")
        return False
    
    print(f"  旧ファイル: {len(old_df)} 行 × {len(old_df.columns)} 列")
    print(f"  新ファイル: {len(new_df)} 行 × {len(new_df.columns)} 列")
    print("  ✅ データ読み込み OK")
    print()
    
    # Step 3: データ正規化
    print("【Step 3】データ正規化")
    normalizer = DataNormalizer()
    old_df_normalized = normalizer.normalize_dataframe(old_df)
    new_df_normalized = normalizer.normalize_dataframe(new_df)
    
    old_df_aligned, new_df_aligned = normalizer.align_columns(
        old_df_normalized, new_df_normalized
    )
    print(f"  正規化後の列数: {len(old_df_aligned.columns)}")
    print("  ✅ データ正規化 OK")
    print()
    
    # Step 4: 差分検出
    print("【Step 4】差分検出")
    diff_results = DiffEngine.compare_dataframes(old_df_aligned, new_df_aligned)
    
    added_count = sum(1 for r in diff_results if r.change_type == 'added')
    deleted_count = sum(1 for r in diff_results if r.change_type == 'deleted')
    changed_count = sum(1 for r in diff_results if r.change_type == 'changed')
    
    print(f"  🟢 追加: {added_count} 行")
    print(f"  🔴 削除: {deleted_count} 行")
    print(f"  🟡 変更: {changed_count} 行")
    print(f"  合計: {len(diff_results)} 件の差分")
    print("  ✅ 差分検出 OK")
    print()
    
    # Step 5: Excel 出力
    print("【Step 5】Excel 出力")
    output_dir = test_dir
    writer = ExcelWriter(output_dir, sheet_name)
    columns = new_df_aligned.columns.tolist()
    output_path = writer.write_diff_results(columns, diff_results)
    
    print(f"  出力ファイル: {output_path.name}")
    print(f"  場所: {output_path.parent}")
    
    if output_path.exists():
        file_size = output_path.stat().st_size / 1024  # KB
        print(f"  ファイルサイズ: {file_size:.2f} KB")
        print("  ✅ Excel 出力 OK")
    else:
        print("  ❌ ファイルが作成されませんでした")
        return False
    
    print()
    print("=" * 70)
    print("✅ すべてのテストが成功しました！")
    print("=" * 70)
    print()
    print(f"💡 出力ファイル: {output_path}")
    
    return True


if __name__ == "__main__":
    success = test_full_pipeline()
    sys.exit(0 if success else 1)
