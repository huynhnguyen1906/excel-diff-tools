"""
Haritsuke Processor
貼付シート専用プロセッサ - メイン処理ロジック
"""
from pathlib import Path
from typing import Optional, Tuple, List, Callable

from core.processors.base_processor import BaseProcessor
from core.data_normalizer import DataNormalizer
from .reader import HaritsukeExcelReader
from .diff_engine import HaritsukeDiffEngine
from .writer import HaritsukeExcelWriter


class HaritsukeProcessor(BaseProcessor):
    """貼付シート専用プロセッサ"""
    
    SHEET_NAME = "貼付"
    RECORD_NUMBER_COLUMN = 1  # Column B (0-indexed)
    
    def get_sheet_name(self) -> str:
        """対応するシート名を返す"""
        return self.SHEET_NAME
    
    def process(
        self, 
        old_file: Path, 
        new_file: Path, 
        sheet_name: str, 
        output_dir: Path,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ) -> Tuple[Optional[Path], Optional[List], str]:
        """
        貼付シートの差分処理を実行
        
        Args:
            old_file: 旧ファイルのパス
            new_file: 新ファイルのパス
            sheet_name: シート名
            output_dir: 出力先ディレクトリ
            progress_callback: 進捗コールバック関数 (progress%, message)
            
        Returns:
            (output_path, diff_results, error_message):
                - output_path: 成功時は出力ファイルパス、失敗時は None
                - diff_results: 差分結果のリスト、失敗時は None
                - error_message: エラーがあればメッセージ
        """
        try:
            # Step 1: ファイル読み込み
            if progress_callback:
                progress_callback(20, "ファイルを読み込んでいます...")
            
            old_reader = HaritsukeExcelReader(old_file)
            new_reader = HaritsukeExcelReader(new_file)
            
            old_df, error_msg = old_reader.read_sheet(sheet_name)
            if old_df is None:
                return None, None, f"旧ファイル読み込みエラー:\n{error_msg}"
            
            new_df, error_msg = new_reader.read_sheet(sheet_name)
            if new_df is None:
                return None, None, f"新ファイル読み込みエラー:\n{error_msg}"
            
            # Step 2: データ正規化
            if progress_callback:
                progress_callback(40, "データを正規化しています...")
            
            normalizer = DataNormalizer()
            old_df_normalized = normalizer.normalize_dataframe(old_df)
            new_df_normalized = normalizer.normalize_dataframe(new_df)
            
            # 列を揃える
            old_df_aligned, new_df_aligned = normalizer.align_columns(
                old_df_normalized, new_df_normalized
            )
            
            # Step 3: 差分検出
            if progress_callback:
                progress_callback(60, "差分を検出しています...")
            
            diff_results = HaritsukeDiffEngine.compare_dataframes(
                old_df_aligned, new_df_aligned
            )
            
            # 更新日時列のみの変更を除外
            # 更新日時列名（複数の可能性を考慮）
            update_time_columns = ['更新日時', '更新日', 'UpdateTime', 'UpdateDate']
            
            filtered_results = []
            for result in diff_results:
                if result.change_type == 'changed':
                    # 更新日時以外の列に変更があるかチェック
                    has_other_changes = any(
                        col not in update_time_columns 
                        for col in result.changed_columns
                    )
                    
                    # 更新日時以外の変更がある場合のみ追加
                    if has_other_changes:
                        # changed_columns はそのまま維持（更新日時も含む）
                        filtered_results.append(result)
                    # else: 更新日時のみの変更 → 除外
                else:
                    # added/deleted は常に含める
                    filtered_results.append(result)
            
            # row_index を再割り当て（連番にする）
            for idx, result in enumerate(filtered_results, start=1):
                result.row_index = idx
            
            diff_results = filtered_results
            
            # 差分が無い場合
            if not diff_results:
                return None, None, "データに差分がありませんでした。\n旧ファイルと新ファイルのデータは完全に一致しています。"
            
            # Step 4: Excel 出力
            if progress_callback:
                progress_callback(80, "結果を出力しています...")
            
            writer = HaritsukeExcelWriter(output_dir, sheet_name)
            columns = new_df_aligned.columns.tolist()
            output_path = writer.write_diff_results(columns, diff_results)
            
            if progress_callback:
                progress_callback(100, "完了しました")
            
            return output_path, diff_results, ""
            
        except Exception as e:
            return None, None, f"処理中にエラーが発生しました:\n{str(e)}"
