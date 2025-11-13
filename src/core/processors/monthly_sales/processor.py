"""
Monthly Sales Processor
月別売上２シート専用プロセッサ - メイン処理ロジック
"""
from pathlib import Path
from typing import Optional, Tuple, List, Callable

from core.processors.base_processor import BaseProcessor
from .reader import MonthlySalesExcelReader
from .diff_engine import MonthlySalesDiffEngine
from .writer import MonthlySalesExcelWriter


class MonthlySalesProcessor(BaseProcessor):
    """月別売上２シート専用プロセッサ"""
    
    SHEET_NAME = "月別売上２"
    
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
        月別売上２シートの差分処理
        
        Args:
            old_file: 旧ファイルのパス
            new_file: 新ファイルのパス
            sheet_name: シート名
            output_dir: 出力先ディレクトリ
            progress_callback: 進捗報告用コールバック関数
            
        Returns:
            Tuple of (output_path, diff_results, error_message)
        """
        try:
            # 1. ファイル読み取り (20%)
            if progress_callback:
                progress_callback(20, "ファイルを読み取り中...")
            
            reader = MonthlySalesExcelReader()
            
            # 旧ファイル検証
            is_valid, error = reader.validate_file(old_file)
            if not is_valid:
                return None, None, f"旧ファイルエラー: {error}"
            
            is_valid, error = reader.validate_sheet(old_file, sheet_name)
            if not is_valid:
                return None, None, f"旧ファイルシートエラー:\n{error}"
            
            # 新ファイル検証
            is_valid, error = reader.validate_file(new_file)
            if not is_valid:
                return None, None, f"新ファイルエラー: {error}"
            
            is_valid, error = reader.validate_sheet(new_file, sheet_name)
            if not is_valid:
                return None, None, f"新ファイルシートエラー:\n{error}"
            
            # データ読み取り
            old_data = reader.read_sheet(old_file, sheet_name)
            new_data = reader.read_sheet(new_file, sheet_name)
            
            # 2. 差分検出 (60%)
            if progress_callback:
                progress_callback(60, "差分を検出中...")
            
            diff_engine = MonthlySalesDiffEngine()
            month_diffs = diff_engine.compare_sheets(old_data, new_data)
            
            # 共通月が存在するか確認
            old_months = set(old_data['month_blocks'].keys())
            new_months = set(new_data['month_blocks'].keys())
            common_months = old_months & new_months
            
            if not common_months:
                # 共通月が無い場合
                return None, None, "共通の月が見つかりませんでした。\n旧ファイルと新ファイルで同じ月のデータが存在しません。"
            elif not month_diffs:
                # 共通月はあるが差分が無い場合
                return None, None, "データに差分がありませんでした。\n旧ファイルと新ファイルのデータは完全に一致しています。"
            
            # 3. Excel出力 (80%)
            if progress_callback:
                progress_callback(80, "結果を出力中...")
            
            writer = MonthlySalesExcelWriter(output_dir, sheet_name)
            output_path = writer.write_diff_result(
                header_df=new_data['header'],
                category_df=new_data['categories'],
                month_diffs=month_diffs,
                source_data=new_data
            )
            
            # 4. サマリー作成
            summary = diff_engine.get_summary(month_diffs)
            
            # 差分結果を SimpleNamespace 形式で返す (UI表示用)
            # UI が .change_type でアクセスするため
            from types import SimpleNamespace
            
            diff_results = []
            for month_diff in month_diffs:
                for cell_diff in month_diff.cell_diffs:
                    if cell_diff.change_type != 'unchanged':
                        # change_type を UI が期待する形式にマッピング
                        # increased → changed (緑), decreased → changed (赤)
                        ui_change_type = 'changed'  # 月別売上は全て「変更」扱い
                        
                        # SimpleNamespace でラップして attribute アクセスを可能にする
                        result = SimpleNamespace(
                            month=month_diff.month_name,
                            row=cell_diff.row_index + 7,  # 実際の行番号
                            column=cell_diff.column_name,
                            change_type=ui_change_type,  # UI が期待する形式
                            original_change_type=cell_diff.change_type,  # 元の情報も保持
                            old_value=cell_diff.old_value,
                            new_value=cell_diff.new_value
                        )
                        diff_results.append(result)
            
            # 5. 完了 (100%)
            if progress_callback:
                progress_callback(100, "完了")
            
            return output_path, diff_results, ""
            
        except Exception as e:
            error_msg = f"月別売上２シート処理エラー:\n{str(e)}"
            return None, None, error_msg
