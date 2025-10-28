"""
Monthly Sales (月別売上２) Sheet Processor
月別売上２ シートの差分処理を担当するモジュール
"""
from pathlib import Path
from typing import Optional, Tuple, List, Callable

from .base_processor import BaseProcessor


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
        月別売上２シートの差分処理を実行（未実装）
        
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
        return None, None, "月別売上２シートの処理は未実装です"
        return None, "月別売上２シートの処理はまだ実装されていません"
