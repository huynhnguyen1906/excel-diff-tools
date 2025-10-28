"""
Base Processor - Abstract base class for all sheet processors
すべてのシートプロセッサの抽象基底クラス
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, Tuple, List, Any, Callable


class BaseProcessor(ABC):
    """シートプロセッサの抽象基底クラス"""
    
    @abstractmethod
    def get_sheet_name(self) -> str:
        """
        このプロセッサが対応するシート名を返す
        
        Returns:
            シート名
        """
        pass
    
    @abstractmethod
    def process(
        self, 
        old_file: Path, 
        new_file: Path, 
        sheet_name: str, 
        output_dir: Path,
        progress_callback: Optional[Callable[[int, str], None]] = None
    ) -> Tuple[Optional[Path], Optional[List[Any]], str]:
        """
        差分処理を実行
        
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
        pass
