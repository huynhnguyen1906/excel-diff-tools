"""
Processor Factory
シート名に基づいて適切なプロセッサを返すファクトリ
"""
from .base_processor import BaseProcessor
from .haritsuke import HaritsukeProcessor
from .monthly_sales import MonthlySalesProcessor


class ProcessorFactory:
    """プロセッサファクトリクラス"""
    
    @staticmethod
    def get_processor(sheet_name: str) -> BaseProcessor:
        """
        シート名に基づいてプロセッサを返す
        
        Args:
            sheet_name: シート名
            
        Returns:
            適切なプロセッサインスタンス
            
        Raises:
            ValueError: サポートされていないシート名の場合
        """
        if sheet_name == "貼付":
            return HaritsukeProcessor()
        elif sheet_name == "月別売上２":
            return MonthlySalesProcessor()
        else:
            raise ValueError(f"サポートされていないシート: {sheet_name}")
