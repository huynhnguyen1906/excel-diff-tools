"""
データ正規化モジュール
比較前のデータ前処理を担当
"""
import pandas as pd
import numpy as np
from typing import Tuple


class DataNormalizer:
    """データを正規化して比較可能な状態にするクラス"""
    
    @staticmethod
    def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        """
        DataFrame を正規化
        
        Args:
            df: 正規化する DataFrame
            
        Returns:
            正規化された DataFrame
        """
        df = df.copy()
        
        # 各列を正規化
        for col in df.columns:
            df[col] = DataNormalizer._normalize_column(df[col])
        
        return df
    
    @staticmethod
    def _normalize_column(series: pd.Series) -> pd.Series:
        """
        1つの列を正規化
        
        Args:
            series: 正規化する Series
            
        Returns:
            正規化された Series
        """
        # 空値を統一: None, NaN, "" → pd.NA
        series = series.replace(['', 'None', 'none', 'NULL', 'null'], pd.NA)
        series = series.replace({np.nan: pd.NA, None: pd.NA})
        
        # 各セルを正規化
        normalized = []
        for value in series:
            if pd.isna(value):
                normalized.append(pd.NA)
            else:
                normalized.append(DataNormalizer._normalize_value(value))
        
        return pd.Series(normalized, index=series.index)
    
    @staticmethod
    def _normalize_value(value):
        """
        単一の値を正規化
        
        Args:
            value: 正規化する値
            
        Returns:
            正規化された値
        """
        # 既に NA の場合
        if pd.isna(value):
            return pd.NA
        
        # 文字列の場合: trim
        if isinstance(value, str):
            trimmed = value.strip()
            if trimmed == "":
                return pd.NA
            
            # 数値として解釈可能かチェック
            try:
                # 整数として解釈
                if '.' not in trimmed:
                    return str(int(trimmed))
                # 浮動小数点数として解釈
                float_val = float(trimmed)
                # 整数値なら整数として返す (1.0 → "1")
                if float_val.is_integer():
                    return str(int(float_val))
                return str(float_val)
            except (ValueError, AttributeError):
                # 数値でない場合はそのまま返す
                return trimmed
        
        # 数値の場合: 文字列に変換
        if isinstance(value, (int, float)):
            # NaN チェック
            if pd.isna(value):
                return pd.NA
            # 整数値なら整数として
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)
        
        # 日付時刻の場合: ISO format に変換
        if isinstance(value, pd.Timestamp):
            return value.strftime('%Y-%m-%d %H:%M:%S')
        
        # その他の型: 文字列に変換
        return str(value).strip()
    
    @staticmethod
    def align_columns(old_df: pd.DataFrame, new_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        新ファイルの column 順序に合わせて old_df を再配置
        
        Args:
            old_df: 旧データ
            new_df: 新データ
            
        Returns:
            (aligned_old_df, aligned_new_df): column が揃った DataFrame のペア
        """
        # 新ファイルの column を基準とする
        new_columns = new_df.columns.tolist()
        
        # Old に存在しない column は空値で追加
        for col in new_columns:
            if col not in old_df.columns:
                old_df[col] = pd.NA
        
        # Old の column を New の順序に並び替え (New にない column は削除)
        old_df = old_df[new_columns]
        
        return old_df, new_df
    
    @staticmethod
    def create_row_signature(row: pd.Series) -> str:
        """
        行の署名（ハッシュ値の元）を作成
        
        Args:
            row: DataFrame の 1 行
            
        Returns:
            行を表す文字列
        """
        # 各セルを文字列に変換して連結
        parts = []
        for value in row:
            if pd.isna(value):
                parts.append("<NA>")
            else:
                parts.append(str(value))
        
        return "|".join(parts)
