"""
Monthly Sales Diff Engine
月別売上２シート専用の差分検出エンジン - 月別ブロック比較アルゴリズム
"""
from dataclasses import dataclass
from typing import List, Dict, Literal, Optional
import pandas as pd


@dataclass
class MonthlySalesCellDiff:
    """セルの差分情報"""
    row_index: int  # 行番号 (0-indexed, データ行基準)
    column_name: str  # 列名 ("売上", "外部原価", etc.)
    old_value: Optional[float]  # 旧値
    new_value: Optional[float]  # 新値
    change_type: Literal['increased', 'decreased', 'unchanged']  # 変更タイプ
    
    def get_diff_value(self) -> float:
        """差分値を取得"""
        if self.old_value is None or self.new_value is None:
            return 0.0
        return self.new_value - self.old_value
    
    def get_formatted_text(self) -> str:
        """フォーマットされたテキストを取得"""
        if self.change_type == 'unchanged':
            return str(self.new_value) if self.new_value is not None else ''
        
        old_str = f"{self.old_value:,.0f}" if self.old_value is not None else '0'
        new_str = f"{self.new_value:,.0f}" if self.new_value is not None else '0'
        
        if self.change_type == 'increased':
            return f"↑~{old_str}~ → {new_str}"
        else:  # decreased
            return f"↓~{old_str}~ → {new_str}"


@dataclass
class MonthlySalesMonthDiff:
    """月ブロック全体の差分情報"""
    month_name: str  # 月名
    column_range: tuple  # (start_col, end_col) 出力Excel での列範囲
    cell_diffs: List[MonthlySalesCellDiff]  # セル単位の差分リスト


class MonthlySalesDiffEngine:
    """月別売上２シート専用 - 差分検出エンジン"""
    
    def __init__(self):
        """初期化"""
        self.tolerance = 0.01  # 浮動小数点の許容誤差
    
    def compare_sheets(
        self,
        old_data: Dict,
        new_data: Dict
    ) -> List[MonthlySalesMonthDiff]:
        """
        2つのシートデータを比較
        
        Args:
            old_data: 旧データ (reader.read_sheet() の戻り値)
            new_data: 新データ (reader.read_sheet() の戻り値)
            
        Returns:
            月別差分のリスト
        """
        old_months = set(old_data['month_blocks'].keys())
        new_months = set(new_data['month_blocks'].keys())
        
        # 共通の月のみを比較対象とする
        common_months = old_months & new_months
        
        # 月の順序を保持 (新ファイルの順序を使用)
        ordered_months = [m for m in new_data['month_order'] if m in common_months]
        
        month_diffs = []
        current_col = 3  # D列から開始 (0-indexed: 3)
        
        for month_name in ordered_months:
            old_block = old_data['month_blocks'][month_name]
            new_block = new_data['month_blocks'][month_name]
            
            # ブロックの差分を計算
            cell_diffs = self._compare_blocks(old_block, new_block)
            
            # 差分がある場合のみ追加（unchanged 以外が存在する場合）
            has_changes = any(cd.change_type != 'unchanged' for cd in cell_diffs)
            
            if has_changes:
                # 列範囲を計算 (4列分)
                month_diff = MonthlySalesMonthDiff(
                    month_name=month_name,
                    column_range=(current_col, current_col + 3),
                    cell_diffs=cell_diffs
                )
                
                month_diffs.append(month_diff)
                current_col += 4  # 次のブロックへ
        
        return month_diffs
    
    def _compare_blocks(
        self,
        old_block: pd.DataFrame,
        new_block: pd.DataFrame
    ) -> List[MonthlySalesCellDiff]:
        """
        2つの月ブロックを比較
        
        Args:
            old_block: 旧ブロックのDataFrame
            new_block: 新ブロックのDataFrame
            
        Returns:
            セル差分のリスト
        """
        cell_diffs = []
        
        # 行数を合わせる (多い方に合わせる)
        max_rows = max(len(old_block), len(new_block))
        
        # 各セルを比較
        for row_idx in range(max_rows):
            for col_name in old_block.columns:
                old_value = None
                new_value = None
                
                # 旧値を取得
                if row_idx < len(old_block):
                    old_value = old_block.iloc[row_idx][col_name]
                    if pd.isna(old_value):
                        old_value = None
                
                # 新値を取得
                if row_idx < len(new_block):
                    new_value = new_block.iloc[row_idx][col_name]
                    if pd.isna(new_value):
                        new_value = None
                
                # 変更タイプを判定
                change_type = self._determine_change_type(old_value, new_value)
                
                # 差分がある場合のみ記録 (unchanged でも記録する)
                cell_diff = MonthlySalesCellDiff(
                    row_index=row_idx,
                    column_name=col_name,
                    old_value=old_value,
                    new_value=new_value,
                    change_type=change_type
                )
                
                cell_diffs.append(cell_diff)
        
        return cell_diffs
    
    def _determine_change_type(
        self,
        old_value: Optional[float],
        new_value: Optional[float]
    ) -> Literal['increased', 'decreased', 'unchanged']:
        """
        変更タイプを判定
        
        Args:
            old_value: 旧値
            new_value: 新値
            
        Returns:
            変更タイプ
        """
        # 両方Noneまたは0の場合は unchanged
        if (old_value is None or old_value == 0) and (new_value is None or new_value == 0):
            return 'unchanged'
        
        # どちらかがNoneの場合
        if old_value is None:
            old_value = 0.0
        if new_value is None:
            new_value = 0.0
        
        # 差分を計算
        diff = new_value - old_value
        
        # 許容誤差内なら unchanged
        if abs(diff) <= self.tolerance:
            return 'unchanged'
        
        # 増加 or 減少
        if diff > 0:
            return 'increased'
        else:
            return 'decreased'
    
    def get_summary(self, month_diffs: List[MonthlySalesMonthDiff]) -> Dict:
        """
        差分のサマリーを取得
        
        Args:
            month_diffs: 月別差分のリスト
            
        Returns:
            サマリー情報
        """
        total_changed = 0
        total_increased = 0
        total_decreased = 0
        
        for month_diff in month_diffs:
            for cell_diff in month_diff.cell_diffs:
                if cell_diff.change_type == 'increased':
                    total_increased += 1
                    total_changed += 1
                elif cell_diff.change_type == 'decreased':
                    total_decreased += 1
                    total_changed += 1
        
        return {
            'total_months': len(month_diffs),
            'total_changed_cells': total_changed,
            'total_increased': total_increased,
            'total_decreased': total_decreased,
            'months': [m.month_name for m in month_diffs]
        }
