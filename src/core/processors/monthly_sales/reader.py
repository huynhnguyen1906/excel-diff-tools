"""
Monthly Sales Excel Reader
月別売上２シート専用の Excel 読み取りモジュール
"""
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


class MonthlySalesExcelReader:
    """月別売上２シート専用 - Excel ファイルを読み取るクラス"""
    
    # 定数
    HEADER_ROWS = 6  # ヘッダーは6行
    CATEGORY_COLUMNS = 3  # カテゴリ列はA, B, C (3列)
    DATA_START_ROW = 7  # データは7行目から開始
    DATA_START_COLUMN = 4  # データは D列 (index 3) から開始
    BLOCK_SIZE = 4  # 1ヶ月のブロックは4列
    MONTH_ROW = 5  # 月名は5行目 (0-indexed: row 4)
    
    def __init__(self):
        """初期化"""
        pass
    
    def validate_file(self, file_path: Path) -> Tuple[bool, Optional[str]]:
        """
        ファイルの妥当性をチェック
        
        Args:
            file_path: ファイルパス
            
        Returns:
            (成功フラグ, エラーメッセージ)
        """
        if not file_path.exists():
            return False, f"ファイルが見つかりません: {file_path}"
        
        if not file_path.suffix.lower() in ['.xlsx', '.xlsm']:
            return False, f"サポートされていないファイル形式です: {file_path.suffix}"
        
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            wb.close()
            return True, None
        except Exception as e:
            return False, f"ファイルを開けませんでした: {str(e)}"
    
    def get_sheet_names(self, file_path: Path) -> List[str]:
        """
        ファイル内のシート名リストを取得
        
        Args:
            file_path: ファイルパス
            
        Returns:
            シート名のリスト
        """
        wb = load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        return sheet_names
    
    def validate_sheet(self, file_path: Path, sheet_name: str) -> Tuple[bool, Optional[str]]:
        """
        シートの妥当性をチェック
        
        Args:
            file_path: ファイルパス
            sheet_name: シート名
            
        Returns:
            (成功フラグ, エラーメッセージ)
        """
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            
            if sheet_name not in wb.sheetnames:
                wb.close()
                return False, f"シート '{sheet_name}' が見つかりません"
            
            ws = wb[sheet_name]
            
            # 最低限の行数チェック
            if ws.max_row < self.DATA_START_ROW:
                wb.close()
                return False, f"シートの行数が不足しています (最小: {self.DATA_START_ROW}行)"
            
            # 最低限の列数チェック (カテゴリ3列 + 最低1ブロック4列)
            min_columns = self.CATEGORY_COLUMNS + self.BLOCK_SIZE
            if ws.max_column < min_columns:
                wb.close()
                return False, f"シートの列数が不足しています (最小: {min_columns}列)"
            
            wb.close()
            return True, None
            
        except Exception as e:
            return False, f"シート検証エラー: {str(e)}"
    
    def _extract_month_blocks(self, ws: Worksheet) -> Dict[str, Tuple[int, int]]:
        """
        月ブロックの情報を抽出 (月名 → (開始列index, 終了列index))
        
        Args:
            ws: ワークシート
            
        Returns:
            {月名: (start_col, end_col)} の辞書
        """
        month_blocks = {}
        
        # 5行目 (MONTH_ROW) をスキャンして月名を探す
        current_col = self.DATA_START_COLUMN  # D列 (index 3) から開始
        
        while current_col <= ws.max_column:
            # openpyxl は 1-indexed なので +1
            cell_value = ws.cell(row=self.MONTH_ROW, column=current_col + 1).value
            
            if cell_value and isinstance(cell_value, str):
                # "2024/12月" のような形式を期待
                if '月' in cell_value or '/' in cell_value:
                    # ブロックの終了列 = 開始列 + 3 (4列分)
                    end_col = current_col + self.BLOCK_SIZE - 1
                    month_blocks[cell_value] = (current_col, end_col)
                    current_col += self.BLOCK_SIZE
                else:
                    current_col += 1
            else:
                current_col += 1
        
        return month_blocks
    
    def read_sheet(self, file_path: Path, sheet_name: str) -> Dict:
        """
        シートを読み取り、構造化されたデータを返す
        
        Args:
            file_path: ファイルパス
            sheet_name: シート名
            
        Returns:
            {
                'header': DataFrame (行1-6),
                'categories': DataFrame (列A-C, 全行),
                'month_blocks': {月名: DataFrame},
                'month_order': [月名リスト]
            }
        """
        wb = load_workbook(file_path, data_only=True)
        ws = wb[sheet_name]
        
        # ヘッダー部分を読み取り (行1-6, 全列)
        header_data = []
        for row_idx in range(1, self.HEADER_ROWS + 1):
            row_data = []
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                row_data.append(cell.value)
            header_data.append(row_data)
        
        header_df = pd.DataFrame(header_data)
        
        # カテゴリ列 (A-C) を読み取り
        category_data = []
        for row_idx in range(1, ws.max_row + 1):
            row_data = []
            for col_idx in range(1, self.CATEGORY_COLUMNS + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                row_data.append(cell.value)
            category_data.append(row_data)
        
        category_df = pd.DataFrame(category_data, columns=['A', 'B', 'C'])
        
        # 月ブロックを抽出
        month_blocks = self._extract_month_blocks(ws)
        
        # 各月ブロックのデータを読み取り
        month_data = {}
        month_order = []
        
        for month_name, (start_col, end_col) in month_blocks.items():
            month_order.append(month_name)
            
            # データ行 (7行目以降) のみを読み取り
            block_data = []
            for row_idx in range(self.DATA_START_ROW, ws.max_row + 1):
                row_data = []
                for col_idx in range(start_col + 1, end_col + 2):  # +1 for 1-indexed
                    cell = ws.cell(row=row_idx, column=col_idx)
                    value = cell.value
                    
                    # 数値に変換を試みる
                    if value is not None:
                        try:
                            value = float(value)
                        except (ValueError, TypeError):
                            pass
                    
                    row_data.append(value)
                block_data.append(row_data)
            
            # 列名は "売上", "外部原価", "内部原価", "営業利益" と仮定
            columns = ['売上', '外部原価', '内部原価', '営業利益']
            month_df = pd.DataFrame(block_data, columns=columns)
            month_data[month_name] = month_df
        
        wb.close()
        
        return {
            'header': header_df,
            'categories': category_df,
            'month_blocks': month_data,
            'month_order': month_order
        }
    
    def get_sheet_info(self, file_path: Path, sheet_name: str) -> Dict:
        """
        シート情報を取得 (デバッグ用)
        
        Args:
            file_path: ファイルパス
            sheet_name: シート名
            
        Returns:
            シート情報の辞書
        """
        wb = load_workbook(file_path, read_only=True, data_only=True)
        ws = wb[sheet_name]
        
        info = {
            'total_rows': ws.max_row,
            'total_columns': ws.max_column,
            'data_rows': ws.max_row - self.HEADER_ROWS,
            'month_blocks': self._extract_month_blocks(ws)
        }
        
        wb.close()
        return info
