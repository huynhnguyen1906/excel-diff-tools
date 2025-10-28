"""
差分エンジン
レコード番号ベースの比較アルゴリズムを実装
"""
from dataclasses import dataclass
from typing import List, Literal, Optional, Tuple, Dict
from difflib import SequenceMatcher
from collections import defaultdict
import pandas as pd

from core.data_normalizer import DataNormalizer


@dataclass
class DiffResult:
    """差分結果を表すデータクラス"""
    row_index: int  # 出力 Excel での行番号
    change_type: Literal['added', 'deleted', 'changed']  # 変更タイプ
    old_data: Optional[dict]  # 旧データ (deleted/changed のみ)
    new_data: Optional[dict]  # 新データ (added/changed のみ)
    changed_columns: Optional[List[str]]  # 変更された列名 (changed のみ)
    record_number: Optional[str] = None  # レコード番号 (ソート用)


class DiffEngine:
    """レコード番号ベース差分比較エンジン"""
    
    # 類似度のしきい値
    SIMILARITY_THRESHOLD = 0.6
    
    # レコード番号列のインデックス (Column B = index 1)
    RECORD_NUMBER_COL_INDEX = 1
    
    def __init__(self, old_df: pd.DataFrame, new_df: pd.DataFrame):
        """
        初期化
        
        Args:
            old_df: 旧データ (正規化済み、列揃え済み)
            new_df: 新データ (正規化済み、列揃え済み)
        """
        # レコード番号列名を取得
        if len(old_df.columns) > self.RECORD_NUMBER_COL_INDEX:
            self.record_col = old_df.columns[self.RECORD_NUMBER_COL_INDEX]
        else:
            self.record_col = None
        
        # レコード番号が空でない行のみを抽出
        if self.record_col is not None:
            self.old_df = old_df[old_df[self.record_col].notna() & (old_df[self.record_col] != '')].copy()
            self.new_df = new_df[new_df[self.record_col].notna() & (new_df[self.record_col] != '')].copy()
        else:
            self.old_df = old_df.copy()
            self.new_df = new_df.copy()
        
        self.results: List[DiffResult] = []
    
    @staticmethod
    def compare_dataframes(old_df: pd.DataFrame, new_df: pd.DataFrame) -> List[DiffResult]:
        """
        2つの DataFrame を比較する便利メソッド
        
        Args:
            old_df: 旧データ (正規化済み)
            new_df: 新データ (正規化済み)
            
        Returns:
            差分結果のリスト
        """
        engine = DiffEngine(old_df, new_df)
        return engine.compare()
    
    def compare(self) -> List[DiffResult]:
        """
        レコード番号ベースの比較を実行
        
        Returns:
            差分結果のリスト (レコード番号でソート済み)
        """
        if self.record_col is None:
            # レコード番号列がない場合は従来の方法
            return self._compare_without_record_number()
        
        # Phase 1: レコード番号でグループ化
        old_groups = self._group_by_record_number(self.old_df)
        new_groups = self._group_by_record_number(self.new_df)
        
        # すべてのレコード番号を取得
        all_record_numbers = set(old_groups.keys()) | set(new_groups.keys())
        
        results = []
        
        # Phase 2: レコード番号ごとに比較
        for record_num in all_record_numbers:
            old_rows = old_groups.get(record_num, [])
            new_rows = new_groups.get(record_num, [])
            
            if not old_rows and new_rows:
                # 新規レコード番号 → ADDED
                for new_idx in new_rows:
                    new_row = self.new_df.iloc[new_idx]
                    results.append(DiffResult(
                        row_index=0,  # 後で更新
                        change_type='added',
                        old_data=None,
                        new_data=new_row.to_dict(),
                        changed_columns=None,
                        record_number=str(record_num)
                    ))
            
            elif old_rows and not new_rows:
                # 削除されたレコード番号 → DELETED
                for old_idx in old_rows:
                    old_row = self.old_df.iloc[old_idx]
                    results.append(DiffResult(
                        row_index=0,  # 後で更新
                        change_type='deleted',
                        old_data=old_row.to_dict(),
                        new_data=None,
                        changed_columns=None,
                        record_number=str(record_num)
                    ))
            
            else:
                # 同じレコード番号のグループ内で比較
                group_results = self._compare_within_group(
                    old_rows, new_rows, record_num
                )
                results.extend(group_results)
        
        # Phase 3: レコード番号でソート
        # 数値の場合のみソート (降順: 大→小)
        try:
            # すべてのレコード番号が数値かチェック
            all_numeric = all(
                r.record_number and r.record_number.replace('.', '', 1).replace('-', '', 1).isdigit()
                for r in results if r.record_number
            )
            
            if all_numeric:
                # 数値としてソート
                results.sort(key=lambda r: float(r.record_number) if r.record_number else 0, reverse=True)
        except (ValueError, AttributeError):
            # ソートに失敗した場合は元の順序を維持
            pass
        
        # Phase 4: 行番号を更新
        for idx, result in enumerate(results, start=1):
            result.row_index = idx
        
        return results
    
    def _group_by_record_number(self, df: pd.DataFrame) -> Dict[str, List[int]]:
        """
        レコード番号でグループ化
        
        Args:
            df: DataFrame
            
        Returns:
            {record_number: [row_indices]}
        """
        groups = defaultdict(list)
        
        for idx, row in df.iterrows():
            record_num = str(row[self.record_col]) if not pd.isna(row[self.record_col]) else ""
            groups[record_num].append(idx)
        
        return dict(groups)
    
    def _compare_within_group(self, 
                              old_indices: List[int], 
                              new_indices: List[int],
                              record_num: str) -> List[DiffResult]:
        """
        同じレコード番号のグループ内で比較
        
        Args:
            old_indices: 旧データの行インデックス
            new_indices: 新データの行インデックス
            record_num: レコード番号
            
        Returns:
            差分結果のリスト
        """
        results = []
        
        # グループ内の行データを取得
        old_rows = [self.old_df.iloc[i] for i in old_indices]
        new_rows = [self.new_df.iloc[i] for i in new_indices]
        
        # 行数が同じ場合: 順番にマッチング (高速)
        if len(old_rows) == len(new_rows):
            for old_row, new_row in zip(old_rows, new_rows):
                changed_cols = self._diff_cells(old_row, new_row)
                
                if changed_cols:
                    results.append(DiffResult(
                        row_index=0,  # 後で更新
                        change_type='changed',
                        old_data=old_row.to_dict(),
                        new_data=new_row.to_dict(),
                        changed_columns=changed_cols,
                        record_number=record_num
                    ))
        else:
            # 行数が異なる場合: 簡略化マッチング
            # 長いリストと短いリストを特定
            if len(old_rows) > len(new_rows):
                # Old が多い → 削除が多い
                for i, new_row in enumerate(new_rows):
                    if i < len(old_rows):
                        changed_cols = self._diff_cells(old_rows[i], new_row)
                        if changed_cols:
                            results.append(DiffResult(
                                row_index=0,
                                change_type='changed',
                                old_data=old_rows[i].to_dict(),
                                new_data=new_row.to_dict(),
                                changed_columns=changed_cols,
                                record_number=record_num
                            ))
                
                # 残りの old rows は削除
                for i in range(len(new_rows), len(old_rows)):
                    results.append(DiffResult(
                        row_index=0,
                        change_type='deleted',
                        old_data=old_rows[i].to_dict(),
                        new_data=None,
                        changed_columns=None,
                        record_number=record_num
                    ))
            else:
                # New が多い → 追加が多い
                for i, old_row in enumerate(old_rows):
                    if i < len(new_rows):
                        changed_cols = self._diff_cells(old_row, new_rows[i])
                        if changed_cols:
                            results.append(DiffResult(
                                row_index=0,
                                change_type='changed',
                                old_data=old_row.to_dict(),
                                new_data=new_rows[i].to_dict(),
                                changed_columns=changed_cols,
                                record_number=record_num
                            ))
                
                # 残りの new rows は追加
                for i in range(len(old_rows), len(new_rows)):
                    results.append(DiffResult(
                        row_index=0,
                        change_type='added',
                        old_data=None,
                        new_data=new_rows[i].to_dict(),
                        changed_columns=None,
                        record_number=record_num
                    ))
        
        return results
    
    def _compute_similarity(self, row1: pd.Series, row2: pd.Series) -> float:
        """
        2つの行の類似度を計算
        
        Args:
            row1: 行1
            row2: 行2
            
        Returns:
            類似度 (0.0 ~ 1.0)
        """
        sig1 = DataNormalizer.create_row_signature(row1.to_dict())
        sig2 = DataNormalizer.create_row_signature(row2.to_dict())
        
        matcher = SequenceMatcher(None, sig1, sig2)
        return matcher.ratio()
    
    def _compare_without_record_number(self) -> List[DiffResult]:
        """
        レコード番号なしの従来の比較方法
        (Fallback用 - 現在は使用しない想定)
        """
        # 従来のアルゴリズムをそのまま使用
        return []
    
    def _create_signatures(self, df: pd.DataFrame) -> List[str]:
        """
        各行の署名を作成
        
        Args:
            df: DataFrame
            
        Returns:
            各行の署名文字列のリスト
        """
        signatures = []
        for idx, row in df.iterrows():
            sig = DataNormalizer.create_row_signature(row)
            signatures.append(sig)
        return signatures
    
    def _diff_cells(self, old_row: pd.Series, new_row: pd.Series) -> List[str]:
        """
        2つの行のセルレベルの差分を検出
        
        Args:
            old_row: 旧データの行
            new_row: 新データの行
            
        Returns:
            変更された列名のリスト
        """
        changed_columns = []
        
        for col in old_row.index:
            old_val = old_row[col]
            new_val = new_row[col]
            
            # 両方とも NA の場合は変更なし
            if pd.isna(old_val) and pd.isna(new_val):
                continue
            
            # 片方だけ NA の場合は変更あり
            if pd.isna(old_val) or pd.isna(new_val):
                changed_columns.append(col)
                continue
            
            # 値が異なる場合は変更あり
            if str(old_val) != str(new_val):
                changed_columns.append(col)
        
        return changed_columns
