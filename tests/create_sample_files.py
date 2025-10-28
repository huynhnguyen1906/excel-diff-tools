"""
テスト用のサンプル Excel ファイルを生成
"""
import pandas as pd
from datetime import datetime
from pathlib import Path

# テストデータ作成
def create_sample_files():
    """サンプル Excel ファイルを作成"""
    
    # 旧ファイルのデータ
    old_data = {
        'ID': [1, 2, 3, 4, 5],
        'レコード番号': ['REC001', 'REC002', 'REC003', 'REC004', 'REC005'],
        '名前': ['田中太郎', '佐藤花子', '鈴木一郎', '高橋美咲', '渡辺健太'],
        '年齢': [25, 30, 28, 35, 22],
        '部署': ['営業部', '開発部', '営業部', '人事部', '開発部'],
        '給与': [300000, 450000, 350000, 500000, 280000],
        '作成日時': [
            datetime(2025, 5, 8, 15, 28, 0),
            datetime(2025, 5, 9, 9, 15, 30),
            datetime(2025, 5, 10, 14, 45, 15),
            datetime(2025, 5, 11, 11, 20, 45),
            datetime(2025, 5, 12, 16, 50, 0)
        ]
    }
    
    # 新ファイルのデータ（いくつか変更を追加）
    new_data = {
        'ID': [1, 2, 3, 5, 6],  # 4が削除、6が追加
        'レコード番号': ['REC001', 'REC002', 'REC003', 'REC005', 'REC006'],  # REC004削除、REC006追加
        '名前': ['田中太郎', '佐藤花子', '鈴木一郎', '渡辺健太', '山田次郎'],  # 新規追加
        '年齢': [25, 31, 28, 23, 27],  # 佐藤の年齢変更、渡辺の年齢変更
        '部署': ['営業部', '企画部', '営業部', '開発部', '営業部'],  # 佐藤の部署変更
        '給与': [300000, 480000, 350000, 290000, 320000],  # 佐藤、渡辺の給与変更
        '作成日時': [
            datetime(2025, 5, 8, 15, 28, 0),    # 変更なし
            datetime(2025, 5, 9, 10, 30, 0),     # 時刻変更
            datetime(2025, 5, 10, 14, 45, 15),  # 変更なし
            datetime(2025, 5, 12, 16, 50, 0),   # 変更なし (ID=5)
            datetime(2025, 6, 1, 9, 0, 0)        # 新規追加
        ]
    }
    
    # DataFrame作成
    df_old = pd.DataFrame(old_data)
    df_new = pd.DataFrame(new_data)
    
    # Excelファイルとして保存
    tests_dir = Path(__file__).parent
    
    old_file = tests_dir / 'sample_old.xlsx'
    new_file = tests_dir / 'sample_new.xlsx'
    
    df_old.to_excel(old_file, index=False, sheet_name='社員リスト')
    df_new.to_excel(new_file, index=False, sheet_name='社員リスト')
    
    print(f"✅ サンプルファイル作成完了:")
    print(f"   - {old_file}")
    print(f"   - {new_file}")
    
    return old_file, new_file


if __name__ == '__main__':
    create_sample_files()
