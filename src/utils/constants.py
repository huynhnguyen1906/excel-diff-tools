"""
定数定義
アプリケーション全体で使用する定数、色、フォーマット設定
"""

# 差分タイプの色設定
DIFF_COLORS = {
    'deleted': 'FFCCCC',   # 赤
    'changed': 'FFF2CC',   # 黄
    'added': 'CCFFCC'      # 緑
}

# フォント設定
FONT_NAME = 'Segoe UI'
FONT_SIZE = 10

# 日付フォーマット
DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss'

# カラム幅の最大値（文字数）
MAX_COLUMN_WIDTH = 100

# バッチ処理のサイズ（行数）
BATCH_SIZE = 1000

# 比較の類似度しきい値
SIMILARITY_THRESHOLD = 0.6
