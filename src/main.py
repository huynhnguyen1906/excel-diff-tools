"""
Excel Sheet Diff - メインエントリーポイント
エクセルシート比較ツール
"""
import sys
import os

# PyInstallerでビルドされた場合の対応
if hasattr(sys, '_MEIPASS'):
    # EXEとして実行: _MEIPASSをパスに追加
    os.chdir(sys._MEIPASS)

from PySide6.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    """アプリケーションのメインエントリーポイント"""
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Sheet Diff")
    app.setOrganizationName("YourCompany")
    
    # メインウィンドウを作成して表示
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
