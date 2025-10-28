"""
メインウィンドウ
Excel Diff アプリケーションのメイン UI
"""
import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QMessageBox,
    QProgressDialog
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

# Core modules をインポート
from core.excel_reader import ExcelReader
from core.data_normalizer import DataNormalizer
from core.diff_engine import DiffEngine
from core.excel_writer import ExcelWriter


def get_executable_dir():
    """実行ファイルのディレクトリを取得"""
    if getattr(sys, 'frozen', False):
        # PyInstaller で実行されている場合
        return Path(sys.executable).parent
    else:
        # 開発環境で実行されている場合
        return Path.cwd()


class DiffWorker(QThread):
    """差分処理を行うワーカースレッド"""
    
    # シグナル定義
    progress = Signal(int, str)  # (進捗%, メッセージ)
    finished = Signal(object, str)  # (結果, エラーメッセージ)
    
    def __init__(self, old_file, new_file, sheet_name, output_dir):
        super().__init__()
        self.old_file = old_file
        self.new_file = new_file
        self.sheet_name = sheet_name
        self.output_dir = output_dir
        self.error_msg = None
        self.result = None
    
    def run(self):
        """バックグラウンドで処理を実行"""
        try:
            # Step 1: データ読み込み
            self.progress.emit(10, "データを読み込んでいます...")
            
            old_reader = ExcelReader(self.old_file)
            new_reader = ExcelReader(self.new_file)
            
            old_df, error_msg = old_reader.read_sheet(self.sheet_name)
            if old_df is None:
                self.finished.emit(None, f"旧ファイル読み込みエラー:\n{error_msg}")
                return
            
            self.progress.emit(20, "データを読み込んでいます...")
            
            new_df, error_msg = new_reader.read_sheet(self.sheet_name)
            if new_df is None:
                self.finished.emit(None, f"新ファイル読み込みエラー:\n{error_msg}")
                return
            
            self.progress.emit(30, "データを読み込んでいます...")
            
            # Step 2: データ正規化
            self.progress.emit(40, "データを正規化しています...")
            normalizer = DataNormalizer()
            old_df_normalized = normalizer.normalize_dataframe(old_df)
            new_df_normalized = normalizer.normalize_dataframe(new_df)
            
            # 列を揃える
            old_df_aligned, new_df_aligned = normalizer.align_columns(
                old_df_normalized, new_df_normalized
            )
            self.progress.emit(50, "データを正規化しています...")
            
            # Step 3: 差分検出
            self.progress.emit(60, "差分を検出しています...")
            print(f"[DEBUG] Starting diff comparison... Old rows: {len(old_df_aligned)}, New rows: {len(new_df_aligned)}")
            
            diff_results = DiffEngine.compare_dataframes(
                old_df_aligned, new_df_aligned
            )
            
            print(f"[DEBUG] Diff completed. Found {len(diff_results)} differences")
            self.progress.emit(80, "差分を検出しています...")
            
            # Step 4: Excel 出力
            self.progress.emit(90, "Excel ファイルを作成しています...")
            print(f"[DEBUG] Writing Excel file...")
            writer = ExcelWriter(self.output_dir, self.sheet_name)
            columns = new_df_aligned.columns.tolist()
            
            output_path = writer.write_diff_results(columns, diff_results)
            
            print(f"[DEBUG] Excel file written: {output_path}")
            self.progress.emit(100, "完了しました")
            
            # 結果を返す
            result = {
                'output_path': output_path,
                'diff_results': diff_results
            }
            self.finished.emit(result, None)
            
        except Exception as e:
            self.finished.emit(None, f"処理中にエラーが発生しました:\n{str(e)}")


class MainWindow(QMainWindow):
    """メインウィンドウクラス"""
    
    def __init__(self):
        super().__init__()
        self.old_file_path = None
        self.new_file_path = None
        self.output_dir = get_executable_dir()  # デフォルトは実行ファイルのディレクトリ
        self._init_ui()
    
    def _init_ui(self):
        """UI 初期化"""
        self.setWindowTitle("Excel Sheet Diff - シート比較ツール")
        self.setFixedSize(750, 650)
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # タイトル
        title = QLabel("Excel シート比較ツール")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # 説明
        description = QLabel("2つのExcelファイルのシートを比較して、差分を出力します")
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(description)
        
        layout.addSpacing(20)
        
        # シート名入力行
        sheet_layout = QHBoxLayout()
        sheet_label = QLabel("シート名:")
        sheet_label.setFixedWidth(100)
        sheet_label.setStyleSheet("font-size: 12px;")
        
        self.sheet_input = QLineEdit()
        self.sheet_input.setPlaceholderText("例: Sheet1")
        self.sheet_input.setFixedHeight(35)
        self.sheet_input.setStyleSheet("""
            QLineEdit {
                padding: 5px 10px;
                border: 2px solid #ccc;
                border-radius: 4px;
                font-size: 13px;
            }
            QLineEdit:focus {
                border: 2px solid #0078d4;
            }
        """)
        self.sheet_input.textChanged.connect(self._on_sheet_name_changed)
        
        sheet_layout.addWidget(sheet_label)
        sheet_layout.addWidget(self.sheet_input)
        layout.addLayout(sheet_layout)
        
        layout.addSpacing(3)
        
        # 説明テキスト
        sheet_note = QLabel("         ※ シート名を入力すると、ファイル選択ボタンが有効になります")
        sheet_note.setStyleSheet("color: #888; font-size: 12px;")
        layout.addWidget(sheet_note)
        
        layout.addSpacing(20)
        
        # ファイルセクションタイトル
        file_section_title = QLabel("② Excelファイルを選択してください")
        file_section_title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        file_section_title.setStyleSheet("margin-bottom: 10px;")
        layout.addWidget(file_section_title)
        
        # 旧ファイルセクション
        self._add_file_section(
            layout, 
            "旧ファイル（比較元）:",
            "old"
        )
        
        # 新ファイルセクション
        self._add_file_section(
            layout,
            "新ファイル（比較先）:",
            "new"
        )
        
        layout.addSpacing(20)
        
        # 出力先セクション
        output_section_title = QLabel("③ 出力先フォルダを選択してください")
        output_section_title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        output_section_title.setStyleSheet("margin-bottom: 10px;")
        layout.addWidget(output_section_title)
        
        # 出力先フォルダ選択
        output_layout = QHBoxLayout()
        
        output_label = QLabel("出力先:")
        output_label.setFixedWidth(100)
        output_label.setStyleSheet("font-size: 12px;")
        
        self.output_input = QLineEdit()
        self.output_input.setReadOnly(True)
        self.output_input.setText(str(self.output_dir))
        self.output_input.setFixedHeight(38)
        self.output_input.setStyleSheet("""
            QLineEdit {
                padding: 8px 12px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f9f9f9;
                color: #333;
                font-size: 12px;
            }
        """)
        
        output_browse_button = QPushButton("📁 参照")
        output_browse_button.setFixedSize(110, 38)
        output_browse_button.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
        """)
        output_browse_button.clicked.connect(self._on_output_browse_clicked)
        
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(output_browse_button)
        layout.addLayout(output_layout)
        
        layout.addStretch()
        
        # 比較ボタン
        self.compare_button = QPushButton("比較実行")
        self.compare_button.setFixedHeight(32)
        self.compare_button.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.compare_button.clicked.connect(self._on_compare_clicked)
        layout.addWidget(self.compare_button)
    
    def _add_file_section(self, parent_layout, label_text, file_type):
        """ファイル選択セクションを追加"""
        section_layout = QVBoxLayout()
        section_layout.setSpacing(8)
        
        # ラベル
        label = QLabel(label_text)
        label.setFont(QFont("Segoe UI", 10))
        label.setStyleSheet("color: #333; margin-bottom: 3px;")
        section_layout.addWidget(label)
        
        # ファイルパス表示 + ボタン
        file_layout = QHBoxLayout()
        file_layout.setSpacing(10)
        
        # ファイルパス表示
        file_input = QLineEdit()
        file_input.setReadOnly(True)
        file_input.setPlaceholderText("ファイルが選択されていません")
        file_input.setFixedHeight(38)
        file_input.setStyleSheet("""
            QLineEdit {
                padding: 8px 12px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f9f9f9;
                color: #333;
                font-size: 12px;
            }
        """)
        
        # 参照ボタン
        browse_button = QPushButton("📁 参照")
        browse_button.setFixedSize(110, 38)
        browse_button.setEnabled(False)  # デフォルトは無効
        browse_button.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover:enabled {
                background-color: #106ebe;
            }
            QPushButton:pressed:enabled {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #e0e0e0;
                color: #999;
            }
        """)
        browse_button.clicked.connect(
            lambda: self._on_browse_clicked(file_input, file_type)
        )
        
        file_layout.addWidget(file_input)
        file_layout.addWidget(browse_button)
        
        section_layout.addLayout(file_layout)
        parent_layout.addLayout(section_layout)
        
        # オブジェクトを保存
        if file_type == "old":
            self.old_file_input = file_input
            self.old_browse_button = browse_button
        else:
            self.new_file_input = file_input
            self.new_browse_button = browse_button
    
    def _on_sheet_name_changed(self, text):
        """シート名入力の変更イベント"""
        # シート名が入力されているかチェック
        has_sheet_name = text.strip() != ""
        
        # ファイル選択ボタンの有効/無効を切り替え
        self.old_browse_button.setEnabled(has_sheet_name)
        self.new_browse_button.setEnabled(has_sheet_name)
        
        # ファイルが選択済みの場合、シート名変更でクリア
        if not has_sheet_name:
            self.old_file_path = None
            self.new_file_path = None
            self.old_file_input.clear()
            self.new_file_input.clear()
    
    def _on_browse_clicked(self, input_widget, file_type):
        """ファイル参照ボタンのクリックイベント"""
        # 入力されたシート名を取得
        sheet_name = self.sheet_input.text().strip()
        
        if not sheet_name:
            QMessageBox.warning(
                self,
                "エラー",
                "先にシート名を入力してください"
            )
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Excelファイルを選択",
            str(get_executable_dir()),
            "Excel Files (*.xlsx *.xlsm);;All Files (*)"
        )
        
        if file_path:
            # ファイルパスを設定
            file_path_obj = Path(file_path)
            
            # Excel Reader で検証
            reader = ExcelReader(file_path_obj)
            is_valid, error_msg = reader.validate_file()
            
            if not is_valid:
                QMessageBox.warning(
                    self,
                    "ファイルエラー",
                    f"ファイルを読み込めません:\n\n{error_msg}"
                )
                return
            
            # 指定されたシートが存在するか確認
            is_valid, error_msg = reader.validate_sheet(sheet_name)
            if not is_valid:
                QMessageBox.warning(
                    self,
                    "シートエラー",
                    f"指定されたシートが見つかりません:\n\n{error_msg}"
                )
                return
            
            # すべて検証OKなら設定
            input_widget.setText(file_path)
            if file_type == "old":
                self.old_file_path = file_path_obj
            else:
                self.new_file_path = file_path_obj
    
    def _on_output_browse_clicked(self):
        """出力先フォルダ選択ボタンのクリックイベント"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "出力先フォルダを選択",
            str(self.output_dir)
        )
        
        if folder_path:
            self.output_dir = Path(folder_path)
            self.output_input.setText(str(self.output_dir))
    
    def _on_compare_clicked(self):
        """比較ボタンのクリックイベント"""
        # バリデーション
        if not self.old_file_path:
            QMessageBox.warning(self, "エラー", "旧ファイルを選択してください")
            return
        
        if not self.new_file_path:
            QMessageBox.warning(self, "エラー", "新ファイルを選択してください")
            return
        
        sheet_name = self.sheet_input.text().strip()
        if not sheet_name:
            QMessageBox.warning(self, "エラー", "シート名を入力してください")
            return
        
        # Excel Reader で検証
        old_reader = ExcelReader(self.old_file_path)
        new_reader = ExcelReader(self.new_file_path)
        
        # 旧ファイルのシート検証
        is_valid, error_msg = old_reader.validate_sheet(sheet_name)
        if not is_valid:
            QMessageBox.warning(
                self,
                "シートエラー（旧ファイル）",
                f"旧ファイルでエラーが発生しました:\n\n{error_msg}"
            )
            return
        
        # 新ファイルのシート検証
        is_valid, error_msg = new_reader.validate_sheet(sheet_name)
        if not is_valid:
            QMessageBox.warning(
                self,
                "シートエラー（新ファイル）",
                f"新ファイルでエラーが発生しました:\n\n{error_msg}"
            )
            return
        
        # プログレスダイアログを表示
        self.progress_dialog = QProgressDialog("処理を開始しています...", None, 0, 100, self)
        self.progress_dialog.setWindowTitle("差分検出")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)
        
        # 比較ボタンを無効化
        self.compare_button.setEnabled(False)
        
        # ワーカースレッド作成
        self.worker = DiffWorker(
            self.old_file_path,
            self.new_file_path,
            sheet_name,
            self.output_dir
        )
        
        # シグナル接続
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.finished.connect(self._on_worker_finished)
        
        # スレッド開始
        self.worker.start()
    
    def _on_worker_progress(self, value, message):
        """ワーカースレッドの進捗更新"""
        self.progress_dialog.setValue(value)
        self.progress_dialog.setLabelText(message)
    
    def _on_worker_finished(self, result, error_msg):
        """ワーカースレッド完了時の処理"""
        # プログレスダイアログを閉じる
        self.progress_dialog.close()
        
        # 比較ボタンを有効化
        self.compare_button.setEnabled(True)
        
        if error_msg:
            # エラーの場合
            QMessageBox.critical(self, "エラー", error_msg)
            return
        
        # 成功の場合
        diff_results = result['diff_results']
        output_path = result['output_path']
        
        added_count = sum(1 for r in diff_results if r.change_type == 'added')
        deleted_count = sum(1 for r in diff_results if r.change_type == 'deleted')
        changed_count = sum(1 for r in diff_results if r.change_type == 'changed')
        
        # 結果メッセージ
        if len(diff_results) == 0:
            result_msg = (
                "差分は検出されませんでした。\n\n"
                "2つのファイルは同じ内容です。"
            )
        else:
            result_msg = (
                f"差分検出が完了しました！\n\n"
                f"【検出結果】\n"
                f"  🟢 追加: {added_count} 行\n"
                f"  🔴 削除: {deleted_count} 行\n"
                f"  🟡 変更: {changed_count} 行\n"
                f"  合計: {len(diff_results)} 件の差分\n\n"
                f"【出力ファイル】\n"
                f"  {output_path.name}\n\n"
                f"📂 場所: {output_path.parent}"
            )
        
        QMessageBox.information(self, "完了", result_msg)
