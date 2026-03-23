import sys
import os
import time
import json
import cv2
import numpy as np
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTextEdit, QLabel, QMainWindow, QLineEdit, 
                             QMessageBox, QComboBox, QCheckBox, QMenuBar, QMenu, QScrollArea)
from PyQt6.QtCore import Qt, QRect, QPoint, QTimer, QThread, pyqtSignal, QBuffer, QIODevice, QSettings
from PyQt6.QtGui import QPainter, QColor, QPen, QGuiApplication, QScreen, QPixmap, QIcon, QAction

import google.generativeai as genai

class HelpWindow(QWidget):
    """
    README.md の内容を表示するウィンドウ
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Glans_Miya - 使い方")
        self.resize(500, 600)
        self.setWindowFlags(Qt.WindowType.Window)
        
        layout = QVBoxLayout()
        
        self.text_browser = QTextEdit()
        self.text_browser.setReadOnly(True)
        
        # README.md の読み込み
        readme_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "README.md")
        if os.path.exists(readme_path):
            with open(readme_path, "r", encoding="utf-8") as f:
                content = f.read()
                self.text_browser.setPlainText(content)
        else:
            self.text_browser.setPlainText("README.md が見つかりませんでした。\nアプリのフォルダを確認してください。")
            
        layout.addWidget(self.text_browser)
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)

class MainWindow(QMainWindow):
    """
    常に最前面に表示される、操作方法の案内とキャプチャ起動ボタンを持つメインウィンドウ
    """
    def __init__(self):
        super().__init__()
        self.version = "v1.0.0"
        self.setWindowTitle(f"Glans_Miya {self.version}")
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(360, 440)
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #ffffff;
                color: #333333;
                font-family: 'Segoe UI', Meiryo, sans-serif;
            }
            QLabel {
                font-size: 13px;
                line-height: 1.5;
            }
            QLineEdit, QComboBox {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: white;
            }
            QCheckBox {
                font-size: 13px;
                padding: 4px;
            }
            QPushButton#SnipBtn {
                background-color: #1a73e8;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px;
                font-size: 14px;
                font-weight: bold;
                margin-top: 10px;
            }
            QPushButton#SnipBtn:hover {
                background-color: #1557b0;
            }
            QPushButton#SaveBtn, QPushButton#TestBtn, QPushButton#RecommendBtn, QPushButton#UpdateBtn {
                background-color: #f1f3f4;
                color: #333;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 8px 12px;
                font-weight: bold;
            }
            QPushButton#SaveBtn:hover, QPushButton#TestBtn:hover, QPushButton#RecommendBtn:hover, QPushButton#UpdateBtn:hover {
                background-color: #e8eaed;
            }
            /* モード選択ボタンのスタイル */
            QPushButton.ModeBtn {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 10px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton.ModeBtn:hover {
                background-color: #e9ecef;
            }
            QPushButton.ModeBtn[selected="true"] {
                background-color: #e8f0fe;
                border: 2px solid #1a73e8;
                color: #1a73e8;
            }
            QPushButton#TestBtn:disabled, QPushButton#UpdateBtn:disabled {
                color: #999;
                background-color: #f8f9fa;
            }
        """)

        # メニューバーの作成
        self.create_menu_bar()

        # 中央ウィジェットの設定
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        # 操作方法の表示
        guide_label = QLabel(
            "<b>【使い方】</b><br>"
            "1. モードを選択し、APIキーを設定します<br>"
            "2. 下のボタンで画面を切り取ります<br>"
            "3. 選択した機能で結果を表示します"
        )
        guide_label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(guide_label)
        
        layout.addSpacing(10)

        # 設定ファイル (Cドライブのユーザー直下) の準備
        app_data_dir = os.path.join(os.path.expanduser('~'), 'Glans_Miya')
        os.makedirs(app_data_dir, exist_ok=True)
        settings_path = os.path.join(app_data_dir, "settings.ini")
        self.settings = QSettings(settings_path, QSettings.Format.IniFormat)

        # --- モード選択エリア ---
        mode_label = QLabel("✨ 機能を選択:")
        layout.addWidget(mode_label)
        
        mode_btn_layout = QHBoxLayout()
        self.btn_ja = QPushButton("日本語翻訳")
        self.btn_en = QPushButton("英語翻訳")
        self.btn_dict = QPushButton("辞書")
        
        for btn in [self.btn_ja, self.btn_en, self.btn_dict]:
            btn.setProperty("class", "ModeBtn")
            btn.setCheckable(True)
            btn.clicked.connect(self.on_mode_changed)
            mode_btn_layout.addWidget(btn)
            
        layout.addLayout(mode_btn_layout)
        
        # 保存されているモードを読み込む (デフォルト: 日本語翻訳)
        self.current_mode = self.settings.value("last_mode", "ja_translate")
        self.update_mode_ui()

        layout.addSpacing(10)

        # --- モデル選択エリア ---
        model_label = QLabel("🤖 AIモデル:")
        model_selector_layout = QHBoxLayout()
        self.model_combo = QComboBox()
        self.model_combo.addItems([
            "gemini-3.1-flash-lite-preview",
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-1.5-flash",
            "gemini-1.5-pro"
        ])
        
        self.update_models_btn = QPushButton("更新")
        self.update_models_btn.setObjectName("UpdateBtn")
        self.update_models_btn.clicked.connect(self.fetch_models_list)
        
        self.recommend_btn = QPushButton("推奨")
        self.recommend_btn.setObjectName("RecommendBtn")
        self.recommend_btn.clicked.connect(self.set_recommended_model)
        
        model_selector_layout.addWidget(self.model_combo, stretch=1)
        model_selector_layout.addWidget(self.update_models_btn)
        model_selector_layout.addWidget(self.recommend_btn)

        saved_model = self.settings.value("model_name", "gemini-3.1-flash-lite-preview")
        index = self.model_combo.findText(saved_model)
        if index >= 0:
            self.model_combo.setCurrentIndex(index)
        
        layout.addWidget(model_label)
        layout.addLayout(model_selector_layout)

        layout.addSpacing(10)

        # --- APIキー設定エリア ---
        api_label = QLabel("🔑 Gemini API キー:")
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_input.setPlaceholderText("APIキーを入力")
        
        saved_key = self.settings.value("api_key", "")
        if saved_key:
            self.api_input.setText(saved_key)

        api_btn_layout = QHBoxLayout()
        self.show_api_key_cb = QCheckBox("表示")
        self.show_api_key_cb.stateChanged.connect(self.toggle_api_key_visibility)
        
        self.test_btn = QPushButton("接続テスト")
        self.test_btn.setObjectName("TestBtn")
        self.test_btn.clicked.connect(self.test_api_key)
        
        save_btn = QPushButton("設定保存")
        save_btn.setObjectName("SaveBtn")
        save_btn.clicked.connect(self.save_settings)
        
        api_btn_layout.addWidget(self.show_api_key_cb)
        api_btn_layout.addStretch()
        api_btn_layout.addWidget(self.test_btn)
        api_btn_layout.addWidget(save_btn)
        
        layout.addWidget(api_label)
        layout.addWidget(self.api_input)
        layout.addLayout(api_btn_layout)

        layout.addStretch()

        # キャプチャ起動ボタン
        self.snip_btn = QPushButton("🔍 画面を切り取って実行")
        self.snip_btn.setObjectName("SnipBtn")
        self.snip_btn.clicked.connect(self.start_snipping)
        layout.addWidget(self.snip_btn)

        self.snipping_tool = None
        self.test_worker = None
        self.model_fetch_worker = None
        self.help_window = None

    def create_menu_bar(self):
        """メニューバーの作成"""
        menubar = self.menuBar()
        help_menu = menubar.addMenu("ヘルプ")
        
        # 使い方アクション
        help_action = QAction("使い方を表示", self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)
        
        # バージョン情報
        about_action = QAction("バージョン情報", self)
        about_action.triggered.connect(lambda: QMessageBox.information(self, "バージョン情報", f"Glans_Miya\nバージョン: {self.version}"))
        help_menu.addAction(about_action)

    def show_help(self):
        """READMEを表示する別ウィンドウを立ち上げる"""
        if self.help_window is None:
            self.help_window = HelpWindow()
        self.help_window.show()
        self.help_window.raise_()

    def on_mode_changed(self):
        """ボタンクリック時にモードを切り替える"""
        sender = self.sender()
        if sender == self.btn_ja:
            self.current_mode = "ja_translate"
        elif sender == self.btn_en:
            self.current_mode = "en_translate"
        elif sender == self.btn_dict:
            self.current_mode = "dictionary"
        self.update_mode_ui()

    def update_mode_ui(self):
        """選択されているモードに応じてボタンの見た目を更新する"""
        self.btn_ja.setProperty("selected", str(self.current_mode == "ja_translate").lower())
        self.btn_en.setProperty("selected", str(self.current_mode == "en_translate").lower())
        self.btn_dict.setProperty("selected", str(self.current_mode == "dictionary").lower())
        
        # スタイルを再適用
        self.btn_ja.style().unpolish(self.btn_ja)
        self.btn_ja.style().polish(self.btn_ja)
        self.btn_en.style().unpolish(self.btn_en)
        self.btn_en.style().polish(self.btn_en)
        self.btn_dict.style().unpolish(self.btn_dict)
        self.btn_dict.style().polish(self.btn_dict)

    def toggle_api_key_visibility(self, state):
        if state == Qt.CheckState.Checked.value:
            self.api_input.setEchoMode(QLineEdit.EchoMode.Normal)
        else:
            self.api_input.setEchoMode(QLineEdit.EchoMode.Password)

    def fetch_models_list(self):
        api_key = self.api_input.text().strip()
        if not api_key:
            QMessageBox.warning(self, "エラー", "APIキーを入力してください。")
            return
        self.update_models_btn.setEnabled(False)
        self.update_models_btn.setText("取得中...")
        self.model_fetch_worker = ModelFetchWorker(api_key)
        self.model_fetch_worker.finished.connect(self.on_models_fetched)
        self.model_fetch_worker.finished.connect(self.model_fetch_worker.deleteLater)
        self.model_fetch_worker.start()

    def on_models_fetched(self, success, models, error_msg):
        self.update_models_btn.setEnabled(True)
        self.update_models_btn.setText("更新")
        if success and models:
            self.model_combo.clear()
            self.model_combo.addItems(models)
            target_model = self.settings.value("model_name", "gemini-3.1-flash-lite-preview")
            index = self.model_combo.findText(target_model)
            if index >= 0: self.model_combo.setCurrentIndex(index)
            else:
                self.model_combo.insertItem(0, "gemini-3.1-flash-lite-preview")
                self.model_combo.setCurrentIndex(0)
            QMessageBox.information(self, "更新完了", "モデル一覧を取得しました。")
        elif not success:
            QMessageBox.critical(self, "取得失敗", f"失敗しました。\n\n詳細:\n{error_msg}")

    def set_recommended_model(self):
        index = self.model_combo.findText("gemini-3.1-flash-lite-preview")
        if index >= 0:
            self.model_combo.setCurrentIndex(index)

    def test_api_key(self):
        api_key = self.api_input.text().strip()
        model_name = self.model_combo.currentText()
        if not api_key:
            QMessageBox.warning(self, "エラー", "APIキーを入力してください。")
            return
        self.test_btn.setEnabled(False)
        self.test_btn.setText("テスト中...")
        self.test_worker = ApiTestWorker(api_key, model_name)
        self.test_worker.finished.connect(self.on_test_finished)
        self.test_worker.finished.connect(self.test_worker.deleteLater)
        self.test_worker.start()

    def on_test_finished(self, success, message):
        self.test_btn.setEnabled(True)
        self.test_btn.setText("接続テスト")
        if success:
            QMessageBox.information(self, "テスト成功", f"APIへの接続に成功しました。")
        else:
            QMessageBox.critical(self, "テスト失敗", f"詳細:\n{message}")

    def save_settings(self):
        api_key = self.api_input.text().strip()
        model_name = self.model_combo.currentText()
        if api_key:
            self.settings.setValue("api_key", api_key)
            self.settings.setValue("model_name", model_name)
            self.settings.setValue("last_mode", self.current_mode)
            self.settings.sync()
            QMessageBox.information(self, "保存完了", "設定を保存しました。")
        else:
            QMessageBox.warning(self, "エラー", "APIキーを入力してください。")

    def closeEvent(self, event):
        QApplication.quit()
        super().closeEvent(event)

    def start_snipping(self):
        api_key = self.api_input.text().strip()
        if not api_key:
            QMessageBox.warning(self, "エラー", "APIキーを設定してください。")
            return
        self.hide()
        # 画面非表示の時間を少し確保してからキャプチャ開始
        QTimer.singleShot(300, self.show_snipping_widget)

    def show_snipping_widget(self):
        self.snipping_tool = SnippingWidget(main_window=self)
        self.snipping_tool.show()


class ModelFetchWorker(QThread):
    finished = pyqtSignal(bool, list, str)
    def __init__(self, api_key):
        super().__init__()
        self.api_key = api_key
    def run(self):
        try:
            genai.configure(api_key=self.api_key)
            models = []
            for m in genai.list_models():
                if 'generateContent' in m.supported_generation_methods:
                    name = m.name.replace('models/', '')
                    if name.startswith('gemini'): models.append(name)
            models.sort(reverse=True)
            self.finished.emit(True, models, "")
        except Exception as e:
            self.finished.emit(False, [], str(e))


class ApiTestWorker(QThread):
    finished = pyqtSignal(bool, str)
    def __init__(self, api_key, model_name):
        super().__init__()
        self.api_key = api_key
        self.model_name = model_name
    def run(self):
        try:
            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel(self.model_name)
            response = model.generate_content("Ping")
            self.finished.emit(True, response.text.strip())
        except Exception as e:
            self.finished.emit(False, str(e))


class ResultWindow(QWidget):
    """
    結果表示ウィンドウ（モードの再選択機能を搭載）
    """
    def __init__(self, original_text, processed_text, image_bytes, api_key, model_name, current_mode):
        super().__init__()
        self.setWindowTitle("Glans_Miya - 結果")
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(500, 550)
        
        # 必要なデータを保持
        self.image_bytes = image_bytes
        self.api_key = api_key
        self.model_name = model_name
        self.current_mode = current_mode
        self.worker = None

        self.setStyleSheet("""
            QWidget { background-color: #ffffff; color: #333333; font-family: 'Segoe UI', Meiryo, sans-serif; }
            QLabel.SectionTitle { font-weight: bold; font-size: 13px; color: #1a73e8; margin-top: 10px; }
            QTextEdit { border: 1px solid #e0e0e0; border-radius: 8px; padding: 8px; font-size: 13px; background-color: #f8f9fa; }
            QPushButton { background-color: #1a73e8; color: white; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; }
            QPushButton:hover { background-color: #1557b0; }
            QPushButton:disabled { background-color: #ccc; }
            
            QPushButton.ModeBtn {
                background-color: #f8f9fa;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 6px;
                font-size: 11px;
                color: #333;
            }
            QPushButton.ModeBtn[selected="true"] {
                background-color: #e8f0fe;
                border: 2px solid #1a73e8;
                color: #1a73e8;
            }
        """)

        layout = QVBoxLayout()

        # オリジナルテキスト
        title_orig = QLabel("📝 読み取ったテキスト:")
        title_orig.setProperty("class", "SectionTitle")
        layout.addWidget(title_orig)
        self.original_edit = QTextEdit()
        self.original_edit.setPlainText(original_text)
        layout.addWidget(self.original_edit)

        # モード切替エリア
        layout.addSpacing(10)
        mode_box = QHBoxLayout()
        mode_box.addWidget(QLabel("✨ モード切替:"))
        
        self.btn_ja = QPushButton("日本語翻訳")
        self.btn_en = QPushButton("英語翻訳")
        self.btn_dict = QPushButton("辞書")
        
        self.mode_buttons = {
            "ja_translate": self.btn_ja,
            "en_translate": self.btn_en,
            "dictionary": self.btn_dict
        }
        
        for key, btn in self.mode_buttons.items():
            btn.setProperty("class", "ModeBtn")
            btn.clicked.connect(lambda checked, k=key: self.reprocess(k))
            mode_box.addWidget(btn)
        
        layout.addLayout(mode_box)

        # 処理結果
        title_proc = QLabel("💡 処理結果:")
        title_proc.setProperty("class", "SectionTitle")
        layout.addWidget(title_proc)
        self.processed_edit = QTextEdit()
        self.processed_edit.setPlainText(processed_text)
        layout.addWidget(self.processed_edit)

        # 閉じるボタン
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.close_btn = QPushButton("閉じる")
        self.close_btn.clicked.connect(self.close)
        btn_layout.addWidget(self.close_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        self.update_mode_ui()

    def update_mode_ui(self):
        """現在のモードに合わせてボタンのスタイルを更新"""
        for key, btn in self.mode_buttons.items():
            btn.setProperty("selected", str(key == self.current_mode).lower())
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    def reprocess(self, new_mode):
        """指定したモードで再処理を行う"""
        if self.current_mode == new_mode and self.processed_edit.toPlainText() != "処理中...":
            return
            
        self.current_mode = new_mode
        self.update_mode_ui()
        self.processed_edit.setPlainText("処理中...")
        
        # ボタンを一時無効化
        for btn in self.mode_buttons.values(): btn.setEnabled(False)

        self.worker = OcrTranslateWorker(self.image_bytes, self.api_key, self.model_name, self.current_mode)
        self.worker.finished.connect(self.on_reprocess_finished)
        self.worker.error.connect(self.on_reprocess_error)
        self.worker.start()

    def on_reprocess_finished(self, original_text, processed_text):
        self.original_edit.setPlainText(original_text)
        self.processed_edit.setPlainText(processed_text)
        for btn in self.mode_buttons.values(): btn.setEnabled(True)

    def on_reprocess_error(self, error_msg):
        self.processed_edit.setPlainText(f"エラーが発生しました:\n{error_msg}")
        for btn in self.mode_buttons.values(): btn.setEnabled(True)


class OcrTranslateWorker(QThread):
    """
    Gemini APIを利用して画像からテキスト抽出と処理を行う
    """
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str)

    def __init__(self, image_bytes, api_key, model_name, mode):
        super().__init__()
        self.image_bytes = image_bytes
        self.api_key = api_key
        self.model_name = model_name
        self.mode = mode

    def run(self):
        try:
            nparr = np.frombuffer(self.image_bytes, np.uint8)
            img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            max_width = 1200
            if gray.shape[1] > max_width:
                scale = max_width / gray.shape[1]
                gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            gray = clahe.apply(gray)
            is_success, buffer = cv2.imencode(".png", gray)
            processed_image_bytes = buffer.tobytes()

            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel(self.model_name)

            if self.mode == "ja_translate":
                role_prompt = "画像内のテキストをすべて抽出し、自然な日本語に翻訳して出力してください。"
            elif self.mode == "en_translate":
                role_prompt = "画像内のテキストをすべて抽出し、自然な英語に翻訳して出力してください。"
            elif self.mode == "dictionary":
                role_prompt = "画像内のメインのテキストを抽出し、その意味、発音(カタカナ)、使い方を辞書のように日本語で詳しく解説してください。"
            else:
                role_prompt = "画像内のテキストを抽出してください。"

            prompt = f"""
            {role_prompt}
            以下のJSON形式で厳密に出力してください。余計な説明は不要です。
            {{
              "original_text": "抽出した元のテキスト全体",
              "processed_text": "翻訳または辞書としての回答内容"
            }}
            """

            image_part = {"mime_type": "image/png", "data": processed_image_bytes}
            generation_config = genai.types.GenerationConfig(response_mime_type="application/json")
            response = model.generate_content([prompt, image_part], generation_config=generation_config)

            try:
                result_json = json.loads(response.text)
                original_text = result_json.get("original_text", "検出なし")
                processed_text = result_json.get("processed_text", "")
            except:
                original_text = "解析エラー"
                processed_text = response.text

            self.finished.emit(original_text, processed_text)

        except Exception as e:
            self.error.emit(f"Gemini API エラー:\n{str(e)}")


class SnippingWidget(QMainWindow):
    """
    画面切り取りオーバーレイ（マルチモニター対応）
    """
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window 
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCursor(Qt.CursorShape.CrossCursor)
        
        # 全画面（マルチモニター対応）の範囲を取得
        total_rect = QRect()
        for screen in QGuiApplication.screens():
            total_rect = total_rect.united(screen.geometry())
        
        # 全画面をキャプチャ（仮想デスクトップ全体）
        primary_screen = QGuiApplication.primaryScreen()
        # 仮想デスクトップの全領域を画像として取得
        self.original_pixmap = primary_screen.grabWindow(
            0, total_rect.x(), total_rect.y(), total_rect.width(), total_rect.height()
        )
        
        # ウィジェットを仮想デスクトップ全体のサイズに設定
        self.setGeometry(total_rect)
        
        self.begin_point = QPoint()
        self.end_point = QPoint()
        self.is_drawing = False

    def paintEvent(self, event):
        painter = QPainter(self)
        # 全画面画像を描画
        painter.drawPixmap(self.rect(), self.original_pixmap)
        # 半透明のオーバーレイを被せる
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100))

        # キャンセル案内の描画（プライマリモニターに表示するように計算）
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        font = painter.font()
        font.setPointSize(12)
        font.setBold(True)
        painter.setFont(font)
        
        guide_text = "ドラッグで範囲指定 / Escキーでキャンセル"
        metrics = painter.fontMetrics()
        tw = metrics.horizontalAdvance(guide_text)
        th = metrics.height()
        
        # プライマリ画面の範囲を取得して、その中での相対位置を計算
        primary_rect = QGuiApplication.primaryScreen().geometry()
        # ウィジェットの原点（仮想デスクトップ全体の左上）からの相対座標に変換
        local_primary = primary_rect.translated(-self.geometry().topLeft())
        
        # プライマリ画面内での中央下部に配置
        bg_x = local_primary.left() + (local_primary.width() - tw) // 2 - 20
        bg_y = local_primary.bottom() - 100
        rect_bg = QRect(bg_x, bg_y, tw + 40, th + 20)

        painter.setBrush(QColor(0, 0, 0, 150))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(rect_bg, 10, 10)
        
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(rect_bg, Qt.AlignmentFlag.AlignCenter, guide_text)

        # 選択範囲の描画
        if not self.begin_point.isNull() and not self.end_point.isNull():
            rect = QRect(self.begin_point, self.end_point).normalized()
            # 選択された部分だけ元の画像を上書き（明るく見える）
            painter.drawPixmap(rect, self.original_pixmap.copy(rect))
            painter.setPen(QPen(QColor(26, 115, 232), 2))
            painter.drawRect(rect)

    def keyPressEvent(self, event):
        """Escキーが押されたらキャンセルして閉じる"""
        if event.key() == Qt.Key.Key_Escape:
            self.close()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.begin_point = event.pos()
            self.end_point = self.begin_point
            self.is_drawing = True
            self.update()

    def mouseMoveEvent(self, event):
        if self.is_drawing:
            self.end_point = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.is_drawing = False
            rect = QRect(self.begin_point, self.end_point).normalized()
            if rect.width() <= 10 or rect.height() <= 10:
                self.close()
                return
            self.process_selected_area(rect)

    def process_selected_area(self, rect: QRect):
        self.setCursor(Qt.CursorShape.WaitCursor)
        # 切り抜かれたピックスマップをバイトデータに変換
        cropped_pixmap = self.original_pixmap.copy(rect)
        buffer = QBuffer()
        buffer.open(QIODevice.OpenModeFlag.ReadWrite)
        cropped_pixmap.save(buffer, "PNG")
        self.last_image_bytes = buffer.data().data()

        api_key = self.main_window.api_input.text().strip()
        model_name = self.main_window.model_combo.currentText()
        mode = self.main_window.current_mode
        
        self.worker = OcrTranslateWorker(self.last_image_bytes, api_key, model_name, mode)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.error.connect(self.on_worker_error)
        self.worker.start()

    def on_worker_finished(self, original_text, processed_text):
        self.setCursor(Qt.CursorShape.CrossCursor)
        # ResultWindowへデータを渡す
        api_key = self.main_window.api_input.text().strip()
        model_name = self.main_window.model_combo.currentText()
        mode = self.main_window.current_mode
        
        self.main_window.result_window = ResultWindow(
            original_text, processed_text, self.last_image_bytes, api_key, model_name, mode
        )
        self.main_window.result_window.show()
        self.close()

    def on_worker_error(self, error_msg):
        self.setCursor(Qt.CursorShape.CrossCursor)
        self.main_window.result_window = ResultWindow("エラー", error_msg, None, "", "", "")
        self.main_window.result_window.show()
        self.close()

    def closeEvent(self, event):
        if self.main_window: self.main_window.showNormal()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Glans_Miya")
    app.setQuitOnLastWindowClosed(False)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()