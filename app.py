import sys
import os
import time
import json
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTextEdit, QLabel, QMainWindow, QLineEdit, 
                             QMessageBox, QComboBox, QCheckBox, QMenuBar, QMenu)
from PyQt6.QtCore import Qt, QRect, QPoint, QTimer, QThread, pyqtSignal, QBuffer, QIODevice, QSettings
from PyQt6.QtGui import QPainter, QColor, QPen, QGuiApplication, QScreen, QPixmap, QIcon, QAction

import google.generativeai as genai

def resource_path(relative_path):
    """ PyInstallerの一時フォルダまたは実行フォルダから絶対パスを取得する """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class HelpWindow(QWidget):
    """ README.md の内容を表示するウィンドウ """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Glans_Miya - 使い方")
        self.resize(500, 600)
        self.setWindowFlags(Qt.WindowType.Window)
        layout = QVBoxLayout()
        self.text_browser = QTextEdit()
        self.text_browser.setReadOnly(True)
        readme_path = resource_path("README.md")
        if os.path.exists(readme_path):
            with open(readme_path, "r", encoding="utf-8") as f:
                self.text_browser.setPlainText(f.read())
        else:
            self.text_browser.setPlainText("README.md が見つかりませんでした。")
        layout.addWidget(self.text_browser)
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        self.setLayout(layout)

class ResultWindow(QWidget):
    """ 結果表示ウィンドウ（モードの再選択機能を搭載） """
    def __init__(self, original_text, processed_text, image_bytes, api_key, model_name, current_mode):
        super().__init__()
        self.setWindowTitle("Glans_Miya v1.0.2 - 結果")
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(500, 550)
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
            QPushButton.ModeBtn { background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 4px; padding: 6px; font-size: 11px; color: #333; }
            QPushButton.ModeBtn[selected="true"] { background-color: #e8f0fe; border: 2px solid #1a73e8; color: #1a73e8; }
        """)
        layout = QVBoxLayout()
        layout.addWidget(QLabel("📝 読み取ったテキスト:", objectName="SectionTitle"))
        self.original_edit = QTextEdit()
        self.original_edit.setPlainText(original_text)
        layout.addWidget(self.original_edit)
        
        mode_box = QHBoxLayout()
        mode_box.addWidget(QLabel("✨ モード切替:"))
        self.mode_buttons = {}
        for k, n in [("ja_translate", "日本語翻訳"), ("en_translate", "英語翻訳"), ("dictionary", "辞書")]:
            btn = QPushButton(n)
            btn.setProperty("class", "ModeBtn")
            btn.clicked.connect(lambda checked, mode=k: self.reprocess(mode))
            self.mode_buttons[k] = btn
            mode_box.addWidget(btn)
        layout.addLayout(mode_box)
        
        layout.addWidget(QLabel("💡 処理結果:", objectName="SectionTitle"))
        self.processed_edit = QTextEdit()
        self.processed_edit.setPlainText(processed_text)
        layout.addWidget(self.processed_edit)
        
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        self.setLayout(layout)
        self.update_mode_ui()

    def update_mode_ui(self):
        for k, btn in self.mode_buttons.items():
            btn.setProperty("selected", str(k == self.current_mode).lower())
            btn.style().unpolish(btn); btn.style().polish(btn)

    def reprocess(self, new_mode):
        if self.current_mode == new_mode and self.processed_edit.toPlainText() != "処理中...":
            return
        self.current_mode = new_mode
        self.update_mode_ui()
        self.processed_edit.setPlainText("処理中...")
        self.worker = OcrTranslateWorker(self.image_bytes, self.api_key, self.model_name, self.current_mode)
        self.worker.finished.connect(lambda orig, proc: (self.original_edit.setPlainText(orig), self.processed_edit.setPlainText(proc)))
        self.worker.error.connect(lambda msg: self.processed_edit.setPlainText(f"エラーが発生しました:\n{msg}"))
        self.worker.start()

class OcrTranslateWorker(QThread):
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
            # OpenCVを使用せず、PyQtで生成されたPNGバイトデータをそのまま使用
            image_data = self.image_bytes

            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel(self.model_name)
            prompts = {
                "ja_translate": "画像内のテキストをすべて抽出し、自然な日本語に翻訳してください。",
                "en_translate": "画像内のテキストをすべて抽出し、自然な英語に翻訳してください。",
                "dictionary": "画像内のメインテキストを抽出し、意味・発音・使い方を辞書的に日本語で解説してください。"
            }
            prompt = f"{prompts.get(self.mode, '')}\n必ず以下のJSON形式のみで出力してください。余計な文章は一切含めないでください: {{'original_text': '原文', 'processed_text': '結果'}}"
            
            response = model.generate_content(
                [prompt, {"mime_type": "image/png", "data": image_data}], 
                generation_config=genai.types.GenerationConfig(response_mime_type="application/json")
            )
            
            # JSON解析 (Markdownタグが含まれる場合を考慮)
            raw_text = response.text.strip()
            if raw_text.startswith("```"):
                lines = raw_text.splitlines()
                if len(lines) >= 2:
                    raw_text = "\n".join(lines[1:-1])
            
            res = json.loads(raw_text)
            self.finished.emit(res.get("original_text", ""), res.get("processed_text", ""))
        except Exception as e:
            self.error.emit(str(e))

class ModelFetchWorker(QThread):
    """ 利用可能なモデル一覧を取得するスレッド """
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
    """ APIの疎通確認を行うスレッド """
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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.version = "v1.0.2"
        self.setWindowTitle(f"Glans_Miya {self.version}")
        # 初期状態は最前面に設定。setup_ui 内の設定読み込みで上書きされます。
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(360, 480)
        self.setStyleSheet("""
            QMainWindow, QWidget { background-color: #ffffff; color: #333333; font-family: 'Segoe UI', Meiryo, sans-serif; }
            QLineEdit, QComboBox { padding: 8px; border: 1px solid #ccc; border-radius: 4px; }
            QPushButton#SnipBtn { background-color: #1a73e8; color: white; border-radius: 6px; padding: 12px; font-weight: bold; }
            QPushButton.ModeBtn { background-color: #f8f9fa; border: 1px solid #ddd; padding: 10px; border-radius: 4px; font-size: 11px; }
            QPushButton.ModeBtn[selected="true"] { background-color: #e8f0fe; border: 2px solid #1a73e8; color: #1a73e8; }
            QPushButton#ActionBtn { background-color: #f1f3f4; border: 1px solid #ccc; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
            QLabel#StatusLabel { font-size: 11px; color: #666; padding: 2px; }
        """)
        self.active_results = [] 
        self.snipping_tool = None 
        self.worker = None
        self.setup_ui()

    def setup_ui(self):
        menubar = self.menuBar()
        help_menu = menubar.addMenu("ヘルプ")
        help_menu.addAction("使い方を表示", self.show_help)
        help_menu.addAction("バージョン情報", lambda: QMessageBox.information(self, "情報", f"Glans_Miya {self.version}"))

        central = QWidget(); self.setCentralWidget(central); layout = QVBoxLayout(central)
        layout.addWidget(QLabel("<b>【使い方】</b><br>1. モードを選択しAPIキーを設定<br>2. 画面を切り取って実行"))
        
        mode_layout = QHBoxLayout()
        self.mode_btns = {}
        for k, n in [("ja_translate", "日本語翻訳"), ("en_translate", "英語翻訳"), ("dictionary", "辞書")]:
            btn = QPushButton(n); btn.setProperty("class", "ModeBtn"); btn.clicked.connect(lambda _, m=k: self.set_mode(m))
            self.mode_btns[k] = btn; mode_layout.addWidget(btn)
        layout.addLayout(mode_layout)

        layout.addWidget(QLabel("🤖 AIモデル:"))
        model_box = QHBoxLayout()
        self.model_combo = QComboBox()
        self.model_combo.addItems(["gemini-3.1-flash-lite-preview", "gemini-1.5-flash"])
        self.update_btn = QPushButton("更新"); self.update_btn.setObjectName("ActionBtn"); self.update_btn.clicked.connect(self.fetch_models_list)
        self.reco_btn = QPushButton("推奨"); self.reco_btn.setObjectName("ActionBtn"); self.reco_btn.clicked.connect(self.set_recommended_model)
        model_box.addWidget(self.model_combo, stretch=1); model_box.addWidget(self.update_btn); model_box.addWidget(self.reco_btn)
        layout.addLayout(model_box)

        layout.addWidget(QLabel("🔑 Gemini API キー:"))
        self.api_input = QLineEdit(); self.api_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.api_input)
        
        btn_box = QHBoxLayout()
        self.show_cb = QCheckBox("表示")
        self.show_cb.stateChanged.connect(lambda s: self.api_input.setEchoMode(QLineEdit.EchoMode.Normal if s == 2 else QLineEdit.EchoMode.Password))
        
        # 最前面表示の切り替えボタン（チェックボックス）
        self.ontop_cb = QCheckBox("最前面")
        self.ontop_cb.clicked.connect(self.toggle_always_on_top)
        
        self.test_btn = QPushButton("テスト"); self.test_btn.setObjectName("ActionBtn"); self.test_btn.clicked.connect(self.test_api_key)
        save_btn = QPushButton("設定保存"); save_btn.setObjectName("ActionBtn"); save_btn.clicked.connect(self.save_settings)
        
        btn_box.addWidget(self.show_cb)
        btn_box.addWidget(self.ontop_cb)
        btn_box.addStretch()
        btn_box.addWidget(self.test_btn)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

        # ステータス表示ラベル
        self.status_label = QLabel("待機中", objectName="StatusLabel")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()
        layout.addWidget(self.status_label)

        snip_btn = QPushButton("🔍 画面を切り取って実行", objectName="SnipBtn"); snip_btn.clicked.connect(self.start_snipping)
        layout.addWidget(snip_btn)

        # 設定読み込み
        self.settings = QSettings(os.path.join(os.path.expanduser('~'), 'Glans_Miya', 'settings.ini'), QSettings.Format.IniFormat)
        self.api_input.setText(self.settings.value("api_key", ""))
        saved_model = self.settings.value("model_name", "gemini-3.1-flash-lite-preview")
        idx = self.model_combo.findText(saved_model)
        if idx >= 0: self.model_combo.setCurrentIndex(idx)
        
        # 最前面表示の設定反映
        is_ontop = self.settings.value("always_on_top", "true").lower() == "true"
        self.ontop_cb.setChecked(is_ontop)
        if not is_ontop:
            self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.Tool)
            self.show()

        self.current_mode = self.settings.value("last_mode", "ja_translate")
        self.set_mode(self.current_mode)

    def toggle_always_on_top(self):
        """ ウィンドウの最前面表示を切り替える """
        is_checked = self.ontop_cb.isChecked()
        if is_checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
        # フラグ変更を適用するために再表示が必要
        self.show()

    def set_mode(self, mode):
        self.current_mode = mode
        for k, b in self.mode_btns.items():
            b.setProperty("selected", str(k == mode).lower())
            b.style().unpolish(b); b.style().polish(b)

    def fetch_models_list(self):
        api_key = self.api_input.text().strip()
        if not api_key: QMessageBox.warning(self, "エラー", "APIキーを入力してください"); return
        self.update_btn.setEnabled(False); self.update_btn.setText("...")
        self.mfw = ModelFetchWorker(api_key)
        self.mfw.finished.connect(self.on_models_fetched)
        self.mfw.start()

    def on_models_fetched(self, success, models, err):
        self.update_btn.setEnabled(True); self.update_btn.setText("更新")
        if success:
            self.model_combo.clear(); self.model_combo.addItems(models)
            QMessageBox.information(self, "完了", "モデル一覧を更新しました。")
        else: QMessageBox.critical(self, "エラー", f"失敗しました: {err}")

    def set_recommended_model(self):
        idx = self.model_combo.findText("gemini-3.1-flash-lite-preview")
        if idx >= 0: self.model_combo.setCurrentIndex(idx)

    def test_api_key(self):
        api_key = self.api_input.text().strip()
        if not api_key: QMessageBox.warning(self, "エラー", "APIキーを入力してください"); return
        self.test_btn.setEnabled(False); self.test_btn.setText("...")
        self.taw = ApiTestWorker(api_key, self.model_combo.currentText())
        self.taw.finished.connect(self.on_test_finished)
        self.taw.start()

    def on_test_finished(self, success, msg):
        self.test_btn.setEnabled(True); self.test_btn.setText("テスト")
        if success: QMessageBox.information(self, "成功", "API接続に成功しました。")
        else: QMessageBox.critical(self, "失敗", f"接続エラー: {msg}")

    def save_settings(self):
        self.settings.setValue("api_key", self.api_input.text())
        self.settings.setValue("model_name", self.model_combo.currentText())
        self.settings.setValue("last_mode", self.current_mode)
        self.settings.setValue("always_on_top", str(self.ontop_cb.isChecked()).lower())
        QMessageBox.information(self, "完了", "設定を保存しました。")

    def show_help(self):
        self.hw = HelpWindow(); self.hw.show()

    def start_snipping(self):
        if not self.api_input.text(): QMessageBox.warning(self, "エラー", "APIキーを入力してください"); return
        self.hide()
        QTimer.singleShot(300, self.init_snipping)

    def init_snipping(self):
        self.status_label.setText("範囲を選択してください...")
        self.status_label.setStyleSheet("color: #666;")
        self.snipping_tool = SnippingWidget(self)
        self.snipping_tool.show()

    def process_captured_image(self, img_data):
        api = self.api_input.text()
        model = self.model_combo.currentText()
        mode = self.current_mode
        
        # ステータス更新
        self.status_label.setText("⏳ AI処理中... しばらくお待ちください")
        self.status_label.setStyleSheet("color: #1a73e8; font-weight: bold;")
        
        self.worker = OcrTranslateWorker(img_data, api, model, mode)
        self.worker.finished.connect(lambda o, p: self.show_result(o, p, img_data, api, model, mode))
        self.worker.error.connect(self.handle_processing_error)
        self.worker.start()

    def handle_processing_error(self, message):
        self.status_label.setText("❌ エラーが発生しました")
        self.status_label.setStyleSheet("color: #d93025;")
        QMessageBox.critical(self, "処理エラー", f"AI処理中にエラーが発生しました。\n\n詳細: {message}")

    def show_result(self, original_text, processed_text, img_data, api, model, mode):
        self.status_label.setText("✅ 処理完了")
        self.status_label.setStyleSheet("color: #188038;")
        
        res_win = ResultWindow(original_text, processed_text, img_data, api, model, mode)
        self.active_results.append(res_win) 
        res_win.show()
        res_win.raise_()
        res_win.activateWindow()
        
        # 3秒後にステータスを待機中に戻す
        QTimer.singleShot(3000, lambda: self.status_label.setText("待機中"))
        QTimer.singleShot(3000, lambda: self.status_label.setStyleSheet("color: #666;"))

    def closeEvent(self, event):
        """ メインウィンドウが閉じられたときに確実にプロセスを終了させる """
        QApplication.quit()
        super().closeEvent(event)

class SnippingWidget(QMainWindow):
    def __init__(self, main_win):
        super().__init__()
        self.main_win = main_win
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCursor(Qt.CursorShape.CrossCursor)
        
        rect = QRect()
        for s in QGuiApplication.screens(): rect = rect.united(s.geometry())
        
        self.original_pixmap = QGuiApplication.primaryScreen().grabWindow(0, rect.x(), rect.y(), rect.width(), rect.height())
        self.setGeometry(rect)
        self.begin, self.end = QPoint(), QPoint()

    def paintEvent(self, _):
        p = QPainter(self); p.drawPixmap(self.rect(), self.original_pixmap)
        p.fillRect(self.rect(), QColor(0, 0, 0, 100))
        
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        guide = "ドラッグで範囲指定 / Escでキャンセル"
        tw = p.fontMetrics().horizontalAdvance(guide)
        pr = QGuiApplication.primaryScreen().geometry()
        local_pr = pr.translated(-self.geometry().topLeft())
        rect_bg = QRect(local_pr.left() + (local_pr.width()-tw)//2-10, local_pr.bottom()-80, tw+20, 30)
        p.setBrush(QColor(0,0,0,150)); p.setPen(Qt.PenStyle.NoPen); p.drawRoundedRect(rect_bg, 5, 5)
        p.setPen(QColor(255,255,255)); p.drawText(rect_bg, Qt.AlignmentFlag.AlignCenter, guide)

        if not self.begin.isNull() and not self.end.isNull():
            r = QRect(self.begin, self.end).normalized()
            p.drawPixmap(r, self.original_pixmap.copy(r))
            p.setPen(QPen(QColor(26, 115, 232), 2)); p.drawRect(r)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key.Key_Escape:
            self.close()
            self.main_win.showNormal()

    def mousePressEvent(self, e): 
        self.begin = e.pos()
        self.end = self.begin
        self.update()

    def mouseMoveEvent(self, e): 
        self.end = e.pos()
        self.update()

    def mouseReleaseEvent(self, e):
        r = QRect(self.begin, e.pos()).normalized()
        if r.width() > 10 and r.height() > 10:
            buf = QBuffer(); buf.open(QIODevice.OpenModeFlag.ReadWrite)
            self.original_pixmap.copy(r).save(buf, "PNG")
            self.main_win.process_captured_image(buf.data().data())
        
        self.close()
        self.main_win.showNormal()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())