import sys
import os
import time
import json
import re
import asyncio
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTextEdit, QLabel, QMainWindow, QLineEdit, 
                             QMessageBox, QComboBox, QCheckBox, QMenuBar, QMenu,
                             QDialog, QTabWidget, QGroupBox, QRadioButton, QSpinBox,
                             QDoubleSpinBox, QListWidget, QListWidgetItem, QAbstractItemView,
                             QButtonGroup, QFrame)
from PyQt6.QtCore import Qt, QRect, QPoint, QTimer, QThread, pyqtSignal, QBuffer, QIODevice, QSettings, QUrl
from PyQt6.QtGui import QPainter, QColor, QPen, QGuiApplication, QScreen, QPixmap, QIcon, QAction, QDesktopServices

import google.generativeai as genai

# Windows ローカルOCR用のライブラリをインポート
try:
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    WIN_OCR_AVAILABLE = True
except ImportError:
    WIN_OCR_AVAILABLE = False

# OCRエンジンをキャッシュして使い回すためのグローバル変数
_ocr_engine = None

async def perform_local_ocr(image_bytes):
    """ Windows標準のOCR機能（WinRT）を使用して画像から瞬時にテキストを抽出する """
    global _ocr_engine
    try:
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(bytearray(image_bytes))
        await writer.store_async()
        await writer.flush_async()
        writer.detach_stream()
        
        stream.seek(0)
        
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        
        # エンジンを毎回生成せず、キャッシュして高速化する
        if _ocr_engine is None:
            _ocr_engine = OcrEngine.try_create_from_user_profile_languages()
            
        if _ocr_engine is None:
            return ""
            
        result = await _ocr_engine.recognize_async(bitmap)
        return result.text
    except Exception as e:
        print(f"Local OCR Error: {e}")
        return ""

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
        self.setWindowTitle("MiyashitaLens - 使い方")
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

class SettingsWindow(QDialog):
    """ API詳細設定画面 """
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("AI詳細設定 (Gemini API)")
        self.resize(900, 650)
        self.setStyleSheet("""
            QDialog { background-color: #f8f9fa; font-family: 'Segoe UI', Meiryo, sans-serif; color: #333; }
            QGroupBox { font-weight: bold; color: #1a73e8; border: 1px solid #ccc; border-radius: 5px; margin-top: 10px; padding-top: 15px; background-color: white;}
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
            QPushButton { background-color: #e8eaed; border: 1px solid #dadce0; border-radius: 4px; padding: 6px 12px; color: #3c4043; }
            QPushButton:hover { background-color: #f1f3f4; }
            QPushButton#PrimaryBtn { background-color: #1a73e8; color: white; border: none; font-weight: bold; }
            QPushButton#PrimaryBtn:hover { background-color: #1557b0; }
            QLineEdit, QComboBox, QListWidget { border: 1px solid #ccc; border-radius: 4px; padding: 5px; background-color: white;}
            QTabWidget::pane { border: 1px solid #ccc; background-color: white; border-radius: 4px;}
            QTabBar::tab { background: #e8eaed; border: 1px solid #ccc; padding: 8px 20px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background: white; border-bottom-color: white; font-weight: bold; color: #1a73e8; }
            QLabel { font-size: 12px; }
        """)

        self.tab_ui = {"free": {}, "paid": {}}
        self.workers = []

        layout = QVBoxLayout(self)

        title = QLabel("Gemini API 詳細設定")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #1a73e8; margin-bottom: 10px;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        plan_group = QGroupBox("実行プランの選択")
        plan_group.setStyleSheet("QGroupBox { background-color: #e8f0fe; border-color: #1a73e8; } QGroupBox::title { color: #1a73e8; }")
        plan_layout = QHBoxLayout()
        
        desc_label = QLabel("実際に抽出で使用するプランを選んでください：")
        desc_label.setStyleSheet("font-weight: bold; color: #333;")
        plan_layout.addWidget(desc_label)
        
        self.plan_combo = QComboBox()
        self.plan_combo.addItems(["🟢 無料枠 (Free Tier)", "🟠 課金枠 (Paid Tier)"])
        self.plan_combo.setCursor(Qt.CursorShape.PointingHandCursor)
        self.plan_combo.setStyleSheet("""
            QComboBox { font-weight: bold; font-size: 13px; color: #1a73e8; padding: 6px 15px; border: 1px solid #ccc; border-radius: 4px; background-color: white; min-width: 150px; }
            QComboBox::drop-down { border-left: 1px solid #ccc; width: 20px; }
        """)
        
        plan_layout.addWidget(self.plan_combo)
        plan_layout.addStretch()
        
        plan_group.setLayout(plan_layout)
        layout.addWidget(plan_group)

        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_tab_content("free"), "🟢 無料枠 (Free Tier) の設定")
        self.tabs.addTab(self.create_tab_content("paid"), "🟠 課金枠 (Paid Tier) の設定")
        layout.addWidget(self.tabs)

        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("キャンセル")
        btn_cancel.clicked.connect(self.reject)
        btn_save = QPushButton("設定を適用して閉じる", objectName="PrimaryBtn")
        btn_save.clicked.connect(self.save_settings)
        btn_layout.addWidget(btn_cancel)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        layout.addLayout(btn_layout)

        self.load_settings()

    def create_tab_content(self, tab_key):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        ui = self.tab_ui[tab_key]

        # ① APIキー
        g1 = QGroupBox("① APIキー")
        l1 = QHBoxLayout()
        l1.addWidget(QLabel(f"{'Free' if tab_key == 'free' else 'Paid'} 用のAPIキー:"))
        ui["api_key"] = QLineEdit()
        ui["api_key"].setEchoMode(QLineEdit.EchoMode.Password)
        l1.addWidget(ui["api_key"], stretch=1)
        btn_confirm = QPushButton("確認")
        btn_confirm.clicked.connect(lambda _, le=ui["api_key"]: le.setEchoMode(QLineEdit.EchoMode.Normal if le.echoMode() == QLineEdit.EchoMode.Password else QLineEdit.EchoMode.Password))
        l1.addWidget(btn_confirm)
        btn_test = QPushButton("テスト")
        btn_test.clicked.connect(lambda _, tk=tab_key: self.test_api_key(tk))
        l1.addWidget(btn_test)
        g1.setLayout(l1)
        layout.addWidget(g1)

        h23 = QHBoxLayout()
        # ② モデル
        g2 = QGroupBox("② モデル・パフォーマンス設定")
        l2 = QVBoxLayout()
        l2_top = QHBoxLayout()
        l2_top.addWidget(QLabel("使用モデル:"))
        ui["model"] = QComboBox()
        ui["model"].addItems(["gemini-3.1-flash-lite-preview", "gemini-1.5-flash", "gemini-1.5-pro"])
        ui["model"].setEditable(True)
        l2_top.addWidget(ui["model"], stretch=1)
        btn_update = QPushButton("🌐 更新")
        btn_update.clicked.connect(lambda _, tk=tab_key: self.fetch_models(tk))
        l2_top.addWidget(btn_update)
        l2.addLayout(l2_top)

        link_label = QLabel("<a href='https://ai.google.dev/models/gemini' style='color:#1a73e8; text-decoration:none;'>🔗 各モデルの特徴を確認する(公式)</a>")
        link_label.setOpenExternalLinks(True)
        link_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        l2.addWidget(link_label)

        l2_mid = QHBoxLayout()
        l2_mid.addWidget(QLabel("RPM:"))
        ui["rpm"] = QSpinBox(); ui["rpm"].setMaximum(9999); ui["rpm"].setValue(15)
        l2_mid.addWidget(ui["rpm"])
        l2_mid.addWidget(QLabel("スレッド:"))
        ui["threads"] = QSpinBox(); ui["threads"].setMinimum(1); ui["threads"].setMaximum(10); ui["threads"].setValue(1)
        l2_mid.addWidget(ui["threads"])
        l2_mid.addStretch()
        l2.addLayout(l2_mid)

        l2_bot = QHBoxLayout()
        btn_limit = QPushButton("ℹ️ 制限と仕様を確認")
        btn_limit.clicked.connect(self.show_api_limits_summary)
        l2_bot.addWidget(btn_limit)
        l2_bot.addStretch()
        btn_reco_model = QPushButton("🔄 推奨値")
        btn_reco_model.clicked.connect(lambda _, tk=tab_key: self.set_reco_model(tk))
        l2_bot.addWidget(btn_reco_model)
        l2.addLayout(l2_bot)
        g2.setLayout(l2)
        h23.addWidget(g2, stretch=2)

        # ③ パラメータ
        g3 = QGroupBox("③ AI抽出パラメータ設定")
        l3 = QVBoxLayout()
        l3_top = QHBoxLayout()
        l3_top.addWidget(QLabel("Temp:"))
        ui["temp"] = QDoubleSpinBox(); ui["temp"].setDecimals(1); ui["temp"].setSingleStep(0.1); ui["temp"].setMaximum(2.0)
        l3_top.addWidget(ui["temp"])
        l3_top.addWidget(QLabel("最大トークン:"))
        ui["tokens"] = QSpinBox(); ui["tokens"].setMaximum(1000000); ui["tokens"].setSingleStep(1024)
        l3_top.addWidget(ui["tokens"])
        l3.addLayout(l3_top)

        ui["safety"] = QCheckBox("安全フィルタ無効化 (エラー回避)")
        l3.addWidget(ui["safety"])
        
        l3_bot = QHBoxLayout()
        l3_bot.addStretch()
        btn_reco_param = QPushButton("🔄 推奨値")
        btn_reco_param.clicked.connect(lambda _, tk=tab_key: self.set_reco_param(tk))
        l3_bot.addWidget(btn_reco_param)
        l3.addLayout(l3_bot)
        g3.setLayout(l3)
        h23.addWidget(g3, stretch=1)

        layout.addLayout(h23)

        # ④ 追加指示
        g4 = QGroupBox("④ 独自の追加指示 (カスタムプロンプト) - 任意")
        l4 = QVBoxLayout()
        l4_top = QHBoxLayout()
        ui["prompt_input"] = QLineEdit()
        ui["prompt_input"].setPlaceholderText("例: 専門用語は分かりやすく噛み砕いて説明してください。")
        l4_top.addWidget(ui["prompt_input"], stretch=1)
        btn_add_prompt = QPushButton("+ 指示を追加", objectName="PrimaryBtn")
        btn_add_prompt.clicked.connect(lambda _, tk=tab_key: self.add_prompt(tk))
        l4_top.addWidget(btn_add_prompt)
        l4.addLayout(l4_top)

        l4_mid = QHBoxLayout()
        v_curr = QVBoxLayout()
        v_curr.addWidget(QLabel("▼ 現在の抽出に使用する指示"))
        ui["list_curr"] = QListWidget()
        v_curr.addWidget(ui["list_curr"])
        h_curr_btn = QHBoxLayout()
        btn_del_curr = QPushButton("🗑 選択を削除")
        btn_del_curr.clicked.connect(lambda _, tk=tab_key: self.remove_item(ui["list_curr"]))
        h_curr_btn.addWidget(btn_del_curr)
        btn_fav = QPushButton("⭐ 選択をお気に入りに保存")
        btn_fav.clicked.connect(lambda _, tk=tab_key: self.save_fav(tk))
        h_curr_btn.addWidget(btn_fav)
        v_curr.addLayout(h_curr_btn)
        l4_mid.addLayout(v_curr)

        v_fav = QVBoxLayout()
        lbl_fav = QLabel("⭐ お気に入り (よく使う指示)")
        lbl_fav.setStyleSheet("color: #fbbc04; font-weight: bold;")
        v_fav.addWidget(lbl_fav)
        ui["list_fav"] = QListWidget()
        v_fav.addWidget(ui["list_fav"])
        h_fav_btn = QHBoxLayout()
        btn_add_left = QPushButton("◀ 選択を左に追加", objectName="PrimaryBtn")
        btn_add_left.clicked.connect(lambda _, tk=tab_key: self.add_to_curr(tk))
        h_fav_btn.addWidget(btn_add_left)
        btn_del_fav = QPushButton("🗑 選択を削除")
        btn_del_fav.clicked.connect(lambda _, tk=tab_key: self.remove_item(ui["list_fav"]))
        h_fav_btn.addWidget(btn_del_fav)
        v_fav.addLayout(h_fav_btn)
        l4_mid.addLayout(v_fav)

        l4.addLayout(l4_mid)
        g4.setLayout(l4)
        layout.addWidget(g4)

        return tab

    def show_api_limits_summary(self):
        """ ドキュメント内容に基づいた制限と仕様の要約ダイアログを表示する """
        dialog = QDialog(self)
        dialog.setWindowTitle("Gemini API 制限と仕様 (要約)")
        dialog.resize(550, 450)
        dialog.setStyleSheet("""
            QDialog { background-color: #f8f9fa; color: #333; }
            QTextEdit { border: 1px solid #ccc; border-radius: 4px; padding: 10px; font-size: 13px; background-color: white; line-height: 1.5; }
            QPushButton { background-color: #1a73e8; color: white; border: none; border-radius: 4px; padding: 8px 16px; font-weight: bold; }
            QPushButton:hover { background-color: #1557b0; }
        """)
        layout = QVBoxLayout(dialog)
        
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        summary_text = """<b>【Gemini API 制限と仕様の要約 (2026年最新版)】</b><br><br>

<b>■ 無料枠 (Free Tier) の厳格な制限</b><br>
・現在、無料枠のリクエスト制限は非常に厳しく設定されています（例: Gemini 2.5 Proで約2 RPM / 50 RPD）。<br>
・連続してリクエストを送ったり並列処理を行うと、すぐにAPI制限エラー（429 Resource Exhausted）が発生します。<br>
・入力したデータは、Googleの製品改善やモデル学習に利用される可能性があります。<br><br>

<b>■ 課金枠 (Paid Tier / Tier 1) のメリット</b><br>
・課金設定（クレジットカード登録）を行うだけで即時アップグレードされ、実用的な水準（例: 3.1 Flash-Liteで300 RPM / 1500 RPDなど）に制限が大幅に引き上げられます。<br>
・入力データがGoogleの学習に利用されなくなるため、プライバシー保護や業務利用における必須条件となります。<br>
・バッチ処理やコンテキストキャッシングといった、コスト・負荷を削減する高度な機能が利用可能になります。<br><br>

<b>■ 同時実行（スレッド処理）に関する注意点</b><br>
・1分間の上限（RPM）内であっても、非同期処理等で一度に大量のリクエストを同時送信すると「バースト制限」に引っかかりエラーとなります。<br>
・安定して高速処理を行うには、スレッド数を絞るか、システム側で適切な待機時間（スリープ）やエラー時の再試行（指数的バックオフ）を設定する必要があります。"""
        text_edit.setHtml(summary_text)
        layout.addWidget(text_edit)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        
        dialog.exec()

    def add_prompt(self, tab_key):
        text = self.tab_ui[tab_key]["prompt_input"].text().strip()
        if text:
            self.tab_ui[tab_key]["list_curr"].addItem(text)
            self.tab_ui[tab_key]["prompt_input"].clear()

    def remove_item(self, list_widget):
        for item in list_widget.selectedItems():
            list_widget.takeItem(list_widget.row(item))

    def save_fav(self, tab_key):
        for item in self.tab_ui[tab_key]["list_curr"].selectedItems():
            text = item.text()
            items = [self.tab_ui[tab_key]["list_fav"].item(i).text() for i in range(self.tab_ui[tab_key]["list_fav"].count())]
            if text not in items:
                self.tab_ui[tab_key]["list_fav"].addItem(text)

    def add_to_curr(self, tab_key):
        for item in self.tab_ui[tab_key]["list_fav"].selectedItems():
            text = item.text()
            self.tab_ui[tab_key]["list_curr"].addItem(text)

    def test_api_key(self, tab_key):
        api_key = self.tab_ui[tab_key]["api_key"].text().strip()
        model = self.tab_ui[tab_key]["model"].currentText()
        if not api_key:
            QMessageBox.warning(self, "エラー", "APIキーを入力してください")
            return
        worker = ApiTestWorker(api_key, model)
        worker.finished.connect(self.on_test_finished)
        self.workers.append(worker)
        worker.start()

    def on_test_finished(self, success, msg):
        if success:
            QMessageBox.information(self, "成功", msg)
        else:
            QMessageBox.critical(self, "失敗", msg)

    def fetch_models(self, tab_key):
        api_key = self.tab_ui[tab_key]["api_key"].text().strip()
        if not api_key:
            QMessageBox.warning(self, "エラー", "APIキーを入力してください")
            return
        worker = ModelFetchWorker(api_key)
        worker.finished.connect(lambda s, m, e, tk=tab_key: self.on_fetch_models_finished(s, m, e, tk))
        self.workers.append(worker)
        worker.start()

    def on_fetch_models_finished(self, success, models, err, tab_key):
        if success:
            combo = self.tab_ui[tab_key]["model"]
            current = combo.currentText()
            combo.clear()
            combo.addItems(models)
            idx = combo.findText(current)
            if idx >= 0: combo.setCurrentIndex(idx)
            QMessageBox.information(self, "完了", "モデル一覧を更新しました。")
        else:
            QMessageBox.critical(self, "エラー", f"失敗しました: {err}")

    def set_reco_model(self, tab_key):
        self.tab_ui[tab_key]["model"].setCurrentText("gemini-3.1-flash-lite-preview")
        self.tab_ui[tab_key]["rpm"].setValue(15 if tab_key == "free" else 1000)
        self.tab_ui[tab_key]["threads"].setValue(1)

    def set_reco_param(self, tab_key):
        self.tab_ui[tab_key]["temp"].setValue(0.0)
        self.tab_ui[tab_key]["tokens"].setValue(8192)
        self.tab_ui[tab_key]["safety"].setChecked(True)

    def load_settings(self):
        plan = self.settings.value("plan", "free")
        if plan == "paid":
            self.plan_combo.setCurrentIndex(1)
        else:
            self.plan_combo.setCurrentIndex(0)

        for tk in ["free", "paid"]:
            ui = self.tab_ui[tk]
            ui["api_key"].setText(self.settings.value(f"{tk}_api_key", ""))
            ui["model"].setCurrentText(self.settings.value(f"{tk}_model", "gemini-3.1-flash-lite-preview"))
            ui["rpm"].setValue(int(self.settings.value(f"{tk}_rpm", 15 if tk=="free" else 1000)))
            ui["threads"].setValue(int(self.settings.value(f"{tk}_threads", 1)))
            ui["temp"].setValue(float(self.settings.value(f"{tk}_temp", 0.0)))
            ui["tokens"].setValue(int(self.settings.value(f"{tk}_tokens", 8192)))
            ui["safety"].setChecked(self.settings.value(f"{tk}_safety", "true") == "true")

            curr_json = self.settings.value(f"{tk}_current_prompts", "[]")
            fav_json = self.settings.value(f"{tk}_fav_prompts", "[]")
            try:
                for text in json.loads(curr_json): ui["list_curr"].addItem(text)
                for text in json.loads(fav_json): ui["list_fav"].addItem(text)
            except: pass

    def save_settings(self):
        self.settings.setValue("plan", "paid" if self.plan_combo.currentIndex() == 1 else "free")
        
        for tk in ["free", "paid"]:
            ui = self.tab_ui[tk]
            self.settings.setValue(f"{tk}_api_key", ui["api_key"].text())
            self.settings.setValue(f"{tk}_model", ui["model"].currentText())
            self.settings.setValue(f"{tk}_rpm", ui["rpm"].value())
            self.settings.setValue(f"{tk}_threads", ui["threads"].value())
            self.settings.setValue(f"{tk}_temp", ui["temp"].value())
            self.settings.setValue(f"{tk}_tokens", ui["tokens"].value())
            self.settings.setValue(f"{tk}_safety", "true" if ui["safety"].isChecked() else "false")

            curr = [ui["list_curr"].item(i).text() for i in range(ui["list_curr"].count())]
            fav = [ui["list_fav"].item(i).text() for i in range(ui["list_fav"].count())]
            self.settings.setValue(f"{tk}_current_prompts", json.dumps(curr))
            self.settings.setValue(f"{tk}_fav_prompts", json.dumps(fav))

        self.accept()

class ResultWindow(QWidget):
    """ 結果表示ウィンドウ """
    def __init__(self, image_bytes, config, current_mode, worker=None):
        super().__init__()
        self.setWindowTitle("MiyashitaLens v1.1.0 - 結果")
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(500, 550)
        self.image_bytes = image_bytes
        self.config = config
        self.current_mode = current_mode
        self.worker = worker
        self.setStyleSheet("""
            QWidget { background-color: #ffffff; color: #333333; font-family: 'Segoe UI', Meiryo, sans-serif; }
            QLabel.SectionTitle { font-weight: bold; font-size: 13px; color: #1a73e8; margin-top: 10px; }
            QTextEdit { border: 1px solid #e0e0e0; border-radius: 8px; padding: 8px; font-size: 13px; background-color: #f8f9fa; }
            QPushButton { background-color: #1a73e8; color: white; border: none; border-radius: 6px; padding: 8px 16px; font-weight: bold; }
            QPushButton.ModeBtn { background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 4px; padding: 6px; font-size: 11px; color: #333; }
            QPushButton.ModeBtn[selected="true"] { background-color: #e8f0fe; border: 2px solid #1a73e8; color: #1a73e8; }
            QPushButton#CancelBtn { background-color: #d93025; }
            QPushButton#CancelBtn:disabled { background-color: #f1f3f4; color: #a8a8a8; }
        """)
        layout = QVBoxLayout()
        layout.addWidget(QLabel("📝 読み取ったテキスト:", objectName="SectionTitle"))
        self.original_edit = QTextEdit()
        self.original_edit.setPlainText("解析中...")
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
        self.processed_edit.setPlainText("待機中...")
        layout.addWidget(self.processed_edit)
        
        btn_box = QHBoxLayout()
        self.cancel_btn = QPushButton("処理を中止", objectName="CancelBtn")
        self.cancel_btn.clicked.connect(self.cancel_processing)
        btn_box.addWidget(self.cancel_btn)

        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.close)
        btn_box.addWidget(close_btn)
        
        layout.addLayout(btn_box)
        self.setLayout(layout)
        self.update_mode_ui()

    def update_mode_ui(self):
        for k, btn in self.mode_buttons.items():
            btn.setProperty("selected", str(k == self.current_mode).lower())
            btn.style().unpolish(btn); btn.style().polish(btn)

    def cancel_processing(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.original_edit.setPlainText("キャンセル済")
            self.processed_edit.setPlainText("処理がキャンセルされました。")
            self.cancel_btn.setEnabled(False)

    def on_chunk_received(self, text):
        """ ストリーミングからの断片データを受け取りUIに逐次反映する """
        self.processed_edit.setPlainText(text)
        scrollbar = self.processed_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_processing_finished(self, original_text, processed_text):
        self.original_edit.setPlainText(original_text)
        self.processed_edit.setPlainText(processed_text)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.hide()

    def on_processing_error(self, error_msg):
        self.original_edit.setPlainText("エラー")
        self.processed_edit.setPlainText(f"エラーが発生しました:\n{error_msg}")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.hide()

    def reprocess(self, new_mode):
        if self.current_mode == new_mode and self.worker and self.worker.isRunning():
            return
        
        if self.worker and self.worker.isRunning():
            self.worker.cancel()

        self.current_mode = new_mode
        self.update_mode_ui()
        self.original_edit.setPlainText("解析中...")
        self.processed_edit.setPlainText("待機中...")
        
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.show()

        self.worker = OcrTranslateWorker(self.image_bytes, self.current_mode, self.config)
        self.worker.chunk_received.connect(self.on_chunk_received)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.error.connect(self.on_processing_error)
        self.worker.start()

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
        super().closeEvent(event)

class OcrTranslateWorker(QThread):
    finished = pyqtSignal(str, str)
    chunk_received = pyqtSignal(str) # ストリーミング用
    error = pyqtSignal(str)
    
    def __init__(self, image_bytes, mode, config):
        super().__init__()
        self.image_bytes = image_bytes
        self.mode = mode
        self.config = config
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            if self._is_cancelled: return
            image_data = self.image_bytes

            ocr_text = ""
            if WIN_OCR_AVAILABLE:
                try:
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    ocr_text = loop.run_until_complete(perform_local_ocr(image_data))
                    loop.close()
                except Exception as e:
                    print(f"OCR Exception: {e}")
                    ocr_text = ""

            if self._is_cancelled: return

            genai.configure(api_key=self.config["api_key"])
            
            sys_instruct = "挨拶や前置きは不要です。要求された結果のみを直接出力してください。"
            model = genai.GenerativeModel(self.config["model_name"], system_instruction=sys_instruct)
            
            add_prompt = ""
            if self.config.get("custom_prompts"):
                add_prompt = "\n\n【追加指示】\n" + "\n".join(self.config["custom_prompts"])

            if ocr_text and ocr_text.strip():
                clean_text = ocr_text.strip().replace('"', '”').replace('\n', ' ')
                prompts = {
                    "ja_translate": f"次のテキストを自然な日本語に翻訳してください。\n\n{clean_text}",
                    "en_translate": f"次のテキストを自然な英語に翻訳してください。\n\n{clean_text}",
                    "dictionary": f"次のテキストの意味・発音・使い方を辞書として簡潔に解説してください。\n\n{clean_text}"
                }
                prompt = f"{prompts.get(self.mode, '')}{add_prompt}"
                contents = [prompt]
                self.chunk_received.emit("翻訳中...")
            else:
                prompts = {
                    "ja_translate": "画像内のテキストを読み取り、自然な日本語に翻訳してください。",
                    "en_translate": "画像内のテキストを読み取り、自然な英語に翻訳してください。",
                    "dictionary": "画像内のメインテキストを読み取り、意味と使い方を簡潔に解説してください。"
                }
                prompt = f"{prompts.get(self.mode, '')}{add_prompt}"
                contents = [prompt, {"mime_type": "image/jpeg", "data": image_data}]
                self.chunk_received.emit("画像を解析中...")
            
            if self._is_cancelled: return

            safety_settings = None
            if self.config.get("safety", True):
                safety_settings = [
                    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
                ]

            generation_config = genai.types.GenerationConfig(
                temperature=self.config.get("temp", 0.0),
                max_output_tokens=self.config.get("max_tokens", 8192)
            )

            response = model.generate_content(
                contents,
                generation_config=generation_config,
                safety_settings=safety_settings,
                stream=True
            )
            
            proc_text = ""
            for chunk in response:
                if self._is_cancelled: return
                if chunk.text:
                    proc_text += chunk.text
                    self.chunk_received.emit(proc_text)
            
            if self._is_cancelled: return

            if ocr_text and ocr_text.strip():
                orig_text = ocr_text.strip()
            else:
                orig_text = "（画像から直接読み取りました）"

            if not self._is_cancelled:
                self.finished.emit(orig_text, proc_text)
                
        except Exception as e:
            if not self._is_cancelled:
                error_msg = str(e)
                if "429" in error_msg or "quota" in error_msg.lower():
                    error_msg = "APIの無料枠制限（1分間あたりの回数制限）を超過しました。\n約1分ほど待ってから再度実行してください。"
                elif "404" in error_msg and "models/" in error_msg:
                    error_msg = f"指定されたモデルが見つかりません: {self.config['model_name']}\n⚙️ API詳細設定から「🌐 更新」ボタンを押し、一覧から有効なモデル（gemini-3.1-flash-lite-previewなど）を選択し直してください。"
                self.error.emit(error_msg)

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
            error_msg = str(e)
            if "429" in error_msg or "quota" in error_msg.lower():
                error_msg = "APIの無料枠制限に達しています。1分ほどお待ちください。"
            self.finished.emit(False, [], error_msg)

class ApiTestWorker(QThread):
    finished = pyqtSignal(bool, str)
    def __init__(self, api_key, model_name):
        super().__init__()
        self.api_key = api_key
        self.model_name = model_name
    def run(self):
        try:
            genai.configure(api_key=self.api_key)
            model_info = genai.get_model(f"models/{self.model_name}")
            self.finished.emit(True, f"接続成功: {model_info.display_name}")
        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "quota" in error_msg.lower():
                error_msg = "APIの無料枠制限に達しています。1分ほどお待ちください。"
            self.finished.emit(False, error_msg)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.version = "v1.1.0"
        self.setWindowTitle(f"MiyashitaLens {self.version}")
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(360, 480)
        self.setStyleSheet("""
            QMainWindow, QWidget { background-color: #ffffff; color: #333333; font-family: 'Segoe UI', Meiryo, sans-serif; }
            QPushButton#SnipBtn { background-color: #1a73e8; color: white; border-radius: 6px; padding: 12px; font-weight: bold; }
            QPushButton.ModeBtn { background-color: #f8f9fa; border: 1px solid #ddd; padding: 10px; border-radius: 4px; font-size: 11px; }
            QPushButton.ModeBtn[selected="true"] { background-color: #e8f0fe; border: 2px solid #1a73e8; color: #1a73e8; }
            QPushButton#ActionBtn { background-color: #f1f3f4; border: 1px solid #ccc; padding: 8px 10px; border-radius: 4px; font-weight: bold; color: #3c4043; }
            QPushButton#ActionBtn:hover { background-color: #e8eaed; }
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
        help_menu.addAction("バージョン情報", lambda: QMessageBox.information(self, "情報", f"MiyashitaLens {self.version}"))

        central = QWidget(); self.setCentralWidget(central); layout = QVBoxLayout(central)
        layout.addWidget(QLabel("<b>【使い方】</b><br>1. モードを選択しAPI設定を行う<br>2. 画面を切り取って実行"))
        
        mode_layout = QHBoxLayout()
        self.mode_btns = {}
        for k, n in [("ja_translate", "日本語翻訳"), ("en_translate", "英語翻訳"), ("dictionary", "辞書")]:
            btn = QPushButton(n); btn.setProperty("class", "ModeBtn"); btn.clicked.connect(lambda _, m=k: self.set_mode(m))
            self.mode_btns[k] = btn; mode_layout.addWidget(btn)
        layout.addLayout(mode_layout)

        setting_box = QHBoxLayout()
        api_btn = QPushButton("⚙️ API詳細設定")
        api_btn.setObjectName("ActionBtn")
        api_btn.clicked.connect(self.open_settings)
        setting_box.addWidget(api_btn)
        
        self.ontop_cb = QCheckBox("最前面表示")
        self.ontop_cb.clicked.connect(self.toggle_always_on_top)
        setting_box.addStretch()
        setting_box.addWidget(self.ontop_cb)
        
        layout.addLayout(setting_box)

        self.status_label = QLabel("待機中", objectName="StatusLabel")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()
        layout.addWidget(self.status_label)

        snip_btn = QPushButton("🔍 画面を切り取って実行", objectName="SnipBtn"); snip_btn.clicked.connect(self.start_snipping)
        layout.addWidget(snip_btn)

        self.settings = QSettings(os.path.join(os.path.expanduser('~'), 'MiyashitaLens', 'settings.ini'), QSettings.Format.IniFormat)
        
        is_ontop = self.settings.value("always_on_top", "true").lower() == "true"
        self.ontop_cb.setChecked(is_ontop)
        if not is_ontop:
            self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.Tool)
            self.show()

        self.current_mode = self.settings.value("last_mode", "ja_translate")
        self.set_mode(self.current_mode)

    def toggle_always_on_top(self):
        is_checked = self.ontop_cb.isChecked()
        if is_checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
        self.settings.setValue("always_on_top", str(is_checked).lower())
        self.show()

    def set_mode(self, mode):
        self.current_mode = mode
        self.settings.setValue("last_mode", mode)
        for k, b in self.mode_btns.items():
            b.setProperty("selected", str(k == mode).lower())
            b.style().unpolish(b); b.style().polish(b)

    def open_settings(self):
        settings_win = SettingsWindow(self.settings, self)
        settings_win.exec()

    def show_help(self):
        self.hw = HelpWindow(); self.hw.show()

    def start_snipping(self):
        plan = self.settings.value("plan", "free")
        api_key = self.settings.value(f"{plan}_api_key", "")
        if not api_key:
            QMessageBox.warning(self, "エラー", "APIキーが設定されていません。\n「⚙️ API詳細設定」からキーを設定してください。")
            return
        
        self.hide()
        QTimer.singleShot(300, self.init_snipping)

    def init_snipping(self):
        self.status_label.setText("範囲を選択してください...")
        self.status_label.setStyleSheet("color: #666;")
        self.snipping_tool = SnippingWidget(self)
        self.snipping_tool.show()

    def process_captured_image(self, img_data):
        plan = self.settings.value("plan", "free")
        config = {
            "api_key": self.settings.value(f"{plan}_api_key", ""),
            "model_name": self.settings.value(f"{plan}_model", "gemini-3.1-flash-lite-preview"),
            "temp": float(self.settings.value(f"{plan}_temp", 0.0)),
            "max_tokens": int(self.settings.value(f"{plan}_tokens", 8192)),
            "safety": self.settings.value(f"{plan}_safety", "true") == "true",
            "custom_prompts": []
        }
        
        try:
            config["custom_prompts"] = json.loads(self.settings.value(f"{plan}_current_prompts", "[]"))
        except:
            pass

        mode = self.current_mode
        
        self.status_label.setText("⏳ AI処理中... しばらくお待ちください")
        self.status_label.setStyleSheet("color: #1a73e8; font-weight: bold;")
        
        self.worker = OcrTranslateWorker(img_data, mode, config)
        
        res_win = ResultWindow(img_data, config, mode, self.worker)
        self.active_results.append(res_win) 
        res_win.show()
        res_win.raise_()
        res_win.activateWindow()

        self.worker.chunk_received.connect(res_win.on_chunk_received)
        self.worker.finished.connect(res_win.on_processing_finished)
        self.worker.error.connect(res_win.on_processing_error)
        
        self.worker.finished.connect(self.reset_status)
        self.worker.error.connect(self.reset_status_error)

        self.worker.start()

    def reset_status(self, original_text="", processed_text=""):
        self.status_label.setText("✅ 処理完了")
        self.status_label.setStyleSheet("color: #188038;")
        QTimer.singleShot(3000, lambda: self.status_label.setText("待機中"))
        QTimer.singleShot(3000, lambda: self.status_label.setStyleSheet("color: #666;"))

    def reset_status_error(self, msg=""):
        self.status_label.setText("❌ エラーが発生しました")
        self.status_label.setStyleSheet("color: #d93025;")
        QTimer.singleShot(3000, lambda: self.status_label.setText("待機中"))
        QTimer.singleShot(3000, lambda: self.status_label.setStyleSheet("color: #666;"))

    def closeEvent(self, event):
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
            pixmap = self.original_pixmap.copy(r)
            
            if pixmap.width() < 300 or pixmap.height() < 100:
                scale_factor = 2
                pixmap = pixmap.scaled(pixmap.width() * scale_factor, pixmap.height() * scale_factor, 
                                       Qt.AspectRatioMode.IgnoreAspectRatio, 
                                       Qt.TransformationMode.FastTransformation)
            
            margin = 40
            padded_pixmap = QPixmap(pixmap.width() + margin * 2, pixmap.height() + margin * 2)
            padded_pixmap.fill(Qt.GlobalColor.white)
            painter = QPainter(padded_pixmap)
            painter.drawPixmap(margin, margin, pixmap)
            painter.end()

            buf = QBuffer()
            buf.open(QIODevice.OpenModeFlag.ReadWrite)
            
            padded_pixmap.save(buf, "JPG", 85)
            
            img_bytes = bytes(buf.data())
            self.main_win.process_captured_image(img_bytes)
        
        self.close()
        self.main_win.showNormal()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())