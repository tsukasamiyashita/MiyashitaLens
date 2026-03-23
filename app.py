import sys
import os
import time
import requests
import urllib3
import cv2
import numpy as np
import asyncio
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QTextEdit, QLabel, QMainWindow)
from PyQt6.QtCore import Qt, QRect, QPoint, QTimer, QThread, pyqtSignal, QBuffer, QIODevice
from PyQt6.QtGui import QPainter, QColor, QPen, QGuiApplication, QScreen, QPixmap, QIcon

# WinRT API (Windows標準機能) へのアクセスモジュール
from winsdk.windows.media.ocr import OcrEngine
from winsdk.windows.globalization import Language
from winsdk.windows.graphics.imaging import BitmapDecoder
from winsdk.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

class MainWindow(QWidget):
    """
    常に最前面に表示される、操作方法の案内とキャプチャ起動ボタンを持つメインウィンドウ
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Glans_Miya")
        # 常に手前に表示される小さなツールウィンドウとして設定
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(320, 220)
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #333333;
                font-family: 'Segoe UI', Meiryo, sans-serif;
            }
            QLabel {
                font-size: 13px;
                line-height: 1.5;
            }
            QPushButton {
                background-color: #1a73e8;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px;
                font-size: 14px;
                font-weight: bold;
                margin-top: 10px;
            }
            QPushButton:hover {
                background-color: #1557b0;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # 操作方法の表示
        guide_label = QLabel(
            "<b>【使い方】</b><br><br>"
            "1. 下のボタンをクリックします<br>"
            "2. 画面が暗くなったら、翻訳したい<br>"
            "　文字や画像をドラッグして囲みます<br>"
            "3. 翻訳結果のウィンドウが表示されます"
        )
        guide_label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(guide_label)

        layout.addStretch()

        # キャプチャ起動ボタン
        self.snip_btn = QPushButton("🔍 画面を切り取って翻訳")
        self.snip_btn.clicked.connect(self.start_snipping)
        layout.addWidget(self.snip_btn)

        self.setLayout(layout)
        self.snipping_tool = None

    def closeEvent(self, event):
        # メインウィンドウが閉じられたらアプリ全体を終了する
        QApplication.quit()
        super().closeEvent(event)

    def start_snipping(self):
        # 自身を非表示にして、画面キャプチャに写り込まないようにする
        self.hide()
        # 非表示が完全に画面に反映されるよう少し待機してからキャプチャ開始
        QTimer.singleShot(300, self.show_snipping_widget)

    def show_snipping_widget(self):
        self.snipping_tool = SnippingWidget(main_window=self)
        self.snipping_tool.show()


class ResultWindow(QWidget):
    """
    抽出したテキストと翻訳結果を表示するモダンなポップアップウィンドウ
    """
    def __init__(self, original_text, translated_text):
        super().__init__()
        self.setWindowTitle("Glans_Miya - 翻訳結果")
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.resize(400, 300)
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #333333;
                font-family: 'Segoe UI', Meiryo, sans-serif;
            }
            QLabel {
                font-weight: bold;
                font-size: 14px;
                color: #1a73e8;
                margin-top: 10px;
            }
            QTextEdit {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 8px;
                font-size: 13px;
                background-color: #f8f9fa;
            }
            QPushButton {
                background-color: #1a73e8;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1557b0;
            }
        """)

        layout = QVBoxLayout()

        # 抽出テキストエリア
        layout.addWidget(QLabel("📝 読み取ったテキスト:"))
        self.original_edit = QTextEdit()
        self.original_edit.setPlainText(original_text)
        layout.addWidget(self.original_edit)

        # 翻訳テキストエリア
        layout.addWidget(QLabel("🌐 翻訳結果 (日本語):"))
        self.translated_edit = QTextEdit()
        self.translated_edit.setPlainText(translated_text)
        layout.addWidget(self.translated_edit)

        # 閉じるボタン
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("閉じる")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)


class OcrTranslateWorker(QThread):
    """
    UIのフリーズを防ぐため、OCRと翻訳処理をバックグラウンドで行うワーカースレッド
    """
    finished = pyqtSignal(str, str) # 成功時のシグナル (抽出テキスト, 翻訳テキスト)
    error = pyqtSignal(str)         # エラー時のシグナル (エラーメッセージ)

    def __init__(self, image_bytes):
        super().__init__()
        # 物理ファイルではなくメモリ上のバイトデータを保持
        self.image_bytes = image_bytes

    async def _recognize_text(self, cv_img):
        """
        Windows標準のWinRT OCRを利用して極めて高速・高精度にテキストを抽出する非同期メソッド
        """
        # OpenCVの画像データをPNG形式のバイト列にメモリ上でエンコード
        is_success, buffer = cv2.imencode(".png", cv_img)
        if not is_success:
            raise Exception("画像データの内部エンコードに失敗しました。")
        image_bytes = buffer.tobytes()

        # バイト列をWinRTのストリーム形式に変換
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(image_bytes)
        await writer.store_async()
        stream.seek(0)

        # ストリームからSoftwareBitmapを生成
        decoder = await BitmapDecoder.create_async(stream)
        software_bitmap = await decoder.get_software_bitmap_async()

        # 日本語対応のOCRエンジンを初期化（日本語パックがない場合はシステムのデフォルト言語）
        lang = Language("ja-JP")
        if OcrEngine.is_language_supported(lang):
            engine = OcrEngine.try_create_from_language(lang)
        else:
            engine = OcrEngine.try_create_from_user_profile_languages()

        if engine is None:
            raise Exception("OCRエンジンを初期化できません。Windowsの言語設定(日本語)を確認してください。")

        # 文字認識を実行
        result = await engine.recognize_async(software_bitmap)
        return result.text

    def _translate_text(self, text):
        """
        Google翻訳APIを直接叩き、プロキシ等のSSLエラーに耐性を持たせた堅牢な翻訳処理
        """
        url = "https://translate.googleapis.com/translate_a/single"
        params = {
            "client": "gtx",
            "sl": "auto",
            "tl": "ja",
            "dt": "t",
            "q": text
        }
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        
        max_retries = 3
        last_exception = None

        for attempt in range(max_retries):
            try:
                # 通常のSSL検証ありでリクエスト
                response = requests.get(url, params=params, headers=headers, timeout=10)
                response.raise_for_status()
                
                # レスポンスのJSONから翻訳結果を抽出
                result = response.json()
                translated_text = "".join([sentence[0] for sentence in result[0] if sentence[0]])
                return translated_text
                
            except requests.exceptions.SSLError as e:
                # SSLエラー（プロキシやセキュリティソフトによるインスペクション等）が起きた場合のフォールバック
                try:
                    # InsecureRequestWarning を抑制
                    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                    # verify=False で証明書検証をスキップして強行接続
                    response = requests.get(url, params=params, headers=headers, timeout=10, verify=False)
                    response.raise_for_status()
                    
                    result = response.json()
                    translated_text = "".join([sentence[0] for sentence in result[0] if sentence[0]])
                    return translated_text
                except Exception as e_fallback:
                    last_exception = e_fallback
                    time.sleep(2 ** attempt) # 指数的バックオフ (1秒, 2秒...)
            
            except Exception as e:
                last_exception = e
                time.sleep(2 ** attempt)

        # 全てのリトライに失敗した場合
        raise Exception(f"翻訳サーバーに接続できませんでした。(詳細: {str(last_exception)})")

    def run(self):
        try:
            # --- OpenCVによる高速なインメモリ画像前処理 ---
            # バイトデータからNumPy配列を生成し、OpenCVの画像データとしてデコード
            nparr = np.frombuffer(self.image_bytes, np.uint8)
            img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

            # 1. グレースケール化
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

            # 2. リサイズ処理（巨大な領域が選択された場合のOCR遅延を防ぐ）
            max_width = 1200
            if gray.shape[1] > max_width:
                scale = max_width / gray.shape[1]
                gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)

            # 3. コントラスト強調（適応的ヒストグラム平坦化）でかすれ文字を補正
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            gray = clahe.apply(gray)
            # ------------------------------------------------

            # WinRT OCRエンジンの非同期処理を同期的に実行して結果を取得
            extracted_text = asyncio.run(self._recognize_text(gray)).strip()
            
            translated_text = ""
            if extracted_text:
                try:
                    # 独自の堅牢なメソッドで翻訳を実行
                    translated_text = self._translate_text(extracted_text)
                except Exception as e:
                    translated_text = f"【翻訳通信エラー】\n{str(e)}"
            else:
                extracted_text = "テキストが検出されませんでした。"
            
            self.finished.emit(extracted_text, translated_text)

        except Exception as e:
            self.error.emit(str(e))


class SnippingWidget(QMainWindow):
    """
    画面全体を暗くし、ドラッグで領域を選択するオーバーレイウィンドウ
    """
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window # 呼び出し元のメインウィンドウを保持
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCursor(Qt.CursorShape.CrossCursor)

        # 全画面のスクリーンショットを取得
        screen: QScreen = QGuiApplication.primaryScreen()
        self.original_pixmap = screen.grabWindow(0)
        
        # 画面サイズに合わせてウィンドウを展開
        self.setGeometry(screen.geometry())

        self.begin_point = QPoint()
        self.end_point = QPoint()
        self.is_drawing = False
        self.result_window = None

    def paintEvent(self, event):
        painter = QPainter(self)
        
        # 取得したスクリーンショットを背景として描画
        painter.drawPixmap(self.rect(), self.original_pixmap)
        
        # 画面全体に半透明の黒いレイヤーをかける
        overlay_color = QColor(0, 0, 0, 100)
        painter.fillRect(self.rect(), overlay_color)

        # 選択領域をクリアして元の明るい画像を見せる
        if not self.begin_point.isNull() and not self.end_point.isNull():
            rect = QRect(self.begin_point, self.end_point).normalized()
            
            # 選択範囲の半透明マスクをくり抜く（元の画像を上書き描画）
            cropped_pixmap = self.original_pixmap.copy(rect)
            painter.drawPixmap(rect, cropped_pixmap)

            # 選択範囲の境界線を描画
            pen = QPen(QColor(26, 115, 232), 2) # Googleブルー
            painter.setPen(pen)
            painter.drawRect(rect)
            
        # 操作案内の表示（まだドラッグを開始していない初期状態の時のみ）
        if self.begin_point.isNull() and self.end_point.isNull():
            guide_text = "マウスをドラッグして、翻訳したい範囲を選択してください"
            font = painter.font()
            font.setPointSize(14)
            font.setBold(True)
            font.setFamily('Segoe UI')
            painter.setFont(font)
            
            metrics = painter.fontMetrics()
            text_width = metrics.horizontalAdvance(guide_text)
            text_height = metrics.height()
            
            # 画面中央の座標を計算
            center_x = self.rect().width() // 2
            center_y = self.rect().height() // 2
            
            # テキスト背景の余白
            padding_x = 30
            padding_y = 15
            
            bg_rect = QRect(center_x - text_width // 2 - padding_x,
                            center_y - text_height // 2 - padding_y,
                            text_width + padding_x * 2,
                            text_height + padding_y * 2)
            
            # 角丸の半透明黒背景を描画
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(0, 0, 0, 180))
            painter.drawRoundedRect(bg_rect, 10, 10)
            
            # 白文字でテキストを描画
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(bg_rect, Qt.AlignmentFlag.AlignCenter, guide_text)

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
            
            # 範囲が小さすぎる（クリックのみ）場合は処理せず閉じる
            if rect.width() <= 10 or rect.height() <= 10:
                self.close()
                return

            # 領域選択された場合は非同期処理を開始する
            self.process_selected_area(rect)

    def process_selected_area(self, rect: QRect):
        """
        選択された領域の画像を切り抜き、別スレッドでOCRと翻訳を実行する
        """
        # 処理中であることをユーザーに伝えるためカーソルを砂時計にする
        self.setCursor(Qt.CursorShape.WaitCursor)

        # 選択領域のピックスマップを取得
        cropped_pixmap = self.original_pixmap.copy(rect)
        
        # 物理ファイルを介さず、QBufferを使ってメモリ上でバイトデータを取得する（高速化）
        buffer = QBuffer()
        buffer.open(QIODevice.OpenModeFlag.ReadWrite)
        cropped_pixmap.save(buffer, "PNG")
        image_bytes = buffer.data().data()

        # ワーカースレッドの生成とシグナルの接続
        self.worker = OcrTranslateWorker(image_bytes)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.error.connect(self.on_worker_error)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.worker.start()

    def on_worker_finished(self, extracted_text, translated_text):
        # 処理成功時のコールバック
        self.setCursor(Qt.CursorShape.CrossCursor)
        # 結果ウィンドウをMainWindowに保持させる（ガベージコレクションによる消滅を防止）
        if self.main_window:
            self.main_window.result_window = ResultWindow(extracted_text, translated_text)
            self.main_window.result_window.show()
        self.cleanup_and_close()

    def on_worker_error(self, error_msg):
        # 処理失敗時のコールバック
        self.setCursor(Qt.CursorShape.CrossCursor)
        if self.main_window:
            self.main_window.result_window = ResultWindow("エラーが発生しました", error_msg)
            self.main_window.result_window.show()
        self.cleanup_and_close()

    def cleanup_and_close(self):
        # 物理ファイルを介していないため削除処理は不要。オーバーレイウィンドウのみ閉じる
        self.close()

    def closeEvent(self, event):
        # SnippingWidgetが閉じる際に、元の操作パネルを再表示させる
        if self.main_window:
            self.main_window.showNormal()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Glans_Miya")
    app.setQuitOnLastWindowClosed(False) # 意図せぬサイレント終了を防止
    
    # 常に表示される小さな操作ウィンドウを起動
    main_win = MainWindow()
    main_win.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()