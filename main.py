# -*- coding: utf-8 -*-
"""
小学生にもわかる説明：
  「layout.json」という設計図を読みこんで、画面の部品(ラベルや入力など)を
  決められた場所に並べます。Excel(xlsm)から読み書きもできます。
"""

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtGui import QIntValidator, QDoubleValidator
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPlainTextEdit, QPushButton, QFileDialog,
    QScrollArea, QFrame, QStatusBar, QMessageBox, QComboBox,
    QHBoxLayout
)
from qt_material import apply_stylesheet
from openpyxl import load_workbook

import sys
import json
import os
import re
from typing import Dict, Optional, Any, List, Callable
import threading


# 数字入力専用のラインエディットです。フォーカス時に入力モードを半角英数字に固定します。
class NumericLineEdit(QLineEdit):
    def __init__(self, *args, **kwargs):
        # まずは親クラス(QLineEdit)の初期化を行います。
        super().__init__(*args, **kwargs)
        # フォーカス前の入力ヒントを保存するための変数を用意します。
        self._prev_hints: Optional[QtCore.Qt.InputMethodHints] = None

    def focusInEvent(self, event: QtGui.QFocusEvent) -> None:
        # この欄にフォーカスが当たったときの処理です。
        # 現在の入力ヒントを記録し、半角英数字のみになるようヒントを追加します。
        self._prev_hints = self.inputMethodHints()
        self.setInputMethodHints(self._prev_hints | QtCore.Qt.ImhLatinOnly)
        super().focusInEvent(event)

    def focusOutEvent(self, event: QtGui.QFocusEvent) -> None:
        # フォーカスが外れたときに元の入力ヒントへ戻します。
        if self._prev_hints is not None:
            self.setInputMethodHints(self._prev_hints)
        super().focusOutEvent(event)


# === 起動時のエクセル読み込み待機ウインドウ ===
class LoadingSpinner(QtWidgets.QDialog):
    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        # 親クラスの初期化を行います。
        super().__init__(parent)
        # 枠を消してシンプルな小窓にします。
        self.setWindowFlags(QtCore.Qt.Dialog | QtCore.Qt.FramelessWindowHint)
        # 画面中央に配置されるように少し大きめのサイズを固定し、
        # スピナーと文字が重ならないようにします。
        self.setFixedSize(160, 160)

        # 縦に部品を並べるレイアウトを用意します。
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        # スピナーとラベルの間に適度な空白を確保します。
        layout.setSpacing(10)

        # 実際にぐるぐるを描画する小さなウィジェットを追加します。
        self.indicator = _SpinnerWidget(self)
        layout.addWidget(self.indicator, alignment=QtCore.Qt.AlignCenter)

        # エクセルファイルを読み込んでいることを示すラベルを、2行に分けて表示します。
        label = QLabel("エクセルファイル\n読み込み中......", self)
        # ラベル内の文字列を上下左右ともに中央揃えにします。
        label.setAlignment(QtCore.Qt.AlignCenter)
        # ラベルをレイアウト中央に配置します。
        layout.addWidget(label, alignment=QtCore.Qt.AlignCenter)


class _SpinnerWidget(QtWidgets.QWidget):
    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        # 親クラスの初期化を行います。
        super().__init__(parent)
        # 描画に使う角度を管理する変数を初期化します。
        self._angle = 0
        # 線の太さを調整するための変数です。
        self._pen_width = 8
        # 高速でタイマーを回して滑らかに回転させます。16ms 間隔はおよそ60fpsです。
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._rotate)
        self._timer.start(16)
        # 大きさを決めます。ここでは正方形の領域とします。
        self.setFixedSize(64, 64)
        # スピナーが四角い枠で切れないよう、背景を透過させます。
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)

    def _rotate(self) -> None:
        # 角度を細かく増やして滑らかに回転させます。360度で一周です。
        self._angle = (self._angle + 5) % 360
        # 値が変わったので再描画を依頼します。
        self.update()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        # ぐるぐるの一部(円弧)を描画します。
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        rect = self.rect()
        painter.translate(rect.center())
        # 現在の角度だけ回転させます。
        painter.rotate(self._angle)
        # ペンの太さ分だけ半径を小さくし、円弧が切れないよう余白を確保します。
        radius = min(rect.width(), rect.height()) / 2 - self._pen_width
        # スピナーを青色にするため、RGB(0,0,255)を指定したペンを使用します。
        pen = QtGui.QPen(QtGui.QColor(0, 0, 255), self._pen_width)
        # ペンの端を丸くして自然な見た目にします。
        pen.setCapStyle(QtCore.Qt.RoundCap)
        painter.setPen(pen)
        # 270度分だけ線を描いて空白を作り、回転で動いているように見せます。
        painter.drawArc(QtCore.QRectF(-radius, -radius, radius * 2, radius * 2), 0, 270 * 16)


# === シリンダー入力ユニット ===
class CylinderUnit(QWidget):
    def __init__(self, get_item_no: Callable[[], str]) -> None:
        """シリンダー情報を1行分まとめる部品です。"""
        super().__init__()
        # 品目番号を取得する関数を保持します。
        self._get_item_no = get_item_no

        # 全体を縦に並べるレイアウトを用意します。
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # 上段の横並びレイアウトを用意します。
        row = QHBoxLayout()

        # 色の順番を示すコンボボックスです。後で番号を設定します。
        self.order_combo = QComboBox()
        row.addWidget(self.order_combo)

        # 使用シリンダー番号を入力するコンボボックスです。入力補完のため編集可能にします。
        self.cylinder_combo = QComboBox()
        self.cylinder_combo.setEditable(True)
        self._refresh_cylinder_list()
        row.addWidget(self.cylinder_combo)

        # 色名を入力するテキストボックスです。
        self.color_edit = QLineEdit()
        row.addWidget(self.color_edit)

        # ベタ巾を入力する数値専用の欄です。
        self.width_edit = NumericLineEdit()
        dv = QDoubleValidator(0.0, 1e12, 3)
        dv.setNotation(QDoubleValidator.Notation.StandardNotation)
        dv.setLocale(QtCore.QLocale("C"))
        self.width_edit.setValidator(dv)
        self.width_edit.setInputMethodHints(QtCore.Qt.ImhPreferNumbers)
        row.addWidget(self.width_edit)

        # 旧版処理を選ぶコンボボックスです。
        self.process_combo = QComboBox()
        self.process_combo.addItems([
            "変更無し",
            "同名製版",
            "落組行き",
            "名義変更",
            "廃棄行き",
        ])
        row.addWidget(self.process_combo)

        # 上段レイアウトをウィジェット全体に追加します。
        layout.addLayout(row)

        # 名義変更を選んだ場合にのみ表示する数値入力欄です。
        self.rename_edit = NumericLineEdit()
        iv = QIntValidator(0, 999999999)
        iv.setLocale(QtCore.QLocale("C"))
        self.rename_edit.setValidator(iv)
        self.rename_edit.setInputMethodHints(QtCore.Qt.ImhDigitsOnly)
        self.rename_edit.setPlaceholderText("名義変更先の番号")
        self.rename_edit.hide()
        layout.addWidget(self.rename_edit)

        # 旧版処理の選択が変わったときに表示・非表示を切り替えます。
        self.process_combo.currentTextChanged.connect(self._on_process_changed)

    def _refresh_cylinder_list(self) -> None:
        """品目番号に基づくシリンダー候補を更新します。"""
        self.cylinder_combo.clear()
        item_no = self._get_item_no()
        # ここでは簡易的に1から10までの番号を候補としています。
        # 実際の仕様に合わせて調整してください。
        candidates = [f"{item_no}-{i}" if item_no else str(i) for i in range(1, 11)]
        self.cylinder_combo.addItems(candidates)

    def _on_process_changed(self, text: str) -> None:
        """旧版処理の内容によって追加欄を表示します。"""
        self.rename_edit.setVisible(text == "名義変更")

# === Excel の読み書き関数 ===
def read_record_from_xlsm(path: str, item_no: str, sheet_name: str) -> Optional[Dict[str, str]]:
    """Excel ファイルから品目番号に一致する行を辞書にして返します。"""
    # Excel ファイルを開きます。
    wb = load_workbook(path, keep_vba=True, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シート『{sheet_name}』がありません")
    ws = wb[sheet_name]

    # 1 行目の見出しを調べて、列名と列番号の対応表を作ります。
    header_map: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v)] = c

    # 必要な「品目番号」列が存在するか確認します。
    if "品目番号" not in header_map:
        raise ValueError("『品目番号』の列が見つかりません")

    col_item = header_map["品目番号"]
    target_row = None
    # 2 行目以降を順番に見て品目番号が一致する行を探します。
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=col_item).value or "") == item_no:
            target_row = r
            break
    # 見つからなければ None を返します。
    if target_row is None:
        return None

    # 見つかった行の値を列名とセットで辞書にまとめて返します。
    return {name: str(ws.cell(row=target_row, column=header_map.get(name)).value or "")
            for name in header_map.keys()}


def upsert_record_to_xlsm(path: str, data: Dict[str, str], sheet_name: str) -> None:
    """品目番号をキーとして Excel ファイルへ行の追加または更新を行います。"""
    # Excel ファイルを開きます。
    wb = load_workbook(path, keep_vba=True, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シート『{sheet_name}』がありません")
    ws = wb[sheet_name]

    # 見出し行を読み取り、列名と列番号の対応表を作ります。
    header_map: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v)] = c

    # 「品目番号」列が無いときはエラーにします。
    if "品目番号" not in header_map:
        raise ValueError("『品目番号』の列が見つかりません")

    col_item = header_map["品目番号"]
    target_row = None
    # 既存の行に同じ品目番号があるか確認します。
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=col_item).value or "") == data.get("品目番号", ""):
            target_row = r
            break
    # 無ければ最後の行の次に追加します。
    if target_row is None:
        target_row = ws.max_row + 1

    # 対応する列に値を書き込みます。存在しない列は無視します。
    for k, v in data.items():
        c = header_map.get(k)
        if c is None:
            continue
        ws.cell(row=target_row, column=c).value = v

    # Excel ファイルを保存します。
    wb.save(path)


# === 起動時データ抽出関数 ===
def extract_initial_data(path: str, progress: Optional[Callable[[int, int], None]] = None) -> Dict[str, List[List[Any]]]:
    """指定された Excel ファイルから二つのシートのデータをまとめて取得します。"""
    # Excel ファイル全体を開きます。
    wb = load_workbook(path, keep_vba=True, data_only=False)

    # 結果を格納する辞書を用意します。
    result: Dict[str, List[List[Any]]] = {}

    # 進捗管理のために対象シートの一覧を用意します。
    sheets = ("受注データ", "シリンダーデータ")
    total = len(sheets)

    # 必要なシート名を順番に処理します。
    for idx, sheet in enumerate(sheets, start=1):
        if sheet in wb.sheetnames:
            # 各シートから必要範囲のデータを取得します。
            ws = wb[sheet]
            result[sheet] = _extract_range_from_sheet(ws)
        # 進捗コールバックがあれば現在の状況を通知します。
        if progress is not None:
            progress(idx, total)

    # 取得したデータを返します。
    return result


def _extract_range_from_sheet(ws) -> List[List[Any]]:
    """1つのシートから必要な範囲のデータだけを取り出します。"""
    # 列ヘッダー行の最後に入力がある列番号を求めます。
    max_col = ws.max_column
    while max_col > 0 and ws.cell(row=1, column=max_col).value in (None, ""):
        max_col -= 1

    # ヘッダー名と列番号の対応表を作ります。
    header_map: Dict[str, int] = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v)] = c

    # 「品目番号」と「品目番号+刷順」列の最後の入力行を確認します。
    # ここではデータが入力されている最終行を求めています。
    last_row = 1
    for key in ("品目番号", "品目番号+刷順"):
        col = header_map.get(key)
        if col is None:
            continue
        for r in range(ws.max_row, 1, -1):
            if ws.cell(row=r, column=col).value not in (None, ""):
                if r > last_row:
                    last_row = r
                break

    # 取得する範囲のセルの値を二次元リストにまとめます。
    data: List[List[Any]] = []
    for r in range(1, last_row + 1):
        row_values: List[Any] = []
        for c in range(1, max_col + 1):
            row_values.append(ws.cell(row=r, column=c).value)
        data.append(row_values)

    # 作成したデータを返します。
    return data


def find_record_by_column(data: List[List[Any]], column: str, value: str) -> Optional[Dict[str, str]]:
    """事前に読み込んだ二次元リストから指定列の値を探します。"""
    # データが空の場合はすぐに終了します。
    if not data:
        return None

    # 1 行目は見出しとして扱い、列名の一覧を作ります。
    headers = [str(v) if v is not None else "" for v in data[0]]

    # 探したい列名が存在するか確認します。
    if column not in headers:
        raise ValueError(f"『{column}』の列が見つかりません")

    idx = headers.index(column)

    # 2 行目以降を順番に調べて一致する値を探します。
    for row in data[1:]:
        cell_value = ""
        if idx < len(row) and row[idx] is not None:
            cell_value = str(row[idx])
        if cell_value == value:
            # 見つかった行の値を列名と対応させた辞書に変換して返します。
            return {h: (str(row[i]) if i < len(row) and row[i] is not None else "")
                    for i, h in enumerate(headers) if h}

    # 一致する行が無かった場合は None を返します。
    return None


# === マテリアル風のカード ===
class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setStyleSheet("""
            QFrame#card {
                background: #FFFFFF;
                border-radius: 12px;
                border: 1px solid rgba(0,0,0,0.08);
            }
        """)


# === メイン画面 ===
class MainWindow(QMainWindow):
    def __init__(self, layout_path: str):
        super().__init__()

        self.config = self._load_layout(layout_path)
        # JSON全体に共通の文字サイズがあれば読み込み、無ければ12を使います。
        self.default_font_size = int(self.config.get("font_size", 12))

        win = self.config.get("window", {})
        self.setWindowTitle(win.get("title", "フォーム"))
        self.resize(int(win.get("width", 980)), int(win.get("height", 680)))

        self.excel_sheet = self.config.get("excel", {}).get("sheet", "受注データ")
        # データファイルのパスを保存しておくための専用ファイルの位置を決めます。
        self.path_store = os.path.join(os.path.dirname(layout_path), "data_file_path.txt")
        # 既に保存済みのパスを読み込むか、無い場合はダイアログで一度だけ選んでもらいます。
        self.current_xlsm: Optional[str] = self.load_xlsm_path()

        # 起動時に参照するデータを保存するための辞書を初期化します。
        self.preloaded_data: Dict[str, List[List[Any]]] = {}
        # Excel の初期データを読み込む処理を繰り返し試みます。
        while self.current_xlsm is not None:
            # 読み込み中であることを示す小さなウインドウを表示します。
            spinner = LoadingSpinner(self)
            spinner.show()
            QApplication.processEvents()

            # バックグラウンドで Excel を読み込む処理を用意します。
            container = {"data": {}, "error": None}

            def load() -> None:
                # スレッド内で Excel からデータを取得します。
                try:
                    container["data"] = extract_initial_data(self.current_xlsm)
                except Exception as e:
                    container["error"] = e

            thread = threading.Thread(target=load)
            thread.start()

            # 読み込みが終わるまでイベントを処理しながら待ちます。
            while thread.is_alive():
                QApplication.processEvents()
                QtCore.QThread.msleep(50)

            thread.join()
            spinner.close()

            # 結果を確認し、エラーであれば別のファイルを選び直します。
            if container["error"] is not None:
                QMessageBox.warning(
                    self,
                    "読み込みエラー",
                    f"初期データの読み込みに失敗しました: {container['error']}",
                )
                # ユーザーに別のファイルを選択してもらいます。
                new_path = self.ask_xlsm_path()
                if new_path:
                    # 選択されたパスを保存し、再度ループを試みます。
                    with open(self.path_store, "w", encoding="utf-8") as f:
                        f.write(new_path)
                    self.current_xlsm = new_path
                    continue
                # キャンセルされた場合は current_xlsm を None にしてループを終了します。
                self.current_xlsm = None
            else:
                # 正常に読み込めた場合は結果を保持してループを抜けます。
                self.preloaded_data = container["data"]
                break

        # 最終的にファイルが設定されていない場合は情報メッセージを表示します。
        if self.current_xlsm is None:
            QMessageBox.information(
                self,
                "設定情報",
                "データファイルが設定されていないため、読み書きは行えません。",
            )

        self.status = QStatusBar(self)
        self.setStatusBar(self.status)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)

        container = QWidget()
        root = QVBoxLayout(container)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(24)

        if "title" in self.config:
            title = QLabel(self.config["title"])
            title.setStyleSheet("font-size:22px; font-weight:600;")
            root.addWidget(title)

        self.card = Card()
        wrap = QVBoxLayout(self.card)
        wrap.setContentsMargins(24, 24, 24, 24)
        wrap.setSpacing(16)

        self.grid = QGridLayout()
        self.grid.setHorizontalSpacing(16)
        self.grid.setVerticalSpacing(12)
        ncols = int(self.config.get("grid_columns", 4))
        for i in range(ncols * 2):
            stretch = 1 if i % 2 == 1 else 0
            self.grid.setColumnStretch(i, stretch)

        self.widgets: Dict[str, QtWidgets.QWidget] = {}
        # 「保存」ボタンへの参照を保存する変数を用意します。
        # 初期値は何もない状態(None)とします。
        self.save_button: Optional[QPushButton] = None
        self._build_from_config(self.config.get("fields", []), ncols)

        # 直近に自動取得した品目番号を記録する変数を用意します。
        self._last_fetched_item: str = ""

        # 品目番号の入力内容に応じて自動でデータ取得とボタン状態の更新を行います。
        item_widget = self.widgets.get("品目番号")
        if isinstance(item_widget, QLineEdit):
            # 入力が変わるたびに専用の処理を呼び出します。
            item_widget.textChanged.connect(self.on_item_no_changed)
        # 起動直後にも一度状態を確認しておきます。
        self.update_button_states()
        # 色数の入力内容に応じてシリンダー欄を更新します。
        color_widget = self.widgets.get("色数")
        if isinstance(color_widget, QLineEdit):
            color_widget.textChanged.connect(self.on_color_count_changed)

        # シリンダーデータ登録欄を配置するための部品を用意します。
        self.cyl_title = QLabel("シリンダーデータ登録")
        self.cyl_title.setStyleSheet("font-weight:600;")
        self.cylinder_layout = QVBoxLayout()
        self.cylinder_layout.setSpacing(8)
        self.cylinder_units: List[CylinderUnit] = []

        wrap.addLayout(self.grid)
        wrap.addWidget(self.cyl_title)
        wrap.addLayout(self.cylinder_layout)
        root.addWidget(self.card)
        scroll.setWidget(container)
        self.setCentralWidget(scroll)

    def _load_layout(self, path: str) -> Dict[str, Any]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"設計図ファイルが見つかりません: {path}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _build_from_config(self, fields: list, ncols: int):
        for f in fields:
            ftype = f.get("type", "")
            row = int(f.get("row", 0))
            col = int(f.get("col", 0))
            col_span = int(f.get("col_span", 1))
            grid_col_label = col * 2
            grid_col_edit = col * 2 + 1
            grid_span = max(1, min(ncols - col, col_span)) * 2 - 1

            # 各部品に設定された文字サイズを取得し、無ければ共通設定を利用します。
            font_size = int(f.get("font_size", self.default_font_size))

            if ftype == "header":
                lbl = QLabel(f.get("text", ""))
                lbl.setStyleSheet(
                    f"font-size:{font_size}px; font-weight:600; color:#333;")
                self.grid.addWidget(lbl, row, 0, 1, ncols * 2)
                continue

            if ftype in ("line", "text"):
                label_text = f.get("label", "")
                key = f.get("key", label_text)

                lbl = QLabel(label_text)
                lbl.setStyleSheet(f"color:#333; font-size:{font_size}px;")
                self.grid.addWidget(lbl, row, grid_col_label)

                if ftype == "text":
                    edit = QPlainTextEdit()
                    # JSON で指定された高さを読み取り、その大きさに固定します。
                    h = int(f.get("height", 120))
                    edit.setFixedHeight(h)
                else:
                    val = f.get("validator", "")
                    # 数値専用欄では専用のラインエディットを利用します。
                    if val in ("int", "float"):
                        edit = NumericLineEdit()
                    else:
                        edit = QLineEdit()

                    if val == "int":
                        # 最小値と最大値を読み込み、整数用のバリデータを準備します。
                        imin = int(f.get("min", 0))
                        imax = int(f.get("max", 2147483647))
                        iv = QIntValidator(imin, imax)
                        # 英語ロケールを指定して全角数字を受け付けないようにします。
                        iv.setLocale(QtCore.QLocale("C"))
                        edit.setValidator(iv)
                        # フォーカス時に入力モードを半角英数字に固定し、半角数字のみを受け付けます。
                        edit.setInputMethodHints(QtCore.Qt.ImhDigitsOnly)
                    elif val == "float":
                        # 小数を扱う欄なので小数用のバリデータを設定します。
                        fmin = float(f.get("min", 0.0))
                        fmax = float(f.get("max", 1e12))
                        dec = int(f.get("decimals", 3))
                        v = QDoubleValidator(fmin, fmax, dec)
                        v.setNotation(QDoubleValidator.Notation.StandardNotation)
                        # 英語ロケールに固定して全角文字を排除します。
                        v.setLocale(QtCore.QLocale("C"))
                        edit.setValidator(v)
                        # フォーカス時に入力モードを半角英数字に固定し、数値と小数点のみを許可します。
                        edit.setInputMethodHints(QtCore.Qt.ImhPreferNumbers)

                edit.setStyleSheet(f"""
                    QLineEdit {{
                        padding: 8px 10px;
                        border-radius: 8px;
                        border:1px solid rgba(0,0,0,0.15);
                        background:#FAFAFA;
                        font-size:{font_size}px;
                    }}
                    QPlainTextEdit {{
                        padding: 10px;
                        border-radius: 10px;
                        border:1px solid rgba(0,0,0,0.15);
                        background:#FAFAFA;
                        font-size:{font_size}px;
                    }}
                    QLineEdit:focus, QPlainTextEdit:focus {{
                        border:1.4px solid #2962FF;
                        background:#FFFFFF;
                    }}
                """)

                # JSONで指定された横幅を取得し、0より大きければ固定幅を設定します。
                w = int(f.get("width", 0))
                if w > 0:
                    edit.setFixedWidth(w)
                self.grid.addWidget(edit, row, grid_col_edit, 1, grid_span)
                self.widgets[key] = edit
                continue

            if ftype == "button":
                text = f.get("text", "ボタン")
                action = f.get("action", "")
                btn = QPushButton(text)
                btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
                # ボタンの色や形を一括で指定します。無効化されているときは灰色にして、
                # 利用できない状態であることが直感的に分かるようにしています。
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background:#2962FF;
                        color:white;
                        padding:8px 16px;
                        border:none;
                        border-radius:8px;
                        font-weight:600;
                        font-size:{font_size}px;
                    }}
                    QPushButton:hover {{ background:#2F6BFF; }}
                    QPushButton:pressed {{ background:#2554CC; }}
                    QPushButton:disabled {{
                        background:#E0E0E0;
                        color:#9E9E9E;
                    }}
                """)
                # JSONで指定された横幅を取得し、0より大きければ固定幅を設定します。
                w = int(f.get("width", 0))
                if w > 0:
                    btn.setFixedWidth(w)
                self.grid.addWidget(btn, row, grid_col_edit,
                                    1, max(1, grid_span))
                if action == "save":
                    # 「保存」ボタンを後で参照できるよう保存します。
                    self.save_button = btn
                    btn.clicked.connect(self.on_save)
                elif action == "clear":
                    btn.clicked.connect(self.on_clear)
                elif action == "close":
                    btn.clicked.connect(self.close)
                continue

    def on_item_no_changed(self, text: str) -> None:
        """品目番号の入力が変わったときの共通処理です。"""
        # まずは保存ボタンの状態を更新します。
        self.update_button_states()

        # 前後の空白を取り除いた文字列を用意します。
        item_no = text.strip()
        # データファイルが設定されているかを確認します。
        has_file = self.current_xlsm is not None
        # 正規表現で8桁の数字のみを判定します。
        is_item_eight_digits = re.fullmatch(r"\d{8}", item_no) is not None

        # 条件を満たし、まだ同じ品目番号を読み込んでいない場合に自動で取得します。
        if has_file and is_item_eight_digits and item_no != self._last_fetched_item:
            self.on_fetch()
            self._last_fetched_item = item_no

    def update_button_states(self) -> None:
        """入力内容に応じて『保存』ボタンの状態を切り替えます。"""
        # 品目番号の入力欄を取り出し、未設定なら空文字として扱います。
        item_widget = self.widgets.get("品目番号")
        item_text = ""
        if isinstance(item_widget, QLineEdit):
            item_text = item_widget.text().strip()

        # データファイルが設定されているかを確認します。
        has_file = self.current_xlsm is not None

        # 「保存」ボタンは、品目番号が空でなく、かつデータファイルが設定されている場合のみ有効にします。
        is_save_valid = bool(item_text) and has_file

        # 保存ボタンに表示する文字列を決めます。既定では新規登録とします。
        save_text = "新規登録"
        if item_text and has_file:
            # 事前に読み込んだデータから同じ品目番号が存在するか調べます。
            sheet_data = self.preloaded_data.get(self.excel_sheet)
            if sheet_data:
                try:
                    exists = find_record_by_column(sheet_data, "品目番号", item_text)
                    if exists is not None:
                        save_text = "上書き保存"
                except Exception:
                    # 読み込みに失敗した場合は既定の表示のままとします。
                    pass

        if self.save_button is not None:
            # 保存ボタンの有効・無効を切り替えます。
            self.save_button.setEnabled(is_save_valid)
            # 判定結果に応じた文字列をボタンに表示します。
            self.save_button.setText(save_text)

    def on_color_count_changed(self, text: str) -> None:
        """色数に応じてシリンダー入力欄を増減させます。"""
        # 入力された文字を整数に変換し、無効な場合は0とします。
        count = int(text) if text.isdigit() else 0
        # 既存のシリンダー入力欄をすべて削除します。
        self._clear_cylinder_units()
        # 指定された数だけ新しいユニットを追加します。
        for _ in range(count):
            unit = CylinderUnit(self._get_item_no)
            self.cylinder_layout.addWidget(unit)
            self.cylinder_units.append(unit)
        # 番号の初期状態は1から始めます。
        self.update_color_numbers(1)
        # 先頭の番号欄だけは0への切り替えを受け付けます。
        if self.cylinder_units:
            first = self.cylinder_units[0].order_combo
            first.currentTextChanged.connect(self.on_first_color_changed)

    def _get_item_no(self) -> str:
        """現在入力されている品目番号を取得します。"""
        w = self.widgets.get("品目番号")
        if isinstance(w, QLineEdit):
            return w.text().strip()
        return ""

    def _clear_cylinder_units(self) -> None:
        """シリンダー入力欄をすべて取り除きます。"""
        while self.cylinder_layout.count():
            item = self.cylinder_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self.cylinder_units = []

    def update_color_numbers(self, start: int) -> None:
        """表示される色番号を0または1から順番に並べ直します。"""
        nums = [str(start + i) for i in range(len(self.cylinder_units))]
        for idx, unit in enumerate(self.cylinder_units):
            unit.order_combo.blockSignals(True)
            unit.order_combo.clear()
            unit.order_combo.addItems(nums)
            unit.order_combo.setCurrentIndex(idx)
            unit.order_combo.blockSignals(False)

    def on_first_color_changed(self, text: str) -> None:
        """先頭の色番号が0か1かで全体の番号を調整します。"""
        start = 0 if text.strip() == "0" else 1
        self.update_color_numbers(start)

    def collect_form_data(self) -> Dict[str, str]:
        d: Dict[str, str] = {}
        for k, w in self.widgets.items():
            if isinstance(w, QPlainTextEdit):
                d[k] = w.toPlainText().strip()
            elif isinstance(w, QLineEdit):
                d[k] = w.text().strip()
        return d

    def fill_form(self, data: Dict[str, str]) -> None:
        for k, w in self.widgets.items():
            v = data.get(k, "")
            if isinstance(w, QPlainTextEdit):
                w.setPlainText(v)
            elif isinstance(w, QLineEdit):
                w.setText(v)

    def load_xlsm_path(self) -> Optional[str]:
        """保存済みのデータファイルのパスを読み込みます。"""
        # まず保存用のテキストファイルが存在するか確認します。
        if os.path.exists(self.path_store):
            # ファイルがある場合は、中身を読み込んでそのまま返します。
            with open(self.path_store, "r", encoding="utf-8") as f:
                saved = f.read().strip()
            if saved:
                return saved
        # 保存されたパスが無いときは、ダイアログで一度だけ選んでもらいます。
        selected = self.ask_xlsm_path()
        if selected:
            # 選ばれたパスをファイルに書き込み、次回以降も使えるようにします。
            with open(self.path_store, "w", encoding="utf-8") as f:
                f.write(selected)
            return selected
        # 何も選ばなければ None を返します。
        return None

    def ask_xlsm_path(self) -> Optional[str]:
        """ユーザーに Excel ファイルを選んでもらうダイアログを表示します。"""
        path, _ = QFileDialog.getOpenFileName(
            self, "xlsm を選んでください", "", "Excel マクロ有効ブック (*.xlsm)")
        # 選ばれなければ空文字が返るので、その場合は None に変換します。
        return path or None

    @QtCore.Slot()
    def on_fetch(self):
        """『データ取得』ボタンが押されたときの処理です。"""
        # 品目番号の入力欄から文字列を取り出します。存在しなければ空文字です。
        item_no = ""
        if "品目番号" in self.widgets and isinstance(self.widgets["品目番号"], QLineEdit):
            item_no = self.widgets["品目番号"].text().strip()

        # 何も入力されていない場合は警告を出して処理を中断します。
        if not item_no:
            QMessageBox.warning(self, "入力エラー", "品目番号を入力してください。")
            return

        # 入力された文字列が8桁の数字でなければ警告を出して処理を中断します。
        if re.fullmatch(r"\d{8}", item_no) is None:
            QMessageBox.warning(self, "入力エラー", "品目番号は8桁の半角数字で入力してください。")
            return

        # 起動時に読み込んでおいたデータを使って該当行を探します。
        sheet_data = self.preloaded_data.get(self.excel_sheet)
        if not sheet_data:
            QMessageBox.warning(self, "データなし", "事前に読み込んだデータがありません。")
            return

        try:
            rec = find_record_by_column(sheet_data, "品目番号", item_no)
            if rec is None:
                # 見つからなかった場合は新規入力として扱います。
                QMessageBox.information(self, "見つかりません", "新規入力できます。")
                self.on_clear(keep_item=True)
            else:
                # 読み込んだ値を画面の入力欄に反映させます。
                filtered = {k: rec.get(k, "") for k in self.widgets.keys()}
                self.fill_form(filtered)
                self.status.showMessage("事前データから読み込みました。", 3000)
        except Exception as e:
            # 何らかのエラーが発生した場合はメッセージを表示します。
            QMessageBox.critical(self, "読み込みエラー", str(e))


    @QtCore.Slot()
    def on_save(self):
        data = self.collect_form_data()
        if not data.get("品目番号"):
            QMessageBox.warning(self, "入力エラー", "品目番号は必須です。")
            return
        # 起動時にファイルが設定されていない場合は、保存を中止して警告します。
        if self.current_xlsm is None:
            QMessageBox.warning(
                self,
                "設定エラー",
                "データファイルが設定されていないため、保存できません。"
            )
            return
        try:
            upsert_record_to_xlsm(self.current_xlsm, data, self.excel_sheet)
            self.status.showMessage("Excel に保存しました。", 3000)
            QMessageBox.information(self, "保存", "保存が完了しました。")

            # Excel に書き込んだ内容を事前読込データにも反映させます。
            sheet_data = self.preloaded_data.get(self.excel_sheet)
            if sheet_data:
                # 1 行目から列名の一覧を取得します。
                headers = [str(v) if v is not None else "" for v in sheet_data[0]]
                row = [data.get(h, "") for h in headers]
                if "品目番号" in headers:
                    idx = headers.index("品目番号")
                    # 既存の行を探し、あれば置き換え、無ければ追加します。
                    for i in range(1, len(sheet_data)):
                        cell = sheet_data[i][idx]
                        cell_text = str(cell) if cell is not None else ""
                        if cell_text == data.get("品目番号", ""):
                            sheet_data[i] = row
                            break
                    else:
                        sheet_data.append(row)

            # 保存後にボタンの表示や状態を更新します。
            self.update_button_states()
        except Exception as e:
            QMessageBox.critical(self, "保存エラー", str(e))

    @QtCore.Slot()
    def on_clear(self, keep_item: bool = False):
        for k, w in self.widgets.items():
            if keep_item and k == "品目番号":
                continue
            if isinstance(w, (QLineEdit, QPlainTextEdit)):
                w.clear()

        # 品目番号も消した場合は自動取得の記録をリセットします。
        if not keep_item:
            self._last_fetched_item = ""


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    layout_path = os.path.join(base, "layout.json")

    app = QApplication(sys.argv)
    apply_stylesheet(app, theme="light_blue.xml")
    # 無効状態の入力欄やボタンを灰色で表示するためのスタイルを追加します。
    # 利用できないことが見た目で分かるようにしています。
    app.setStyleSheet(app.styleSheet() + """
        QLineEdit:disabled,
        QPlainTextEdit:disabled,
        QPushButton:disabled {
            background-color: #E0E0E0;
            color: #9E9E9E;
        }
    """)

    w = MainWindow(layout_path)
    w.showMaximized()  # ← 起動時に最大化
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
