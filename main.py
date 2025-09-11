# -*- coding: utf-8 -*-
"""
小学生にもわかる説明：
  「layout.json」という設計図を読みこんで、画面の部品(ラベルや入力など)を
  決められた場所に並べます。Excel(xlsm)から読み書きもできます。
"""

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtGui import QIntValidator, QDoubleValidator, QRegularExpressionValidator
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPlainTextEdit, QPushButton, QFileDialog,
    QScrollArea, QFrame, QStatusBar, QMessageBox, QComboBox,
    QHBoxLayout, QListView, QStyledItemDelegate
)
from qt_material import apply_stylesheet
from openpyxl import load_workbook

import sys
import json
import os
import re
from typing import Dict, Optional, Any, List, Callable
import threading

# 半角数字を全角数字に直すためのテーブルを用意します
_FW_TABLE = str.maketrans("0123456789", "０１２３４５６７８９")


def to_full_width(num: int) -> str:
    """
    小学生にもわかる説明：
      ふつうの数字(半角)を、見た目が広い数字(全角)に変えて返します。
    """
    return str(num).translate(_FW_TABLE)

# Windows の IME を制御するための準備（他OSでは使いません）
_IS_WINDOWS = sys.platform.startswith("win")
if _IS_WINDOWS:
    import ctypes
    _imm32 = ctypes.windll.imm32  # ImmAssociateContext を使う
    HWND = ctypes.wintypes.HWND if hasattr(
        ctypes, "wintypes") else ctypes.c_void_p
    HIMC = ctypes.c_void_p
    _imm32.ImmAssociateContext.argtypes = [HWND, HIMC]
    _imm32.ImmAssociateContext.restype = HIMC


# =========================================
# コンボボックスを左寄せ＆はみ出し防止にする仕組み
# =========================================
class LeftAlignDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        # 小学生にもわかる説明：
        #   リストの文字を左にそろえ、長すぎると「…」にします。
        super().initStyleOption(option, index)
        option.displayAlignment = QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter
        option.textElideMode = QtCore.Qt.ElideRight


def setup_left_aligned_combo(combo: QComboBox) -> None:
    """
    小学生にもわかる説明：
      この関数はプルダウンを『左寄せ』に直し、
      リストが狭くて文字がはみ出すのを防ぎます。
    """
    originally_editable = combo.isEditable()

    view = QListView(combo)
    view.setUniformItemSizes(True)
    view.setTextElideMode(QtCore.Qt.ElideRight)
    view.setItemDelegate(LeftAlignDelegate(view))
    combo.setView(view)

    combo.setLayoutDirection(QtCore.Qt.LeftToRight)

    if not originally_editable:
        combo.setEditable(True)
        combo.lineEdit().setReadOnly(True)
    if combo.lineEdit() is not None:
        combo.lineEdit().setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)

    try:
        width_hint = max(view.sizeHintForColumn(
            0), combo.view().sizeHintForColumn(0)) + 32
        view.setMinimumWidth(width_hint)
    except Exception:
        pass
# =========================================


# 数字入力専用のラインエディットです。フォーカス中は IME を完全に効かなくします。
class NumericLineEdit(QLineEdit):
    """
    小学生にもわかる説明：
      この入力欄は数字用です。カーソルが入っている間は
      日本語入力(IME)を使えないようにします。
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._prev_hints: Optional[QtCore.Qt.InputMethodHints] = None
        self._prev_ime_enabled: Optional[bool] = None
        # Windows 専用：元の IME ハンドルを覚えておく場所
        self._prev_himc = None

    def _hwnd(self) -> Optional[int]:
        if not _IS_WINDOWS:
            return None
        wid = self.winId()
        try:
            # PySide6 の winId() は sip.voidptr → int へ
            return int(wid)
        except Exception:
            return None

    def focusInEvent(self, event: QtGui.QFocusEvent) -> None:
        # いまのヒントと IME 状態を記録します
        self._prev_hints = self.inputMethodHints()
        self._prev_ime_enabled = self.testAttribute(
            QtCore.Qt.WA_InputMethodEnabled)

        # 半角英数字優先（数値入力向け）ヒント
        self.setInputMethodHints(
            self._prev_hints | QtCore.Qt.ImhLatinOnly | QtCore.Qt.ImhPreferNumbers)

        # Qt 側のIMEを無効（保険）
        self.setAttribute(QtCore.Qt.WA_InputMethodEnabled, False)

        # Windows では OS 側の IME をこのウィジェットから切り離します
        if _IS_WINDOWS:
            hwnd = self._hwnd()
            if hwnd:
                # 0 を関連付けると IME 無効化。戻り値が元の IME ハンドル
                self._prev_himc = _imm32.ImmAssociateContext(hwnd, HIMC(0))

        super().focusInEvent(event)

    def focusOutEvent(self, event: QtGui.QFocusEvent) -> None:
        # フォーカスが外れたら元の設定に戻します
        if self._prev_hints is not None:
            self.setInputMethodHints(self._prev_hints)
        if self._prev_ime_enabled is not None:
            self.setAttribute(QtCore.Qt.WA_InputMethodEnabled,
                              self._prev_ime_enabled)

        if _IS_WINDOWS:
            hwnd = self._hwnd()
            if hwnd:
                # 元の IME を戻します（他の欄で使えるように）
                _imm32.ImmAssociateContext(hwnd, self._prev_himc or HIMC(0))
                self._prev_himc = None

        super().focusOutEvent(event)

    def inputMethodEvent(self, event: QtGui.QInputMethodEvent) -> None:
        # 念のため：IME からの合成入力は無視（保険）
        event.ignore()


# === 起動時のエクセル読み込み待機ウインドウ ===
class LoadingSpinner(QtWidgets.QDialog):
    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.Dialog | QtCore.Qt.FramelessWindowHint)
        self.setFixedSize(160, 160)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        self.indicator = _SpinnerWidget(self)
        layout.addWidget(self.indicator, alignment=QtCore.Qt.AlignCenter)

        label = QLabel("エクセルファイル\n読み込み中......", self)
        label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(label, alignment=QtCore.Qt.AlignCenter)


class _SpinnerWidget(QtWidgets.QWidget):
    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self._angle = 0
        self._pen_width = 8
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._rotate)
        self._timer.start(16)
        self.setFixedSize(64, 64)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)

    def _rotate(self) -> None:
        self._angle = (self._angle + 5) % 360
        self.update()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        rect = self.rect()
        painter.translate(rect.center())
        painter.rotate(self._angle)
        radius = min(rect.width(), rect.height()) / 2 - self._pen_width
        pen = QtGui.QPen(QtGui.QColor(0, 0, 255), self._pen_width)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        painter.setPen(pen)
        painter.drawArc(QtCore.QRectF(-radius, -radius,
                        radius * 2, radius * 2), 0, 270 * 16)


# === シリンダー入力ユニット ===
class CylinderUnit(QWidget):
    def __init__(
        self,
        get_item_no: Callable[[], str],
        get_candidates: Callable[[str], List[str]],
    ) -> None:
        """シリンダー情報を1行分まとめる部品です。"""
        super().__init__()
        # 品目番号を取得する関数を保持します
        self._get_item_no = get_item_no
        # シリンダー候補を取得する関数を保持します
        self._get_candidates = get_candidates

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        row = QHBoxLayout()

        # 〇色目
        self.order_combo = QComboBox()
        self.order_combo.addItems(["0", "1"])
        self.order_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToContents)
        row.addWidget(self.order_combo)

        # シリンダー番号（編集可）
        self.cylinder_combo = QComboBox()
        self.cylinder_combo.setEditable(True)
        self.cylinder_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToContents)
        # 9桁の数字のみ入力できるようにします
        regex = QtCore.QRegularExpression(r"\d{9}")
        validator = QRegularExpressionValidator(regex)
        line_edit = self.cylinder_combo.lineEdit()
        if line_edit is not None:
            line_edit.setValidator(validator)
            line_edit.setInputMethodHints(QtCore.Qt.ImhDigitsOnly)
        # Excel から読み取った候補を表示します
        self.refresh_cylinder_list()
        row.addWidget(self.cylinder_combo)

        # 色名
        self.color_edit = QLineEdit()
        row.addWidget(self.color_edit)

        # ベタ巾（数値）
        self.width_edit = NumericLineEdit()
        dv = QDoubleValidator(0.0, 1e12, 3)
        dv.setNotation(QDoubleValidator.Notation.StandardNotation)
        dv.setLocale(QtCore.QLocale("C"))
        self.width_edit.setValidator(dv)
        self.width_edit.setInputMethodHints(QtCore.Qt.ImhPreferNumbers)
        row.addWidget(self.width_edit)

        # 旧版処理
        self.process_combo = QComboBox()
        self.process_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.process_combo.addItems([
            "変更無し",
            "同名製版",
            "落組行き",
            "名義変更",
            "廃棄行き",
        ])
        row.addWidget(self.process_combo)

        # 左寄せ＆はみ出し防止を適用
        setup_left_aligned_combo(self.order_combo)
        setup_left_aligned_combo(self.cylinder_combo)
        setup_left_aligned_combo(self.process_combo)

        layout.addLayout(row)

        # 名義変更時のみ表示する欄（数値）
        self.rename_edit = NumericLineEdit()
        iv = QIntValidator(0, 999999999)
        iv.setLocale(QtCore.QLocale("C"))
        self.rename_edit.setValidator(iv)
        self.rename_edit.setInputMethodHints(QtCore.Qt.ImhDigitsOnly)
        self.rename_edit.setPlaceholderText("名義変更先の番号")
        self.rename_edit.hide()
        layout.addWidget(self.rename_edit)

        self.process_combo.currentTextChanged.connect(self._on_process_changed)

    def refresh_cylinder_list(self) -> None:
        """Excel シートから取得したシリンダー候補を表示し直します。"""
        # いったんリストを空にします
        self.cylinder_combo.clear()
        # 現在入力されている品目番号を取得します（候補自体は品目番号に依存しません）
        item_no = self._get_item_no()
        # すべての候補を取得してプルダウンに追加します
        candidates = self._get_candidates(item_no)
        self.cylinder_combo.addItems(candidates)

    def _on_process_changed(self, text: str) -> None:
        """旧版処理の内容によって追加欄を表示します。"""
        self.rename_edit.setVisible(text == "名義変更")


# === Excel の読み書き関数（省略なし・そのまま） ===
def read_record_from_xlsm(path: str, item_no: str, sheet_name: str) -> Optional[Dict[str, str]]:
    wb = load_workbook(path, keep_vba=True, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シート『{sheet_name}』がありません")
    ws = wb[sheet_name]

    header_map: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v)] = c

    if "品目番号" not in header_map:
        raise ValueError("『品目番号』の列が見つかりません")

    col_item = header_map["品目番号"]
    target_row = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=col_item).value or "") == item_no:
            target_row = r
            break
    if target_row is None:
        return None

    return {name: str(ws.cell(row=target_row, column=header_map.get(name)).value or "")
            for name in header_map.keys()}


def upsert_record_to_xlsm(path: str, data: Dict[str, str], sheet_name: str) -> None:
    wb = load_workbook(path, keep_vba=True, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シート『{sheet_name}』がありません")
    ws = wb[sheet_name]

    header_map: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v)] = c

    if "品目番号" not in header_map:
        raise ValueError("『品目番号』の列が見つかりません")

    col_item = header_map["品目番号"]
    target_row = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=col_item).value or "") == data.get("品目番号", ""):
            target_row = r
            break
    if target_row is None:
        target_row = ws.max_row + 1

    for k, v in data.items():
        c = header_map.get(k)
        if c is None:
            continue
        ws.cell(row=target_row, column=c).value = v

    wb.save(path)


# === 起動時データ抽出関数（省略なし・そのまま） ===
def extract_initial_data(path: str, progress: Optional[Callable[[int, int], None]] = None) -> Dict[str, List[List[Any]]]:
    wb = load_workbook(path, keep_vba=True, data_only=False)
    result: Dict[str, List[List[Any]]] = {}
    sheets = ("受注データ", "シリンダーデータ")
    total = len(sheets)
    for idx, sheet in enumerate(sheets, start=1):
        if sheet in wb.sheetnames:
            ws = wb[sheet]
            result[sheet] = _extract_range_from_sheet(ws)
        if progress is not None:
            progress(idx, total)
    return result


def _extract_range_from_sheet(ws) -> List[List[Any]]:
    max_col = ws.max_column
    while max_col > 0 and ws.cell(row=1, column=max_col).value in (None, ""):
        max_col -= 1

    header_map: Dict[str, int] = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            header_map[str(v)] = c

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

    data: List[List[Any]] = []
    for r in range(1, last_row + 1):
        row_values: List[Any] = []
        for c in range(1, max_col + 1):
            row_values.append(ws.cell(row=r, column=c).value)
        data.append(row_values)
    return data


def find_record_by_column(data: List[List[Any]], column: str, value: str) -> Optional[Dict[str, str]]:
    if not data:
        return None
    headers = [str(v) if v is not None else "" for v in data[0]]
    if column not in headers:
        raise ValueError(f"『{column}』の列が見つかりません")
    idx = headers.index(column)
    for row in data[1:]:
        cell_value = ""
        if idx < len(row) and row[idx] is not None:
            cell_value = str(row[idx])
        if cell_value == value:
            return {h: (str(row[i]) if i < len(row) and row[i] is not None else "")
                    for i, h in enumerate(headers) if h}
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
        self.default_font_size = int(self.config.get("font_size", 12))

        win = self.config.get("window", {})
        self.setWindowTitle(win.get("title", "フォーム"))
        self.resize(int(win.get("width", 980)), int(win.get("height", 680)))

        self.excel_sheet = self.config.get("excel", {}).get("sheet", "受注データ")
        self.path_store = os.path.join(
            os.path.dirname(layout_path), "data_file_path.txt")
        self.current_xlsm: Optional[str] = self.load_xlsm_path()

        self.preloaded_data: Dict[str, List[List[Any]]] = {}
        while self.current_xlsm is not None:
            spinner = LoadingSpinner(self)
            spinner.show()
            QApplication.processEvents()

            container = {"data": {}, "error": None}

            def load() -> None:
                try:
                    container["data"] = extract_initial_data(self.current_xlsm)
                except Exception as e:
                    container["error"] = e

            thread = threading.Thread(target=load)
            thread.start()

            while thread.is_alive():
                QApplication.processEvents()
                QtCore.QThread.msleep(50)

            thread.join()
            spinner.close()

            if container["error"] is not None:
                QMessageBox.warning(
                    self,
                    "読み込みエラー",
                    f"初期データの読み込みに失敗しました: {container['error']}",
                )
                new_path = self.ask_xlsm_path()
                if new_path:
                    with open(self.path_store, "w", encoding="utf-8") as f:
                        f.write(new_path)
                    self.current_xlsm = new_path
                    continue
                self.current_xlsm = None
            else:
                self.preloaded_data = container["data"]
                break

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
        self.save_button: Optional[QPushButton] = None
        self._build_from_config(self.config.get("fields", []), ncols)

        self._last_fetched_item: str = ""

        item_widget = self.widgets.get("品目番号")
        if isinstance(item_widget, QLineEdit):
            item_widget.textChanged.connect(self.on_item_no_changed)
        self.update_button_states()

        color_widget = self.widgets.get("色数")
        if isinstance(color_widget, QLineEdit):
            color_widget.textChanged.connect(self.on_color_count_changed)

        self.cyl_title = QLabel("シリンダーデータ登録")
        self.cyl_title.setStyleSheet("font-weight:600;")

        self.cylinder_layout = QVBoxLayout()
        self.cylinder_layout.setSpacing(8)

        self.cylinder_header = QHBoxLayout()
        header_labels = ["〇色目", "シリンダー番号", "色名", "ベタ巾", "旧版処理"]
        for text in header_labels:
            lbl = QLabel(text)
            lbl.setStyleSheet("font-weight:600;")
            self.cylinder_header.addWidget(lbl)
        self.cylinder_layout.addLayout(self.cylinder_header)

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
                    h = int(f.get("height", 120))
                    edit.setFixedHeight(h)
                else:
                    val = f.get("validator", "")
                    if val in ("int", "float"):
                        edit = NumericLineEdit()
                    else:
                        edit = QLineEdit()

                    if val == "int":
                        imin = int(f.get("min", 0))
                        imax = int(f.get("max", 2147483647))
                        iv = QIntValidator(imin, imax)
                        iv.setLocale(QtCore.QLocale("C"))
                        edit.setValidator(iv)
                        edit.setInputMethodHints(QtCore.Qt.ImhDigitsOnly)
                    elif val == "float":
                        fmin = float(f.get("min", 0.0))
                        fmax = float(f.get("max", 1e12))
                        dec = int(f.get("decimals", 3))
                        v = QDoubleValidator(fmin, fmax, dec)
                        v.setNotation(
                            QDoubleValidator.Notation.StandardNotation)
                        v.setLocale(QtCore.QLocale("C"))
                        edit.setValidator(v)
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
                w = int(f.get("width", 0))
                if w > 0:
                    btn.setFixedWidth(w)
                self.grid.addWidget(btn, row, grid_col_edit,
                                    1, max(1, grid_span))
                if action == "save":
                    self.save_button = btn
                    btn.clicked.connect(self.on_save)
                elif action == "clear":
                    btn.clicked.connect(self.on_clear)
                elif action == "close":
                    btn.clicked.connect(self.close)
                continue

    def on_item_no_changed(self, text: str) -> None:
        """品目番号の入力が変わったときの共通処理です。"""
        self.update_button_states()

        # 既存のシリンダー入力欄の候補を更新します
        for unit in self.cylinder_units:
            unit.refresh_cylinder_list()

        item_no = text.strip()
        has_file = self.current_xlsm is not None
        is_item_eight_digits = re.fullmatch(r"\d{8}", item_no) is not None

        if has_file and is_item_eight_digits and item_no != self._last_fetched_item:
            self.on_fetch()
            self._last_fetched_item = item_no

    def update_button_states(self) -> None:
        """入力内容に応じて『保存』ボタンの状態を切り替えます。"""
        item_widget = self.widgets.get("品目番号")
        item_text = ""
        if isinstance(item_widget, QLineEdit):
            item_text = item_widget.text().strip()

        has_file = self.current_xlsm is not None
        is_save_valid = bool(item_text) and has_file

        save_text = "新規登録"
        if item_text and has_file:
            sheet_data = self.preloaded_data.get(self.excel_sheet)
            if sheet_data:
                try:
                    exists = find_record_by_column(
                        sheet_data, "品目番号", item_text)
                    if exists is not None:
                        save_text = "上書き保存"
                except Exception:
                    pass

        if self.save_button is not None:
            self.save_button.setEnabled(is_save_valid)
            self.save_button.setText(save_text)

    def on_color_count_changed(self, text: str) -> None:
        """色数に応じてシリンダー入力欄を増減させます。"""
        count = int(text) if text.isdigit() else 0
        self._clear_cylinder_units()
        for _ in range(count):
            unit = CylinderUnit(self._get_item_no, self._get_cylinder_candidates)
            self.cylinder_layout.addWidget(unit)
            self.cylinder_units.append(unit)
        self.update_color_numbers(1)
        if self.cylinder_units:
            first = self.cylinder_units[0].order_combo
            first.currentTextChanged.connect(self.on_first_color_changed)

    def _get_item_no(self) -> str:
        """現在入力されている品目番号を取得します。"""
        w = self.widgets.get("品目番号")
        if isinstance(w, QLineEdit):
            return w.text().strip()
        return ""

    def _get_cylinder_candidates(self, item_no: str) -> List[str]:
        """Excel シートから『品目番号+刷順列』のすべての値を集めます。

        品目番号は引数として受け取りますが、候補の抽出には利用しません。
        """
        # 事前に読み込んだ「シリンダーデータ」シートを取得します
        data = self.preloaded_data.get("シリンダーデータ")
        if not data:
            return []
        # 1 行目のヘッダー文字列を取得します
        headers = [str(v) if v is not None else "" for v in data[0]]
        # シリンダー番号が記載された列名を探します
        cyl_header = "品目番号+刷順列"
        if cyl_header not in headers:
            if "品目番号+刷順" in headers:
                cyl_header = "品目番号+刷順"
            else:
                return []
        idx_cyl = headers.index(cyl_header)
        result: List[str] = []
        seen = set()
        for row in data[1:]:
            # 各行の『品目番号+刷順列』の値を取り出します
            cyl_cell = (
                str(row[idx_cyl])
                if idx_cyl < len(row) and row[idx_cyl] is not None
                else ""
            )
            # 9 桁の数字のみを候補とし、重複は除きます
            if re.fullmatch(r"\d{9}", cyl_cell) and cyl_cell not in seen:
                result.append(cyl_cell)
                seen.add(cyl_cell)
        return result

    def _clear_cylinder_units(self) -> None:
        """シリンダー入力欄をすべて取り除きます。"""
        while self.cylinder_layout.count() > 1:
            item = self.cylinder_layout.takeAt(1)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self.cylinder_units = []

    def update_color_numbers(self, start: int) -> None:
        """表示される色番号を0または1から順番に並べ直します。"""

        # すべての行に対して、先頭から順番に色番号を設定します
        for idx, unit in enumerate(self.cylinder_units):
            unit.order_combo.blockSignals(True)  # 値を書き換える間はシグナルを止めます
            unit.order_combo.clear()             # 以前の選択肢を消します

            # 表示する番号を計算します（start が0なら0から、1なら1から）
            number = start + idx

            if idx == 0:
                # 先頭の行だけは0と1を選べるようにして、初期値を設定します
                unit.order_combo.addItems(["0", "1"])
                unit.order_combo.setCurrentText(str(number))
                unit.order_combo.setEnabled(True)
            else:
                # 2行目以降は計算した番号を表示し、変更できないようにします
                unit.order_combo.addItem(str(number))
                unit.order_combo.setCurrentIndex(0)
                unit.order_combo.setEnabled(False)

            unit.order_combo.blockSignals(False)  # シグナルを再び有効にします

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
        if os.path.exists(self.path_store):
            with open(self.path_store, "r", encoding="utf-8") as f:
                saved = f.read().strip()
            if saved:
                return saved
        selected = self.ask_xlsm_path()
        if selected:
            with open(self.path_store, "w", encoding="utf-8") as f:
                f.write(selected)
            return selected
        return None

    def ask_xlsm_path(self) -> Optional[str]:
        path, _ = QFileDialog.getOpenFileName(
            self, "xlsm を選んでください", "", "Excel マクロ有効ブック (*.xlsm)")
        return path or None

    @QtCore.Slot()
    def on_fetch(self):
        item_no = ""
        if "品目番号" in self.widgets and isinstance(self.widgets["品目番号"], QLineEdit):
            item_no = self.widgets["品目番号"].text().strip()

        if not item_no:
            QMessageBox.warning(self, "入力エラー", "品目番号を入力してください。")
            return

        if re.fullmatch(r"\d{8}", item_no) is None:
            QMessageBox.warning(self, "入力エラー", "品目番号は8桁の半角数字で入力してください。")
            return

        sheet_data = self.preloaded_data.get(self.excel_sheet)
        if not sheet_data:
            QMessageBox.warning(self, "データなし", "事前に読み込んだデータがありません。")
            return

        try:
            rec = find_record_by_column(sheet_data, "品目番号", item_no)
            if rec is None:
                QMessageBox.information(self, "見つかりません", "新規入力できます。")
                self.on_clear(keep_item=True)
            else:
                filtered = {k: rec.get(k, "") for k in self.widgets.keys()}
                self.fill_form(filtered)
                # シリンダー番号を画面に反映します
                #   Excel の「０色目シリンダー」〜「１０色目シリンダー」の値を
                #   調べて、対応する入力欄に順番に入れます。
                #   ここでは先頭の「０色目シリンダー」に値があるか確認し、
                #   あれば色番号を０から開始します。なければ色番号を１から
                #   開始し、「０色目シリンダー」の値は無視します。
                zero_key = f"{to_full_width(0)}色目シリンダー"  # 「０色目シリンダー」の列名を作ります
                start = 0 if rec.get(zero_key, "") else 1  # 0番目が空かどうかで開始番号を決めます
                # 〇色目プルダウンの表示を０または１から始まるように整えます
                self.update_color_numbers(start)
                for idx, unit in enumerate(self.cylinder_units):
                    # 列名に使う数字を全角に変え、必要に応じてずらします
                    key = f"{to_full_width(idx + start)}色目シリンダー"
                    value = rec.get(key, "")
                    value = "" if value is None else str(value)
                    line = unit.cylinder_combo.lineEdit()
                    if line is not None:
                        line.setText(value)
                self.status.showMessage("事前データから読み込みました。", 3000)
        except Exception as e:
            QMessageBox.critical(self, "読み込みエラー", str(e))

    @QtCore.Slot()
    def on_save(self):
        data = self.collect_form_data()
        if not data.get("品目番号"):
            QMessageBox.warning(self, "入力エラー", "品目番号は必須です。")
            return
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

            sheet_data = self.preloaded_data.get(self.excel_sheet)
            if sheet_data:
                headers = [
                    str(v) if v is not None else "" for v in sheet_data[0]]
                row = [data.get(h, "") for h in headers]
                if "品目番号" in headers:
                    idx = headers.index("品目番号")
                    for i in range(1, len(sheet_data)):
                        cell = sheet_data[i][idx]
                        cell_text = str(cell) if cell is not None else ""
                        if cell_text == data.get("品目番号", ""):
                            sheet_data[i] = row
                            break
                    else:
                        sheet_data.append(row)

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

        if not keep_item:
            self._last_fetched_item = ""


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    layout_path = os.path.join(base, "layout.json")

    app = QApplication(sys.argv)
    apply_stylesheet(app, theme="light_blue.xml")
    # 無効状態の入力欄やボタンを灰色に、プルダウンの文字色を黒にします。
    app.setStyleSheet(app.styleSheet() + """
        QLineEdit:disabled,
        QPlainTextEdit:disabled,
        QPushButton:disabled {
            background-color: #E0E0E0;
            color: #9E9E9E;
        }
        QComboBox {
            color: #000000;
        }
    """)

    w = MainWindow(layout_path)
    w.showMaximized()  # ← 起動時に最大化
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
