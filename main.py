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
    QScrollArea, QFrame, QStatusBar, QMessageBox, QProgressDialog
)
from qt_material import apply_stylesheet
from openpyxl import load_workbook

import sys
import json
import os
import re
from typing import Dict, Optional, Any, List, Callable


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


# === Excel の読み書き関数 ===
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
        # ファイルが選択されなかった場合は、後の操作ができない旨を表示します。
        if self.current_xlsm is None:
            QMessageBox.information(
                self,
                "設定情報",
                "データファイルが設定されていないため、読み書きは行えません。"
            )

        # 起動時に参照するデータを保存するための辞書を初期化します。
        self.preloaded_data: Dict[str, List[List[Any]]] = {}
        if self.current_xlsm is not None:
            # 進捗を表示するためのダイアログを準備します。
            progress_dialog = QProgressDialog(
                "Excel からデータを読み込んでいます...",
                "",
                0,
                2,
                self,
            )
            progress_dialog.setWindowTitle("データ取得中")
            progress_dialog.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
            progress_dialog.setMinimumDuration(0)
            progress_dialog.setValue(0)
            progress_dialog.show()
            # ここでイベントを処理してダイアログをすぐに表示させます。
            QApplication.processEvents()

            # コールバック関数を定義し、シート読込ごとに進捗を更新します。
            def update_progress(current: int, total: int) -> None:
                progress_dialog.setMaximum(total)
                progress_dialog.setValue(current)
                QApplication.processEvents()

            try:
                # Excel ファイルから指定された二つのシートを読み込みます。
                self.preloaded_data = extract_initial_data(
                    self.current_xlsm, update_progress
                )
            except Exception as e:
                # 読み込みに失敗した場合は警告を表示し、空のデータを保持します。
                QMessageBox.warning(
                    self,
                    "読み込みエラー",
                    f"初期データの読み込みに失敗しました: {e}",
                )
                self.preloaded_data = {}
            finally:
                # 進捗ダイアログを閉じて画面を元に戻します。
                progress_dialog.close()

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
        # 「データ取得」ボタンと「保存」ボタンへの参照を保存する変数を用意します。
        # 初期値は何もない状態(None)とします。
        self.fetch_button: Optional[QPushButton] = None
        self.save_button: Optional[QPushButton] = None
        self._build_from_config(self.config.get("fields", []), ncols)

        # 品目番号の入力内容に応じて各ボタンの有効・無効を切り替える設定を行います。
        item_widget = self.widgets.get("品目番号")
        if isinstance(item_widget, QLineEdit):
            # 入力が変わるたびに状態を更新するよう signal を接続します。
            item_widget.textChanged.connect(self.update_button_states)
        # 起動直後にも一度状態を確認しておきます。
        self.update_button_states()

        wrap.addLayout(self.grid)
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
                if action == "fetch":
                    # 「データ取得」ボタンを後で参照できるよう保存します。
                    self.fetch_button = btn
                    btn.clicked.connect(self.on_fetch)
                elif action == "save":
                    # 「保存」ボタンを後で参照できるよう保存します。
                    self.save_button = btn
                    btn.clicked.connect(self.on_save)
                elif action == "clear":
                    btn.clicked.connect(self.on_clear)
                elif action == "close":
                    btn.clicked.connect(self.close)
                continue

    def update_button_states(self) -> None:
        """入力内容に応じて『データ取得』と『保存』のボタンを切り替えます。"""
        # 品目番号の入力欄を取り出し、未設定なら空文字として扱います。
        item_widget = self.widgets.get("品目番号")
        item_text = ""
        if isinstance(item_widget, QLineEdit):
            item_text = item_widget.text().strip()

        # データファイルが設定されているかを確認します。
        has_file = self.current_xlsm is not None

        # 正規表現を使って「8桁の数字」のみを判定します。
        # r"\d{8}" は「半角数字が8文字続く」という意味です。
        is_item_eight_digits = re.fullmatch(r"\d{8}", item_text) is not None

        # 「データ取得」ボタンは、8桁の数字が入力され、かつデータファイルが設定されている場合のみ有効にします。
        is_fetch_valid = is_item_eight_digits and has_file
        if self.fetch_button is not None:
            self.fetch_button.setEnabled(is_fetch_valid)

        # 「保存」ボタンは、品目番号が空でなく、かつデータファイルが設定されている場合のみ有効にします。
        is_save_valid = bool(item_text) and has_file
        if self.save_button is not None:
            self.save_button.setEnabled(is_save_valid)

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
        item = ""
        if "品目番号" in self.widgets and isinstance(self.widgets["品目番号"], QLineEdit):
            item = self.widgets["品目番号"].text().strip()

        # 何も入力されていない場合は警告を出して処理を中断します。
        if not item:
            QMessageBox.warning(self, "入力エラー", "品目番号を入力してください。")
            return

        # 入力された文字列が8桁の数字でなければ警告を出して処理を中断します。
        if re.fullmatch(r"\d{8}", item) is None:
            QMessageBox.warning(self, "入力エラー", "品目番号は8桁の半角数字で入力してください。")
            return

        # 起動時にファイルが設定されていない場合は、読み込みを中止して警告します。
        if self.current_xlsm is None:
            QMessageBox.warning(
                self,
                "設定エラー",
                "データファイルが設定されていないため、読み込みできません。",
            )
            return

        # Excel ファイルから該当するレコードを読み込みます。
        try:
            rec = read_record_from_xlsm(
                self.current_xlsm, item, self.excel_sheet)
            if rec is None:
                # 見つからなかった場合は新規入力として扱います。
                QMessageBox.information(self, "見つかりません", "新規入力できます。")
                self.on_clear(keep_item=True)
            else:
                # 読み込んだ値を画面の入力欄に反映させます。
                filtered = {k: rec.get(k, "") for k in self.widgets.keys()}
                self.fill_form(filtered)
                self.status.showMessage("Excel から読み込みました。", 3000)
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
        except Exception as e:
            QMessageBox.critical(self, "保存エラー", str(e))

    @QtCore.Slot()
    def on_clear(self, keep_item: bool = False):
        for k, w in self.widgets.items():
            if keep_item and k == "品目番号":
                continue
            if isinstance(w, (QLineEdit, QPlainTextEdit)):
                w.clear()


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
