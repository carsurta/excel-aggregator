# app.py – Excel Aggregator (Drag & Drop) – FINAL UI/Features
from __future__ import annotations
import sys, os, traceback, importlib
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

def _pd():
    return importlib.import_module("pandas")

from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QListWidget,
    QFileDialog, QHBoxLayout, QMessageBox, QTableView, QProgressBar, QCheckBox,
    QSplitter, QListWidgetItem, QDialog, QDialogButtonBox, QListWidget
)

from header_multirow import load_sheet_merge_aware, detect_header_band_and_build
from sheet_match import list_sheet_names, auto_match_with_headers
from validators import compute_violations

# ======= Meta columns (Korean, official tone) =======
META_MAP = {
    "file": "출처 파일명",
    "sheet": "출처 시트명",
    "header_row": "헤더 행 번호(원본)",
    "violations": "유효성 점검 메모",
}
META_COLS_KR = list(META_MAP.values())

class _FileRow(QWidget):
    def __init__(self, parent, path: str, on_remove, on_choose):
        super().__init__(parent)
        self.path = path
        from PyQt6.QtWidgets import QHBoxLayout, QPushButton, QLabel, QSizePolicy
        from PyQt6.QtCore import Qt

        hl = QHBoxLayout(self)
        hl.setContentsMargins(6, 2, 6, 2)
        hl.setSpacing(6)

        # 좌측 X 버튼
        self.btn_x = QPushButton("✕")
        self.btn_x.setFixedWidth(28)
        self.btn_x.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_x.clicked.connect(lambda: on_remove(self.path))
        hl.addWidget(self.btn_x)

        # 가운데 경로 라벨(길면 말줄임)
        self.lbl = QLabel(path)
        self.lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.lbl.setToolTip(path)
        hl.addWidget(self.lbl, 1)

        # 우측 시트 버튼
        self.btn_sheet = QPushButton("시트 선택")
        self.btn_sheet.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_sheet.clicked.connect(lambda: on_choose(self.path, True))
        hl.addWidget(self.btn_sheet)

# ============== Qt Model for pandas ==================
class PandasModel(QAbstractTableModel):
    def __init__(self, df):
        super().__init__()
        self._df = df

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._df.columns)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            val = self._df.iat[index.row(), index.column()]
            try:
                import pandas as pd
                if pd.isna(val):
                    return ""
            except Exception:
                pass
            return "" if val is None else str(val)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            try:
                return str(self._df.columns[section])
            except Exception:
                return ""
        else:
            return str(section + 1)

# ================= Data structures ====================
@dataclass
class ParsedSheet:
    file_path: str
    sheet_name: str
    header_row: int
    columns: List[str]
    data: "pandas.DataFrame"

# ==================== Main UI =========================
class ExcelAggregator(QWidget):

    def _on_meta_label_clicked(self, event):
        self.chk_meta.toggle()
        self._refresh_preview()
        
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Aggregator – Drag & Drop")
        self.setAcceptDrops(True)
        self.resize(1200, 780)

        # preferences: sheet names chosen in the first file
        self.pref_sheet_names: List[str] = []  # used for auto-preselect on later files
        # selected sheets per file
        self.file_sheets: Dict[str, List[str]] = {}

        # ---- Layout: vertical splitter (resizable list + preview) ----
        root = QVBoxLayout(self)
        self.info = QLabel("xlsx 파일/폴더를 드래그 앤 드롭하세요.")
        self.info.setStyleSheet("font-size:14px;color:#333;margin:6px 4px;")
        root.addWidget(self.info)

        splitter = QSplitter(Qt.Orientation.Vertical)
        root.addWidget(splitter, 1)

        # Top: file list area (with Clear & per-item remove via Del key)
        top_widget = QWidget(); top_layout = QVBoxLayout(top_widget); top_layout.setContentsMargins(0,0,0,0)
        self.list = QListWidget()
        self.list.setStyleSheet("QListWidget{border:1px dashed #999;padding:8px;}")
        top_layout.addWidget(self.list, 1)

        hl = QHBoxLayout(); hl.setContentsMargins(4,4,4,0)
        self.btn_add = QPushButton("파일 불러오기")
        self.btn_add.clicked.connect(self._browse_files)
        hl.addWidget(self.btn_add)
        self.btn_clear = QPushButton("모두 제거")
        self.btn_clear.clicked.connect(self._clear)
        hl.addWidget(self.btn_clear)
        hl.addStretch(1)
        top_layout.addLayout(hl)

        splitter.addWidget(top_widget)

        # Bottom: preview table + bottom bar (meta checkbox + Save)
        bottom_widget = QWidget(); bottom_layout = QVBoxLayout(bottom_widget); bottom_layout.setContentsMargins(0,0,0,0)
        self.progress = QProgressBar(); self.progress.setVisible(False)
        bottom_layout.addWidget(self.progress)

        self.table = QTableView()
        bottom_layout.addWidget(self.table, 1)

        bottom_bar = QHBoxLayout()
        self.chk_meta = QCheckBox(); self.chk_meta.setChecked(False)
        self.lbl_meta = QLabel("원본 파일 정보 표시")
        self.lbl_meta.setStyleSheet("padding-left:4px;")
        self.lbl_meta.setCursor(Qt.CursorShape.PointingHandCursor)
        tooltip = (
            "출력 결과물에 다음 정보를 표기합니다:\n"
            f"• {META_MAP['file']}: 취합의 출처가 된 원본 엑셀 파일 이름 (예: 자료제출_부서A.xlsx)\n"
            f"• {META_MAP['sheet']}: 해당 데이터가 있었던 원본 시트 이름 (예: 제출본, 입력데이터)\n"
            f"• {META_MAP['header_row']}: 원본 시트에서 컬럼명으로 판단된 헤더 행 번호 (예: 2)\n"
            f"• {META_MAP['violations']}: 전화/이메일 형식 오류, 빈값, 중복 등 간단한 유효성 점검 메모 (예: [연락처:전화형식] [이메일:형식] [번호:중복])\n\n"
            "체크를 해제하면 위 정보는 출력 파일에 포함되지 않으며 미리보기에서도 숨겨집니다."
        )
        self.chk_meta.setToolTip(tooltip); self.lbl_meta.setToolTip(tooltip)
        self.lbl_meta.mousePressEvent = self._on_meta_label_clicked
        self.chk_meta.stateChanged.connect(self._refresh_preview)
        bottom_bar.addWidget(self.chk_meta)
        bottom_bar.addWidget(self.lbl_meta)
        bottom_bar.addStretch(1)

        self.btn_save = QPushButton("출력 저장…")
        self.btn_save.clicked.connect(self._save_output)
        bottom_bar.addWidget(self.btn_save)
        bottom_layout.addLayout(bottom_bar)

        splitter.addWidget(bottom_widget)
        splitter.setStretchFactor(0, 0)  # top list minimal height by default
        splitter.setStretchFactor(1, 1)  # table expands

        # Data
        self.file_paths: List[str] = []
        self.parsed: List[ParsedSheet] = []
        self.combined = None  # pandas.DataFrame
        self.pref_sheet_names: List[str] = []        # 첫 파일에서 선택한 시트명들
        self.pref_headers: Dict[str, List[str]] = {} # 첫 파일에서 선택한 시트별 헤더 (합친 열명 리스트)

        # Remove key for selected list items
        self.list.keyPressEvent = self._list_keypress

    # ---------- Drag & Drop ----------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        added_files = []
        for u in urls:
            p = u.toLocalFile()
            if os.path.isdir(p):
                for root, _, files in os.walk(p):
                    for f in files:
                        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                            added_files.append(os.path.join(root, f))
            elif p.lower().endswith(".xlsx") and not os.path.basename(p).startswith("~$"):
                added_files.append(p)
        if not added_files:
            return

        # For each file, run sheet chooser with auto-preselection based on pref_sheet_names
        for path in added_files:
            self._open_sheet_chooser_for(path)

    def _list_keypress(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            for it in self.list.selectedItems():
                w = self.list.itemWidget(it)
                if w and hasattr(w, "path"):
                    self._remove_file(w.path)
        else:
            try:
                from PyQt6.QtWidgets import QListWidget
                QListWidget.keyPressEvent(self.list, event)
            except Exception:
                pass

    def _browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx)"
        )
        for path in paths or []:
            if path and path.lower().endswith(".xlsx") and not os.path.basename(path).startswith("~$"):
                self._open_sheet_chooser_for(path)

    # ---------- Sheet selection ----------
    def _add_with_sheet_selection(self, path: str):
        try:
            sheets = list_sheet_names(path)
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "시트 읽기 오류", f"{os.path.basename(path)}: 시트 목록을 읽을 수 없습니다.\n{e}")
            return

        # auto-preselect: match to preferred names if exists
        preselect = []
        if self.pref_sheet_names:
            base_norm = [s.strip().lower() for s in self.pref_sheet_names]
            for s in sheets:
                sn = s.strip().lower()
                if any(b in sn or sn in b for b in base_norm):
                    preselect.append(s)

        dlg = SheetChooser(self, path, sheets, preselect)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        chosen = dlg.selected_sheets()
        if not chosen:
            return

        if not self.file_paths and not self.pref_sheet_names:
            self.pref_sheet_names = chosen[:]

        if path not in self.file_paths:
            self.file_paths.append(path)
            self.file_sheets[path] = chosen
            self.list.addItem(path)
        else:
            self.file_sheets[path] = chosen

    def _add_list_item(self, path: str):
        from PyQt6.QtWidgets import QListWidgetItem
        # 이미 있으면 무시(혹은 덮어쓰기)
        for i in range(self.list.count()):
            it = self.list.item(i)
            w = self.list.itemWidget(it)
            if w and getattr(w, "path", "") == path:
                return it  # 이미 존재
        it = QListWidgetItem(self.list)
        w = _FileRow(self, path, self._remove_file, self._open_sheet_chooser_for)
        it.setSizeHint(w.sizeHint())
        self.list.addItem(it)
        self.list.setItemWidget(it, w)
        return it

    def _refresh_all_items(self):
        # 목록 전체를 현재 self.file_paths 순서대로 재구성
        self.list.clear()
        for p in self.file_paths:
            self._add_list_item(p)

    def _remove_file(self, path: str):
        self.file_paths = [p for p in self.file_paths if p != path]
        if path in self.file_sheets:
            del self.file_sheets[path]
        self._refresh_all_items()
        self._parse_and_preview()

    def _has_preview(self) -> bool:
        return bool(self.parsed)

    def _record_pref_from(self, path: str, sheets: list[str]):
        if self.file_paths or self.pref_sheet_names or not sheets:
            return
        self.pref_sheet_names = sheets[:]
        self.pref_headers = {}
        for s in sheets:
            try:
                raw0 = load_sheet_merge_aware(path, s)
                _st, _ed, _hdrs = detect_header_band_and_build(raw0)
                self.pref_headers[s] = _hdrs
            except Exception:
                self.pref_headers[s] = []

    def _finalize_file_selection(self, path: str, chosen: list[str]):
        self.file_sheets[path] = chosen[:]
        if path not in self.file_paths:
            self.file_paths.append(path)
            self._add_list_item(path)

    def _open_sheet_chooser_for(self, path: str, force_dialog: bool = False):
        try:
            sheets = list_sheet_names(path)
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "시트 읽기 오류", f"{os.path.basename(path)}: 시트 목록을 불러올 수 없습니다.\n{e}")
            return

        # 1) 각 시트의 헤더 미리 추출
        cand_name_headers = []
        for s in sheets:
            try:
                raw_preview = load_sheet_merge_aware(path, s)
                _st, _ed, _hdrs = detect_header_band_and_build(raw_preview)
                cand_name_headers.append((s, _hdrs))
            except Exception:
                cand_name_headers.append((s, []))

        # 2) preselect 계산: 기존 지식> (첫파일 로직+헤더 추천)
        preselect = []
        current = self.file_sheets.get(path, [])
        if current:
            preselect = current[:]
        elif getattr(self, "pref_sheet_names", None):
            base = [(bn, self.pref_headers.get(bn, [])) for bn in self.pref_sheet_names]
            preselect = auto_match_with_headers(base, cand_name_headers, threshold=0.55)

        # 3) 미리보기(기존 파일)가 있고 시트가 1개뿐이면 팝업 없이 즉시 적용
        auto_allowed = (not force_dialog) and self._has_preview()
        if auto_allowed and len(sheets) == 1:
            chosen = sheets[:]
            self._finalize_file_selection(path, chosen)
            self._parse_and_preview()
            return

        # 4) 후보가 0개 또는 2개 이상)팝업 유도: 여러 개면 여러 개를 체크상태로
        dlg = SheetChooser(self, path, sheets, preselect)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        chosen = dlg.selected_sheets()
        if not chosen:
            return

        # 첫파일의 첫 선택이면 기준 시트/헤더 저장
        self._record_pref_from(path, chosen)

        # 파일 목록/시트 매핑 반영
        self._finalize_file_selection(path, chosen)

        # 즉시 미리보기 반영
        self._parse_and_preview()
    # ---------- Core parse & combine ----------
    def _clear(self):
        self.file_paths.clear()
        self.file_sheets.clear()
        self.list.clear()
        self.parsed.clear()
        self.combined = None
        self.table.setModel(None)

    def _parse_and_preview(self):
        if not self.file_paths:
            self.table.setModel(None)
            return
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.progress.setMaximum(len(self.file_paths))

        parsed: List[ParsedSheet] = []
        for i, p in enumerate(self.file_paths, 1):
            try:
                sheets = self.file_sheets.get(p) or list_sheet_names(p)[:1]
                for sh in sheets:
                    raw = load_sheet_merge_aware(p, sh)
                    start, end, headers = detect_header_band_and_build(raw)
                    pd = _pd()
                    data = raw.iloc[end+1:, :].copy()
                    data.columns = headers[: data.shape[1]]
                    if len(headers) > data.shape[1]:
                        for j in range(data.shape[1], len(headers)):
                            data[headers[j]] = pd.NA
                    data = data.dropna(how='all')
                    for c in data.columns:
                        data[c] = data[c].apply(lambda x: str(x).strip() if (x is not None and str(x).strip() != "nan") else x)
                    parsed.append(ParsedSheet(file_path=p, sheet_name=sh, header_row=start, columns=headers, data=data))
            except Exception as e:
                traceback.print_exc()
                QMessageBox.critical(self, "읽기 오류", f"{os.path.basename(p)} 처리 중 오류\n{e}")
            finally:
                self.progress.setValue(i)
        self.parsed = parsed

        if not parsed:
            QMessageBox.warning(self, "알림", "유효한 시트를 찾지 못했습니다.")
            self.progress.setVisible(False)
            return

        pd = _pd()
        all_cols: List[str] = []
        for ps in parsed:
            for c in ps.data.columns:
                if c not in all_cols:
                    all_cols.append(c)

        frames = []
        for ps in parsed:
            df = ps.data.copy()
            for c in all_cols:
                if c not in df.columns:
                    df[c] = pd.NA
            df = df[all_cols]
            df[META_MAP["file"]] = os.path.basename(ps.file_path)
            df[META_MAP["sheet"]] = str(ps.sheet_name)
            df[META_MAP["header_row"]] = ps.header_row
            frames.append(df)

        combined = _pd().concat(frames, ignore_index=True)
        combined = combined.dropna(how='all', subset=[c for c in all_cols])
        combined[META_MAP["violations"]] = compute_violations(combined)

        self.combined = combined
        self.progress.setVisible(False)
        self._refresh_preview()

    def _filtered_for_preview(self):
        if self.combined is None:
            return None
        df = self.combined.copy()
        if not self.chk_meta.isChecked():
            drop_cols = [c for c in META_COLS_KR if c in df.columns]
            if drop_cols:
                df = df.drop(columns=drop_cols)
        return df.head(500).copy()

    def _refresh_preview(self):
        if self.combined is None:
            return
        preview = self._filtered_for_preview()
        self.table.setModel(PandasModel(preview))

    def _save_output(self):
        if self.combined is None or len(self.combined) == 0:
            QMessageBox.warning(self, "알림", "먼저 파일을 드래그하여 미리보기를 확인하세요.")
            return
        default_name = "취합_결과.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "저장 위치 선택", default_name, "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            pd = _pd()
            df_out = self.combined.copy()
            if not self.chk_meta.isChecked():
                drop_cols = [c for c in META_COLS_KR if c in df_out.columns]
                if drop_cols:
                    df_out = df_out.drop(columns=drop_cols)

            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "합본"

            split_cols = [str(c).split(" | ") for c in df_out.columns]
            depth = max(len(parts) for parts in split_cols) if split_cols else 1
            depth = min(max(depth, 1), 5)

            norm_cols = []
            for parts in split_cols:
                parts = parts[-depth:]
                if len(parts) < depth:
                    parts = ([""] * (depth - len(parts))) + parts
                norm_cols.append(parts)

            for r in range(depth):
                for c, parts in enumerate(norm_cols, start=1):
                    val = parts[r] if parts[r] != "" else None
                    ws.cell(row=r+1, column=c, value=val)

            for r in range(depth):
                cur, start = None, 1
                for c in range(1, len(norm_cols)+1):
                    v = ws.cell(row=r+1, column=c).value or ""
                    if cur is None:
                        cur, start = v, c
                    elif v != cur:
                        if cur != "" and (c-1) > start:
                            ws.merge_cells(start_row=r+1, start_column=start, end_row=r+1, end_column=c-1)
                        cur, start = v, c
                if cur is not None and cur != "" and len(norm_cols) >= start+1:
                    ws.merge_cells(start_row=r+1, start_column=start, end_row=r+1, end_column=len(norm_cols))

            for c in range(1, len(norm_cols)+1):
                for r in range(1, depth):
                    top = ws.cell(row=r, column=c).value
                    bot = ws.cell(row=r+1, column=c).value
                    if (top is not None and str(top) != "") and (bot is None or str(bot) == ""):
                        ws.merge_cells(start_row=r, start_column=c, end_row=r+1, end_column=c)

            for i, row in enumerate(df_out.itertuples(index=False), start=depth+1):
                for j, val in enumerate(row, start=1):
                    try:
                        import pandas as pd
                        ws.cell(row=i, column=j, value=None if pd.isna(val) else val)
                    except Exception:
                        ws.cell(row=i, column=j, value=val)

            wb.save(path)

            log_path = os.path.join(os.path.dirname(path), "취합_로그.csv")
            _pd().DataFrame([
                {"file": os.path.basename(p.file_path), "sheet": p.sheet_name, "header_row": p.header_row}
                for p in self.parsed
            ]).to_csv(log_path, index=False, encoding="utf-8-sig")

            QMessageBox.information(self, "저장됨", f"출력: {path}\n로그: {log_path}")
        except Exception as err:
            traceback.print_exc()
            QMessageBox.critical(self, "저장 오류", str(err))
            


class SheetChooser(QDialog):
    """시트 목록을 체크박스로 보여주는 단순 다이얼로그.
       preselect에 포함된 시트는 미리 체크됨.
    """
    def __init__(self, parent, file_path: str, sheets: list[str], preselect: list[str] | None = None):
        super().__init__(parent)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowTitle(f"시트 선택 ― {os.path.basename(file_path)}")
        self.resize(420, 520)
        self._sheets = sheets
        self._pre = set(preselect or [])

        layout = QVBoxLayout(self)
        self.info = QLabel("취합할 시트를 선택하세요. (복수 선택 가능)")
        self.info.setStyleSheet("color:#333;")
        layout.addWidget(self.info)

        self.listw = QListWidget(self)
        for s in sheets:
            it = QListWidgetItem(s)
            flags = it.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
            it.setFlags(flags)
            if s in self._pre:
                it.setCheckState(Qt.CheckState.Checked)
            else:
                it.setCheckState(Qt.CheckState.Unchecked)
            self.listw.addItem(it)
        layout.addWidget(self.listw, 1)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            parent=self
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        QTimer.singleShot(0, self._ensure_attention)

    def selected_sheets(self) -> list[str]:
        out: list[str] = []
        for i in range(self.listw.count()):
            it = self.listw.item(i)
            if it.checkState() == Qt.CheckState.Checked:
                out.append(it.text())
        return out

    def _ensure_attention(self):
        try:
            self.raise_()
            self.activateWindow()
            QApplication.alert(self, 0)
        except Exception:
            pass

def main():
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    os.environ.setdefault("QT_SCALE_FACTOR", "1")
    app = QApplication(sys.argv)
    w = ExcelAggregator()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
