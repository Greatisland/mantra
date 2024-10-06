import sys
import subprocess
import json
import os
import logging
import base64
from io import BytesIO
import time
import urllib.request
from urllib.parse import urljoin, urlparse

# 로깅 설정 (패키지 설치 이전에 로깅 설정)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        # logging.FileHandler("app.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# 기본 매핑 파일 경로
DEFAULT_MAPPING_FILE = os.path.join('config', 'default_mappings.json')

def load_mapping(file_path):
    """
    JSON 파일에서 매핑을 로드합니다.
    파일이 없거나 JSON이 유효하지 않은 경우 기본 매핑을 반환합니다.
    """
    if not os.path.exists(file_path):
        logging.warning(f"매핑 파일 '{file_path}'이 존재하지 않습니다. 기본 매핑을 사용합니다.")
        return DEFAULT_MAPPING.copy()
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            mapping = json.load(f)
        logging.info(f"매핑이 '{file_path}'에서 성공적으로 로드되었습니다.")
        return mapping
    except json.JSONDecodeError:
        logging.error(f"매핑 파일 '{file_path}'이 유효한 JSON 형식이 아닙니다.")
        return DEFAULT_MAPPING.copy()
    except Exception as e:
        logging.error(f"매핑 로드 중 오류 발생: {str(e)}")
        return DEFAULT_MAPPING.copy()
    
# 기본 매핑 정의 (프로그램 초기 로드용)
DEFAULT_MAPPING = {
    "빈값": {"1", "2", "3"}
}

import importlib

# 필요한 패키지 목록 (튜플 형식: (패키지명, 임포트명, 버전))
required_packages = [
    ("pandas", "pandas", None),
    ("beautifulsoup4", "bs4", None),
    ("xlrd", "xlrd", "==1.2.0"),
    ("PyQt5", "PyQt5", None),
    ("openpyxl", "openpyxl", None),
    ("matplotlib", "matplotlib", None),
    ("Pillow", "PIL", None)
]

def check_and_install_packages():
    missing_packages = []
    for package, import_name, version in required_packages:
        try:
            importlib.import_module(import_name)
            logging.info(f"'{package}' 이미 설치되어 있습니다.")
        except ImportError:
            missing_packages.append((package, import_name, version))

    if missing_packages:
        logging.info("필수 패키지가 설치되어 있지 않습니다. 설치를 시작합니다...")
        for package, import_name, version in missing_packages:
            try:
                if version:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", f"{package}{version}"])
                else:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                logging.info(f"'{package}' 설치가 완료되었습니다.")
            except subprocess.CalledProcessError:
                logging.error(f"'{package}' 설치에 실패했습니다. 수동으로 설치해 주세요.")
                sys.exit(1)
        # 패키지 설치 후 스크립트 재시작
        logging.info("패키지 설치가 완료되었습니다. 스크립트를 재시작합니다.")
        os.execv(sys.executable, [sys.executable] + sys.argv)

# 패키지 설치 및 임포트 확인
check_and_install_packages()

# 이제 필요한 서드파티 패키지를 임포트
from PIL import Image, UnidentifiedImageError
from PyQt5.QtGui import QPixmap, QIcon, QKeySequence, QTextOption  
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QComboBox,
    QTableView, QDialog, QLineEdit, QGridLayout, QHeaderView,
    QAbstractItemView, QScrollArea, QDialogButtonBox, QProgressBar,
    QStatusBar, QTextEdit, QAction, QMenu, QCompleter, QCheckBox,
    QToolBar, QTabWidget, QInputDialog, QTableWidget, QTableWidgetItem,
    QShortcut, QUndoStack, QUndoCommand, QStyledItemDelegate, QTextEdit
)
from PyQt5.QtCore import (
    Qt, QAbstractTableModel, QVariant, pyqtSignal, QObject, QThread,
    QModelIndex, QPoint, QItemSelection, QItemSelectionModel, QTranslator,
    QLocale
)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import pandas as pd
from bs4 import BeautifulSoup
from collections import defaultdict

class EditCellCommand(QUndoCommand):
    def __init__(self, model, index, old_value, new_value, description="Edit Cell"):
        super().__init__(description)
        self.model = model
        self.index = index
        self.old_value = old_value
        self.new_value = new_value

    def undo(self):
        self.model.ignore_undo = True
        self.model.setData(self.index, self.old_value, Qt.EditRole)
        self.model.ignore_undo = False

    def redo(self):
        self.model.ignore_undo = True
        self.model.setData(self.index, self.new_value, Qt.EditRole)
        self.model.ignore_undo = False


class DeleteCellsCommand(QUndoCommand):
    def __init__(self, model, selected_indexes, description="Delete Cells"):
        super().__init__(description)
        self.model = model
        self.selected_indexes = selected_indexes
        # Store old values
        self.old_values = {index: self.model.data(index, Qt.DisplayRole) for index in selected_indexes}

    def undo(self):
        self.model.ignore_undo = True
        for index, value in self.old_values.items():
            self.model.setData(index, value, Qt.EditRole)
        self.model.ignore_undo = False

    def redo(self):
        self.model.ignore_undo = True
        for index in self.selected_indexes:
            self.model.setData(index, '', Qt.EditRole)  # 빈 문자열로 삭제 처리
        self.model.ignore_undo = False


class PasteCellsCommand(QUndoCommand):
    def __init__(self, model, paste_data, start_index, description="Paste Cells"):
        super().__init__(description)
        self.model = model
        self.paste_data = paste_data  # list of lists
        self.start_index = start_index
        # Store old values
        self.old_values = {}
        for r, row in enumerate(paste_data):
            for c, value in enumerate(row):
                index = self.model.index(start_index.row() + r, start_index.column() + c)
                if index.isValid():
                    self.old_values[index] = self.model.data(index, Qt.DisplayRole)

    def undo(self):
        self.model.ignore_undo = True
        for index, value in self.old_values.items():
            self.model.setData(index, value, Qt.EditRole)
        self.model.ignore_undo = False

    def redo(self):
        self.model.ignore_undo = True
        for r, row in enumerate(self.paste_data):
            for c, value in enumerate(row):
                index = self.model.index(self.start_index.row() + r, self.start_index.column() + c)
                if index.isValid():
                    self.model.setData(index, value, Qt.EditRole)
        self.model.ignore_undo = False

class PasteMultipleCellsCommand(QUndoCommand):
    def __init__(self, model, indexes, value, description="Paste Multiple Cells"):
        super().__init__(description)
        self.model = model
        self.indexes = indexes
        self.value = value
        # 이전 값을 저장하여 Undo 시 복원
        self.old_values = {index: self.model.data(index, Qt.DisplayRole) for index in indexes}

    def undo(self):
        self.model.ignore_undo = True
        for index, old_value in self.old_values.items():
            self.model.setData(index, old_value, Qt.EditRole)
        self.model.ignore_undo = False

    def redo(self):
        self.model.ignore_undo = True
        for index in self.indexes:
            self.model.setData(index, self.value, Qt.EditRole)
        self.model.ignore_undo = False

# 텍스트 길이제한 해제
class TextEditDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QTextEdit(parent)
        editor.setWordWrapMode(QTextOption.WrapAnywhere)
        return editor

    def setEditorData(self, editor, index):
        value = index.data(Qt.EditRole)
        if value:
            editor.setPlainText(str(value))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), Qt.EditRole)

# 데이터 프레임 모델
class DataFrameModel(QAbstractTableModel):
    dataChangedSignal = pyqtSignal()

    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df.copy()
        self._original_order = df.reset_index(drop=True).index.tolist()
        self.undo_stack = None  # Undo 스택을 나중에 설정
        self.ignore_undo = False  # Undo/Redo 중 명령 푸시를 무시하기 위한 플래그

    def setUndoStack(self, undo_stack):
        self.undo_stack = undo_stack

    def rowCount(self, parent=None):
        return len(self._df.index)

    def columnCount(self, parent=None):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return QVariant()
        if role in (Qt.DisplayRole, Qt.EditRole):
            value = self._df.iloc[index.row(), index.column()]
            if pd.isna(value):
                return ""
            return str(value)
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return QVariant()
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(section + 1)  # 1부터 시작하는 행 번호

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable
    

    def setData(self, index, value, role=Qt.EditRole):
        
        if not index.isValid():
            logging.debug("setData 호출 실패: Invalid index")
            return False
        if role not in (Qt.DisplayRole, Qt.EditRole):
            logging.debug(f"setData 호출 실패: Unsupported role {role}")
            return False

        if self.ignore_undo:
            # Undo/Redo 작업 중인 경우, value를 직접 설정하고 비교를 건너뜁니다.
            try:
                if isinstance(value, str):
                    if value == "" or value == "(필드 값 없음)":
                        self._df.iloc[index.row(), index.column()] = pd.NA
                        logging.debug(f"setData: Set pd.NA at ({index.row()}, {index.column()})")
                    else:
                        self._df.iloc[index.row(), index.column()] = value
                        logging.debug(f"setData: Set value '{value}' at ({index.row()}, {index.column()})")
                else:
                    self._df.iloc[index.row(), index.column()] = value
                    logging.debug(f"setData: Set value '{value}' at ({index.row()}, {index.column()})")
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                self.dataChangedSignal.emit()
                return True
            except Exception as e:
                logging.error(f"데이터 업데이트 실패: {e}")
                return False

        # Undo/Redo 작업이 아닌 경우, 기존 로직을 수행합니다.
        old_value = self._df.iloc[index.row(), index.column()]

        # value의 타입에 따라 new_value를 안전하게 할당
        if pd.isna(value):
            new_value = pd.NA
        elif isinstance(value, str):
            new_value = value if value != "" else pd.NA
        else:
            new_value = value

        # pandas.NA를 안전하게 처리하기 위한 비교
        try:
            if pd.isna(old_value) and pd.isna(new_value):
                values_differ = False
            elif pd.isna(old_value) or pd.isna(new_value):
                values_differ = True
            else:
                values_differ = old_value != new_value
        except Exception as e:
            logging.error(f"비교 중 오류 발생: {e}")
            values_differ = False

        logging.debug(f"setData: values_differ={values_differ}")

        if values_differ:
            if self.undo_stack is not None and not self.ignore_undo:
                try:
                    command = EditCellCommand(self, index, old_value, new_value)
                    self.undo_stack.push(command)
                    logging.debug(f"setData: Pushed EditCellCommand for ({index.row()}, {index.column()})")
                except Exception as e:
                    logging.error(f"Undo 명령 실패: {e}")
                    return False
            else:
                # Undo 스택이 설정되지 않았거나, Undo/Redo 중인 경우 직접 데이터 변경
                try:
                    if isinstance(new_value, str) and (new_value == "(필드 값 없음)" or new_value == ""):
                        self._df.iloc[index.row(), index.column()] = pd.NA
                    else:
                        self._df.iloc[index.row(), index.column()] = new_value
                    self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                    self.dataChangedSignal.emit()
                    return True
                except Exception as e:
                    logging.error(f"데이터 업데이트 실패: {e}")
                    return False
            return True
        return False


    def get_dataframe(self):
        return self._df.copy()

    def set_dataframe(self, df):
        self.beginResetModel()
        self._df = df.copy().reset_index(drop=True)
        self._original_order = self._df.index.tolist()
        self.endResetModel()

    def sort_original_order(self):
        self.beginResetModel()
        self._df = self._df.loc[self._original_order].reset_index(drop=True)
        self.endResetModel()

    def clear_selected_cells_bulk(self, selected_indexes):
        """
        선택된 셀들을 빈 문자열로 일괄적으로 비웁니다.
        """
        try:
            if not selected_indexes:
                return

            # 선택된 셀들을 행별로 그룹화
            from collections import defaultdict
            cells_by_row = defaultdict(list)
            for index in selected_indexes:
                cells_by_row[index.row()].append(index.column())

            # 데이터 수정 및 신호 최소화
            for row, cols in cells_by_row.items():
                sorted_cols = sorted(cols)
                # Find contiguous column ranges
                ranges = []
                start = end = sorted_cols[0]
                for col in sorted_cols[1:]:
                    if col == end + 1:
                        end = col
                    else:
                        ranges.append((start, end))
                        start = end = col
                ranges.append((start, end))
                
                # Update data and emit dataChanged for each range
                for start_col, end_col in ranges:
                    for col in range(start_col, end_col + 1):
                        # 올바른 인덱싱 사용
                        if row < self.rowCount() and col < self.columnCount():
                            self._df.iloc[row, col] = ''
                        else:
                            logging.warning(f"행 {row}, 열 {col}이 DataFrame 범위를 벗어났습니다.")
                    topLeft = self.index(row, start_col)
                    bottomRight = self.index(row, end_col)
                    self.dataChanged.emit(topLeft, bottomRight, [Qt.DisplayRole, Qt.EditRole])

            logging.info("선택된 셀의 내용이 빈 값으로 변환되었습니다.")
        except Exception as e:
            logging.error(f"셀 데이터 변환 실패 - 오류: {str(e)}")


# 로그 관련 클래스
class QTextEditLogger(logging.Handler, QObject):
    log_signal = pyqtSignal(str)

    def __init__(self):
        logging.Handler.__init__(self)
        QObject.__init__(self)

    def emit(self, record):
        msg = self.format(record)
        self.log_signal.emit(msg)

# 매핑 편집 다이얼로그 클래스
class MappingEditorDialog(QDialog):
    def __init__(self, mapping, source_columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("매칭 표 수정")
        self.mapping = mapping.copy()
        self.source_columns = source_columns  # 소스 파일의 컬럼 목록
        self.init_ui()
        self.set_default_size()

    def init_ui(self):
        layout = QVBoxLayout()

        # 스타일 시트 적용
        self.setStyleSheet("""
            QDialog {
                background-color: #ECEFF4; /* 파스텔 톤 배경 */
            }
            QWidget {
                background-color: #FFFFFF; /* 파스텔 톤 배경 */
            }
            QLabel {
                font-size: 12px;
                color: #4C566A; /* 다크 그레이 텍스트 색상 */
            }
                           
            QComboBox {
                background-color: #E5E9F0; /* 파스텔 블루 콤보박스 배경 */
                color: #3B4252;  /* 다크 그레이 텍스트 */
                border: 1px solid #D8DEE9;
                padding: 4px;
                border-radius: 4px;
            }

            QLineEdit {
                color: #3B4252;  /* 다크 그레이 텍스트 */
                border: 1px solid #D8DEE9;
                padding: 4px;
                border-radius: 4px;
            }
                           
            QPushButton {
                font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
                background-color: #88C0D0;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
        """)
        # 검색 바 추가
        search_layout = QHBoxLayout()
        search_label = QLabel("검색:")
        self.search_line_edit = QLineEdit()
        self.search_line_edit.setPlaceholderText("매칭 항목을 검색하세요...")
        self.search_line_edit.textChanged.connect(self.filter_items)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_line_edit)
        layout.addLayout(search_layout)

        # 스크롤 가능한 영역
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QGridLayout()

        self.combo_boxes = {}
        self.labels = {}
        self.all_a_cols = list(self.mapping.keys())

        row_num = 0
        for a_col, b_col in self.mapping.items():
            # 'Index' 열은 매핑에서 제외
            if a_col.lower() == 'index':
                continue

            label = QLabel(a_col)
            self.labels[a_col] = label
            combo_box = QComboBox()
            combo_box.setEditable(True)  # 자동완성을 위해 QComboBox를 Editable로 설정
            combo_box.addItem("")  # 빈 선택지를 추가
            combo_box.addItems(self.source_columns)
            # if b_col in self.source_columns:
            #     combo_box.setCurrentText(b_col)
            # else:
            #     combo_box.setCurrentIndex(0)
            combo_box.setCurrentText(b_col)

            # 자동완성 설정
            completer = QCompleter(self.source_columns)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            combo_box.setCompleter(completer)

            # 텍스트가 변경될 때마다 검증
            combo_box.lineEdit().textChanged.connect(lambda text, cb=combo_box: self.validate_combo_box(cb, text))

            # 초기 값에 대한 검증 수행
            self.validate_combo_box(combo_box, b_col)

            scroll_layout.addWidget(label, row_num, 0)
            scroll_layout.addWidget(combo_box, row_num, 1)
            self.combo_boxes[a_col] = combo_box
            row_num += 1

        scroll_content.setLayout(scroll_layout)
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        # 버튼 박스
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        reset_button = button_box.addButton("Reset", QDialogButtonBox.ResetRole)
        reset_button.clicked.connect(self.reset_mapping)
        button_box.accepted.connect(self.save_mapping)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        self.setLayout(layout)

    def set_default_size(self):
        """
        매칭 표 수정 창의 기본 크기를 설정합니다.
        필요에 따라 크기를 조정하세요.
        """
        self.setMinimumSize(200, 600)  # 최소 크기 설정
        self.resize(400, 800)           # 초기 크기 설정
        # 또는 고정 크기로 설정하려면 아래 주석을 해제하세요.
        # self.setFixedSize(800, 600)

    def filter_items(self, text):
        """
        검색어에 따라 매칭 항목을 필터링합니다.
        """
        text = text.lower()
        for a_col in self.all_a_cols:
            if text in a_col.lower():
                self.labels[a_col].show()
                self.combo_boxes[a_col].show()
            else:
                self.labels[a_col].hide()
                self.combo_boxes[a_col].hide()

    def validate_combo_box(self, combo_box, text):
        """콤보박스의 QLineEdit 텍스트 색상을 유효성에 따라 변경"""
        if text not in self.source_columns:
            combo_box.setStyleSheet("""
                QComboBox {
                    color: red;
                }
            """)
        else:
            combo_box.setStyleSheet("""
                QComboBox {
                    color: green;
                }
            """)

    def reset_mapping(self):
        """
        데이터 매칭을 기본값으로 초기화합니다.
        """
        reply = QMessageBox.question(
            self,
            '데이터 매칭 초기화',
            '매칭을 기본 설정으로 초기화하시겠습니까?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 현재 매핑의 모든 값(value)을 빈 문자열로 설정
            for a_col in self.mapping.keys():
                self.mapping[a_col] = ""
                combo_box = self.combo_boxes.get(a_col)
                if combo_box:
                    combo_box.setCurrentText("")
                    self.validate_combo_box(combo_box, "")
            logging.info("매칭이 기본 설정으로 초기화되었습니다.")

    def save_mapping(self):
        for a_col, combo_box in self.combo_boxes.items():
            self.mapping[a_col] = combo_box.currentText().strip()
        # 로그 추가: 매핑이 어떻게 저장되는지 확인
        self.accept()

    def get_mapping(self):
        return self.mapping

# 커스텀 QTableView 클래스 (좌측 상단 꼭지점 클릭 감지)
class CustomTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        # 이미 context menu 정책은 메인 클래스에서 설정됨

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            pos = event.pos()
            index = self.indexAt(pos)
            # Check if the click is on the top-left corner
            if not index.isValid():
                # Check if both headers return -1, indicating top-left corner
                if self.horizontalHeader().logicalIndexAt(pos) == -1 and self.verticalHeader().logicalIndexAt(pos) == -1:
                    # 전체 선택
                    self.selectAll()
            super().mousePressEvent(event)
        else:
            # Right-click or other buttons
            super().mousePressEvent(event)

# 이미지 클래스 정의
class ImageProcessor(QObject):
    progress = pyqtSignal(int)           # 진행률 (%)을 전달하는 신호
    log = pyqtSignal(str)                # 로그 메시지를 전달하는 신호
    finished = pyqtSignal()              # 작업 완료를 알리는 신호
    error = pyqtSignal(str)              # 오류 메시지를 전달하는 신호
    update_cell = pyqtSignal(QModelIndex, str)  # 셀 업데이트를 위한 시그널

    def __init__(self, selected_indexes, new_path, download_images, download_folder, base64_download_folder, model):
        super().__init__()
        self.selected_indexes = selected_indexes
        self.new_path = new_path
        self.download_images = download_images
        self.download_folder = download_folder
        self.base64_download_folder = base64_download_folder
        self.model = model

    def run(self):
        total = len(self.selected_indexes)
        if total == 0:
            self.progress.emit(100)  # 작업이 없을 경우 바로 100% 방출
            self.finished.emit()
            return

        for i, index in enumerate(self.selected_indexes, 1):
            try:
                current_data = self.model.data(index, Qt.DisplayRole)
                if not current_data:
                    continue

                soup = BeautifulSoup(current_data, 'html.parser')
                imgs = soup.find_all('img')
                if not imgs:
                    continue  # img 태그가 없는 셀은 건너뜀

                for img in imgs:
                    src = img.get('src', '')
                    if src.startswith('data:image'):
                        # Base64 데이터 처리
                        try:
                            header, encoded = src.split(',', 1)
                            file_ext = header.split('/')[1].split(';')[0]  # 예: 'png', 'jpeg'

                            # 파일명 생성 (임의의 규칙: img_row_col_timestamp.ext)
                            timestamp = int(time.time() * 1000)
                            file_name = f"img_{index.row()}_{index.column()}_{timestamp}.{file_ext}"
                            file_path = os.path.join(self.base64_download_folder, file_name)

                            # 파일명에 확장자가 없을 경우 '.jpg' 추가
                            if not os.path.splitext(file_name)[1]:
                                file_name += '.jpg'
                                file_path = os.path.join(self.base64_download_folder, file_name)

                            # 이미지 데이터 디코딩 및 저장
                            image_data = base64.b64decode(encoded)
                            with open(file_path, 'wb') as f:
                                f.write(image_data)
                            self.log.emit(f"셀({index.row()}, {index.column()})의 Base64 이미지를 '{file_path}'에 저장했습니다.")

                            # img src를 새로운 경로 + 파일명으로 변경
                            new_src = os.path.join(self.new_path, file_name).replace('\\', '/')
                            img['src'] = new_src
                            self.log.emit(f"셀({index.row()}, {index.column()})의 img src를 '{src}'에서 '{new_src}'로 변경했습니다.")

                        except Exception as e:
                            error_msg = f"셀({index.row()}, {index.column()})의 Base64 이미지 처리 실패: {e}"
                            self.log.emit(error_msg)
                            self.error.emit(error_msg)
                            continue
                    else:
                        # 일반 URL 이미지 처리
                        try:
                            # 프로토콜이 없는 URL 처리
                            parsed_url = urlparse(src)
                            if not parsed_url.scheme:
                                # 기본적으로 https를 사용
                                src = 'https:' + src
                                self.log.emit(f"프로토콜이 없는 URL을 https로 보완: {src}")

                            # 기존 src에서 파일명 추출
                            file_name = os.path.basename(urlparse(src).path)
                            if not file_name:
                                continue  # 파일명이 없으면 건너뜀

                            # 파일명에 확장자가 없을 경우 '.jpg' 추가
                            if not os.path.splitext(file_name)[1]:
                                file_name += '.jpg'

                            # 이미지 다운로드
                            if self.download_images:
                                # 이미지 다운로드 시도
                                image_path = os.path.join(self.download_folder, file_name)
                                urllib.request.urlretrieve(src, image_path)
                                self.log.emit(f"셀({index.row()}, {index.column()})의 이미지를 '{image_path}'에 다운로드했습니다.")

                            # img src를 새로운 경로 + 파일명으로 변경
                            new_src = os.path.join(self.new_path, file_name).replace('\\', '/')
                            img['src'] = new_src
                            self.log.emit(f"셀({index.row()}, {index.column()})의 img src를 '{src}'에서 '{new_src}'로 변경했습니다.")

                        except Exception as e:
                            error_msg = f"셀({index.row()}, {index.column()})의 이미지 처리 실패: {e}"
                            self.log.emit(error_msg)
                            self.error.emit(error_msg)
                            continue

                # 수정된 HTML을 문자열로 변환
                new_data = str(soup)

                # 모델 업데이트를 메인 스레드로 전달
                self.update_cell.emit(index, new_data)

                # 진행률 업데이트
                progress_percent = int((i / total) * 100)
                self.progress.emit(progress_percent)

            except Exception as e:
                error_msg = f"셀({index.row()}, {index.column()})의 이미지 처리 중 예외 발생: {e}"
                self.log.emit(error_msg)
                self.error.emit(error_msg)
                continue

        self.finished.emit()

class DataLoaderThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(pd.DataFrame)
    error = pyqtSignal(str)

    def __init__(self, file_path, file_ext, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.file_ext = file_ext

    def run(self):
        try:
            if self.file_ext == ".csv":
                df = self.load_csv(self.file_path)
            elif self.file_ext in [".xlsx", ".xls"]:
                if self.file_ext == ".xlsx":
                    engine = 'openpyxl'
                else:
                    # 먼저 파일이 HTML인지 확인
                    with open(self.file_path, 'r', encoding='utf-8', errors='ignore') as file:
                        content = file.read(1024).lower()  # 첫 1KB만 읽어 확인
                    if '<html' in content:
                        with open(self.file_path, 'r', encoding='utf-8', errors='ignore') as file:
                            html_content = file.read()
                        df = self.parse_custom_html_xls(html_content)
                        if df is not None:
                            df = self.clean_column_names(df)
                    else:
                        engine = 'xlrd'
                        df = self.load_excel(self.file_path, engine=engine)
            elif self.file_ext == ".html":
                with open(self.file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                df = self.parse_html(content)
            else:
                self.error.emit("지원되지 않는 파일 형식입니다.")
                return

            if df is not None:
                self.finished.emit(df)
            else:
                self.error.emit("DataFrame 로드 실패.")
        except Exception as e:
            self.error.emit(str(e))

    def load_csv(self, file_path):
        """
        다양한 인코딩을 시도하여 CSV 파일을 로드합니다.
        """
        encodings = ['utf-8-sig', 'utf-8', 'cp949', 'ISO-8859-1']
        for enc in encodings:
            try:
                df = pd.read_csv(file_path, encoding=enc, index_col=None)
                logging.info(f"CSV 파일이 '{enc}' 인코딩으로 성공적으로 로드되었습니다.")

                # 불필요한 'Unnamed: 0' 컬럼 제거
                if 'Unnamed: 0' in df.columns:
                    df = df.drop(columns=['Unnamed: 0'])

                # 인덱스 리셋
                df = df.reset_index(drop=True)

                return df
            except UnicodeDecodeError:
                logging.warning(f"'{enc}' 인코딩으로 읽을 수 없습니다. 다음 인코딩을 시도합니다.")
                continue
            except FileNotFoundError:
                self.error.emit("파일을 찾을 수 없습니다.")
                logging.error("파일을 찾을 수 없습니다.")
                return None
            except pd.errors.ParserError as e:
                self.error.emit(f"CSV 파싱 오류: {str(e)}")
                logging.error(f"CSV 파싱 오류: {str(e)}")
                return None
            except Exception as e:
                self.error.emit(f"알 수 없는 오류가 발생했습니다: {str(e)}")
                logging.error(f"알 수 없는 오류: {str(e)}")
                return None

        self.error.emit("파일을 열 수 없습니다. 다른 인코딩을 시도해보세요.")
        logging.error("모든 인코딩 시도가 실패했습니다.")
        return None

    def load_excel(self, file_path, engine):
        """
        Excel 파일을 로드합니다.
        """
        try:
            df = pd.read_excel(file_path, engine=engine, index_col=None)
            logging.info(f"Excel 파일이 '{engine}' 엔진으로 성공적으로 로드되었습니다.")

            # 불필요한 'Unnamed: 0' 컬럼 제거
            if 'Unnamed: 0' in df.columns:
                df = df.drop(columns=['Unnamed: 0'])

            # 인덱스 리셋
            df = df.reset_index(drop=True)

            return df
        except Exception as e:
            self.error.emit(f"Excel 파일을 열 수 없습니다. 오류: {str(e)}")
            logging.error(f"Excel 파일 열기 실패: {str(e)}")
            return None

    def parse_custom_html_xls(self, html_content):
        """
        사용자 정의 HTML 형식의 .xls 파일을 파싱하여 DataFrame으로 변환합니다.
        """
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            table = soup.find('table')
            if not table:
                self.error.emit("HTML에서 테이블을 찾을 수 없습니다.")
                logging.error("HTML에서 테이블을 찾을 수 없습니다.")
                return None

            rows = table.find_all('tr')
            if not rows:
                self.error.emit("HTML 테이블에 tr 요소가 없습니다.")
                logging.error("HTML 테이블에 tr 요소가 없습니다.")
                return None

            # 첫 번째 tr은 헤더
            header_tr = rows[0]
            header_tds = header_tr.find_all('td')
            if not header_tds:
                self.error.emit("헤더 tr에 td 요소가 없습니다.")
                logging.error("헤더 tr에 td 요소가 없습니다.")
                return None

            # 첫 번째 td에 class="title"이 있는지 확인
            first_td = header_tds[0]
            if 'title' not in first_td.get('class', []):
                self.error.emit("헤더의 첫 번째 td에 class='title'이 없습니다.")
                logging.error("헤더의 첫 번째 td에 class='title'이 없습니다.")
                return None

            # 열 이름 추출
            columns = [td.get_text(strip=True) for td in header_tds]

            data = []
            for tr in rows[1:]:
                tds = tr.find_all('td')
                if not tds:
                    continue  # 빈 tr 건너뜀
                row = [td.get_text(strip=True) for td in tds]
                # 열 수에 맞게 행을 패딩
                if len(row) < len(columns):
                    row += [''] * (len(columns) - len(row))
                elif len(row) > len(columns):
                    row = row[:len(columns)]
                data.append(row)

            df = pd.DataFrame(data, columns=columns)

            return df
        except Exception as e:
            self.error.emit(f"HTML을 파싱하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"HTML 파싱 실패: {str(e)}")
            return None

    def parse_html(self, html_content):
        """
        HTML에서 첫 번째 테이블을 찾아 DataFrame으로 변환하고 '내용' 열을 HTML로 처리합니다.
        """
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            df_list = pd.read_html(str(soup), header=0)  # header=0을 명시적으로 지정
            if not df_list:
                self.error.emit("HTML에서 테이블을 찾을 수 없습니다.")
                logging.error("HTML에서 테이블을 찾을 수 없습니다.")
                return None

            df = df_list[0]
            df = self.clean_column_names(df)

            # 불필요한 'Unnamed: 0' 컬럼 제거
            if 'Unnamed: 0' in df.columns:
                df = df.drop(columns=['Unnamed: 0'])

            # '내용' 열 처리
            if '내용' in df.columns:
                content_column_index = df.columns.get_loc('내용')
                rows = soup.find_all('tr')

                content_data = []
                for row in rows[1:]:  # 헤더 제외
                    cells = row.find_all('td')
                    if len(cells) > content_column_index:
                        cell_text = cells[content_column_index].get_text(strip=True)
                        if cell_text == '':
                            content_data.append("")
                        else:
                            content_data.append(cell_text)
                    else:
                        content_data.append("")
                df['내용'] = content_data
                logging.info("'내용' 열이 성공적으로 처리되었습니다.")
            else:
                self.error.emit("'내용' 열이 존재하지 않습니다.")
                logging.warning("'내용' 열이 존재하지 않습니다.")

            return df
        except Exception as e:
            self.error.emit(f"HTML을 파싱하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"HTML 파싱 실패: {str(e)}")
            return None

    def clean_column_names(self, df):
        """
        DataFrame의 열 이름을 정리합니다.
        """
        try:
            df.columns = df.columns.astype(str).str.strip()
            df.columns = df.columns.str.replace('\n', ' ', regex=True)
            df.columns = df.columns.str.replace('\r', ' ', regex=True)
            df.columns = df.columns.str.replace(' +', ' ', regex=True)
            logging.info("열 이름이 성공적으로 정리되었습니다.")
        except Exception as e:
            self.error.emit(f"열 이름을 정리하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"열 이름 정리 실패: {str(e)}")

            
# 메인 애플리케이션 클래스
class CSVMatcherApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mantra")
        self.setGeometry(100, 100, 1400, 900)
        # self.showMaximized()  # 창을 최대화된 상태로 표시

        # 프로그램 아이콘 설정 (아이콘 파일 'mantra.ico'이 동일 디렉토리에 있다고 가정)
        if os.path.exists('mantra.ico'):
            self.setWindowIcon(QIcon('mantra.ico'))
        else:
            logging.warning("아이콘 파일 'mantra.ico'을 찾을 수 없습니다.")

        self.df_source = None
        self.initial_df = None           # 초기 데이터프레임
        self.mapping = {}
        self.option = "회원"  # 기본값

        # 매핑 정의
        self.mappings = load_mapping(DEFAULT_MAPPING_FILE).copy()
        
        # Undo 스택 초기화
        self.undo_stack = QUndoStack(self)

        # Undo 스택 크기 제한 설정 (예: 100)
        self.undo_stack.setUndoLimit(50)

        # 모델 초기화 및 Undo 스택 설정
        self.model = DataFrameModel()
        self.model.setUndoStack(self.undo_stack)

        self.init_ui()
        self.setup_logging()

        self.setup_shortcuts()
        # 테스트 로그 메시지
        logging.info("애플리케이션이 시작되었습니다.")

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()

        # 상단 컨트롤 패널 레이아웃
        control_layout = QHBoxLayout()



        # 파일 열기 버튼
        open_button = QPushButton("파일 열기")
        open_button.setToolTip("CSV, xls, xlsx, html 파일을 열어 데이터를 불러옵니다.")
        open_button.setStatusTip("파일을 열어 데이터를 로드합니다.")
        open_button.setStyleSheet("""
            QPushButton {
                background-color: #81A1C1;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #5E81AC;
            }
        """)
        open_button.clicked.connect(self.open_file)
        control_layout.addWidget(open_button)
        
        # 데이터 유형 선택 콤보박스
        self.option_combo = QComboBox()
        self.option_combo.addItems(self.mappings.keys())
        self.option_combo.currentTextChanged.connect(self.update_mapping)
        control_layout.addWidget(QLabel("매칭 데이터 선택:"))
        control_layout.addWidget(self.option_combo)


        edit_mapping_button = QPushButton("매칭 표 수정")
        edit_mapping_button.clicked.connect(self.edit_mapping)
        edit_mapping_button.setToolTip("카페24 양식으로 매칭하기 위한 표를 작성합니다.")
        edit_mapping_button.setStatusTip("매핑 표를 편집합니다.")
        control_layout.addWidget(edit_mapping_button)

        save_mapping_button = QPushButton("매칭 상태 저장")
        save_mapping_button.clicked.connect(self.save_mapping_to_file)
        save_mapping_button.setToolTip("매칭한 값을 나중에 사용할 수 있도록 저장합니다.")
        save_mapping_button.setStatusTip("현재 매핑 상태를 파일로 저장합니다.")
        control_layout.addWidget(save_mapping_button)

        load_mapping_button = QPushButton("매칭 상태 불러오기")
        load_mapping_button.clicked.connect(self.load_mapping_from_file)
        load_mapping_button.setToolTip("저장한 매칭 상태를 불러옵니다.")
        load_mapping_button.setStatusTip("저장된 매핑 상태를 불러옵니다.")
        control_layout.addWidget(load_mapping_button)

        create_file_button = QPushButton("매칭된 파일 저장")
        create_file_button.setToolTip("매칭된 상태의 파일을 저장합니다.")
        create_file_button.setStatusTip("매핑된 데이터를 기반으로 파일을 저장합니다.")
        create_file_button.setStyleSheet("""
            QPushButton {
                background-color: #A3BE8C;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #B5D19A;
            }            
        """)
        create_file_button.clicked.connect(self.create_a_file)
        control_layout.addWidget(create_file_button)

        save_without_mapping_button = QPushButton("매칭 없이 파일 저장")
        save_without_mapping_button.setToolTip("매칭을 적용하지 않고 현재 수정된 데이터를 저장합니다.")
        save_without_mapping_button.setStatusTip("매핑을 무시하고 데이터를 저장합니다.")
        save_without_mapping_button.setStyleSheet("""
            QPushButton {
                background-color: #A3BE8C;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #B5D19A;
            }            
        """)
        save_without_mapping_button.clicked.connect(self.save_without_mapping)
        control_layout.addWidget(save_without_mapping_button)

        # 프로그램 소개 버튼
        info_button = QPushButton("Mantra")
        info_button.setToolTip("프로그램을 소개합니다.")
        info_button.setStatusTip("프로그램을 소개합니다.")
        info_button.setStyleSheet("""
            QPushButton {
                background-color: #81A1C1;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #5E81AC;
            }
        """)
        info_button.clicked.connect(self.show_info_popup)
        control_layout.addWidget(info_button)


        # 상단 컨트롤 패널을 메인 레이아웃에 추가
        main_layout.addLayout(control_layout)

        # 탭 위젯 추가
        self.tabs = QTabWidget()

        # 데이터 탭
        self.data_tab = QWidget()
        data_layout = QVBoxLayout()

        self.file_label = QLabel("파일이 선택되지 않았습니다.")
        data_layout.addWidget(self.file_label)

        self.table_view = CustomTableView()
        self.table_view.setItemDelegate(TextEditDelegate(self.table_view))
        # self.model = DataFrameModel()
        self.table_view.setModel(self.model)
        self.table_view.setSortingEnabled(False)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        self.table_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self.open_context_menu)

        # 헤더 클릭 시 전체 행/열 선택
        self.table_view.horizontalHeader().sectionClicked.connect(self.select_column)
        self.table_view.verticalHeader().sectionClicked.connect(self.select_row)

        data_layout.addWidget(self.table_view)

        # 메인 프로그레스 바
        self.main_progress_bar = QProgressBar()
        self.main_progress_bar.setValue(0)
        self.main_progress_bar.setAlignment(Qt.AlignCenter)
        self.main_progress_bar.setMaximum(100)
        self.main_progress_bar.setVisible(False)
        data_layout.addWidget(self.main_progress_bar)

        self.main_progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 20px;
            }
        """)

        # 행 인덱스 정보 레이블 추가
        self.row_info_label = QLabel("총 행 수: 0")
        data_layout.addWidget(self.row_info_label)

        self.data_tab.setLayout(data_layout)
        self.tabs.addTab(self.data_tab, "데이터")

        # 로그 탭
        self.log_tab = QWidget()
        log_layout = QVBoxLayout()

        # 로그 텍스트 에디트 추가
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.log_text_edit.setVisible(True)      
        log_layout.addWidget(self.log_text_edit) 

        self.log_tab.setLayout(log_layout)
        self.tabs.addTab(self.log_tab, "로그")

        # 메인 탭 위젯을 메인 레이아웃에 추가
        main_layout.addWidget(self.tabs)

        # 중앙 위젯 설정
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # 상태 바에 프로그레스 바 추가
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.status_progress_bar = QProgressBar()
        self.status_progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.status_progress_bar)

        # 인터페이스 스타일시트 적용
        self.apply_styles()

        # 행 헤더 표시
        self.table_view.verticalHeader().setVisible(True)
        # 초기 매핑 설정
        if self.option_combo.count() > 0:
            self.option_combo.setCurrentIndex(0)
            self.update_mapping(self.option_combo.currentText())
        else:
            logging.error("매핑 유형이 정의되지 않았습니다.")

    def setup_logging(self):
        # 이미 QTextEditLogger가 추가되어 있는지 확인
        if any(isinstance(handler, QTextEditLogger) for handler in logging.getLogger().handlers):
            return  # 이미 설정되어 있으므로 종료
        
        # QTextEdit 핸들러 설정 및 인스턴스 변수로 저장
        self.text_edit_handler = QTextEditLogger()
        self.text_edit_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.text_edit_handler.setFormatter(formatter)
        self.text_edit_handler.log_signal.connect(self.append_log)
        logging.getLogger().addHandler(self.text_edit_handler)

    def append_log(self, message):
        """로그 메시지를 QTextEdit에 추가합니다."""
        self.log_text_edit.append(message)

    def show_info_popup(self):
        """프로그램 소개 팝업창을 띄움"""
        dialog = QDialog(self)
        dialog.setWindowTitle("프로그램 소개")
        # 팝업창 배경 및 기본 스타일 적용
        dialog.setStyleSheet("""
            QDialog {
                background-color: #E5E9F0;  /* 파스텔 그레이 배경 */
                border-radius: 8px;  /* 부드러운 테두리 */
            }
        """)
        # 레이아웃 설정
        layout = QVBoxLayout(dialog)

        description = (
            "<strong style='font-size: 13px;'>만트라(Mantra)는 무엇입니까?</strong><br><br>"
            "우리가 스트레스를 받거나 강박적인 부정적인 생각으로 마음이 혼란스러울 때 "
            "<span style='color: #D08770;'>만트라</span>를 주장하면 마음을 안정시킬 수 있습니다. "
            "마음속의 잡음을 가라앉히고 부작용 없이 긍정적인 정신 및 감정 상태를 생성합니다.<br><br>"
            "불교의 수행에서는 자비심, 명료성, 깊은 이해를 갖기 위해서 <strong>특정 만트라를 외웁니다</strong>. "
            "만트라는 정신적인 수행자들 뿐만 아니라 현대 사회에서 문제를 직면한 모든 사람에게 "
            "혜택을 가져다 주는 강력한 도구입니다."
        )
        text_label = QLabel(description)

        # 텍스트 스타일 시트 설정
        text_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
                color: #4C566A;  /* 다크 그레이 파스텔 톤 텍스트 색상 */
                padding: 10px;
                background-color: #FFFFFF;  /* 연한 파스텔 블루 배경 */
                border-radius: 6px;  /* 부드러운 테두리 */
                border: 1px solid #D8DEE9;  /* 테두리 색상 */
            }
        """)
        text_label.setWordWrap(True)  # 자동 줄바꿈 활성화
        text_label.setAlignment(Qt.AlignCenter)

        # 레이아웃에 텍스트 추가
        layout.addWidget(text_label)

        # 창 크기 조정
        dialog.setFixedSize(800, 250)
        dialog.exec_()

    def apply_styles(self):
        # 기본 스타일시트
        self.setStyleSheet("""
            QLabel, QPushButton, QComboBox, QTextEdit, QLineEdit {
                font-family: "Noto Sans KR", "Helvetica Neue", Arial, sans-serif;
            }
            QMainWindow {
                background-color: #f0f0f0;
            }
            QPushButton {
                font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
                background-color: #88C0D0;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
            QComboBox {
                font-size: 12px;
                padding: 4px;
                text-align: center;
                border-radius: 4px;
                background-color: white;
                color: #3B4252;
                border: 1px solid #F2F2F2;
            }
                           
            QComboBox QAbstractItemView {
                color: #3B4252;  /* 리스트 아이템 텍스트 색상 */
                border: 1px solid #D8DEE9;
                selection-background-color: #88C0D0; /* 선택된 아이템의 배경 */
                selection-color: white;
            }

            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;  /* 화살표 버튼의 너비 */
                
                background-color: #88C0D0;  /* 배경색 */
                border-top-right-radius: 4px;
                border-bottom-right-radius: 4px;
            }

            QComboBox::down-arrow {
                border-style: solid;
                border-width: 6px 0 0 6px;
                border-color: #FFFFFF #FFFFFF #FFFFFF #FFFFFF; 
                width: 0px;
                height: 0px;
            }
                           
            QComboBox::down-arrow:hover {
                border-width: 6px 0 0 6px;
                border-color: #3B4252 transparent transparent transparent;  
            }
                           
            QLabel {
                font-size: 12px;
            }
            QTableView {
                background-color: white;
                font-size: 12px;
            }
            QLineEdit {
                font-size: 12px;
            }
            QTextEdit {
                font-size: 12px;
            }
            QCheckBox {
                font-size: 12px;
            }
            QToolBar {
                background-color: #0e0e0e;
            }
            QStatusBar {
                font-size: 12px;
            }

            QScrollBar:horizontal {
                border: none;
                background: #FFFFFF;  /* 파스텔 톤 배경 */
                height: 12px;  /* 스크롤바의 높이 */
                margin: 0px 20px 0 20px;
            }
            QScrollBar::handle:horizontal {
                background: #999;  /* 핸들의 배경색  */
                min-width: 20px;
                border-radius: 5px;  /* 모서리를 둥글게 처리 */
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                width: 0px;
            }

            QScrollBar:vertical {
                border: none;
                background: #ECEFF4;  /* 파스텔 톤 배경 */
                width: 12px;  /* 스크롤바의 너비 */
                margin: 20px 0px 20px 0px;
            }
            QScrollBar::handle:vertical {
                background: #999999;  /* 핸들의 배경색 (파스텔 그린) */
                min-height: 20px;
                border-radius: 5px;  /* 모서리를 둥글게 처리 */
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }

            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
        """)

    def closeEvent(self, event):
        """
        애플리케이션 종료 시 확인 메시지를 표시하고 로깅 핸들러를 안전하게 정리합니다.
        """
        reply = QMessageBox.question(
            self,
            '종료 확인',
            "정말로 종료하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                # QTextEditLogger 핸들러를 로거에서 제거
                logging.getLogger().removeHandler(self.text_edit_handler)
                # 핸들러를 닫아 리소스 해제
                self.text_edit_handler.close()
                # 로깅 시스템 종료
                logging.shutdown()
            except Exception as e:
                print(f"로깅 종료 중 오류 발생: {e}")
            event.accept()
        else:
            event.ignore()

    def update_mapping(self, selected_option):
        """
        선택한 옵션에 따라 매핑을 업데이트합니다.
        """
        if selected_option in self.mappings:
            self.mapping = self.mappings[selected_option].copy()
            logging.info(f"'{selected_option}' 매핑이 업데이트되었습니다.")
        else:
            logging.error(f"선택된 옵션 '{selected_option}'에 해당하는 매핑이 존재하지 않습니다.")
            self.mapping = {}

    def open_file(self):
        """
        파일 열기 다이얼로그를 열고 파일을 로드합니다.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "파일 열기",
            "",
            "데이터 파일 (*.csv *.xlsx *.xls *.html)"
        )
        if not file_path:
            logging.warning("파일 선택이 취소되었습니다.")
            return

        self.file_label.setText(f"선택된 파일: {file_path}")
        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            if file_ext == ".csv":
                self.df_source = self.load_csv(file_path)
            elif file_ext in [".xlsx", ".xls"]:
                if file_ext == ".xlsx":
                    engine = 'openpyxl'
                    self.df_source = self.load_excel(file_path, engine=engine)
                else:  # .xls
                    # 먼저 파일이 HTML인지 확인
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                        content = file.read(1024).lower()  # 첫 1KB만 읽어 확인
                    if '<html' in content:
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                            html_content = file.read()
                        self.df_source = self.parse_custom_html_xls(html_content)
                        if self.df_source is not None:
                            self.clean_column_names(self.df_source)
                    else:
                        engine = 'xlrd'
                        self.df_source = self.load_excel(file_path, engine=engine)
            elif file_ext == ".html":
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                self.df_source = self.parse_html(content)
            else:
                QMessageBox.critical(self, "파일 열기 실패", "지원되지 않는 파일 형식입니다.")
                logging.error(f"지원되지 않는 파일 형식: {file_ext}")
                self.main_progress_bar.setVisible(False)
                self.status_bar.clearMessage()
                return

            if self.df_source is not None:
                # 초기 데이터 저장
                self.initial_df = self.df_source.copy()

                QMessageBox.information(self, "파일 열기 성공", "파일을 성공적으로 불러왔습니다.")
                logging.info(f"파일 '{file_path}'을 성공적으로 불러왔습니다.")
                self.display_data()
            else:
                logging.error("DataFrame 로드 실패.")
        except Exception as e:
            QMessageBox.critical(self, "파일 열기 실패", f"파일을 열 수 없습니다. 오류: {str(e)}")
            logging.error(f"파일 열기 실패: {str(e)}")
        finally:
            self.main_progress_bar.setVisible(False)
            self.status_bar.clearMessage()

    def setup_shortcuts(self):
        """
        단축키를 설정합니다.
        """
        # Ctrl+C 복사 단축키 설정
        copy_action = QAction(self)
        copy_action.setShortcut(QKeySequence.Copy)  # Ctrl+C 단축키 설정
        copy_action.triggered.connect(self.perform_copy)
        self.addAction(copy_action)  # 메인 윈도우에 액션 추가

        # Ctrl+V 붙여넣기 단축키 설정
        paste_action = QAction(self)
        paste_action.setShortcut(QKeySequence.Paste)  # Ctrl+V 단축키 설정
        paste_action.triggered.connect(self.perform_paste)
        self.addAction(paste_action)  # 메인 윈도우에 액션 추가

        # Delete 키 단축키 설정
        delete_action = QAction(self)
        delete_action.setShortcut(QKeySequence(Qt.Key.Key_Delete))  # Delete 키 설정
        delete_action.triggered.connect(self.delete_selected_cells)
        self.addAction(delete_action)  # 메인 윈도우에 액션 추가

        # Undo 단축키 설정 (Ctrl+Z)
        undo_action = QAction(self)
        undo_action.setShortcut(QKeySequence.Undo)  # Ctrl+Z
        undo_action.triggered.connect(self.undo_stack.undo)
        self.addAction(undo_action)

        # Redo 단축키 설정 (Ctrl+Y)
        redo_action = QAction(self)
        redo_action.setShortcut(QKeySequence.Redo)  # Ctrl+Y
        redo_action.triggered.connect(self.undo_stack.redo)
        self.addAction(redo_action)

    def perform_copy(self):
        """
        선택된 셀의 인덱스를 가져와 copy_selected_cells 메서드에 전달합니다.
        """
        selected_indexes = self.table_view.selectedIndexes()
        self.copy_selected_cells(selected_indexes)

    def perform_paste(self):
        """
        선택된 셀의 인덱스를 가져와 paste_clipboard_data 메서드에 전달합니다.
        """
        selected_indexes = self.table_view.selectedIndexes()
        self.paste_clipboard_data(selected_indexes)

    def delete_selected_cells(self):
        """
        Delete 키를 눌렀을 때 선택된 셀을 빈 값으로 만드는 메서드.
        """
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            self.clear_selected_cells(selected_indexes)
        else:
            QMessageBox.warning(self, "선택 없음", "비울 셀을 먼저 선택해주세요.")

    def parse_custom_html_xls(self, html_content):
        """
        사용자 정의 HTML 형식의 .xls 파일을 파싱하여 DataFrame으로 변환합니다.
        """
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            table = soup.find('table')
            if not table:
                QMessageBox.critical(self, "파싱 실패", "HTML에서 테이블을 찾을 수 없습니다.")
                logging.error("HTML에서 테이블을 찾을 수 없습니다.")
                return None

            rows = table.find_all('tr')
            if not rows:
                QMessageBox.critical(self, "파싱 실패", "HTML 테이블에 tr 요소가 없습니다.")
                logging.error("HTML 테이블에 tr 요소가 없습니다.")
                return None

            # 첫 번째 tr은 헤더
            header_tr = rows[0]
            header_tds = header_tr.find_all('td')
            if not header_tds:
                QMessageBox.critical(self, "파싱 실패", "헤더 tr에 td 요소가 없습니다.")
                logging.error("헤더 tr에 td 요소가 없습니다.")
                return None

            # 첫 번째 td에 class="title"이 있는지 확인
            first_td = header_tds[0]
            if 'title' not in first_td.get('class', []):
                QMessageBox.critical(self, "파싱 실패", "헤더의 첫 번째 td에 class='title'이 없습니다.")
                logging.error("헤더의 첫 번째 td에 class='title'이 없습니다.")
                return None

            # 열 이름 추출
            columns = [td.get_text(strip=True) for td in header_tds]

            data = []
            for tr in rows[1:]:
                tds = tr.find_all('td')
                if not tds:
                    continue  # 빈 tr 건너뜀
                row = [td.get_text(strip=True) for td in tds]
                # 열 수에 맞게 행을 패딩
                if len(row) < len(columns):
                    row += [''] * (len(columns) - len(row))
                elif len(row) > len(columns):
                    row = row[:len(columns)]
                data.append(row)

            df = pd.DataFrame(data, columns=columns)

            return df
        except Exception as e:
            QMessageBox.critical(self, "HTML 파싱 실패", f"HTML을 파싱하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"HTML 파싱 실패: {str(e)}")
            return None

    def parse_html(self, html_content):
        """
        HTML에서 첫 번째 테이블을 찾아 DataFrame으로 변환하고 '내용' 열을 HTML로 처리합니다.
        """
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            df_list = pd.read_html(str(soup), header=0)  # header=0을 명시적으로 지정
            if not df_list:
                QMessageBox.critical(self, "파싱 실패", "HTML에서 테이블을 찾을 수 없습니다.")
                logging.error("HTML에서 테이블을 찾을 수 없습니다.")
                return None

            df = df_list[0]
            self.clean_column_names(df)

            # 불필요한 'Unnamed: 0' 컬럼 제거
            if 'Unnamed: 0' in df.columns:
                df = df.drop(columns=['Unnamed: 0'])

            # '내용' 열 처리
            if '내용' in df.columns:
                content_column_index = df.columns.get_loc('내용')
                rows = soup.find_all('tr')

                content_data = []
                for row in rows[1:]:  # 헤더 제외
                    cells = row.find_all('td')
                    if len(cells) > content_column_index:
                        cell_text = cells[content_column_index].get_text(strip=True)
                        if cell_text == '':
                            content_data.append("")
                        else:
                            content_data.append(cell_text)
                    else:
                        content_data.append("")
                df['내용'] = content_data
                logging.info("'내용' 열이 성공적으로 처리되었습니다.")
            else:
                QMessageBox.warning(self, "경고", "'내용' 열이 존재하지 않습니다.")
                logging.warning("'내용' 열이 존재하지 않습니다.")

            # DataFrame 정보 로깅

            return df
        except Exception as e:
            QMessageBox.critical(self, "HTML 파싱 실패", f"HTML을 파싱하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"HTML 파싱 실패: {str(e)}")
            return None

    def clean_column_names(self, df):
        """
        DataFrame의 열 이름을 정리합니다.
        """
        try:
            df.columns = df.columns.astype(str).str.strip()
            df.columns = df.columns.str.replace('\n', ' ', regex=True)
            df.columns = df.columns.str.replace('\r', ' ', regex=True)
            df.columns = df.columns.str.replace(' +', ' ', regex=True)
            logging.info("열 이름이 성공적으로 정리되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "열 이름 정리 실패", f"열 이름을 정리하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"열 이름 정리 실패: {str(e)}")

    def load_csv(self, file_path):
        """
        다양한 인코딩을 시도하여 CSV 파일을 로드합니다.
        """
        encodings = ['utf-8-sig', 'utf-8', 'cp949', 'ISO-8859-1']
        for enc in encodings:
            try:
                df = pd.read_csv(file_path, encoding=enc, index_col=None)
                logging.info(f"CSV 파일이 '{enc}' 인코딩으로 성공적으로 로드되었습니다.")

                # 불필요한 'Unnamed: 0' 컬럼 제거
                if 'Unnamed: 0' in df.columns:
                    df = df.drop(columns=['Unnamed: 0'])

                # 인덱스 리셋
                df = df.reset_index(drop=True)

                return df
            except UnicodeDecodeError:
                logging.warning(f"'{enc}' 인코딩으로 읽을 수 없습니다. 다음 인코딩을 시도합니다.")
                continue
            except FileNotFoundError:
                QMessageBox.critical(self, "파일 열기 실패", "파일을 찾을 수 없습니다.")
                logging.error("파일을 찾을 수 없습니다.")
                return None
            except pd.errors.ParserError as e:
                QMessageBox.critical(self, "CSV 파싱 오류", f"CSV 파싱 오류: {str(e)}")
                logging.error(f"CSV 파싱 오류: {str(e)}")
                return None
            except Exception as e:
                QMessageBox.critical(self, "파일 열기 실패", f"알 수 없는 오류가 발생했습니다: {str(e)}")
                logging.error(f"알 수 없는 오류: {str(e)}")
                return None

        QMessageBox.critical(self, "파일 열기 실패", "파일을 열 수 없습니다. 다른 인코딩을 시도해보세요.")
        logging.error("모든 인코딩 시도가 실패했습니다.")
        return None

    def load_excel(self, file_path, engine):
        """
        Excel 파일을 로드합니다.
        """
        try:
            df = pd.read_excel(file_path, engine=engine, index_col=None)
            logging.info(f"Excel 파일이 '{engine}' 엔진으로 성공적으로 로드되었습니다.")

            # 불필요한 'Unnamed: 0' 컬럼 제거
            if 'Unnamed: 0' in df.columns:
                df = df.drop(columns=['Unnamed: 0'])

            # 인덱스 리셋
            df = df.reset_index(drop=True)

            return df
        except Exception as e:
            QMessageBox.critical(self, "Excel 파일 열기 실패", f"Excel 파일을 열 수 없습니다. 오류: {str(e)}")
            logging.error(f"Excel 파일 열기 실패: {str(e)}")
            return None

    def edit_mapping(self):
        """
        매핑 표를 수정할 수 있는 다이얼로그를 엽니다.
        """
        if self.df_source is None:
            QMessageBox.critical(self, "오류", "먼저 파일을 열어주세요.")
            logging.error("파일이 열려 있지 않습니다.")
            return

        dialog = MappingEditorDialog(self.mapping, self.df_source.columns.tolist(), self)
        if dialog.exec_() == QDialog.Accepted:
            self.mapping = dialog.get_mapping()

            # 로그 추가: 수정된 매핑 확인
            logging.info(f"수정된 매핑: {self.mapping}")

            # 선택된 옵션을 가져와서 self.mappings에 업데이트
            selected_option = self.option_combo.currentText()
            self.mappings[selected_option] = self.mapping.copy()

            # 로그 추가: 전체 매핑 상태 확인
            logging.info(f"선택된 옵션 '{selected_option}'의 전체 매핑: {self.mappings[selected_option]}")

            QMessageBox.information(self, "저장 완료", "매칭 표가 저장되었습니다.")
            logging.info("매칭 표가 저장되었습니다.")

    def select_encoding(self):
        """
        사용자가 CSV 파일을 저장할 때 인코딩을 선택할 수 있는 다이얼로그를 표시합니다.
        """
        items = ["utf-8-sig", "cp949"]
        item, ok = QInputDialog.getItem(self, "인코딩 선택", "CSV 파일 인코딩을 선택하세요:", items, 0, False)
        if ok and item:
            return item
        return None

    def create_a_file(self):
        """
        매핑을 기반으로 새로운 파일을 생성합니다.
        """
        if self.df_source is None:
            QMessageBox.critical(self, "오류", "먼저 파일을 열어주세요.")
            logging.error("파일이 열려 있지 않습니다.")
            return

        # 모델에서 최신 DataFrame 가져오기
        df_source_updated = self.model.get_dataframe()

        columns_a = list(self.mapping.keys())
        df_target = pd.DataFrame(columns=columns_a)

        for a_col, b_col in self.mapping.items():
            if b_col and b_col in df_source_updated.columns:
                df_target[a_col] = df_source_updated[b_col]
                logging.info(f"열 '{b_col}'이 '{a_col}'으로 매핑되었습니다.")
            else:
                df_target[a_col] = pd.NA  # 매핑되지 않은 경우 NaN으로 채움
                if b_col == "":
                    logging.info(f"열 '{a_col}'은 소스 컬럼이 없으므로 NaN으로 설정됩니다.")
                else:
                    logging.warning(f"열 '{b_col}'이 데이터 파일에 존재하지 않거나 비어 있습니다. 열 '{a_col}'은 NaN으로 설정됩니다.")

        # 매핑된 데이터 프레임의 컬럼 확인
        logging.info(f"타겟 데이터 프레임 컬럼: {df_target.columns.tolist()}")

        # 진행 바 표시
        self.status_progress_bar.setVisible(True)
        self.status_progress_bar.setMaximum(0)  # Indeterminate
        self.status_bar.showMessage("파일을 생성하는 중...")

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "파일 저장",
            os.path.expanduser("~"),
            "CSV 파일 (*.csv);;Excel 파일 (*.xlsx)"
        )
        if file_path:
            file_ext = os.path.splitext(file_path)[1].lower()
            try:
                if file_ext == ".csv":
                    encoding = self.select_encoding()
                    if not encoding:
                        QMessageBox.warning(self, "경고", "인코딩 선택이 취소되었습니다.")
                        self.status_progress_bar.setVisible(False)
                        self.status_bar.clearMessage()
                        return
                    df_target.to_csv(file_path, index=False, encoding=encoding)
                elif file_ext == ".xlsx":
                    df_target.to_excel(file_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "저장 완료", "파일이 생성되었습니다.")
                logging.info(f"파일 '{file_path}'이 성공적으로 생성되었습니다.")
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", f"파일을 저장하는 데 오류가 발생했습니다: {str(e)}")
                logging.error(f"파일 저장 실패: {str(e)}")
        self.status_progress_bar.setVisible(False)
        self.status_bar.clearMessage()


    def save_without_mapping(self):
        if self.df_source is None:
            QMessageBox.critical(self, "오류", "먼저 파일을 열어주세요.")
            logging.error("파일이 열려 있지 않습니다.")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "파일 저장",
            os.path.expanduser("~"),
            "CSV 파일 (*.csv);;Excel 파일 (*.xlsx)"
        )

        if file_path:
            try:
                file_ext = os.path.splitext(file_path)[1].lower()
                # 수정된 데이터프레임 가져오기
                df_to_save = self.model.get_dataframe()
                if file_ext == ".csv":
                    encoding = self.select_encoding()
                    if not encoding:
                        QMessageBox.warning(self, "경고", "인코딩 선택이 취소되었습니다.")
                        return
                    df_to_save.to_csv(file_path, index=False, encoding=encoding)
                elif file_ext == ".xlsx":
                    df_to_save.to_excel(file_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "저장 완료", "파일이 성공적으로 저장되었습니다.")
                logging.info(f"파일 '{file_path}'이 성공적으로 저장되었습니다.")
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", f"파일을 저장하는 중 오류가 발생했습니다: {str(e)}")
                logging.error(f"파일 저장 실패: {str(e)}")


    def save_mapping_to_file(self):
        """
        현재 매핑 상태를 JSON 파일로 저장합니다.
        """
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "매핑 저장",
            os.path.expanduser("~"),
            "JSON 파일 (*.json)"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.mappings, f, ensure_ascii=False, indent=4)  # self.mappings 전체를 저장
                QMessageBox.information(self, "저장 완료", "매칭 상태가 파일에 저장되었습니다.")
                logging.info(f"매칭 상태가 '{file_path}'에 저장되었습니다.")
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", f"매칭 상태를 저장하는 데 오류가 발생했습니다: {str(e)}")
                logging.error(f"매칭 상태 저장 실패: {str(e)}")

    def load_mapping_from_file(self):
        """
        JSON 파일에서 매핑 상태를 불러옵니다.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "매핑 불러오기",
            os.path.expanduser("~"),
            "JSON 파일 (*.json)"
        )
        if not file_path:
            logging.warning("매핑 파일 선택이 취소되었습니다.")
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                loaded_mappings = json.load(f)  # 전체 매핑을 불러옴

            # 현재 선택된 옵션에 맞게 self.mappings 업데이트
            selected_option = self.option_combo.currentText()
            if selected_option in loaded_mappings:
                self.mappings[selected_option] = loaded_mappings[selected_option]
                self.mapping = self.mappings[selected_option].copy()
                self.update_mapping(selected_option)
                QMessageBox.information(self, "불러오기 완료", "매칭 상태가 불러와졌습니다.")
                logging.info(f"매칭 상태가 '{file_path}'에서 불러와졌습니다.")
            else:
                QMessageBox.warning(self, "경고", f"불러온 파일에 '{selected_option}' 매핑이 없습니다.")
                logging.warning(f"불러온 파일에 '{selected_option}' 매핑이 없습니다.")
        except json.JSONDecodeError:
            QMessageBox.critical(self, "오류", "매칭 상태 파일을 읽을 수 없습니다.")
            logging.error("매칭 상태 파일 JSON 디코딩 실패.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"매핑을 불러오는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"매핑 불러오기 실패: {str(e)}")

    def display_data(self):
        """
        로드된 데이터를 테이블에 표시합니다.
        """
        if self.df_source is not None:
            # 인덱스를 데이터의 일부로 포함시키지 않고 원본 데이터 프레임 사용
            df_display = self.df_source.copy()

            self.model.set_dataframe(df_display)
            self.table_view.setModel(self.model)  # 모델을 다시 설정
            self.table_view.resizeColumnsToContents()
            logging.info("데이터가 테이블에 성공적으로 표시되었습니다.")

            # 행 인덱스 정보 업데이트
            total_rows = self.model.rowCount()
            self.row_info_label.setText(f"총 행 수: {total_rows}")

            # 행 헤더 표시 (이미 init_ui에서 설정)
            self.table_view.verticalHeader().setVisible(True)

    def sort_original_order(self):
        """
        원본 인덱스 순으로 데이터를 정렬합니다.
        """
        if self.df_source is None:
            QMessageBox.warning(self, "경고", "먼저 파일을 열어주세요.")
            logging.warning("파일이 열려 있지 않습니다.")
            return
        self.model.sort_original_order()
        logging.info("데이터가 원본 순서대로 정렬되었습니다.")


    def open_context_menu(self, position):
        """
        선택된 셀 범위를 우클릭할 때 컨텍스트 메뉴를 열고 선택된 범위에 적용되는 로직을 추가할 수 있는 기본 구조를 제공합니다.
        """
        selected_indexes = self.table_view.selectedIndexes()
        if not selected_indexes:
            # 선택된 셀이 없을 경우, 현재 클릭한 셀만 처리
            index = self.table_view.indexAt(position)
            if not index.isValid():
                return
            selected_indexes = [index]

        # 컨텍스트 메뉴 생성
        menu = QMenu(self)

        # 데이터 보기 액션 추가
        view_data_action = QAction("데이터 보기", self)
        view_data_action.triggered.connect(lambda: self.show_selected_data(selected_indexes))
        menu.addAction(view_data_action)

        # 선택된 셀의 내용을 복사
        copy_action = QAction("복사", self)
        copy_action.triggered.connect(lambda: self.copy_selected_cells(selected_indexes))
        menu.addAction(copy_action)

        # 붙여넣기 
        paste_action = QAction("붙여넣기", self)
        paste_action.triggered.connect(lambda: self.paste_clipboard_data(selected_indexes))
        menu.addAction(paste_action)
        
        # 선택한 셀의 데이터 지우기
        add_clear_action = QAction("데이터 지우기", self)
        add_clear_action.triggered.connect(lambda: self.clear_selected_cells(selected_indexes))
        menu.addAction(add_clear_action)

        # 선택한 셀의 데이터 앞에 공백 추가
        add_space_action = QAction("앞에 공백 추가", self)
        add_space_action.triggered.connect(lambda: self.front_space_selected_cells(selected_indexes, lambda x: ' ' + x))
        menu.addAction(add_space_action)

        # 이미지 처리 (src 변경 및 Base64 이미지 다운로드)
        process_images_action = QAction("이미지 처리 (src값 변경 및 다운로드)", self)
        process_images_action.triggered.connect(lambda: self.process_images(selected_indexes))
        menu.addAction(process_images_action)



        # 예:
        # custom_action = QAction("커스텀 작업", self)
        # custom_action.triggered.connect(lambda: self.custom_logic(selected_indexes))
        # menu.addAction(custom_action)

        menu.exec_(self.table_view.viewport().mapToGlobal(position))

    def is_base64_image(self, data):
        """
        Base64 인코딩된 이미지인지 확인하는 함수
        """
        try:
            # Base64 인코딩 여부 확인
            if isinstance(data, str) and len(data) > 100:  # 기본적인 길이 체크
                base64_bytes = base64.b64decode(data, validate=True)
                img = Image.open(BytesIO(base64_bytes))
                return True, base64_bytes
            return False, None
        except Exception:
            return False, None        
        
    def show_selected_data(self, selected_indexes):
        """
        선택된 셀 범위의 데이터를 별도 팝업 창으로 표시합니다.
        """
        if not selected_indexes:
            QMessageBox.warning(self, "경고", "선택된 셀이 없습니다.")
            return

        # 셀을 행별로 정렬
        selected_indexes = sorted(selected_indexes, key=lambda x: (x.row(), x.column()))

        # 행별로 그룹화
        rows = {}
        for index in selected_indexes:
            if index.row() not in rows:
                rows[index.row()] = {}
            rows[index.row()][index.column()] = self.model.data(index, Qt.DisplayRole)

        # 팝업 창 생성
        dialog = QDialog(self)
        dialog.setWindowTitle("선택된 데이터 보기")
        dialog.setMinimumSize(800, 800)
        layout = QVBoxLayout()

        # 이미지가 포함된 셀 처리
        for row in sorted(rows.keys()):
            for col in sorted(rows[row].keys()):
                value = rows[row][col]
                is_image, image_data = self.is_base64_image(value)
                if is_image:
                    # Base64 이미지인 경우 이미지로 표시
                    image = Image.open(BytesIO(image_data))
                    qimage = QPixmap.fromImage(image)
                    image_label = QLabel()
                    image_label.setPixmap(qimage)
                    layout.addWidget(image_label)
                else:
                    # Base64 이미지가 아닌 경우 텍스트로 표시
                    text_edit = QTextEdit()
                    text_edit.setReadOnly(True)
                    text_edit.setText(value)
                    layout.addWidget(text_edit)

        close_button = QPushButton("닫기")
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)

        dialog.setLayout(layout)
        dialog.exec_()

    def copy_selected_cells(self, selected_indexes):
        """
        선택된 셀의 내용을 클립보드에 복사합니다.
        """
        if not selected_indexes:
            return

        # 셀을 행별로 정렬
        selected_indexes = sorted(selected_indexes, key=lambda x: (x.row(), x.column()))

        # 행별로 그룹화
        rows = {}
        for index in selected_indexes:
            if index.row() not in rows:
                rows[index.row()] = {}
            rows[index.row()][index.column()] = self.model.data(index, Qt.DisplayRole)

        # 문자열로 변환 (줄 바꿈을 한 번에 추가하여 마지막에 빈 줄이 생기지 않도록 함)
        clipboard_rows = []
        for row in sorted(rows.keys()):
            row_data = [rows[row][col] for col in sorted(rows[row].keys())]
            clipboard_rows.append("\t".join(row_data))
        clipboard_text = "\n".join(clipboard_rows)

        # 클립보드에 복사
        clipboard = QApplication.clipboard()
        clipboard.setText(clipboard_text)
        # QMessageBox.information(self, "복사 완료", "선택된 셀의 내용이 클립보드에 복사되었습니다.")
        # logging.info("선택된 셀의 내용이 클립보드에 복사되었습니다.")


    def select_column(self, column):
        """
        열 헤더를 클릭했을 때 해당 열 전체를 선택합니다.
        """
        selection_model = self.table_view.selectionModel()
        selection = QItemSelection()
        top_left = self.model.index(0, column)
        bottom_right = self.model.index(self.model.rowCount() - 1, column)
        selection.select(top_left, bottom_right)
        selection_model.select(selection, QItemSelectionModel.Select | QItemSelectionModel.Columns)

    def select_row(self, row):
        """
        행 헤더를 클릭했을 때 해당 행 전체를 선택합니다.
        """
        selection_model = self.table_view.selectionModel()
        selection = QItemSelection()
        top_left = self.model.index(row, 0)
        bottom_right = self.model.index(row, self.model.columnCount() - 1)
        selection.select(top_left, bottom_right)
        selection_model.select(selection, QItemSelectionModel.Select | QItemSelectionModel.Rows)

    def process_images(self, selected_indexes):
        """
        선택된 셀들에서 img 태그의 src 값을 변경하고, Base64 인코딩된 이미지를 다운로드합니다.
        진행 상황을 프로그레스 바와 실시간 로그로 표시합니다.
        """
        if not selected_indexes:
            return

        # 사용자로부터 새로운 경로 입력 받기
        new_path, ok = QInputDialog.getText(self, "이미지 src 변경", "새로운 이미지 경로를 입력하세요:")
        if not ok or not new_path:
            return

        # 이미지 다운로드 여부 확인
        download_choice = QMessageBox.question(
            self,
            "이미지 다운로드 여부",
            "선택된 이미지의 다운로드를 진행하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        download_images = download_choice == QMessageBox.Yes

        # 다운로드 폴더 경로 요청
        if download_images:
            download_folder = QFileDialog.getExistingDirectory(
                self,
                "이미지 다운로드 폴더 선택",
                os.getcwd(),
                QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
            )
            if not download_folder:
                QMessageBox.warning(self, "경고", "이미지 다운로드를 취소했습니다.")
                download_images = False
        else:
            download_folder = None

        # Base64 이미지 다운로드를 위한 폴더 설정
        base64_download_folder = os.path.join(os.getcwd(), "base64이미지")
        os.makedirs(base64_download_folder, exist_ok=True)
        logging.info(f"Base64 이미지 다운로드 폴더: {base64_download_folder}")

        # UI 요소 초기화
        self.main_progress_bar.setValue(0)
        self.main_progress_bar.setVisible(True)  # 프로그레스 바 표시

        # 스레드와 워커 초기화
        self.thread = QThread()
        self.worker = ImageProcessor(
            selected_indexes,
            new_path,
            download_images,
            download_folder,
            base64_download_folder,
            self.model
        )
        self.worker.moveToThread(self.thread)

        # 신호 연결
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_main_progress)
        self.worker.log.connect(self.update_log)
        self.worker.update_cell.connect(self.handle_update_cell)  # 새로 추가된 시그널 연결
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.error.connect(self.update_log)  

        # 스레드 시작
        self.thread.start()

        # 스레드가 끝났을 때의 처리
        self.thread.finished.connect(lambda: self.on_processing_finished())

    def handle_update_cell(self, index, new_data):
        """
        ImageProcessor로부터 전달된 셀 업데이트 데이터를 처리하는 슬롯.
        """
        self.model.setData(index, new_data, Qt.EditRole)


    def update_progress(self, value):
        """
        프로그레스 바 업데이트
        """
        self.status_progress_bar.setValue(value)

    def update_main_progress(self, value):
        """
        메인 레이아웃의 프로그레스 바 업데이트
        """
        self.main_progress_bar.setValue(value)


    def on_processing_finished(self):
        """
        이미지 처리 작업이 완료되었을 때 호출되는 메서드
        """
        self.main_progress_bar.setValue(100)  # 100%로 설정
        self.main_progress_bar.setVisible(False)  # 프로그레스 바 숨김
        QMessageBox.information(self, "처리 완료", "선택된 셀들의 이미지가 처리되었습니다.")
        logging.info("선택된 셀들의 이미지가 처리되었습니다.")

    def update_log(self, message):
        """
        로그 텍스트 에디트 업데이트
        """
        if self.log_text_edit.isVisible():
            if "실패" in message or "Error" in message:
                formatted_message = f"<span style='color:red;'>{message}</span>"
            else:
                formatted_message = f"<span style='color:green;'>{message}</span>"
            self.log_text_edit.append(formatted_message)
            # 자동 스크롤
            self.log_text_edit.verticalScrollBar().setValue(self.log_text_edit.verticalScrollBar().maximum())

    def handle_error(self, error_message):
        """
        오류 메시지를 처리하는 슬롯입니다.
        """
        self.log_text_edit.append(f"<span style='color:red;'>오류: {error_message}</span>")

    def front_space_selected_cells(self, selected_indexes, transform_function):
        """
        선택된 셀의 데이터 앞에 공백을 추가하는 메서드.
        transform_function: 각 셀 데이터에 적용할 함수 (예: lambda x: ' ' + x)
        """
        try:
            for index in selected_indexes:
                current_data = self.model.data(index, Qt.DisplayRole)
                if isinstance(current_data, str) and current_data.strip() != '':
                    transformed_data = transform_function(current_data)
                    self.model.setData(index, transformed_data, Qt.EditRole)
                    # logging.info(f"셀({index.row()}, {index.column()})의 데이터가 '{current_data}'에서 '{transformed_data}'로 변경되었습니다.")
                else:
                    logging.warning(f"셀({index.row()}, {index.column()})의 데이터가 문자열이 아니거나 비어 있습니다. 변환을 건너뜁니다.")
            QMessageBox.information(self, "변환 완료", "선택된 셀의 내용이 변환되었습니다.")
            logging.info("선택된 셀의 내용이 변환되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "변환 실패", f"셀 데이터를 변환하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"셀 데이터 변환 실패: {str(e)}")

    def clear_selected_cells(self, selected_indexes):
        """
        선택된 셀의 데이터를 빈 값으로 만드는 메서드.
        """
        try:
            if self.undo_stack is not None:
                command = DeleteCellsCommand(self.model, selected_indexes)
                self.undo_stack.push(command)
            else:
                self.model.clear_selected_cells_bulk(selected_indexes)
            # QMessageBox.information(self, "변환 완료", "선택된 셀의 내용이 빈 값으로 변환되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "변환 실패", f"셀 데이터를 빈 값으로 변환하는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"셀 데이터 변환 실패: {str(e)}")

    def paste_clipboard_data(self, selected_indexes):
        try:
            # 클립보드에서 데이터를 가져옴
            clipboard = QApplication.clipboard()
            clipboard_text = clipboard.text()

            if not clipboard_text:
                QMessageBox.warning(self, "클립보드 없음", "클립보드에 데이터가 없습니다.")
                return

            # 클립보드 데이터를 줄과 탭을 기준으로 나눔
            rows = clipboard_text.strip().split('\n')  # 양 끝 공백 및 줄 바꿈 제거
            data_matrix = [row.split('\t') for row in rows if row.strip()]  # 빈 행 무시

            if not data_matrix:
                QMessageBox.warning(self, "붙여넣기 실패", "붙여넣을 데이터가 유효하지 않습니다.")
                return

            # 클립보드 데이터가 단일 값인지 확인
            if len(data_matrix) == 1 and len(data_matrix[0]) == 1:
                # 단일 값인 경우, 선택된 모든 셀에 동일한 값을 붙여넣음
                paste_value = data_matrix[0][0]
                command = PasteMultipleCellsCommand(self.model, selected_indexes, paste_value)
            else:
                # 다중 값인 경우, 기존 동작대로 영역에 맞게 붙여넣음
                # 시작점을 첫 번째 선택된 셀로 설정
                start_index = selected_indexes[0]
                command = PasteCellsCommand(self.model, data_matrix, self.model.index(start_index.row(), start_index.column()))

            # Undo 스택을 사용하여 명령을 푸시
            if self.undo_stack is not None:
                self.undo_stack.push(command)
            else:
                # Undo 스택이 없는 경우, 직접 실행
                command.redo()

            logging.info("클립보드 데이터가 성공적으로 붙여넣어졌습니다.")
        except Exception as e:
            QMessageBox.critical(self, "붙여넣기 실패", f"클립보드 데이터를 붙여넣는 중 오류가 발생했습니다: {str(e)}")
            logging.error(f"클립보드 데이터 붙여넣기 실패: {str(e)}")




    # def transform_selected_cells(self, selected_indexes, transform_function):
    #     """
    #     선택된 셀의 데이터를 변환하는 메서드.
    #     transform_function: 각 셀 데이터에 적용할 함수 (예: lambda x: ' ' + x)
    #     """
    #     try:
    #         for index in selected_indexes:
    #             current_data = self.model.data(index, Qt.DisplayRole)
    #             if current_data:
    #                 transformed_data = transform_function(current_data)
    #                 self.model.setData(index, transformed_data, Qt.EditRole)
    #                 logging.info(f"셀({index.row()}, {index.column()})의 데이터가 '{current_data}'에서 '{transformed_data}'로 변경되었습니다.")
    #         QMessageBox.information(self, "변환 완료", "선택된 셀의 내용이 변환되었습니다.")
    #         logging.info("선택된 셀의 내용이 변환되었습니다.")
    #     except Exception as e:
    #         QMessageBox.critical(self, "변환 실패", f"셀 데이터를 변환하는 중 오류가 발생했습니다: {str(e)}")
    #         logging.error(f"셀 데이터 변환 실패: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = CSVMatcherApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
