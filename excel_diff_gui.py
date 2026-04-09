#!/usr/bin/env python3
from __future__ import annotations

import gc
import re
import sys
import time
import traceback
import unicodedata
from pathlib import Path
from typing import Any

from PySide6.QtCore import QElapsedTimer, QObject, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QCloseEvent, QImage, QPixmap, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from excel_diff import SUPPORTED_EXTENSIONS, generate_highlight_workbook


LOG_LINE_LIMIT = 1200
APP_DIR = Path(__file__).resolve().parent
MASCOT_DIR = APP_DIR / "assets" / "mascot"
CHELSEA_DIR = APP_DIR / "assets" / "chelsea"
THEME_GREEN = "green"
THEME_CUTE = "cute"
THEME_CHELSEA = "chelsea"
EASTER_EGG_COMMAND = "먼작귀"
CHELSEA_EASTER_EGG_COMMAND = "첼시"

HANGUL_CHO = [
    "ㄱ",
    "ㄲ",
    "ㄴ",
    "ㄷ",
    "ㄸ",
    "ㄹ",
    "ㅁ",
    "ㅂ",
    "ㅃ",
    "ㅅ",
    "ㅆ",
    "ㅇ",
    "ㅈ",
    "ㅉ",
    "ㅊ",
    "ㅋ",
    "ㅌ",
    "ㅍ",
    "ㅎ",
]
HANGUL_JUNG = [
    "ㅏ",
    "ㅐ",
    "ㅑ",
    "ㅒ",
    "ㅓ",
    "ㅔ",
    "ㅕ",
    "ㅖ",
    "ㅗ",
    "ㅘ",
    "ㅙ",
    "ㅚ",
    "ㅛ",
    "ㅜ",
    "ㅝ",
    "ㅞ",
    "ㅟ",
    "ㅠ",
    "ㅡ",
    "ㅢ",
    "ㅣ",
]
HANGUL_JONG = [
    "",
    "ㄱ",
    "ㄲ",
    "ㄳ",
    "ㄴ",
    "ㄵ",
    "ㄶ",
    "ㄷ",
    "ㄹ",
    "ㄺ",
    "ㄻ",
    "ㄼ",
    "ㄽ",
    "ㄾ",
    "ㄿ",
    "ㅀ",
    "ㅁ",
    "ㅂ",
    "ㅄ",
    "ㅅ",
    "ㅆ",
    "ㅇ",
    "ㅈ",
    "ㅊ",
    "ㅋ",
    "ㅌ",
    "ㅍ",
    "ㅎ",
]
HANGUL_JUNG_COMBINE = {
    ("ㅗ", "ㅏ"): "ㅘ",
    ("ㅗ", "ㅐ"): "ㅙ",
    ("ㅗ", "ㅣ"): "ㅚ",
    ("ㅜ", "ㅓ"): "ㅝ",
    ("ㅜ", "ㅔ"): "ㅞ",
    ("ㅜ", "ㅣ"): "ㅟ",
    ("ㅡ", "ㅣ"): "ㅢ",
}
HANGUL_CHO_INDEX = {value: idx for idx, value in enumerate(HANGUL_CHO)}
HANGUL_JUNG_INDEX = {value: idx for idx, value in enumerate(HANGUL_JUNG)}
HANGUL_JONG_INDEX = {value: idx for idx, value in enumerate(HANGUL_JONG)}
HANGUL_CHO_SET = set(HANGUL_CHO)
HANGUL_JUNG_SET = set(HANGUL_JUNG)
CELL_ADDRESS_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$")
CELL_LOOKUP_LOG_LIMIT = 40


def now_text() -> str:
    return time.strftime("%H:%M:%S")


def compose_compat_hangul(text: str) -> str:
    if not text:
        return text

    result: list[str] = []
    index = 0
    text_len = len(text)

    while index < text_len:
        first = text[index]
        if first not in HANGUL_CHO_SET or index + 1 >= text_len or text[index + 1] not in HANGUL_JUNG_SET:
            result.append(first)
            index += 1
            continue

        choseong = first
        index += 1

        jungseong = text[index]
        index += 1

        if index < text_len and text[index] in HANGUL_JUNG_SET:
            combined_jung = HANGUL_JUNG_COMBINE.get((jungseong, text[index]))
            if combined_jung:
                jungseong = combined_jung
                index += 1

        jongseong = ""
        if index < text_len and text[index] in HANGUL_CHO_SET:
            tail = text[index]
            can_use_tail = tail in HANGUL_JONG_INDEX and not (
                index + 1 < text_len and text[index + 1] in HANGUL_JUNG_SET
            )
            if can_use_tail:
                jongseong = tail
                index += 1

        code = 0xAC00 + ((HANGUL_CHO_INDEX[choseong] * 21) + HANGUL_JUNG_INDEX[jungseong]) * 28 + HANGUL_JONG_INDEX[jongseong]
        result.append(chr(code))

    return "".join(result)


def normalize_command_text(text: str) -> str:
    normalized = unicodedata.normalize("NFC", text or "")
    normalized = (
        normalized.replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
        .strip()
    )
    normalized = compose_compat_hangul(normalized)
    if normalized.startswith(">"):
        normalized = normalized[1:].strip()
    normalized = "".join(ch for ch in normalized if not ch.isspace())
    return normalized


def extract_excel_paths_from_mime_data(mime_data: Any) -> list[Path]:
    paths: list[Path] = []

    if mime_data is None:
        return paths

    if mime_data.hasUrls():
        for url in mime_data.urls():
            if not url.isLocalFile():
                continue

            candidate = Path(url.toLocalFile()).expanduser()
            if candidate.is_file() and candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
                paths.append(candidate)
                continue

            if candidate.is_dir():
                try:
                    for child in sorted(candidate.iterdir(), key=lambda p: (p.name.lower(), str(p).lower())):
                        if child.is_file() and child.suffix.lower() in SUPPORTED_EXTENSIONS:
                            paths.append(child)
                except OSError:
                    continue

    if not paths and mime_data.hasText():
        text = mime_data.text().strip()
        if text:
            raw = text.strip().strip('"').strip("'")
            candidate = Path(raw).expanduser()
            if candidate.is_file() and candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
                paths.append(candidate)

    # 순서 유지 중복 제거
    unique: list[Path] = []
    seen: set[Path] = set()
    for path in paths:
        if path in seen:
            continue
        seen.add(path)
        unique.append(path)
    return unique


class DropComboBox(QComboBox):
    fileDropped = Signal(str)
    dragHoverChanged = Signal(bool)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self._hovering = False

    def _set_hovering(self, hovering: bool) -> None:
        if self._hovering == hovering:
            return
        self._hovering = hovering
        self.dragHoverChanged.emit(hovering)

    def dragEnterEvent(self, event: Any) -> None:
        if extract_excel_paths_from_mime_data(event.mimeData()):
            self._set_hovering(True)
            event.acceptProposedAction()
        else:
            self._set_hovering(False)
            event.ignore()

    def dragMoveEvent(self, event: Any) -> None:
        if extract_excel_paths_from_mime_data(event.mimeData()):
            self._set_hovering(True)
            event.acceptProposedAction()
        else:
            self._set_hovering(False)
            event.ignore()

    def dragLeaveEvent(self, event: Any) -> None:
        self._set_hovering(False)
        event.accept()

    def dropEvent(self, event: Any) -> None:
        paths = extract_excel_paths_from_mime_data(event.mimeData())
        if not paths:
            self._set_hovering(False)
            event.ignore()
            return

        self.fileDropped.emit(str(paths[0]))
        self._set_hovering(False)
        event.acceptProposedAction()


class TerminalConsole(QPlainTextEdit):
    commandEntered = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("logView")
        self.setReadOnly(False)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setTextInteractionFlags(Qt.TextInteractionFlag.TextEditorInteraction)
        self.setUndoRedoEnabled(True)
        self.setAttribute(Qt.WidgetAttribute.WA_InputMethodEnabled, True)

        self._prompt = "> "
        self._history: list[str] = []
        self._history_index = 0
        self._set_content([], "")
        QTimer.singleShot(0, self.ensure_input_focus)

    def append_log(self, message: str) -> None:
        scrollbar = self.verticalScrollBar()
        previous_value = scrollbar.value()
        previous_max = scrollbar.maximum()
        was_at_bottom = previous_value >= max(0, previous_max - 2)

        cursor = self.textCursor()
        input_start = self._input_start_position()
        if cursor.anchor() >= input_start and cursor.position() >= input_start:
            anchor_offset = cursor.anchor() - input_start
            position_offset = cursor.position() - input_start
        else:
            anchor_offset = None
            position_offset = None

        log_lines = self._get_log_lines()
        for line in str(message).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            log_lines.append(line)
        if len(log_lines) > LOG_LINE_LIMIT:
            log_lines = log_lines[-LOG_LINE_LIMIT:]

        current_input = self._current_input_text()
        self._set_content(log_lines, current_input)

        if anchor_offset is not None and position_offset is not None:
            new_input_start = self._input_start_position()
            restored = self.textCursor()
            anchor_pos = new_input_start + min(max(anchor_offset, 0), len(current_input))
            position_pos = new_input_start + min(max(position_offset, 0), len(current_input))
            restored.setPosition(anchor_pos)
            restored.setPosition(position_pos, QTextCursor.MoveMode.KeepAnchor)
            self.setTextCursor(restored)

        if was_at_bottom:
            scrollbar.setValue(scrollbar.maximum())
        else:
            delta = max(0, scrollbar.maximum() - previous_max)
            scrollbar.setValue(previous_value + delta)

    def clear_logs(self) -> None:
        self._set_content([], "")
        self._history_index = len(self._history)
        self.ensure_input_focus()

    def ensure_input_focus(self) -> None:
        if not self.hasFocus():
            self.setFocus(Qt.FocusReason.OtherFocusReason)
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.setTextCursor(cursor)
        self.ensureCursorVisible()

    def _navigate_history(self, direction: int) -> None:
        if not self._history:
            return

        if direction < 0 and self._history_index > 0:
            self._history_index -= 1
        elif direction > 0 and self._history_index < len(self._history):
            self._history_index += 1
        else:
            return

        if self._history_index >= len(self._history):
            self._set_current_input_text("")
            return

        history_item = self._history[self._history_index]
        self._set_current_input_text(history_item)

    def keyPressEvent(self, event: Any) -> None:
        key = event.key()
        modifiers = event.modifiers()

        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self._submit_current_command()
            return

        if (modifiers & Qt.KeyboardModifier.ControlModifier) and key == Qt.Key.Key_L:
            self.clear_logs()
            return

        if key == Qt.Key.Key_Up and self._cursor_is_on_input():
            self._navigate_history(-1)
            return

        if key == Qt.Key.Key_Down and self._cursor_is_on_input():
            self._navigate_history(1)
            return

        if key == Qt.Key.Key_Home and not (modifiers & Qt.KeyboardModifier.ControlModifier):
            cursor = self.textCursor()
            input_start = self._input_start_position()
            if modifiers & Qt.KeyboardModifier.ShiftModifier:
                cursor.setPosition(input_start, QTextCursor.MoveMode.KeepAnchor)
            else:
                cursor.setPosition(input_start)
            self.setTextCursor(cursor)
            return

        if key == Qt.Key.Key_Left and not self.textCursor().hasSelection():
            if self.textCursor().position() <= self._input_start_position():
                return

        if key == Qt.Key.Key_Backspace and not self.textCursor().hasSelection():
            if self.textCursor().position() <= self._input_start_position():
                return

        if self._is_editing_key(event):
            self._prepare_cursor_for_edit()

        super().keyPressEvent(event)
        self._enforce_cursor_boundary()
        if self._cursor_is_on_input():
            self._history_index = len(self._history)

    def insertFromMimeData(self, source: Any) -> None:
        self._prepare_cursor_for_edit()
        raw = source.text() if source is not None else ""
        sanitized = (raw or "").replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
        self.textCursor().insertText(sanitized)
        self._enforce_cursor_boundary()
        self._history_index = len(self._history)

    def cut(self) -> None:
        cursor = self.textCursor()
        if not cursor.hasSelection():
            return
        input_start = self._input_start_position()
        if cursor.selectionEnd() <= input_start:
            return
        if cursor.selectionStart() < input_start:
            cursor.setPosition(input_start)
            cursor.setPosition(cursor.selectionEnd(), QTextCursor.MoveMode.KeepAnchor)
            self.setTextCursor(cursor)
        super().cut()
        self._enforce_cursor_boundary()
        self._history_index = len(self._history)

    def _submit_current_command(self) -> None:
        command = compose_compat_hangul(self._current_input_text())
        if command:
            if not self._history or self._history[-1] != command:
                self._history.append(command)
        self._history_index = len(self._history)

        log_lines = self._get_log_lines()
        log_lines.append(f"{self._prompt}{command}")
        if len(log_lines) > LOG_LINE_LIMIT:
            log_lines = log_lines[-LOG_LINE_LIMIT:]
        self._set_content(log_lines, "")
        self.ensure_input_focus()
        self.commandEntered.emit(command)

    def _is_editing_key(self, event: Any) -> bool:
        key = event.key()
        modifiers = event.modifiers()
        if key in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
            return True
        if (modifiers & Qt.KeyboardModifier.ControlModifier) and key in (Qt.Key.Key_V, Qt.Key.Key_X):
            return True
        text = event.text()
        if not text:
            return False
        blocked = Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.MetaModifier | Qt.KeyboardModifier.AltModifier
        return not (modifiers & blocked)

    def _prepare_cursor_for_edit(self) -> None:
        cursor = self.textCursor()
        input_start = self._input_start_position()

        if cursor.hasSelection():
            selection_start = cursor.selectionStart()
            selection_end = cursor.selectionEnd()
            if selection_end <= input_start:
                cursor.clearSelection()
                cursor.movePosition(QTextCursor.MoveOperation.End)
            elif selection_start < input_start:
                cursor.setPosition(input_start)
                cursor.setPosition(selection_end, QTextCursor.MoveMode.KeepAnchor)
            self.setTextCursor(cursor)
            return

        if cursor.position() < input_start:
            cursor.movePosition(QTextCursor.MoveOperation.End)
            self.setTextCursor(cursor)

    def _enforce_cursor_boundary(self) -> None:
        cursor = self.textCursor()
        if cursor.hasSelection():
            return
        input_start = self._input_start_position()
        if cursor.position() < input_start:
            cursor.setPosition(input_start)
            self.setTextCursor(cursor)

    def _cursor_is_on_input(self) -> bool:
        cursor = self.textCursor()
        input_start = self._input_start_position()
        if cursor.hasSelection():
            return cursor.selectionStart() >= input_start
        return cursor.position() >= input_start

    def _set_current_input_text(self, text: str) -> None:
        sanitized = (text or "").replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
        input_start = self._input_start_position()
        cursor = self.textCursor()
        cursor.setPosition(input_start)
        cursor.movePosition(QTextCursor.MoveOperation.End, QTextCursor.MoveMode.KeepAnchor)
        cursor.insertText(sanitized)
        self.setTextCursor(cursor)
        self.ensureCursorVisible()

    def _current_input_text(self) -> str:
        text = self.toPlainText()
        input_start = self._input_start_position()
        return text[input_start:]

    def _get_log_lines(self) -> list[str]:
        text = self.toPlainText()
        if not text:
            return []
        line_start = text.rfind("\n") + 1
        last_line = text[line_start:]
        if last_line.startswith(self._prompt):
            logs_text = text[: max(0, line_start - 1)]
        else:
            logs_text = text
        if not logs_text:
            return []
        return logs_text.split("\n")

    def _set_content(self, log_lines: list[str], input_text: str) -> None:
        cleaned_lines: list[str] = []
        for raw in log_lines:
            for line in str(raw).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
                cleaned_lines.append(line)
        if len(cleaned_lines) > LOG_LINE_LIMIT:
            cleaned_lines = cleaned_lines[-LOG_LINE_LIMIT:]

        sanitized_input = (input_text or "").replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
        rendered_lines = [*cleaned_lines, f"{self._prompt}{sanitized_input}"]
        self.setPlainText("\n".join(rendered_lines))
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.setTextCursor(cursor)

    def _input_start_position(self) -> int:
        text = self.toPlainText()
        line_start = text.rfind("\n") + 1
        if text[line_start:].startswith(self._prompt):
            return line_start + len(self._prompt)

        # 예상치 못한 편집으로 프롬프트가 깨졌으면 현재 마지막 줄을 입력으로 보존해 복구
        raw_input = text[line_start:]
        logs_text = text[: max(0, line_start - 1)]
        log_lines = logs_text.split("\n") if logs_text else []
        self._set_content(log_lines, raw_input)
        restored = self.toPlainText()
        restored_line_start = restored.rfind("\n") + 1
        return restored_line_start + len(self._prompt)


class CompareWorker(QObject):
    finished = Signal(dict)
    failed = Signal(str)

    def __init__(
        self,
        base_file: Path,
        target_file: Path,
        output_file: Path,
        include_format: bool,
        unprotect_second: bool,
    ) -> None:
        super().__init__()
        self.base_file = base_file
        self.target_file = target_file
        self.output_file = output_file
        self.include_format = include_format
        self.unprotect_second = unprotect_second

    def run(self) -> None:
        try:
            summary = generate_highlight_workbook(
                left_file=self.base_file,
                right_file=self.target_file,
                output_file=self.output_file,
                include_format=self.include_format,
                unprotect_second=self.unprotect_second,
            )
            payload = {
                "output_file": str(self.output_file),
                "modified_cells": summary.modified_cells,
                "added_cells": summary.added_cells,
                "removed_cells": summary.removed_cells,
                "changed_sheets": summary.changed_sheets,
                "diff_rows": summary.diff_rows,
            }
            self.finished.emit(payload)
        except Exception as exc:  # noqa: BLE001
            detail = f"{exc}\n\n{traceback.format_exc()}"
            self.failed.emit(detail)


class ExcelDiffMainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel Diff GUI")
        self.resize(1060, 860)
        self.setMinimumSize(940, 760)
        self.setAcceptDrops(True)

        self.base_file: Path | None = None
        self.target_file: Path | None = None
        self.output_file: Path | None = None
        self.output_file_manually_set = False

        self.recent_files: list[Path] = []
        self.base_folder_files: list[Path] = []
        self.target_folder_files: list[Path] = []
        self.base_choice_map: dict[str, Path] = {}
        self.target_choice_map: dict[str, Path] = {}
        self.last_diff_rows: list[tuple[str, str, str, str, str]] = []
        self.last_diff_by_cell: dict[str, list[tuple[str, str, str, str, str]]] = {}
        self.last_diff_by_sheet_cell: dict[str, list[tuple[str, str, str, str, str]]] = {}

        self.worker_thread: QThread | None = None
        self.worker: CompareWorker | None = None
        self.running = False

        self.progress_timer = QTimer(self)
        self.progress_timer.setInterval(400)
        self.progress_timer.timeout.connect(self._update_progress_text)
        self.elapsed_timer = QElapsedTimer()

        self._interactive_widgets: list[QWidget] = []
        self.mascot_paths = self._discover_mascot_paths()
        self.chelsea_paths = self._discover_chelsea_paths()
        self.current_theme = THEME_GREEN

        self._build_ui()
        self._apply_styles()
        self._initialize_dropdown_sources()
        self._append_log("준비 완료. 기준 파일과 비교 파일을 선택한 뒤 [비교 실행]을 누르세요.")
        self._append_log("콘솔 명령어: run, clear, reset, help, find <키워드>, <셀주소>, <시트명>!<셀주소>")
        QTimer.singleShot(0, self.log_view.setFocus)

    def _build_ui(self) -> None:
        root = QWidget(self)
        self.setCentralWidget(root)

        main = QVBoxLayout(root)
        main.setContentsMargins(16, 14, 16, 14)
        main.setSpacing(10)

        header_card = self._make_card()
        header_card.setObjectName("headerCard")
        header_card.setFixedHeight(110)
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(16, 10, 16, 10)
        header_layout.setSpacing(10)
        title = QLabel("엑셀 비교 도구")
        title.setObjectName("titleText")
        header_layout.addWidget(title)
        header_layout.addStretch(1)

        self.mascot_wrap = QWidget()
        mascot_layout = QHBoxLayout(self.mascot_wrap)
        mascot_layout.setContentsMargins(0, 0, 0, 0)
        mascot_layout.setSpacing(8)

        if self.mascot_paths:
            for mascot_path in self.mascot_paths[:3]:
                mascot_layout.addWidget(self._make_mascot_avatar(mascot_path))
        else:
            fallback = QLabel("캐릭터 이미지 없음")
            fallback.setObjectName("mutedText")
            mascot_layout.addWidget(fallback)

        header_layout.addWidget(self.mascot_wrap)

        self.chelsea_wrap = QWidget()
        chelsea_layout = QHBoxLayout(self.chelsea_wrap)
        chelsea_layout.setContentsMargins(0, 0, 0, 0)
        chelsea_layout.setSpacing(8)

        if self.chelsea_paths:
            for chelsea_path in self.chelsea_paths[:2]:
                chelsea_layout.addWidget(self._make_chelsea_avatar(chelsea_path))
        else:
            fallback = QLabel("첼시 이미지 없음")
            fallback.setObjectName("mutedText")
            chelsea_layout.addWidget(fallback)

        header_layout.addWidget(self.chelsea_wrap)
        main.addWidget(header_card)

        summary_row = QHBoxLayout()
        summary_row.setSpacing(8)
        self.base_name_label, self.base_path_label = self._make_selected_summary_card("기준 파일")
        self.target_name_label, self.target_path_label = self._make_selected_summary_card("비교 파일")
        summary_row.addWidget(self.base_summary_card, 1)
        summary_row.addWidget(self.target_summary_card, 1)
        main.addLayout(summary_row)

        picker_card = self._make_card()
        picker_grid = QGridLayout(picker_card)
        picker_grid.setContentsMargins(12, 12, 12, 12)
        picker_grid.setHorizontalSpacing(8)
        picker_grid.setVerticalSpacing(8)
        picker_grid.setColumnStretch(0, 1)
        picker_grid.setColumnStretch(1, 1)

        self.base_picker_block, self.base_combo = self._make_picker_block(
            parent_layout=picker_grid,
            column=0,
            title="기준 파일 선택",
            on_browse=self._pick_base_file,
            on_combo_change=self._on_base_combo_changed,
            on_drop_file=self._on_base_combo_file_dropped,
            on_drag_state=self._on_base_combo_drag_state,
        )
        self.target_picker_block, self.target_combo = self._make_picker_block(
            parent_layout=picker_grid,
            column=1,
            title="비교 파일 선택",
            on_browse=self._pick_target_file,
            on_combo_change=self._on_target_combo_changed,
            on_drop_file=self._on_target_combo_file_dropped,
            on_drag_state=self._on_target_combo_drag_state,
        )
        main.addWidget(picker_card)

        option_card = self._make_card()
        option_layout = QVBoxLayout(option_card)
        option_layout.setContentsMargins(12, 10, 12, 10)
        option_layout.setSpacing(8)

        option_title = QLabel("옵션")
        option_title.setProperty("sectionTitle", True)
        option_layout.addWidget(option_title)

        checks = QHBoxLayout()
        checks.setSpacing(16)

        self.include_format_check = QCheckBox("숫자 형식(number format) 차이도 변경으로 처리")
        self.unprotect_check = QCheckBox("비교 파일 보호/숨김 해제 후 비교")
        self.unprotect_check.setChecked(True)

        checks.addWidget(self.include_format_check)
        checks.addWidget(self.unprotect_check)
        checks.addStretch(1)
        option_layout.addLayout(checks)

        output_row = QHBoxLayout()
        output_row.setSpacing(8)

        output_label = QLabel("결과 파일")
        output_label.setFixedWidth(72)
        output_label.setProperty("fieldLabel", True)

        self.output_line = QLineEdit()
        self.output_line.setReadOnly(True)
        self.output_line.setPlaceholderText("자동 생성 예정")

        self.output_browse_button = QPushButton("경로 변경")
        self.output_browse_button.clicked.connect(self._pick_output_file)

        output_row.addWidget(output_label)
        output_row.addWidget(self.output_line, 1)
        output_row.addWidget(self.output_browse_button)
        option_layout.addLayout(output_row)

        main.addWidget(option_card)

        action_card = self._make_card()
        action_layout = QHBoxLayout(action_card)
        action_layout.setContentsMargins(12, 10, 12, 10)
        action_layout.setSpacing(8)

        self.run_button = QPushButton("비교 실행")
        self.run_button.setObjectName("primaryButton")
        self.run_button.clicked.connect(self._on_run_clicked)

        self.clear_log_button = QPushButton("로그 지우기")
        self.clear_log_button.clicked.connect(self._clear_log)

        action_layout.addWidget(self.run_button, 1)
        action_layout.addWidget(self.clear_log_button)
        main.addWidget(action_card)

        status_card = self._make_card()
        status_layout = QHBoxLayout(status_card)
        status_layout.setContentsMargins(12, 6, 12, 6)
        status_layout.setSpacing(8)
        status_card.setMinimumHeight(38)

        self.progress_label = QLabel("")
        self.progress_label.setObjectName("mutedText")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)

        self.status_label = QLabel("준비됨")
        self.status_label.setProperty("statusText", True)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)

        status_layout.addWidget(self.progress_label, 1)
        status_layout.addWidget(self.status_label, 0)

        main.addWidget(status_card)

        log_card = self._make_card()
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(12, 10, 12, 12)
        log_layout.setSpacing(6)

        log_title = QLabel("실행 로그 / 콘솔")
        log_title.setProperty("sectionTitle", True)
        log_layout.addWidget(log_title)

        self.log_view = TerminalConsole()
        log_layout.addWidget(self.log_view, 1)
        self.log_view.commandEntered.connect(self._on_console_command_entered)

        main.addWidget(log_card, 1)

        self._interactive_widgets.extend(
            [
                self.base_combo,
                self.target_combo,
                self.include_format_check,
                self.unprotect_check,
                self.output_browse_button,
                self.run_button,
                self.clear_log_button,
            ]
        )

        self._sync_summary_labels()

    def _make_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        card.setFrameShape(QFrame.Shape.StyledPanel)
        return card

    def _make_selected_summary_card(self, title: str) -> tuple[QLabel, QLabel]:
        card = self._make_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setObjectName("mutedText")

        file_name = QLabel("선택되지 않음")
        file_name.setObjectName("selectedName")
        file_name.setWordWrap(True)

        file_path = QLabel("")
        file_path.setObjectName("pathText")
        file_path.setWordWrap(True)

        layout.addWidget(title_label)
        layout.addWidget(file_name)
        layout.addWidget(file_path)

        if title == "기준 파일":
            self.base_summary_card = card
        else:
            self.target_summary_card = card

        return file_name, file_path

    def _make_picker_block(
        self,
        parent_layout: QGridLayout,
        column: int,
        title: str,
        on_browse,
        on_combo_change,
        on_drop_file,
        on_drag_state,
    ) -> tuple[QFrame, QComboBox]:
        block = QFrame()
        block.setObjectName("pickerCard")
        block.setProperty("dragActive", False)

        block_layout = QVBoxLayout(block)
        block_layout.setContentsMargins(12, 10, 12, 10)
        block_layout.setSpacing(8)

        title_label = QLabel(title)
        title_label.setProperty("sectionTitle", True)

        combo = DropComboBox()
        combo.setEditable(False)
        combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        combo.currentTextChanged.connect(on_combo_change)
        combo.fileDropped.connect(on_drop_file)
        combo.dragHoverChanged.connect(on_drag_state)

        browse_button = QPushButton("찾아보기")
        browse_button.clicked.connect(on_browse)

        block_layout.addWidget(title_label)
        block_layout.addWidget(combo)
        block_layout.addWidget(browse_button, 0, Qt.AlignmentFlag.AlignLeft)

        parent_layout.addWidget(block, 0, column)
        self._interactive_widgets.append(browse_button)
        return block, combo

    def _discover_mascot_paths(self) -> list[Path]:
        if not MASCOT_DIR.exists():
            return []
        mascots = [p for p in sorted(MASCOT_DIR.iterdir()) if p.is_file() and p.suffix.lower() in {".webp", ".png", ".jpg", ".jpeg"}]
        return mascots

    def _discover_chelsea_paths(self) -> list[Path]:
        if not CHELSEA_DIR.exists():
            return []

        candidates = [p for p in sorted(CHELSEA_DIR.iterdir()) if p.is_file() and p.suffix.lower() in {".webp", ".png", ".jpg", ".jpeg"}]
        ordered: list[Path] = []

        for keyword in ("drogba", "player", "face", "badge", "crest", "logo"):
            for path in candidates:
                if keyword in path.stem.lower() and path not in ordered:
                    ordered.append(path)
                    break

        for path in candidates:
            if path not in ordered:
                ordered.append(path)

        return ordered

    def _make_mascot_avatar(self, image_path: Path) -> QLabel:
        label = QLabel()
        label.setObjectName("mascotAvatar")
        label.setFixedSize(78, 78)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        pixmap = QPixmap(str(image_path))
        if not pixmap.isNull():
            pixmap = self._remove_flat_background(pixmap)
            scaled = pixmap.scaled(
                74,
                74,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            label.setPixmap(scaled)
        return label

    def _make_chelsea_avatar(self, image_path: Path) -> QLabel:
        label = QLabel()
        label.setObjectName("chelseaAvatar")
        label.setFixedSize(78, 78)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        pixmap = QPixmap(str(image_path))
        if not pixmap.isNull():
            scaled = pixmap.scaled(
                74,
                74,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            label.setPixmap(scaled)
        return label

    def _remove_flat_background(self, pixmap: QPixmap) -> QPixmap:
        image = pixmap.toImage().convertToFormat(QImage.Format.Format_ARGB32)
        width = image.width()
        height = image.height()
        if width <= 2 or height <= 2:
            return pixmap

        # 모서리 평균색을 배경으로 간주하고 유사 픽셀을 투명 처리
        sample_points = [
            (0, 0),
            (width - 1, 0),
            (0, height - 1),
            (width - 1, height - 1),
            (width // 2, 0),
            (0, height // 2),
            (width - 1, height // 2),
            (width // 2, height - 1),
        ]
        samples = [image.pixelColor(x, y) for x, y in sample_points]
        bg_r = int(sum(c.red() for c in samples) / len(samples))
        bg_g = int(sum(c.green() for c in samples) / len(samples))
        bg_b = int(sum(c.blue() for c in samples) / len(samples))

        threshold = 34
        threshold_sq = threshold * threshold

        for y in range(height):
            for x in range(width):
                color = image.pixelColor(x, y)
                if color.alpha() <= 10:
                    continue

                dr = color.red() - bg_r
                dg = color.green() - bg_g
                db = color.blue() - bg_b
                dist_sq = dr * dr + dg * dg + db * db
                if dist_sq <= threshold_sq:
                    color.setAlpha(0)
                    image.setPixelColor(x, y, color)

        return QPixmap.fromImage(image)

    def _apply_styles(self, *, keep_console_focus: bool = False) -> None:
        had_focus = self.log_view.hasFocus() if hasattr(self, "log_view") else False
        self.setStyleSheet(self._stylesheet_for_theme(self.current_theme))
        self._sync_theme_visibility()
        if keep_console_focus or had_focus:
            QTimer.singleShot(0, self.log_view.ensure_input_focus)
            QTimer.singleShot(60, self.log_view.ensure_input_focus)

    def _sync_theme_visibility(self) -> None:
        self.mascot_wrap.setVisible(self.current_theme == THEME_CUTE)
        self.chelsea_wrap.setVisible(self.current_theme == THEME_CHELSEA)

    def _toggle_theme_via_easter_egg(self) -> None:
        self.current_theme = THEME_CUTE if self.current_theme != THEME_CUTE else THEME_GREEN
        self._apply_styles(keep_console_focus=True)
        if self.current_theme == THEME_CUTE:
            self._append_log("이스터에그 테마 토글: 먼작귀 테마")
        else:
            self._append_log("이스터에그 테마 토글: 초록 테마")

    def _toggle_chelsea_theme_via_easter_egg(self) -> None:
        self.current_theme = THEME_CHELSEA if self.current_theme != THEME_CHELSEA else THEME_GREEN
        self._apply_styles(keep_console_focus=True)
        if self.current_theme == THEME_CHELSEA:
            self._append_log("이스터에그 테마 토글: 첼시 FC 테마")
        else:
            self._append_log("이스터에그 테마 토글: 초록 테마")

    def _on_console_command_entered(self, raw: str) -> None:
        try:
            raw_command = self._normalize_console_text(raw).strip("`'\".,!?")
            command = normalize_command_text(raw).strip("`'\".,!?")
            command_lower = command.lower()
            easter_egg = normalize_command_text(EASTER_EGG_COMMAND)
            chelsea_easter_egg = normalize_command_text(CHELSEA_EASTER_EGG_COMMAND)

            if not command:
                return

            if command == easter_egg:
                self._toggle_theme_via_easter_egg()
                return

            if command == chelsea_easter_egg or command_lower in {"chelsea", "cfc", "drogba"}:
                self._toggle_chelsea_theme_via_easter_egg()
                return

            if command_lower == "run":
                self._on_run_clicked()
                return

            if command_lower in {"clear", "cls"}:
                self._clear_log()
                return

            if command_lower == "reset":
                self._reset_inputs()
                return

            if command_lower in {"help", "?"}:
                self._append_log("명령어: run, clear, reset, find <키워드>, <셀주소>, <시트명>!<셀주소>")
                self._append_log("예시: find 차량명 / B12 / 견적서!B12")
                return

            if self._handle_search_command(raw_command):
                return

            if self._handle_cell_lookup_command(raw_command):
                return

            self._append_log(f"알 수 없는 명령: {raw}")
        except Exception as exc:  # noqa: BLE001
            self._append_log(f"명령 처리 오류: {exc}")
        finally:
            QTimer.singleShot(0, self.log_view.ensure_input_focus)

    def _stylesheet_for_theme(self, theme_key: str) -> str:
        if theme_key == THEME_GREEN:
            return """
            QMainWindow {
                background: #EEF2F7;
            }
            QWidget {
                font-family: "Apple SD Gothic Neo", "Noto Sans KR", "Malgun Gothic", sans-serif;
                font-size: 13px;
                color: #1F2937;
            }
            QFrame#card {
                background: #FFFFFF;
                border: 1px solid #D9E0EA;
                border-radius: 10px;
            }
            QFrame#headerCard {
                background: #FFFFFF;
                border: 1px solid #D9E0EA;
                border-radius: 10px;
            }
            QFrame#pickerCard {
                background: #F7F9FC;
                border: 1px solid #D7DEE9;
                border-radius: 10px;
            }
            QFrame#pickerCard[dragActive="true"] {
                background: #ECF9F2;
                border: 2px dashed #0F6B3E;
            }
            QLabel#titleText {
                font-size: 20px;
                font-weight: 700;
                color: #0F6B3E;
            }
            QLabel#selectedName {
                font-size: 15px;
                font-weight: 700;
                color: #1F2937;
            }
            QLabel#pathText {
                color: #6B7280;
                font-size: 12px;
            }
            QLabel#mutedText {
                color: #6B7280;
                font-size: 12px;
            }
            QLabel[sectionTitle="true"] {
                font-size: 13px;
                font-weight: 700;
                color: #1F2937;
            }
            QLabel[fieldLabel="true"] {
                color: #6B7280;
                font-weight: 700;
            }
            QLabel[statusText="true"] {
                font-weight: 700;
                color: #1F2937;
            }
            QComboBox, QLineEdit {
                border: 1px solid #D1D9E6;
                border-radius: 8px;
                padding: 7px 10px;
                background: #FFFFFF;
            }
            QComboBox::drop-down {
                border: none;
                width: 22px;
            }
            QComboBox:disabled, QLineEdit:disabled {
                background: #F3F6FA;
                color: #A0A7B3;
            }
            QPushButton {
                border: 1px solid #D1D9E6;
                border-radius: 8px;
                background: #E8EDF5;
                color: #1F2937;
                padding: 7px 12px;
                font-weight: 700;
            }
            QPushButton:hover {
                background: #DBE5F1;
            }
            QPushButton:pressed {
                background: #CCD9EA;
            }
            QPushButton:disabled {
                background: #EEF2F7;
                color: #9AA3AF;
            }
            QPushButton#primaryButton {
                background: #0F6B3E;
                color: #FFFFFF;
                border: none;
                font-size: 14px;
                padding: 10px 14px;
            }
            QPushButton#primaryButton:hover {
                background: #0D5C36;
            }
            QPushButton#primaryButton:disabled {
                background: #93BFA8;
                color: #F2F7F4;
            }
            QPlainTextEdit#logView {
                background: #17233A;
                color: #D9E7FF;
                border: 1px solid #243A5F;
                border-radius: 8px;
                padding: 6px;
                font-family: Menlo, Monaco, monospace;
                font-size: 12px;
            }
            QPlainTextEdit#logView:focus {
                border: 2px solid #0F6B3E;
            }
            """

        if theme_key == THEME_CHELSEA:
            return """
            QMainWindow {
                background: #E8EEF8;
            }
            QWidget {
                font-family: "Apple SD Gothic Neo", "Noto Sans KR", "Malgun Gothic", sans-serif;
                font-size: 13px;
                color: #10254D;
            }
            QFrame#card {
                background: #FFFFFF;
                border: 1px solid #C4D2EA;
                border-radius: 10px;
            }
            QFrame#headerCard {
                background: #034694;
                border: 1px solid #D2AB3F;
                border-radius: 10px;
            }
            QFrame#pickerCard {
                background: #F6F9FF;
                border: 1px solid #C7D4ED;
                border-radius: 10px;
            }
            QFrame#pickerCard[dragActive="true"] {
                background: #EBF2FF;
                border: 2px dashed #034694;
            }
            QLabel#titleText {
                font-size: 20px;
                font-weight: 700;
                color: #F8E5A6;
            }
            QLabel#selectedName {
                font-size: 15px;
                font-weight: 700;
                color: #132A56;
            }
            QLabel#pathText {
                color: #4E6288;
                font-size: 12px;
            }
            QLabel#mutedText {
                color: #5A6E95;
                font-size: 12px;
            }
            QLabel#chelseaAvatar {
                background: transparent;
                border: none;
                padding: 0px;
            }
            QLabel[sectionTitle="true"] {
                font-size: 13px;
                font-weight: 700;
                color: #1A2F5A;
            }
            QLabel[fieldLabel="true"] {
                color: #506288;
                font-weight: 700;
            }
            QLabel[statusText="true"] {
                font-weight: 700;
                color: #1A2F5A;
            }
            QComboBox, QLineEdit {
                border: 1px solid #BFCDE7;
                border-radius: 8px;
                padding: 7px 10px;
                background: #FFFFFF;
            }
            QComboBox::drop-down {
                border: none;
                width: 22px;
            }
            QComboBox:disabled, QLineEdit:disabled {
                background: #EEF3FD;
                color: #8EA1C6;
            }
            QPushButton {
                border: 1px solid #BFCDE7;
                border-radius: 8px;
                background: #E8EFFB;
                color: #16305D;
                padding: 7px 12px;
                font-weight: 700;
            }
            QPushButton:hover {
                background: #DDE8FA;
            }
            QPushButton:pressed {
                background: #D0DDF5;
            }
            QPushButton:disabled {
                background: #EFF4FD;
                color: #9AAACC;
            }
            QPushButton#primaryButton {
                background: #034694;
                color: #FFFFFF;
                border: 1px solid #D2AB3F;
                font-size: 14px;
                padding: 10px 14px;
            }
            QPushButton#primaryButton:hover {
                background: #003B7B;
            }
            QPushButton#primaryButton:disabled {
                background: #5F84BF;
                color: #E9F1FF;
            }
            QPlainTextEdit#logView {
                background: #0B1F46;
                color: #E6EEFF;
                border: 1px solid #31558E;
                border-radius: 8px;
                padding: 6px;
                font-family: Menlo, Monaco, monospace;
                font-size: 12px;
            }
            QPlainTextEdit#logView:focus {
                border: 2px solid #D2AB3F;
            }
            """

        return """
        QMainWindow {
            background: #F7F4FA;
        }
        QWidget {
            font-family: "Apple SD Gothic Neo", "Noto Sans KR", "Malgun Gothic", sans-serif;
            font-size: 13px;
            color: #2F3140;
        }
        QFrame#card {
            background: #FFFFFF;
            border: 1px solid #E7DCEC;
            border-radius: 14px;
        }
        QFrame#headerCard {
            background: #FFF9F3;
            border: 1px solid #F1D9C7;
            border-radius: 16px;
        }
        QFrame#pickerCard {
            background: #FFF9FD;
            border: 1px solid #F2DBE8;
            border-radius: 14px;
        }
        QFrame#pickerCard[dragActive="true"] {
            background: #FFEFF8;
            border: 2px dashed #E47BA3;
        }
        QLabel#titleText {
            font-size: 22px;
            font-weight: 700;
            color: #D96E9A;
        }
        QLabel#selectedName {
            font-size: 15px;
            font-weight: 700;
            color: #35364A;
        }
        QLabel#pathText {
            color: #86879A;
            font-size: 12px;
        }
        QLabel#mutedText {
            color: #7E7F94;
            font-size: 12px;
        }
        QLabel#mascotAvatar {
            background: transparent;
            border: none;
            border-radius: 16px;
            padding: 0px;
        }
        QLabel[sectionTitle="true"] {
            font-size: 13px;
            font-weight: 700;
            color: #4B4D68;
        }
        QLabel[fieldLabel="true"] {
            color: #7E7F94;
            font-weight: 700;
        }
        QLabel[statusText="true"] {
            font-weight: 700;
            color: #4B4D68;
        }
        QComboBox, QLineEdit {
            border: 1px solid #E4D8EA;
            border-radius: 8px;
            padding: 7px 10px;
            background: #FFFFFF;
        }
        QComboBox::drop-down {
            border: none;
            width: 22px;
        }
        QComboBox:disabled, QLineEdit:disabled {
            background: #F6F1F9;
            color: #A3A1B2;
        }
        QPushButton {
            border: 1px solid #E4D8EA;
            border-radius: 8px;
            background: #F3EBF7;
            color: #494B61;
            padding: 7px 12px;
            font-weight: 700;
        }
        QPushButton:hover {
            background: #EBDFF2;
        }
        QPushButton:pressed {
            background: #E3D3EC;
        }
        QPushButton:disabled {
            background: #F7F1FB;
            color: #A9A4B8;
        }
        QPushButton#primaryButton {
            background: #E985AE;
            color: #FFFFFF;
            border: none;
            font-size: 14px;
            padding: 10px 14px;
        }
        QPushButton#primaryButton:hover {
            background: #E072A1;
        }
        QPushButton#primaryButton:disabled {
            background: #F0B7CF;
            color: #FFF8FC;
        }
        QPlainTextEdit#logView {
            background: #2C2B45;
            color: #EEEAFD;
            border: 1px solid #45436A;
            border-radius: 8px;
            padding: 6px;
            font-family: Menlo, Monaco, monospace;
            font-size: 12px;
        }
        QPlainTextEdit#logView:focus {
            border: 2px solid #E072A1;
        }
        """

    def _append_log(self, message: str) -> None:
        self.log_view.append_log(f"[{now_text()}] {message}")

    def _clear_log(self) -> None:
        self.log_view.clear_logs()

    def _normalize_console_text(self, text: str) -> str:
        normalized = unicodedata.normalize("NFC", text or "")
        normalized = (
            normalized.replace("\u200b", "")
            .replace("\u200c", "")
            .replace("\u200d", "")
            .replace("\ufeff", "")
            .strip()
        )
        normalized = compose_compat_hangul(normalized)
        if normalized.startswith(">"):
            normalized = normalized[1:].strip()
        return normalized

    def _normalize_cell_address_input(self, text: str) -> str:
        candidate = text.strip().upper().replace("$", "")
        if not CELL_ADDRESS_RE.fullmatch(candidate):
            return ""
        return candidate

    def _normalize_lookup_key(self, text: str) -> str:
        normalized = unicodedata.normalize("NFC", text or "").casefold()
        normalized = compose_compat_hangul(normalized)
        normalized = re.sub(r"[\s!,:;|/\\]+", " ", normalized).strip()
        return normalized

    def _normalize_exact_match_key(self, text: str) -> str:
        normalized = unicodedata.normalize("NFC", text or "")
        normalized = compose_compat_hangul(normalized)
        normalized = normalized.replace("\r\n", "\n").replace("\r", "\n")
        normalized = re.sub(r"\s+", " ", normalized).strip().casefold()
        return normalized

    def _safe_console_value(self, value: str) -> str:
        text = (value or "").replace("\r\n", "\n").replace("\r", "\n").replace("\n", " ⏎ ").strip()
        if not text:
            return "(빈값)"
        if len(text) > 140:
            return text[:140] + "..."
        return text

    def _set_last_diff_rows(self, rows: Any) -> None:
        self.last_diff_rows = []
        self.last_diff_by_cell = {}
        self.last_diff_by_sheet_cell = {}

        if not isinstance(rows, list):
            return

        for row in rows:
            if not isinstance(row, (list, tuple)) or len(row) != 5:
                continue
            change_type, sheet_name, cell_addr, before_val, after_val = row
            sheet = str(sheet_name or "")
            cell = self._normalize_cell_address_input(str(cell_addr or ""))
            if not cell:
                continue
            record = (
                str(change_type or ""),
                sheet,
                cell,
                str(before_val or ""),
                str(after_val or ""),
            )
            self.last_diff_rows.append(record)
            self.last_diff_by_cell.setdefault(cell, []).append(record)
            sheet_key = f"{self._normalize_lookup_key(sheet)}!{cell}"
            self.last_diff_by_sheet_cell.setdefault(sheet_key, []).append(record)

    def _render_lookup_results(
        self,
        records: list[tuple[str, str, str, str, str]],
        title: str,
    ) -> None:
        if not records:
            self._append_log(f"{title}: 결과 없음")
            return

        change_label = {"modified": "수정", "added": "추가", "removed": "삭제"}
        total = len(records)
        self._append_log(f"{title}: {total}건")

        for change_type, sheet_name, cell_addr, before_val, after_val in records[:CELL_LOOKUP_LOG_LIMIT]:
            label = change_label.get(change_type, change_type)
            before_text = self._safe_console_value(before_val)
            after_text = self._safe_console_value(after_val)
            self._append_log(
                f"[{label}] {sheet_name}!{cell_addr} | 이전: {before_text} | 변경: {after_text}"
            )

        if total > CELL_LOOKUP_LOG_LIMIT:
            self._append_log(f"... {total - CELL_LOOKUP_LOG_LIMIT}건 더 있음")

    def _handle_search_command(self, raw_command: str) -> bool:
        parts = raw_command.split(maxsplit=1)
        if not parts:
            return False

        head = normalize_command_text(parts[0]).lower()
        if head not in {"find", "search", "검색"}:
            return False

        if len(parts) < 2 or not parts[1].strip():
            self._append_log("검색어를 입력해주세요. 예: find 할인")
            return True

        if not self.last_diff_rows:
            self._append_log("최근 실행 결과가 없습니다. 먼저 run으로 비교를 실행하세요.")
            return True

        keyword = parts[1].strip()
        needle = self._normalize_exact_match_key(keyword)
        matches: list[tuple[str, str, str, str, str]] = []
        change_label = {"modified": "수정", "added": "추가", "removed": "삭제"}

        for record in self.last_diff_rows:
            change_type, sheet_name, cell_addr, before_val, after_val = record
            candidates = {
                self._normalize_exact_match_key(change_type),
                self._normalize_exact_match_key(change_label.get(change_type, change_type)),
                self._normalize_exact_match_key(sheet_name),
                self._normalize_exact_match_key(cell_addr),
                self._normalize_exact_match_key(f"{sheet_name}!{cell_addr}"),
                self._normalize_exact_match_key(before_val),
                self._normalize_exact_match_key(after_val),
            }
            if needle in candidates:
                matches.append(record)

        self._render_lookup_results(matches, f"검색 '{keyword}'")
        return True

    def _handle_cell_lookup_command(self, raw_command: str) -> bool:
        if not raw_command:
            return False

        if "!" in raw_command:
            sheet_raw, cell_raw = raw_command.rsplit("!", 1)
            sheet_name = sheet_raw.strip()
            cell = self._normalize_cell_address_input(cell_raw)
            if not sheet_name or not cell:
                return False
            if not self.last_diff_rows:
                self._append_log("최근 실행 결과가 없습니다. 먼저 run으로 비교를 실행하세요.")
                return True
            sheet_key = f"{self._normalize_lookup_key(sheet_name)}!{cell}"
            records = self.last_diff_by_sheet_cell.get(sheet_key, [])
            self._render_lookup_results(records, f"조회 {sheet_name}!{cell}")
            return True

        cell = self._normalize_cell_address_input(raw_command)
        if not cell:
            return False
        if not self.last_diff_rows:
            self._append_log("최근 실행 결과가 없습니다. 먼저 run으로 비교를 실행하세요.")
            return True

        records = self.last_diff_by_cell.get(cell, [])
        self._render_lookup_results(records, f"조회 {cell}")
        return True

    def _reset_inputs(self) -> None:
        if self.running:
            self._append_log("실행 중에는 reset 명령을 사용할 수 없습니다.")
            return

        self.last_diff_rows = []
        self.last_diff_by_cell = {}
        self.last_diff_by_sheet_cell = {}
        self.base_file = None
        self.target_file = None
        self.output_file = None
        self.output_file_manually_set = False
        self.base_choice_map = {}
        self.target_choice_map = {}
        self._clear_picker_drag_highlight()
        self._initialize_dropdown_sources()
        self._sync_summary_labels()
        self._append_log("입력 파일/결과 경로를 초기화했습니다.")

    def _scan_excel_files(self, folder: Path) -> list[Path]:
        files: list[Path] = []
        try:
            for item in folder.iterdir():
                if not item.is_file():
                    continue
                if item.name.startswith("~$"):
                    continue
                if item.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                files.append(item)
        except OSError:
            return []

        files.sort(key=lambda p: (p.name.lower(), str(p).lower()))
        return files

    def _load_folder_files(self, folder: Path, slot: str, *, log: bool = True) -> None:
        files = self._scan_excel_files(folder)
        if slot == "base":
            self.base_folder_files = files
            if log:
                self._append_log(f"기준 드롭다운 목록 {len(files)}개 로드: {folder}")
        else:
            self.target_folder_files = files
            if log:
                self._append_log(f"비교 드롭다운 목록 {len(files)}개 로드: {folder}")

        self._refresh_dropdown_values()

    def _initialize_dropdown_sources(self) -> None:
        default_folder = Path.cwd()
        self._load_folder_files(default_folder, "base", log=False)
        self._load_folder_files(default_folder, "target", log=False)

    def _label_for_path(self, path: Path) -> str:
        return path.name

    def _merge_file_choices(self, slot_files: list[Path]) -> list[Path]:
        merged: list[Path] = []
        seen: set[Path] = set()

        for path in slot_files + self.recent_files:
            if path in seen:
                continue
            seen.add(path)
            merged.append(path)
        return merged

    def _build_choice_map(self, paths: list[Path]) -> tuple[list[str], dict[str, Path]]:
        labels: list[str] = []
        mapping: dict[str, Path] = {}

        for path in paths:
            label = self._label_for_path(path)
            if label in mapping:
                label = f"{path.name} ({path.parent})"
                suffix = 2
                candidate = f"{label} #{suffix}"
                while candidate in mapping:
                    suffix += 1
                    candidate = f"{label} #{suffix}"
                label = candidate
            labels.append(label)
            mapping[label] = path

        return labels, mapping

    def _find_label_by_path(self, mapping: dict[str, Path], target_path: Path) -> str:
        for label, mapped_path in mapping.items():
            if mapped_path == target_path:
                return label
        return ""

    def _refresh_dropdown_values(self) -> None:
        base_values, self.base_choice_map = self._build_choice_map(self._merge_file_choices(self.base_folder_files))
        target_values, self.target_choice_map = self._build_choice_map(
            self._merge_file_choices(self.target_folder_files)
        )

        self.base_combo.blockSignals(True)
        self.target_combo.blockSignals(True)
        self.base_combo.clear()
        self.target_combo.clear()
        self.base_combo.addItems(base_values)
        self.target_combo.addItems(target_values)

        if self.base_file is not None:
            label = self._find_label_by_path(self.base_choice_map, self.base_file)
            if label:
                self.base_combo.setCurrentText(label)
            else:
                self.base_combo.setCurrentIndex(-1)
        else:
            self.base_combo.setCurrentIndex(-1)

        if self.target_file is not None:
            label = self._find_label_by_path(self.target_choice_map, self.target_file)
            if label:
                self.target_combo.setCurrentText(label)
            else:
                self.target_combo.setCurrentIndex(-1)
        else:
            self.target_combo.setCurrentIndex(-1)

        self.base_combo.blockSignals(False)
        self.target_combo.blockSignals(False)

    def _push_recent(self, path: Path) -> None:
        try:
            resolved = path.resolve()
        except OSError:
            resolved = path

        self.recent_files = [p for p in self.recent_files if p != resolved]
        self.recent_files.insert(0, resolved)
        self.recent_files = self.recent_files[:30]

    def _sync_summary_labels(self) -> None:
        if self.base_file is None:
            self.base_name_label.setText("선택되지 않음")
            self.base_path_label.setText("")
        else:
            self.base_name_label.setText(self.base_file.name)
            self.base_path_label.setText(str(self.base_file))

        if self.target_file is None:
            self.target_name_label.setText("선택되지 않음")
            self.target_path_label.setText("")
        else:
            self.target_name_label.setText(self.target_file.name)
            self.target_path_label.setText(str(self.target_file))

        if self.output_file is None:
            self.output_line.setText("")
        else:
            self.output_line.setText(str(self.output_file))

    def _set_base_file(self, path: Path, reason: str = "기준 파일 선택") -> None:
        self.base_file = path
        self._push_recent(path)
        self._refresh_dropdown_values()
        self._sync_summary_labels()
        self._append_log(f"{reason}: {path.name}")

    def _set_target_file(self, path: Path, reason: str = "비교 파일 선택") -> None:
        self.target_file = path
        self._push_recent(path)

        if self.output_file is None or not self.output_file_manually_set:
            self.output_file = self._make_default_output(path)

        self._refresh_dropdown_values()
        self._sync_summary_labels()
        self._append_log(f"{reason}: {path.name}")

    def _apply_file_to_slot(self, slot: str, path: Path, reason: str) -> None:
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            return

        self._load_folder_files(path.parent, slot, log=False)
        if slot == "base":
            self._set_base_file(path, reason=reason)
        else:
            self._set_target_file(path, reason=reason)

    def _set_picker_drag_highlight(self, slot: str, active: bool) -> None:
        block = self.base_picker_block if slot == "base" else self.target_picker_block
        block.setProperty("dragActive", active)
        style = block.style()
        if style is not None:
            style.unpolish(block)
            style.polish(block)
        block.update()

    def _clear_picker_drag_highlight(self) -> None:
        self._set_picker_drag_highlight("base", False)
        self._set_picker_drag_highlight("target", False)

    def _slot_from_drop_position(self, event: Any) -> str | None:
        if not hasattr(event, "position"):
            return None
        point = event.position().toPoint()
        global_point = self.mapToGlobal(point)

        if self.base_picker_block.rect().contains(self.base_picker_block.mapFromGlobal(global_point)):
            return "base"
        if self.target_picker_block.rect().contains(self.target_picker_block.mapFromGlobal(global_point)):
            return "target"
        return None

    def _preferred_drop_slot(self) -> str:
        if self.base_file is None:
            return "base"
        return "target"

    def _on_base_combo_changed(self, text: str) -> None:
        selected = self.base_choice_map.get(text)
        if selected is None:
            return
        if self.base_file == selected:
            return
        self._set_base_file(selected, reason="드롭다운 기준 파일 선택")

    def _on_target_combo_changed(self, text: str) -> None:
        selected = self.target_choice_map.get(text)
        if selected is None:
            return
        if self.target_file == selected:
            return
        self._set_target_file(selected, reason="드롭다운 비교 파일 선택")

    def _on_base_combo_file_dropped(self, file_path_text: str) -> None:
        path = Path(file_path_text).expanduser()
        if not path.exists():
            return
        self._apply_file_to_slot("base", path, reason="드래그앤드롭 기준 파일 적용")
        self._clear_picker_drag_highlight()

    def _on_target_combo_file_dropped(self, file_path_text: str) -> None:
        path = Path(file_path_text).expanduser()
        if not path.exists():
            return
        self._apply_file_to_slot("target", path, reason="드래그앤드롭 비교 파일 적용")
        self._clear_picker_drag_highlight()

    def _on_base_combo_drag_state(self, active: bool) -> None:
        self._set_picker_drag_highlight("base", active)
        if active:
            self._set_picker_drag_highlight("target", False)

    def _on_target_combo_drag_state(self, active: bool) -> None:
        self._set_picker_drag_highlight("target", active)
        if active:
            self._set_picker_drag_highlight("base", False)

    def _pick_excel_file(self, title: str, initial_dir: Path | None = None) -> Path | None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            title,
            str(initial_dir or Path.cwd()),
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not file_path:
            return None

        candidate = Path(file_path).expanduser()
        if candidate.suffix.lower() not in SUPPORTED_EXTENSIONS:
            QMessageBox.critical(self, "파일 형식 오류", "xlsx 또는 xlsm 파일만 선택할 수 있습니다.")
            return None

        return candidate

    def _pick_base_file(self) -> None:
        initial = self.base_file.parent if self.base_file else Path.cwd()
        candidate = self._pick_excel_file("기준 파일 선택", initial)
        if candidate is None:
            return

        self._apply_file_to_slot("base", candidate, reason="기준 파일 선택")

    def _pick_target_file(self) -> None:
        initial = self.target_file.parent if self.target_file else Path.cwd()
        candidate = self._pick_excel_file("비교 파일 선택", initial)
        if candidate is None:
            return

        self._apply_file_to_slot("target", candidate, reason="비교 파일 선택")

    def _pick_output_file(self) -> None:
        if self.output_file is not None:
            initial = self.output_file
        elif self.target_file is not None:
            initial = self._make_default_output(self.target_file)
        else:
            initial = Path.cwd() / "result_diff_highlight.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "결과 파일 저장 위치",
            str(initial),
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not file_path:
            return

        candidate = Path(file_path).expanduser()
        if candidate.suffix.lower() not in SUPPORTED_EXTENSIONS:
            QMessageBox.critical(self, "파일 형식 오류", "결과 파일은 xlsx 또는 xlsm 이어야 합니다.")
            return

        self.output_file = candidate
        self.output_file_manually_set = True
        self._sync_summary_labels()

    def _make_default_output(self, target_file: Path) -> Path:
        return target_file.with_name(f"{target_file.stem}_diff_highlight{target_file.suffix}")

    def _validate_input(self) -> tuple[Path, Path, Path]:
        if self.base_file is None:
            raise ValueError("기준 파일을 선택해주세요.")
        if self.target_file is None:
            raise ValueError("비교 파일을 선택해주세요.")

        if self.output_file is None:
            self.output_file = self._make_default_output(self.target_file)

        if not self.base_file.exists():
            raise ValueError("기준 파일 경로를 찾을 수 없습니다.")
        if not self.target_file.exists():
            raise ValueError("비교 파일 경로를 찾을 수 없습니다.")
        if self.base_file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError("기준 파일은 xlsx/xlsm 파일이어야 합니다.")
        if self.target_file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError("비교 파일은 xlsx/xlsm 파일이어야 합니다.")
        if self.output_file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError("결과 파일은 xlsx/xlsm 파일이어야 합니다.")

        self._sync_summary_labels()
        return self.base_file, self.target_file, self.output_file

    def _set_running(self, running: bool) -> None:
        self.running = running
        for widget in self._interactive_widgets:
            widget.setEnabled(not running)

        if running:
            self.status_label.setText("실행 중")
            self.elapsed_timer.start()
            self.progress_timer.start()
            self._update_progress_text()
        else:
            self.status_label.setText("준비됨")
            self.progress_timer.stop()
            self.progress_label.setText("")

    def _update_progress_text(self) -> None:
        if not self.running:
            self.progress_label.setText("")
            return
        elapsed = int(self.elapsed_timer.elapsed() / 1000)
        minutes = elapsed // 60
        seconds = elapsed % 60
        self.progress_label.setText(f"처리 중 · {minutes:02d}:{seconds:02d}")

    def _on_run_clicked(self) -> None:
        if self.running:
            self._append_log("이미 비교 작업이 실행 중입니다.")
            return

        try:
            base_file, target_file, output_file = self._validate_input()
        except ValueError as exc:
            QMessageBox.critical(self, "입력 오류", str(exc))
            return

        self._set_running(True)
        self._append_log("비교 시작")
        self._append_log(f"- 기준 파일: {base_file}")
        self._append_log(f"- 비교 파일: {target_file}")
        self._append_log(f"- 결과 파일: {output_file}")

        self.worker_thread = QThread(self)
        self.worker = CompareWorker(
            base_file=base_file,
            target_file=target_file,
            output_file=output_file,
            include_format=self.include_format_check.isChecked(),
            unprotect_second=self.unprotect_check.isChecked(),
        )
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.failed.connect(self._on_worker_failed)

        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)

        self.worker_thread.finished.connect(self._on_worker_thread_finished)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)

        self.worker_thread.start()

    def _on_worker_finished(self, payload: dict[str, Any]) -> None:
        self._set_last_diff_rows(payload.get("diff_rows", []))
        self._append_log("하이라이트 파일 생성 완료")
        self._append_log(
            "- 수정 셀: {modified_cells}, 추가 셀: {added_cells}, 삭제 셀: {removed_cells}".format(
                modified_cells=payload.get("modified_cells", 0),
                added_cells=payload.get("added_cells", 0),
                removed_cells=payload.get("removed_cells", 0),
            )
        )
        self._append_log(f"- 셀 조회 인덱스 준비: {len(self.last_diff_rows)}건")

        changed_sheets = payload.get("changed_sheets", [])
        if changed_sheets:
            self._append_log(f"- 변경 시트: {', '.join(changed_sheets)}")
        else:
            self._append_log("- 변경 시트: 없음")

        output_file = payload.get("output_file", "")
        QMessageBox.information(
            self,
            "완료",
            "비교가 완료되었습니다.\n\n"
            f"결과 파일:\n{output_file}\n\n"
            "색상 안내: 수정(노랑), 추가(초록), 삭제(주황)\n"
            "시트 탭 안내: 변경 시트(노랑), 신규 시트(초록)",
        )

    def _on_worker_failed(self, detail: str) -> None:
        first_line = detail.splitlines()[0] if detail else "알 수 없는 오류"
        self._append_log("오류 발생")
        self._append_log(first_line)
        self._append_log(detail)

        QMessageBox.critical(
            self,
            "실행 실패",
            f"오류가 발생했습니다.\n\n{first_line}\n\n자세한 내용은 실행 로그를 확인해주세요.",
        )

    def _on_worker_thread_finished(self) -> None:
        self.worker = None
        self.worker_thread = None
        self._set_running(False)
        gc.collect()

    def dragEnterEvent(self, event: Any) -> None:
        if self.running:
            self._clear_picker_drag_highlight()
            event.ignore()
            return
        if extract_excel_paths_from_mime_data(event.mimeData()):
            target_slot = self._slot_from_drop_position(event) or self._preferred_drop_slot()
            self._set_picker_drag_highlight("base", target_slot == "base")
            self._set_picker_drag_highlight("target", target_slot == "target")
            event.acceptProposedAction()
        else:
            self._clear_picker_drag_highlight()
            event.ignore()

    def dragMoveEvent(self, event: Any) -> None:
        if self.running:
            self._clear_picker_drag_highlight()
            event.ignore()
            return
        if extract_excel_paths_from_mime_data(event.mimeData()):
            target_slot = self._slot_from_drop_position(event) or self._preferred_drop_slot()
            self._set_picker_drag_highlight("base", target_slot == "base")
            self._set_picker_drag_highlight("target", target_slot == "target")
            event.acceptProposedAction()
        else:
            self._clear_picker_drag_highlight()
            event.ignore()

    def dropEvent(self, event: Any) -> None:
        if self.running:
            self._clear_picker_drag_highlight()
            event.ignore()
            return

        paths = extract_excel_paths_from_mime_data(event.mimeData())
        if not paths:
            self._clear_picker_drag_highlight()
            event.ignore()
            return

        preferred = self._slot_from_drop_position(event) or self._preferred_drop_slot()
        if len(paths) >= 2:
            self._apply_file_to_slot("base", paths[0], reason="드래그앤드롭 기준 파일 적용")
            self._apply_file_to_slot("target", paths[1], reason="드래그앤드롭 비교 파일 적용")
            self._append_log(f"윈도우 드롭으로 파일 2개 적용: {paths[0].name}, {paths[1].name}")
        else:
            dropped = paths[0]
            if preferred == "base":
                self._apply_file_to_slot("base", dropped, reason="드래그앤드롭 기준 파일 적용")
            else:
                self._apply_file_to_slot("target", dropped, reason="드래그앤드롭 비교 파일 적용")
            self._append_log(f"윈도우 드롭으로 파일 적용: {dropped.name}")

        self._clear_picker_drag_highlight()
        event.acceptProposedAction()

    def dragLeaveEvent(self, event: Any) -> None:
        self._clear_picker_drag_highlight()
        event.accept()

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.running:
            QMessageBox.warning(
                self,
                "실행 중",
                "현재 비교 작업이 실행 중입니다. 완료 후 창을 닫아주세요.",
            )
            event.ignore()
            return

        self.progress_timer.stop()
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    window = ExcelDiffMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
