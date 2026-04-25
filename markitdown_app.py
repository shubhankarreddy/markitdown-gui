"""
MarkItDown - polished PyQt6 desktop interface for Microsoft's markitdown library.
"""
import os
import re
import sys
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional
from uuid import uuid4

from PyQt6.QtCore import QEvent, QSettings, QSize, QThread, QTimer, QUrl, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QDesktopServices, QFont, QIcon, QKeySequence, QPalette, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QCheckBox,
    QComboBox,
    QProgressBar,
    QScrollArea,
    QSplitter,
    QStyleFactory,
    QTabWidget,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


def exception_hook(exc_type, exc_value, exc_tb):
    lines = traceback.format_exception(exc_type, exc_value, exc_tb)
    msg = "".join(lines)
    print(msg, file=sys.stderr)
    try:
        QMessageBox.critical(None, "MarkItDown - Error", msg)
    except Exception:
        pass


sys.excepthook = exception_hook


DARK = {
    "bg": "#0a1118",
    "surface": "#0f1821",
    "surface2": "#14202b",
    "card": "#101b25",
    "card2": "#162634",
    "border": "#223445",
    "border2": "#365066",
    "accent": "#44cdee",
    "accent2": "#16aeca",
    "accent_soft": "#123544",
    "text": "#eff7ff",
    "muted": "#98aec2",
    "success": "#4fd79a",
    "warning": "#f3c56f",
    "error": "#ff817d",
    "btn_text": "#041016",
}
LIGHT = {
    "bg": "#eef4f8",
    "surface": "#ffffff",
    "surface2": "#f3f8fc",
    "card": "#ffffff",
    "card2": "#edf7fb",
    "border": "#d4e0ea",
    "border2": "#b7cada",
    "accent": "#0ea6bf",
    "accent2": "#0a8ca3",
    "accent_soft": "#d9f2f7",
    "text": "#10202d",
    "muted": "#64798b",
    "success": "#119b6c",
    "warning": "#b77707",
    "error": "#c84545",
    "btn_text": "#ffffff",
}

EXTS = (
    "pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls",
    "html", "htm", "xml", "txt", "csv", "json", "yaml", "yml",
    "png", "jpg", "jpeg", "gif", "bmp",
    "mp3", "wav", "mp4", "ipynb", "zip", "md",
)
FILTER_STR = "Supported files (%s)" % " ".join("*.%s" % e for e in EXTS)
BADGES = {
    "queued": "[Q]",
    "converting": "[~]",
    "done": "[OK]",
    "error": "[!]",
}
STATUS_LABELS = {
    "queued": "Queued",
    "converting": "Converting",
    "done": "Done",
    "error": "Needs attention",
}


def resource_path(*parts):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*parts)


def find_icon_path():
    candidates = [
        resource_path("assets", "markitdown.ico"),
        Path(__file__).resolve().parent / "assets" / "markitdown.ico",
        Path(__file__).resolve().parent / "Native" / "MarkItDown.Native" / "Assets" / "MarkItDown.ico",
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return ""


def default_output_folder():
    docs = Path.home() / "Documents"
    base = docs if docs.exists() else Path.home()
    return str(base / "MarkItDown Exports")


def detect_dark():
    try:
        s = QApplication.instance().styleHints().colorScheme()
        if hasattr(Qt, "ColorScheme"):
            return s == Qt.ColorScheme.Dark
    except Exception:
        pass
    c = QApplication.instance().palette().color(QPalette.ColorRole.Window)
    return (0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue()) < 128


def build_palette(theme):
    palette = QPalette()
    mapping = [
        (QPalette.ColorRole.Window, "bg"),
        (QPalette.ColorRole.WindowText, "text"),
        (QPalette.ColorRole.Base, "surface"),
        (QPalette.ColorRole.AlternateBase, "surface2"),
        (QPalette.ColorRole.Text, "text"),
        (QPalette.ColorRole.Button, "surface2"),
        (QPalette.ColorRole.ButtonText, "text"),
        (QPalette.ColorRole.Highlight, "accent"),
        (QPalette.ColorRole.HighlightedText, "btn_text"),
        (QPalette.ColorRole.ToolTipBase, "card2"),
        (QPalette.ColorRole.ToolTipText, "text"),
        (QPalette.ColorRole.PlaceholderText, "muted"),
        (QPalette.ColorRole.Link, "accent"),
    ]
    for role, key in mapping:
        palette.setColor(role, QColor(theme[key]))
    return palette


def build_ui_font():
    if sys.platform == "win32":
        for family in ("Segoe UI Variable Text", "Aptos", "Segoe UI"):
            font = QFont(family, 10)
            if family == "Segoe UI" or font.exactMatch():
                return font
    return QFont("Sans Serif", 10)


def build_editor_font():
    for family in ("Cascadia Mono", "Consolas", "JetBrains Mono", "Courier New"):
        font = QFont(family)
        font.setStyleHint(QFont.StyleHint.Monospace)
        if font.exactMatch():
            font.setPointSize(10)
            return font
    font = QFont("Courier New")
    font.setStyleHint(QFont.StyleHint.Monospace)
    font.setPointSize(10)
    return font


def shorten_middle(text, limit=66):
    if len(text) <= limit:
        return text
    half = max(8, (limit - 3) // 2)
    return text[:half] + "..." + text[-half:]


def sanitize_stem(text):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", text).strip("-._")
    return cleaned or "converted-file"


def normalize_markdown_output(text):
    """Clean common PDF extraction artifacts so the saved file is real Markdown."""
    if not text:
        return ""

    text = text.replace("\u00a0", " ")
    bullet_chars = "\u2022\u25cf\u25aa\u25ab\u25e6\u2023\u2043\u2219\u00b7\uf0b7\uf0a7\uf0d8\uf0fc"
    bullet_re = re.compile(r"^(\s*)[" + re.escape(bullet_chars) + r"]\s*(\S.*)$")
    normalized = []
    for line in text.splitlines():
        match = bullet_re.match(line)
        if match:
            normalized.append("%s- %s" % (match.group(1), match.group(2)))
        else:
            normalized.append(line.rstrip())

    return "\n".join(normalized).strip() + "\n"


def format_dt(dt):
    if not dt:
        return "Not yet"
    return dt.strftime("%d %b %Y, %I:%M %p")


def unique_output_path(folder, entry_name):
    folder_path = Path(folder)
    folder_path.mkdir(parents=True, exist_ok=True)
    stem = sanitize_stem(Path(entry_name).stem)
    out = folder_path / (stem + ".md")
    counter = 1
    while out.exists():
        out = folder_path / ("%s_%d.md" % (stem, counter))
        counter += 1
    return str(out)


def collect_supported_files(folder_path, errors=None):
    found = []
    root = Path(folder_path)
    if not root.exists():
        return found
    try:
        for current_root, _, files in os.walk(root, onerror=lambda exc: errors.append(str(exc)) if errors is not None else None):
            for name in sorted(files):
                path = Path(current_root) / name
                if path.suffix.lstrip(".").lower() in EXTS:
                    found.append(str(path))
    except Exception as exc:
        if errors is not None:
            errors.append(str(exc))
    return sorted(found)


@dataclass
class QueueEntry:
    source: str
    is_url: bool
    name: str
    entry_id: str = field(default_factory=lambda: uuid4().hex)
    status: str = "queued"
    result: str = ""
    error: str = ""
    saved_path: str = ""
    added_at: datetime = field(default_factory=datetime.now)
    completed_at: Optional[datetime] = None

    def key(self):
        if self.is_url:
            return self.source.strip().lower()
        try:
            return str(Path(self.source).resolve()).lower()
        except Exception:
            return os.path.abspath(self.source).lower()

    def status_label(self):
        return STATUS_LABELS.get(self.status, "Queued")

    def source_label(self):
        return shorten_middle(self.source, 78)


class ConvertWorker(QThread):
    item_started = pyqtSignal(int, str)
    item_done = pyqtSignal(int, str, str)
    all_done = pyqtSignal()

    def __init__(self, items, llm_cfg=None):
        super().__init__()
        self.items = items
        self.llm_cfg = llm_cfg

    def run(self):
        try:
            from markitdown import MarkItDown
        except Exception as exc:
            for i in range(len(self.items)):
                self.item_done.emit(i, "", "Failed to load markitdown: %s" % str(exc))
            self.all_done.emit()
            return

        kwargs = {}
        if self.llm_cfg:
            try:
                from openai import OpenAI
                kwargs["llm_client"] = OpenAI(api_key=self.llm_cfg["api_key"])
                kwargs["llm_model"] = self.llm_cfg["model"]
            except Exception as exc:
                for i in range(len(self.items)):
                    self.item_done.emit(i, "", "OpenAI init failed: %s" % str(exc))
                self.all_done.emit()
                return

        try:
            converter = MarkItDown(**kwargs)
        except Exception as exc:
            for i in range(len(self.items)):
                self.item_done.emit(i, "", "MarkItDown init failed: %s" % str(exc))
            self.all_done.emit()
            return

        for i, source in enumerate(self.items):
            self.item_started.emit(i, source)
            try:
                result = converter.convert(source)
                self.item_done.emit(i, normalize_markdown_output(result.text_content or ""), "")
            except Exception as exc:
                self.item_done.emit(i, "", str(exc))

        self.all_done.emit()


class FolderScanWorker(QThread):
    scan_done = pyqtSignal(list, list, list)

    def __init__(self, folders):
        super().__init__()
        self.folders = folders

    def run(self):
        found = []
        errors = []
        seen = set()
        for folder in self.folders:
            for path in collect_supported_files(folder, errors):
                norm = str(Path(path))
                if norm not in seen:
                    seen.add(norm)
                    found.append(norm)
        self.scan_done.emit(self.folders, found, errors)


class DropZone(QLabel):
    files_dropped = pyqtSignal(list)
    clicked = pyqtSignal()

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self.theme = theme
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setWordWrap(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setText("Drop files or folders here")
        self.setMinimumHeight(64)
        self.setMaximumHeight(76)
        self.apply_style(False)

    def apply_style(self, hover):
        theme = self.theme
        border_color = theme["accent"] if hover else theme["border2"]
        background = theme["accent_soft"] if hover else theme["card2"]
        self.setStyleSheet(
            "QLabel {"
            " border: 2px dashed %s;"
            " border-radius: 8px;"
            " background: %s;"
            " color: %s;"
            " padding: 10px;"
            " font-size: 13px;"
            " font-weight: 600;"
            "}" % (border_color, background, theme["text"])
        )

    def update_theme(self, theme):
        self.theme = theme
        self.apply_style(False)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.apply_style(True)

    def dragLeaveEvent(self, event):
        self.apply_style(False)

    def dropEvent(self, event):
        self.apply_style(False)
        paths = []
        for url in event.mimeData().urls():
            fp = url.toLocalFile()
            if fp and (os.path.isfile(fp) or os.path.isdir(fp)):
                paths.append(fp)
        if paths:
            self.files_dropped.emit(paths)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setObjectName("mainWindow")
        self.setWindowTitle("MarkItDown")
        self.setMinimumSize(1080, 680)
        self.resize(1320, 820)

        self.settings = QSettings("MarkItDown", "MarkItDown")
        self.queue = []
        self.worker = None
        self.scan_worker = None
        self.output_folder = ""
        self.last_browse_dir = str(Path.home())
        self._run_entry_ids = []
        self._prog = 0
        self._active_source = ""
        self._active_scan_folders = []
        self._queued_scan_folders = []
        self._block_show = False

        # Keep the desktop app dark-first for long document work and low-glare batch sessions.
        self._dark = True
        self.theme = DARK
        self.icon_path = find_icon_path()

        self.build_ui()
        self.apply_theme()
        self.wire_signals()
        self.setup_shortcuts()
        self.restore_settings()
        self.update_summary()
        self.update_preview(-1)
        self.show_message("Ready. Add files, folders, or URLs to build a local Markdown conversion queue.")

        try:
            QApplication.instance().styleHints().colorSchemeChanged.connect(
                self.on_system_theme_changed
            )
        except Exception:
            pass

    def build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        self.statusBar().setSizeGripEnabled(False)

        main_lay = QHBoxLayout(central)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.setSpacing(0)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        main_lay.addWidget(self.splitter)

        left_panel = QWidget()
        left_panel.setObjectName("leftRail")
        left_panel.setMinimumWidth(440)
        left_panel.setMaximumWidth(520)
        left_outer = QVBoxLayout(left_panel)
        left_outer.setContentsMargins(0, 0, 0, 0)
        left_outer.setSpacing(0)

        scroll = QScrollArea()
        self.left_scroll = scroll
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.left_scroll_viewport = scroll.viewport()
        self.left_scroll_viewport.installEventFilter(self)

        self.left_widget = QWidget()
        lv = QVBoxLayout(self.left_widget)
        lv.setContentsMargins(16, 16, 16, 10)
        lv.setSpacing(10)

        hero = QFrame()
        hero.setObjectName("heroCard")
        hero_lay = QVBoxLayout(hero)
        hero_lay.setContentsMargins(14, 14, 14, 14)
        hero_lay.setSpacing(6)

        eyebrow = QLabel("DOCUMENT TO MARKDOWN")
        eyebrow.setObjectName("eyebrowLbl")
        hero_lay.addWidget(eyebrow)

        title = QLabel("MarkItDown")
        title.setObjectName("titleLbl")
        hero_lay.addWidget(title)

        subtitle = QLabel(
            "Local batch conversion for clean, reusable Markdown."
        )
        subtitle.setObjectName("subLbl")
        subtitle.setWordWrap(True)
        hero_lay.addWidget(subtitle)

        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(6)
        total_card, self.total_metric_lbl = self.make_metric_card("Ready")
        done_card, self.done_metric_lbl = self.make_metric_card("Saved")
        issue_card, self.issue_metric_lbl = self.make_metric_card("Issues")
        metrics_row.addWidget(total_card)
        metrics_row.addWidget(done_card)
        metrics_row.addWidget(issue_card)
        hero_lay.addLayout(metrics_row)
        lv.addWidget(hero)

        source_card = QFrame()
        source_card.setObjectName("sourceCard")
        source_lay = QVBoxLayout(source_card)
        source_lay.setContentsMargins(12, 12, 12, 12)
        source_lay.setSpacing(8)

        source_title = QLabel("Collect sources")
        source_title.setObjectName("sectionLbl")
        source_lay.addWidget(source_title)

        source_sub = QLabel("Add files, folders, or a URL.")
        source_sub.setObjectName("mutedLbl")
        source_sub.setWordWrap(True)
        source_lay.addWidget(source_sub)

        self.drop_zone = DropZone(self.theme)
        source_lay.addWidget(self.drop_zone)

        add_row = QHBoxLayout()
        add_row.setSpacing(8)
        self.browse_btn = QPushButton("Add Files")
        self.browse_btn.setObjectName("secondaryBtn")
        self.browse_btn.setMinimumWidth(0)
        add_row.addWidget(self.browse_btn, 1)

        self.add_folder_btn = QPushButton("Add Folder")
        self.add_folder_btn.setObjectName("ghostBtn")
        self.add_folder_btn.setMinimumWidth(0)
        add_row.addWidget(self.add_folder_btn, 1)
        source_lay.addLayout(add_row)

        web_row = QHBoxLayout()
        web_row.setSpacing(8)
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("Paste URL")
        self.url_input.setMinimumWidth(0)
        web_row.addWidget(self.url_input, 1)
        self.url_btn = QPushButton("Add URL")
        self.url_btn.setObjectName("ghostBtn")
        self.url_btn.setFixedWidth(82)
        web_row.addWidget(self.url_btn)
        source_lay.addLayout(web_row)

        lv.addWidget(source_card)

        output_card = QFrame()
        output_card.setObjectName("sectionCard")
        out_lay = QVBoxLayout(output_card)
        out_lay.setContentsMargins(12, 12, 12, 12)
        out_lay.setSpacing(8)

        out_header = QHBoxLayout()
        out_title = QLabel("Output folder")
        out_title.setObjectName("sectionLbl")
        out_header.addWidget(out_title)
        out_header.addStretch()
        self.open_output_btn = QPushButton("Open")
        self.open_output_btn.setObjectName("ghostBtn")
        self.open_output_btn.setFixedWidth(58)
        out_header.addWidget(self.open_output_btn)
        out_lay.addLayout(out_header)

        self.folder_btn = QPushButton("Change Output Folder")
        self.folder_btn.setObjectName("secondaryBtn")
        self.folder_btn.setMinimumHeight(34)
        out_lay.addWidget(self.folder_btn)

        self.folder_lbl = QLabel("")
        self.folder_lbl.setObjectName("mutedLbl")
        self.folder_lbl.setWordWrap(True)
        out_lay.addWidget(self.folder_lbl)
        lv.addWidget(output_card)

        queue_card = QFrame()
        queue_card.setObjectName("queueCard")
        q_lay = QVBoxLayout(queue_card)
        q_lay.setContentsMargins(12, 12, 12, 12)
        q_lay.setSpacing(8)

        q_hdr = QHBoxLayout()
        q_title = QLabel("Queue")
        q_title.setObjectName("sectionLbl")
        q_hdr.addWidget(q_title)
        q_hdr.addStretch()
        self.queue_meta_lbl = QLabel("No items yet")
        self.queue_meta_lbl.setObjectName("mutedLbl")
        q_hdr.addWidget(self.queue_meta_lbl)
        q_lay.addLayout(q_hdr)

        self.q_list = QListWidget()
        self.q_list.setObjectName("queueList")
        self.q_list.setSpacing(4)
        self.q_list.setMinimumHeight(236)
        self.q_list.setMaximumHeight(236)
        self.q_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.q_list.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.q_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        q_lay.addWidget(self.q_list)

        q_btn_row = QHBoxLayout()
        q_btn_row.setSpacing(8)
        self.convert_selected_btn = QPushButton("Convert")
        self.convert_selected_btn.setObjectName("secondaryBtn")
        q_btn_row.addWidget(self.convert_selected_btn, 1)

        self.retry_btn = QPushButton("Retry")
        self.retry_btn.setObjectName("ghostBtn")
        q_btn_row.addWidget(self.retry_btn, 1)
        q_lay.addLayout(q_btn_row)

        manage_row = QHBoxLayout()
        manage_row.setSpacing(8)
        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setObjectName("ghostBtn")
        manage_row.addWidget(self.remove_btn, 1)

        self.clear_done_btn = QPushButton("Clear Saved")
        self.clear_done_btn.setObjectName("ghostBtn")
        manage_row.addWidget(self.clear_done_btn, 1)

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setObjectName("dangerGhostBtn")
        manage_row.addWidget(self.clear_btn)
        q_lay.addLayout(manage_row)

        lv.addWidget(queue_card)

        llm_card = QFrame()
        llm_card.setObjectName("sectionCard")
        llm_lay = QVBoxLayout(llm_card)
        llm_lay.setContentsMargins(12, 12, 12, 12)
        llm_lay.setSpacing(8)

        self.llm_chk = QCheckBox("Use OpenAI descriptions")
        llm_lay.addWidget(self.llm_chk)

        llm_note = QLabel(
            "Optional for charts, screenshots, and diagrams."
        )
        llm_note.setObjectName("mutedLbl")
        llm_note.setWordWrap(True)
        llm_lay.addWidget(llm_note)

        self.llm_grp = QFrame()
        self.llm_grp.setObjectName("subSectionCard")
        llm_grp_lay = QVBoxLayout(self.llm_grp)
        llm_grp_lay.setContentsMargins(12, 12, 12, 12)
        llm_grp_lay.setSpacing(8)

        key_lbl = QLabel("OpenAI API key")
        key_lbl.setObjectName("sectionMiniLbl")
        llm_grp_lay.addWidget(key_lbl)
        self.api_input = QLineEdit()
        self.api_input.setPlaceholderText("Paste your key for this session")
        self.api_input.setEchoMode(QLineEdit.EchoMode.Password)
        llm_grp_lay.addWidget(self.api_input)

        model_lbl = QLabel("Model")
        model_lbl.setObjectName("sectionMiniLbl")
        llm_grp_lay.addWidget(model_lbl)
        self.model_cb = QComboBox()
        self.model_cb.addItems(["gpt-4o", "gpt-4o-mini", "gpt-4.1"])
        llm_grp_lay.addWidget(self.model_cb)
        self.llm_grp.setVisible(False)
        llm_lay.addWidget(self.llm_grp)
        lv.addWidget(llm_card)

        lv.addStretch(1)

        scroll.setWidget(self.left_widget)
        left_outer.addWidget(scroll, 1)

        footer = QFrame()
        footer.setObjectName("stickyActionCard")
        footer_lay = QVBoxLayout(footer)
        footer_lay.setContentsMargins(14, 10, 14, 10)
        footer_lay.setSpacing(7)

        self.progress_lbl = QLabel(
            "Converted Markdown auto-saves to the output folder."
        )
        self.progress_lbl.setObjectName("mutedLbl")
        self.progress_lbl.setWordWrap(False)
        self.progress_lbl.setMaximumHeight(20)
        footer_lay.addWidget(self.progress_lbl)

        self.pbar = QProgressBar()
        self.pbar.setTextVisible(False)
        self.pbar.setVisible(False)
        self.pbar.setFixedHeight(7)
        footer_lay.addWidget(self.pbar)

        self.conv_btn = QPushButton("Add sources to begin")
        self.conv_btn.setObjectName("primaryBtn")
        self.conv_btn.setEnabled(False)
        self.conv_btn.setMinimumHeight(44)
        footer_lay.addWidget(self.conv_btn)

        left_outer.addWidget(footer, 0)
        self.splitter.addWidget(left_panel)

        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(18, 18, 18, 18)
        rv.setSpacing(14)

        info_card = QFrame()
        info_card.setObjectName("infoCard")
        info_lay = QVBoxLayout(info_card)
        info_lay.setContentsMargins(18, 16, 18, 16)
        info_lay.setSpacing(6)

        info_top = QHBoxLayout()
        info_top.setSpacing(10)
        self.file_lbl = QLabel("No item selected")
        self.file_lbl.setObjectName("fileLbl")
        info_top.addWidget(self.file_lbl, 1)
        self.status_chip = QLabel("Idle")
        self.status_chip.setObjectName("statusChip")
        self.status_chip.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_chip.setMinimumWidth(118)
        info_top.addWidget(self.status_chip, 0)
        info_lay.addLayout(info_top)

        self.stats_lbl = QLabel("Choose a queue item to inspect the source, rendered preview, and saved Markdown.")
        self.stats_lbl.setObjectName("metaLbl")
        self.stats_lbl.setWordWrap(True)
        info_lay.addWidget(self.stats_lbl)

        self.path_lbl = QLabel("")
        self.path_lbl.setObjectName("pathLbl")
        self.path_lbl.setWordWrap(True)
        info_lay.addWidget(self.path_lbl)
        rv.addWidget(info_card)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)

        self.raw_edit = QTextEdit()
        self.raw_edit.setReadOnly(True)
        self.raw_edit.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.raw_edit.setPlaceholderText("Converted Markdown will appear here.")
        self.tabs.addTab(self.raw_edit, "Markdown")

        self.prev_edit = QTextBrowser()
        self.prev_edit.setOpenExternalLinks(True)
        self.prev_edit.setPlaceholderText("Preview the rendered Markdown here.")
        self.tabs.addTab(self.prev_edit, "Preview")

        self.details_edit = QTextEdit()
        self.details_edit.setReadOnly(True)
        self.details_edit.setPlaceholderText("Metadata, save location, and conversion details appear here.")
        self.tabs.addTab(self.details_edit, "Details")

        rv.addWidget(self.tabs, 1)

        bottom_bar = QFrame()
        bottom_bar.setObjectName("actionBar")
        bb_lay = QHBoxLayout(bottom_bar)
        bb_lay.setContentsMargins(0, 0, 0, 0)
        bb_lay.setSpacing(8)

        self.open_source_btn = QPushButton("Open Source")
        self.open_source_btn.setObjectName("ghostBtn")
        bb_lay.addWidget(self.open_source_btn)

        self.open_saved_btn = QPushButton("Open Saved File")
        self.open_saved_btn.setObjectName("ghostBtn")
        bb_lay.addWidget(self.open_saved_btn)

        self.open_saved_folder_btn = QPushButton("Open Saved Folder")
        self.open_saved_folder_btn.setObjectName("ghostBtn")
        bb_lay.addWidget(self.open_saved_folder_btn)

        bb_lay.addStretch()

        self.copy_btn = QPushButton("Copy Markdown")
        self.copy_btn.setObjectName("secondaryBtn")
        bb_lay.addWidget(self.copy_btn)

        self.saveas_btn = QPushButton("Save As...")
        self.saveas_btn.setObjectName("secondaryBtn")
        bb_lay.addWidget(self.saveas_btn)
        rv.addWidget(bottom_bar)

        self.splitter.addWidget(right)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        self.splitter.splitterMoved.connect(lambda *_: self.sync_left_content_width())

        button_tips = [
            (self.browse_btn, "Add one or more files to the queue (Ctrl+O)"),
            (self.add_folder_btn, "Scan a folder in the background and add supported files (Ctrl+Shift+O)"),
            (self.url_btn, "Clip a single URL into the queue"),
            (self.folder_btn, "Choose where successful Markdown files are saved"),
            (self.open_output_btn, "Open the current auto-save folder"),
            (self.convert_selected_btn, "Convert only the currently selected queue item"),
            (self.retry_btn, "Return failed items to the ready state"),
            (self.remove_btn, "Remove the selected item from the queue"),
            (self.clear_done_btn, "Remove items that were already saved successfully"),
            (self.clear_btn, "Remove every queued item"),
            (self.conv_btn, "Convert ready items, or retry failed items when no ready items remain (Ctrl+Enter)"),
            (self.copy_btn, "Copy the current Markdown result to the clipboard"),
            (self.saveas_btn, "Save the current Markdown result to a custom location"),
        ]
        for button, tooltip in button_tips:
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setToolTip(tooltip)

    def make_metric_card(self, label_text):
        frame = QFrame()
        frame.setObjectName("metricCard")
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(0)

        value_lbl = QLabel("0")
        value_lbl.setObjectName("metricValueLbl")
        lay.addWidget(value_lbl)

        title_lbl = QLabel(label_text)
        title_lbl.setObjectName("metricTitleLbl")
        lay.addWidget(title_lbl)
        frame.setMinimumWidth(0)
        return frame, value_lbl

    def wire_signals(self):
        self.drop_zone.files_dropped.connect(self.add_sources)
        self.drop_zone.clicked.connect(self.browse)
        self.browse_btn.clicked.connect(self.browse)
        self.add_folder_btn.clicked.connect(self.browse_folder)
        self.url_btn.clicked.connect(self.add_url)
        self.url_input.returnPressed.connect(self.add_url)
        self.folder_btn.clicked.connect(self.pick_folder)
        self.open_output_btn.clicked.connect(self.open_output_folder)
        self.clear_btn.clicked.connect(self.clear_queue)
        self.remove_btn.clicked.connect(self.remove_selected)
        self.convert_selected_btn.clicked.connect(self.convert_selected)
        self.retry_btn.clicked.connect(self.retry_failed)
        self.clear_done_btn.clicked.connect(self.clear_saved)
        self.llm_chk.toggled.connect(self.on_llm_toggled)
        self.conv_btn.clicked.connect(self.handle_primary_action)
        self.q_list.currentRowChanged.connect(self.update_preview)
        self.q_list.itemDoubleClicked.connect(self.activate_current_item)
        self.q_list.customContextMenuRequested.connect(self.show_queue_menu)
        self.copy_btn.clicked.connect(self.copy_output)
        self.saveas_btn.clicked.connect(self.save_as)
        self.open_source_btn.clicked.connect(self.open_selected_source)
        self.open_saved_btn.clicked.connect(self.open_saved_file)
        self.open_saved_folder_btn.clicked.connect(self.open_saved_folder)

    def sync_left_content_width(self):
        scroll = getattr(self, "left_scroll", None)
        widget = getattr(self, "left_widget", None)
        if scroll is None or widget is None:
            return
        width = scroll.viewport().width()
        if width > 0 and widget.width() != width:
            widget.setFixedWidth(width)

    def eventFilter(self, obj, event):
        if obj is getattr(self, "left_scroll_viewport", None) and event.type() == QEvent.Type.Resize:
            QTimer.singleShot(0, self.sync_left_content_width)
        return super().eventFilter(obj, event)

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self.sync_left_content_width)
        QTimer.singleShot(50, self.sync_left_content_width)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QTimer.singleShot(0, self.sync_left_content_width)

    def setup_shortcuts(self):
        self.shortcuts = []
        bindings = [
            ("Ctrl+O", self.browse),
            ("Ctrl+Shift+O", self.browse_folder),
            ("Ctrl+L", self.url_input.setFocus),
            ("Ctrl+Return", self.start_conversion),
            ("Ctrl+Enter", self.start_conversion),
            ("Delete", self.remove_selected),
        ]
        for sequence, handler in bindings:
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.activated.connect(handler)
            self.shortcuts.append(shortcut)

    def queue_entry_index(self, entry_id):
        for idx, entry in enumerate(self.queue):
            if entry.entry_id == entry_id:
                return idx
        return -1

    def ensure_queue_mutable(self, action_text):
        if self.worker is None:
            return True
        self.show_message("Wait for the current conversion to finish before you %s." % action_text)
        return False

    def convert_selected(self):
        row = self.q_list.currentRow()
        if not (0 <= row < len(self.queue)):
            self.show_message("Select a ready item to convert.")
            return
        if self.queue[row].status != "queued":
            self.show_message("Select a ready item to convert.")
            return
        self.start_conversion_rows([row])

    def activate_current_item(self, *_args):
        entry = self.current_entry()
        if not entry:
            return
        if entry.status == "queued":
            self.convert_selected()
        elif entry.saved_path:
            self.open_saved_file()
        else:
            self.tabs.setCurrentWidget(self.prev_edit)

    def on_system_theme_changed(self):
        self._dark = True
        self.theme = DARK
        self.apply_theme()
        self.refresh_queue()
        self.update_preview(self.q_list.currentRow())

    def apply_theme(self):
        theme = self.theme
        app = QApplication.instance()
        app.setStyle(QStyleFactory.create("Fusion"))
        app.setPalette(build_palette(theme))
        app.setFont(build_ui_font())
        self.setStyleSheet(self.build_css(theme))

        if self.icon_path:
            icon = QIcon(self.icon_path)
            app.setWindowIcon(icon)
            self.setWindowIcon(icon)

        self.raw_edit.setFont(build_editor_font())
        self.details_edit.setFont(build_editor_font())
        self.prev_edit.document().setDefaultStyleSheet(
            "body { font-family: 'Segoe UI Variable Text', 'Aptos', 'Segoe UI'; color: %s; } "
            "h1, h2, h3, h4, h5 { color: %s; } "
            "code, pre { font-family: 'Cascadia Mono', 'Consolas', monospace; "
            "background: %s; } "
            "a { color: %s; }"
            % (theme["text"], theme["accent"], theme["card2"], theme["accent"])
        )
        self.update_status_chip(self.status_chip, "queued", "Idle")

    def build_css(self, theme):
        return """
        QMainWindow#mainWindow {
            background: %(bg)s;
        }
        QWidget#leftRail {
            background: %(bg)s;
        }
        QStatusBar {
            background: %(surface)s;
            color: %(muted)s;
            border-top: 1px solid %(border)s;
            padding: 3px 10px;
        }
        QFrame#stickyActionCard {
            background: %(card)s;
            border: 1px solid %(border)s;
            border-radius: 8px;
        }
        QFrame#actionBar {
            background: transparent;
        }
        QScrollArea {
            border: none;
            background: transparent;
        }
        QScrollBar:vertical {
            background: %(surface)s;
            width: 8px;
            margin: 4px 2px 4px 0;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical {
            background: %(border2)s;
            min-height: 28px;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical:hover {
            background: %(accent2)s;
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical,
        QScrollBar::up-arrow:vertical,
        QScrollBar::down-arrow:vertical,
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical {
            height: 0px;
            background: transparent;
        }
        QFrame#sectionCard, QFrame#infoCard, QFrame#subSectionCard, QFrame#sourceCard, QFrame#queueCard {
            background: %(card)s;
            border: 1px solid %(border)s;
            border-radius: 8px;
        }
        QFrame#heroCard {
            background: %(card)s;
            border: 1px solid %(border)s;
            border-radius: 8px;
        }
        QFrame#sourceCard {
            background: %(card)s;
        }
        QFrame#queueCard {
            background: %(card)s;
        }
        QFrame#subSectionCard {
            background: %(card2)s;
            border-radius: 8px;
        }
        QFrame#metricCard {
            background: %(surface2)s;
            border: 1px solid %(border)s;
            border-radius: 8px;
        }
        QLabel#eyebrowLbl {
            color: %(accent)s;
            font-size: 10px;
            letter-spacing: 1px;
            font-weight: 700;
        }
        QLabel#titleLbl {
            color: %(text)s;
            font-family: 'Segoe UI Variable Display', 'Aptos Display', 'Segoe UI';
            font-size: 26px;
            font-weight: 800;
        }
        QLabel#subLbl {
            color: %(muted)s;
            font-size: 13px;
            line-height: 1.35;
        }
        QLabel#sectionLbl {
            color: %(text)s;
            font-size: 14px;
            font-weight: 700;
        }
        QLabel#sectionMiniLbl {
            color: %(text)s;
            font-size: 12px;
            font-weight: 600;
        }
        QLabel#mutedLbl {
            color: %(muted)s;
            font-size: 12px;
            line-height: 1.3;
        }
        QLabel#metricValueLbl {
            color: %(text)s;
            font-size: 20px;
            font-weight: 800;
        }
        QLabel#metricTitleLbl {
            color: %(muted)s;
            font-size: 11px;
            font-weight: 600;
        }
        QLabel#fileLbl {
            color: %(text)s;
            font-family: 'Segoe UI Variable Display', 'Aptos Display', 'Segoe UI';
            font-size: 24px;
            font-weight: 800;
        }
        QLabel#metaLbl {
            color: %(muted)s;
            font-size: 13px;
        }
        QLabel#pathLbl {
            color: %(muted)s;
            font-size: 12px;
        }
        QLabel#statusChip {
            background: %(accent_soft)s;
            color: %(accent)s;
            border: 1px solid %(border2)s;
            border-radius: 4px;
            padding: 6px 12px;
            font-size: 11px;
            font-weight: 700;
        }
        QPushButton {
            min-height: 34px;
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid %(border2)s;
            background: %(surface2)s;
            color: %(text)s;
            font-weight: 600;
        }
        QPushButton:hover {
            border-color: %(accent2)s;
            background: %(card2)s;
        }
        QPushButton:pressed {
            border-color: %(accent2)s;
            background: %(surface)s;
        }
        QPushButton:disabled {
            background: %(surface)s;
            color: %(muted)s;
            border-color: %(border)s;
        }
        QPushButton#primaryBtn {
            background: %(accent)s;
            border: none;
            color: %(btn_text)s;
            font-size: 14px;
            font-weight: 800;
            padding: 12px 16px;
        }
        QPushButton#primaryBtn:hover {
            background: %(accent2)s;
        }
        QPushButton#secondaryBtn {
            background: %(card2)s;
            border-color: %(accent2)s;
        }
        QPushButton#ghostBtn {
            background: transparent;
        }
        QPushButton#dangerGhostBtn {
            background: transparent;
            color: %(error)s;
            border-color: %(border2)s;
        }
        QLineEdit, QComboBox, QTextEdit, QTextBrowser, QListWidget {
            background: %(surface)s;
            color: %(text)s;
            border: 1px solid %(border)s;
            border-radius: 6px;
        }
        QLineEdit, QComboBox {
            padding: 9px 12px;
        }
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QTextBrowser:focus, QListWidget:focus {
            border-color: %(accent2)s;
        }
        QListWidget#queueList {
            padding: 6px;
            outline: none;
        }
        QListWidget#queueList::item {
            background: %(surface)s;
            border: 1px solid %(border)s;
            border-radius: 6px;
            padding: 6px 8px;
            margin: 0px 0px 3px 0px;
        }
        QListWidget#queueList::item:selected {
            background: %(accent_soft)s;
            border-color: %(accent2)s;
            color: %(text)s;
        }
        QTabWidget::pane {
            border: 1px solid %(border)s;
            border-radius: 8px;
            background: %(card)s;
            top: -1px;
        }
        QTabBar::tab {
            background: %(surface)s;
            color: %(muted)s;
            border: 1px solid %(border)s;
            border-bottom: none;
            padding: 10px 18px;
            min-width: 112px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
            font-weight: 700;
        }
        QTabBar::tab:selected {
            background: %(card)s;
            color: %(text)s;
            border-color: %(accent2)s;
        }
        QTabBar::tab:hover:!selected {
            color: %(text)s;
            border-color: %(border2)s;
        }
        QTextEdit, QTextBrowser {
            padding: 12px;
            selection-background-color: %(accent)s;
            selection-color: %(btn_text)s;
        }
        QProgressBar {
            background: %(surface)s;
            border: 1px solid %(border)s;
            border-radius: 4px;
        }
        QProgressBar::chunk {
            background: %(accent)s;
            border-radius: 4px;
        }
        QCheckBox {
            color: %(text)s;
            spacing: 8px;
            font-weight: 600;
        }
        QCheckBox::indicator {
            width: 18px;
            height: 18px;
            border-radius: 4px;
            border: 2px solid %(border2)s;
            background: %(surface)s;
        }
        QCheckBox::indicator:checked {
            background: %(accent)s;
            border-color: %(accent)s;
        }
        """ % theme

    def restore_settings(self):
        self.output_folder = self.settings.value("output_folder", default_output_folder(), type=str)
        self.last_browse_dir = self.settings.value("last_browse_dir", str(Path.home()), type=str)
        self.model_cb.setCurrentText(
            self.settings.value("llm_model", "gpt-4o-mini", type=str)
        )
        self.llm_chk.setChecked(self.settings.value("llm_enabled", False, type=bool))

        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)

        layout_version = "2026-04-practical-rail"
        saved_layout_version = self.settings.value("layout_version", "", type=str)
        splitter_state = self.settings.value("splitter_state")
        if splitter_state and saved_layout_version == layout_version:
            self.splitter.restoreState(splitter_state)
        else:
            self.splitter.setSizes([470, 980])
            self.settings.setValue("layout_version", layout_version)

        self.update_output_label()
        self.on_llm_toggled(self.llm_chk.isChecked())

    def closeEvent(self, event):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("splitter_state", self.splitter.saveState())
        self.settings.setValue("layout_version", "2026-04-practical-rail")
        self.settings.setValue("output_folder", self.output_folder)
        self.settings.setValue("last_browse_dir", self.last_browse_dir)
        self.settings.setValue("llm_enabled", self.llm_chk.isChecked())
        self.settings.setValue("llm_model", self.model_cb.currentText())
        super().closeEvent(event)

    def on_llm_toggled(self, enabled):
        self.llm_grp.setVisible(enabled)
        self.update_summary()

    def update_output_label(self):
        if self.output_folder:
            self.folder_lbl.setText(
                "New conversions will auto-save to:\n%s" % shorten_middle(self.output_folder, 92)
            )
        else:
            self.folder_lbl.setText("Choose an output folder to auto-save converted Markdown.")
        self.open_output_btn.setEnabled(bool(self.output_folder))

    def source_exists(self, source, is_url):
        key = source.strip().lower() if is_url else str(Path(source).resolve()).lower()
        return any(entry.key() == key for entry in self.queue)

    def start_folder_scan(self, folders):
        pending = []
        seen = set()
        for folder in folders:
            normalized = str(Path(folder))
            if normalized not in seen:
                seen.add(normalized)
                pending.append(normalized)
        if not pending:
            return

        if self.scan_worker is not None:
            queued_now = 0
            known = set(self._queued_scan_folders) | set(self._active_scan_folders)
            for folder in pending:
                if folder not in known:
                    self._queued_scan_folders.append(folder)
                    known.add(folder)
                    queued_now += 1
            if queued_now:
                self.show_message("Scanning in progress. Queued %d more folder%s." % (
                    queued_now, "" if queued_now == 1 else "s"
                ))
                self.update_progress_note()
            return

        self._active_scan_folders = pending
        self.scan_worker = FolderScanWorker(pending)
        self.scan_worker.scan_done.connect(self.on_folder_scan_done)
        self.scan_worker.start()
        self.update_progress_note()
        self.show_message("Scanning %d folder%s for supported files..." % (
            len(pending), "" if len(pending) == 1 else "s"
        ), 10000)

    def on_folder_scan_done(self, folders, files, errors):
        if self.scan_worker is not None:
            self.scan_worker.deleteLater()
        self.scan_worker = None
        self._active_scan_folders = []

        added = 0
        duplicates = 0
        first_added_index = -1
        for path in files:
            a, d = self.add_source_item(path, False)
            added += a
            duplicates += d
            if a and first_added_index < 0:
                first_added_index = len(self.queue) - 1

        self.refresh_queue()
        if first_added_index >= 0 and self.q_list.currentRow() < 0:
            self.q_list.setCurrentRow(first_added_index)

        parts = ["Scanned %d folder%s." % (len(folders), "" if len(folders) == 1 else "s")]
        if added:
            parts.append("Added %d item%s." % (added, "" if added == 1 else "s"))
        if duplicates:
            parts.append("Skipped %d duplicate%s." % (duplicates, "" if duplicates == 1 else "s"))
        if not files:
            parts.append("No supported files were found.")
        if errors:
            parts.append("%d folder access issue%s were skipped." % (
                len(errors), "" if len(errors) == 1 else "s"
            ))
        self.show_message(" ".join(parts), 10000)

        if self._queued_scan_folders:
            queued = self._queued_scan_folders[:]
            self._queued_scan_folders = []
            self.start_folder_scan(queued)
        else:
            self.update_progress_note()

    def add_sources(self, paths):
        added = 0
        duplicates = 0
        unsupported = 0
        first_added_index = -1
        dir_paths = []

        for raw_path in paths:
            path = Path(raw_path)
            if path.is_dir():
                dir_paths.append(str(path))
                continue

            if path.is_file() and path.suffix.lstrip(".").lower() in EXTS:
                a, d = self.add_source_item(str(path), False)
                added += a
                duplicates += d
                if a and first_added_index < 0:
                    first_added_index = len(self.queue) - 1
            else:
                unsupported += 1

        self.refresh_queue()
        if first_added_index >= 0 and self.q_list.currentRow() < 0:
            self.q_list.setCurrentRow(first_added_index)

        parts = []
        if added:
            parts.append("Added %d item%s." % (added, "" if added == 1 else "s"))
        if duplicates:
            parts.append("Skipped %d duplicate%s." % (duplicates, "" if duplicates == 1 else "s"))
        if unsupported:
            parts.append("Ignored %d unsupported path%s." % (unsupported, "" if unsupported == 1 else "s"))
        if dir_paths:
            self.start_folder_scan(dir_paths)
            parts.append("Scanning %d folder%s in the background." % (
                len(dir_paths), "" if len(dir_paths) == 1 else "s"
            ))
        if not parts:
            parts.append("Nothing new was added to the queue.")
        self.show_message(" ".join(parts))

    def add_source_item(self, source, is_url):
        if self.source_exists(source, is_url):
            return 0, 1

        if is_url:
            name = shorten_middle(source.split("//", 1)[-1], 42)
        else:
            name = os.path.basename(source)

        self.queue.append(QueueEntry(source=source, is_url=is_url, name=name))
        return 1, 0

    def browse(self):
        self._block_show = True
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select files to convert",
            self.last_browse_dir,
            FILTER_STR,
        )
        self._block_show = False
        if files:
            self.last_browse_dir = str(Path(files[0]).parent)
            self.add_sources(files)

    def browse_folder(self):
        self._block_show = True
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select a folder to scan for supported files",
            self.last_browse_dir,
        )
        self._block_show = False
        if folder:
            self.last_browse_dir = folder
            self.add_sources([folder])

    def add_url(self):
        url = self.url_input.text().strip()
        if not url:
            self.show_message("Paste an http or https URL first.")
            return
        if not url.startswith(("http://", "https://")):
            self.show_message("Only http and https URLs are supported.")
            return

        added, duplicates = self.add_source_item(url, True)
        if duplicates:
            self.show_message("That URL is already in the queue.")
            return

        self.url_input.clear()
        self.refresh_queue()
        if added:
            self.q_list.setCurrentRow(len(self.queue) - 1)
            self.show_message("Added URL to the queue.")

    def pick_folder(self):
        self._block_show = True
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select where converted Markdown should be saved",
            self.output_folder or self.last_browse_dir,
        )
        self._block_show = False
        if folder:
            self.output_folder = folder
            self.update_output_label()
            self.show_message("Output folder updated.")

    def open_output_folder(self):
        if not self.output_folder:
            self.show_message("Choose an output folder first.")
            return
        Path(self.output_folder).mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_folder))

    def refresh_queue(self):
        selected_key = None
        row = self.q_list.currentRow()
        if 0 <= row < len(self.queue):
            selected_key = self.queue[row].key()

        self.q_list.clear()
        selected_row = -1
        for idx, entry in enumerate(self.queue):
            item = QListWidgetItem(self.build_queue_text(entry))
            item.setToolTip(self.build_queue_tooltip(entry))
            item.setSizeHint(QSize(0, 34))

            color_map = {
                "done": self.theme["success"],
                "error": self.theme["error"],
                "converting": self.theme["warning"],
            }
            item.setForeground(QColor(color_map.get(entry.status, self.theme["text"])))
            self.q_list.addItem(item)

            if selected_key and entry.key() == selected_key:
                selected_row = idx

        if selected_row >= 0:
            self.q_list.setCurrentRow(selected_row)
        elif self.queue and row == -1:
            self.q_list.setCurrentRow(0)

        self.update_summary()

    def build_queue_text(self, entry):
        return "%s  %s" % (BADGES[entry.status], shorten_middle(entry.name, 42))

    def build_queue_tooltip(self, entry):
        lines = [
            "Source: %s" % entry.source,
            "Type: %s" % ("URL" if entry.is_url else "File"),
            "Status: %s" % entry.status_label(),
            "Added: %s" % format_dt(entry.added_at),
        ]
        if entry.completed_at:
            lines.append("Completed: %s" % format_dt(entry.completed_at))
        if entry.saved_path:
            lines.append("Saved to: %s" % entry.saved_path)
        if entry.error:
            lines.append("Error: %s" % entry.error)
        return "\n".join(lines)

    def show_queue_menu(self, pos):
        item = self.q_list.itemAt(pos)
        if not item:
            return

        row = self.q_list.row(item)
        if not (0 <= row < len(self.queue)):
            return
        self.q_list.setCurrentRow(row)
        entry = self.queue[row]

        menu = QMenu(self)
        convert_action = menu.addAction("Convert this item")
        convert_action.setEnabled(entry.status == "queued" and self.worker is None)
        menu.addSeparator()
        open_source_action = menu.addAction("Open source")
        open_saved_action = menu.addAction("Open saved file")
        open_saved_action.setEnabled(bool(entry.saved_path))
        menu.addSeparator()
        retry_action = menu.addAction("Retry this item")
        retry_action.setEnabled(entry.status == "error" and self.worker is None)
        remove_action = menu.addAction("Remove from queue")
        remove_action.setEnabled(self.worker is None)

        chosen = menu.exec(self.q_list.viewport().mapToGlobal(pos))
        if chosen == convert_action:
            self.start_conversion_rows([row])
        elif chosen == open_source_action:
            self.open_selected_source()
        elif chosen == open_saved_action:
            self.open_saved_file()
        elif chosen == retry_action:
            self.retry_failed(selected_only=True)
        elif chosen == remove_action:
            self.remove_selected()

    def remove_selected(self):
        if not self.ensure_queue_mutable("change the queue"):
            return
        row = self.q_list.currentRow()
        if not (0 <= row < len(self.queue)):
            self.show_message("Select a queue item to remove.")
            return
        removed = self.queue.pop(row)
        self.refresh_queue()
        if self.queue:
            self.q_list.setCurrentRow(min(row, len(self.queue) - 1))
        self.show_message("Removed %s from the queue." % removed.name)

    def retry_failed(self, selected_only=False):
        if not self.ensure_queue_mutable("retry failed items"):
            return
        rows = []
        if selected_only:
            row = self.q_list.currentRow()
            if 0 <= row < len(self.queue) and self.queue[row].status == "error":
                rows = [row]
        else:
            rows = [i for i, entry in enumerate(self.queue) if entry.status == "error"]

        if not rows:
            self.show_message("There are no failed items to retry.")
            return

        for row in rows:
            entry = self.queue[row]
            entry.status = "queued"
            entry.error = ""
            entry.result = ""
            entry.saved_path = ""
            entry.completed_at = None

        self.refresh_queue()
        self.q_list.setCurrentRow(rows[0])
        self.show_message("Returned %d failed item%s to the queue." % (
            len(rows), "" if len(rows) == 1 else "s"
        ))

    def handle_primary_action(self):
        queued = sum(1 for entry in self.queue if entry.status == "queued")
        errors = sum(1 for entry in self.queue if entry.status == "error")
        if queued:
            self.start_conversion()
            return
        if errors and self.worker is None:
            rows = [i for i, entry in enumerate(self.queue) if entry.status == "error"]
            for row in rows:
                entry = self.queue[row]
                entry.status = "queued"
                entry.error = ""
                entry.result = ""
                entry.saved_path = ""
                entry.completed_at = None
            self.refresh_queue()
            self.q_list.setCurrentRow(rows[0])
            self.start_conversion_rows(rows)
            return
        self.start_conversion()

    def clear_saved(self):
        if not self.ensure_queue_mutable("clear saved items"):
            return
        before = len(self.queue)
        self.queue = [entry for entry in self.queue if entry.status != "done"]
        removed = before - len(self.queue)
        self.refresh_queue()
        self.update_preview(self.q_list.currentRow())
        self.show_message(
            "Cleared %d saved item%s." % (removed, "" if removed == 1 else "s")
            if removed else "There are no saved items to clear."
        )

    def clear_queue(self):
        if not self.ensure_queue_mutable("clear the queue"):
            return
        if not self.queue:
            self.show_message("The queue is already empty.")
            return

        answer = QMessageBox.question(
            self,
            "Clear queue",
            "Remove every item from the queue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        self.queue = []
        self.refresh_queue()
        self.update_preview(-1)
        self.show_message("Queue cleared.")

    def refresh_convert_btn(self):
        queued = sum(1 for entry in self.queue if entry.status == "queued")
        errors = sum(1 for entry in self.queue if entry.status == "error")
        can_convert = (queued > 0 or errors > 0) and self.worker is None
        self.conv_btn.setEnabled(can_convert)
        if queued:
            self.conv_btn.setText("Convert %d ready item%s" % (queued, "" if queued == 1 else "s"))
        elif self.worker is not None:
            self.conv_btn.setText("Converting...")
        elif errors:
            self.conv_btn.setText("Retry %d failed item%s" % (errors, "" if errors == 1 else "s"))
        else:
            self.conv_btn.setText("Add sources to begin")

    def update_summary(self):
        queued = sum(1 for entry in self.queue if entry.status == "queued")
        converting = sum(1 for entry in self.queue if entry.status == "converting")
        done = sum(1 for entry in self.queue if entry.status == "done")
        errors = sum(1 for entry in self.queue if entry.status == "error")
        selected_row = self.q_list.currentRow()
        selected_ready = 0 <= selected_row < len(self.queue) and self.queue[selected_row].status == "queued"
        self.total_metric_lbl.setText(str(queued))
        self.done_metric_lbl.setText(str(done))
        self.issue_metric_lbl.setText(str(errors))

        meta_bits = []
        if queued:
            meta_bits.append("%d ready" % queued)
        if converting:
            meta_bits.append("%d converting" % converting)
        if done:
            meta_bits.append("%d saved" % done)
        if errors:
            meta_bits.append("%d issue%s" % (errors, "" if errors == 1 else "s"))
        if meta_bits:
            self.queue_meta_lbl.setText(" | ".join(meta_bits))
        else:
            self.queue_meta_lbl.setText("No items yet")

        self.convert_selected_btn.setEnabled(selected_ready and self.worker is None)
        self.remove_btn.setEnabled(self.q_list.currentRow() >= 0 and self.worker is None)
        self.retry_btn.setEnabled(errors > 0 and self.worker is None)
        self.clear_done_btn.setEnabled(done > 0 and self.worker is None)
        self.clear_btn.setEnabled(bool(self.queue) and self.worker is None)
        self.open_output_btn.setEnabled(bool(self.output_folder))
        self.refresh_convert_btn()
        self.update_progress_note()

    def update_progress_note(self):
        queued = sum(1 for entry in self.queue if entry.status == "queued")
        converting = sum(1 for entry in self.queue if entry.status == "converting")
        errors = sum(1 for entry in self.queue if entry.status == "error")

        if self.worker is not None and self._run_entry_ids:
            total = len(self._run_entry_ids)
            current = min(total, self._prog + 1)
            label = shorten_middle(os.path.basename(self._active_source) or self._active_source, 34)
            self.progress_lbl.setText(
                "Converting %d of %d%s" % (
                    current,
                    total,
                    (" - %s" % label) if label else "",
                )
            )
            return

        if self.scan_worker is not None and self._active_scan_folders:
            self.progress_lbl.setText("Scanning %d folder%s for supported files..." % (
                len(self._active_scan_folders), "" if len(self._active_scan_folders) == 1 else "s"
            ))
            return

        bits = []
        if queued:
            bits.append("%d ready" % queued)
        if converting:
            bits.append("%d converting" % converting)
        if errors:
            bits.append("%d issue%s" % (errors, "" if errors == 1 else "s"))
        if bits:
            self.progress_lbl.setText("Queue summary: " + " | ".join(bits))
        else:
            self.progress_lbl.setText(
                "Add sources. Successful conversions auto-save."
            )

    def start_conversion(self):
        self.start_conversion_rows()

    def start_conversion_rows(self, rows=None):
        if self.worker is not None:
            self.show_message("Wait for the current conversion to finish before starting another batch.")
            return

        sources = []
        run_entry_ids = []
        target_rows = rows if rows is not None else range(len(self.queue))
        for i in target_rows:
            if not (0 <= i < len(self.queue)):
                continue
            entry = self.queue[i]
            if entry.status == "queued":
                entry.status = "converting"
                entry.error = ""
                entry.result = ""
                run_entry_ids.append(entry.entry_id)
                sources.append(entry.source)

        if not sources:
            if rows is None:
                self.show_message("Add or retry at least one queue item first.")
            else:
                self.show_message("Select a queued item to convert.")
            return

        llm_cfg = None
        if self.llm_chk.isChecked():
            key = self.api_input.text().strip()
            if not key:
                QMessageBox.information(
                    self,
                    "OpenAI key required",
                    "Paste your OpenAI API key or turn off LLM enrichment before converting.",
                )
                self.api_input.setFocus()
                for entry_id in run_entry_ids:
                    idx = self.queue_entry_index(entry_id)
                    if idx >= 0:
                        self.queue[idx].status = "queued"
                self.refresh_queue()
                return
            llm_cfg = {
                "api_key": key,
                "model": self.model_cb.currentText(),
            }

        self._run_entry_ids = run_entry_ids
        self._prog = 0
        self._active_source = ""
        self.pbar.setVisible(True)
        self.pbar.setMaximum(len(sources))
        self.pbar.setValue(0)

        self.worker = ConvertWorker(sources, llm_cfg)
        self.worker.item_started.connect(self.on_item_started)
        self.worker.item_done.connect(self.on_item_done)
        self.worker.all_done.connect(self.on_all_done)
        self.worker.start()

        self.refresh_queue()
        if self.q_list.currentRow() < 0 and self.queue:
            first_row = self.queue_entry_index(run_entry_ids[0])
            if first_row >= 0:
                self.q_list.setCurrentRow(first_row)
        self.show_message("Started converting %d item%s." % (
            len(sources), "" if len(sources) == 1 else "s"
        ))

    def on_item_started(self, worker_idx, source):
        if worker_idx < len(self._run_entry_ids):
            self._active_source = source
            self.update_progress_note()

    def on_item_done(self, worker_idx, result, error):
        if worker_idx >= len(self._run_entry_ids):
            return
        real_idx = self.queue_entry_index(self._run_entry_ids[worker_idx])
        if real_idx < 0:
            return

        entry = self.queue[real_idx]
        entry.completed_at = datetime.now()
        if error:
            entry.status = "error"
            entry.error = error
            entry.result = ""
            entry.saved_path = ""
        else:
            entry.result = result
            entry.saved_path, save_error = self.auto_save(entry)
            if save_error:
                entry.status = "error"
                entry.error = save_error
            else:
                entry.status = "done"
                entry.error = ""

        self._prog += 1
        self.pbar.setValue(self._prog)

        self.refresh_queue()
        if self.q_list.currentRow() < 0:
            self.q_list.setCurrentRow(real_idx)
        elif self.q_list.currentRow() == real_idx:
            self.update_preview(real_idx)
            if entry.result:
                self.tabs.setCurrentWidget(self.prev_edit)

    def on_all_done(self):
        self.pbar.setVisible(False)
        self._active_source = ""
        if self.worker is not None:
            self.worker.deleteLater()
        self.worker = None
        self._run_entry_ids = []
        self.refresh_queue()

        done = sum(1 for entry in self.queue if entry.status == "done")
        errors = sum(1 for entry in self.queue if entry.status == "error")
        self.show_message(
            "Conversion complete - %d saved, %d issue%s." % (
                done,
                errors,
                "" if errors == 1 else "s",
            )
        )

    def auto_save(self, entry):
        folder = self.output_folder or default_output_folder()
        self.output_folder = folder
        self.update_output_label()
        try:
            output_path = unique_output_path(folder, entry.name)
            with open(output_path, "w", encoding="utf-8") as fh:
                fh.write(entry.result)
            return output_path, ""
        except Exception as exc:
            return "", "Auto-save failed: %s" % str(exc)

    def update_preview(self, row):
        if self._block_show:
            return
        if row < 0 or row >= len(self.queue):
            self.file_lbl.setText("No item selected")
            self.stats_lbl.setText(
                "Choose a queue item to inspect the source, rendered preview, and saved Markdown."
            )
            self.path_lbl.setText("")
            self.update_status_chip(self.status_chip, "queued", "Idle")
            self.raw_edit.clear()
            self.prev_edit.setHtml(
                "<p style='color:%s;'>Select a converted item to preview rendered Markdown and inspect the final output.</p>"
                % self.theme["muted"]
            )
            self.details_edit.setPlainText(
                "Use the left workflow to collect sources, review the queue, and convert them into reusable Markdown here."
            )
            self.copy_btn.setEnabled(False)
            self.saveas_btn.setEnabled(False)
            self.open_source_btn.setEnabled(False)
            self.open_saved_btn.setEnabled(False)
            self.open_saved_folder_btn.setEnabled(bool(self.output_folder))
            return

        entry = self.queue[row]
        self.file_lbl.setText(entry.name)
        self.stats_lbl.setText(self.build_stats_text(entry))
        self.path_lbl.setText(self.build_path_text(entry))
        self.update_status_chip(self.status_chip, entry.status, entry.status_label())

        if entry.result:
            self.raw_edit.setPlainText(entry.result)
            self.prev_edit.setMarkdown(entry.result)
        elif entry.status == "error":
            error_text = "Conversion error\n\n%s" % entry.error
            self.raw_edit.setPlainText(error_text)
            self.prev_edit.setHtml(
                "<h3 style='color:%s;'>Conversion error</h3><p>%s</p>"
                % (self.theme["error"], entry.error.replace("\n", "<br>"))
            )
        else:
            if entry.status == "queued":
                note = "This item is ready. Convert it from the queue or run the main Convert action to generate Markdown."
            elif entry.status == "converting":
                note = "This item is converting now. The rendered preview will appear here as soon as the conversion finishes."
            else:
                note = "This item has not finished converting yet."
            self.raw_edit.setPlainText(note)
            self.prev_edit.setHtml("<p style='color:%s;'>%s</p>" % (self.theme["muted"], note))

        self.details_edit.setPlainText(self.build_details_text(entry))
        self.copy_btn.setEnabled(bool(entry.result))
        self.saveas_btn.setEnabled(bool(entry.result))
        self.open_source_btn.setEnabled(True)
        self.open_saved_btn.setEnabled(bool(entry.saved_path))
        self.open_saved_folder_btn.setEnabled(bool(entry.saved_path or self.output_folder))

    def update_status_chip(self, label, status, text):
        color_map = {
            "queued": (self.theme["accent_soft"], self.theme["accent"]),
            "converting": (self.theme["accent_soft"], self.theme["warning"]),
            "done": ("#14372d" if self._dark else "#dff7ee", self.theme["success"]),
            "error": ("#3a1919" if self._dark else "#fde7e7", self.theme["error"]),
        }
        background, color = color_map.get(status, (self.theme["accent_soft"], self.theme["accent"]))
        label.setText(text)
        label.setStyleSheet(
            "QLabel { background:%s; color:%s; border:1px solid %s; border-radius:4px; "
            "padding:6px 12px; font-size:11px; font-weight:700; }"
            % (background, color, self.theme["border2"])
        )

    def build_stats_text(self, entry):
        if entry.result:
            words = len(entry.result.split())
            lines = entry.result.count("\n") + 1
            chars = len(entry.result)
            parts = [
                entry.status_label(),
                "%d words" % words,
                "%d lines" % lines,
                "%d chars" % chars,
            ]
            if entry.error:
                parts.append("save issue")
            return " | ".join(parts)
        if entry.status == "error":
            return "%s | %s" % (entry.status_label(), shorten_middle(entry.error, 72))
        return "%s | Added %s" % (entry.status_label(), format_dt(entry.added_at))

    def build_path_text(self, entry):
        bits = ["Source: %s" % shorten_middle(entry.source, 110)]
        if entry.saved_path:
            bits.append("Saved: %s" % shorten_middle(entry.saved_path, 110))
        elif self.output_folder:
            bits.append("Auto-save folder: %s" % shorten_middle(self.output_folder, 110))
        if entry.error:
            bits.append("Issue: %s" % shorten_middle(entry.error, 110))
        return "\n".join(bits)

    def build_details_text(self, entry):
        lines = [
            "Name: %s" % entry.name,
            "Type: %s" % ("URL" if entry.is_url else "File"),
            "Status: %s" % entry.status_label(),
            "Source: %s" % entry.source,
            "Added: %s" % format_dt(entry.added_at),
            "Completed: %s" % format_dt(entry.completed_at),
        ]
        if entry.saved_path:
            lines.append("Saved file: %s" % entry.saved_path)
        elif self.output_folder:
            lines.append("Output folder: %s" % self.output_folder)
        if entry.status == "queued":
            lines.extend(["", "Next step:", "  Run conversion to generate Markdown for this item."])
        elif entry.status == "converting":
            lines.extend(["", "Next step:", "  Conversion is in progress. Preview will update automatically."])
        if entry.error:
            lines.extend(["", "Error:", entry.error])
        if entry.result:
            lines.extend([
                "",
                "Result stats:",
                "  Words: %d" % len(entry.result.split()),
                "  Lines: %d" % (entry.result.count("\n") + 1),
                "  Characters: %d" % len(entry.result),
            ])
        return "\n".join(lines)

    def current_entry(self):
        row = self.q_list.currentRow()
        if 0 <= row < len(self.queue):
            return self.queue[row]
        return None

    def copy_output(self):
        entry = self.current_entry()
        if entry and entry.result:
            QApplication.clipboard().setText(entry.result)
            self.show_message("Copied Markdown to the clipboard.")

    def save_as(self):
        entry = self.current_entry()
        if not entry or not entry.result:
            self.show_message("Convert an item first, then save the Markdown.")
            return
        self._block_show = True
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save converted Markdown",
            Path(entry.name).stem + ".md",
            "Markdown (*.md)",
        )
        self._block_show = False
        if path:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(entry.result)
            entry.saved_path = path
            self.refresh_queue()
            self.update_preview(self.q_list.currentRow())
            self.show_message("Saved Markdown to %s." % shorten_middle(path, 80))

    def open_selected_source(self):
        entry = self.current_entry()
        if not entry:
            self.show_message("Select a queue item first.")
            return

        if entry.is_url:
            QDesktopServices.openUrl(QUrl(entry.source))
            return

        if os.path.exists(entry.source):
            QDesktopServices.openUrl(QUrl.fromLocalFile(entry.source))
        else:
            self.show_message("The source file is no longer available.")

    def open_saved_file(self):
        entry = self.current_entry()
        if not entry or not entry.saved_path:
            self.show_message("This item does not have a saved Markdown file yet.")
            return
        if os.path.exists(entry.saved_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(entry.saved_path))
        else:
            self.show_message("The saved Markdown file could not be found.")

    def open_saved_folder(self):
        entry = self.current_entry()
        if entry and entry.saved_path and os.path.exists(entry.saved_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(entry.saved_path).parent)))
            return
        self.open_output_folder()

    def show_message(self, text, timeout=5000):
        self.statusBar().showMessage(text, timeout)


def main():
    if "--self-test-convert" in sys.argv:
        idx = sys.argv.index("--self-test-convert")
        log_path = os.environ.get(
            "MARKITDOWN_SELF_TEST_LOG",
            str(Path.cwd() / "markitdown_self_test.log"),
        )
        if idx + 1 >= len(sys.argv):
            Path(log_path).write_text("Missing path after --self-test-convert\n", encoding="utf-8")
            print("Missing path after --self-test-convert", file=sys.stderr)
            sys.exit(2)
        try:
            from markitdown import MarkItDown

            result = MarkItDown().convert(sys.argv[idx + 1])
            text = normalize_markdown_output(result.text_content or "")
            Path(log_path).write_text("self-test-ok %d\n" % len(text), encoding="utf-8")
            print("self-test-ok %d" % len(text))
            sys.exit(0 if text else 3)
        except Exception as exc:
            Path(log_path).write_text(
                "self-test-failed %s: %s\n\n%s" % (
                    type(exc).__name__,
                    exc,
                    traceback.format_exc(),
                ),
                encoding="utf-8",
            )
            print("self-test-failed %s: %s" % (type(exc).__name__, exc), file=sys.stderr)
            sys.exit(1)

    app = QApplication(sys.argv)
    app.setApplicationName("MarkItDown")
    app.setOrganizationName("MarkItDown")

    icon_path = find_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
