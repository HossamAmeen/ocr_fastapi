import os
import sys
import traceback
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QProgressBar,
    QMessageBox, QFrame, QListWidget, QListWidgetItem, QStyle, QSizePolicy, QPlainTextEdit,
    QLineEdit, QRadioButton, QButtonGroup, QTabWidget, QScrollArea,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QSettings
from PyQt6.QtGui import QFont, QIcon, QColor

# Adjust system path to import app services correctly from root directory
sys.path.append(str(Path(__file__).resolve().parent))

from app.config import APP_VERSION
from app.services.combined_service import process_combined
from app.services.soe_service import parse_table_names
from excel_soe.writer import read_template_rig


class WorkerThread(QThread):
    """
    Background worker thread to run the OCR extraction and workbook generation.
    Prevents the main GUI thread from freezing during heavy operations.
    """
    finished_signal = pyqtSignal(dict, str)      # Emits (results_dict, output_path_str)
    error_signal = pyqtSignal(str)               # Emits (error_message)
    progress_signal = pyqtSignal(int, int, str)  # Emits (current_step, total_steps, message)

    def __init__(
        self,
        excel_path: Path,
        proforma_path: Path | None,
        soe_entries: list[tuple[Path, str]] | None,
        job_order_path: Path | None,
        job_order_source: str,
        job_order_start_marker: str = "",
        job_order_end_marker: str = "",
        soe_table_names: list[str] | None = None,
        soe_is_summary: bool = False
    ):
        super().__init__()
        self.excel_path = excel_path
        self.proforma_path = proforma_path
        self.soe_entries = soe_entries
        self.job_order_path = job_order_path
        self.job_order_source = job_order_source
        self.job_order_start_marker = job_order_start_marker
        self.job_order_end_marker = job_order_end_marker
        self.soe_table_names = soe_table_names
        self.soe_is_summary = soe_is_summary
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def is_cancelled(self) -> bool:
        return self._cancelled

    def _on_progress(self, current: int, total: int, msg: str):
        self.progress_signal.emit(current, total, msg)

    def run(self):
        try:
            result_data, output_path = process_combined(
                self.excel_path,
                proforma_pdf=self.proforma_path,
                soe_pdfs=self.soe_entries,
                job_order_pdf=self.job_order_path,
                job_order_source=self.job_order_source,
                job_order_start_marker=self.job_order_start_marker,
                job_order_end_marker=self.job_order_end_marker,
                soe_table_names=self.soe_table_names,
                soe_is_summary=self.soe_is_summary,
                progress_callback=self._on_progress,
                is_cancelled=self.is_cancelled,
            )
            self.finished_signal.emit(result_data, str(output_path))
        except InterruptedError as exc:
            self.error_signal.emit(str(exc))
        except Exception as exc:
            traceback.print_exc()
            self.error_signal.emit(str(exc))


class DeletableListWidget(QListWidget):
    """QListWidget that supports deleting selected items with Delete/Backspace key."""
    def __init__(self, parent=None, on_delete_requested=None):
        super().__init__(parent)
        self.on_delete_requested = on_delete_requested

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            if self.selectedItems() and self.on_delete_requested:
                self.on_delete_requested(self)
                return
        super().keyPressEvent(event)


class MainWindow(QMainWindow):
    _TEMPLATE_HINTS = {
        "auto": "Auto-detect scans the PDF and picks the first matching built-in template.",
        "1c": 'Start: "1-C Upper Completion - Running Procedure". End: "1-D Additional Information".',
        "running": 'Start: "Running Completion" (line start). End: "32 Perform final tests of TR-SCSSSV".',
        "completion_procedure": 'Start: "Completion Procedure". End: "CT OPERATION FOR MILLING GLASS DISC" (or end of PDF).',
        "custom": "Enter the exact start phrase from your PDF below. Optionally add an end phrase.",
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"OCR Excel Generator v{APP_VERSION} - Desktop")
        self.setMinimumSize(1020, 720)
        self.resize(1140, 820)

        # File states
        self.excel_file_path: Path | None = None
        self.proforma_file_path: Path | None = None
        self.job_order_file_path: Path | None = None
        self.selected_soe_files: list[tuple[Path, str]] = []

        # Widget registry for multi-tab synchronization
        self.excel_filename_labels: list[QLabel] = []
        self.excel_clear_buttons: list[QPushButton] = []
        self.excel_choose_buttons: list[QPushButton] = []

        self.proforma_filename_labels: list[QLabel] = []
        self.proforma_clear_buttons: list[QPushButton] = []
        self.proforma_choose_buttons: list[QPushButton] = []

        self.job_filename_labels: list[QLabel] = []
        self.job_clear_buttons: list[QPushButton] = []
        self.job_choose_buttons: list[QPushButton] = []
        self.job_combo_boxes: list[QComboBox] = []
        self.job_hint_labels: list[QLabel] = []
        self.job_custom_widgets: list[QWidget] = []
        self.job_start_inputs: list[QLineEdit] = []
        self.job_end_inputs: list[QLineEdit] = []

        self.soe_list_widgets: list[DeletableListWidget] = []
        self.soe_count_badges: list[QLabel] = []
        self.soe_choose_files_buttons: list[QPushButton] = []
        self.soe_choose_folder_buttons: list[QPushButton] = []
        self.soe_remove_buttons: list[QPushButton] = []
        self.soe_clear_buttons: list[QPushButton] = []
        self.soe_radio_tables: list[QRadioButton] = []
        self.soe_radio_summaries: list[QRadioButton] = []
        self.soe_table_names_inputs: list[QPlainTextEdit] = []
        self.soe_table_names_labels: list[QLabel] = []
        self.soe_table_names_hint_labels: list[QLabel] = []

        self.generated_output_path: Path | None = None
        self.last_results_dict: dict | None = None

        self.setAcceptDrops(True)
        self.setup_ui()
        self.apply_styles()

    def _get_last_dir(self, key: str) -> str:
        settings = QSettings("OCRGenerator", "OCRExcelApp")
        return str(settings.value(key, ""))

    def _set_last_dir(self, key: str, path_str: str):
        settings = QSettings("OCRGenerator", "OCRExcelApp")
        p = Path(path_str)
        settings.setValue(key, str(p.parent if p.is_file() else p))

    # =========================================================================
    # DRAG & DROP SUPPORT
    # =========================================================================
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return

        paths = [Path(url.toLocalFile()) for url in urls if url.isLocalFile()]
        if not paths:
            return

        excel_files = [p for p in paths if p.suffix.lower() in ('.xlsm', '.xlsx') and p.is_file()]
        pdf_files = [p for p in paths if p.suffix.lower() == '.pdf' and p.is_file()]
        directories = [p for p in paths if p.is_dir()]

        for d in directories:
            folder_pdfs = sorted(
                [p for p in d.glob("*.pdf") if p.is_file()],
                key=lambda x: x.name.lower()
            )
            pdf_files.extend(folder_pdfs)

        # 1. Handle Excel template
        if excel_files:
            self.excel_file_path = excel_files[0]
            self._set_last_dir("last_excel_dir", str(self.excel_file_path))
            self._update_excel_ui()
            self.show_status(f"✓ Template set via drag & drop: {self.excel_file_path.name}", "info")

        # 2. Handle PDF files
        if pdf_files:
            current_tab_idx = self.tabs.currentIndex()
            if current_tab_idx == 3 and len(pdf_files) == 1:
                # Proforma tab
                self.proforma_file_path = pdf_files[0]
                self._set_last_dir("last_proforma_dir", str(self.proforma_file_path))
                self._update_proforma_ui()
                self.show_status(f"✓ Proforma PDF set: {self.proforma_file_path.name}", "info")
            elif current_tab_idx == 2 and len(pdf_files) == 1:
                # Job Order tab
                self.job_order_file_path = pdf_files[0]
                self._set_last_dir("last_job_dir", str(self.job_order_file_path))
                self._update_job_ui()
                self.show_status(f"✓ Job Order PDF set: {self.job_order_file_path.name}", "info")
            else:
                # Smart detection on Combined tab or SOE tab
                name_lower = pdf_files[0].name.lower()
                if len(pdf_files) == 1 and ("proforma" in name_lower or "po" in name_lower):
                    self.proforma_file_path = pdf_files[0]
                    self._set_last_dir("last_proforma_dir", str(self.proforma_file_path))
                    self._update_proforma_ui()
                    self.show_status(f"✓ Proforma PDF set: {self.proforma_file_path.name}", "info")
                elif len(pdf_files) == 1 and ("job" in name_lower or "swi" in name_lower or "procedure" in name_lower):
                    self.job_order_file_path = pdf_files[0]
                    self._set_last_dir("last_job_dir", str(self.job_order_file_path))
                    self._update_job_ui()
                    self.show_status(f"✓ Job Order PDF set: {self.job_order_file_path.name}", "info")
                else:
                    self.add_soe_paths(pdf_files)
                    if pdf_files:
                        self._set_last_dir("last_soe_dir", str(pdf_files[0]))
                    self.show_status(f"✓ Added {len(pdf_files)} PDF file(s) via drag & drop.", "info")

        event.acceptProposedAction()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(12)

        # --- TOP HEADER ---
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(4, 0, 4, 4)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)

        eyebrow = QLabel("OCR + EXCEL WORKBOOK GENERATION")
        eyebrow.setObjectName("eyebrow")
        title_vbox.addWidget(eyebrow)

        title = QLabel(f"Excel Workbook Generator")
        title.setObjectName("main-title")
        title_vbox.addWidget(title)

        subtitle = QLabel(
            "Extract Proforma, SOE time logs (filtered by Rig), and Job Order procedures into your Excel template. Drag & drop files anytime."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        title_vbox.addWidget(subtitle)

        header_layout.addLayout(title_vbox)
        header_layout.addStretch()

        version_badge = QLabel(f"v{APP_VERSION}")
        version_badge.setObjectName("version-badge")
        header_layout.addWidget(version_badge, alignment=Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)

        main_layout.addWidget(header_widget, stretch=0)

        # --- TAB WIDGET ---
        self.tabs = QTabWidget()
        self.tabs.setObjectName("main-tabs")

        # Tab 1: Combined Generator
        tab_combined = self._create_combined_tab()
        self.tabs.addTab(tab_combined, "⚡ Combined Generator")

        # Tab 2: SOE Only
        tab_soe = self._create_soe_tab()
        self.tabs.addTab(tab_soe, "📊 SOE Time Log Only")

        # Tab 3: Job Order Only
        tab_job = self._create_job_order_tab()
        self.tabs.addTab(tab_job, "📋 Job Order Only")

        # Tab 4: Proforma Only
        tab_proforma = self._create_proforma_tab()
        self.tabs.addTab(tab_proforma, "📑 Proforma Only")

        # Tab 5: Results & Activity
        tab_results = self._create_results_tab()
        self.tabs.addTab(tab_results, "📈 Results & Activity")

        main_layout.addWidget(self.tabs, stretch=1)

        # --- GLOBAL STATUS, PROGRESS BAR & CANCEL ACTION ---
        bottom_bar = QWidget()
        bottom_layout = QVBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(4, 0, 4, 0)
        bottom_layout.setSpacing(6)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setObjectName("global-progress")
        bottom_layout.addWidget(self.progress_bar)

        status_action_layout = QHBoxLayout()
        status_action_layout.setContentsMargins(0, 0, 0, 0)

        self.lbl_status = QLabel("")
        self.lbl_status.setObjectName("status-label")
        self.lbl_status.setVisible(False)
        self.lbl_status.setWordWrap(True)
        status_action_layout.addWidget(self.lbl_status, stretch=1)

        self.btn_cancel = QPushButton("🛑 Cancel Operation")
        self.btn_cancel.setObjectName("cancel-btn")
        self.btn_cancel.setVisible(False)
        self.btn_cancel.clicked.connect(self.cancel_generation)
        status_action_layout.addWidget(self.btn_cancel)

        bottom_layout.addLayout(status_action_layout)
        main_layout.addWidget(bottom_bar, stretch=0)

        # Initial UI synchronization across all tabs
        self._update_excel_ui()
        self._update_proforma_ui()
        self._update_job_ui()
        self._update_soe_ui()
        self._on_soe_mode_changed()
        self._on_job_template_changed()

    # =========================================================================
    # TAB 1: COMBINED GENERATOR
    # =========================================================================
    def _create_combined_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setObjectName("tab-scroll-area")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(14)

        cols_widget = QWidget()
        cols_layout = QHBoxLayout(cols_widget)
        cols_layout.setContentsMargins(0, 0, 0, 0)
        cols_layout.setSpacing(16)

        # --- LEFT COLUMN (Excel, Proforma, Job Order) ---
        left_col = QWidget()
        left_layout = QVBoxLayout(left_col)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(12)

        excel_card = self._build_excel_card()
        left_layout.addWidget(excel_card)

        proforma_card = self._build_proforma_card()
        left_layout.addWidget(proforma_card)

        job_card = self._build_job_order_card()
        left_layout.addWidget(job_card)

        left_layout.addStretch()

        # --- RIGHT COLUMN (SOE Time Log) ---
        right_col = QWidget()
        right_layout = QVBoxLayout(right_col)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(12)

        soe_card = self._build_soe_card()
        right_layout.addWidget(soe_card)

        cols_layout.addWidget(left_col, stretch=1)
        cols_layout.addWidget(right_col, stretch=1)
        layout.addWidget(cols_widget)

        # Generate Button
        self.btn_generate_combined = QPushButton("🚀 Generate Combined Workbook")
        self.btn_generate_combined.setObjectName("generate-btn")
        self.btn_generate_combined.clicked.connect(lambda: self.start_generation())
        layout.addWidget(self.btn_generate_combined)

        scroll.setWidget(container)
        return scroll

    # =========================================================================
    # TAB 2: SOE TIME LOG ONLY
    # =========================================================================
    def _create_soe_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        info_lbl = QLabel(
            "<b>SOE Time Log Extraction:</b> Extracts operational time entries from daily report PDFs "
            "and appends them to the <i>SOE</i> sheet. If a Rig name is configured in the Excel template, "
            "non-matching PDFs are skipped automatically."
        )
        info_lbl.setWordWrap(True)
        info_lbl.setObjectName("tab-info-label")
        layout.addWidget(info_lbl)

        excel_reminder = self._build_excel_reminder_box()
        layout.addWidget(excel_reminder)

        soe_box = self._build_soe_card()
        layout.addWidget(soe_box)

        btn_gen_soe = QPushButton("📊 Generate SOE Workbook")
        btn_gen_soe.setObjectName("generate-btn")
        btn_gen_soe.clicked.connect(lambda: self.start_single_generation(target="soe"))
        self.btn_gen_soe = btn_gen_soe
        layout.addWidget(btn_gen_soe)

        layout.addStretch()
        scroll.setWidget(container)
        return scroll

    # =========================================================================
    # TAB 3: JOB ORDER ONLY
    # =========================================================================
    def _create_job_order_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        info_lbl = QLabel(
            "<b>Job Order Procedure Extraction:</b> Extracts step-by-step procedures from completion PDFs "
            "and appends them to the <i>JOB ORDER</i> sheet with live sequential formulas and Abadi 12 styling."
        )
        info_lbl.setWordWrap(True)
        info_lbl.setObjectName("tab-info-label")
        layout.addWidget(info_lbl)

        excel_reminder = self._build_excel_reminder_box()
        layout.addWidget(excel_reminder)

        job_box = self._build_job_order_card()
        layout.addWidget(job_box)

        btn_gen_job = QPushButton("📋 Generate Job Order Workbook")
        btn_gen_job.setObjectName("generate-btn")
        btn_gen_job.clicked.connect(lambda: self.start_single_generation(target="job"))
        self.btn_gen_job = btn_gen_job
        layout.addWidget(btn_gen_job)

        layout.addStretch()
        scroll.setWidget(container)
        return scroll

    # =========================================================================
    # TAB 4: PROFORMA ONLY
    # =========================================================================
    def _create_proforma_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        info_lbl = QLabel(
            "<b>Proforma Purchase Order Extraction:</b> Extracts line items, rates, and day quantities "
            "from Proforma PO PDFs and appends them to the <i>Proforma</i> sheet continuing from row 8."
        )
        info_lbl.setWordWrap(True)
        info_lbl.setObjectName("tab-info-label")
        layout.addWidget(info_lbl)

        excel_reminder = self._build_excel_reminder_box()
        layout.addWidget(excel_reminder)

        prof_box = self._build_proforma_card()
        layout.addWidget(prof_box)

        btn_gen_prof = QPushButton("📑 Generate Proforma Workbook")
        btn_gen_prof.setObjectName("generate-btn")
        btn_gen_prof.clicked.connect(lambda: self.start_single_generation(target="proforma"))
        self.btn_gen_prof = btn_gen_prof
        layout.addWidget(btn_gen_prof)

        layout.addStretch()
        scroll.setWidget(container)
        return scroll

    # =========================================================================
    # TAB 5: RESULTS & ACTIVITY
    # =========================================================================
    def _create_results_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        # Top Action Bar
        action_bar = QHBoxLayout()
        self.lbl_results_heading = QLabel("No workbook generated yet.")
        self.lbl_results_heading.setObjectName("results-title")
        action_bar.addWidget(self.lbl_results_heading)
        action_bar.addStretch()

        self.btn_open_file = QPushButton("📂 Open Excel Workbook")
        self.btn_open_file.setObjectName("open-file-btn")
        self.btn_open_file.setEnabled(False)
        self.btn_open_file.clicked.connect(self.open_generated_file)

        self.btn_open_folder = QPushButton("📁 Open Output Folder")
        self.btn_open_folder.setObjectName("secondary-btn")
        self.btn_open_folder.setEnabled(False)
        self.btn_open_folder.clicked.connect(self.open_generated_folder)

        self.btn_copy_summary = QPushButton("📋 Copy Summary")
        self.btn_copy_summary.setObjectName("secondary-btn")
        self.btn_copy_summary.setEnabled(False)
        self.btn_copy_summary.clicked.connect(self.copy_summary_text)

        action_bar.addWidget(self.btn_open_file)
        action_bar.addWidget(self.btn_open_folder)
        action_bar.addWidget(self.btn_copy_summary)
        layout.addLayout(action_bar)

        # Summary Metric Cards Container
        self.summary_cards_widget = QWidget()
        summary_cards_layout = QHBoxLayout(self.summary_cards_widget)
        summary_cards_layout.setContentsMargins(0, 0, 0, 0)
        summary_cards_layout.setSpacing(12)

        self.card_summary_proforma = self._create_metric_card("Proforma", "Not processed", "#38bdf8")
        self.card_summary_soe = self._create_metric_card("SOE Time Log", "Not processed", "#10b981")
        self.card_summary_job = self._create_metric_card("Job Order", "Not processed", "#a855f7")

        summary_cards_layout.addWidget(self.card_summary_proforma)
        summary_cards_layout.addWidget(self.card_summary_soe)
        summary_cards_layout.addWidget(self.card_summary_job)
        layout.addWidget(self.summary_cards_widget)

        # SOE Detailed PDF Summary Table (QTableWidget)
        soe_table_lbl = QLabel("SOE PDF Processing Breakdown:")
        soe_table_lbl.setObjectName("section-title")
        layout.addWidget(soe_table_lbl)

        self.table_soe_results = QTableWidget(0, 6)
        self.table_soe_results.setObjectName("results-table")
        self.table_soe_results.setHorizontalHeaderLabels([
            "PDF Filename", "Source", "Well / Date", "Period", "Rig", "Rows / Status"
        ])
        self.table_soe_results.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_soe_results.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table_soe_results.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table_soe_results.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table_soe_results.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.table_soe_results.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.table_soe_results.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_soe_results.setAlternatingRowColors(True)
        self.table_soe_results.setMinimumHeight(220)
        layout.addWidget(self.table_soe_results)

        # Text Summary Box for detailed logs
        details_lbl = QLabel("Activity Log & Extraction Details:")
        details_lbl.setObjectName("section-title")
        layout.addWidget(details_lbl)

        self.txt_results_log = QPlainTextEdit()
        self.txt_results_log.setReadOnly(True)
        self.txt_results_log.setMinimumHeight(150)
        self.txt_results_log.setObjectName("txt-results-log")
        layout.addWidget(self.txt_results_log)

        scroll.setWidget(container)
        return scroll

    def _create_metric_card(self, title: str, initial_val: str, accent_color: str) -> QFrame:
        card = QFrame()
        card.setObjectName("metric-card")
        card.setStyleSheet(f"border-top: 3px solid {accent_color};")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)
        card_layout.setSpacing(4)

        t_lbl = QLabel(title)
        t_lbl.setObjectName("metric-title")
        v_lbl = QLabel(initial_val)
        v_lbl.setObjectName("metric-value")
        v_lbl.setWordWrap(True)

        card_layout.addWidget(t_lbl)
        card_layout.addWidget(v_lbl)
        card.value_label = v_lbl
        return card

    # =========================================================================
    # REUSABLE CARD BUILDERS
    # =========================================================================
    def _build_excel_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("section-card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 14, 14, 14)
        card_layout.setSpacing(8)

        hdr = QHBoxLayout()
        title = QLabel("Excel Template")
        title.setObjectName("section-title")
        badge = QLabel("Required")
        badge.setObjectName("badge-required")
        hdr.addWidget(title)
        hdr.addWidget(badge)
        hdr.addStretch()
        card_layout.addLayout(hdr)

        hint = QLabel("The master .xlsm template with Proforma, SOE, and JOB ORDER sheets.")
        hint.setObjectName("field-hint")
        card_layout.addWidget(hint)

        action_layout = QHBoxLayout()
        btn_choose = QPushButton("📁 Choose Excel Template")
        btn_choose.clicked.connect(self.choose_excel_file)
        self.excel_choose_buttons.append(btn_choose)

        lbl_filename = QLabel("No template selected")
        lbl_filename.setObjectName("filename-label")
        lbl_filename.setWordWrap(True)
        self.excel_filename_labels.append(lbl_filename)

        btn_clear = QPushButton("✕")
        btn_clear.setObjectName("btn-mini-clear")
        btn_clear.setToolTip("Clear Excel template selection")
        btn_clear.setVisible(False)
        btn_clear.clicked.connect(self.clear_excel_file)
        self.excel_clear_buttons.append(btn_clear)

        action_layout.addWidget(btn_choose)
        action_layout.addWidget(btn_clear)
        action_layout.addWidget(lbl_filename, stretch=1)
        card_layout.addLayout(action_layout)

        return card

    def _build_excel_reminder_box(self) -> QFrame:
        box = QFrame()
        box.setObjectName("reminder-box")
        layout = QHBoxLayout(box)
        layout.setContentsMargins(12, 10, 12, 10)

        icon = QLabel("📌")
        icon.setStyleSheet("font-size: 16px;")
        layout.addWidget(icon)

        lbl_filename = QLabel("No template selected")
        lbl_filename.setObjectName("filename-label")
        lbl_filename.setWordWrap(True)
        self.excel_filename_labels.append(lbl_filename)
        layout.addWidget(lbl_filename, stretch=1)

        btn_pick = QPushButton("Select Template")
        btn_pick.setObjectName("secondary-btn")
        btn_pick.clicked.connect(self.choose_excel_file)
        self.excel_choose_buttons.append(btn_pick)
        layout.addWidget(btn_pick)

        return box

    def _build_proforma_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("section-card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 14, 14, 14)
        card_layout.setSpacing(8)

        hdr = QHBoxLayout()
        title = QLabel("Proforma Purchase Order")
        title.setObjectName("section-title")
        badge = QLabel("Optional")
        badge.setObjectName("badge-optional")
        hdr.addWidget(title)
        hdr.addWidget(badge)
        hdr.addStretch()
        card_layout.addLayout(hdr)

        hint = QLabel("Purchase order PDF to extract into the Proforma sheet (row 8+).")
        hint.setObjectName("field-hint")
        card_layout.addWidget(hint)

        action_layout = QHBoxLayout()
        btn_choose = QPushButton("📑 Choose Proforma PDF")
        btn_choose.clicked.connect(self.choose_proforma_file)
        self.proforma_choose_buttons.append(btn_choose)

        lbl_filename = QLabel("No file selected")
        lbl_filename.setObjectName("filename-label")
        lbl_filename.setWordWrap(True)
        self.proforma_filename_labels.append(lbl_filename)

        btn_clear = QPushButton("✕")
        btn_clear.setObjectName("btn-mini-clear")
        btn_clear.setToolTip("Clear Proforma selection")
        btn_clear.setVisible(False)
        btn_clear.clicked.connect(self.clear_proforma_file)
        self.proforma_clear_buttons.append(btn_clear)

        action_layout.addWidget(btn_choose)
        action_layout.addWidget(btn_clear)
        action_layout.addWidget(lbl_filename, stretch=1)
        card_layout.addLayout(action_layout)

        return card

    def _build_job_order_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("section-card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 14, 14, 14)
        card_layout.setSpacing(8)

        hdr = QHBoxLayout()
        title = QLabel("Job Order Procedure")
        title.setObjectName("section-title")
        badge = QLabel("Optional")
        badge.setObjectName("badge-optional")
        hdr.addWidget(title)
        hdr.addWidget(badge)
        hdr.addStretch()
        card_layout.addLayout(hdr)

        hint = QLabel("Completion procedure PDF for the JOB ORDER sheet.")
        hint.setObjectName("field-hint")
        card_layout.addWidget(hint)

        action_layout = QHBoxLayout()
        btn_choose = QPushButton("🛠 Choose Job Order PDF")
        btn_choose.clicked.connect(self.choose_job_file)
        self.job_choose_buttons.append(btn_choose)

        lbl_filename = QLabel("No file selected")
        lbl_filename.setObjectName("filename-label")
        lbl_filename.setWordWrap(True)
        self.job_filename_labels.append(lbl_filename)

        btn_clear = QPushButton("✕")
        btn_clear.setObjectName("btn-mini-clear")
        btn_clear.setToolTip("Clear Job Order selection")
        btn_clear.setVisible(False)
        btn_clear.clicked.connect(self.clear_job_file)
        self.job_clear_buttons.append(btn_clear)

        action_layout.addWidget(btn_choose)
        action_layout.addWidget(btn_clear)
        action_layout.addWidget(lbl_filename, stretch=1)
        card_layout.addLayout(action_layout)

        # Template selector
        template_layout = QHBoxLayout()
        template_lbl = QLabel("Extraction template:")
        template_lbl.setObjectName("template-label")
        combo = QComboBox()
        combo.addItem("Auto-detect", "auto")
        combo.addItem("1-C procedure (detailed lines)", "1c")
        combo.addItem("Running completion", "running")
        combo.addItem("Completion Procedure (SWI)", "completion_procedure")
        combo.addItem("Custom start/end text", "custom")
        self.job_combo_boxes.append(combo)
        template_layout.addWidget(template_lbl)
        template_layout.addWidget(combo, stretch=1)
        card_layout.addLayout(template_layout)

        # Template hint label
        lbl_hint = QLabel("")
        lbl_hint.setObjectName("field-hint")
        lbl_hint.setWordWrap(True)
        self.job_hint_labels.append(lbl_hint)
        card_layout.addWidget(lbl_hint)

        # Custom markers widget
        custom_markers_widget = QWidget()
        custom_markers_widget.setVisible(False)
        self.job_custom_widgets.append(custom_markers_widget)
        custom_markers_layout = QVBoxLayout(custom_markers_widget)
        custom_markers_layout.setContentsMargins(0, 4, 0, 0)
        custom_markers_layout.setSpacing(6)

        start_lbl = QLabel("Start text (required):")
        start_lbl.setObjectName("template-label")
        txt_start = QLineEdit()
        txt_start.setObjectName("marker-input")
        txt_start.setPlaceholderText("e.g. COMPLETION PROCEDURE")
        self.job_start_inputs.append(txt_start)

        end_lbl = QLabel("End text (optional):")
        end_lbl.setObjectName("template-label")
        txt_end = QLineEdit()
        txt_end.setObjectName("marker-input")
        txt_end.setPlaceholderText("e.g. APPENDIX A")
        self.job_end_inputs.append(txt_end)

        custom_markers_layout.addWidget(start_lbl)
        custom_markers_layout.addWidget(txt_start)
        custom_markers_layout.addWidget(end_lbl)
        custom_markers_layout.addWidget(txt_end)

        card_layout.addWidget(custom_markers_widget)

        combo.currentIndexChanged.connect(lambda idx, c=combo: self._on_job_template_changed(c))
        txt_start.textChanged.connect(lambda text, inp=txt_start: self._on_job_start_marker_changed(inp))
        txt_end.textChanged.connect(lambda text, inp=txt_end: self._on_job_end_marker_changed(inp))

        return card

    def _build_soe_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("section-card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 14, 14, 14)
        card_layout.setSpacing(8)

        hdr = QHBoxLayout()
        title = QLabel("SOE Time Log Reports")
        title.setObjectName("section-title")
        badge = QLabel("Optional")
        badge.setObjectName("badge-optional")

        lbl_count = QLabel("0 files")
        lbl_count.setObjectName("badge-count")
        self.soe_count_badges.append(lbl_count)

        hdr.addWidget(title)
        hdr.addWidget(badge)
        hdr.addWidget(lbl_count)
        hdr.addStretch()
        card_layout.addLayout(hdr)

        hint = QLabel(
            "Upload daily operations report PDFs or a folder. "
            "Matches Rig if defined in the template. OAMN markers are stripped automatically."
        )
        hint.setObjectName("field-hint")
        hint.setWordWrap(True)
        card_layout.addWidget(hint)

        action_layout = QHBoxLayout()
        btn_choose_files = QPushButton("📚 Select PDF Files")
        btn_choose_files.clicked.connect(self.choose_soe_files)
        self.soe_choose_files_buttons.append(btn_choose_files)

        btn_choose_folder = QPushButton("📁 Select Folder")
        btn_choose_folder.clicked.connect(self.choose_soe_folder)
        btn_choose_folder.setObjectName("secondary-btn")
        self.soe_choose_folder_buttons.append(btn_choose_folder)

        btn_remove_item = QPushButton("🗑 Remove Selected")
        btn_remove_item.setObjectName("secondary-btn")
        self.soe_remove_buttons.append(btn_remove_item)

        btn_clear = QPushButton("✕ Clear All")
        btn_clear.setObjectName("secondary-btn")
        btn_clear.clicked.connect(self.clear_soe_files)
        self.soe_clear_buttons.append(btn_clear)

        action_layout.addWidget(btn_choose_files)
        action_layout.addWidget(btn_choose_folder)
        action_layout.addWidget(btn_remove_item)
        action_layout.addWidget(btn_clear)
        action_layout.addStretch()
        card_layout.addLayout(action_layout)

        # File list with Delete key support
        list_widget = DeletableListWidget(on_delete_requested=self.remove_selected_soe)
        list_widget.setObjectName("list-soe-files")
        list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        list_widget.setMinimumHeight(120)
        self.soe_list_widgets.append(list_widget)
        btn_remove_item.clicked.connect(lambda checked, lw=list_widget: self.remove_selected_soe(lw))
        card_layout.addWidget(list_widget)

        # Extraction mode: Table vs Summary
        mode_box = QHBoxLayout()
        mode_lbl = QLabel("Extraction mode:")
        mode_lbl.setObjectName("template-label")
        radio_table = QRadioButton("Table (Multiple Rows)")
        radio_summary = QRadioButton("Summary (Single Paragraph Row)")
        radio_table.setChecked(True)

        mode_group = QButtonGroup(card)
        mode_group.addButton(radio_table)
        mode_group.addButton(radio_summary)
        self.soe_radio_tables.append(radio_table)
        self.soe_radio_summaries.append(radio_summary)

        mode_box.addWidget(mode_lbl)
        mode_box.addWidget(radio_table)
        mode_box.addWidget(radio_summary)
        mode_box.addStretch()
        card_layout.addLayout(mode_box)

        # Table names / Paragraph title
        lbl_names = QLabel("Table names to extract:")
        lbl_names.setObjectName("template-label")
        self.soe_table_names_labels.append(lbl_names)

        txt_names = QPlainTextEdit()
        txt_names.setPlaceholderText("Time Log\nJob Time Log\nOperational Time Summary")
        txt_names.setPlainText("Time Log\nJob Time Log\nOperational Time Summary")
        txt_names.setMaximumHeight(85)
        txt_names.setObjectName("txt-table-names")
        self.soe_table_names_inputs.append(txt_names)

        lbl_names_hint = QLabel("One table name per line. Only matching tables are extracted.")
        lbl_names_hint.setObjectName("field-hint")
        self.soe_table_names_hint_labels.append(lbl_names_hint)

        event_merge_hint = QLabel("💡 Each SOE row's Event cell automatically spans columns E–R in the Excel workbook.")
        event_merge_hint.setObjectName("field-hint-accent")

        card_layout.addWidget(lbl_names)
        card_layout.addWidget(txt_names)
        card_layout.addWidget(lbl_names_hint)
        card_layout.addWidget(event_merge_hint)

        radio_table.toggled.connect(lambda checked, r=radio_table: self._on_soe_mode_changed(r))
        radio_summary.toggled.connect(lambda checked, r=radio_summary: self._on_soe_mode_changed(r))
        txt_names.textChanged.connect(lambda txt=txt_names: self._on_soe_table_names_changed(txt))

        return card

    # =========================================================================
    # STYLESHEET (QSS)
    # =========================================================================
    def apply_styles(self):
        qss = """
        QWidget {
            background-color: #0b0f19;
            color: #cbd5e1;
            font-family: 'Segoe UI', 'Inter', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            font-size: 13px;
        }

        #tab-scroll-area {
            background-color: transparent;
            border: none;
        }

        #eyebrow {
            color: #38bdf8;
            font-weight: 800;
            font-size: 11px;
            letter-spacing: 1.5px;
            text-transform: uppercase;
        }

        #main-title {
            color: #f8fafc;
            font-size: 22px;
            font-weight: 700;
        }

        #subtitle {
            color: #94a3b8;
            font-size: 13px;
            line-height: 1.4;
        }

        #version-badge {
            background-color: #1e293b;
            color: #38bdf8;
            font-size: 11px;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 6px;
            border: 1px solid #334155;
        }

        /* Tabs */
        QTabWidget::pane {
            border: 1px solid #232f48;
            background-color: #111726;
            border-radius: 8px;
            top: -1px;
        }

        QTabBar::tab {
            background-color: #161f30;
            color: #94a3b8;
            padding: 10px 18px;
            font-weight: 600;
            font-size: 13px;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
            margin-right: 4px;
            border: 1px solid #232f48;
            border-bottom: none;
        }

        QTabBar::tab:selected {
            background-color: #111726;
            color: #38bdf8;
            border-top: 2px solid #0ea5e9;
            font-weight: 700;
        }

        QTabBar::tab:hover:!selected {
            background-color: #1e293b;
            color: #f8fafc;
        }

        /* Section Cards */
        #section-card {
            background-color: #161f30;
            border: 1px solid #27354f;
            border-radius: 8px;
        }

        #reminder-box {
            background-color: #0f172a;
            border: 1px dashed #334155;
            border-radius: 6px;
        }

        #tab-info-label {
            background-color: #0f172a;
            border-left: 3px solid #0ea5e9;
            padding: 8px 12px;
            border-radius: 4px;
            color: #cbd5e1;
            font-size: 13px;
            line-height: 1.4;
        }

        #section-title {
            color: #f1f5f9;
            font-size: 14px;
            font-weight: 700;
        }

        #badge-required {
            background-color: #4c1d1d;
            color: #fca5a5;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 4px;
            border: 1px solid #ef4444;
        }

        #badge-optional {
            background-color: #1e293b;
            color: #94a3b8;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 4px;
            border: 1px solid #475569;
        }

        #badge-count {
            background-color: #0284c7;
            color: #ffffff;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 10px;
        }

        #field-hint {
            color: #94a3b8;
            font-size: 12px;
        }

        #field-hint-accent {
            color: #38bdf8;
            font-size: 12px;
            font-style: italic;
        }

        #filename-label {
            color: #94a3b8;
            font-size: 12px;
            padding-left: 4px;
        }

        /* Buttons */
        QPushButton {
            background-color: #2563eb;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: 600;
            font-size: 13px;
            min-height: 20px;
        }

        QPushButton:hover {
            background-color: #3b82f6;
        }

        QPushButton:pressed {
            background-color: #1d4ed8;
        }

        #secondary-btn {
            background-color: #27354f;
            color: #cbd5e1;
        }

        #secondary-btn:hover {
            background-color: #334466;
            color: #ffffff;
        }

        #secondary-btn:pressed {
            background-color: #1e293b;
        }

        #btn-mini-clear {
            background-color: #475569;
            color: #f8fafc;
            border-radius: 4px;
            padding: 2px 8px;
            font-weight: bold;
            font-size: 11px;
            min-height: 18px;
        }

        #btn-mini-clear:hover {
            background-color: #ef4444;
        }

        #generate-btn {
            background-color: #0ea5e9;
            color: #ffffff;
            font-size: 15px;
            padding: 12px 24px;
            border-radius: 8px;
            font-weight: 700;
            min-height: 24px;
        }

        #generate-btn:hover {
            background-color: #38bdf8;
        }

        #generate-btn:pressed {
            background-color: #0284c7;
        }

        #cancel-btn {
            background-color: #dc2626;
            color: #ffffff;
            font-size: 13px;
            padding: 6px 14px;
            border-radius: 6px;
            font-weight: 700;
        }

        #cancel-btn:hover {
            background-color: #ef4444;
        }

        #cancel-btn:pressed {
            background-color: #b91c1c;
        }

        #generate-btn:disabled {
            background-color: #1e293b;
            color: #64748b;
            border: 1px solid #334155;
        }

        /* Inputs */
        QComboBox {
            background-color: #0c101a;
            border: 1px solid #27354f;
            border-radius: 6px;
            padding: 7px 12px;
            color: #f1f5f9;
            min-width: 220px;
        }

        QComboBox::drop-down {
            border: none;
            width: 24px;
        }

        QComboBox QAbstractItemView {
            background-color: #0c101a;
            border: 1px solid #27354f;
            color: #f1f5f9;
            selection-background-color: #1e293b;
            selection-color: #38bdf8;
        }

        QListWidget {
            background-color: #0c101a;
            border: 1px solid #27354f;
            border-radius: 6px;
            color: #e2e8f0;
            padding: 6px;
        }

        QListWidget::item {
            padding: 6px 10px;
            border-bottom: 1px solid #161f30;
            border-radius: 4px;
        }

        QListWidget::item:hover {
            background-color: #1e293b;
            color: #38bdf8;
        }

        QListWidget::item:selected {
            background-color: #0284c7;
            color: #ffffff;
        }

        QPlainTextEdit, QLineEdit {
            background-color: #0c101a;
            border: 1px solid #27354f;
            border-radius: 6px;
            color: #f1f5f9;
            padding: 8px 10px;
            font-size: 13px;
        }

        QPlainTextEdit:focus, QLineEdit:focus {
            border: 1px solid #0ea5e9;
        }

        #template-label {
            color: #cbd5e1;
            font-weight: 600;
        }

        /* Progress Bar */
        #global-progress {
            border: 1px solid #27354f;
            border-radius: 6px;
            text-align: center;
            background-color: #0c101a;
            height: 20px;
            color: #ffffff;
            font-weight: bold;
        }

        #global-progress::chunk {
            background-color: #0ea5e9;
            border-radius: 5px;
        }

        /* Status Banner */
        #status-label {
            font-size: 13px;
            font-weight: 600;
            padding: 10px 14px;
            border-radius: 6px;
        }

        .status-info {
            background-color: #082f49;
            color: #38bdf8;
            border: 1px solid #0ea5e9;
        }

        .status-error {
            background-color: #451a1a;
            color: #fca5a5;
            border: 1px solid #ef4444;
        }

        .status-success {
            background-color: #064e3b;
            color: #6ee7b7;
            border: 1px solid #10b981;
        }

        /* Metric Cards */
        #metric-card {
            background-color: #161f30;
            border: 1px solid #27354f;
            border-radius: 8px;
        }

        #metric-title {
            color: #94a3b8;
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        #metric-value {
            color: #f8fafc;
            font-size: 14px;
            font-weight: 600;
        }

        #results-title {
            color: #f8fafc;
            font-size: 18px;
            font-weight: 700;
        }

        #open-file-btn {
            background-color: #10b981;
            font-weight: 700;
        }

        #open-file-btn:hover {
            background-color: #059669;
        }

        #open-file-btn:pressed {
            background-color: #047857;
        }

        /* Table */
        #results-table {
            background-color: #0c101a;
            border: 1px solid #27354f;
            border-radius: 6px;
            gridline-color: #1e293b;
            color: #f1f5f9;
        }

        #results-table QHeaderView::section {
            background-color: #1e293b;
            color: #f8fafc;
            font-weight: 700;
            padding: 8px 10px;
            border: none;
            border-right: 1px solid #27354f;
            border-bottom: 2px solid #0ea5e9;
        }

        #results-table::item {
            padding: 6px 10px;
        }

        #results-table::item:alternate {
            background-color: #111726;
        }

        #results-table::item:selected {
            background-color: #0284c7;
            color: #ffffff;
        }

        QScrollBar:vertical {
            border: none;
            background: #0b0f19;
            width: 10px;
            margin: 0px;
        }

        QScrollBar::handle:vertical {
            background: #27354f;
            min-height: 24px;
            border-radius: 5px;
        }

        QScrollBar::handle:vertical:hover {
            background: #38bdf8;
        }

        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        """
        self.setStyleSheet(qss)

    # =========================================================================
    # FILE PICKER & MANAGEMENT HANDLERS
    # =========================================================================
    def _update_excel_ui(self):
        rig_name = None
        if self.excel_file_path and self.excel_file_path.is_file():
            try:
                rig_name = read_template_rig(self.excel_file_path)
            except Exception:
                rig_name = None

        for lbl in self.excel_filename_labels:
            if self.excel_file_path:
                rig_badge = f" [Rig: {rig_name}]" if rig_name else " [Rig: None (All)]"
                lbl.setText(f"✓ {self.excel_file_path.name}{rig_badge}")
                lbl.setToolTip(str(self.excel_file_path.resolve()))
                lbl.setStyleSheet("color: #38bdf8; font-weight: 700;")
            else:
                lbl.setText("No template selected")
                lbl.setToolTip("")
                lbl.setStyleSheet("")
        for btn in self.excel_clear_buttons:
            btn.setVisible(bool(self.excel_file_path))

    def choose_excel_file(self):
        start_dir = self._get_last_dir("last_excel_dir")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel Template",
            start_dir,
            "Excel Files (*.xlsm *.xlsx)"
        )
        if file_path:
            self.excel_file_path = Path(file_path)
            self._set_last_dir("last_excel_dir", file_path)
            self._update_excel_ui()

    def clear_excel_file(self):
        self.excel_file_path = None
        self._update_excel_ui()

    def _update_proforma_ui(self):
        for lbl in self.proforma_filename_labels:
            if self.proforma_file_path:
                lbl.setText(f"✓ {self.proforma_file_path.name}")
                lbl.setToolTip(str(self.proforma_file_path.resolve()))
                lbl.setStyleSheet("color: #38bdf8; font-weight: 700;")
            else:
                lbl.setText("No file selected")
                lbl.setToolTip("")
                lbl.setStyleSheet("")
        for btn in self.proforma_clear_buttons:
            btn.setVisible(bool(self.proforma_file_path))

    def choose_proforma_file(self):
        start_dir = self._get_last_dir("last_proforma_dir") or self._get_last_dir("last_pdf_dir")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Proforma PDF",
            start_dir,
            "PDF Files (*.pdf)"
        )
        if file_path:
            self.proforma_file_path = Path(file_path)
            self._set_last_dir("last_proforma_dir", file_path)
            self._set_last_dir("last_pdf_dir", file_path)
            self._update_proforma_ui()

    def clear_proforma_file(self):
        self.proforma_file_path = None
        self._update_proforma_ui()

    def _update_job_ui(self):
        for lbl in self.job_filename_labels:
            if self.job_order_file_path:
                lbl.setText(f"✓ {self.job_order_file_path.name}")
                lbl.setToolTip(str(self.job_order_file_path.resolve()))
                lbl.setStyleSheet("color: #38bdf8; font-weight: 700;")
            else:
                lbl.setText("No file selected")
                lbl.setToolTip("")
                lbl.setStyleSheet("")
        for btn in self.job_clear_buttons:
            btn.setVisible(bool(self.job_order_file_path))

    def choose_job_file(self):
        start_dir = self._get_last_dir("last_job_dir") or self._get_last_dir("last_pdf_dir")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Job Order PDF",
            start_dir,
            "PDF Files (*.pdf)"
        )
        if file_path:
            self.job_order_file_path = Path(file_path)
            self._set_last_dir("last_job_dir", file_path)
            self._set_last_dir("last_pdf_dir", file_path)
            self._update_job_ui()

    def clear_job_file(self):
        self.job_order_file_path = None
        self._update_job_ui()

    def choose_soe_files(self):
        start_dir = self._get_last_dir("last_soe_dir") or self._get_last_dir("last_pdf_dir")
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select SOE PDF Files",
            start_dir,
            "PDF Files (*.pdf)"
        )
        if file_paths:
            self._set_last_dir("last_soe_dir", file_paths[0])
            self._set_last_dir("last_pdf_dir", file_paths[0])
            self.add_soe_paths([Path(p) for p in file_paths])
            self.show_status(f"Added {len(file_paths)} SOE PDF file(s).", "info")

    def choose_soe_folder(self):
        start_dir = self._get_last_dir("last_soe_dir")
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Folder with SOE Reports",
            start_dir
        )
        if folder_path:
            self._set_last_dir("last_soe_dir", folder_path)
            folder = Path(folder_path)
            pdf_paths = sorted(
                [p for p in folder.glob("*.pdf")],
                key=lambda x: x.name.lower()
            )
            if not pdf_paths:
                QMessageBox.warning(
                    self,
                    "No PDFs Found",
                    "The selected folder does not contain any PDF files."
                )
                return
            self.add_soe_paths(pdf_paths)
            self.show_status(f"Added {len(pdf_paths)} SOE PDF file(s) from folder.", "info")

    def add_soe_paths(self, paths: list[Path]):
        existing_paths = {p for p, _ in self.selected_soe_files}
        for p in paths:
            if p not in existing_paths:
                self.selected_soe_files.append((p, p.name))
                existing_paths.add(p)
        self._update_soe_ui()

    def remove_selected_soe(self, source_list_widget=None):
        items_to_remove = set()
        if source_list_widget and hasattr(source_list_widget, "selectedItems"):
            for item in source_list_widget.selectedItems():
                items_to_remove.add(item.text())
        else:
            for list_w in self.soe_list_widgets:
                for item in list_w.selectedItems():
                    items_to_remove.add(item.text())

        if not items_to_remove:
            return

        self.selected_soe_files = [
            (p, name) for (p, name) in self.selected_soe_files
            if name not in items_to_remove
        ]
        self._update_soe_ui()

    def clear_soe_files(self):
        self.selected_soe_files = []
        self._update_soe_ui()

    def _update_soe_ui(self):
        count = len(self.selected_soe_files)
        badge_text = f"{count} file{'s' if count != 1 else ''}"
        for badge in self.soe_count_badges:
            badge.setText(badge_text)

        for list_w in self.soe_list_widgets:
            list_w.blockSignals(True)
            list_w.clear()
            for p, name in self.selected_soe_files:
                item = QListWidgetItem(name)
                item.setToolTip(str(p.resolve()))
                list_w.addItem(item)
            list_w.blockSignals(False)

    def _on_soe_mode_changed(self, source_radio=None):
        is_summary = False
        if source_radio and hasattr(source_radio, "isChecked"):
            if source_radio in self.soe_radio_summaries:
                is_summary = source_radio.isChecked()
            elif source_radio in self.soe_radio_tables:
                is_summary = not source_radio.isChecked()
        elif self.soe_radio_summaries:
            is_summary = any(r.isChecked() for r in self.soe_radio_summaries)

        for r_tbl in self.soe_radio_tables:
            r_tbl.blockSignals(True)
            r_tbl.setChecked(not is_summary)
            r_tbl.blockSignals(False)
        for r_sum in self.soe_radio_summaries:
            r_sum.blockSignals(True)
            r_sum.setChecked(is_summary)
            r_sum.blockSignals(False)

        for lbl, hint in zip(self.soe_table_names_labels, self.soe_table_names_hint_labels):
            if is_summary:
                lbl.setText("Paragraph title to extract:")
                hint.setText(
                    "Exact paragraph heading in the PDF. Its text is captured as one consolidated Excel row."
                )
            else:
                lbl.setText("Table names to extract:")
                hint.setText(
                    "One table name per line. Only matching tables are extracted into rows."
                )

        for txt in self.soe_table_names_inputs:
            current = txt.toPlainText().strip()
            if is_summary:
                txt.setPlaceholderText("24 Hrs Summary")
                if "Time Log" in current or "Job Time Log" in current or "Operational Time Summary" in current:
                    txt.blockSignals(True)
                    txt.setPlainText("24 Hrs Summary")
                    txt.blockSignals(False)
            else:
                txt.setPlaceholderText("Time Log\nJob Time Log\nOperational Time Summary")
                if current in ("24 Hrs Summary", "24 Hours Summary"):
                    txt.blockSignals(True)
                    txt.setPlainText("Time Log\nJob Time Log\nOperational Time Summary")
                    txt.blockSignals(False)

    def _on_soe_table_names_changed(self, source_txt=None):
        if not source_txt:
            return
        text = source_txt.toPlainText()
        for txt in self.soe_table_names_inputs:
            if txt is not source_txt:
                txt.blockSignals(True)
                txt.setPlainText(text)
                txt.blockSignals(False)

    def _on_job_template_changed(self, source_combo=None):
        source = source_combo.currentData() if source_combo else (
            self.job_combo_boxes[0].currentData() if self.job_combo_boxes else "auto"
        )
        hint_text = self._TEMPLATE_HINTS.get(source, "")
        is_custom = (source == "custom")

        for combo in self.job_combo_boxes:
            if combo is not source_combo:
                combo.blockSignals(True)
                idx = combo.findData(source)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
                combo.blockSignals(False)

        for lbl in self.job_hint_labels:
            lbl.setText(hint_text)

        for widget in self.job_custom_widgets:
            widget.setVisible(is_custom)

    def _on_job_start_marker_changed(self, source_input=None):
        if not source_input:
            return
        text = source_input.text()
        for inp in self.job_start_inputs:
            if inp is not source_input:
                inp.blockSignals(True)
                inp.setText(text)
                inp.blockSignals(False)

    def _on_job_end_marker_changed(self, source_input=None):
        if not source_input:
            return
        text = source_input.text()
        for inp in self.job_end_inputs:
            if inp is not source_input:
                inp.blockSignals(True)
                inp.setText(text)
                inp.blockSignals(False)

    # =========================================================================
    # GENERATION LOGIC
    # =========================================================================
    def show_status(self, message: str, status_type: str = "info"):
        self.lbl_status.setText(message)
        self.lbl_status.setVisible(True)
        if status_type == "info":
            self.lbl_status.setProperty("class", "status-info")
        elif status_type == "success":
            self.lbl_status.setProperty("class", "status-success")
        else:
            self.lbl_status.setProperty("class", "status-error")
        self.lbl_status.style().polish(self.lbl_status)

    def hide_status(self):
        self.lbl_status.setVisible(False)

    def start_single_generation(self, target: str):
        """Helper to run extraction for a single specific section."""
        if target == "soe":
            self.start_generation(force_single="soe")
        elif target == "job":
            self.start_generation(force_single="job")
        elif target == "proforma":
            self.start_generation(force_single="proforma")

    def cancel_generation(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.btn_cancel.setEnabled(False)
            self.show_status("Cancelling operation, please wait...", "error")
            self.worker.cancel()

    def on_generation_progress(self, current: int, total: int, message: str):
        if total > 0:
            percent = int((current / total) * 100)
            self.progress_bar.setValue(percent)
            self.show_status(f"[{percent}%] {message}", "info")
        else:
            self.show_status(message, "info")

    def start_generation(self, force_single: str | None = None):
        if not isinstance(force_single, str) or force_single not in ("soe", "job", "proforma"):
            force_single = None
        self.hide_status()

        if not self.excel_file_path:
            self.show_status("Excel template file (.xlsm) is required.", "error")
            QMessageBox.warning(self, "Missing Template", "Please select an Excel template (.xlsm) file first.")
            return

        # Pre-check: Ensure Excel template is readable and not locked by another process
        try:
            with open(self.excel_file_path, "r+b") as test_f:
                pass
        except PermissionError:
            self.show_status("Excel template is locked! Please close it in Microsoft Excel.", "error")
            QMessageBox.critical(
                self,
                "File Is Locked",
                f"The Excel file '{self.excel_file_path.name}' is currently open in Microsoft Excel or another program.\n\nPlease close it before generating."
            )
            return
        except Exception:
            pass

        # Determine active files based on mode
        proforma_path = self.proforma_file_path if (force_single in (None, "proforma")) else None
        soe_entries = self.selected_soe_files if (force_single in (None, "soe") and len(self.selected_soe_files) > 0) else None
        job_order_path = self.job_order_file_path if (force_single in (None, "job")) else None

        if not proforma_path and not soe_entries and not job_order_path:
            msg = "Please select at least one PDF section to process."
            if force_single == "soe":
                msg = "Please select at least one SOE PDF file or folder."
            elif force_single == "job":
                msg = "Please select a Job Order PDF file."
            elif force_single == "proforma":
                msg = "Please select a Proforma PDF file."
            self.show_status(msg, "error")
            QMessageBox.warning(self, "No PDFs Selected", msg)
            return

        job_order_source = self.job_combo_boxes[0].currentData() if self.job_combo_boxes else "auto"
        job_order_start_marker = self.job_start_inputs[0].text().strip() if self.job_start_inputs else ""
        job_order_end_marker = self.job_end_inputs[0].text().strip() if self.job_end_inputs else ""

        if job_order_path and job_order_source == "custom" and not job_order_start_marker:
            self.show_status("Start text is required when using the Custom Job Order template.", "error")
            return

        # Prepare SOE parameters
        soe_table_names_text = self.soe_table_names_inputs[0].toPlainText() if self.soe_table_names_inputs else ""
        soe_table_names_list = [line.strip() for line in soe_table_names_text.split('\n') if line.strip()]
        parsed_table_names = parse_table_names(soe_table_names_list) if soe_entries else None
        soe_is_summary = self.soe_radio_summaries[0].isChecked() if self.soe_radio_summaries else False

        # Update UI state for processing
        self.set_controls_enabled(False)
        self.btn_cancel.setVisible(True)
        self.btn_cancel.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.show_status("Extracting PDF contents and generating Excel workbook...", "info")

        # Spawn background worker thread
        self.worker = WorkerThread(
            excel_path=self.excel_file_path,
            proforma_path=proforma_path,
            soe_entries=soe_entries,
            job_order_path=job_order_path,
            job_order_source=job_order_source,
            job_order_start_marker=job_order_start_marker,
            job_order_end_marker=job_order_end_marker,
            soe_table_names=parsed_table_names,
            soe_is_summary=soe_is_summary
        )
        self.worker.finished_signal.connect(self.on_generation_finished)
        self.worker.error_signal.connect(self.on_generation_error)
        self.worker.progress_signal.connect(self.on_generation_progress)
        self.worker.start()

    def set_controls_enabled(self, enabled: bool):
        self.btn_generate_combined.setEnabled(enabled)
        if hasattr(self, 'btn_gen_soe'):
            self.btn_gen_soe.setEnabled(enabled)
        if hasattr(self, 'btn_gen_job'):
            self.btn_gen_job.setEnabled(enabled)
        if hasattr(self, 'btn_gen_prof'):
            self.btn_gen_prof.setEnabled(enabled)

        for btn in self.excel_choose_buttons + self.excel_clear_buttons:
            btn.setEnabled(enabled)
        for btn in self.proforma_choose_buttons + self.proforma_clear_buttons:
            btn.setEnabled(enabled)
        for btn in self.job_choose_buttons + self.job_clear_buttons:
            btn.setEnabled(enabled)
        for btn in self.soe_choose_files_buttons:
            btn.setEnabled(enabled)
        for btn in self.soe_choose_folder_buttons:
            btn.setEnabled(enabled)
        for btn in self.soe_remove_buttons:
            btn.setEnabled(enabled)
        for btn in self.soe_clear_buttons:
            btn.setEnabled(enabled)
        for combo in self.job_combo_boxes:
            combo.setEnabled(enabled)
        for txt in self.soe_table_names_inputs:
            txt.setEnabled(enabled)
        for r in self.soe_radio_tables:
            r.setEnabled(enabled)
        for r in self.soe_radio_summaries:
            r.setEnabled(enabled)
        for inp in self.job_start_inputs:
            inp.setEnabled(enabled)
        for inp in self.job_end_inputs:
            inp.setEnabled(enabled)

    def on_generation_finished(self, results: dict, output_path: str):
        self.set_controls_enabled(True)
        self.btn_cancel.setVisible(False)
        self.progress_bar.setVisible(False)
        self.generated_output_path = Path(output_path)
        self.last_results_dict = results

        self.show_status("✓ Workbook generated successfully!", "success")

        # Update Results & Activity Tab
        self.lbl_results_heading.setText(f"✓ Output Ready: {self.generated_output_path.name}")
        self.btn_open_file.setEnabled(True)
        self.btn_open_folder.setEnabled(True)
        self.btn_copy_summary.setEnabled(True)

        log_lines = []
        log_lines.append(f"Output File: {self.generated_output_path.resolve()}")
        sections = results.get("processed_sections", [])
        log_lines.append(f"Processed Sections: {', '.join(sections)}")

        # 1. Proforma Metric
        if "proforma" in results and results["proforma"]:
            prof = results["proforma"]
            prof_txt = f"{prof['item_count']} items (${prof['gross_total']:,.2f})"
            self.card_summary_proforma.value_label.setText(prof_txt)
            log_lines.append(f"\n[PROFORMA]\n• Extracted {prof['item_count']} items. Total Amount: ${prof['gross_total']:,.2f}")
        else:
            self.card_summary_proforma.value_label.setText("Skipped / Not selected")

        # 2. SOE Metric & Table
        self.table_soe_results.setRowCount(0)
        if "soe" in results and results["soe"]:
            soe = results["soe"]
            rig_str = f" (Rig Filter: '{soe.get('rig_filter')}')" if soe.get('rig_filter') else ""
            self.card_summary_soe.value_label.setText(f"{soe['row_count']} rows from {soe['pdf_count']} PDF(s){rig_str}")
            log_lines.append(f"\n[SOE TIME LOG]\n• Processed {soe['row_count']} rows across {soe['pdf_count']} PDF reports{rig_str}")

            summaries = soe.get("pdf_summaries", [])
            self.table_soe_results.setRowCount(len(summaries))
            for row_idx, sm in enumerate(summaries):
                fn = sm.get("filename", "")
                src = sm.get("source", "")
                src_disp = (
                    "Time Log" if src == "time_log"
                    else "Op Time Summary" if src == "operational_time_summary"
                    else "Summary (paragraph)" if src == "paragraph_summary"
                    else src
                )
                well_date = sm.get("report_date", "") if src == "operational_time_summary" else sm.get("well_name", "")
                well_date = well_date or "-"

                p_from = sm.get("report_period_from", "") or ""
                p_to = sm.get("report_period_to", "") or ""
                period = f"{p_from} – {p_to}" if (p_from and p_to and p_from != p_to) else (p_from or "-")
                rig = sm.get("rig", "") or "-"

                if sm.get("skipped"):
                    r = sm.get("skip_reason", "")
                    status = (
                        "Skipped (rig mismatch)" if r == "rig_mismatch"
                        else "Skipped (no table)" if r == "no_matching_table"
                        else "Skipped (empty table)" if r == "empty_table"
                        else "Skipped"
                    )
                else:
                    status = f"{sm.get('row_count', 0)} rows"

                item_fn = QTableWidgetItem(fn)
                item_src = QTableWidgetItem(src_disp)
                item_wd = QTableWidgetItem(well_date)
                item_per = QTableWidgetItem(period)
                item_rig = QTableWidgetItem(rig)
                item_st = QTableWidgetItem(status)

                if sm.get("skipped"):
                    item_st.setForeground(QColor("#fca5a5"))
                else:
                    item_st.setForeground(QColor("#6ee7b7"))

                self.table_soe_results.setItem(row_idx, 0, item_fn)
                self.table_soe_results.setItem(row_idx, 1, item_src)
                self.table_soe_results.setItem(row_idx, 2, item_wd)
                self.table_soe_results.setItem(row_idx, 3, item_per)
                self.table_soe_results.setItem(row_idx, 4, item_rig)
                self.table_soe_results.setItem(row_idx, 5, item_st)
        else:
            self.card_summary_soe.value_label.setText("Skipped / Not selected")

        # 3. Job Order Metric
        if "job_order" in results and results["job_order"]:
            jo = results["job_order"]
            lines_data = jo.get("lines", [])
            steps = sum(1 for ln in lines_data if ln.get("kind") == "step")
            others = jo["line_count"] - steps
            detail = f"{steps} step{'s' if steps != 1 else ''}"
            if others > 0:
                detail += f", {others} text/table rows"
            self.card_summary_job.value_label.setText(f"{jo['line_count']} rows ({detail})")
            log_lines.append(f"\n[JOB ORDER]\n• Appended {jo['line_count']} rows ({detail}) using '{jo['source']}' template.")
        else:
            self.card_summary_job.value_label.setText("Skipped / Not selected")

        self.txt_results_log.setPlainText("\n".join(log_lines))

        # Switch to Results Tab (Index 4)
        self.tabs.setCurrentIndex(4)

        QMessageBox.information(
            self,
            "Generation Complete",
            f"Excel workbook generated successfully!\n\nSaved to: {self.generated_output_path.name}"
        )

    def on_generation_error(self, err_msg: str):
        self.set_controls_enabled(True)
        self.btn_cancel.setVisible(False)
        self.progress_bar.setVisible(False)

        if "cancelled" in err_msg.lower():
            self.show_status("Operation was cancelled by user.", "info")
            return

        first_line = err_msg.split('\n')[0] if err_msg else "Workbook generation failed."
        self.show_status(f"Error: {first_line}", "error")

        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setWindowTitle("Extraction & Generation Error")
        msg_box.setText(err_msg)
        msg_box.exec()

    # =========================================================================
    # ACTIONS
    # =========================================================================
    def open_generated_file(self):
        if not self.generated_output_path or not self.generated_output_path.is_file():
            QMessageBox.warning(self, "File Not Found", "The generated file could not be found.")
            return
        self.launch_path(self.generated_output_path)

    def open_generated_folder(self):
        if not self.generated_output_path:
            QMessageBox.warning(self, "Directory Not Found", "No file has been generated yet.")
            return
        self.launch_path(self.generated_output_path.parent)

    def copy_summary_text(self):
        summary_text = self.txt_results_log.toPlainText()
        if summary_text:
            QApplication.clipboard().setText(summary_text)
            self.show_status("✓ Summary copied to clipboard!", "success")

    def launch_path(self, path: Path):
        """Cross-platform path launcher."""
        target_str = str(path.resolve())
        if sys.platform == 'win32':
            os.startfile(target_str)
        elif sys.platform == 'darwin':
            import subprocess
            subprocess.call(['open', target_str])
        else:
            import subprocess
            subprocess.call(['xdg-open', target_str])


if __name__ == "__main__":
    # High-DPI Scaling configuration for modern crisp rendering
    if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    
    if hasattr(Qt.HighDpiScaleFactorRoundingPolicy, "PassThrough"):
        app.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    window = MainWindow()

    try:
        app_icon = window.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)
        window.setWindowIcon(app_icon)
    except Exception:
        pass

    window.show()
    sys.exit(app.exec())
