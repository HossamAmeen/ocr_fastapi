import os
import sys
import traceback
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QProgressBar,
    QMessageBox, QFrame, QListWidget, QStyle, QSizePolicy, QPlainTextEdit,
    QLineEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QIcon

# Adjust system path to import app services correctly from root directory
sys.path.append(str(Path(__file__).resolve().parent))

from app.config import APP_VERSION
from app.services.combined_service import process_combined
from app.services.soe_service import parse_table_names


class WorkerThread(QThread):
    """
    Background worker thread to run the OCR extraction and workbook generation.
    Prevents the main GUI thread from freezing during heavy operations.
    """
    finished_signal = pyqtSignal(dict, str)  # Emits (results_dict, output_path_str)
    error_signal = pyqtSignal(str)           # Emits (error_message)

    def __init__(
        self,
        excel_path: Path,
        proforma_path: Path | None,
        soe_entries: list[tuple[Path, str]] | None,
        job_order_path: Path | None,
        job_order_source: str,
        job_order_start_marker: str = "",
        job_order_end_marker: str = "",
        soe_table_names: list[str] | None = None
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
            )
            self.finished_signal.emit(result_data, str(output_path))
        except Exception as exc:
            # Capture the traceback for console logging
            traceback.print_exc()
            self.error_signal.emit(str(exc))


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
        self.setMinimumSize(900, 680)
        self.resize(1020, 750)

        # File states
        self.excel_file_path = None
        self.proforma_file_path = None
        self.job_order_file_path = None
        self.selected_soe_files = []  # List of tuples (Path, display_name)

        self.setup_ui()
        self.apply_styles()

    def setup_ui(self):
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # --- HEADER SECTION ---
        header_widget = QWidget()
        header_layout = QVBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 5)
        header_layout.setSpacing(4)

        eyebrow = QLabel("OCR + EXCEL")
        eyebrow.setObjectName("eyebrow")
        header_layout.addWidget(eyebrow)

        title = QLabel(f"Excel Workbook Generator  v{APP_VERSION}")
        title.setObjectName("main-title")
        header_layout.addWidget(title)

        subtitle = QLabel(
            "Upload one Excel template and any combination of PDFs. Proforma, SOE time logs\n"
            "(filtered by Rig if Rig is set in template), and Job Order procedures are extracted."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        header_layout.addWidget(subtitle)

        main_layout.addWidget(header_widget, stretch=0)

        # --- COLUMNS CONTAINER (Form Layout) ---
        columns_widget = QWidget()
        columns_layout = QHBoxLayout(columns_widget)
        columns_layout.setContentsMargins(0, 0, 0, 0)
        columns_layout.setSpacing(15)

        # --- LEFT COLUMN (Excel, Proforma, Job Order) ---
        left_column = QWidget()
        left_layout = QVBoxLayout(left_column)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(12)

        # 1. Excel Template Card (REQUIRED)
        excel_card = QFrame()
        excel_card.setObjectName("excel-card")
        excel_card.setFrameShape(QFrame.Shape.StyledPanel)
        excel_layout = QVBoxLayout(excel_card)
        excel_layout.setSpacing(8)
        
        excel_hdr = QHBoxLayout()
        excel_title = QLabel("Excel Template")
        excel_title.setObjectName("section-title")
        excel_badge = QLabel("Required")
        excel_badge.setObjectName("badge-required")
        excel_hdr.addWidget(excel_title)
        excel_hdr.addWidget(excel_badge)
        excel_hdr.addStretch()
        excel_layout.addLayout(excel_hdr)

        excel_hint = QLabel("One .xlsm template used for all selected sections.")
        excel_hint.setObjectName("field-hint")
        excel_layout.addWidget(excel_hint)

        excel_action_layout = QHBoxLayout()
        self.btn_choose_excel = QPushButton("Choose Excel template")
        self.btn_choose_excel.clicked.connect(self.choose_excel_file)
        self.lbl_excel_filename = QLabel("No file selected")
        self.lbl_excel_filename.setObjectName("filename-label")
        excel_action_layout.addWidget(self.btn_choose_excel)
        excel_action_layout.addWidget(self.lbl_excel_filename)
        excel_action_layout.addStretch()
        excel_layout.addLayout(excel_action_layout)
        
        left_layout.addWidget(excel_card)

        # 2. Proforma Card (OPTIONAL)
        proforma_card = QFrame()
        proforma_card.setObjectName("proforma-card")
        proforma_card.setFrameShape(QFrame.Shape.StyledPanel)
        proforma_layout = QVBoxLayout(proforma_card)
        proforma_layout.setSpacing(8)

        proforma_hdr = QHBoxLayout()
        proforma_title = QLabel("Proforma")
        proforma_title.setObjectName("section-title")
        proforma_badge = QLabel("Optional")
        proforma_badge.setObjectName("badge-optional")
        proforma_hdr.addWidget(proforma_title)
        proforma_hdr.addWidget(proforma_badge)
        proforma_hdr.addStretch()
        proforma_layout.addLayout(proforma_hdr)

        proforma_hint = QLabel("Purchase order PDF for the Proforma sheet.")
        proforma_hint.setObjectName("field-hint")
        proforma_layout.addWidget(proforma_hint)

        proforma_action_layout = QHBoxLayout()
        self.btn_choose_proforma = QPushButton("Choose Proforma PDF")
        self.btn_choose_proforma.clicked.connect(self.choose_proforma_file)
        self.lbl_proforma_filename = QLabel("No file selected")
        self.lbl_proforma_filename.setObjectName("filename-label")
        proforma_action_layout.addWidget(self.btn_choose_proforma)
        proforma_action_layout.addWidget(self.lbl_proforma_filename)
        proforma_action_layout.addStretch()
        proforma_layout.addLayout(proforma_action_layout)

        left_layout.addWidget(proforma_card)

        # 4. Job Order Card (OPTIONAL)
        job_card = QFrame()
        job_card.setObjectName("job-card")
        job_card.setFrameShape(QFrame.Shape.StyledPanel)
        job_layout = QVBoxLayout(job_card)
        job_layout.setSpacing(8)

        job_hdr = QHBoxLayout()
        job_title = QLabel("Job Order")
        job_title.setObjectName("section-title")
        job_badge = QLabel("Optional")
        job_badge.setObjectName("badge-optional")
        job_hdr.addWidget(job_title)
        job_hdr.addWidget(job_badge)
        job_hdr.addStretch()
        job_layout.addLayout(job_hdr)

        job_hint = QLabel("Completion program PDF for the JOB ORDER sheet.")
        job_hint.setObjectName("field-hint")
        job_layout.addWidget(job_hint)

        job_action_layout = QHBoxLayout()
        self.btn_choose_job = QPushButton("Choose Job Order PDF")
        self.btn_choose_job.clicked.connect(self.choose_job_file)
        self.lbl_job_filename = QLabel("No file selected")
        self.lbl_job_filename.setObjectName("filename-label")
        job_action_layout.addWidget(self.btn_choose_job)
        job_action_layout.addWidget(self.lbl_job_filename)
        job_action_layout.addStretch()
        job_layout.addLayout(job_action_layout)

        template_layout = QHBoxLayout()
        template_lbl = QLabel("Extraction template:")
        template_lbl.setObjectName("template-label")
        self.combo_job_template = QComboBox()
        self.combo_job_template.addItem("Auto-detect", "auto")
        self.combo_job_template.addItem("1-C procedure (detailed lines)", "1c")
        self.combo_job_template.addItem("Running completion", "running")
        self.combo_job_template.addItem("Completion Procedure (SWI)", "completion_procedure")
        self.combo_job_template.addItem("Custom start/end text", "custom")
        template_layout.addWidget(template_lbl)
        template_layout.addWidget(self.combo_job_template)
        template_layout.addStretch()
        job_layout.addLayout(template_layout)

        # Template hint label
        self.lbl_job_template_hint = QLabel("")
        self.lbl_job_template_hint.setObjectName("field-hint")
        self.lbl_job_template_hint.setWordWrap(True)
        job_layout.addWidget(self.lbl_job_template_hint)

        # Custom markers widget
        self.custom_markers_widget = QWidget()
        self.custom_markers_widget.setVisible(False)
        self.custom_markers_widget.setMinimumHeight(125)
        custom_markers_layout = QVBoxLayout(self.custom_markers_widget)
        custom_markers_layout.setContentsMargins(0, 4, 0, 0)
        custom_markers_layout.setSpacing(6)

        start_lbl = QLabel("Start text (required):")
        start_lbl.setObjectName("template-label")
        self.txt_job_start_marker = QLineEdit()
        self.txt_job_start_marker.setObjectName("marker-input")
        self.txt_job_start_marker.setPlaceholderText("e.g. COMPLETION PROCEDURE")

        end_lbl = QLabel("End text (optional):")
        end_lbl.setObjectName("template-label")
        self.txt_job_end_marker = QLineEdit()
        self.txt_job_end_marker.setObjectName("marker-input")
        self.txt_job_end_marker.setPlaceholderText("e.g. APPENDIX A")

        custom_markers_layout.addWidget(start_lbl)
        custom_markers_layout.addWidget(self.txt_job_start_marker)
        custom_markers_layout.addWidget(end_lbl)
        custom_markers_layout.addWidget(self.txt_job_end_marker)

        job_layout.addWidget(self.custom_markers_widget)

        self.combo_job_template.currentIndexChanged.connect(self._on_job_template_changed)
        self._on_job_template_changed()

        left_layout.addWidget(job_card)
        left_layout.addStretch()  # Push cards to the top of left column

        # --- RIGHT COLUMN (SOE Time Log - OPTIONAL) ---
        right_column = QWidget()
        right_layout = QVBoxLayout(right_column)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        soe_card = QFrame()
        soe_card.setObjectName("soe-card")
        soe_card.setFrameShape(QFrame.Shape.StyledPanel)
        soe_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        soe_layout = QVBoxLayout(soe_card)
        soe_layout.setSpacing(8)

        soe_hdr = QHBoxLayout()
        soe_title = QLabel("SOE Time Log")
        soe_title.setObjectName("section-title")
        soe_badge = QLabel("Optional")
        soe_badge.setObjectName("badge-optional")
        soe_hdr.addWidget(soe_title)
        soe_hdr.addWidget(soe_badge)
        soe_hdr.addStretch()
        soe_layout.addLayout(soe_hdr)

        soe_hint = QLabel(
            "One or more SOE PDFs or folder. Matches Rig filter if set in the Excel template.\n"
            "OAMN markers and trailing off-duty text are stripped automatically from Operation Details."
        )
        soe_hint.setObjectName("field-hint")
        soe_layout.addWidget(soe_hint)

        soe_action_layout = QHBoxLayout()
        self.btn_choose_soe_files = QPushButton("Select PDF files")
        self.btn_choose_soe_files.clicked.connect(self.choose_soe_files)
        self.btn_choose_soe_folder = QPushButton("Select folder")
        self.btn_choose_soe_folder.clicked.connect(self.choose_soe_folder)
        self.btn_choose_soe_folder.setObjectName("secondary-btn")
        
        soe_action_layout.addWidget(self.btn_choose_soe_files)
        soe_action_layout.addWidget(self.btn_choose_soe_folder)
        soe_action_layout.addStretch()
        soe_layout.addLayout(soe_action_layout)

        # File list starts visible and occupies the remaining space in the SOE card
        self.list_soe_files = QListWidget()
        self.list_soe_files.setObjectName("list-soe-files")
        self.list_soe_files.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        soe_layout.addWidget(self.list_soe_files)

        # Table names section
        table_names_lbl = QLabel("Table names to extract:")
        table_names_lbl.setObjectName("template-label")
        self.txt_soe_table_names = QPlainTextEdit()
        self.txt_soe_table_names.setPlaceholderText("Time Log\nJob Time Log\nOperational Time Summary")
        self.txt_soe_table_names.setPlainText("Time Log\nJob Time Log\nOperational Time Summary")
        self.txt_soe_table_names.setMaximumHeight(95)
        self.txt_soe_table_names.setObjectName("txt-table-names")
        
        table_names_hint = QLabel("One table name per line. Only matching tables are extracted.")
        table_names_hint.setObjectName("field-hint")

        event_merge_hint = QLabel("Each SOE row's Event cell automatically spans columns E\u2013R in the output workbook.")
        event_merge_hint.setObjectName("field-hint")

        soe_layout.addWidget(table_names_lbl)
        soe_layout.addWidget(self.txt_soe_table_names)
        soe_layout.addWidget(table_names_hint)
        soe_layout.addWidget(event_merge_hint)

        right_layout.addWidget(soe_card)

        # Add left and right columns to columns container
        columns_layout.addWidget(left_column, stretch=1)
        columns_layout.addWidget(right_column, stretch=1)

        main_layout.addWidget(columns_widget, stretch=1)

        # --- BOTTOM SECTION ---
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(10)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        bottom_layout.addWidget(self.progress_bar)

        # Status label
        self.lbl_status = QLabel("")
        self.lbl_status.setObjectName("status-label")
        self.lbl_status.setVisible(False)
        self.lbl_status.setWordWrap(True)
        bottom_layout.addWidget(self.lbl_status)

        # Generate button
        self.btn_generate = QPushButton("Generate Workbook")
        self.btn_generate.setObjectName("generate-btn")
        self.btn_generate.clicked.connect(self.start_generation)
        bottom_layout.addWidget(self.btn_generate)

        # Results card
        self.results_card = QFrame()
        self.results_card.setObjectName("results-card")
        self.results_card.setFrameShape(QFrame.Shape.StyledPanel)
        self.results_card.setVisible(False)
        results_layout = QVBoxLayout(self.results_card)

        results_hdr = QLabel("Generation Successful")
        results_hdr.setObjectName("results-title")
        results_layout.addWidget(results_hdr)

        self.lbl_results_summary = QLabel("")
        self.lbl_results_summary.setObjectName("results-summary")
        self.lbl_results_summary.setWordWrap(True)
        self.lbl_results_summary.setTextFormat(Qt.TextFormat.RichText)
        results_layout.addWidget(self.lbl_results_summary)

        results_btn_layout = QHBoxLayout()
        self.btn_open_file = QPushButton("Open Excel Workbook")
        self.btn_open_file.setObjectName("open-file-btn")
        self.btn_open_file.clicked.connect(self.open_generated_file)
        
        self.btn_open_folder = QPushButton("Open Folder")
        self.btn_open_folder.setObjectName("secondary-btn")
        self.btn_open_folder.clicked.connect(self.open_generated_folder)
        
        results_btn_layout.addWidget(self.btn_open_file)
        results_btn_layout.addWidget(self.btn_open_folder)
        results_btn_layout.addStretch()
        results_layout.addLayout(results_btn_layout)

        bottom_layout.addWidget(self.results_card)

        main_layout.addWidget(bottom_widget, stretch=0)

        # Keep output path reference
        self.generated_output_path = None

    def apply_styles(self):
        # Modern CSS-like style sheet (QSS) targeting dark layout aesthetic
        qss = """
        QWidget {
            background-color: #0b0f19;
            color: #cbd5e1;
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            font-size: 13px;
        }

        #eyebrow {
            color: #0ea5e9;
            font-weight: 800;
            font-size: 10px;
            letter-spacing: 1.5px;
            text-transform: uppercase;
        }

        #main-title {
            color: #f8fafc;
            font-size: 24px;
            font-weight: 700;
        }

        #subtitle {
            color: #64748b;
            font-size: 13px;
            line-height: 1.4;
        }

        #excel-card, #proforma-card, #job-card, #soe-card {
            background-color: #161b26;
            border: 1px solid #273043;
            border-radius: 8px;
            padding: 14px;
        }

        #section-title {
            color: #f1f5f9;
            font-size: 14px;
            font-weight: 600;
        }

        #badge-required {
            background-color: #451a1a;
            color: #fca5a5;
            font-size: 9px;
            font-weight: bold;
            padding: 1px 5px;
            border-radius: 4px;
            border: 1px solid #ef4444;
        }

        #badge-optional {
            background-color: #1e293b;
            color: #94a3b8;
            font-size: 9px;
            font-weight: bold;
            padding: 1px 5px;
            border-radius: 4px;
            border: 1px solid #475569;
        }

        #field-hint {
            color: #64748b;
            font-size: 12px;
        }

        #filename-label {
            color: #64748b;
            font-style: italic;
            font-size: 12px;
            margin-left: 10px;
        }

        QPushButton {
            background-color: #2563eb;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            padding: 7px 14px;
            font-weight: 600;
        }

        QPushButton:hover {
            background-color: #3b82f6;
        }

        QPushButton:pressed {
            background-color: #1d4ed8;
        }

        #secondary-btn {
            background-color: #334155;
        }

        #secondary-btn:hover {
            background-color: #475569;
        }

        #secondary-btn:pressed {
            background-color: #1e293b;
        }

        #generate-btn {
            background-color: #0ea5e9;
            color: #ffffff;
            font-size: 14px;
            padding: 11px 20px;
            border-radius: 6px;
            font-weight: 700;
        }

        #generate-btn:hover {
            background-color: #38bdf8;
        }

        #generate-btn:pressed {
            background-color: #0284c7;
        }

        #generate-btn:disabled {
            background-color: #1e293b;
            color: #475569;
            border: 1px solid #273043;
        }

        QComboBox {
            background-color: #0f131c;
            border: 1px solid #273043;
            border-radius: 6px;
            padding: 5px 10px;
            color: #e2e8f0;
            min-width: 200px;
        }

        QComboBox::drop-down {
            border: none;
            width: 20px;
        }

        QComboBox QAbstractItemView {
            background-color: #0f131c;
            border: 1px solid #273043;
            color: #e2e8f0;
            selection-background-color: #1e293b;
            selection-color: #ffffff;
        }

        QListWidget {
            background-color: #0f131c;
            border: 1px solid #273043;
            border-radius: 6px;
            color: #e2e8f0;
            padding: 5px;
        }

        QPlainTextEdit {
            background-color: #0f131c;
            border: 1px solid #273043;
            border-radius: 6px;
            color: #e2e8f0;
            padding: 5px;
        }

        QLineEdit#marker-input {
            background-color: #0f131c;
            border: 1px solid #273043;
            border-radius: 6px;
            padding: 5px 10px;
            color: #e2e8f0;
            min-height: 28px;
        }

        QLineEdit#marker-input:focus {
            border: 1px solid #0ea5e9;
        }

        QListWidget::item {
            padding: 6px 10px;
            border-bottom: 1px solid #161b26;
            color: #cbd5e1;
        }

        QListWidget::item:hover {
            background-color: #1e293b;
            border-radius: 4px;
            color: #f8fafc;
        }

        #template-label {
            color: #94a3b8;
            font-weight: 600;
        }

        QProgressBar {
            border: 1px solid #273043;
            border-radius: 6px;
            text-align: center;
            background-color: #0f131c;
            height: 18px;
            color: #ffffff;
            font-weight: bold;
        }

        QProgressBar::chunk {
            background-color: #0ea5e9;
            border-radius: 5px;
        }

        #status-label {
            font-size: 12px;
            font-weight: 600;
            padding: 8px 12px;
            border-radius: 6px;
        }

        .status-info {
            background-color: #0c2540;
            color: #38bdf8;
            border: 1px solid #0ea5e9;
        }

        .status-error {
            background-color: #451a1a;
            color: #fca5a5;
            border: 1px solid #ef4444;
        }

        #results-card {
            background-color: #0c1c16;
            border: 1px solid #10b981;
            border-radius: 8px;
            padding: 14px;
        }

        #results-title {
            color: #10b981;
            font-size: 16px;
            font-weight: bold;
            margin-bottom: 4px;
        }

        #results-summary {
            color: #a7f3d0;
            font-size: 13px;
            line-height: 1.5;
            margin-bottom: 10px;
        }

        #open-file-btn {
            background-color: #10b981;
        }

        #open-file-btn:hover {
            background-color: #059669;
        }

        #open-file-btn:pressed {
            background-color: #047857;
        }

        QScrollBar:vertical {
            border: none;
            background: #0f131c;
            width: 8px;
            margin: 0px;
        }

        QScrollBar::handle:vertical {
            background: #334155;
            min-height: 20px;
            border-radius: 4px;
        }

        QScrollBar::handle:vertical:hover {
            background: #475569;
        }

        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        """
        self.setStyleSheet(qss)

    # --- FILE CHOOSERS ---

    def choose_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel Template",
            "",
            "Excel Files (*.xlsm *.xlsx)"
        )
        if file_path:
            self.excel_file_path = Path(file_path)
            self.lbl_excel_filename.setText(f"✓ {self.excel_file_path.name}")
            self.lbl_excel_filename.setStyleSheet("color: #10b981; font-weight: 600;")
        else:
            self.excel_file_path = None
            self.lbl_excel_filename.setText("No file selected")
            self.lbl_excel_filename.setStyleSheet("")

    def choose_proforma_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Proforma PDF",
            "",
            "PDF Files (*.pdf)"
        )
        if file_path:
            self.proforma_file_path = Path(file_path)
            self.lbl_proforma_filename.setText(f"✓ {self.proforma_file_path.name}")
            self.lbl_proforma_filename.setStyleSheet("color: #10b981; font-weight: 600;")
        else:
            self.proforma_file_path = None
            self.lbl_proforma_filename.setText("No file selected")
            self.lbl_proforma_filename.setStyleSheet("")

    def choose_soe_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select SOE PDF Files",
            "",
            "PDF Files (*.pdf)"
        )
        if file_paths:
            self.process_soe_paths([Path(p) for p in file_paths])
            self.show_status("Selected SOE files.", "info")

    def choose_soe_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Folder with SOE Reports"
        )
        if folder_path:
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
            self.process_soe_paths(pdf_paths)

    def process_soe_paths(self, paths: list[Path]):
        # Clear previous selection
        self.selected_soe_files = []
        self.list_soe_files.clear()

        for p in paths:
            # Entry: (file_path, display_name)
            self.selected_soe_files.append((p, p.name))
            self.list_soe_files.addItem(p.name)

    def choose_job_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Job Order PDF",
            "",
            "PDF Files (*.pdf)"
        )
        if file_path:
            self.job_order_file_path = Path(file_path)
            self.lbl_job_filename.setText(f"✓ {self.job_order_file_path.name}")
            self.lbl_job_filename.setStyleSheet("color: #10b981; font-weight: 600;")
        else:
            self.job_order_file_path = None
            self.lbl_job_filename.setText("No file selected")
            self.lbl_job_filename.setStyleSheet("")

    def _on_job_template_changed(self):
        source = self.combo_job_template.currentData()
        self.lbl_job_template_hint.setText(self._TEMPLATE_HINTS.get(source, ""))
        is_custom = source == "custom"
        self.custom_markers_widget.setVisible(is_custom)

    # --- OPERATION ---

    def show_status(self, message: str, status_type: str = "info"):
        self.lbl_status.setText(message)
        self.lbl_status.setVisible(True)
        if status_type == "info":
            self.lbl_status.setProperty("class", "status-info")
        else:
            self.lbl_status.setProperty("class", "status-error")
        # Refresh style
        self.lbl_status.style().polish(self.lbl_status)

    def hide_status(self):
        self.lbl_status.setVisible(False)

    def start_generation(self):
        self.hide_status()
        self.results_card.setVisible(False)

        # Validation
        if not self.excel_file_path:
            self.show_status("Excel template file is required.", "error")
            return

        has_proforma = self.proforma_file_path is not None
        has_soe = len(self.selected_soe_files) > 0
        has_job = self.job_order_file_path is not None

        if not has_proforma and not has_soe and not has_job:
            self.show_status(
                "Please select at least one PDF section to process: Proforma, SOE, or Job Order.",
                "error"
            )
            return

        job_order_source = self.combo_job_template.currentData()
        job_order_start_marker = self.txt_job_start_marker.text().strip()
        job_order_end_marker = self.txt_job_end_marker.text().strip()

        if has_job and job_order_source == "custom" and not job_order_start_marker:
            self.show_status("Start text is required when using the Custom template.", "error")
            return

        # Prepare parameters
        soe_table_names_text = self.txt_soe_table_names.toPlainText()
        soe_table_names_list = [line.strip() for line in soe_table_names_text.split('\n') if line.strip()]
        parsed_table_names = parse_table_names(soe_table_names_list) if has_soe else None

        # Update UI for processing state
        self.btn_generate.setEnabled(False)
        self.btn_generate.setText("Processing...")
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate loading spinner style
        self.show_status("Extracting PDF contents and writing workbook...", "info")

        # Disable selection buttons
        self.set_controls_enabled(False)

        # Spawn background worker thread
        self.worker = WorkerThread(
            excel_path=self.excel_file_path,
            proforma_path=self.proforma_file_path,
            soe_entries=self.selected_soe_files if has_soe else None,
            job_order_path=self.job_order_file_path,
            job_order_source=job_order_source,
            job_order_start_marker=job_order_start_marker,
            job_order_end_marker=job_order_end_marker,
            soe_table_names=parsed_table_names
        )
        self.worker.finished_signal.connect(self.on_generation_finished)
        self.worker.error_signal.connect(self.on_generation_error)
        self.worker.start()

    def set_controls_enabled(self, enabled: bool):
        self.btn_choose_excel.setEnabled(enabled)
        self.btn_choose_proforma.setEnabled(enabled)
        self.btn_choose_soe_files.setEnabled(enabled)
        self.btn_choose_soe_folder.setEnabled(enabled)
        self.btn_choose_job.setEnabled(enabled)
        self.combo_job_template.setEnabled(enabled)
        self.txt_soe_table_names.setEnabled(enabled)
        self.txt_job_start_marker.setEnabled(enabled)
        self.txt_job_end_marker.setEnabled(enabled)

    def on_generation_finished(self, results: dict, output_path: str):
        # Restore GUI
        self.btn_generate.setEnabled(True)
        self.btn_generate.setText("Generate Workbook")
        self.progress_bar.setVisible(False)
        self.set_controls_enabled(True)
        self.hide_status()

        self.generated_output_path = Path(output_path)

        # Build nice summary block text
        summary_lines = []
        summary_lines.append(f"<b>Workbook saved to:</b><br>{self.generated_output_path.name}<br>")
        
        sections = results.get("processed_sections", [])
        summary_lines.append(f"<b>Processed sections:</b> {', '.join(sections)}")

        if "proforma" in results and results["proforma"]:
            prof = results["proforma"]
            summary_lines.append(
                f"• Proforma: Extracted {prof['item_count']} items (Total: ${prof['gross_total']:,.2f})"
            )
        if "soe" in results and results["soe"]:
            soe = results["soe"]
            rig_filter = soe.get("rig_filter", "")
            rig_filter_msg = f" (Rig Filter: '{rig_filter}')" if rig_filter else ""
            summary_lines.append(
                f"• SOE Log: Processed {soe['row_count']} rows across {soe['pdf_count']} PDF reports{rig_filter_msg}"
            )
            
            # Render a detailed HTML table for each SOE PDF processed
            summaries = soe.get("pdf_summaries", [])
            if summaries:
                table_html = [
                    "<br><table border='0' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%; border: 1px solid #273043; font-size: 11px;'>"
                    "<tr style='background-color: #1e293b; color: #f8fafc; font-weight: bold;'>"
                    "  <th align='left'>PDF Filename</th>"
                    "  <th align='left'>Source</th>"
                    "  <th align='left'>Well / Date</th>"
                    "  <th align='left'>Period</th>"
                    "  <th align='left'>Rig</th>"
                    "  <th align='right'>Rows</th>"
                    "</tr>"
                ]
                for summary in summaries:
                    filename = summary.get("filename", "")
                    src = summary.get("source", "")
                    src_display = "Time Log" if src == "time_log" else ("Op Time Summary" if src == "operational_time_summary" else src)

                    well_or_date = summary.get("report_date", "") if src == "operational_time_summary" else summary.get("well_name", "")
                    well_or_date = well_or_date or "-"

                    # Build period string (from – to)
                    period_from = summary.get("report_period_from", "") or ""
                    period_to = summary.get("report_period_to", "") or ""
                    if period_from and period_to and period_from != period_to:
                        period = f"{period_from} – {period_to}"
                    elif period_from:
                        period = period_from
                    else:
                        period = "-"

                    rig = summary.get("rig", "") or "-"

                    if summary.get("skipped"):
                        reason = summary.get("skip_reason", "")
                        if reason == "rig_mismatch":
                            status = "Skipped (rig mismatch)"
                        elif reason == "no_matching_table":
                            status = "Skipped (no table)"
                        elif reason == "empty_table":
                            status = "Skipped (empty table)"
                        else:
                            status = "Skipped"
                        status_style = "color: #fca5a5; font-weight: bold;"
                    else:
                        status = str(summary.get("row_count", 0))
                        status_style = ""

                    table_html.append(
                        f"<tr style='border-bottom: 1px solid #273043;'>"
                        f"  <td>{filename}</td>"
                        f"  <td>{src_display}</td>"
                        f"  <td>{well_or_date}</td>"
                        f"  <td>{period}</td>"
                        f"  <td>{rig}</td>"
                        f"  <td align='right' style='{status_style}'>{status}</td>"
                        f"</tr>"
                    )
                table_html.append("</table>")
                summary_lines.append("".join(table_html))

        if "job_order" in results and results["job_order"]:
            jo = results["job_order"]
            lines_data = jo.get("lines", [])
            steps = sum(1 for ln in lines_data if ln.get("kind") == "step")
            others = jo["line_count"] - steps
            jo_detail = f"{steps} numbered step{'' if steps == 1 else 's'}"
            if others > 0:
                jo_detail += f", {others} text/bullet/table row{'' if others == 1 else 's'}"
            summary_lines.append(
                f"• Job Order: Appended {jo['line_count']} rows — {jo_detail} ('{jo['source']}')"
            )

        self.lbl_results_summary.setText("<br>".join(summary_lines))
        self.results_card.setVisible(True)

        QMessageBox.information(
            self,
            "Success",
            "Workbook generated successfully!"
        )

    def on_generation_error(self, err_msg: str):
        # Restore GUI
        self.btn_generate.setEnabled(True)
        self.btn_generate.setText("Generate Workbook")
        self.progress_bar.setVisible(False)
        self.set_controls_enabled(True)
        
        first_line = err_msg.split('\n')[0] if err_msg else "Workbook generation failed."
        self.show_status(f"Error: {first_line}", "error")

        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setWindowTitle("Processing Error")
        msg_box.setText(err_msg)
        msg_box.exec()

    # --- ACTION BUTTONS ---

    def open_generated_file(self):
        if not self.generated_output_path or not self.generated_output_path.is_file():
            QMessageBox.warning(self, "File Not Found", "The generated file could not be found.")
            return

        try:
            self.launch_path(self.generated_output_path)
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to open the generated file:\n{exc}")

    def open_generated_folder(self):
        if not self.generated_output_path:
            QMessageBox.warning(self, "Directory Not Found", "No file has been generated yet.")
            return

        folder = self.generated_output_path.parent
        try:
            self.launch_path(folder)
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Failed to open folder:\n{exc}")

    def launch_path(self, path: Path):
        """Cross-platform opening of files or folders."""
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
    app = QApplication(sys.argv)
    
    # Optional: set a styling icon using standard built-in QStyle
    window = MainWindow()
    
    # Try to set application icon if available
    try:
        app_icon = window.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)
        window.setWindowIcon(app_icon)
    except Exception:
        pass

    window.show()
    sys.exit(app.exec())
