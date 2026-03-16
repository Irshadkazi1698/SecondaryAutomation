import os
import sys

from PyQt5 import QtGui
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import main
from SanityCheckModule.SanityChecking import createSanityCheck
from SanityCheckModule.SanityCheckingTabPlan2 import createSanityCheck as createSanityCheckTabPlan2
from GridTable.CreateGridTables import GenerateGridTables


class WorkThread(QThread):
    progress = pyqtSignal(str)
    progress_value = pyqtSignal(int)

    def __init__(
        self,
        input_dir,
        output_dir,
        banner_file,
        count_file,
        num_var,
        tabplan_file,
        tabplan_choice,
        question_index,
        label_index,
        base_index,
        sheet_name,
        grid_enable,
        grid_counts_file,
    ):
        super().__init__()
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.banner_file = banner_file
        self.count_file = count_file
        self.num_var = num_var
        self.tabplan_file = tabplan_file
        self.tabplan_choice = tabplan_choice
        self.question_index = question_index
        self.label_index = label_index
        self.base_index = base_index
        self.sheet_name = sheet_name
        self.grid_enable = grid_enable
        self.grid_counts_file = grid_counts_file

    def run(self):
        try:
            self.progress_value.emit(5)
            self.progress.emit("Running Banner Validation...")

            validation = main.BannerValidation(
                self.input_dir,
                self.output_dir,
                self.banner_file,
                self.count_file,
                self.num_var,
                self.tabplan_file,
                self.tabplan_choice,
                self.question_index,
                self.label_index,
                self.base_index,
                self.sheet_name,
            )

            output_files = os.listdir(self.output_dir)
            self.progress_value.emit(15)
            if "Matched_Variables.xlsx" not in output_files:
                self.progress.emit("Preparing Matched Variables file...")
                self.progress_value.emit(25)
                validation.CreatingMatchingFileInOutput()

            self.progress.emit("Validating banner data...")
            self.progress_value.emit(35)
            validation.BannerValidationAutomation()
            self.progress_value.emit(70)
            self.progress.emit("Validation completed successfully.")

            final_comparison_path = os.path.join(self.output_dir, "Final Comparison.xlsx")
            if not os.path.exists(final_comparison_path):
                self.progress.emit("Final Comparison file not found in output directory.")

            tabplan_filename = f"{self.tabplan_file}.xlsm"
            print(f"Looking for TabPlan file at: {os.path.join(self.input_dir, tabplan_filename)}")
            tabplan_path = os.path.join(self.input_dir, tabplan_filename)

            self.progress.emit("Running sanity check...")
            self.progress_value.emit(80)
            if self.tabplan_choice == 2:
                createSanityCheckTabPlan2(final_comparison_path, tabplan_path)
            else:   
                createSanityCheck(final_comparison_path, tabplan_path)
            self.progress_value.emit(90)
            self.progress.emit("Sanity Check completed.")

            if self.grid_enable:
                self.progress.emit("Generating Grid Tables...")
                self.progress.emit("This may take a few moments depending on the size of the data.")
                self.progress_value.emit(93)
                grid_counts_path = os.path.join(self.input_dir, f"{self.grid_counts_file}.xlsx")
                grid_output_path = os.path.join(self.output_dir, "Final Comparison.xlsx")
                GenerateGridTables(grid_counts_path, grid_output_path)
                self.progress_value.emit(98)
                self.progress.emit("Grid Tables generated successfully.")

            self.progress_value.emit(100)
            self.progress.emit("All steps completed.")
        except BaseException as exc:
            self.progress.emit(f"Error: {exc}")
            self.progress_value.emit(0)


class CustomTabPlanDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Custom TabPlan Details")
        self.setWindowIcon(QtGui.QIcon("icon.png"))
        self.setMinimumWidth(460)

        self.question_edit = QLineEdit()
        self.label_edit = QLineEdit()
        self.base_edit = QLineEdit()
        self.sheet_edit = QLineEdit()

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.addRow("Question column index:", self.question_edit)
        form.addRow("Label column index:", self.label_edit)
        form.addRow("Base text column index:", self.base_edit)
        form.addRow("Sheet name:", self.sheet_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        root = QVBoxLayout(self)
        root.addLayout(form)
        root.addWidget(buttons)

    def values(self):
        return {
            "question_index": self.question_edit.text().strip(),
            "label_index": self.label_edit.text().strip(),
            "base_index": self.base_edit.text().strip(),
            "sheet_name": self.sheet_edit.text().strip(),
        }


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Ipsos Banner Validation Tool")
        self.setWindowIcon(QtGui.QIcon("icon.png"))
        self._configure_window_for_screen()

        self.custom_values = {}
        self.worker = None

        self._build_ui()

    def _configure_window_for_screen(self):
        screen = QtGui.QGuiApplication.primaryScreen()
        if not screen:
            self.resize(1100, 760)
            self.setMinimumSize(800, 620)
            return

        available = screen.availableGeometry()
        width = max(900, int(available.width() * 0.85))
        height = max(650, int(available.height() * 0.85))
        self.resize(width, height)
        self.setMinimumSize(min(900, width), min(650, height))

    def _build_ui(self):
        title = QLabel("Ipsos Banner Validation Tool")
        title.setObjectName("titleLabel")
        title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.input_edit, input_row = self._file_row(
            "Input directory", self.input_folder_dialog, is_folder=True
        )
        self.output_edit, output_row = self._file_row(
            "Output directory", self.output_folder_dialog, is_folder=True
        )
        self.banner_edit, banner_row = self._file_row(
            "Banner file", self.banner_file_name
        )
        self.count_edit, count_row = self._file_row(
            "Counts file", self.count_file_name
        )
        self.num_edit, num_row = self._file_row(
            "Numeric variable (.inc)", self.numeric_var_file_name
        )
        self.tabplan_edit, tabplan_row = self._file_row(
            "TabPlan file", self.tabplan_file_name
        )

        file_group = QGroupBox("Files")
        file_layout = QFormLayout()
        file_layout.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        file_layout.setHorizontalSpacing(14)
        file_layout.setVerticalSpacing(10)
        file_layout.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        file_layout.addRow("Input directory", input_row)
        file_layout.addRow("Output directory", output_row)
        file_layout.addRow("Banner file", banner_row)
        file_layout.addRow("Counts file", count_row)
        file_layout.addRow("Numeric variable (.inc)", num_row)
        file_layout.addRow("TabPlan file", tabplan_row)
        file_group.setLayout(file_layout)

        tabplan_group = QGroupBox("TabPlan Category")
        tabplan_layout = QHBoxLayout()
        tabplan_hint = QLabel("Choose one:")

        self.tabplan_choice_combo = QComboBox()
        self.tabplan_choice_combo.addItem("1 - AutoPlan", 1)
        self.tabplan_choice_combo.addItem("2 - Tabplan", 2)
        self.tabplan_choice_combo.addItem("3 - Custom", 3)
        self.tabplan_choice_combo.currentIndexChanged.connect(self._on_tabplan_choice_changed)

        self.custom_btn = QPushButton("Custom Details")
        self.custom_btn.clicked.connect(self.open_custom_dialog)
        self.custom_btn.setEnabled(False)

        tabplan_layout.addWidget(tabplan_hint)
        tabplan_layout.addWidget(self.tabplan_choice_combo, 1)
        tabplan_layout.addWidget(self.custom_btn)
        tabplan_group.setLayout(tabplan_layout)

        grid_group = QGroupBox("Grid Tables (Optional)")
        grid_layout = QVBoxLayout()
        self.grid_enable_check = QCheckBox("Enable Grid Tables Generation")
        self.grid_enable_check.stateChanged.connect(self._on_grid_enable_changed)

        self.grid_counts_edit, grid_counts_row = self._file_row(
            "Grid Counts file", self.grid_counts_file_name
        )

        self.grid_counts_edit.setEnabled(False)

        grid_form = QFormLayout()
        grid_form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        grid_form.setHorizontalSpacing(14)
        grid_form.setVerticalSpacing(10)
        grid_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        grid_form.addRow("Grid Counts file", grid_counts_row)

        grid_layout.addWidget(self.grid_enable_check)
        grid_layout.addLayout(grid_form)
        grid_group.setLayout(grid_layout)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setPlaceholderText("Execution logs will appear here...")
        self.output.setMinimumHeight(170)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Progress: %p%")
        self.progress_bar.setTextVisible(True)

        self.run_btn = QPushButton("Run")
        self.run_btn.clicked.connect(self.run_background_task)
        self.run_btn.setMinimumWidth(130)

        bottom_row = QHBoxLayout()
        bottom_row.addStretch(1)
        bottom_row.addWidget(self.run_btn)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 16, 16, 16)
        content_layout.setSpacing(12)
        content_layout.addWidget(title)
        content_layout.addWidget(file_group)
        content_layout.addWidget(tabplan_group)
        content_layout.addWidget(grid_group)
        content_layout.addWidget(self.progress_bar)
        content_layout.addWidget(self.output, 1)
        content_layout.addLayout(bottom_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(content)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(scroll)

        self.setStyleSheet(
            """
            QWidget {
                font-size: 12px;
            }
            QLabel#titleLabel {
                font-size: 20px;
                font-weight: 600;
                padding: 4px 0;
            }
            QLineEdit {
                min-height: 30px;
            }
            QPushButton {
                min-height: 30px;
                padding: 4px 10px;
            }
            """
        )

    def _file_row(self, label_text, browse_handler, is_folder=False):
        edit = QLineEdit()
        edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        edit.setMinimumWidth(140)

        browse = QPushButton("Browse")
        browse.clicked.connect(browse_handler)
        browse.setProperty("is_folder", is_folder)
        browse.setMinimumWidth(84)
        browse.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(edit, 1)
        layout.addWidget(browse)
        layout.setStretch(0, 1)
        layout.setStretch(1, 0)

        return edit, row

    def _pick_file_name_without_extension(self, title, filters):
        path, _ = QFileDialog.getOpenFileName(self, title, "", filters)
        if not path:
            return ""
        file_name = os.path.basename(path)
        return os.path.splitext(file_name)[0]

    def input_folder_dialog(self):
        path = QFileDialog.getExistingDirectory(self, "Select Input Folder", "")
        if path:
            self.input_edit.setText(path)

    def output_folder_dialog(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder", "")
        if path:
            self.output_edit.setText(path)

    def banner_file_name(self):
        clean_name = self._pick_file_name_without_extension(
            "Select Banner File", "Excel Files (*.xlsx *.xls)"
        )
        if clean_name:
            self.banner_edit.setText(clean_name)

    def count_file_name(self):
        clean_name = self._pick_file_name_without_extension(
            "Select Counts File", "Excel Files (*.xlsx *.xls)"
        )
        if clean_name:
            self.count_edit.setText(clean_name)

    def numeric_var_file_name(self):
        clean_name = self._pick_file_name_without_extension(
            "Select Numeric Variable File", "Allowed Files (*.xlsx *.xls *.inc)"
        )
        if clean_name:
            self.num_edit.setText(clean_name)

    def tabplan_file_name(self):
        clean_name = self._pick_file_name_without_extension(
            "Select TabPlan File", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        if clean_name:
            self.tabplan_edit.setText(clean_name)

    def _on_tabplan_choice_changed(self):
        choice = self.tabplan_choice_combo.currentData()
        self.custom_btn.setEnabled(choice == 3)

    def _on_grid_enable_changed(self):
        enabled = self.grid_enable_check.isChecked()
        self.grid_counts_edit.setEnabled(enabled)

    def grid_counts_file_name(self):
        clean_name = self._pick_file_name_without_extension(
            "Select Grid Counts File", "Excel Files (*.xlsx *.xls)"
        )
        if clean_name:
            self.grid_counts_edit.setText(clean_name)

    def open_custom_dialog(self):
        dialog = CustomTabPlanDialog(self)

        if self.custom_values:
            dialog.question_edit.setText(self.custom_values.get("question_index", ""))
            dialog.label_edit.setText(self.custom_values.get("label_index", ""))
            dialog.base_edit.setText(self.custom_values.get("base_index", ""))
            dialog.sheet_edit.setText(self.custom_values.get("sheet_name", ""))

        if dialog.exec_() == QDialog.Accepted:
            self.custom_values = dialog.values()

    def run_background_task(self):
        try:
            input_dir = self.input_edit.text().strip()
            output_dir = self.output_edit.text().strip()
            num_var = self.num_edit.text().strip()
            tabplan_file = self.tabplan_edit.text().strip()
            tabplan_choice = int(self.tabplan_choice_combo.currentData())
            banner_file = self.banner_edit.text().strip()
            count_file = self.count_edit.text().strip()
            grid_enable = self.grid_enable_check.isChecked()
            grid_counts_file = self.grid_counts_edit.text().strip() if grid_enable else ""

            if not input_dir or not output_dir:
                self.output.append("Please select both Input and Output directories.")
                return

            if grid_enable and not grid_counts_file:
                self.output.append("Please select Grid Counts file.")
                return

            question_index = None
            label_index = None
            base_index = None
            sheet_name = None

            if tabplan_choice == 3:
                if not self.custom_values:
                    self.output.append("Please add Custom TabPlan details first.")
                    return

                try:
                    question_index = int(self.custom_values.get("question_index", ""))
                    label_index = int(self.custom_values.get("label_index", ""))
                    base_index = int(self.custom_values.get("base_index", ""))
                    sheet_name = self.custom_values.get("sheet_name", "")

                    if not sheet_name:
                        raise ValueError("Sheet name is required for custom tabplan.")
                except ValueError as exc:
                    self.output.append(f"Invalid Custom TabPlan details: {exc}")
                    return

            self.output.append("Starting background task...")
            self.run_btn.setEnabled(False)
            self.progress_bar.setValue(0)

            self.worker = WorkThread(
                input_dir,
                output_dir,
                banner_file,
                count_file,
                num_var,
                tabplan_file,
                tabplan_choice,
                question_index,
                label_index,
                base_index,
                sheet_name,
                grid_enable,
                grid_counts_file,
            )
            self.worker.progress.connect(self.update_output)
            self.worker.progress_value.connect(self.update_progress)
            self.worker.finished.connect(self._on_worker_finished)
            self.worker.start()

        except Exception as exc:
            self.output.append(f"Error while starting task: {exc}")

    def update_output(self, text):
        self.output.append(text)

    def update_progress(self, value):
        self.progress_bar.setValue(max(0, min(100, int(value))))

    def _on_worker_finished(self):
        self.run_btn.setEnabled(True)


if __name__ == "__main__":
    # Helps scaling across monitors with different DPI.
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    window = MainWindow()
    window.output.append("Please fill in the details and click 'Run' to start the validation process.")
    window.output.append("Execution logs will appear here.")
    window.output.append("Note: If Grid is Enabled, it may take additional time to generate the grid tables based on the size of the data.")
    window.showMaximized()
    sys.exit(app.exec())
