"""
文档处理 Tab (Phase B-3)
合并"批量处理"+"Word转MD"为单一 Tab，支持 3 种状态
"""

from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTextEdit, QProgressBar,
    QListWidget, QGroupBox, QButtonGroup, QRadioButton,
    QMessageBox,
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt

from ..core.word2md_converter import Word2MDConverter
from ..core.md2word_converter import MD2WordConverter
from ..templates.template_io import load_template
from ..utils.logger import get_log_emitter, get_logger

logger = get_logger()


class ProcessWorker(QThread):
    """文档处理工作线程"""
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(bool, str)
    log = pyqtSignal(str)

    def __init__(self, state, input_paths, output_dir, template_path=None):
        super().__init__()
        self.state = state
        self.input_paths = input_paths
        self.output_dir = output_dir
        self.template_path = template_path

    def run(self):
        try:
            total = len(self.input_paths)
            output_dir = Path(self.output_dir)

            if self.state == "A":
                # Word 格式整理：Word→MD→Word
                template = load_template(self.template_path)
                w2m = Word2MDConverter()

                for i, input_path in enumerate(self.input_paths):
                    name = Path(input_path).stem
                    self.progress.emit(i + 1, total, f"处理中: {name}")
                    self.log.emit(f"[{i+1}/{total}] {name}")

                    try:
                        # Stage 1: Word→MD
                        self.log.emit(f"  ① Word→MD...")
                        md_dir = output_dir / f"{name}_md"
                        md_path = w2m.convert(input_path, str(md_dir))
                        self.log.emit(f"    MD: {md_path}")

                        # Stage 2: MD→Word
                        self.log.emit(f"  ② MD→Word...")
                        m2w = MD2WordConverter(template)
                        output_path = output_dir / f"{name}_formatted.docx"
                        m2w.convert(md_path, str(output_path))
                        self.log.emit(f"    ✓ {output_path}")

                    except Exception as e:
                        self.log.emit(f"    ✗ 失败: {e}")
                        logger.error(f"处理 {input_path} 失败: {e}")

            elif self.state == "B":
                # Word → MD
                w2m = Word2MDConverter()
                for i, input_path in enumerate(self.input_paths):
                    name = Path(input_path).stem
                    self.progress.emit(i + 1, total, f"转换中: {name}")
                    self.log.emit(f"[{i+1}/{total}] {name}")

                    try:
                        md_dir = output_dir / f"{name}_md"
                        md_path = w2m.convert(input_path, str(md_dir))
                        self.log.emit(f"    ✓ {md_path}")
                    except Exception as e:
                        self.log.emit(f"    ✗ 失败: {e}")

            elif self.state == "C":
                # MD → Word
                template = load_template(self.template_path)
                m2w = MD2WordConverter(template)

                for i, input_path in enumerate(self.input_paths):
                    name = Path(input_path).stem
                    self.progress.emit(i + 1, total, f"生成中: {name}")
                    self.log.emit(f"[{i+1}/{total}] {name}")

                    try:
                        output_path = output_dir / f"{name}.docx"
                        m2w.convert(input_path, str(output_path))
                        self.log.emit(f"    ✓ {output_path}")
                    except Exception as e:
                        self.log.emit(f"    ✗ 失败: {e}")

            self.progress.emit(total, total, "完成")
            self.finished.emit(True, f"处理完成！共 {total} 个文件")

        except Exception as e:
            logger.error(f"处理失败: {e}")
            self.finished.emit(False, f"处理失败: {e}")


class DocProcessWidget(QWidget):
    """文档处理 Tab"""

    STATE_LABELS = {
        "A": "Word 格式整理 (Word→MD→Word)",
        "B": "Word → MD",
        "C": "MD → Word",
    }

    def __init__(self):
        super().__init__()
        self.current_state = "A"
        self.input_files = []
        self.template_path = None
        self.output_dir = None
        self.worker = None

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        # ===== 状态选择 =====
        state_group = QGroupBox("处理模式")
        state_layout = QHBoxLayout(state_group)

        self.state_group = QButtonGroup(self)
        self.btn_state_a = QRadioButton("A: Word 格式整理")
        self.btn_state_b = QRadioButton("B: Word → MD")
        self.btn_state_c = QRadioButton("C: MD → Word")
        self.btn_state_a.setChecked(True)

        for btn in [self.btn_state_a, self.btn_state_b, self.btn_state_c]:
            self.state_group.addButton(btn)
            state_layout.addWidget(btn)

        state_layout.addStretch()
        layout.addWidget(state_group)

        self.btn_state_a.toggled.connect(self._on_state_changed)
        self.btn_state_b.toggled.connect(self._on_state_changed)
        self.btn_state_c.toggled.connect(self._on_state_changed)

        # ===== 文件选择区 =====
        file_group = QGroupBox("源文件")
        file_layout = QVBoxLayout(file_group)

        btn_layout = QHBoxLayout()
        self.btn_select_files = QPushButton("添加文件")
        self.btn_select_files.clicked.connect(self._select_files)
        btn_layout.addWidget(self.btn_select_files)

        self.btn_select_folder = QPushButton("添加文件夹")
        self.btn_select_folder.clicked.connect(self._select_folder)
        btn_layout.addWidget(self.btn_select_folder)

        self.btn_clear = QPushButton("清空")
        self.btn_clear.clicked.connect(self._clear_files)
        btn_layout.addWidget(self.btn_clear)

        btn_layout.addStretch()
        file_layout.addLayout(btn_layout)

        self.file_list = QListWidget()
        file_layout.addWidget(self.file_list)
        layout.addWidget(file_group)

        # ===== 模板选择 =====
        self.template_group = QGroupBox("模板")
        template_layout = QVBoxLayout(self.template_group)

        tmpl_btn_layout = QHBoxLayout()
        self.btn_select_template = QPushButton("选择模板文件")
        self.btn_select_template.clicked.connect(self._select_template)
        tmpl_btn_layout.addWidget(self.btn_select_template)
        self.lbl_template = QLabel("未选择模板")
        self.lbl_template.setWordWrap(True)
        tmpl_btn_layout.addWidget(self.lbl_template)
        tmpl_btn_layout.addStretch()
        template_layout.addLayout(tmpl_btn_layout)

        layout.addWidget(self.template_group)

        # ===== 输出目录 =====
        output_group = QGroupBox("输出目录")
        output_layout = QHBoxLayout(output_group)
        self.btn_select_output = QPushButton("选择输出目录")
        self.btn_select_output.clicked.connect(self._select_output)
        output_layout.addWidget(self.btn_select_output)
        self.lbl_output = QLabel("未选择目录")
        self.lbl_output.setWordWrap(True)
        output_layout.addWidget(self.lbl_output)
        output_layout.addStretch()
        layout.addWidget(output_group)

        # ===== 日志区 =====
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)

        # ===== 进度和按钮 =====
        bottom_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        bottom_layout.addWidget(self.progress_bar)

        self.btn_start = QPushButton("开始处理")
        self.btn_start.clicked.connect(self._start_processing)
        self.btn_start.setEnabled(False)
        bottom_layout.addWidget(self.btn_start)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.clicked.connect(self._cancel)
        self.btn_cancel.setEnabled(False)
        bottom_layout.addWidget(self.btn_cancel)
        layout.addLayout(bottom_layout)

        # 连接日志信号
        log_emitter = get_log_emitter()
        log_emitter.add_callback(self._append_log)

    # ==================== 状态切换 ====================

    def _on_state_changed(self):
        if self.btn_state_a.isChecked():
            self.current_state = "A"
        elif self.btn_state_b.isChecked():
            self.current_state = "B"
        elif self.btn_state_c.isChecked():
            self.current_state = "C"

        # 状态 B 不需要模板
        self.template_group.setVisible(self.current_state != "B")
        self._update_file_filter()
        self._check_ready()

    def _update_file_filter(self):
        """根据状态更新提示"""
        if self.current_state == "C":
            self.btn_select_files.setToolTip("选择 Markdown 文件 (.md)")
        else:
            self.btn_select_files.setToolTip("选择 Word 文档 (.docx)")

    # ==================== 文件选择 ====================

    def _select_files(self):
        if self.current_state == "C":
            filter_str = "Markdown文件 (*.md);;所有文件 (*.*)"
        else:
            filter_str = "Word文档 (*.docx);;所有文件 (*.*)"

        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", filter_str)
        if files:
            for f in files:
                if f not in self.input_files:
                    self.input_files.append(f)
            self._refresh_file_list()
            self._check_ready()

    def _select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            if self.current_state == "C":
                ext = "*.md"
            else:
                ext = "*.docx"

            for f in Path(folder).glob(ext):
                f_str = str(f)
                if f_str not in self.input_files:
                    self.input_files.append(f_str)

            self._refresh_file_list()
            self._check_ready()

            if not self.input_files:
                QMessageBox.information(self, "提示", f"文件夹中未找到 {ext} 文件")

    def _clear_files(self):
        self.input_files.clear()
        self.file_list.clear()
        self._check_ready()

    def _refresh_file_list(self):
        self.file_list.clear()
        for f in self.input_files:
            self.file_list.addItem(Path(f).name)

    # ==================== 模板/输出 ====================

    def _select_template(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", "", "JSON文件 (*.json);;所有文件 (*.*)"
        )
        if file:
            self.template_path = file
            self.lbl_template.setText(Path(file).name)
            self._check_ready()

    def _select_output(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if folder:
            self.output_dir = folder
            self.lbl_output.setText(folder)
            self._check_ready()

    # ==================== 就绪检查 ====================

    def _check_ready(self):
        has_files = len(self.input_files) > 0
        has_output = self.output_dir is not None
        needs_template = self.current_state != "B"
        has_template = self.template_path is not None

        if needs_template:
            ready = has_files and has_output and has_template
        else:
            ready = has_files and has_output

        self.btn_start.setEnabled(ready)

    # ==================== 处理 ====================

    def _start_processing(self):
        if not self.input_files or not self.output_dir:
            return
        if self.current_state != "B" and not self.template_path:
            return

        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.log_text.clear()
        self._append_log(f"开始处理 - 模式: {self.STATE_LABELS[self.current_state]}")
        self._append_log(f"文件数: {len(self.input_files)}")

        self.worker = ProcessWorker(
            self.current_state,
            self.input_files,
            self.output_dir,
            self.template_path
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self._append_log)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _cancel(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self._append_log("处理已取消")
        self._reset_buttons()

    def _on_progress(self, current, total, message):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.progress_bar.setFormat(message)

    def _on_finished(self, success, message):
        self._reset_buttons()
        if success:
            self.progress_bar.setFormat("处理完成")
            QMessageBox.information(self, "完成", message)
        else:
            self.progress_bar.setFormat("处理失败")
            QMessageBox.warning(self, "失败", message)

    def _reset_buttons(self):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)

    def _append_log(self, message: str):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
