"""
批量处理模块
批量选择文件、启动处理、显示进度
"""

import traceback
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTextEdit, QProgressBar,
    QListWidget, QListWidgetItem, QMessageBox, QGroupBox
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt

from ..core.formatter import DocumentFormatter
from ..templates.template_io import load_template
from ..utils.logger import get_log_emitter, get_logger

logger = get_logger()


class FormatWorker(QThread):
    """格式化工作线程"""
    progress = pyqtSignal(int, int, str)  # current, total, message
    finished = pyqtSignal(bool, str)  # success, message
    log = pyqtSignal(str)  # log message
    
    def __init__(self, input_paths, output_dir, template_path):
        super().__init__()
        self.input_paths = input_paths
        self.output_dir = output_dir
        self.template_path = template_path
    
    def run(self):
        try:
            # 加载模板
            template = load_template(self.template_path)
            
            # 创建格式化器
            formatter = DocumentFormatter(template)
            
            # 批量处理
            total = len(self.input_paths)
            for i, input_path in enumerate(self.input_paths):
                input_name = Path(input_path).stem
                output_path = str(Path(self.output_dir) / f"{input_name}_formatted.docx")
                
                self.progress.emit(i + 1, total, f"正在处理: {input_name}...")
                self.log.emit(f"[{i+1}/{total}] 开始处理: {input_path}")
                
                try:
                    result = formatter.format(input_path, output_path)
                    
                    if result.success:
                        self.log.emit(f"  ✓ 完成: {output_path}")
                        self.log.emit(f"    图片: {result.figures_processed}, 表格: {result.tables_processed}, 公式: {result.equations_processed}")
                    else:
                        self.log.emit(f"  ✗ 失败: {', '.join(result.errors)}")
                
                except Exception as e:
                    self.log.emit(f"  ✗ 异常: {str(e)}")
                    logger.error(f"处理 {input_path} 时出错: {e}")
                    logger.error(traceback.format_exc())
            
            self.progress.emit(total, total, "处理完成")
            self.finished.emit(True, f"批量处理完成！成功: {total}/{total}")
            
        except Exception as e:
            logger.error(f"批量处理失败: {e}")
            logger.error(traceback.format_exc())
            self.finished.emit(False, f"批量处理失败: {str(e)}")


class BatchProcessWidget(QWidget):
    """
    批量处理界面
    """
    
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.template_path = None
        self.output_dir = None
        self.worker = None
        
        self._init_ui()
    
    def _init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)
        
        # ===== 文件选择区 =====
        file_group = QGroupBox("源文件")
        file_layout = QVBoxLayout(file_group)
        
        # 文件选择按钮
        btn_layout = QHBoxLayout()
        
        self.btn_select_files = QPushButton("选择文件")
        self.btn_select_files.clicked.connect(self._select_files)
        btn_layout.addWidget(self.btn_select_files)
        
        self.btn_select_folder = QPushButton("选择文件夹")
        self.btn_select_folder.clicked.connect(self._select_folder)
        btn_layout.addWidget(self.btn_select_folder)
        
        btn_layout.addStretch()
        file_layout.addLayout(btn_layout)
        
        # 文件列表
        self.file_list = QListWidget()
        file_layout.addWidget(self.file_list)
        
        layout.addWidget(file_group)
        
        # ===== 模板和输出区 =====
        config_group = QGroupBox("配置")
        config_layout = QHBoxLayout(config_group)
        
        # 模板选择
        template_layout = QVBoxLayout()
        template_layout.addWidget(QLabel("模板文件:"))
        self.btn_select_template = QPushButton("选择模板文件")
        self.btn_select_template.clicked.connect(self._select_template)
        template_layout.addWidget(self.btn_select_template)
        self.lbl_template = QLabel("未选择模板")
        self.lbl_template.setWordWrap(True)
        template_layout.addWidget(self.lbl_template)
        config_layout.addLayout(template_layout)
        
        # 输出目录
        output_layout = QVBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.btn_select_output = QPushButton("选择输出目录")
        self.btn_select_output.clicked.connect(self._select_output)
        output_layout.addWidget(self.btn_select_output)
        self.lbl_output = QLabel("未选择目录")
        self.lbl_output.setWordWrap(True)
        output_layout.addWidget(self.lbl_output)
        config_layout.addLayout(output_layout)
        
        layout.addWidget(config_group)
        
        # ===== 日志区 =====
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)
        
        # ===== 进度和按钮区 =====
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
    
    def _select_files(self):
        """选择文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if files:
            self.input_files = files
            self.file_list.clear()
            for f in files:
                self.file_list.addItem(Path(f).name)
            self._check_ready()
    
    def _select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        
        if folder:
            # 扫描文件夹中的所有 docx 文件
            docx_files = list(Path(folder).glob("*.docx"))
            
            if docx_files:
                self.input_files = [str(f) for f in docx_files]
                self.file_list.clear()
                for f in self.input_files:
                    self.file_list.addItem(Path(f).name)
                self._check_ready()
            else:
                QMessageBox.information(self, "提示", "文件夹中没有找到Word文档")
    
    def _select_template(self):
        """选择模板文件"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择模板文件",
            "",
            "JSON文件 (*.json);;所有文件 (*.*)"
        )
        
        if file:
            self.template_path = file
            self.lbl_template.setText(Path(file).name)
            self._check_ready()
    
    def _select_output(self):
        """选择输出目录"""
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录")
        
        if folder:
            self.output_dir = folder
            self.lbl_output.setText(folder)
            self._check_ready()
    
    def _check_ready(self):
        """检查是否可以开始处理"""
        ready = (
            len(self.input_files) > 0 and
            self.template_path is not None and
            self.output_dir is not None
        )
        self.btn_start.setEnabled(ready)
    
    def _start_processing(self):
        """开始处理"""
        if not self.input_files or not self.template_path or not self.output_dir:
            return
        
        # 禁用按钮
        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        
        # 清空日志
        self.log_text.clear()
        self._append_log("开始批量处理...")
        
        # 创建工作线程
        self.worker = FormatWorker(
            self.input_files,
            self.output_dir,
            self.template_path
        )
        
        # 连接信号
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self._append_log)
        self.worker.finished.connect(self._on_finished)
        
        # 启动
        self.worker.start()
    
    def _cancel(self):
        """取消处理"""
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self._append_log("已取消")
        
        self._on_finished(False, "已取消")
    
    def _on_progress(self, current, total, message):
        """进度更新"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.progress_bar.setFormat(f"{current}/{total} - {message}")
    
    def _on_finished(self, success, message):
        """处理完成"""
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        
        if success:
            self.progress_bar.setFormat("处理完成")
            QMessageBox.information(self, "完成", message)
        else:
            self.progress_bar.setFormat("处理失败")
            QMessageBox.warning(self, "处理失败", message)
    
    def _append_log(self, message: str):
        """追加日志"""
        self.log_text.append(message)
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def get_current_template_path(self) -> str:
        """获取当前模板路径（供主窗口保存模板时使用）"""
        return self.template_path or ""
