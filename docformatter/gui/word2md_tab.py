"""
Word转MD模块 (已废弃)
@deprecated: 此模块已被 doc_process.DocProcessWidget 替代。
"""

from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTextEdit, QGroupBox,
    QMessageBox
)
from PyQt6.QtCore import QThread, pyqtSignal

from ..utils.logger import get_logger

logger = get_logger()


class Word2MDWorker(QThread):
    """转换工作线程"""
    finished = pyqtSignal(bool, str)
    log = pyqtSignal(str)
    
    def __init__(self, input_path, output_path):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
    
    def run(self):
        try:
            self.log.emit(f"开始转换: {self.input_path}")
            
            # 使用 mammoth 进行转换
            import mammoth
            
            with open(self.input_path, 'rb') as docx_file:
                result = mammoth.convert_to_markdown(docx_file)
                
                # 写入输出文件
                with open(self.output_path, 'w', encoding='utf-8') as md_file:
                    md_file.write(result.value)
                
                # 输出转换消息
                if result.messages:
                    self.log.emit(f"转换完成，有 {len(result.messages)} 条消息:")
                    for msg in result.messages:
                        self.log.emit(f"  - {msg}")
                else:
                    self.log.emit("转换完成，无警告")
                
                self.finished.emit(True, f"转换成功:\n{self.output_path}")
                
        except Exception as e:
            logger.error(f"转换失败: {e}")
            self.log.emit(f"转换失败: {str(e)}")
            self.finished.emit(False, f"转换失败:\n{str(e)}")


class Word2MDWidget(QWidget):
    """
    Word转MD界面
    """
    
    def __init__(self):
        super().__init__()
        self.input_path = None
        self.output_path = None
        self.worker = None
        
        self._init_ui()
    
    def _init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)
        
        # ===== 文件选择区 =====
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)
        
        # 输入文件
        input_layout = QHBoxLayout()
        self.btn_select_input = QPushButton("选择Word文件")
        self.btn_select_input.clicked.connect(self._select_input)
        input_layout.addWidget(self.btn_select_input)
        self.lbl_input = QLabel("未选择文件")
        self.lbl_input.setWordWrap(True)
        input_layout.addWidget(self.lbl_input)
        file_layout.addLayout(input_layout)
        
        # 输出文件
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件:"))
        self.output_path_edit = QLabel("未指定")
        self.output_path_edit.setWordWrap(True)
        output_layout.addWidget(self.output_path_edit)
        file_layout.addLayout(output_layout)
        
        layout.addWidget(file_group)
        
        # ===== 转换按钮 =====
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.btn_convert = QPushButton("转换为Markdown")
        self.btn_convert.clicked.connect(self._start_convert)
        self.btn_convert.setEnabled(False)
        btn_layout.addWidget(self.btn_convert)
        
        self.btn_preview = QPushButton("预览")
        self.btn_preview.clicked.connect(self._preview)
        self.btn_preview.setEnabled(False)
        btn_layout.addWidget(self.btn_preview)
        
        layout.addLayout(btn_layout)
        
        # ===== 日志区 =====
        log_group = QGroupBox("转换日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)
        
        # 提示
        hint_layout = QHBoxLayout()
        hint_layout.addWidget(QLabel(
            "提示: 转换结果可能需要手动调整，特别是表格、图片和公式"
        ))
        hint_layout.addStretch()
        layout.addLayout(hint_layout)
    
    def _select_input(self):
        """选择输入文件"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file:
            self.input_path = file
            self.lbl_input.setText(Path(file).name)
            
            # 自动设置输出路径
            self.output_path = str(Path(file).with_suffix('.md'))
            self.output_path_edit.setText(Path(self.output_path).name)
            
            self.btn_convert.setEnabled(True)
    
    def _start_convert(self):
        """开始转换"""
        if not self.input_path:
            return
        
        # 确认输出路径
        suggested_output = self.input_path.replace('.docx', '.md')
        output, _ = QFileDialog.getSaveFileName(
            self,
            "保存Markdown文件",
            suggested_output,
            "Markdown文件 (*.md);;所有文件 (*.*)"
        )
        
        if not output:
            return
        
        self.output_path = output
        
        # 禁用按钮
        self.btn_convert.setEnabled(False)
        self.log_text.clear()
        
        # 创建工作线程
        self.worker = Word2MDWorker(self.input_path, self.output_path)
        self.worker.log.connect(self._append_log)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()
    
    def _on_finished(self, success, message):
        """转换完成"""
        self.btn_convert.setEnabled(True)
        self.btn_preview.setEnabled(success)
        
        if success:
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.warning(self, "失败", message)
    
    def _preview(self):
        """预览转换结果"""
        if self.output_path and Path(self.output_path).exists():
            # 用系统默认程序打开
            import subprocess
            subprocess.Popen(['notepad', self.output_path])
    
    def _append_log(self, message: str):
        """追加日志"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
