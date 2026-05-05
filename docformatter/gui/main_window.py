"""
主窗口模块
程序入口，Tab导航，菜单栏
"""

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QMenuBar, QMenu, QStatusBar, QMessageBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction

from .template_config import TemplateConfigWidget
from .doc_process import DocProcessWidget
from ..utils.logger import get_log_emitter, setup_logger


class MainWindow(QMainWindow):
    """
    DocFormatter 主窗口
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocFormatter - Word文档格式批量处理工具")
        self.setMinimumSize(900, 650)

        # 设置日志
        setup_logger()

        # 初始化UI
        self._init_ui()

    def _init_ui(self):
        """初始化UI"""
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 布局
        layout = QVBoxLayout(central_widget)

        # Tab部件：2个 Tab
        self.tab_widget = QTabWidget()
        self.tab_widget.addTab(TemplateConfigWidget(), "模板编辑")
        self.tab_widget.addTab(DocProcessWidget(), "文档处理")

        layout.addWidget(self.tab_widget)

        # 菜单栏
        self._create_menu_bar()

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
    
    def _create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件(F)")
        
        new_action = QAction("新建模板", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self._on_new_template)
        file_menu.addAction(new_action)
        
        open_action = QAction("打开模板...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._on_open_template)
        file_menu.addAction(open_action)
        
        save_action = QAction("保存模板...", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self._on_save_template)
        file_menu.addAction(save_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("退出", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助(H)")
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self._on_about)
        help_menu.addAction(about_action)
        
        manual_action = QAction("使用说明", self)
        manual_action.triggered.connect(self._on_manual)
        help_menu.addAction(manual_action)
    
    def _on_new_template(self):
        """新建模板"""
        current_widget = self.tab_widget.currentWidget()
        if hasattr(current_widget, 'new_template'):
            current_widget.new_template()
    
    def _on_open_template(self):
        """打开/导入模板"""
        current_widget = self.tab_widget.currentWidget()
        if hasattr(current_widget, '_import_template'):
            current_widget._import_template()
    
    def _on_save_template(self):
        """保存模板"""
        current_widget = self.tab_widget.currentWidget()
        if hasattr(current_widget, 'save_template'):
            current_widget.save_template()
    
    def _on_about(self):
        """关于"""
        QMessageBox.about(
            self,
            "关于 DocFormatter",
            "DocFormatter v2.0.0\n\n"
            "Word文档格式批量处理工具\n"
            "基于Python + PyQt6开发\n"
            "两段式架构：Word→MD→Word\n\n"
            "功能：\n"
            "- 模板编辑：设置文档格式模板\n"
            "- 文档处理：Word格式整理 / Word→MD / MD→Word\n"
            "- 图序/表序/公式序自动编号\n"
            "- 交叉引用/脚注/Visio OLE 提取"
        )

    def _on_manual(self):
        """使用说明"""
        QMessageBox.information(
            self,
            "使用说明",
            "DocFormatter 使用说明\n\n"
            "1. 模板编辑：设置文档格式模板（字体、段落、页眉页脚等）\n"
            "2. 文档处理：\n"
            "   - Word格式整理：Word→MD→Word 两段式格式化\n"
            "   - Word→MD：将Word转换为Markdown中间态\n"
            "   - MD→Word：从Markdown生成格式化Word文档\n\n"
            "详细使用请参考帮助文档。"
        )
    
    def update_status(self, message: str):
        """更新状态栏消息"""
        self.status_bar.showMessage(message)
