"""
DocFormatter - Word文档格式批量处理工具
程序入口
"""

import sys
from pathlib import Path

# 获取项目根目录（docformatter/ 的父目录）
current_file = Path(__file__).resolve()
project_root = current_file.parent.parent

# 添加到路径
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# 确保可以导入子模块
import docformatter
sys.path.insert(0, str(current_file.parent))

from PyQt6.QtWidgets import QApplication
from docformatter.gui.main_window import MainWindow


def main():
    """程序入口"""
    app = QApplication(sys.argv)
    
    app.setApplicationName("DocFormatter")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("DocFormatter")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
