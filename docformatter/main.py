"""
DocFormatter - Word文档格式批量处理工具
程序入口
"""

import sys
from pathlib import Path

if not getattr(sys, 'frozen', False):
    current_file = Path(__file__).resolve()
    project_root = current_file.parent.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    sys.path.insert(0, str(current_file.parent))

from PyQt6.QtWidgets import QApplication
from docformatter.gui.main_window import MainWindow


def main():
    """程序入口"""
    app = QApplication(sys.argv)
    
    app.setApplicationName("DocFormatter")
    app.setApplicationVersion("2.0.0")
    app.setOrganizationName("DocFormatter")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
