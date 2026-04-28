"""
DocFormatter - Word文档格式批量处理工具
程序入口
"""

import sys
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from PyQt6.QtWidgets import QApplication

from gui.main_window import MainWindow


def main():
    """程序入口"""
    app = QApplication(sys.argv)
    
    # 设置应用信息
    app.setApplicationName("DocFormatter")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("DocFormatter")
    
    # 创建并显示主窗口
    window = MainWindow()
    window.show()
    
    # 运行应用
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
