"""
日志工具
统一日志输出到GUI和文件
"""

import logging
import sys
from pathlib import Path
from typing import Optional
from datetime import datetime


class LogEmitter:
    """日志发射器，用于将日志发送到GUI等UI组件"""
    
    def __init__(self):
        self._callbacks = []
    
    def add_callback(self, callback):
        """添加日志回调函数"""
        self._callbacks.append(callback)
    
    def remove_callback(self, callback):
        """移除日志回调函数"""
        if callback in self._callbacks:
            self._callbacks.remove(callback)
    
    def emit(self, level: str, message: str):
        """发射日志消息"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        formatted = f"[{timestamp}] [{level}] {message}"
        for callback in self._callbacks:
            try:
                callback(formatted)
            except Exception:
                pass


# 全局日志发射器
_log_emitter = LogEmitter()


def get_log_emitter() -> LogEmitter:
    """获取全局日志发射器"""
    return _log_emitter


class GuiLogHandler(logging.Handler):
    """logging.Handler 的子类，用于将日志发送到GUI"""
    
    def __init__(self, emitter: LogEmitter):
        super().__init__()
        self.emitter = emitter
    
    def emit(self, record: logging.LogRecord):
        """发送日志记录"""
        try:
            msg = self.format(record)
            level = record.levelname
            self.emitter.emit(level, msg)
        except Exception:
            self.handleError(record)


def setup_logger(name: str = 'docformatter', 
                 log_file: Optional[str] = None,
                 level: int = logging.INFO) -> logging.Logger:
    """
    设置日志记录器
    
    Args:
        name: 日志记录器名称
        log_file: 日志文件路径（可选）
        level: 日志级别
    
    Returns:
        logging.Logger: 配置好的日志记录器
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # 清除已有的处理器
    logger.handlers.clear()
    
    # 格式
    formatter = logging.Formatter(
        '[%(asctime)s] [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 文件处理器（可选）
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str = 'docformatter') -> logging.Logger:
    """获取日志记录器"""
    return logging.getLogger(name)
