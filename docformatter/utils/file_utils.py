"""
文件工具模块
文件路径处理、批量文件扫描等
"""

from pathlib import Path
from typing import List, Optional
import os


def scan_docx_files(directory: str, recursive: bool = False) -> List[str]:
    """
    扫描目录中的所有 docx 文件
    
    Args:
        directory: 目录路径
        recursive: 是否递归扫描子目录
    
    Returns:
        List[str]: docx文件路径列表
    """
    dir_path = Path(directory)
    
    if not dir_path.exists() or not dir_path.is_dir():
        return []
    
    if recursive:
        files = list(dir_path.rglob("*.docx"))
    else:
        files = list(dir_path.glob("*.docx"))
    
    return [str(f.absolute()) for f in files]


def get_output_path(input_path: str, output_dir: str, suffix: str = "_formatted") -> str:
    """
    生成输出文件路径
    
    Args:
        input_path: 输入文件路径
        output_dir: 输出目录
        suffix: 文件名后缀
    
    Returns:
        str: 输出文件完整路径
    """
    input_file = Path(input_path)
    output_dir_path = Path(output_dir)
    
    # 确保输出目录存在
    output_dir_path.mkdir(parents=True, exist_ok=True)
    
    # 生成输出文件名
    output_name = f"{input_file.stem}{suffix}{input_file.suffix}"
    
    return str(output_dir_path / output_name)


def ensure_dir(path: str) -> Path:
    """
    确保目录存在，不存在则创建
    
    Args:
        path: 目录路径
    
    Returns:
        Path: 目录Path对象
    """
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_file_size(path: str) -> int:
    """
    获取文件大小（字节）
    
    Args:
        path: 文件路径
    
    Returns:
        int: 文件大小，文件不存在返回0
    """
    try:
        return Path(path).stat().st_size
    except Exception:
        return 0


def format_file_size(size_bytes: int) -> str:
    """
    格式化文件大小为可读字符串
    
    Args:
        size_bytes: 字节数
    
    Returns:
        str: 格式化后的大小字符串，如 "1.5 MB"
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def is_valid_docx(path: str) -> bool:
    """
    检查文件是否是有效的 docx 文件
    
    Args:
        path: 文件路径
    
    Returns:
        bool: 是否是有效的docx文件
    """
    p = Path(path)
    
    # 检查扩展名
    if p.suffix.lower() != '.docx':
        return False
    
    # 检查文件是否存在
    if not p.exists():
        return False
    
    # 检查文件大小（不能为0）
    if p.stat().st_size == 0:
        return False
    
    return True
