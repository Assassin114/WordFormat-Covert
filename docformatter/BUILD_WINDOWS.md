# Windows 打包说明

## 环境要求

- Windows 10/11
- Python 3.10+ 
- pip

## 安装构建依赖

```batch
cd docformatter
pip install -r requirements.txt
pip install pyinstaller
```

## 打包命令

```batch
pyinstaller build.spec --clean
```

打包完成后，`dist/DocFormatter.exe` 即为可执行文件。

## 依赖列表 (requirements.txt)

```
PyQt6>=6.6.0
python-docx>=1.1.0
lxml>=5.0.0
```

## 输出位置

```
dist/
└── DocFormatter.exe    # 主程序
```

## 注意事项

1. 首次运行可能需要安装 VC++ Redistributable
2. 图标文件 (icon.ico) 暂未提供，可自行添加
3. 模板文件在 `templates/` 目录，打包时已包含

## 快速测试（不打包）

```batch
python main.py
```
