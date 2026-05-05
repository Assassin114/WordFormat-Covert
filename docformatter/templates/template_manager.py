"""
模板管理器 — 内置模板 + 用户模板 CRUD + 导入导出
"""

import os
import sys
import json
import shutil
from pathlib import Path
from datetime import date
from typing import Optional

from .template_io import load_template, save_template, template_to_dict
from ..models.template_config import TemplateConfig


def _user_data_dir() -> Path:
    """跨平台用户数据目录"""
    if sys.platform == "win32":
        base = Path(os.getenv("APPDATA", os.path.expanduser("~/AppData/Roaming")))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.getenv("XDG_DATA_HOME", str(Path.home() / ".local" / "share")))
    return base / "DocFormatter"


def _user_templates_dir() -> Path:
    d = _user_data_dir() / "templates"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _builtin_dir() -> Path:
    return Path(__file__).parent / "builtin"


# ══════════════════════════════════════════════════════
# 元信息读写
# ══════════════════════════════════════════════════════

def read_meta(path: str) -> dict:
    """从模板 JSON 文件读取 _meta，不存在则返回空 dict"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("_meta", {})


def write_meta(path: str, meta: dict):
    """将 _meta 写入模板 JSON 文件"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    data["_meta"] = meta
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _default_meta(name: str = "") -> dict:
    today = date.today().isoformat()
    return {
        "name": name or "未命名模板",
        "created": today,
        "modified": today,
        "source_docs": [],
        "tags": [],
        "notes": "",
    }


# ══════════════════════════════════════════════════════
# TemplateManager
# ══════════════════════════════════════════════════════

class TemplateManager:
    """管理内置和用户模板的发现、加载、保存、导入导出"""

    def __init__(self):
        self._builtin = _builtin_dir()
        self._user = _user_templates_dir()

    @property
    def builtin_dir(self) -> Path:
        return self._builtin

    @property
    def user_dir(self) -> Path:
        return self._user

    # ── 列表 ──

    def list_templates(self) -> list:
        """返回所有可用模板: [{"name", "path", "meta", "source"}, ...]"""
        result = []
        seen_names = set()

        # 用户模板优先
        if self._user.exists():
            for f in sorted(self._user.glob("*.json")):
                meta = read_meta(str(f))
                name = meta.get("name", f.stem)
                result.append({
                    "name": name,
                    "path": str(f),
                    "meta": meta,
                    "source": "user",
                })
                seen_names.add(name)

        # 内置模板（同名时用户模板覆盖）
        if self._builtin.exists():
            for f in sorted(self._builtin.glob("*.json")):
                meta = read_meta(str(f))
                meta["builtin"] = True
                name = meta.get("name", f.stem)
                if name in seen_names:
                    continue
                result.append({
                    "name": name,
                    "path": str(f),
                    "meta": meta,
                    "source": "builtin",
                })

        return result

    def list_user_templates(self) -> list:
        return [t for t in self.list_templates() if t["source"] == "user"]

    # ── 加载 ──

    def load_template(self, path: str) -> TemplateConfig:
        return load_template(path)

    def get_default_template(self) -> TemplateConfig:
        """获取默认模板（优先用户默认，否则内置默认）"""
        # 查用户目录下名称为"默认模板"的文件
        for t in self.list_user_templates():
            if t["name"] == "默认模板":
                return self.load_template(t["path"])
        # 回退到内置默认
        default_path = self._builtin / "default.json"
        if default_path.exists():
            return self.load_template(str(default_path))
        from ..models.template_config import create_default_template
        return create_default_template()

    # ── 保存 ──

    def save_as(self, config: TemplateConfig, name: str, meta_overrides: dict = None) -> str:
        """另存模板到用户目录，返回路径"""
        path = self._user / f"{name}.json"
        data = template_to_dict(config)

        meta = _default_meta(name)
        if meta_overrides:
            meta.update(meta_overrides)
        meta["modified"] = date.today().isoformat()
        data["_meta"] = meta

        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return str(path)

    def save(self, config: TemplateConfig, path: str):
        """保存到指定路径，更新 modified 时间"""
        data = template_to_dict(config)
        try:
            meta = read_meta(path)
        except (FileNotFoundError, json.JSONDecodeError):
            meta = _default_meta()
        meta["modified"] = date.today().isoformat()
        data["_meta"] = meta
        save_template(config, path)
        # 重新写入以包含 _meta
        with open(path, "r", encoding="utf-8") as f:
            current = json.load(f)
        current["_meta"] = meta
        with open(path, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False, indent=2)

    # ── 删除 ──

    def delete(self, path: str) -> bool:
        p = Path(path)
        if not p.exists() or str(p.parent) != str(self._user):
            return False
        p.unlink()
        return True

    # ── 导出导入 ──

    def export_template(self, path: str, export_path: str):
        """导出模板到外部文件"""
        shutil.copy2(path, export_path)

    def import_template(self, file_path: str) -> Optional[str]:
        """导入外部模板到用户目录，返回新路径"""
        src = Path(file_path)
        if not src.exists():
            return None
        dst = self._user / src.name
        # 不覆盖同名文件，加后缀
        if dst.exists():
            dst = self._user / f"{src.stem}_imported{src.suffix}"
        shutil.copy2(file_path, str(dst))
        # 确保有 _meta
        try:
            meta = read_meta(str(dst))
            if "created" not in meta:
                meta = _default_meta(dst.stem)
                write_meta(str(dst), meta)
        except json.JSONDecodeError:
            pass
        return str(dst)
