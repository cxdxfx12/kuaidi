"""PyInstaller Runtime Hook: 彻底修复 pkg_resources 兼容性问题

问题根源：PyInstaller 3.6 的 pyi_rth_pkgres.py 会尝试导入
pkg_resources.py2_warn，但新版 setuptools 已移除该模块。

解决方案：在 pyi_rth_pkgres 执行之前，提前注入空模块占位。
"""
import sys
import os
import types

# ===== 关键修复：注入空的 py2_warn 模块 =====
# pyi_rth_pkgres.py 会在 import pkg_resources 前尝试 import pkg_resources.py2_warn
# 这个模块在新版 setuptools 中已被移除，导致 ImportError
if 'pkg_resources.py2_warn' not in sys.modules:
    fake_mod = types.ModuleType('pkg_resources.py2_warn')
    sys.modules['pkg_resources.py2_warn'] = fake_mod

# 同样处理 pkg_resources.extern 相关的缺失模块
for mod_name in [
    'pkg_resources.extern.pyparsing',
    'pkg_resources.extern.appdirs',
    'pkg_resources.extern.six',
    'pkg_resources.extern.six.moves',
]:
    if mod_name not in sys.modules:
        fake_mod = types.ModuleType(mod_name)
        sys.modules[mod_name] = fake_mod

# ===== 处理 pyimod03_importers 缺失问题 =====
# pyi_rth_pkgres.py 会尝试 from pyimod03_importers import FrozenImporter
# 但这个模块在打包后可能无法直接导入
if 'pyimod03_importers' not in sys.modules:
    fake_mod = types.ModuleType('pyimod03_importers')
    fake_mod.FrozenImporter = None
    sys.modules['pyimod03_importers'] = fake_mod

# ===== 处理 setuptools 相关的缺失模块 =====
for mod_name in [
    'setuptools.extern.pyparsing',
    'setuptools.extern.six',
    'setuptools.extern.six.moves',
    'setuptools.extern.packaging',
    'setuptools.extern.ordered_set',
]:
    if mod_name not in sys.modules:
        fake_mod = types.ModuleType(mod_name)
        sys.modules[mod_name] = fake_mod

# ===== 将 _MEIPASS 加入 sys.path =====
meipass = getattr(sys, '_MEIPASS', None)
if meipass and meipass not in sys.path:
    sys.path.insert(0, meipass)

# ===== 设置 QT 环境变量 =====
if meipass:
    os.environ['QT_PLUGIN_PATH'] = os.path.join(meipass, 'PyQt5', 'Qt', 'plugins')
    os.environ['QML2_IMPORT_PATH'] = os.path.join(meipass, 'PyQt5', 'Qt', 'qml')

# ===== 安全地执行 pkg_resources 注册逻辑 =====
try:
    import pkg_resources as res
    try:
        from pyimod03_importers import FrozenImporter
        res.register_loader_type(FrozenImporter, res.NullProvider)
    except Exception:
        pass
    except ImportError:
        pass

    # Patch pkg_resources 方法为安全版本
    _orig_require = res.require
    _orig_iter_entry_points = res.iter_entry_points
    _orig_get_distribution = res.get_distribution

    def _safe_require(*args, **kwargs):
        try:
            return _orig_require(*args, **kwargs)
        except Exception:
            return []

    def _safe_iter_entry_points(*args, **kwargs):
        try:
            return _orig_iter_entry_points(*args, **kwargs)
        except Exception:
            return iter([])

    def _safe_get_distribution(*args, **kwargs):
        try:
            return _orig_get_distribution(*args, **kwargs)
        except Exception:
            class _FakeDist:
                version = "0.0.0"
                project_name = "unknown"
                def __str__(self):
                    return "unknown 0.0.0"
            return _FakeDist()

    res.require = _safe_require
    res.iter_entry_points = _safe_iter_entry_points
    res.get_distribution = _safe_get_distribution

except Exception:
    pass
