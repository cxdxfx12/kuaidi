"""PyInstaller Runtime Hook: 提前屏蔽 pkg_resources 的有问题初始化"""
import sys
import os

# 1) 将 _MEIPASS 加入 sys.path，确保运行时能找到内部模块
meipass = getattr(sys, '_MEIPASS', None)
if meipass and meipass not in sys.path:
    sys.path.insert(0, meipass)

# 2) Patch pkg_resources：
#    pyi_rth_pkgres 失败的根源是 pkg_resources.require() / iter_entry_points()
#    在冻结环境下试图访问不存在的 .egg-info / dist-info 元数据。
#    这里在它运行前替换相关方法为安全的 no-op。
try:
    import pkg_resources
    # 记录原始方法
    _orig_require = pkg_resources.require
    _orig_iter_entry_points = pkg_resources.iter_entry_points
    _orig_get_distribution = pkg_resources.get_distribution

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
            class _FakeDist(object):
                version = "0.0.0"
                project_name = "unknown"
                def __str__(self):
                    return "unknown 0.0.0"
            return _FakeDist()

    # 应用 patch
    pkg_resources.require = _safe_require
    pkg_resources.iter_entry_points = _safe_iter_entry_points
    pkg_resources.get_distribution = _safe_get_distribution

    # 3) 防止 pkg_resources 的 working_set 初始化失败
    try:
        ws = pkg_resources.working_set
        ws.require = _safe_require
    except Exception:
        pass

except Exception:
    # pkg_resources 不存在时（被排除了），什么都不做
    pass

# 4) Patch setuptools 的 distutils 重定向（若存在）
try:
    import setuptools
except Exception:
    pass

# 5) 设置 QT 环境变量（防止 PyQt5 在某些环境下找不到插件）
if meipass:
    os.environ['QT_PLUGIN_PATH'] = os.path.join(meipass, 'PyQt5', 'Qt', 'plugins')
    os.environ['QML2_IMPORT_PATH'] = os.path.join(meipass, 'PyQt5', 'Qt', 'qml')
