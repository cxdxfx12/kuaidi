# ---------------------------------------------------------------------------
# PyInstaller 3.6 + setuptools 兼容性运行时钩子
# 解决：pkg_resources.py2_warn 在新版 setuptools 中已被移除
# 但 pyi_rth_pkgres.py 仍会尝试加载它，导致 ModuleNotFoundError
# ---------------------------------------------------------------------------
import sys

# 如果 setuptools 太新（没有 py2_warn），注入一个空模块占位
if 'pkg_resources.py2_warn' not in sys.modules:
    import types
    fake_mod = types.ModuleType('pkg_resources.py2_warn')
    sys.modules['pkg_resources.py2_warn'] = fake_mod

# 然后安全地执行原始的 pkg_resources 注册逻辑
try:
    import pkg_resources as res
    from pyimod03_importers import FrozenImporter
    res.register_loader_type(FrozenImporter, res.NullProvider)
except Exception:
    pass
