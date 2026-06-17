"""
数据路径管理（跨平台 + PyInstaller onefile/onedir 兼容）：
- 用户数据（数据库、设置等）：保存在用户可写目录
- 资源文件（图标、默认配置）：从 PyInstaller 打包目录 (_MEIPASS) 或源码目录查找
- 跨平台：Windows (APPDATA)、macOS (~/Library/Application Support)、Linux (~/.local/share)
"""
import os
import sys
import shutil

_resolved_user_data_dir = None
_resolved_data_base = None
_resolved_bundle_dir = None


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def _get_meipass() -> str:
    """PyInstaller onefile 模式临时解压目录"""
    return getattr(sys, "_MEIPASS", None)


def get_app_root() -> str:
    """程序可执行文件所在目录"""
    if _is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_bundle_dir() -> str:
    """打包后的资源根目录。
    onefile 模式 → sys._MEIPASS（临时解压目录）
    onedir 模式  → sys.executable 所在目录（与 data 同级）
    源码模式   → 项目根目录（与 app/ 同级）
    """
    global _resolved_bundle_dir
    if _resolved_bundle_dir:
        return _resolved_bundle_dir

    candidates = []
    meipass = _get_meipass()
    if meipass:
        candidates.append(meipass)
    if _is_frozen():
        candidates.append(os.path.dirname(sys.executable))
    # 源码模式回退
    candidates.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

    for c in candidates:
        if c and os.path.isdir(c):
            # 验证它里面是否有我们关心的 data 目录
            if os.path.isdir(os.path.join(c, "data")):
                _resolved_bundle_dir = c
                return c

    # 如果都没有 data 子目录，返回第一个候选
    _resolved_bundle_dir = candidates[0] if candidates else get_app_root()
    return _resolved_bundle_dir


def _test_writable(path: str) -> bool:
    try:
        os.makedirs(path, exist_ok=True)
        test_file = os.path.join(path, "_write_test_" + str(os.getpid()) + ".tmp")
        with open(test_file, "wb") as f:
            f.write(b"ok")
        os.remove(test_file)
        return True
    except Exception:
        return False


def get_user_data_dir() -> str:
    """用户数据根目录（写权限）"""
    global _resolved_user_data_dir
    if _resolved_user_data_dir:
        return _resolved_user_data_dir

    candidates = []
    home = os.path.expanduser("~")

    # macOS: ~/Library/Application Support/SFeeSystem
    if sys.platform == "darwin":
        if home and home != "~":
            candidates.append(os.path.join(home, "Library", "Application Support", "SFeeSystem"))

    # Windows: %APPDATA%\SFeeSystem, %LOCALAPPDATA%\SFeeSystem
    appdata = os.environ.get("APPDATA")
    if appdata:
        candidates.append(os.path.join(appdata, "SFeeSystem"))
    localappdata = os.environ.get("LOCALAPPDATA")
    if localappdata:
        candidates.append(os.path.join(localappdata, "SFeeSystem"))

    # 跨平台通用：用户主目录
    if home and home != "~":
        candidates.append(os.path.join(home, ".SFeeSystem"))
        candidates.append(os.path.join(home, "SFeeSystem"))

    # 便携模式：程序所在目录
    candidates.append(os.path.join(get_app_root(), "SFeeSystem_data"))

    # 系统临时目录（最后回退）
    import tempfile
    candidates.append(os.path.join(tempfile.gettempdir(), "SFeeSystem"))

    chosen = None
    for p in candidates:
        if _test_writable(p):
            chosen = p
            break
    if not chosen:
        chosen = candidates[-1]

    _resolved_user_data_dir = chosen
    return chosen


def get_data_dir(*sub_paths: str) -> str:
    """用户数据的子目录（可写）"""
    global _resolved_data_base
    if not _resolved_data_base:
        base = os.path.join(get_user_data_dir(), "data")
        base = os.path.abspath(base)
        if _test_writable(base):
            _resolved_data_base = base
        else:
            alt = os.path.join(get_app_root(), "SFeeSystem_data", "data")
            if _test_writable(alt):
                _resolved_data_base = alt
            else:
                import tempfile
                tmp = os.path.join(tempfile.gettempdir(), "SFeeSystem", "data")
                _test_writable(tmp)
                _resolved_data_base = tmp

    path = _resolved_data_base if not sub_paths else os.path.join(_resolved_data_base, *sub_paths)
    if path != _resolved_data_base:
        _test_writable(path)
    return path


def get_resource_path(*sub_paths: str) -> str:
    """获取**打包资源**路径（只读，图标/默认配置等）。
    优先在 bundle dir 中查找，找不到则回退到项目根。
    """
    bundle = get_bundle_dir()
    candidates = [bundle, get_app_root()]

    for base in candidates:
        full = os.path.join(base, *sub_paths) if sub_paths else base
        if os.path.exists(full):
            return full

    # 都不存在，返回 bundle + sub_paths 的组合路径（让上层处理不存在的情况）
    return os.path.join(bundle, *sub_paths) if sub_paths else bundle


def get_db_path() -> str:
    return os.path.join(get_data_dir(), "app.db")


def get_config_file(filename: str) -> str:
    cfg_dir = get_data_dir("config")
    return os.path.join(cfg_dir, filename)


def _copy_file_if_not_exists(src: str, dst: str) -> None:
    if os.path.exists(src) and not os.path.exists(dst):
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
        except Exception:
            pass


def bootstrap_user_data() -> None:
    """首次启动时：把打包资源中的默认配置、图标复制到用户数据目录"""
    if not _is_frozen():
        return

    bundle = get_bundle_dir()
    user_root = get_user_data_dir()

    try:
        for d in ("config", "exports", "uploads", "icons"):
            get_data_dir(d)
    except Exception:
        pass

    # 复制默认配置
    for name in ("fee_rules.json", "default_settings.json"):
        src = os.path.join(bundle, "data", "config", name)
        dst = os.path.join(user_root, "data", "config", name)
        _copy_file_if_not_exists(src, dst)

    # 复制图标（用于 QIcon 加载）
    icon_src = os.path.join(bundle, "data", "icons")
    icon_dst = get_data_dir("icons")
    if os.path.isdir(icon_src):
        for fn in os.listdir(icon_src):
            _copy_file_if_not_exists(
                os.path.join(icon_src, fn),
                os.path.join(icon_dst, fn),
            )


try:
    bootstrap_user_data()
except Exception:
    pass
