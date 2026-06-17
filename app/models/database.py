"""
数据库模型定义（增强版）
- 数据库路径可写检测
- 数据库文件损坏检测 + 自动重建
- SQLite 调优参数（WAL/大缓存/mmap）
"""
from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker, declarative_base
import os
import shutil
import time

from app.models.path_config import get_db_path


def _safe_backup_and_rebuild(db_path: str) -> bool:
    """数据库损坏时：备份旧文件，删除 WAL/SHM，让 SQLite 重建"""
    try:
        # 备份主文件
        if os.path.exists(db_path):
            backup = db_path + ".corrupt_" + str(int(time.time()))
            try:
                shutil.copy2(db_path, backup)
            except Exception:
                pass
            try:
                os.remove(db_path)
            except Exception:
                return False

        # 清理 WAL/SHM
        for ext in ("-wal", "-shm", "-journal"):
            p = db_path + ext
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        return True
    except Exception:
        return False


def _build_engine(db_path: str):
    """构建 SQLAlchemy engine（绝对路径、check_same_thread=False、长超时）"""
    # 确保目录存在
    db_dir = os.path.dirname(db_path)
    try:
        os.makedirs(db_dir, exist_ok=True)
    except Exception:
        pass

    # SQLite 要求绝对路径，且 Windows 下用 / 分隔
    abs_path = os.path.abspath(db_path).replace("\\", "/")
    return create_engine(
        f"sqlite:///{abs_path}",
        echo=False,
        connect_args={
            "check_same_thread": False,
            "timeout": 120,  # 2分钟长超时，避免并发启动失败
        },
        pool_pre_ping=True,
        pool_size=5,
        max_overflow=10,
        future=True,
    )


def _apply_optimization_pragmas(connection):
    """尝试设置优化参数，失败则降级"""
    cursor = connection.cursor()
    try:
        # 1) WAL 模式（关键：写更快、读不阻塞）
        cursor.execute("PRAGMA journal_mode = WAL")
        # 2) 同步策略：NORMAL 比 FULL 快 5-10 倍，但仍有持久化保障
        cursor.execute("PRAGMA synchronous = NORMAL")
        # 3) 页面缓存：200MB（负值=KB）
        cursor.execute("PRAGMA cache_size = -200000")
        # 4) 临时对象放内存
        cursor.execute("PRAGMA temp_store = MEMORY")
        # 5) mmap：1GB（读/写经由 OS 文件缓存，减少用户态拷贝）
        cursor.execute("PRAGMA mmap_size = 1073741824")
        connection.commit()
    except Exception:
        # WAL/mmap 失败（某些文件系统不支持），降级为普通 DELETE 模式 + 基础缓存
        try:
            cursor.execute("PRAGMA journal_mode = DELETE")
            cursor.execute("PRAGMA synchronous = NORMAL")
            cursor.execute("PRAGMA cache_size = -64000")
            connection.commit()
        except Exception:
            pass
    finally:
        cursor.close()


# ===== 初始化 engine =====
DB_PATH = get_db_path()
_engine_attempts = 0

while True:
    try:
        engine = _build_engine(DB_PATH)
        # 测试连接：如果打开失败，尝试备份+重建
        test_conn = engine.raw_connection()
        try:
            test_cursor = test_conn.cursor()
            test_cursor.execute("PRAGMA integrity_check")
            result = test_cursor.fetchone()
            # integrity_check 返回 "ok" 表示数据库正常；其他返回都算损坏
            if result and (result[0] == "ok" or (isinstance(result[0], str) and result[0].lower() == "ok")):
                pass  # 数据库正常
            else:
                # 数据库损坏，重建
                test_conn.close()
                engine.dispose()
                _safe_backup_and_rebuild(DB_PATH)
                # 重建后再构建新 engine
                engine = _build_engine(DB_PATH)
                # 不再测试，直接使用
        except Exception:
            # 任何测试失败都视为损坏
            test_conn.close()
            engine.dispose()
            _safe_backup_and_rebuild(DB_PATH)
            engine = _build_engine(DB_PATH)
        else:
            test_conn.close()
        break  # 初始化成功
    except Exception as e:
        _engine_attempts += 1
        if _engine_attempts >= 2:
            # 第二次还是失败：尝试清理
            _safe_backup_and_rebuild(DB_PATH)
            engine = _build_engine(DB_PATH)
            break
        # 第一次失败：清理后重试
        _safe_backup_and_rebuild(DB_PATH)
        continue

# 连接事件：每次新建连接时应用优化 PRAGMA
@event.listens_for(engine, "connect")
def _on_connect(dbapi_conn, connection_record):
    _apply_optimization_pragmas(dbapi_conn)


# 会话工厂
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)

# 模型基类
Base = declarative_base()


# 是否已初始化标记
_db_initialized = False


def get_session():
    """获取数据库会话（自动建表）"""
    global _db_initialized
    if not _db_initialized:
        init_db()
        _db_initialized = True
    return SessionLocal()


def init_db():
    """初始化数据库（创建所有表）"""
    global engine, SessionLocal
    from app.models.fee_record import FeeRecord
    from app.models.fee_detail import FeeDetail
    from app.models.station import Station
    from app.models.courier import Courier
    from app.models.commission_rule import CommissionRule
    from app.models.column_mapping import ColumnMapping
    # 用户认证表（在 create_all 前 import，确保注册到 Base）
    from app.models.user import User, LoginToken

    try:
        Base.metadata.create_all(bind=engine)
    except Exception:
        # create_all 失败 → 数据库损坏 → 备份后重建
        _safe_backup_and_rebuild(DB_PATH)
        new_engine = _build_engine(DB_PATH)
        engine.dispose()
        engine = new_engine
        SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
        Base.metadata.create_all(bind=engine)
