"""
用户认证模块（内置账号 + 过期机制）
- 内置两个账号（首次启动自动创建，无需用户手动创建）
  - user / user → 3 个月免费
  - admin / admin → 6 个月使用
- 密码加盐哈希（pbkdf2_hmac，标准库自带）
- 支持"记住我"Token（30天免输入密码）
- 账号到期后提示"软件已到期，请联系管理员"，拒绝登录
"""
import os
import base64
import hashlib
import secrets
from datetime import datetime, timedelta
from typing import Tuple, Optional, List
from sqlalchemy import Column, Integer, String, DateTime, ForeignKey
from sqlalchemy.orm import relationship

from app.models.database import Base, get_session


# ===================== 常量 =====================

PBKDF2_ITERATIONS = 100_000
HASH_ALGO = "sha256"
TOKEN_REMEMBER_DAYS = 30   # "记住我" 有效期 30 天

# 内置账号
BUILTIN_ACCOUNTS = [
    {"username": "user",  "password": "user",  "role": "user",   "months": 3},
    {"username": "admin", "password": "admin", "role": "admin", "months": 6},
]

# 角色 → 有效期（月）
ROLE_DURATION_MONTHS = {
    "user": 3,
    "admin": 6,
}

MESSAGE_EXPIRED = "软件已到期，请联系管理员"
MESSAGE_WRONG = "用户名或密码错误"
MESSAGE_DISABLED = "该账号已被禁用"


# ===================== 模型 =====================

class User(Base):
    """用户表"""
    __tablename__ = "users"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(64), unique=True, nullable=False)
    password_hash = Column(String(256), nullable=False)
    salt = Column(String(128), nullable=False)
    role = Column(String(16), nullable=False, default="user")
    is_active = Column(Integer, nullable=False, default=1)
    last_login = Column(DateTime)
    created_at = Column(DateTime, default=datetime.now)


class LoginToken(Base):
    """"记住我" Token 表"""
    __tablename__ = "login_tokens"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    token_hash = Column(String(256), nullable=False)
    expires_at = Column(DateTime, nullable=False)
    created_at = Column(DateTime, default=datetime.now)

    user = relationship("User")


# ===================== 工具函数 =====================

def _add_months(dt: datetime, months: int) -> datetime:
    """将日期加 n 个月（不依赖 dateutil）"""
    # 简单实现：用 timedelta 近似
    # 更精确地：手动计算年月
    year = dt.year
    month = dt.month + months
    day = dt.day
    while month > 12:
        month -= 12
        year += 1
    # 处理月末溢出（如 1月31日 + 1个月 → 3月3日 → 改 2月28日）
    try:
        return dt.replace(year=year, month=month, day=day)
    except ValueError:
        # 2月30日这种情况 → 退到月末
        if month == 12:
            next_month = datetime(year + 1, 1, 1)
        else:
            next_month = datetime(year, month + 1, 1)
        last_day = (next_month - timedelta(days=1)).day
        return dt.replace(year=year, month=month, day=min(day, last_day))


def get_expire_date(user) -> datetime:
    """返回某账号的到期日期"""
    months = ROLE_DURATION_MONTHS.get(user.role, 3)
    return _add_months(user.created_at or datetime.now(), months)


def get_remaining_days(user) -> int:
    """返回剩余天数（负数表示已过期）"""
    exp = get_expire_date(user)
    return (exp.date() - datetime.now().date()).days


def _hash_password(password: str, salt: bytes) -> bytes:
    return hashlib.pbkdf2_hmac(
        HASH_ALGO, password.encode("utf-8"), salt, PBKDF2_ITERATIONS, dklen=32
    )


def make_password_hash(password: str) -> Tuple[str, str]:
    """生成新密码的 (hash_b64, salt_b64)"""
    salt = secrets.token_bytes(16)
    hashed = _hash_password(password, salt)
    return base64.b64encode(hashed).decode(), base64.b64encode(salt).decode()


def verify_password(password: str, password_hash_b64: str, salt_b64: str) -> bool:
    """验证密码"""
    try:
        salt = base64.b64decode(salt_b64)
        expected = base64.b64decode(password_hash_b64)
        actual = _hash_password(password, salt)
        return secrets.compare_digest(actual, expected)
    except Exception:
        return False


# ===================== Token（记住我） =====================

def make_remember_token() -> Tuple[str, str, datetime]:
    """返回 (raw_token, token_hash, expires_at)"""
    raw_token = secrets.token_hex(32)
    token_hash = hashlib.sha256(raw_token.encode()).hexdigest()
    expires_at = datetime.now() + timedelta(days=TOKEN_REMEMBER_DAYS)
    return raw_token, token_hash, expires_at


def _credential_key() -> str:
    return "SFeeSystem/remember_token"


def save_remember_token(username: str, raw_token: str) -> bool:
    """保存记住我的 token（优先 Windows Credential Manager）"""
    try:
        import win32cred
        payload = f"{username}|{raw_token}"
        win32cred.CredWrite({
            "Type": win32cred.CRED_TYPE_GENERIC,
            "TargetName": _credential_key(),
            "UserName": username,
            "CredentialBlob": payload.encode("utf-8"),
            "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        })
        return True
    except Exception:
        pass

    # 回退：APPDATA 下文件
    try:
        from app.models.path_config import get_user_data_dir
        fpath = os.path.join(get_user_data_dir(), ".remember_token")
        payload = f"{username}|{raw_token}"
        encoded = base64.b64encode(payload.encode("utf-8")).decode()
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(encoded)
        try:
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(fpath, 2)  # 2 = HIDDEN
        except Exception:
            pass
        return True
    except Exception:
        return False


def load_remember_token() -> Optional[Tuple[str, str]]:
    """读取记住我的 token"""
    try:
        import win32cred
        cred = win32cred.CredRead(_credential_key(), win32cred.CRED_TYPE_GENERIC, 0)
        if cred and cred.get("CredentialBlob"):
            blob = cred["CredentialBlob"]
            payload = blob.decode("utf-8") if isinstance(blob, bytes) else str(blob)
            if "|" in payload:
                u, t = payload.split("|", 1)
                return u, t
    except Exception:
        pass

    try:
        from app.models.path_config import get_user_data_dir
        fpath = os.path.join(get_user_data_dir(), ".remember_token")
        if not os.path.exists(fpath):
            return None
        with open(fpath, "r", encoding="utf-8") as f:
            encoded = f.read().strip()
        payload = base64.b64decode(encoded).decode("utf-8")
        if "|" in payload:
            u, t = payload.split("|", 1)
            return u, t
    except Exception:
        pass
    return None


def clear_remember_token() -> None:
    """清除记住我"""
    try:
        import win32cred
        win32cred.CredDelete(_credential_key(), win32cred.CRED_TYPE_GENERIC, 0)
    except Exception:
        pass
    try:
        from app.models.path_config import get_user_data_dir
        fpath = os.path.join(get_user_data_dir(), ".remember_token")
        if os.path.exists(fpath):
            os.remove(fpath)
    except Exception:
        pass


# ===================== UserService =====================

class UserService:

    @staticmethod
    def count_users() -> int:
        session = get_session()
        try:
            return session.query(User).count()
        except Exception:
            return 0
        finally:
            session.close()

    @staticmethod
    def ensure_builtin_accounts() -> None:
        """
        首次启动时：自动创建内置账号 user/user 和 admin/admin
        若用户已存在则跳过（不会覆盖已被修改的密码）
        """
        session = get_session()
        try:
            existing = {u.username: u for u in session.query(User).all()}
            created = False
            for acc in BUILTIN_ACCOUNTS:
                if acc["username"] not in existing:
                    pwd_hash, salt = make_password_hash(acc["password"])
                    user = User(
                        username=acc["username"],
                        password_hash=pwd_hash,
                        salt=salt,
                        role=acc["role"],
                        is_active=1,
                        created_at=datetime.now(),
                    )
                    session.add(user)
                    created = True
            if created:
                session.commit()
        except Exception:
            session.rollback()
        finally:
            session.close()

    @staticmethod
    def _check_expired(user) -> bool:
        """账号是否已过期"""
        return get_remaining_days(user) < 0

    @staticmethod
    def verify_login(username: str, password: str) -> Tuple[bool, str, Optional["User"]]:
        """
        验证用户名密码（同时检查过期）
        返回 (是否成功, 消息, User 对象或 None)
        """
        if not username or not password:
            return False, MESSAGE_WRONG, None

        # 确保首次启动时内置账号已创建
        if UserService.count_users() == 0:
            UserService.ensure_builtin_accounts()

        session = get_session()
        try:
            user = session.query(User).filter(User.username == username).first()
            if not user:
                return False, MESSAGE_WRONG, None
            if not user.is_active:
                return False, MESSAGE_DISABLED, None
            if not verify_password(password, user.password_hash, user.salt):
                return False, MESSAGE_WRONG, None
            # 检查是否过期
            if UserService._check_expired(user):
                return False, MESSAGE_EXPIRED, None
            user.last_login = datetime.now()
            session.commit()
            # 显式访问所有后续会用到的字段（防止 DetachedInstanceError）
            _ = user.id, user.username, user.role, user.created_at, user.is_active
            session.expunge(user)
            return True, "登录成功", user
        except Exception as e:
            return False, f"登录验证失败: {e}", None
        finally:
            session.close()

    @staticmethod
    def verify_remember_token(username: str, raw_token: str) -> Tuple[bool, str, Optional["User"]]:
        """
        验证"记住我" Token （同时检查过期）
        返回 (是否有效, 消息, User 对象或 None)
        """
        if UserService.count_users() == 0:
            UserService.ensure_builtin_accounts()

        session = get_session()
        try:
            user = session.query(User).filter(User.username == username).first()
            if not user or not user.is_active:
                return False, MESSAGE_WRONG, None
            token_hash = hashlib.sha256(raw_token.encode()).hexdigest()
            token_rec = (
                session.query(LoginToken)
                .filter(
                    LoginToken.user_id == user.id,
                    LoginToken.token_hash == token_hash,
                    LoginToken.expires_at > datetime.now(),
                )
                .first()
            )
            if token_rec:
                if UserService._check_expired(user):
                    return False, MESSAGE_EXPIRED, None
                user.last_login = datetime.now()
                session.commit()
                _ = user.id, user.username, user.role, user.created_at, user.is_active
                session.expunge(user)
                return True, "登录成功", user
            return False, MESSAGE_WRONG, None
        except Exception:
            return False, MESSAGE_WRONG, None
        finally:
            session.close()

    @staticmethod
    def save_token_for_user(user_id: int, token_hash: str, expires_at: datetime) -> bool:
        session = get_session()
        try:
            session.query(LoginToken).filter(LoginToken.user_id == user_id).delete()
            session.add(LoginToken(user_id=user_id, token_hash=token_hash, expires_at=expires_at))
            session.commit()
            return True
        except Exception:
            session.rollback()
            return False
        finally:
            session.close()

    @staticmethod
    def clear_tokens_for_user(user_id: int) -> None:
        session = get_session()
        try:
            session.query(LoginToken).filter(LoginToken.user_id == user_id).delete()
            session.commit()
        except Exception:
            session.rollback()
        finally:
            session.close()
