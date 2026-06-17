"""
列名映射表 - 用于自动识别不同Excel的列名
"""
from sqlalchemy import Column, Integer, String, DateTime, Boolean, JSON
from datetime import datetime
from app.models.database import Base


class ColumnMapping(Base):
    """列名映射表"""
    __tablename__ = "column_mappings"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    standard_name = Column(String(50), unique=True, nullable=False, comment="标准列名（系统内部用）")
    alias_names = Column(JSON, nullable=False, comment="列名别名列表")
    is_required = Column(Boolean, default=False, comment="是否必填列")
    data_type = Column(String(20), default="string", comment="数据类型：string/number/decimal")
    default_value = Column(String(100), nullable=True, comment="默认值")
    description = Column(String(255), nullable=True, comment="列说明")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")

    def __repr__(self):
        return f"<ColumnMapping(standard={self.standard_name})>"
