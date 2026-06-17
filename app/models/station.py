"""
网点表 - 申通快递的各级网点
"""
from sqlalchemy import Column, Integer, String, DateTime, Numeric, Boolean
from datetime import datetime
from app.models.database import Base


class Station(Base):
    """网点表"""
    __tablename__ = "stations"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    station_code = Column(String(50), unique=True, nullable=False, comment="网点编码")
    station_name = Column(String(100), nullable=False, comment="网点名称")
    parent_code = Column(String(50), nullable=True, comment="上级网点编码")
    region_code = Column(String(50), nullable=True, comment="所属区域编码")
    address = Column(String(255), nullable=True, comment="网点地址")
    contact_person = Column(String(50), nullable=True, comment="联系人")
    contact_phone = Column(String(20), nullable=True, comment="联系电话")
    commission_rate = Column(Numeric(5, 4), default=0.15, comment="网点分成比例（0~1）")
    is_active = Column(Boolean, default=True, comment="是否启用")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment="更新时间")

    def __repr__(self):
        return f"<Station(code={self.station_code}, name={self.station_name})>"
