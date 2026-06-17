"""
快递员表
"""
from sqlalchemy import Column, Integer, String, DateTime, Numeric, Boolean, ForeignKey
from datetime import datetime
from app.models.database import Base


class Courier(Base):
    """快递员表"""
    __tablename__ = "couriers"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    courier_code = Column(String(50), unique=True, nullable=False, comment="快递员编码")
    courier_name = Column(String(100), nullable=False, comment="快递员姓名")
    station_code = Column(String(50), nullable=True, comment="所属网点编码")
    phone = Column(String(20), nullable=True, comment="手机号")
    delivery_commission_rate = Column(Numeric(5, 4), default=0.80, comment="派件提成比例（快递员拿到派费的80%）")
    is_active = Column(Boolean, default=True, comment="是否启用")
    hire_date = Column(DateTime, nullable=True, comment="入职日期")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment="更新时间")

    def __repr__(self):
        return f"<Courier(code={self.courier_code}, name={self.courier_name})>"
