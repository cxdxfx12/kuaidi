"""
派费明细表 - 每一行Excel数据对应一条派费明细
"""
from sqlalchemy import Column, Integer, String, DateTime, Numeric, Text, Boolean, ForeignKey, JSON
from sqlalchemy.orm import relationship
from datetime import datetime
from app.models.database import Base


class FeeDetail(Base):
    """派费明细表"""
    __tablename__ = "fee_details"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    record_id = Column(Integer, ForeignKey("fee_records.id"), nullable=False, comment="关联计算记录")
    row_index = Column(Integer, default=0, comment="原始Excel行号")

    # 申通派费结算Excel常见字段
    tracking_no = Column(String(50), nullable=True, comment="快递单号/运单号")
    station_code = Column(String(50), nullable=True, comment="网点编码")
    station_name = Column(String(100), nullable=True, comment="网点名称")
    courier_code = Column(String(50), nullable=True, comment="快递员编码")
    courier_name = Column(String(100), nullable=True, comment="快递员姓名")
    region_code = Column(String(50), nullable=True, comment="区域编码")
    region_name = Column(String(100), nullable=True, comment="区域名称")
    weight = Column(Numeric(10, 3), default=0, comment="重量（kg）")
    quantity = Column(Integer, default=1, comment="件数")
    service_type = Column(String(20), nullable=True, comment="服务类型")

    # 原始数据（JSON存储所有原始列）
    original_data = Column(JSON, nullable=True, comment="原始行数据")

    # 计算结果
    calculated_fee = Column(Numeric(12, 2), default=0, comment="本行派费")
    rule_name = Column(String(100), nullable=True, comment="应用的规则名称")

    # 异常标记
    is_exception = Column(Boolean, default=False, comment="是否异常")
    exception_type = Column(String(50), nullable=True, comment="异常类型")
    remark = Column(String(500), nullable=True, comment="备注")

    created_at = Column(DateTime, default=datetime.now, comment="创建时间")

    # 关联
    record = relationship("FeeRecord", back_populates="details")

    def __repr__(self):
        return f"<FeeDetail(id={self.id}, tracking={self.tracking_no}, fee={self.calculated_fee})>"
