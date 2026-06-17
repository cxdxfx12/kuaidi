"""
计算记录表 - 每次导入的Excel文件生成一条记录
"""
from sqlalchemy import Column, Integer, String, DateTime, Numeric, Text
from sqlalchemy.orm import relationship
from datetime import datetime
from app.models.database import Base


class FeeRecord(Base):
    """计算记录表"""
    __tablename__ = "fee_records"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    file_name = Column(String(255), nullable=False, comment="原始文件名")
    file_path = Column(String(500), nullable=False, comment="文件存储路径")
    file_size = Column(Integer, default=0, comment="文件大小（字节）")
    total_rows = Column(Integer, default=0, comment="数据总行数")
    success_rows = Column(Integer, default=0, comment="成功计算行数")
    error_rows = Column(Integer, default=0, comment="异常行数")
    total_fee = Column(Numeric(12, 2), default=0, comment="派费总额")
    status = Column(String(20), default="pending", comment="状态：pending/success/failed")
    error_message = Column(Text, nullable=True, comment="错误信息")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")
    completed_at = Column(DateTime, nullable=True, comment="完成时间")

    # 关联明细
    details = relationship("FeeDetail", back_populates="record", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<FeeRecord(id={self.id}, file={self.file_name}, total_fee={self.total_fee})>"
