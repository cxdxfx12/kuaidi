"""
分成规则表 - 网点/快递员的分成配置
"""
from sqlalchemy import Column, Integer, String, DateTime, Numeric, Boolean
from datetime import datetime
from app.models.database import Base


class CommissionRule(Base):
    """分成规则表"""
    __tablename__ = "commission_rules"
    __table_args__ = {"extend_existing": True}

    id = Column(Integer, primary_key=True, autoincrement=True)
    rule_name = Column(String(100), nullable=False, comment="规则名称")
    target_type = Column(String(20), nullable=False, comment="目标类型：station/courier")
    target_code = Column(String(50), nullable=True, comment="目标编码（NULL表示全局默认）")
    commission_rate = Column(Numeric(5, 4), default=0.15, comment="分成比例")
    is_active = Column(Boolean, default=True, comment="是否启用")
    effective_date = Column(DateTime, nullable=True, comment="生效日期")
    expire_date = Column(DateTime, nullable=True, comment="失效日期")
    description = Column(String(255), nullable=True, comment="规则描述")
    created_at = Column(DateTime, default=datetime.now, comment="创建时间")

    def __repr__(self):
        return f"<CommissionRule(name={self.rule_name}, rate={self.commission_rate})>"
