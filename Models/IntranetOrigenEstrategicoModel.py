from sqlalchemy import Column, Integer, String, TIMESTAMP, text, DateTime
from Config.db import BASE
from datetime import datetime

class IntranetOrigenEstrategicoModel(BASE):
    __tablename__ = 'intranet_origen_estrategico'

    id = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(255), nullable=False)
    estado = Column(Integer, nullable=False, default=1)
    created_at = Column(DateTime, default=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'estado': self.estado,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }
