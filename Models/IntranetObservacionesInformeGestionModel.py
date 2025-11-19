from sqlalchemy import Column, BigInteger, Integer, String, DateTime, SmallInteger
from sqlalchemy.sql import func
from Config.db import BASE

class IntranetObservacionesInformeGestionModel(BASE):
    __tablename__ = 'intranet_observaciones_informe_gestion'
    
    id = Column(BigInteger, primary_key=True, autoincrement=True)
    anio = Column(Integer, nullable=False)
    mes = Column(Integer, nullable=False)
    observaciones = Column(String, nullable=True)
    estado = Column(SmallInteger, nullable=False, default=1)
    created_at = Column(DateTime, nullable=False, default=func.getdate())
    
    def __init__(self, data):
        self.anio = data.get('anio')
        self.mes = data.get('mes')
        self.observaciones = data.get('observaciones')
        self.estado = data.get('estado', 1)
    
    def to_dict(self):
        return {
            'id': self.id,
            'anio': self.anio,
            'mes': self.mes,
            'observaciones': self.observaciones,
            'estado': self.estado,
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S') if self.created_at else None
        }
