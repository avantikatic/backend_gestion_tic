from sqlalchemy import Column, Integer, String, Boolean, DateTime
from sqlalchemy.sql import func
from Config.db import BASE

class IntranetGscRiesgos(BASE):
    __tablename__ = 'intranet_gsc_riesgos'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(20), nullable=False, unique=True)
    orden = Column(Integer, nullable=False, default=0)
    activo = Column(Boolean, nullable=False, default=True)
    fecha_creacion = Column(DateTime, nullable=False, server_default=func.getdate())
    fecha_actualizacion = Column(DateTime, nullable=True, onupdate=func.getdate())
    
    def __init__(self, nombre, orden=0, activo=True):
        self.nombre = nombre
        self.orden = orden
        self.activo = activo
    
    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'orden': self.orden,
            'activo': self.activo
        }
