from sqlalchemy import Column, Integer, String, Boolean, DateTime
from sqlalchemy.sql import func
from Config.db import BASE

class IntranetGscOrigenesPlataforma(BASE):
    __tablename__ = 'intranet_gsc_origenes_plataforma'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(50), nullable=False, unique=True)
    activo = Column(Boolean, nullable=False, default=True)
    fecha_creacion = Column(DateTime, nullable=False, server_default=func.getdate())
    fecha_actualizacion = Column(DateTime, nullable=True, onupdate=func.getdate())
    
    def __init__(self, nombre, activo=True):
        self.nombre = nombre
        self.activo = activo
    
    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'activo': self.activo
        }
