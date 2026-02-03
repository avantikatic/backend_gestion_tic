from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean
from datetime import datetime

class IntranetGscEstados(BASE):

    __tablename__= "intranet_gsc_estados"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(100), nullable=False)
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.now)
    fecha_actualizacion = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.nombre = data.get('nombre')
        self.activo = data.get('activo', True)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'nombre': self.nombre,
            'activo': self.activo
        }
