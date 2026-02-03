from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean
from datetime import datetime

class IntranetGscModulos(BASE):

    __tablename__= "intranet_gsc_modulos"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    codigo = Column(String(10), nullable=False, unique=True)
    nombre = Column(String(100), nullable=False)
    descripcion = Column(String(500))
    color_clase = Column(String(50))
    orden = Column(Integer, nullable=False, default=0)
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.now)
    fecha_actualizacion = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.codigo = data.get('codigo')
        self.nombre = data.get('nombre')
        self.descripcion = data.get('descripcion')
        self.color_clase = data.get('color_clase')
        self.orden = data.get('orden', 0)
        self.activo = data.get('activo', True)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'codigo': self.codigo,
            'nombre': self.nombre,
            'descripcion': self.descripcion,
            'color_clase': self.color_clase,
            'orden': self.orden,
            'activo': self.activo
        }
