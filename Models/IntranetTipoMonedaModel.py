from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean
from datetime import datetime

class IntranetTipoMoneda(BASE):

    __tablename__= "intranet_tipo_moneda"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    codigo = Column(String(10), nullable=False, unique=True)
    nombre = Column(String(50), nullable=False)
    simbolo = Column(String(10))
    activo = Column(Boolean, default=True)
    fecha_creacion = Column(DateTime, default=datetime.now)
    fecha_actualizacion = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.codigo = data.get('codigo')
        self.nombre = data.get('nombre')
        self.simbolo = data.get('simbolo')
        self.activo = data.get('activo', True)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'codigo': self.codigo,
            'nombre': self.nombre,
            'simbolo': self.simbolo,
            'activo': self.activo
        }
