from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean
from datetime import datetime

class IntranetContingenciaEstadosDocumento(BASE):

    __tablename__ = "intranet_contingencia_estados_documento"

    id          = Column(Integer, primary_key=True, autoincrement=True)
    nombre      = Column(String(100), nullable=False)
    descripcion = Column(String(300), nullable=True)
    orden       = Column(Integer, default=1, nullable=False)
    activo      = Column(Boolean, default=True, nullable=False)
    created_at  = Column(DateTime, default=datetime.now)

    def __init__(self, data: dict):
        self.nombre      = data.get('nombre')
        self.descripcion = data.get('descripcion')
        self.orden       = data.get('orden', 1)
        self.activo      = data.get('activo', True)

    def to_dict(self):
        return {
            'id':          self.id,
            'nombre':      self.nombre,
            'descripcion': self.descripcion,
            'orden':       self.orden,
            'activo':      self.activo,
        }
