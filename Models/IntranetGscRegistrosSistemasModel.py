from Config.db import BASE
from sqlalchemy import Column, Integer, DateTime, ForeignKey
from datetime import datetime

class IntranetGscRegistrosSistemas(BASE):
    """
    Modelo para la tabla relacional muchos a muchos
    Relaciona registros GSC con sistemas afectados
    """

    __tablename__ = "intranet_gsc_registros_sistemas"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id', ondelete='CASCADE'), nullable=False)
    id_sistema = Column(Integer, ForeignKey('intranet_gsc_sistemas_afectados.id'), nullable=False)
    fecha_asignacion = Column(DateTime, default=datetime.now, nullable=False)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.id_sistema = data.get('id_sistema')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'id_sistema': self.id_sistema,
            'fecha_asignacion': self.fecha_asignacion.isoformat() if self.fecha_asignacion else None
        }
