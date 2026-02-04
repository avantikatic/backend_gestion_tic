from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Text, ForeignKey
from datetime import datetime

class IntranetGscEvidenciasOtro(BASE):
    """
    Modelo para evidencias de tipo OTRO
    Almacena información de evidencias que no encajan en otras categorías
    """

    __tablename__ = "intranet_gsc_evidencias_otro"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_evidencia = Column(Integer, ForeignKey('intranet_gsc_evidencias.id', ondelete='CASCADE'), nullable=False, unique=True)
    descripcion_tipo = Column(String(500))
    detalles = Column(Text)
    referencia = Column(String(500))
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_evidencia = data.get('id_evidencia')
        self.descripcion_tipo = data.get('descripcion_tipo')
        self.detalles = data.get('detalles')
        self.referencia = data.get('referencia')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_evidencia': self.id_evidencia,
            'descripcion_tipo': self.descripcion_tipo,
            'detalles': self.detalles,
            'referencia': self.referencia
        }
