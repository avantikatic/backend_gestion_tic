from Config.db import BASE
from sqlalchemy import Column, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetGscResultados(BASE):
    """
    Modelo para la tabla de resultados (bitácora) del módulo GSC
    Almacena múltiples entradas de texto relacionadas a un registro
    """

    __tablename__ = "intranet_gsc_resultados"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id'), nullable=False)
    texto = Column(Text, nullable=False)
    created_at = Column(DateTime, default=datetime.now, nullable=False)
    activo = Column(Boolean, default=True, nullable=False)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.texto = data.get('texto')
        self.created_at = data.get('created_at', datetime.now())
        self.activo = data.get('activo', True)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'texto': self.texto,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'activo': self.activo
        }
