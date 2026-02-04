from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, ForeignKey
from datetime import datetime

class IntranetGscEvidenciasAlerta(BASE):
    """
    Modelo para evidencias de tipo ALERTA (Plataforma)
    Almacena información específica de alertas de seguridad/monitoreo
    """

    __tablename__ = "intranet_gsc_evidencias_alerta"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_evidencia = Column(Integer, ForeignKey('intranet_gsc_evidencias.id', ondelete='CASCADE'), nullable=False, unique=True)
    id_origen_plataforma = Column(Integer, ForeignKey('intranet_gsc_origenes_plataforma.id'), nullable=False)
    nombre_alerta = Column(String(255), nullable=False)
    severidad = Column(String(50))
    fecha_alerta = Column(DateTime)
    codigo_alerta = Column(String(100))
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_evidencia = data.get('id_evidencia')
        self.id_origen_plataforma = data.get('id_origen_plataforma')
        self.nombre_alerta = data.get('nombre_alerta')
        self.severidad = data.get('severidad')
        self.fecha_alerta = data.get('fecha_alerta')
        self.codigo_alerta = data.get('codigo_alerta')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_evidencia': self.id_evidencia,
            'id_origen_plataforma': self.id_origen_plataforma,
            'nombre_alerta': self.nombre_alerta,
            'severidad': self.severidad,
            'fecha_alerta': self.fecha_alerta.isoformat() if self.fecha_alerta else None,
            'codigo_alerta': self.codigo_alerta
        }
