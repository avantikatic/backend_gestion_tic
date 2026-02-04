from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Text, ForeignKey
from datetime import datetime

class IntranetGscRegistrosSeguridad(BASE):
    """
    Modelo para registros del módulo SEGURIDAD (SEG)
    Almacena información específica de incidentes de seguridad
    """

    __tablename__ = "intranet_gsc_registros_seguridad"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id', ondelete='CASCADE'), nullable=False, unique=True)
    fecha_hora_incidente = Column(DateTime, nullable=False)
    id_fuente_seguridad = Column(Integer, ForeignKey('intranet_gsc_fuentes_seguridad.id'), nullable=False)
    tipo_amenaza = Column(String(200))
    id_impacto = Column(Integer, ForeignKey('intranet_gsc_impactos.id'), nullable=False)
    responsable_tic = Column(String(200))
    acciones_tomadas = Column(Text)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.fecha_hora_incidente = data.get('fecha_hora_incidente')
        self.id_fuente_seguridad = data.get('id_fuente_seguridad')
        self.tipo_amenaza = data.get('tipo_amenaza')
        self.id_impacto = data.get('id_impacto')
        self.responsable_tic = data.get('responsable_tic')
        self.acciones_tomadas = data.get('acciones_tomadas')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'fecha_hora_incidente': self.fecha_hora_incidente.isoformat() if self.fecha_hora_incidente else None,
            'id_fuente_seguridad': self.id_fuente_seguridad,
            'tipo_amenaza': self.tipo_amenaza,
            'id_impacto': self.id_impacto,
            'responsable_tic': self.responsable_tic,
            'acciones_tomadas': self.acciones_tomadas
        }
