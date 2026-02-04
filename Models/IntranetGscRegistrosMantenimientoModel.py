from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetGscRegistrosMantenimiento(BASE):
    """
    Modelo para registros del módulo MANTENIMIENTO (MNT)
    Almacena información específica de mantenimientos preventivos y correctivos
    """

    __tablename__ = "intranet_gsc_registros_mantenimiento"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id', ondelete='CASCADE'), nullable=False, unique=True)
    area = Column(String(200))
    tipo_mantenimiento = Column(String(200))
    fecha_inicio = Column(DateTime, nullable=False)
    fecha_fin = Column(DateTime, nullable=False)
    requiere_parada = Column(Boolean, default=False, nullable=False)
    id_riesgo = Column(Integer, ForeignKey('intranet_gsc_riesgos.id'), nullable=False)
    sistemas_componentes = Column(Text)
    responsable_ejecucion = Column(String(200))
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.area = data.get('area')
        self.tipo_mantenimiento = data.get('tipo_mantenimiento')
        self.fecha_inicio = data.get('fecha_inicio')
        self.fecha_fin = data.get('fecha_fin')
        self.requiere_parada = data.get('requiere_parada', False)
        self.id_riesgo = data.get('id_riesgo')
        self.sistemas_componentes = data.get('sistemas_componentes')
        self.responsable_ejecucion = data.get('responsable_ejecucion')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'area': self.area,
            'tipo_mantenimiento': self.tipo_mantenimiento,
            'fecha_inicio': self.fecha_inicio.isoformat() if self.fecha_inicio else None,
            'fecha_fin': self.fecha_fin.isoformat() if self.fecha_fin else None,
            'requiere_parada': self.requiere_parada,
            'id_riesgo': self.id_riesgo,
            'sistemas_componentes': self.sistemas_componentes,
            'responsable_ejecucion': self.responsable_ejecucion
        }
