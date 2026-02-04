from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetGscRegistrosDisponibilidad(BASE):
    """
    Modelo para registros del módulo DISPONIBILIDAD (DISP)
    Almacena información específica de eventos de disponibilidad de servicios
    """

    __tablename__ = "intranet_gsc_registros_disponibilidad"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id', ondelete='CASCADE'), nullable=False, unique=True)
    servicio_afectado = Column(String(300), nullable=False)
    tipo_evento = Column(String(200))
    tiempo_indisponible_min = Column(Integer, default=0, nullable=False)
    sla_afectado = Column(Boolean, default=False, nullable=False)
    acciones = Column(Text)
    causa_raiz = Column(Text)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.servicio_afectado = data.get('servicio_afectado')
        self.tipo_evento = data.get('tipo_evento')
        self.tiempo_indisponible_min = data.get('tiempo_indisponible_min', 0)
        self.sla_afectado = data.get('sla_afectado', False)
        self.acciones = data.get('acciones')
        self.causa_raiz = data.get('causa_raiz')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'servicio_afectado': self.servicio_afectado,
            'tipo_evento': self.tipo_evento,
            'tiempo_indisponible_min': self.tiempo_indisponible_min,
            'sla_afectado': self.sla_afectado,
            'acciones': self.acciones,
            'causa_raiz': self.causa_raiz
        }
