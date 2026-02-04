from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Text, ForeignKey
from datetime import datetime

class IntranetGscRegistrosDisasterRecovery(BASE):
    """
    Modelo para registros del módulo DISASTER RECOVERY (DR)
    Almacena información específica de pruebas y eventos de recuperación ante desastres
    """

    __tablename__ = "intranet_gsc_registros_disaster_recovery"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id', ondelete='CASCADE'), nullable=False, unique=True)
    escenario = Column(String(300))
    fecha_inicio = Column(DateTime, nullable=False)
    fecha_fin = Column(DateTime, nullable=False)
    objetivo = Column(Text)
    resultado = Column(Text)
    hallazgos = Column(Text)
    lecciones_aprendidas = Column(Text)
    rto_objetivo = Column(Integer)
    rto_real = Column(Integer)
    rpo_objetivo = Column(Integer)
    rpo_real = Column(Integer)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.escenario = data.get('escenario')
        self.fecha_inicio = data.get('fecha_inicio')
        self.fecha_fin = data.get('fecha_fin')
        self.objetivo = data.get('objetivo')
        self.resultado = data.get('resultado')
        self.hallazgos = data.get('hallazgos')
        self.lecciones_aprendidas = data.get('lecciones_aprendidas')
        self.rto_objetivo = data.get('rto_objetivo')
        self.rto_real = data.get('rto_real')
        self.rpo_objetivo = data.get('rpo_objetivo')
        self.rpo_real = data.get('rpo_real')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'escenario': self.escenario,
            'fecha_inicio': self.fecha_inicio.isoformat() if self.fecha_inicio else None,
            'fecha_fin': self.fecha_fin.isoformat() if self.fecha_fin else None,
            'objetivo': self.objetivo,
            'resultado': self.resultado,
            'hallazgos': self.hallazgos,
            'lecciones_aprendidas': self.lecciones_aprendidas,
            'rto_objetivo': self.rto_objetivo,
            'rto_real': self.rto_real,
            'rpo_objetivo': self.rpo_objetivo,
            'rpo_real': self.rpo_real
        }
