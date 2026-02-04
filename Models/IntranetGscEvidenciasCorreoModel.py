from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Text, ForeignKey
from datetime import datetime

class IntranetGscEvidenciasCorreo(BASE):
    """
    Modelo para evidencias de tipo CORREO
    Almacena información específica de correos electrónicos
    """

    __tablename__ = "intranet_gsc_evidencias_correo"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_evidencia = Column(Integer, ForeignKey('intranet_gsc_evidencias.id', ondelete='CASCADE'), nullable=False, unique=True)
    asunto = Column(String(500), nullable=False)
    remitente = Column(String(255))
    destinatarios = Column(Text)
    fecha_envio = Column(DateTime)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_evidencia = data.get('id_evidencia')
        self.asunto = data.get('asunto')
        self.remitente = data.get('remitente')
        self.destinatarios = data.get('destinatarios')
        self.fecha_envio = data.get('fecha_envio')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_evidencia': self.id_evidencia,
            'asunto': self.asunto,
            'remitente': self.remitente,
            'destinatarios': self.destinatarios,
            'fecha_envio': self.fecha_envio.isoformat() if self.fecha_envio else None
        }
