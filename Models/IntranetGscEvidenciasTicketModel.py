from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, ForeignKey
from datetime import datetime

class IntranetGscEvidenciasTicket(BASE):
    """
    Modelo para evidencias de tipo TICKET
    Almacena información específica de tickets
    """

    __tablename__ = "intranet_gsc_evidencias_ticket"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_evidencia = Column(Integer, ForeignKey('intranet_gsc_evidencias.id', ondelete='CASCADE'), nullable=False, unique=True)
    numero_ticket = Column(String(100), nullable=False)
    plataforma = Column(String(100))
    url_ticket = Column(String(500))
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_evidencia = data.get('id_evidencia')
        self.numero_ticket = data.get('numero_ticket')
        self.plataforma = data.get('plataforma')
        self.url_ticket = data.get('url_ticket')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_evidencia': self.id_evidencia,
            'numero_ticket': self.numero_ticket,
            'plataforma': self.plataforma,
            'url_ticket': self.url_ticket
        }
