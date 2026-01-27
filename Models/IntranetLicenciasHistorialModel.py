from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Text, ForeignKey
from sqlalchemy.orm import relationship
from datetime import datetime

class IntranetLicenciasHistorialModel(BASE):

    __tablename__= "intranet_licencias_historial"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    licencia_id = Column(Integer, ForeignKey('intranet_licencias.id', ondelete='CASCADE'), nullable=False)
    fecha = Column(DateTime, default=datetime.now, nullable=False)
    usuario = Column(String(200), nullable=False)
    accion = Column(String(50), nullable=False)  # 'Creación', 'Edición', 'Baja', 'Reactivación'
    cambios = Column(Text)  # JSON string con los cambios
    created_at = Column(DateTime, default=datetime.now)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'licenciaId': self.licencia_id,
            'fecha': self.fecha.isoformat() if self.fecha else None,
            'usuario': self.usuario,
            'accion': self.accion,
            'cambios': self.cambios
        }
