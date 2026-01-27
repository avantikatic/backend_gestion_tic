from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Date, Text, ForeignKey
from datetime import datetime

class IntranetRevisionesModel(BASE):

    __tablename__= "intranet_revisiones"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    fecha = Column(Date, nullable=False)
    tipo_revision_id = Column(Integer, ForeignKey('intranet_tipo_revision.id'), nullable=False)
    observaciones = Column(Text)
    usuario = Column(String(200), nullable=False)
    estado = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'fecha': str(self.fecha) if self.fecha else None,
            'tipo_revision_id': self.tipo_revision_id,
            'observaciones': self.observaciones,
            'usuario': self.usuario,
            'created_at': str(self.created_at) if self.created_at else None
        }
