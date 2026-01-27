from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime
from datetime import datetime

class IntranetTipoRevisionModel(BASE):

    __tablename__= "intranet_tipo_revision"
    
    id = Column(Integer, primary_key=True)
    nombre = Column(String(100))
    estado = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'nombre': self.nombre
        }
