from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime
from datetime import datetime

class IntranetProductosServiciosModel(BASE):

    __tablename__= "intranet_productos_servicios"
    
    id = Column(Integer, primary_key=True)
    nombre = Column(String(300))
    estado = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'nombre': self.nombre
        }
