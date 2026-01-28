from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime
from datetime import datetime

class IntranetProveedoresModel(BASE):
    """
    Modelo para la tabla intranet_proveedores.
    NOTA: A partir de enero 2026, los proveedores se obtienen directamente de la tabla 'terceros'
    mediante la consulta en querys.obtener_proveedores(). Esta tabla se mantiene por compatibilidad
    con código legado.
    """

    __tablename__= "intranet_proveedores"
    
    id = Column(Integer, primary_key=True)
    nombre = Column(String(200))
    estado = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'nombre': self.nombre
        }
