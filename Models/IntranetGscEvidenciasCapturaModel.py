from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Text, ForeignKey
from datetime import datetime

class IntranetGscEvidenciasCaptura(BASE):
    """
    Modelo para evidencias de tipo CAPTURA
    Almacena información específica de capturas de pantalla o archivos
    """

    __tablename__ = "intranet_gsc_evidencias_captura"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_evidencia = Column(Integer, ForeignKey('intranet_gsc_evidencias.id', ondelete='CASCADE'), nullable=False, unique=True)
    nombre_archivo = Column(String(255), nullable=False)
    ruta_archivo = Column(String(500))
    archivo_base64 = Column(Text)
    tipo_mime = Column(String(100))
    tamano_bytes = Column(Integer)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_evidencia = data.get('id_evidencia')
        self.nombre_archivo = data.get('nombre_archivo')
        self.ruta_archivo = data.get('ruta_archivo')
        self.archivo_base64 = data.get('archivo_base64')
        self.tipo_mime = data.get('tipo_mime')
        self.tamano_bytes = data.get('tamano_bytes')

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_evidencia': self.id_evidencia,
            'nombre_archivo': self.nombre_archivo,
            'ruta_archivo': self.ruta_archivo,
            'archivo_base64': self.archivo_base64,
            'tipo_mime': self.tipo_mime,
            'tamano_bytes': self.tamano_bytes
        }
