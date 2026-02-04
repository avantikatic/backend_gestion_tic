from Config.db import BASE
from sqlalchemy import Column, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetGscEvidencias(BASE):
    """
    Modelo para la tabla base de evidencias
    Almacena información común a todos los tipos de evidencias
    """

    __tablename__ = "intranet_gsc_evidencias"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_registro = Column(Integer, ForeignKey('intranet_gsc_registros.id', ondelete='CASCADE'), nullable=False)
    id_tipo_evidencia = Column(Integer, ForeignKey('intranet_gsc_tipos_evidencia.id'), nullable=False)
    observacion = Column(Text)
    fecha_evidencia = Column(DateTime, default=datetime.now, nullable=False)
    fecha_creacion = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    activo = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_registro = data.get('id_registro')
        self.id_tipo_evidencia = data.get('id_tipo_evidencia')
        self.observacion = data.get('observacion')
        self.fecha_evidencia = data.get('fecha_evidencia', datetime.now())
        self.activo = data.get('activo', True)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_registro': self.id_registro,
            'id_tipo_evidencia': self.id_tipo_evidencia,
            'observacion': self.observacion,
            'fecha_evidencia': self.fecha_evidencia.isoformat() if self.fecha_evidencia else None,
            'fecha_creacion': self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion': self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'activo': self.activo
        }
