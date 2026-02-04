from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from sqlalchemy.orm import relationship
from datetime import datetime

class IntranetGscRegistros(BASE):
    """
    Modelo para la tabla principal de registros GSC
    Almacena información común a todos los módulos (SEG, DISP, MNT, DR)
    """

    __tablename__ = "intranet_gsc_registros"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_modulo = Column(Integer, ForeignKey('intranet_gsc_modulos.id'), nullable=False)
    resumen = Column(String(500), nullable=False)
    descripcion = Column(Text)
    id_estado = Column(Integer, ForeignKey('intranet_gsc_estados.id'), nullable=False)
    
    # Fechas según estado (solo se llena la correspondiente)
    fecha_abierto = Column(DateTime)
    fecha_en_analisis = Column(DateTime)
    fecha_mitigado = Column(DateTime)
    fecha_cerrado = Column(DateTime)
    
    # Opciones
    notificar_gerencia = Column(Boolean, default=False, nullable=False)
    
    # Auditoría
    fecha_creacion = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion = Column(String(100))
    usuario_actualizacion = Column(String(100))
    activo = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        self.id_modulo = data.get('id_modulo')
        self.resumen = data.get('resumen')
        self.descripcion = data.get('descripcion')
        self.id_estado = data.get('id_estado')
        self.fecha_abierto = data.get('fecha_abierto')
        self.fecha_en_analisis = data.get('fecha_en_analisis')
        self.fecha_mitigado = data.get('fecha_mitigado')
        self.fecha_cerrado = data.get('fecha_cerrado')
        self.notificar_gerencia = data.get('notificar_gerencia', False)
        self.usuario_creacion = data.get('usuario_creacion')
        self.usuario_actualizacion = data.get('usuario_actualizacion')
        self.activo = data.get('activo', True)

    def to_dict(self):
        """Convierte el modelo a diccionario para serialización JSON"""
        return {
            'id': self.id,
            'id_modulo': self.id_modulo,
            'resumen': self.resumen,
            'descripcion': self.descripcion,
            'id_estado': self.id_estado,
            'fecha_abierto': self.fecha_abierto.isoformat() if self.fecha_abierto else None,
            'fecha_en_analisis': self.fecha_en_analisis.isoformat() if self.fecha_en_analisis else None,
            'fecha_mitigado': self.fecha_mitigado.isoformat() if self.fecha_mitigado else None,
            'fecha_cerrado': self.fecha_cerrado.isoformat() if self.fecha_cerrado else None,
            'notificar_gerencia': self.notificar_gerencia,
            'fecha_creacion': self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion': self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion': self.usuario_creacion,
            'usuario_actualizacion': self.usuario_actualizacion,
            'activo': self.activo
        }
