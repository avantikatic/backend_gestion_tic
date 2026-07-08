from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime


class IntranetCctvRegistrosCambios(BASE):

    __tablename__ = "intranet_cctv_registros_cambios"

    id                = Column(Integer, primary_key=True, autoincrement=True)
    fecha_cambio      = Column(String(20), nullable=False)
    id_sede           = Column(Integer, ForeignKey('intranet_cctv_sedes.id'), nullable=True)
    id_camara         = Column(Integer, ForeignKey('intranet_cctv_camaras.id'), nullable=True)
    descripcion       = Column(Text, nullable=False)
    observaciones     = Column(Text, nullable=True)
    cargo_responsable = Column(String(200), nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.fecha_cambio      = data.get('fecha_cambio')
        self.id_sede           = data.get('id_sede') or None
        self.id_camara         = data.get('id_camara') or None
        self.descripcion       = data.get('descripcion')
        self.observaciones     = data.get('observaciones')
        self.cargo_responsable = data.get('cargo_responsable')
        self.usuario_creacion  = data.get('usuario_creacion')
        self.usuario_actualizacion = data.get('usuario_actualizacion')
        self.activo            = data.get('activo', True)

    def to_dict(self):
        return {
            'id_registro_cambio': self.id,
            'fecha_cambio':       self.fecha_cambio,
            'id_sede':            self.id_sede,
            'id_camara':          self.id_camara,
            'descripcion':        self.descripcion,
            'observaciones':      self.observaciones,
            'cargo_responsable':  self.cargo_responsable,
            'activo':             self.activo,
            'fecha_creacion':     self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion': self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':   self.usuario_creacion,
            'usuario_actualizacion': self.usuario_actualizacion,
        }
