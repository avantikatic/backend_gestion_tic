from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime


class IntranetCctvIncidentes(BASE):

    __tablename__ = "intranet_cctv_incidentes"

    id                   = Column(Integer, primary_key=True, autoincrement=True)
    fecha_incidente      = Column(String(20), nullable=False)
    id_sede              = Column(Integer, ForeignKey('intranet_cctv_sedes.id'), nullable=True)
    id_camara            = Column(Integer, ForeignKey('intranet_cctv_camaras.id'), nullable=True)
    titulo               = Column(String(500), nullable=False)
    id_severidad         = Column(Integer, ForeignKey('intranet_cctv_severidades.id'), nullable=True)
    descripcion          = Column(Text, nullable=False)
    accion_correctiva    = Column(Text, nullable=True)
    id_estado_incidente  = Column(Integer, ForeignKey('intranet_cctv_estados_incidente.id'), nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.fecha_incidente      = data.get('fecha_incidente')
        self.id_sede              = data.get('id_sede') or None
        self.id_camara            = data.get('id_camara') or None
        self.titulo               = data.get('titulo')
        self.id_severidad         = data.get('id_severidad') or None
        self.descripcion          = data.get('descripcion')
        self.accion_correctiva    = data.get('accion_correctiva')
        self.id_estado_incidente  = data.get('id_estado_incidente') or None
        self.usuario_creacion     = data.get('usuario_creacion')
        self.usuario_actualizacion = data.get('usuario_actualizacion')
        self.activo               = data.get('activo', True)

    def to_dict(self):
        return {
            'id_incidente':       self.id,
            'fecha_incidente':    self.fecha_incidente,
            'id_sede':            self.id_sede,
            'id_camara':          self.id_camara,
            'titulo':             self.titulo,
            'id_severidad':       self.id_severidad,
            'descripcion':        self.descripcion,
            'accion_correctiva':  self.accion_correctiva,
            'id_estado_incidente': self.id_estado_incidente,
            'activo':             self.activo,
            'fecha_creacion':     self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion': self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':   self.usuario_creacion,
            'usuario_actualizacion': self.usuario_actualizacion,
        }
