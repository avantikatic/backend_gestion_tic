from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text
from datetime import datetime


class IntranetCctvSedes(BASE):

    __tablename__ = "intranet_cctv_sedes"

    id                          = Column(Integer, primary_key=True, autoincrement=True)
    nombre                      = Column(String(200), nullable=False)
    ubicacion_general           = Column(String(300), nullable=True)
    responsable_operativo       = Column(String(200), nullable=True)
    sistema_grabacion           = Column(String(200), nullable=True)
    dias_almacenamiento_estimado = Column(Integer, default=0, nullable=True)
    observaciones               = Column(Text, nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.nombre                      = data.get('nombre')
        self.ubicacion_general           = data.get('ubicacion_general')
        self.responsable_operativo       = data.get('responsable_operativo')
        self.sistema_grabacion           = data.get('sistema_grabacion')
        self.dias_almacenamiento_estimado = int(data.get('dias_almacenamiento_estimado') or 0)
        self.observaciones               = data.get('observaciones')
        self.usuario_creacion            = data.get('usuario_creacion')
        self.usuario_actualizacion       = data.get('usuario_actualizacion')
        self.activo                      = data.get('activo', True)

    def to_dict(self):
        return {
            'id_sede':                      self.id,
            'nombre':                       self.nombre,
            'ubicacion_general':            self.ubicacion_general,
            'responsable_operativo':        self.responsable_operativo,
            'sistema_grabacion':            self.sistema_grabacion,
            'dias_almacenamiento_estimado': self.dias_almacenamiento_estimado,
            'observaciones':                self.observaciones,
            'activo':                       self.activo,
            'fecha_creacion':               self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':          self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':             self.usuario_creacion,
            'usuario_actualizacion':        self.usuario_actualizacion,
        }
