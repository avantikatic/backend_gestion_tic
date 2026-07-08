from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text
from datetime import datetime


class IntranetCctvRevisiones(BASE):

    __tablename__ = "intranet_cctv_revisiones"

    id                      = Column(Integer, primary_key=True, autoincrement=True)
    fecha_revision          = Column(String(20), nullable=False)
    cargo_revisor           = Column(String(200), nullable=True)
    camaras_activas         = Column(Integer, default=0, nullable=True)
    camaras_mantenimiento   = Column(Integer, default=0, nullable=True)
    camaras_inactivas       = Column(Integer, default=0, nullable=True)
    almacenamiento_verificado = Column(Boolean, default=False, nullable=False)
    backups_verificados     = Column(Boolean, default=False, nullable=False)
    accesos_verificados     = Column(Boolean, default=False, nullable=False)
    hallazgos               = Column(Text, nullable=True)
    proxima_revision        = Column(String(20), nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.fecha_revision          = data.get('fecha_revision')
        self.cargo_revisor           = data.get('cargo_revisor')
        self.camaras_activas         = int(data.get('camaras_activas') or 0)
        self.camaras_mantenimiento   = int(data.get('camaras_mantenimiento') or 0)
        self.camaras_inactivas       = int(data.get('camaras_inactivas') or 0)
        self.almacenamiento_verificado = bool(data.get('almacenamiento_verificado', False))
        self.backups_verificados     = bool(data.get('backups_verificados', False))
        self.accesos_verificados     = bool(data.get('accesos_verificados', False))
        self.hallazgos               = data.get('hallazgos')
        self.proxima_revision        = data.get('proxima_revision')
        self.usuario_creacion        = data.get('usuario_creacion')
        self.usuario_actualizacion   = data.get('usuario_actualizacion')
        self.activo                  = data.get('activo', True)

    def to_dict(self):
        return {
            'id_revision':              self.id,
            'fecha_revision':           self.fecha_revision,
            'cargo_revisor':            self.cargo_revisor,
            'camaras_activas':          self.camaras_activas,
            'camaras_mantenimiento':    self.camaras_mantenimiento,
            'camaras_inactivas':        self.camaras_inactivas,
            'almacenamiento_verificado': bool(self.almacenamiento_verificado),
            'backups_verificados':      bool(self.backups_verificados),
            'accesos_verificados':      bool(self.accesos_verificados),
            'hallazgos':                self.hallazgos,
            'proxima_revision':         self.proxima_revision,
            'activo':                   self.activo,
            'fecha_creacion':           self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':      self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':         self.usuario_creacion,
            'usuario_actualizacion':    self.usuario_actualizacion,
        }
