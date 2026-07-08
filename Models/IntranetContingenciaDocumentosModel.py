from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, ForeignKey
from datetime import datetime

class IntranetContingenciaDocumentos(BASE):

    __tablename__ = "intranet_contingencia_documentos"

    id                    = Column(Integer, primary_key=True, autoincrement=True)
    codigo_interno        = Column(String(20), nullable=False)

    id_evento             = Column(Integer, ForeignKey('intranet_contingencia_eventos.id'), nullable=False)
    id_estado_documento   = Column(Integer, ForeignKey('intranet_contingencia_estados_documento.id'), nullable=False)

    nombre                = Column(String(300), nullable=False)
    codigo_documento      = Column(String(50), nullable=True)   # ej. "FR-TIC-17"
    responsable           = Column(String(200), nullable=True)
    fecha_documento       = Column(String(50), nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.codigo_interno       = data.get('codigo_interno')
        self.id_evento            = data.get('id_evento')
        self.id_estado_documento  = data.get('id_estado_documento')
        self.nombre               = data.get('nombre')
        self.codigo_documento     = data.get('codigo_documento')
        self.responsable          = data.get('responsable')
        self.fecha_documento      = data.get('fecha_documento')
        self.usuario_creacion     = data.get('usuario_creacion')
        self.usuario_actualizacion = data.get('usuario_actualizacion')
        self.activo               = data.get('activo', True)

    def to_dict(self):
        return {
            'id':                    self.id,
            'codigo_interno':        self.codigo_interno,
            'id_evento':             self.id_evento,
            'id_estado_documento':   self.id_estado_documento,
            'nombre':                self.nombre,
            'codigo_documento':      self.codigo_documento,
            'responsable':           self.responsable,
            'fecha_documento':       self.fecha_documento,
            'activo':                self.activo,
            'fecha_creacion':        self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':   self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':      self.usuario_creacion,
            'usuario_actualizacion': self.usuario_actualizacion,
        }
