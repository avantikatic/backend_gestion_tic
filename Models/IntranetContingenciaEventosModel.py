from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetContingenciaEventos(BASE):

    __tablename__ = "intranet_contingencia_eventos"

    id                    = Column(Integer, primary_key=True, autoincrement=True)
    codigo                = Column(String(20), nullable=False, unique=True)

    id_tipo_evento        = Column(Integer, ForeignKey('intranet_contingencia_tipos_evento.id'), nullable=False)
    id_prioridad          = Column(Integer, ForeignKey('intranet_contingencia_prioridades.id'), nullable=False)
    id_estado_evento      = Column(Integer, ForeignKey('intranet_contingencia_estados_evento.id'), nullable=False)

    titulo                = Column(String(500), nullable=False)
    area                  = Column(String(200), nullable=True)
    responsable           = Column(String(200), nullable=True)
    fecha_inicio          = Column(String(100), nullable=True)
    rto_objetivo          = Column(String(50), nullable=True)
    rpo_objetivo          = Column(String(50), nullable=True)
    impacto               = Column(Text, nullable=True)
    causa                 = Column(Text, nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.codigo               = data.get('codigo')
        self.id_tipo_evento       = data.get('id_tipo_evento')
        self.id_prioridad         = data.get('id_prioridad')
        self.id_estado_evento     = data.get('id_estado_evento')
        self.titulo               = data.get('titulo')
        self.area                 = data.get('area')
        self.responsable          = data.get('responsable')
        self.fecha_inicio         = data.get('fecha_inicio')
        self.rto_objetivo         = data.get('rto_objetivo')
        self.rpo_objetivo         = data.get('rpo_objetivo')
        self.impacto              = data.get('impacto')
        self.causa                = data.get('causa')
        self.usuario_creacion     = data.get('usuario_creacion')
        self.usuario_actualizacion = data.get('usuario_actualizacion')
        self.activo               = data.get('activo', True)

    def to_dict(self):
        return {
            'id':                    self.id,
            'codigo':                self.codigo,
            'id_tipo_evento':        self.id_tipo_evento,
            'id_prioridad':          self.id_prioridad,
            'id_estado_evento':      self.id_estado_evento,
            'titulo':                self.titulo,
            'area':                  self.area,
            'responsable':           self.responsable,
            'fecha_inicio':          self.fecha_inicio,
            'rto_objetivo':          self.rto_objetivo,
            'rpo_objetivo':          self.rpo_objetivo,
            'impacto':               self.impacto,
            'causa':                 self.causa,
            'activo':                self.activo,
            'fecha_creacion':        self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':   self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':      self.usuario_creacion,
            'usuario_actualizacion': self.usuario_actualizacion,
        }
