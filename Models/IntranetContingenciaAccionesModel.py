from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetContingenciaAcciones(BASE):

    __tablename__ = "intranet_contingencia_acciones"

    id                    = Column(Integer, primary_key=True, autoincrement=True)
    codigo                = Column(String(20), nullable=False)

    id_evento             = Column(Integer, ForeignKey('intranet_contingencia_eventos.id'), nullable=False)
    id_estado_accion      = Column(Integer, ForeignKey('intranet_contingencia_estados_accion.id'), nullable=False)

    titulo                = Column(String(500), nullable=False)
    responsable           = Column(String(200), nullable=True)
    fecha_compromiso      = Column(String(100), nullable=True)
    evidencia             = Column(Text, nullable=True)
    control_asociado      = Column(String(300), nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.codigo               = data.get('codigo')
        self.id_evento            = data.get('id_evento')
        self.id_estado_accion     = data.get('id_estado_accion')
        self.titulo               = data.get('titulo')
        self.responsable          = data.get('responsable')
        self.fecha_compromiso     = data.get('fecha_compromiso')
        self.evidencia            = data.get('evidencia')
        self.control_asociado     = data.get('control_asociado')
        self.usuario_creacion     = data.get('usuario_creacion')
        self.usuario_actualizacion = data.get('usuario_actualizacion')
        self.activo               = data.get('activo', True)

    def to_dict(self):
        return {
            'id':                    self.id,
            'codigo':                self.codigo,
            'id_evento':             self.id_evento,
            'id_estado_accion':      self.id_estado_accion,
            'titulo':                self.titulo,
            'responsable':           self.responsable,
            'fecha_compromiso':      self.fecha_compromiso,
            'evidencia':             self.evidencia,
            'control_asociado':      self.control_asociado,
            'activo':                self.activo,
            'fecha_creacion':        self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':   self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':      self.usuario_creacion,
            'usuario_actualizacion': self.usuario_actualizacion,
        }
