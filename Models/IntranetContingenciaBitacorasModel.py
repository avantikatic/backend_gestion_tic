from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetContingenciaBitacoras(BASE):

    __tablename__ = "intranet_contingencia_bitacoras"

    id                 = Column(Integer, primary_key=True, autoincrement=True)
    codigo             = Column(String(20), nullable=False)

    id_evento          = Column(Integer, ForeignKey('intranet_contingencia_eventos.id'), nullable=False)
    id_tipo_bitacora   = Column(Integer, ForeignKey('intranet_contingencia_tipos_bitacora.id'), nullable=False)

    hora_registro      = Column(String(50), nullable=True)
    actor              = Column(String(200), nullable=True)
    detalle            = Column(Text, nullable=False)

    activo             = Column(Boolean, default=True, nullable=False)
    fecha_creacion     = Column(DateTime, default=datetime.now, nullable=False)
    usuario_creacion   = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.codigo           = data.get('codigo')
        self.id_evento        = data.get('id_evento')
        self.id_tipo_bitacora = data.get('id_tipo_bitacora')
        self.hora_registro    = data.get('hora_registro')
        self.actor            = data.get('actor')
        self.detalle          = data.get('detalle')
        self.usuario_creacion = data.get('usuario_creacion')
        self.activo           = data.get('activo', True)

    def to_dict(self):
        return {
            'id':               self.id,
            'codigo':           self.codigo,
            'id_evento':        self.id_evento,
            'id_tipo_bitacora': self.id_tipo_bitacora,
            'hora_registro':    self.hora_registro,
            'actor':            self.actor,
            'detalle':          self.detalle,
            'activo':           self.activo,
            'fecha_creacion':   self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'usuario_creacion': self.usuario_creacion,
        }
