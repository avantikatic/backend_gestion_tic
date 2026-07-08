from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime

class IntranetContingenciaRecuperacion(BASE):

    __tablename__ = "intranet_contingencia_recuperacion"

    id                          = Column(Integer, primary_key=True, autoincrement=True)

    id_evento                   = Column(Integer, ForeignKey('intranet_contingencia_eventos.id'), nullable=False, unique=True)
    id_resultado_recuperacion   = Column(Integer, ForeignKey('intranet_contingencia_resultados_recuperacion.id'), nullable=True)

    tiempo_real                 = Column(String(50), nullable=True)
    datos_recuperados           = Column(String(500), nullable=True)
    observaciones               = Column(Text, nullable=True)

    servicio_alterno            = Column(Boolean, default=False, nullable=False)
    integridad_documental       = Column(Boolean, default=False, nullable=False)
    informe_final               = Column(Boolean, default=False, nullable=False)
    lecciones_aprendidas        = Column(Boolean, default=False, nullable=False)

    activo                      = Column(Boolean, default=True, nullable=False)
    fecha_creacion              = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion         = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion            = Column(String(100), nullable=True)
    usuario_actualizacion       = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.id_evento                 = data.get('id_evento')
        self.id_resultado_recuperacion = data.get('id_resultado_recuperacion')
        self.tiempo_real               = data.get('tiempo_real')
        self.datos_recuperados         = data.get('datos_recuperados')
        self.observaciones             = data.get('observaciones')
        self.servicio_alterno          = data.get('servicio_alterno', False)
        self.integridad_documental     = data.get('integridad_documental', False)
        self.informe_final             = data.get('informe_final', False)
        self.lecciones_aprendidas      = data.get('lecciones_aprendidas', False)
        self.usuario_creacion          = data.get('usuario_creacion')
        self.usuario_actualizacion     = data.get('usuario_actualizacion')
        self.activo                    = data.get('activo', True)

    def to_dict(self):
        return {
            'id':                          self.id,
            'id_evento':                   self.id_evento,
            'id_resultado_recuperacion':   self.id_resultado_recuperacion,
            'tiempo_real':                 self.tiempo_real,
            'datos_recuperados':           self.datos_recuperados,
            'observaciones':               self.observaciones,
            'servicio_alterno':            self.servicio_alterno,
            'integridad_documental':       self.integridad_documental,
            'informe_final':               self.informe_final,
            'lecciones_aprendidas':        self.lecciones_aprendidas,
            'activo':                      self.activo,
            'fecha_creacion':              self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':         self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':            self.usuario_creacion,
            'usuario_actualizacion':       self.usuario_actualizacion,
        }
