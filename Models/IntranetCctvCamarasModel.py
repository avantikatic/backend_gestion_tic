from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime


class IntranetCctvCamaras(BASE):

    __tablename__ = "intranet_cctv_camaras"

    id                             = Column(Integer, primary_key=True, autoincrement=True)
    id_sede                        = Column(Integer, ForeignKey('intranet_cctv_sedes.id'), nullable=False)
    codigo_equipo_grabacion        = Column(String(100), nullable=False)
    ubicacion_fisica               = Column(String(300), nullable=False)
    id_estado_camara               = Column(Integer, ForeignKey('intranet_cctv_estados_camara.id'), nullable=True)
    dias_almacenamiento            = Column(Integer, default=0, nullable=True)
    id_metodo_backup               = Column(Integer, ForeignKey('intranet_cctv_metodos_backup.id'), nullable=True)
    fecha_instalacion_actualizacion = Column(String(20), nullable=True)
    observaciones                  = Column(Text, nullable=True)

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.id_sede                        = int(data.get('id_sede') or 0)
        self.codigo_equipo_grabacion        = data.get('codigo_equipo_grabacion')
        self.ubicacion_fisica               = data.get('ubicacion_fisica')
        self.id_estado_camara               = data.get('id_estado_camara') or None
        self.dias_almacenamiento            = int(data.get('dias_almacenamiento') or 0)
        self.id_metodo_backup               = data.get('id_metodo_backup') or None
        self.fecha_instalacion_actualizacion = data.get('fecha_instalacion_actualizacion')
        self.observaciones                  = data.get('observaciones')
        self.usuario_creacion               = data.get('usuario_creacion')
        self.usuario_actualizacion          = data.get('usuario_actualizacion')
        self.activo                         = data.get('activo', True)

    def to_dict(self):
        return {
            'id_camara':                     self.id,
            'id_sede':                       self.id_sede,
            'codigo_equipo_grabacion':       self.codigo_equipo_grabacion,
            'ubicacion_fisica':              self.ubicacion_fisica,
            'id_estado_camara':              self.id_estado_camara,
            'dias_almacenamiento':           self.dias_almacenamiento,
            'id_metodo_backup':              self.id_metodo_backup,
            'fecha_instalacion_actualizacion': self.fecha_instalacion_actualizacion,
            'observaciones':                 self.observaciones,
            'activo':                        self.activo,
            'fecha_creacion':                self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':           self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':              self.usuario_creacion,
            'usuario_actualizacion':         self.usuario_actualizacion,
        }
