from Config.db import BASE
from sqlalchemy import Column, String, Integer, DateTime, Boolean, Text, ForeignKey
from datetime import datetime
import json


class IntranetCctvCargos(BASE):

    __tablename__ = "intranet_cctv_cargos"

    id                              = Column(Integer, primary_key=True, autoincrement=True)
    nombre                          = Column(String(200), nullable=False)
    id_nivel_acceso                 = Column(Integer, ForeignKey('intranet_cctv_niveles_acceso.id'), nullable=True)
    justificacion_acceso            = Column(Text, nullable=True)
    puede_ver_camaras               = Column(Boolean, default=False, nullable=False)
    puede_editar_inventario         = Column(Boolean, default=False, nullable=False)
    puede_administrar_configuracion = Column(Boolean, default=False, nullable=False)
    ids_sedes                       = Column(Text, nullable=True)  # JSON array: "[1, 3, 5]"

    activo                = Column(Boolean, default=True, nullable=False)
    fecha_creacion        = Column(DateTime, default=datetime.now, nullable=False)
    fecha_actualizacion   = Column(DateTime, default=datetime.now, onupdate=datetime.now, nullable=False)
    usuario_creacion      = Column(String(100), nullable=True)
    usuario_actualizacion = Column(String(100), nullable=True)

    def __init__(self, data: dict):
        self.nombre                          = data.get('nombre')
        self.id_nivel_acceso                 = data.get('id_nivel_acceso') or None
        self.justificacion_acceso            = data.get('justificacion_acceso')
        self.puede_ver_camaras               = bool(data.get('puede_ver_camaras', False))
        self.puede_editar_inventario         = bool(data.get('puede_editar_inventario', False))
        self.puede_administrar_configuracion = bool(data.get('puede_administrar_configuracion', False))
        ids = data.get('ids_sedes') or []
        self.ids_sedes                       = json.dumps([int(i) for i in ids])
        self.usuario_creacion                = data.get('usuario_creacion')
        self.usuario_actualizacion           = data.get('usuario_actualizacion')
        self.activo                          = data.get('activo', True)

    def to_dict(self):
        try:
            ids_list = json.loads(self.ids_sedes) if self.ids_sedes else []
        except Exception:
            ids_list = []
        return {
            'id_cargo':                      self.id,
            'nombre':                        self.nombre,
            'id_nivel_acceso':               self.id_nivel_acceso,
            'justificacion_acceso':          self.justificacion_acceso,
            'puede_ver_camaras':             bool(self.puede_ver_camaras),
            'puede_editar_inventario':       bool(self.puede_editar_inventario),
            'puede_administrar_configuracion': bool(self.puede_administrar_configuracion),
            'ids_sedes':                     ids_list,
            'activo':                        self.activo,
            'fecha_creacion':                self.fecha_creacion.isoformat() if self.fecha_creacion else None,
            'fecha_actualizacion':           self.fecha_actualizacion.isoformat() if self.fecha_actualizacion else None,
            'usuario_creacion':              self.usuario_creacion,
            'usuario_actualizacion':         self.usuario_actualizacion,
        }
