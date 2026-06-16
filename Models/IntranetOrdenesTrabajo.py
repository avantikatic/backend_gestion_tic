from sqlalchemy import Column, BigInteger, Integer, Date, Text, TIMESTAMP, text, SmallInteger
from Config.db import BASE


class IntranetOrdenesTrabajo(BASE):
    __tablename__ = 'intranet_ordenes_trabajo'

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    activo_id = Column(BigInteger, nullable=True)
    tipo_mantenimiento = Column(Integer, nullable=True)
    fecha_programacion_desde = Column(Date, nullable=True)
    fecha_programacion_hasta = Column(Date, nullable=True)
    tecnico_asignado = Column(BigInteger, nullable=True)
    descripcion = Column(Text, nullable=True)
    estado_ot = Column(BigInteger, nullable=True)
    estado = Column(SmallInteger, nullable=True)
    created_at = Column(TIMESTAMP, default=text('GETDATE()'))
