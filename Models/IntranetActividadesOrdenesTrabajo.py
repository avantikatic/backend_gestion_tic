from sqlalchemy import Column, BigInteger, Text, TIMESTAMP, text, SmallInteger
from Config.db import BASE


class IntranetActividadesOrdenesTrabajo(BASE):
    __tablename__ = 'intranet_actividades_ordenes_trabajo'

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    orden_trabajo_id = Column(BigInteger, nullable=True)
    descripcion = Column(Text, nullable=True)
    tecnico = Column(Text, nullable=True)
    estado = Column(SmallInteger, nullable=True)
    created_at = Column(TIMESTAMP, default=text('GETDATE()'))
