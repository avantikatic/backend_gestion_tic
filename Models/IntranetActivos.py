from sqlalchemy import Column, BigInteger, Integer, String, Text, Numeric, Date, DateTime, SmallInteger
from Config.db import BASE


class IntranetActivos(BASE):
    __tablename__ = 'intranet_activos'

    id                    = Column(BigInteger,   primary_key=True, autoincrement=True)
    codigo                = Column(String,        nullable=True)
    descripcion           = Column(Text,          nullable=True)
    modelo                = Column(String,        nullable=True)
    serie                 = Column(String,        nullable=True)
    marca                 = Column(String,        nullable=True)
    estado                = Column(BigInteger,    nullable=True)
    vida_util             = Column(Integer,       nullable=True)
    proveedor             = Column(Numeric,       nullable=True)
    tercero               = Column(Numeric,       nullable=True)
    docto_compra          = Column(String,        nullable=True)
    fecha_compra          = Column(Date,          nullable=True)
    caracteristicas       = Column(Text,          nullable=True)
    sede                  = Column(BigInteger,    nullable=True)
    centro                = Column(Integer,       nullable=True)
    grupo                 = Column(String,        nullable=True)
    macroproceso_encargado = Column(BigInteger,   nullable=True)
    macroproceso          = Column(BigInteger,    nullable=True)
    costo_compra          = Column(Numeric,       nullable=True)
    retirado              = Column(SmallInteger,  nullable=True)
    motivo_retiro         = Column(Text,          nullable=True)
    fecha_retiro          = Column(DateTime,      nullable=True)
    created_at            = Column(DateTime,      nullable=True)
