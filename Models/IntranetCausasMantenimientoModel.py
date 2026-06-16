from sqlalchemy import Column, BigInteger, Integer, String, Date, TIMESTAMP, text
from Config.db import BASE


class IntranetCausasMantenimientoModel(BASE):
    __tablename__ = 'intranet_causas_mantenimiento'

    id               = Column(BigInteger, primary_key=True, autoincrement=True)
    anio             = Column(Integer,    nullable=False)
    mes              = Column(Integer,    nullable=False)
    analisis         = Column(String(None), nullable=True)
    acciones         = Column(String(None), nullable=True)
    responsable      = Column(String(None), nullable=True)
    fecha_compromiso = Column(Date,       nullable=True)
    seguimiento      = Column(String(None), nullable=True)
    estado           = Column(Integer,    nullable=False, default=1)
    created_at       = Column(TIMESTAMP,  default=text('GETDATE()'))

    def to_dict(self):
        return {
            'id':               self.id,
            'anio':             self.anio,
            'mes':              self.mes,
            'analisis':         self.analisis,
            'acciones':         self.acciones,
            'responsable':      self.responsable,
            'fecha_compromiso': self.fecha_compromiso.isoformat() if self.fecha_compromiso else None,
            'seguimiento':      self.seguimiento,
            'estado':           self.estado,
            'created_at':       self.created_at.isoformat() if self.created_at else None,
        }
