from sqlalchemy import Column, BigInteger, Integer, String, Date, TIMESTAMP, text
from Config.db import BASE


class IntranetCausasInformeGestion(BASE):
    __tablename__ = 'intranet_causas_informe_gestion'

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    anio = Column(Integer, nullable=False)
    mes = Column(Integer, nullable=False)
    analisis = Column(String(None), nullable=True)  # VARCHAR(MAX)
    acciones = Column(String(None), nullable=True)  # VARCHAR(MAX)
    responsable = Column(String(None), nullable=True)  # VARCHAR(MAX)
    fecha_compromiso = Column(Date, nullable=True)
    seguimiento = Column(String(None), nullable=True)  # VARCHAR(MAX)
    tipo_ticket = Column(Integer, nullable=True)
    estado = Column(Integer, nullable=False, default=1)
    created_at = Column(TIMESTAMP, default=text('GETDATE()'))

    def to_dict(self):
        return {
            'id': self.id,
            'anio': self.anio,
            'mes': self.mes,
            'analisis': self.analisis,
            'acciones': self.acciones,
            'responsable': self.responsable,
            'fecha_compromiso': self.fecha_compromiso.isoformat() if self.fecha_compromiso else None,
            'seguimiento': self.seguimiento,
            'tipo_ticket': self.tipo_ticket,
            'estado': self.estado,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }
