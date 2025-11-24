from sqlalchemy import Column, BigInteger, Integer, String, DateTime, text
from Config.db import BASE


class IntranetAniosInformeGestion(BASE):
    __tablename__ = 'intranet_anios_informe_gestion'

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    anio = Column(Integer, nullable=False, unique=True)
    descripcion = Column(String(255), nullable=True)
    estado = Column(Integer, nullable=False, server_default=text('1'))
    created_at = Column(DateTime, server_default=text('GETDATE()'))

    def to_dict(self):
        return {
            'id': self.id,
            'anio': self.anio,
            'descripcion': self.descripcion,
            'estado': self.estado,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }
