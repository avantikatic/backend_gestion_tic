from Config.db import BASE
from sqlalchemy import Column, String, Integer, Boolean


class IntranetCctvEstadosIncidente(BASE):

    __tablename__ = "intranet_cctv_estados_incidente"

    id     = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(100), nullable=False)
    valor  = Column(String(50),  nullable=False)
    orden  = Column(Integer, nullable=True, default=0)
    activo = Column(Boolean, default=True, nullable=False)

    def to_dict(self):
        return {'id': self.id, 'nombre': self.nombre, 'valor': self.valor, 'orden': self.orden}
