from Config.db import BASE
from sqlalchemy import Column, String, Integer, Boolean


class IntranetCctvResponsablesTic(BASE):

    __tablename__ = "intranet_cctv_responsables_tic"

    id     = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(200), nullable=False)
    orden  = Column(Integer, nullable=True, default=0)
    activo = Column(Boolean, default=True, nullable=False)

    def to_dict(self):
        return {'id': self.id, 'nombre': self.nombre, 'orden': self.orden}
