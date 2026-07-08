from Config.db import BASE
from sqlalchemy import Column, String, Integer, Boolean


class IntranetCctvMetodosBackup(BASE):

    __tablename__ = "intranet_cctv_metodos_backup"

    id     = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(100), nullable=False)
    valor  = Column(String(50),  nullable=False)
    orden  = Column(Integer, nullable=True, default=0)
    activo = Column(Boolean, default=True, nullable=False)

    def to_dict(self):
        return {'id': self.id, 'nombre': self.nombre, 'valor': self.valor, 'orden': self.orden}
