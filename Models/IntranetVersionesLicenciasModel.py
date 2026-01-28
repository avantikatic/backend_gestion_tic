from sqlalchemy import Column, Integer, String, Date, Text, DateTime
from datetime import datetime
from Config.db import BASE

class IntranetVersionesLicenciasModel(BASE):
    __tablename__ = "intranet_versiones_licencias"
    __table_args__ = {'extend_existing': True}
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    fecha = Column(Date, nullable=False)
    version = Column(String(50), nullable=False)
    descripcion = Column(Text, nullable=True)
    estado = Column(Integer, default=1)
    fecha_creacion = Column(DateTime, default=datetime.utcnow)
