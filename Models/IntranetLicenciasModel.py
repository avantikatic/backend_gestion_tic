from Config.db import BASE
from sqlalchemy import Column, String, BigInteger, Text, Integer, DateTime, Date, Numeric, Boolean, ForeignKey
from sqlalchemy.orm import relationship
from datetime import datetime

class IntranetLicenciasModel(BASE):

    __tablename__= "intranet_licencias"
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    tipo_servicio_id = Column(Integer, ForeignKey('intranet_tipos_servicio.id'))
    proveedor_id = Column(Integer, ForeignKey('intranet_proveedores.id'))
    producto_id = Column(Integer, ForeignKey('intranet_productos_servicios.id'))
    cantidad = Column(Integer)
    frecuencia = Column(String(20))
    fecha_compra = Column(Date)
    fecha_vencimiento = Column(Date)
    valor = Column(Numeric(18, 2))
    metodo_pago_id = Column(Integer, ForeignKey('intranet_metodos_pago.id'))
    responsable_nombre = Column(String(200))
    responsable_cargo = Column(String(200))
    observaciones = Column(Text)
    baja = Column(Boolean, default=False)
    fecha_baja = Column(Date)
    motivo_baja = Column(Text)
    estado = Column(Integer, default=1)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __init__(self, data: dict):
        # No se inicializa el ID, se genera automáticamente
        self.tipo_servicio_id = data.get('tipoServicioId')
        self.proveedor_id = data.get('proveedorId')
        self.producto_id = data.get('productoId')
        self.cantidad = data.get('cantidad')
        self.frecuencia = data.get('frecuencia')
        self.fecha_compra = data.get('fechaCompra')
        self.fecha_vencimiento = data.get('fechaVencimiento')
        self.valor = data.get('valor')
        self.metodo_pago_id = data.get('metodoPagoId')
        self.responsable_nombre = data.get('responsable', {}).get('nombre')
        self.responsable_cargo = data.get('responsable', {}).get('cargo')
        self.observaciones = data.get('observaciones')
        self.baja = data.get('baja', False)
        self.fecha_baja = data.get('fechaBaja')
        self.motivo_baja = data.get('motivoBaja')

    def to_dict(self, include_relations=False):
        """Convierte el modelo a diccionario para serialización JSON"""
        result = {
            'id': self.id,
            'tipoServicioId': self.tipo_servicio_id,
            'proveedorId': self.proveedor_id,
            'productoId': self.producto_id,
            'cantidad': self.cantidad,
            'frecuencia': self.frecuencia,
            'fechaCompra': self.fecha_compra.isoformat() if self.fecha_compra else None,
            'fechaVencimiento': self.fecha_vencimiento.isoformat() if self.fecha_vencimiento else None,
            'valor': float(self.valor) if self.valor else 0,
            'metodoPagoId': self.metodo_pago_id,
            'responsable': {
                'nombre': self.responsable_nombre,
                'cargo': self.responsable_cargo
            },
            'observaciones': self.observaciones,
            'baja': self.baja,
            'fechaBaja': self.fecha_baja.isoformat() if self.fecha_baja else None,
            'motivoBaja': self.motivo_baja,
            'historial': []  # Se llenará posteriormente con otra tabla
        }
        return result
