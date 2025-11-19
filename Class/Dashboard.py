import requests
from Utils.tools import Tools, CustomException
from Utils.querys import Querys
from datetime import datetime, timedelta

from Utils.constants import (
    MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, MICROSOFT_TENANT_ID,
    MICROSOFT_API_SCOPE, MICROSOFT_URL, MICROSOFT_URL_GRAPH, PARENT_FOLDER,
    TARGET_FOLDER, EMAIL_USER
)

class Dashboard:

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.querys = Querys(self.db)
        self.token = None

    # Función para obtener métricas del dashboard
    def obtener_metricas_dashboard(self, data=None):
        """
        Obtiene las métricas principales del dashboard:
        - Total de tickets (activo=1, ticket=1)
        - Gestión (tipo_ticket=1)
        - Estratégicos (tipo_ticket=2)
        - Prioridad alta (prioridad=3)
        - Estados: Abiertos (estado=1), En proceso (estado=2), Completados (estado=3)
        """
        try:
            # Obtener filtros de fecha si se proporcionan
            filtros = data or {}
            fecha_inicio = filtros.get('fecha_inicio')
            fecha_fin = filtros.get('fecha_fin')
            
            # Obtener métricas usando query directo
            metricas = self.querys.obtener_metricas_dashboard(fecha_inicio, fecha_fin)
            
            return self.tools.output(200, "Métricas del dashboard obtenidas exitosamente.", metricas)
                
        except Exception as e:
            print(f"Error obteniendo métricas del dashboard: {e}")
            return self.tools.output(500, "Error obteniendo métricas del dashboard.", {})
    
    # Función para obtener indicadores de gestión mensual
    def obtener_indicadores_gestion(self, data=None):
        """
        Obtiene los indicadores de gestión por mes:
        - Tickets completados
        - Tickets cerrados oportunamente (fecha_cierre <= fecha_vencimiento)
        - Tickets cerrados no oportunamente (fecha_cierre > fecha_vencimiento)
        - Porcentaje de cumplimiento
        """
        try:
            filtros = data or {}
            anio = filtros.get('anio', datetime.now().year)
            
            # Obtener indicadores usando query
            indicadores = self.querys.obtener_indicadores_gestion(anio)
            
            return self.tools.output(200, "Indicadores de gestión obtenidos exitosamente.", indicadores)
                
        except Exception as e:
            print(f"Error obteniendo indicadores de gestión: {e}")
            return self.tools.output(500, "Error obteniendo indicadores de gestión.", {})
    
    # Función para obtener observación de un mes
    def obtener_observacion_mes(self, data=None):
        """
        Obtiene la observación de un mes específico
        """
        try:
            filtros = data or {}
            anio = filtros.get('anio')
            mes = filtros.get('mes')
            
            if not anio or not mes:
                return self.tools.output(400, "Año y mes son requeridos.", {})
            
            observacion = self.querys.obtener_observacion_mes(anio, mes)
            
            return self.tools.output(200, "Observación obtenida exitosamente.", observacion or {})
                
        except Exception as e:
            print(f"Error obteniendo observación del mes: {e}")
            return self.tools.output(500, "Error obteniendo observación del mes.", {})
    
    # Función para guardar observación de un mes
    def guardar_observacion_mes(self, data=None):
        """
        Guarda o actualiza la observación de un mes específico
        """
        try:
            filtros = data or {}
            anio = filtros.get('anio')
            mes = filtros.get('mes')
            observaciones = filtros.get('observaciones', '')
            
            if not anio or not mes:
                return self.tools.output(400, "Año y mes son requeridos.", {})
            
            observacion = self.querys.guardar_observacion_mes(anio, mes, observaciones)
            
            return self.tools.output(200, "Observación guardada exitosamente.", observacion)
                
        except Exception as e:
            print(f"Error guardando observación del mes: {e}")
            return self.tools.output(500, "Error guardando observación del mes.", {})
