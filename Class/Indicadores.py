import requests
from Utils.tools import Tools, CustomException
from Utils.querys import Querys
from datetime import datetime, timedelta

from Utils.constants import (
    MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, MICROSOFT_TENANT_ID,
    MICROSOFT_API_SCOPE, MICROSOFT_URL, MICROSOFT_URL_GRAPH, PARENT_FOLDER,
    TARGET_FOLDER, EMAIL_USER
)

class Indicadores:

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.querys = Querys(self.db)
        self.token = None

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
    
    # Función para obtener indicadores estratégicos mensuales
    def obtener_indicadores_estrategicos(self, data=None):
        """
        Obtiene los indicadores de tickets estratégicos (tipo_ticket = 2) agrupados por origen_estrategico:
        - Tickets completados por origen
        - Tickets cerrados oportunamente por origen
        - Tickets cerrados no oportunamente por origen
        - Porcentaje de cumplimiento por origen
        """
        try:
            filtros = data or {}
            anio = filtros.get('anio', datetime.now().year)
            
            # Obtener indicadores usando query
            indicadores = self.querys.obtener_indicadores_estrategicos(anio)
            
            return self.tools.output(200, "Indicadores estratégicos obtenidos exitosamente.", indicadores)
                
        except Exception as e:
            print(f"Error obteniendo indicadores estratégicos: {e}")
            return self.tools.output(500, "Error obteniendo indicadores estratégicos.", {})
    
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
    
    # Función para obtener análisis de causas de un año
    def obtener_analisis_causas(self, data=None):
        """
        Obtiene todos los análisis de causas y acciones de un año específico
        """
        try:
            filtros = data or {}
            anio = filtros.get('anio')
            tipo_ticket = filtros.get('tipo_ticket', 1)  # Default 1 (Gestión)
            
            if not anio:
                return self.tools.output(400, "Año es requerido.", {})
            
            analisis = self.querys.obtener_analisis_causas(anio, tipo_ticket)
            
            return self.tools.output(200, "Análisis de causas obtenidos exitosamente.", analisis)
                
        except Exception as e:
            print(f"Error obteniendo análisis de causas: {e}")
            return self.tools.output(500, "Error obteniendo análisis de causas.", {})
    
    # Función para guardar análisis de causas
    def guardar_analisis_causas(self, data=None):
        """
        Guarda o actualiza un análisis de causas y acciones.
        Valida que no exista otro registro con el mismo año y mes.
        """
        try:
            filtros = data or {}
            id_analisis = filtros.get('id')
            anio = filtros.get('anio')
            mes = filtros.get('mes')
            analisis = filtros.get('analisis', '')
            acciones = filtros.get('acciones', '')
            responsable = filtros.get('responsable', '')
            fecha_compromiso = filtros.get('fecha_compromiso')
            seguimiento = filtros.get('seguimiento', '')
            tipo_ticket = filtros.get('tipo_ticket', 1)  # Default 1 (Gestión)
            
            if not anio or not mes:
                return self.tools.output(400, "Año y mes son requeridos.", {})
            
            # Validar que mes esté entre 1 y 12
            if not isinstance(mes, int) or mes < 1 or mes > 12:
                return self.tools.output(400, "El mes debe ser un número entre 1 y 12.", {})
            
            # Si no hay ID (nuevo registro), validar que no exista año+mes+tipo
            if not id_analisis:
                existe = self.querys.verificar_analisis_existe(anio, mes, tipo_ticket)
                if existe:
                    return self.tools.output(400, f"Ya existe un análisis para el mes {mes} del año {anio}.", {})
            
            resultado = self.querys.guardar_analisis_causas(
                id_analisis, anio, mes, analisis, acciones, 
                responsable, fecha_compromiso, seguimiento, tipo_ticket
            )
            
            return self.tools.output(200, "Análisis de causas guardado exitosamente.", resultado)
                
        except Exception as e:
            print(f"Error guardando análisis de causas: {e}")
            return self.tools.output(500, "Error guardando análisis de causas.", {})

    # Función para obtener tickets del periodo
    def obtener_tickets_periodo(self, data=None):
        """
        Obtiene los tickets del periodo especificado (año y mes)
        """
        try:
            filtros = data or {}
            anio = filtros.get('anio')
            mes = filtros.get('mes')
            page = filtros.get('page', 1)
            limit = filtros.get('limit', 5)
            
            if not anio or not mes:
                return self.tools.output(400, "Año y mes son requeridos.", {})
            
            resultado = self.querys.obtener_tickets_periodo(anio, mes, page, limit)
            
            return self.tools.output(200, "Tickets del periodo obtenidos exitosamente.", resultado)
                
        except Exception as e:
            print(f"Error obteniendo tickets del periodo: {e}")
            return self.tools.output(500, "Error obteniendo tickets del periodo.", {})

    # Función para obtener años disponibles
    def obtener_anios_disponibles(self, data=None):
        """
        Obtiene todos los años disponibles en el sistema
        """
        try:
            anios = self.querys.obtener_anios_disponibles()
            return self.tools.output(200, "Años disponibles obtenidos exitosamente.", anios)
                
        except Exception as e:
            print(f"Error obteniendo años disponibles: {e}")
            return self.tools.output(500, "Error obteniendo años disponibles.", [])

    # Función para crear un nuevo año
    def crear_anio(self, data=None):
        """
        Crea un nuevo año en el sistema
        """
        try:
            if not data:
                return self.tools.output(400, "Datos requeridos.", {})
            
            anio = data.get('anio')
            descripcion = data.get('descripcion', '')
            
            if not anio:
                return self.tools.output(400, "El año es requerido.", {})
            
            # Validar que sea un número de 4 dígitos
            try:
                anio_int = int(anio)
                if anio_int < 1900 or anio_int > 2100:
                    return self.tools.output(400, "El año debe estar entre 1900 y 2100.", {})
            except ValueError:
                return self.tools.output(400, "El año debe ser un número válido.", {})
            
            nuevo_anio = self.querys.crear_anio(anio_int, descripcion)
            
            return self.tools.output(200, "Año creado exitosamente.", nuevo_anio)
                
        except CustomException as ce:
            return self.tools.output(400, str(ce), {})
        except Exception as e:
            print(f"Error creando año: {e}")
            return self.tools.output(500, "Error creando año.", {})
