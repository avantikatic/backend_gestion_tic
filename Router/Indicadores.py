from fastapi import APIRouter, Request, Depends, Query
from sqlalchemy.orm import Session
from Class.Indicadores import Indicadores
from Utils.decorator import http_decorator
from Config.db import get_db

indicadores_router = APIRouter()

@indicadores_router.post('/obtener_indicadores_gestion', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_indicadores_gestion(request: Request, db: Session = Depends(get_db)):
    """Obtiene indicadores de gestión mensual: tickets completados, oportunos y no oportunos"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).obtener_indicadores_gestion(data)
    return response

@indicadores_router.post('/obtener_indicadores_estrategicos', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_indicadores_estrategicos(request: Request, db: Session = Depends(get_db)):
    """Obtiene indicadores de tickets estratégicos agrupados por origen_estrategico"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).obtener_indicadores_estrategicos(data)
    return response

@indicadores_router.post('/obtener_observacion_mes', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_observacion_mes(request: Request, db: Session = Depends(get_db)):
    """Obtiene la observación de un mes específico del informe de gestión"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).obtener_observacion_mes(data)
    return response

@indicadores_router.post('/guardar_observacion_mes', tags=["INDICADORES"], response_model=dict)
@http_decorator
def guardar_observacion_mes(request: Request, db: Session = Depends(get_db)):
    """Guarda o actualiza la observación de un mes específico del informe de gestión"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).guardar_observacion_mes(data)
    return response

@indicadores_router.post('/obtener_analisis_causas', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_analisis_causas(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los análisis de causas y acciones de un año específico"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).obtener_analisis_causas(data)
    return response

@indicadores_router.post('/guardar_analisis_causas', tags=["INDICADORES"], response_model=dict)
@http_decorator
def guardar_analisis_causas(request: Request, db: Session = Depends(get_db)):
    """Guarda o actualiza un análisis de causas y acciones. Valida que no exista otro registro con el mismo año y mes."""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).guardar_analisis_causas(data)
    return response

@indicadores_router.post('/obtener_tickets_periodo', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_tickets_periodo(request: Request, db: Session = Depends(get_db)):
    """Obtiene los tickets del periodo especificado (año y mes)"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).obtener_tickets_periodo(data)
    return response

@indicadores_router.post('/obtener_anios', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_anios(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los años disponibles en el sistema"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).obtener_anios_disponibles(data)
    return response

@indicadores_router.post('/crear_anio', tags=["INDICADORES"], response_model=dict)
@http_decorator
def crear_anio(request: Request, db: Session = Depends(get_db)):
    """Crea un nuevo año en el sistema"""
    data = getattr(request.state, "json_data", {})
    response = Indicadores(db).crear_anio(data)
    return response
