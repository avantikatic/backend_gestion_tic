from fastapi import APIRouter, Request, Depends, Query
from sqlalchemy.orm import Session
from Class.Dashboard import Dashboard
from Utils.decorator import http_decorator
from Config.db import get_db

dashboard_router = APIRouter()

@dashboard_router.post('/obtener_metricas_dashboard', tags=["DASHBOARD"], response_model=dict)
@http_decorator
def obtener_metricas_dashboard(request: Request, db: Session = Depends(get_db)):
    """Obtiene métricas principales del dashboard: totales, tipos, prioridades y estados"""
    data = getattr(request.state, "json_data", {})
    response = Dashboard(db).obtener_metricas_dashboard(data)
    return response

@dashboard_router.post('/obtener_indicadores_gestion', tags=["DASHBOARD"], response_model=dict)
@http_decorator
def obtener_indicadores_gestion(request: Request, db: Session = Depends(get_db)):
    """Obtiene indicadores de gestión mensual: tickets completados, oportunos y no oportunos"""
    data = getattr(request.state, "json_data", {})
    response = Dashboard(db).obtener_indicadores_gestion(data)
    return response

@dashboard_router.post('/obtener_observacion_mes', tags=["DASHBOARD"], response_model=dict)
@http_decorator
def obtener_observacion_mes(request: Request, db: Session = Depends(get_db)):
    """Obtiene la observación de un mes específico del informe de gestión"""
    data = getattr(request.state, "json_data", {})
    response = Dashboard(db).obtener_observacion_mes(data)
    return response

@dashboard_router.post('/guardar_observacion_mes', tags=["DASHBOARD"], response_model=dict)
@http_decorator
def guardar_observacion_mes(request: Request, db: Session = Depends(get_db)):
    """Guarda o actualiza la observación de un mes específico del informe de gestión"""
    data = getattr(request.state, "json_data", {})
    response = Dashboard(db).guardar_observacion_mes(data)
    return response

@dashboard_router.post('/obtener_analisis_causas', tags=["DASHBOARD"], response_model=dict)
@http_decorator
def obtener_analisis_causas(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los análisis de causas y acciones de un año específico"""
    data = getattr(request.state, "json_data", {})
    response = Dashboard(db).obtener_analisis_causas(data)
    return response

@dashboard_router.post('/guardar_analisis_causas', tags=["DASHBOARD"], response_model=dict)
@http_decorator
def guardar_analisis_causas(request: Request, db: Session = Depends(get_db)):
    """Guarda o actualiza un análisis de causas y acciones. Valida que no exista otro registro con el mismo año y mes."""
    data = getattr(request.state, "json_data", {})
    response = Dashboard(db).guardar_analisis_causas(data)
    return response
