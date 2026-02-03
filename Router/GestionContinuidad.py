from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session
from Class.GestionContinuidad import GestionContinuidad
from Config.db import get_db

gestion_continuidad_router = APIRouter()

@gestion_continuidad_router.post('/obtener_estados_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_estados_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los estados disponibles para el módulo de Gestión de Seguridad y Continuidad"""
    response = GestionContinuidad(db).obtener_estados_gsc()
    return response

@gestion_continuidad_router.post('/obtener_sistemas_afectados_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_sistemas_afectados_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los sistemas afectados disponibles para el módulo de Gestión de Seguridad y Continuidad"""
    response = GestionContinuidad(db).obtener_sistemas_afectados_gsc()
    return response

@gestion_continuidad_router.post('/obtener_modulos_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_modulos_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los módulos disponibles para el módulo de Gestión de Seguridad y Continuidad"""
    response = GestionContinuidad(db).obtener_modulos_gsc()
    return response
