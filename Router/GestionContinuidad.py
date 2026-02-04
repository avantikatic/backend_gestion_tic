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
@gestion_continuidad_router.post('/obtener_tipos_evidencia_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_tipos_evidencia_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los tipos de evidencia disponibles para el módulo GSC"""
    response = GestionContinuidad(db).obtener_tipos_evidencia_gsc()
    return response

@gestion_continuidad_router.post('/obtener_origenes_plataforma_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_origenes_plataforma_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los orígenes de plataforma disponibles para alertas en el módulo GSC"""
    response = GestionContinuidad(db).obtener_origenes_plataforma_gsc()
    return response

@gestion_continuidad_router.post('/obtener_fuentes_seguridad_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_fuentes_seguridad_gsc(db: Session = Depends(get_db)):
    """Obtiene todas las fuentes de seguridad disponibles para el módulo SEG"""
    response = GestionContinuidad(db).obtener_fuentes_seguridad_gsc()
    return response

@gestion_continuidad_router.post('/obtener_impactos_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_impactos_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los niveles de impacto disponibles para el módulo SEG"""
    response = GestionContinuidad(db).obtener_impactos_gsc()
    return response

@gestion_continuidad_router.post('/obtener_riesgos_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_riesgos_gsc(db: Session = Depends(get_db)):
    """Obtiene todos los niveles de riesgo disponibles para el módulo MNT"""
    response = GestionContinuidad(db).obtener_riesgos_gsc()
    return response