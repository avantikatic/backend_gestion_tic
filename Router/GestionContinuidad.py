from fastapi import APIRouter, Depends, Body
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session
from Class.GestionContinuidad import GestionContinuidad
from Config.db import get_db
from Utils.file_handler import FileHandler
from typing import Optional
import os

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

@gestion_continuidad_router.post('/crear_registro_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def crear_registro_gsc(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Crea un registro GSC completo con todas sus secciones.
    
    Estructura esperada:
    {
        "id_modulo": int,
        "resumen": str,
        "descripcion": str,
        "id_estado": int,
        "notificar_gerencia": bool,
        "sistemas_afectados": [int],
        "evidencias": [
            {
                "id_tipo_evidencia": int,
                "observacion": str,
                "datos_especificos": {...}
            }
        ],
        "datos_modulo": {...},
        "usuario_creacion": str,
        "resultados_iniciales": [str] (opcional - lista de textos para bitácora)
    }
    """
    response = GestionContinuidad(db).crear_registro_gsc(data)
    return response

@gestion_continuidad_router.post('/obtener_registro_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_registro_gsc(data: dict = Body(...), db: Session = Depends(get_db)):
    """Obtiene un registro GSC completo por su ID"""
    id_registro = data.get('id_registro')
    if not id_registro:
        return {"status": 400, "message": "ID de registro requerido", "data": {}}
    
    response = GestionContinuidad(db).obtener_registro_gsc(id_registro)
    return response

@gestion_continuidad_router.post('/listar_registros_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def listar_registros_gsc(data: dict = Body(default={}), db: Session = Depends(get_db)):
    """
    Lista registros GSC con filtros opcionales.
    
    Filtros disponibles:
    {
        "id_modulo": int,
        "id_estado": int,
        "fecha_desde": str (ISO format),
        "fecha_hasta": str (ISO format),
        "limite": int,
        "offset": int
    }
    """
    filtros = data if data else {}
    response = GestionContinuidad(db).listar_registros_gsc(filtros)
    return response

@gestion_continuidad_router.post('/obtener_contadores_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def obtener_contadores_gsc(data: dict = Body(default={}), db: Session = Depends(get_db)):
    """
    Obtiene contadores de registros por estado.
    Usado para KPIs en el dashboard.
    
    Parámetros opcionales:
    {
        "id_modulo": int  // Si se especifica, filtra por módulo. Si no, retorna totales globales
    }
    
    Retorna:
    {
        "total": int,
        "abiertos": int,
        "en_analisis": int,
        "mitigados": int,
        "cerrados": int
    }
    """
    filtros = data if data else {}
    response = GestionContinuidad(db).obtener_contadores_gsc(filtros)
    return response

@gestion_continuidad_router.post('/actualizar_registro_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def actualizar_registro_gsc(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Actualiza un registro GSC existente.
    
    Estructura esperada:
    {
        "id_registro": int,
        "resumen": str (opcional),
        "descripcion": str (opcional),
        "id_estado": int (opcional),
        "notificar_gerencia": bool (opcional),
        "sistemas_afectados": [int] (opcional),
        "datos_modulo": {...} (opcional),
        "usuario_actualizacion": str
    }
    """
    id_registro = data.get('id_registro')
    if not id_registro:
        return {"status": 400, "message": "ID de registro requerido", "data": {}}
    
    response = GestionContinuidad(db).actualizar_registro_gsc(id_registro, data)
    return response

@gestion_continuidad_router.post('/eliminar_registro_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def eliminar_registro_gsc(data: dict = Body(...), db: Session = Depends(get_db)):
    """Elimina (desactiva) un registro GSC"""
    id_registro = data.get('id_registro')
    if not id_registro:
        return {"status": 400, "message": "ID de registro requerido", "data": {}}
    
    response = GestionContinuidad(db).eliminar_registro_gsc(id_registro)
    return response

@gestion_continuidad_router.post('/crear_resultado_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def crear_resultado_gsc(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Crea un nuevo resultado (entrada de bitácora) para un registro GSC.
    
    Estructura esperada:
    {
        "id_registro": int,
        "texto": str
    }
    """
    id_registro = data.get('id_registro')
    texto = data.get('texto')
    
    if not id_registro or not texto:
        return {"status": 400, "message": "ID de registro y texto son requeridos", "data": {}}
    
    response = GestionContinuidad(db).crear_resultado_gsc(data)
    return response

@gestion_continuidad_router.post('/listar_resultados_gsc', tags=["GESTION_CONTINUIDAD"], response_model=dict)
def listar_resultados_gsc(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Lista todos los resultados (entradas de bitácora) de un registro GSC.
    Ordenados por fecha de creación descendente (más reciente primero).
    
    Parámetros:
    {
        "id_registro": int
    }
    """
    id_registro = data.get('id_registro')
    
    if not id_registro:
        return {"status": 400, "message": "ID de registro requerido", "data": {}}
    
    response = GestionContinuidad(db).listar_resultados_gsc(id_registro)
    return response

@gestion_continuidad_router.get('/obtener_imagen_evidencia/{nombre_archivo}', tags=["GESTION_CONTINUIDAD"])
def obtener_imagen_evidencia(nombre_archivo: str):
    """Sirve una imagen de evidencia desde el sistema de archivos"""
    try:
        file_handler = FileHandler()
        ruta_completa = file_handler.obtener_ruta_completa(nombre_archivo)
        
        if not ruta_completa.exists():
            return {"status": 404, "message": "Imagen no encontrada"}
        
        return FileResponse(
            path=str(ruta_completa),
            media_type="image/png",
            filename=nombre_archivo
        )
    except Exception as e:
        print(f"Error sirviendo imagen: {e}")
        return {"status": 500, "message": "Error al obtener la imagen"}
