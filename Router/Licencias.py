from fastapi import APIRouter, Request, Depends, Query, Path
from sqlalchemy.orm import Session
from fastapi.responses import StreamingResponse
from Class.Licencias import Licencias
from Utils.decorator import http_decorator
from Config.db import get_db
from datetime import datetime

licencias_router = APIRouter()

# ===================================================
# ENDPOINTS DE LICENCIAS (CRUD)
# ===================================================

@licencias_router.post('/crear', tags=["LICENCIAS"], response_model=dict)
@http_decorator
def crear_licencia(request: Request, db: Session = Depends(get_db)):
    """Crea una nueva licencia/servicio"""
    data = getattr(request.state, "json_data", {})
    response = Licencias(db).crear_licencia(data)
    return response

@licencias_router.post('/obtener', tags=["LICENCIAS"], response_model=dict)
@http_decorator
def obtener_licencias(
    request: Request, 
    db: Session = Depends(get_db),
    incluir_bajas: bool = Query(True, description="Incluir licencias dadas de baja"),
    proveedor: str = Query(None, description="Filtrar por proveedor"),
    tipo_servicio: str = Query(None, description="Filtrar por tipo de servicio"),
    page: int = Query(1, ge=1, description="Número de página"),
    per_page: int = Query(5, ge=1, le=100, description="Registros por página")
):
    """Obtiene todas las licencias/servicios con paginación"""
    filtros = {
        'incluirBajas': incluir_bajas,
        'proveedor': proveedor,
        'tipoServicio': tipo_servicio
    }
    response = Licencias(db).obtener_licencias(filtros, page, per_page)
    return response

@licencias_router.post('/obtener/{licencia_id}', tags=["LICENCIAS"], response_model=dict)
@http_decorator
def obtener_licencia_por_id(
    request: Request,
    licencia_id: str = Path(..., description="ID de la licencia"),
    db: Session = Depends(get_db)
):
    """Obtiene una licencia específica por su ID"""
    response = Licencias(db).obtener_licencia_por_id(licencia_id)
    return response

@licencias_router.put('/actualizar/{licencia_id}', tags=["LICENCIAS"], response_model=dict)
@http_decorator
def actualizar_licencia(
    request: Request,
    licencia_id: str = Path(..., description="ID de la licencia"),
    db: Session = Depends(get_db)
):
    """Actualiza una licencia existente"""
    data = getattr(request.state, "json_data", {})
    response = Licencias(db).actualizar_licencia(licencia_id, data)
    return response

@licencias_router.delete('/eliminar/{licencia_id}', tags=["LICENCIAS"], response_model=dict)
@http_decorator
def eliminar_licencia(
    request: Request,
    licencia_id: str = Path(..., description="ID de la licencia"),
    db: Session = Depends(get_db)
):
    """Elimina (marca como inactiva) una licencia"""
    response = Licencias(db).eliminar_licencia(licencia_id)
    return response

@licencias_router.post('/historial/{licencia_id}', tags=["LICENCIAS"], response_model=dict)
@http_decorator
def obtener_historial_licencia(
    request: Request,
    licencia_id: int = Path(..., description="ID de la licencia"),
    db: Session = Depends(get_db)
):
    """Obtiene el historial de cambios de una licencia"""
    response = Licencias(db).obtener_historial_licencia(licencia_id)
    return response

# ===================================================
# ENDPOINTS DE CATÁLOGOS/MAESTROS
# ===================================================

@licencias_router.post('/catalogos/tipos-servicio', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def obtener_tipos_servicio(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los tipos de servicio disponibles"""
    response = Licencias(db).obtener_tipos_servicio()
    return response

@licencias_router.post('/catalogos/proveedores', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def obtener_proveedores(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los proveedores disponibles"""
    response = Licencias(db).obtener_proveedores()
    return response

@licencias_router.post('/catalogos/productos-servicios', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def obtener_productos_servicios(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los productos/servicios disponibles"""
    response = Licencias(db).obtener_productos_servicios()
    return response

@licencias_router.post('/catalogos/metodos-pago', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def obtener_metodos_pago(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los métodos de pago disponibles"""
    response = Licencias(db).obtener_metodos_pago()
    return response

@licencias_router.post('/catalogos/proveedores/crear', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def crear_proveedor(request: Request, db: Session = Depends(get_db)):
    """Crea un nuevo proveedor"""
    data = getattr(request.state, "json_data", {})
    nombre = data.get('nombre', '')
    response = Licencias(db).crear_proveedor(nombre)
    return response

@licencias_router.post('/catalogos/productos-servicios/crear', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def crear_producto_servicio(request: Request, db: Session = Depends(get_db)):
    """Crea un nuevo producto/servicio"""
    data = getattr(request.state, "json_data", {})
    nombre = data.get('nombre', '')
    response = Licencias(db).crear_producto_servicio(nombre)
    return response

@licencias_router.post('/catalogos/tipos-servicio/crear', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def crear_tipo_servicio(request: Request, db: Session = Depends(get_db)):
    """Crea un nuevo tipo de servicio"""
    data = getattr(request.state, "json_data", {})
    nombre = data.get('nombre', '')
    response = Licencias(db).crear_tipo_servicio(nombre)
    return response

@licencias_router.post('/catalogos/metodos-pago/crear', tags=["CATALOGOS"], response_model=dict)
@http_decorator
def crear_metodo_pago(request: Request, db: Session = Depends(get_db)):
    """Crea un nuevo método de pago"""
    data = getattr(request.state, "json_data", {})
    nombre = data.get('nombre', '')
    response = Licencias(db).crear_metodo_pago(nombre)
    return response

# ===================================================
# OTROS ENDPOINTS
# ===================================================

# ===================================================
# ENDPOINTS DE REVISIONES GENERALES
# ===================================================

@licencias_router.post('/tipos-revision', tags=["REVISIONES"], response_model=dict)
@http_decorator
def obtener_tipos_revision(request: Request, db: Session = Depends(get_db)):
    """Obtiene todos los tipos de revisión disponibles"""
    response = Licencias(db).obtener_tipos_revision()
    return response

@licencias_router.post('/revisiones/crear', tags=["REVISIONES"], response_model=dict)
@http_decorator
def crear_revision(request: Request, db: Session = Depends(get_db)):
    """Crea una nueva revisión general"""
    data = getattr(request.state, "json_data", {})
    response = Licencias(db).crear_revision(data)
    return response

@licencias_router.post('/revisiones/obtener', tags=["REVISIONES"], response_model=dict)
@http_decorator
def obtener_revisiones(
    request: Request,
    db: Session = Depends(get_db),
    page: int = Query(1, ge=1, description="Número de página"),
    per_page: int = Query(5, ge=1, le=100, description="Registros por página")
):
    """Obtiene todas las revisiones con paginación"""
    response = Licencias(db).obtener_revisiones(page, per_page)
    return response

@licencias_router.put('/revisiones/eliminar', tags=["REVISIONES"], response_model=dict)
@http_decorator
def eliminar_revision(
    request: Request,
    db: Session = Depends(get_db)
):
    """Elimina (marca como inactiva) una revisión"""
    data = getattr(request.state, "json_data", {})
    response = Licencias(db).eliminar_revision(data)
    return response

@licencias_router.post('/exportar-excel', tags=["LICENCIAS"])
def exportar_licencias_excel(
    request: Request,
    db: Session = Depends(get_db)
):
    """Exporta todas las licencias filtradas a Excel"""
    try:
        data = getattr(request.state, "json_data", {})

        # Generar Excel
        excel_file = Licencias(db).exportar_licencias_excel(data)
        
        # Nombre del archivo con fecha
        filename = f"Control-licencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return StreamingResponse(
            excel_file,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        print(f"Error en endpoint exportar Excel: {e}")
        return {"error": str(e)}

# ===================================================
# OTROS ENDPOINTS
# ===================================================

@licencias_router.post('/obtener_indicadores_gestion', tags=["INDICADORES"], response_model=dict)
@http_decorator
def obtener_indicadores_gestion(request: Request, db: Session = Depends(get_db)):
    """Obtiene indicadores de gestión mensual: tickets completados, oportunos y no oportunos"""
    data = getattr(request.state, "json_data", {})
    response = Licencias(db).obtener_indicadores_gestion(data)
    return response

