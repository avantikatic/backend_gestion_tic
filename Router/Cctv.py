from fastapi import APIRouter, Depends, Body
from sqlalchemy.orm import Session
from Class.Cctv import Cctv
from Config.db import get_db
from Utils.tools import Tools

cctv_router = APIRouter()


# ── CATÁLOGOS ─────────────────────────────────────────────────────────────────

@cctv_router.post('/obtener_catalogos', tags=["CCTV"], response_model=dict)
def obtener_catalogos(db: Session = Depends(get_db)):
    """Devuelve los 5 catálogos: estados_camara, metodos_backup, niveles_acceso, severidades, estados_incidente."""
    return Cctv(db).obtener_catalogos()


# ── DASHBOARD ─────────────────────────────────────────────────────────────────

@cctv_router.post('/dashboard', tags=["CCTV"], response_model=dict)
def cctv_dashboard(db: Session = Depends(get_db)):
    """KPIs y resumen por sede."""
    return Cctv(db).dashboard()


# ── SEDES ─────────────────────────────────────────────────────────────────────

@cctv_router.post('/listar_sedes', tags=["CCTV"], response_model=dict)
def listar_sedes(db: Session = Depends(get_db)):
    return Cctv(db).listar_sedes()


@cctv_router.post('/crear_sede', tags=["CCTV"], response_model=dict)
def crear_sede(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    {
        "nombre":                      str,  // requerido
        "ubicacion_general":           str,
        "responsable_operativo":       str,
        "sistema_grabacion":           str,
        "dias_almacenamiento_estimado": int,
        "observaciones":               str,
        "usuario_creacion":            str
    }
    """
    return Cctv(db).crear_sede(data)


@cctv_router.post('/actualizar_sede', tags=["CCTV"], response_model=dict)
def actualizar_sede(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_sede": int, ...campos... }"""
    id_sede = data.get('id_sede')
    if not id_sede:
        return Tools().output(400, "id_sede requerido.", {})
    return Cctv(db).actualizar_sede(id_sede, data)


@cctv_router.post('/eliminar_sede', tags=["CCTV"], response_model=dict)
def eliminar_sede(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_sede": int }"""
    id_sede = data.get('id_sede')
    if not id_sede:
        return Tools().output(400, "id_sede requerido.", {})
    return Cctv(db).eliminar_sede(id_sede)


# ── CÁMARAS ───────────────────────────────────────────────────────────────────

@cctv_router.post('/listar_camaras', tags=["CCTV"], response_model=dict)
def listar_camaras(data: dict = Body(default={}), db: Session = Depends(get_db)):
    """
    Filtros opcionales: { "query": str, "id_sede": int, "estado": str }
    """
    return Cctv(db).listar_camaras(data)


@cctv_router.post('/crear_camara', tags=["CCTV"], response_model=dict)
def crear_camara(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    {
        "id_sede":                         int,   // requerido
        "codigo_equipo_grabacion":         str,   // requerido
        "ubicacion_fisica":                str,   // requerido
        "estado":                          str,   // activa | inactiva | en_mantenimiento | retirada
        "dias_almacenamiento":             int,
        "metodo_backup":                   str,
        "dias_retencion_backup":           int,
        "fecha_instalacion_actualizacion": str,
        "observaciones":                   str,
        "usuario_creacion":                str
    }
    """
    return Cctv(db).crear_camara(data)


@cctv_router.post('/actualizar_camara', tags=["CCTV"], response_model=dict)
def actualizar_camara(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_camara": int, ...campos... }"""
    id_camara = data.get('id_camara')
    if not id_camara:
        return Tools().output(400, "id_camara requerido.", {})
    return Cctv(db).actualizar_camara(id_camara, data)


@cctv_router.post('/eliminar_camara', tags=["CCTV"], response_model=dict)
def eliminar_camara(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_camara": int }"""
    id_camara = data.get('id_camara')
    if not id_camara:
        return Tools().output(400, "id_camara requerido.", {})
    return Cctv(db).eliminar_camara(id_camara)


# ── CARGOS / PERMISOS ─────────────────────────────────────────────────────────

@cctv_router.post('/listar_cargos', tags=["CCTV"], response_model=dict)
def listar_cargos(db: Session = Depends(get_db)):
    return Cctv(db).listar_cargos()


@cctv_router.post('/crear_cargo', tags=["CCTV"], response_model=dict)
def crear_cargo(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    {
        "nombre":                          str,   // requerido
        "nivel_acceso":                    str,   // sin_acceso | todas_las_sedes | sedes_especificas
        "justificacion_acceso":            str,
        "puede_ver_camaras":               bool,
        "puede_editar_inventario":         bool,
        "puede_administrar_configuracion": bool,
        "ids_sedes":                       [int],
        "usuario_creacion":                str
    }
    """
    return Cctv(db).crear_cargo(data)


@cctv_router.post('/actualizar_cargo', tags=["CCTV"], response_model=dict)
def actualizar_cargo(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_cargo": int, ...campos... }"""
    id_cargo = data.get('id_cargo')
    if not id_cargo:
        return Tools().output(400, "id_cargo requerido.", {})
    return Cctv(db).actualizar_cargo(id_cargo, data)


@cctv_router.post('/eliminar_cargo', tags=["CCTV"], response_model=dict)
def eliminar_cargo(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_cargo": int }"""
    id_cargo = data.get('id_cargo')
    if not id_cargo:
        return Tools().output(400, "id_cargo requerido.", {})
    return Cctv(db).eliminar_cargo(id_cargo)


# ── REGISTROS DE CAMBIOS ──────────────────────────────────────────────────────

@cctv_router.post('/listar_cambios', tags=["CCTV"], response_model=dict)
def listar_cambios(data: dict = Body(default={}), db: Session = Depends(get_db)):
    return Cctv(db).listar_cambios(data)


@cctv_router.post('/crear_cambio', tags=["CCTV"], response_model=dict)
def crear_cambio(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    {
        "fecha_cambio":      str,   // requerido
        "descripcion":       str,   // requerido
        "id_sede":           int,
        "id_camara":         int,
        "observaciones":     str,
        "cargo_responsable": str,
        "usuario_creacion":  str
    }
    """
    return Cctv(db).crear_cambio(data)


# ── REVISIONES ────────────────────────────────────────────────────────────────

@cctv_router.post('/listar_revisiones', tags=["CCTV"], response_model=dict)
def listar_revisiones(data: dict = Body(default={}), db: Session = Depends(get_db)):
    return Cctv(db).listar_revisiones(data)


@cctv_router.post('/crear_revision', tags=["CCTV"], response_model=dict)
def crear_revision(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    {
        "fecha_revision":           str,   // requerido
        "cargo_revisor":            str,
        "camaras_activas":          int,
        "camaras_mantenimiento":    int,
        "camaras_inactivas":        int,
        "almacenamiento_verificado": bool,
        "backups_verificados":      bool,
        "accesos_verificados":      bool,
        "hallazgos":                str,
        "proxima_revision":         str,
        "usuario_creacion":         str
    }
    """
    return Cctv(db).crear_revision(data)


# ── INCIDENTES ────────────────────────────────────────────────────────────────

@cctv_router.post('/listar_incidentes', tags=["CCTV"], response_model=dict)
def listar_incidentes(data: dict = Body(default={}), db: Session = Depends(get_db)):
    return Cctv(db).listar_incidentes(data)


@cctv_router.post('/crear_incidente', tags=["CCTV"], response_model=dict)
def crear_incidente(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    {
        "fecha_incidente":   str,   // requerido
        "titulo":            str,   // requerido
        "descripcion":       str,   // requerido
        "id_sede":           int,
        "id_camara":         int,
        "severidad":         str,   // baja | media | alta | critica
        "estado":            str,   // abierto | en_gestion | cerrado
        "accion_correctiva": str,
        "usuario_creacion":  str
    }
    """
    return Cctv(db).crear_incidente(data)


@cctv_router.post('/actualizar_incidente', tags=["CCTV"], response_model=dict)
def actualizar_incidente(data: dict = Body(...), db: Session = Depends(get_db)):
    """{ "id_incidente": int, ...campos... }"""
    id_incidente = data.get('id_incidente')
    if not id_incidente:
        return Tools().output(400, "id_incidente requerido.", {})
    return Cctv(db).actualizar_incidente(id_incidente, data)
