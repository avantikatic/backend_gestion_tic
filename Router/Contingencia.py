from fastapi import APIRouter, Depends, Body
from sqlalchemy.orm import Session
from Class.Contingencia import Contingencia
from Config.db import get_db

contingencia_router = APIRouter()


# ── CATÁLOGOS ─────────────────────────────────────────────────────────────────

@contingencia_router.post('/obtener_catalogos', tags=["CONTINGENCIA"], response_model=dict)
def obtener_catalogos(db: Session = Depends(get_db)):
    """Retorna todos los catálogos (listas de selección) del módulo."""
    return Contingencia(db).obtener_catalogos()


# ── DASHBOARD / CENTRO DE CONTROL ─────────────────────────────────────────────

@contingencia_router.post('/obtener_contadores', tags=["CONTINGENCIA"], response_model=dict)
def obtener_contadores(db: Session = Depends(get_db)):
    """KPI cards: eventos abiertos, acciones completadas, documentos, bitácoras."""
    return Contingencia(db).obtener_contadores()


@contingencia_router.post('/listar_eventos', tags=["CONTINGENCIA"], response_model=dict)
def listar_eventos(data: dict = Body(default={}), db: Session = Depends(get_db)):
    """
    Lista eventos activos con datos de catálogo resueltos.

    Filtros opcionales:
    {
        "query":         str,   // busca en código, título, área, responsable
        "id_tipo_evento": int,
        "id_estado_evento": int
    }
    """
    return Contingencia(db).listar_eventos(data)


# ── EVENTOS ───────────────────────────────────────────────────────────────────

@contingencia_router.post('/crear_evento', tags=["CONTINGENCIA"], response_model=dict)
def crear_evento(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Crea un evento de contingencia.

    {
        "id_tipo_evento":    int,   // requerido
        "id_prioridad":      int,   // requerido
        "id_estado_evento":  int,   // requerido
        "titulo":            str,   // requerido
        "area":              str,
        "responsable":       str,
        "fecha_inicio":      str,
        "rto_objetivo":      str,
        "rpo_objetivo":      str,
        "impacto":           str,
        "causa":             str,
        "usuario_creacion":  str
    }
    """
    return Contingencia(db).crear_evento(data)


@contingencia_router.post('/obtener_evento', tags=["CONTINGENCIA"], response_model=dict)
def obtener_evento(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Retorna un evento completo con sus relaciones.

    { "id_evento": int }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).obtener_evento(id_evento)


@contingencia_router.post('/actualizar_evento', tags=["CONTINGENCIA"], response_model=dict)
def actualizar_evento(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Actualiza campos de un evento existente.

    { "id_evento": int, ...campos a actualizar... }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).actualizar_evento(id_evento, data)


@contingencia_router.post('/eliminar_evento', tags=["CONTINGENCIA"], response_model=dict)
def eliminar_evento(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Desactiva un evento (soft delete) y todos sus registros hijos.

    { "id_evento": int }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).eliminar_evento(id_evento)


# ── ACCIONES AJUSTADAS ────────────────────────────────────────────────────────

@contingencia_router.post('/crear_accion', tags=["CONTINGENCIA"], response_model=dict)
def crear_accion(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Registra una acción ajustada para un evento.

    {
        "id_evento":          int,   // requerido
        "id_estado_accion":   int,   // requerido
        "titulo":             str,   // requerido
        "responsable":        str,
        "fecha_compromiso":   str,
        "evidencia":          str,
        "control_asociado":   str,
        "usuario_creacion":   str
    }
    """
    return Contingencia(db).crear_accion(data)


@contingencia_router.post('/listar_acciones', tags=["CONTINGENCIA"], response_model=dict)
def listar_acciones(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Lista acciones de un evento.

    { "id_evento": int }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).listar_acciones(id_evento)


@contingencia_router.post('/actualizar_accion', tags=["CONTINGENCIA"], response_model=dict)
def actualizar_accion(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Actualiza una acción existente.

    { "id_accion": int, ...campos a actualizar... }
    """
    id_accion = data.get('id_accion')
    if not id_accion:
        from Utils.tools import Tools
        return Tools().output(400, "id_accion requerido.", {})
    return Contingencia(db).actualizar_accion(id_accion, data)


@contingencia_router.post('/eliminar_accion', tags=["CONTINGENCIA"], response_model=dict)
def eliminar_accion(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Desactiva una acción (soft delete).

    { "id_accion": int }
    """
    id_accion = data.get('id_accion')
    if not id_accion:
        from Utils.tools import Tools
        return Tools().output(400, "id_accion requerido.", {})
    return Contingencia(db).eliminar_accion(id_accion)


# ── RECUPERACIÓN ──────────────────────────────────────────────────────────────

@contingencia_router.post('/guardar_recuperacion', tags=["CONTINGENCIA"], response_model=dict)
def guardar_recuperacion(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Upsert del registro de recuperación de un evento (uno por evento).

    {
        "id_evento":                  int,   // requerido
        "id_resultado_recuperacion":  int,
        "tiempo_real":                str,
        "datos_recuperados":          str,
        "observaciones":              str,
        "servicio_alterno":           bool,
        "integridad_documental":      bool,
        "informe_final":              bool,
        "lecciones_aprendidas":       bool,
        "usuario_actualizacion":      str
    }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).guardar_recuperacion(data)


@contingencia_router.post('/obtener_recuperacion', tags=["CONTINGENCIA"], response_model=dict)
def obtener_recuperacion(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Retorna el registro de recuperación de un evento.

    { "id_evento": int }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).obtener_recuperacion(id_evento)


# ── DOCUMENTOS ────────────────────────────────────────────────────────────────

@contingencia_router.post('/crear_documento', tags=["CONTINGENCIA"], response_model=dict)
def crear_documento(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Agrega un documento al expediente.

    {
        "id_evento":            int,   // requerido
        "id_estado_documento":  int,   // requerido
        "nombre":               str,   // requerido
        "codigo_documento":     str,
        "responsable":          str,
        "fecha_documento":      str,
        "usuario_creacion":     str
    }
    """
    return Contingencia(db).crear_documento(data)


@contingencia_router.post('/listar_documentos', tags=["CONTINGENCIA"], response_model=dict)
def listar_documentos(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Lista documentos de un evento.

    { "id_evento": int }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).listar_documentos(id_evento)


@contingencia_router.post('/eliminar_documento', tags=["CONTINGENCIA"], response_model=dict)
def eliminar_documento(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Desactiva un documento (soft delete).

    { "id_documento": int }
    """
    id_documento = data.get('id_documento')
    if not id_documento:
        from Utils.tools import Tools
        return Tools().output(400, "id_documento requerido.", {})
    return Contingencia(db).eliminar_documento(id_documento)


# ── BITÁCORAS ─────────────────────────────────────────────────────────────────

@contingencia_router.post('/crear_bitacora', tags=["CONTINGENCIA"], response_model=dict)
def crear_bitacora(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Agrega una entrada de bitácora al expediente.

    {
        "id_evento":          int,   // requerido
        "id_tipo_bitacora":   int,   // requerido
        "detalle":            str,   // requerido
        "hora_registro":      str,
        "actor":              str,
        "usuario_creacion":   str
    }
    """
    return Contingencia(db).crear_bitacora(data)


@contingencia_router.post('/listar_bitacoras', tags=["CONTINGENCIA"], response_model=dict)
def listar_bitacoras(data: dict = Body(...), db: Session = Depends(get_db)):
    """
    Lista bitácoras de un evento, ordenadas más reciente primero.

    { "id_evento": int }
    """
    id_evento = data.get('id_evento')
    if not id_evento:
        from Utils.tools import Tools
        return Tools().output(400, "id_evento requerido.", {})
    return Contingencia(db).listar_bitacoras(id_evento)
