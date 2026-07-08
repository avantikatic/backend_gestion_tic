from Utils.tools import Tools
from Utils.querys import Querys
from datetime import datetime
import pytz

COLOMBIA_TZ = pytz.timezone('America/Bogota')


def _now_co():
    return datetime.now(COLOMBIA_TZ).replace(tzinfo=None)


def _gen_codigo(prefijo: str, total: int) -> str:
    """Genera código legible: EV-2025-001, AC-2025-003, etc."""
    anio = datetime.now(COLOMBIA_TZ).year
    return f"{prefijo}-{anio}-{str(total + 1).zfill(3)}"


class Contingencia:

    def __init__(self, db):
        self.db     = db
        self.tools  = Tools()
        self.querys = Querys(self.db)

    # ── CATÁLOGOS ──────────────────────────────────────────────────────────────

    def obtener_catalogos(self):
        try:
            data = self.querys.cont_obtener_catalogos()
            return self.tools.output(200, "Catálogos obtenidos.", data)
        except Exception as e:
            print(f"[Contingencia] obtener_catalogos: {e}")
            return self.tools.output(500, "Error obteniendo catálogos.", {})

    # ── DASHBOARD ──────────────────────────────────────────────────────────────

    def obtener_contadores(self):
        try:
            data = self.querys.cont_obtener_contadores()
            return self.tools.output(200, "Contadores obtenidos.", data)
        except Exception as e:
            print(f"[Contingencia] obtener_contadores: {e}")
            return self.tools.output(500, "Error obteniendo contadores.", {})

    def listar_eventos(self, filtros: dict):
        try:
            data = self.querys.cont_listar_eventos(filtros)
            return self.tools.output(200, "Eventos obtenidos.", data)
        except Exception as e:
            print(f"[Contingencia] listar_eventos: {e}")
            return self.tools.output(500, "Error listando eventos.", {})

    # ── EVENTOS ────────────────────────────────────────────────────────────────

    def crear_evento(self, data: dict):
        try:
            campos_req = ['id_tipo_evento', 'id_prioridad', 'id_estado_evento', 'titulo']
            for campo in campos_req:
                if not data.get(campo):
                    return self.tools.output(400, f"Campo requerido faltante: {campo}.", {})

            total = self.querys.cont_contar_eventos_anio()
            data['codigo'] = _gen_codigo('EV', total)

            evento = self.querys.cont_crear_evento(data)

            # Bitácora automática al crear
            self.querys.cont_crear_bitacora({
                'id_evento':        evento['id'],
                'id_tipo_bitacora': self.querys.cont_id_tipo_bitacora('Activacion'),
                'detalle':          'Evento creado y plan de contingencia iniciado.',
                'actor':            data.get('responsable') or 'Sistema',
                'usuario_creacion': data.get('usuario_creacion'),
                'codigo':           _gen_codigo('BT', self.querys.cont_contar_bitacoras_evento(evento['id'])),
            })

            return self.tools.output(201, "Evento creado exitosamente.", evento)
        except Exception as e:
            print(f"[Contingencia] crear_evento: {e}")
            return self.tools.output(500, "Error creando evento.", {})

    def obtener_evento(self, id_evento: int):
        try:
            data = self.querys.cont_obtener_evento(id_evento)
            if not data:
                return self.tools.output(404, "Evento no encontrado.", {})
            return self.tools.output(200, "Evento obtenido.", data)
        except Exception as e:
            print(f"[Contingencia] obtener_evento: {e}")
            return self.tools.output(500, "Error obteniendo evento.", {})

    def actualizar_evento(self, id_evento: int, data: dict):
        try:
            actualizado = self.querys.cont_actualizar_evento(id_evento, data)
            if not actualizado:
                return self.tools.output(404, "Evento no encontrado.", {})
            return self.tools.output(200, "Evento actualizado.", actualizado)
        except Exception as e:
            print(f"[Contingencia] actualizar_evento: {e}")
            return self.tools.output(500, "Error actualizando evento.", {})

    def eliminar_evento(self, id_evento: int):
        try:
            self.querys.cont_eliminar_evento(id_evento)
            return self.tools.output(200, "Evento eliminado.", {})
        except Exception as e:
            print(f"[Contingencia] eliminar_evento: {e}")
            return self.tools.output(500, "Error eliminando evento.", {})

    # ── ACCIONES ───────────────────────────────────────────────────────────────

    def crear_accion(self, data: dict):
        try:
            if not data.get('id_evento') or not data.get('titulo') or not data.get('id_estado_accion'):
                return self.tools.output(400, "id_evento, id_estado_accion y titulo son requeridos.", {})

            total = self.querys.cont_contar_acciones_anio()
            data['codigo'] = _gen_codigo('AC', total)

            accion = self.querys.cont_crear_accion(data)

            self.querys.cont_crear_bitacora({
                'id_evento':        data['id_evento'],
                'id_tipo_bitacora': self.querys.cont_id_tipo_bitacora('Ejecucion'),
                'detalle':          f"Acción ajustada registrada: {data['titulo']}",
                'actor':            data.get('responsable') or 'Sistema',
                'usuario_creacion': data.get('usuario_creacion'),
                'codigo':           _gen_codigo('BT', self.querys.cont_contar_bitacoras_evento(data['id_evento'])),
            })

            return self.tools.output(201, "Acción creada exitosamente.", accion)
        except Exception as e:
            print(f"[Contingencia] crear_accion: {e}")
            return self.tools.output(500, "Error creando acción.", {})

    def listar_acciones(self, id_evento: int):
        try:
            data = self.querys.cont_listar_acciones(id_evento)
            return self.tools.output(200, "Acciones obtenidas.", data)
        except Exception as e:
            print(f"[Contingencia] listar_acciones: {e}")
            return self.tools.output(500, "Error listando acciones.", {})

    def actualizar_accion(self, id_accion: int, data: dict):
        try:
            actualizado = self.querys.cont_actualizar_accion(id_accion, data)
            if not actualizado:
                return self.tools.output(404, "Acción no encontrada.", {})
            return self.tools.output(200, "Acción actualizada.", actualizado)
        except Exception as e:
            print(f"[Contingencia] actualizar_accion: {e}")
            return self.tools.output(500, "Error actualizando acción.", {})

    def eliminar_accion(self, id_accion: int):
        try:
            self.querys.cont_eliminar_accion(id_accion)
            return self.tools.output(200, "Acción eliminada.", {})
        except Exception as e:
            print(f"[Contingencia] eliminar_accion: {e}")
            return self.tools.output(500, "Error eliminando acción.", {})

    # ── RECUPERACIÓN ───────────────────────────────────────────────────────────

    def guardar_recuperacion(self, data: dict):
        try:
            resultado = self.querys.cont_upsert_recuperacion(data)

            self.querys.cont_crear_bitacora({
                'id_evento':        data['id_evento'],
                'id_tipo_bitacora': self.querys.cont_id_tipo_bitacora('Cierre'),
                'detalle':          'Resultado de recuperación actualizado.',
                'actor':            data.get('usuario_actualizacion') or 'Sistema',
                'usuario_creacion': data.get('usuario_actualizacion'),
                'codigo':           _gen_codigo('BT', self.querys.cont_contar_bitacoras_evento(data['id_evento'])),
            })

            return self.tools.output(200, "Recuperación guardada.", resultado)
        except Exception as e:
            print(f"[Contingencia] guardar_recuperacion: {e}")
            return self.tools.output(500, "Error guardando recuperación.", {})

    def obtener_recuperacion(self, id_evento: int):
        try:
            data = self.querys.cont_obtener_recuperacion(id_evento)
            return self.tools.output(200, "Recuperación obtenida.", data or {})
        except Exception as e:
            print(f"[Contingencia] obtener_recuperacion: {e}")
            return self.tools.output(500, "Error obteniendo recuperación.", {})

    # ── DOCUMENTOS ─────────────────────────────────────────────────────────────

    def crear_documento(self, data: dict):
        try:
            if not data.get('id_evento') or not data.get('nombre') or not data.get('id_estado_documento'):
                return self.tools.output(400, "id_evento, id_estado_documento y nombre son requeridos.", {})

            total = self.querys.cont_contar_documentos_anio()
            data['codigo_interno'] = _gen_codigo('DOC', total)

            documento = self.querys.cont_crear_documento(data)

            self.querys.cont_crear_bitacora({
                'id_evento':        data['id_evento'],
                'id_tipo_bitacora': self.querys.cont_id_tipo_bitacora('Documentacion'),
                'detalle':          f"Documento agregado: {data['nombre']}",
                'actor':            data.get('responsable') or 'Gestión documental',
                'usuario_creacion': data.get('usuario_creacion'),
                'codigo':           _gen_codigo('BT', self.querys.cont_contar_bitacoras_evento(data['id_evento'])),
            })

            return self.tools.output(201, "Documento creado exitosamente.", documento)
        except Exception as e:
            print(f"[Contingencia] crear_documento: {e}")
            return self.tools.output(500, "Error creando documento.", {})

    def listar_documentos(self, id_evento: int):
        try:
            data = self.querys.cont_listar_documentos(id_evento)
            return self.tools.output(200, "Documentos obtenidos.", data)
        except Exception as e:
            print(f"[Contingencia] listar_documentos: {e}")
            return self.tools.output(500, "Error listando documentos.", {})

    def eliminar_documento(self, id_documento: int):
        try:
            self.querys.cont_eliminar_documento(id_documento)
            return self.tools.output(200, "Documento eliminado.", {})
        except Exception as e:
            print(f"[Contingencia] eliminar_documento: {e}")
            return self.tools.output(500, "Error eliminando documento.", {})

    # ── BITÁCORAS ──────────────────────────────────────────────────────────────

    def crear_bitacora(self, data: dict):
        try:
            if not data.get('id_evento') or not data.get('detalle') or not data.get('id_tipo_bitacora'):
                return self.tools.output(400, "id_evento, id_tipo_bitacora y detalle son requeridos.", {})

            total = self.querys.cont_contar_bitacoras_evento(data['id_evento'])
            data['codigo'] = _gen_codigo('BT', total)

            bitacora = self.querys.cont_crear_bitacora(data)
            return self.tools.output(201, "Bitácora creada exitosamente.", bitacora)
        except Exception as e:
            print(f"[Contingencia] crear_bitacora: {e}")
            return self.tools.output(500, "Error creando bitácora.", {})

    def listar_bitacoras(self, id_evento: int):
        try:
            data = self.querys.cont_listar_bitacoras(id_evento)
            return self.tools.output(200, "Bitácoras obtenidas.", data)
        except Exception as e:
            print(f"[Contingencia] listar_bitacoras: {e}")
            return self.tools.output(500, "Error listando bitácoras.", {})
