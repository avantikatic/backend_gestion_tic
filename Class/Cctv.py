from Utils.tools import Tools
from Utils.querys import Querys


class Cctv:

    def __init__(self, db):
        self.db     = db
        self.tools  = Tools()
        self.querys = Querys(self.db)

    # ── Catálogos ──────────────────────────────────────────────────────────────

    def obtener_catalogos(self):
        try:
            data = self.querys.cctv_obtener_catalogos()
            return self.tools.output(200, "Catalogos obtenidos.", data)
        except Exception as e:
            print(f"[Cctv] obtener_catalogos: {e}")
            return self.tools.output(500, "Error obteniendo catalogos.", {})

    # ── Dashboard ──────────────────────────────────────────────────────────────

    def dashboard(self):
        try:
            data = self.querys.cctv_dashboard()
            return self.tools.output(200, "Dashboard obtenido.", data)
        except Exception as e:
            print(f"[Cctv] dashboard: {e}")
            return self.tools.output(500, "Error obteniendo dashboard.", {})

    # ── Sedes ──────────────────────────────────────────────────────────────────

    def listar_sedes(self):
        try:
            data = self.querys.cctv_listar_sedes()
            return self.tools.output(200, "Sedes obtenidas.", data)
        except Exception as e:
            print(f"[Cctv] listar_sedes: {e}")
            return self.tools.output(500, "Error listando sedes.", {})

    def crear_sede(self, data: dict):
        try:
            if not data.get('nombre'):
                return self.tools.output(400, "El campo nombre es requerido.", {})
            sede = self.querys.cctv_crear_sede(data)
            return self.tools.output(201, "Sede creada.", sede)
        except Exception as e:
            print(f"[Cctv] crear_sede: {e}")
            return self.tools.output(500, "Error creando sede.", {})

    def actualizar_sede(self, id_sede: int, data: dict):
        try:
            sede = self.querys.cctv_actualizar_sede(id_sede, data)
            if not sede:
                return self.tools.output(404, "Sede no encontrada.", {})
            return self.tools.output(200, "Sede actualizada.", sede)
        except Exception as e:
            print(f"[Cctv] actualizar_sede: {e}")
            return self.tools.output(500, "Error actualizando sede.", {})

    def eliminar_sede(self, id_sede: int):
        try:
            self.querys.cctv_eliminar_sede(id_sede)
            return self.tools.output(200, "Sede eliminada.", {})
        except Exception as e:
            print(f"[Cctv] eliminar_sede: {e}")
            return self.tools.output(500, "Error eliminando sede.", {})

    # ── Cámaras ────────────────────────────────────────────────────────────────

    def listar_camaras(self, filtros: dict = {}):
        try:
            data = self.querys.cctv_listar_camaras(filtros)
            return self.tools.output(200, "Cámaras obtenidas.", data)
        except Exception as e:
            print(f"[Cctv] listar_camaras: {e}")
            return self.tools.output(500, "Error listando cámaras.", {})

    def crear_camara(self, data: dict):
        try:
            if not data.get('id_sede') or not data.get('codigo_equipo_grabacion') or not data.get('ubicacion_fisica'):
                return self.tools.output(400, "id_sede, codigo_equipo_grabacion y ubicacion_fisica son requeridos.", {})
            camara = self.querys.cctv_crear_camara(data)
            return self.tools.output(201, "Cámara creada.", camara)
        except Exception as e:
            print(f"[Cctv] crear_camara: {e}")
            return self.tools.output(500, "Error creando cámara.", {})

    def actualizar_camara(self, id_camara: int, data: dict):
        try:
            camara = self.querys.cctv_actualizar_camara(id_camara, data)
            if not camara:
                return self.tools.output(404, "Cámara no encontrada.", {})
            return self.tools.output(200, "Cámara actualizada.", camara)
        except Exception as e:
            print(f"[Cctv] actualizar_camara: {e}")
            return self.tools.output(500, "Error actualizando cámara.", {})

    def eliminar_camara(self, id_camara: int):
        try:
            self.querys.cctv_eliminar_camara(id_camara)
            return self.tools.output(200, "Cámara eliminada.", {})
        except Exception as e:
            print(f"[Cctv] eliminar_camara: {e}")
            return self.tools.output(500, "Error eliminando cámara.", {})

    # ── Cargos / permisos ──────────────────────────────────────────────────────

    def listar_cargos(self):
        try:
            data = self.querys.cctv_listar_cargos()
            return self.tools.output(200, "Cargos obtenidos.", data)
        except Exception as e:
            print(f"[Cctv] listar_cargos: {e}")
            return self.tools.output(500, "Error listando cargos.", {})

    def crear_cargo(self, data: dict):
        try:
            if not data.get('nombre'):
                return self.tools.output(400, "El campo nombre es requerido.", {})
            cargo = self.querys.cctv_crear_cargo(data)
            return self.tools.output(201, "Cargo creado.", cargo)
        except Exception as e:
            print(f"[Cctv] crear_cargo: {e}")
            return self.tools.output(500, "Error creando cargo.", {})

    def actualizar_cargo(self, id_cargo: int, data: dict):
        try:
            cargo = self.querys.cctv_actualizar_cargo(id_cargo, data)
            if not cargo:
                return self.tools.output(404, "Cargo no encontrado.", {})
            return self.tools.output(200, "Cargo actualizado.", cargo)
        except Exception as e:
            print(f"[Cctv] actualizar_cargo: {e}")
            return self.tools.output(500, "Error actualizando cargo.", {})

    def eliminar_cargo(self, id_cargo: int):
        try:
            self.querys.cctv_eliminar_cargo(id_cargo)
            return self.tools.output(200, "Cargo eliminado.", {})
        except Exception as e:
            print(f"[Cctv] eliminar_cargo: {e}")
            return self.tools.output(500, "Error eliminando cargo.", {})

    # ── Registros de cambios ───────────────────────────────────────────────────

    def listar_cambios(self):
        try:
            data = self.querys.cctv_listar_cambios()
            return self.tools.output(200, "Cambios obtenidos.", data)
        except Exception as e:
            print(f"[Cctv] listar_cambios: {e}")
            return self.tools.output(500, "Error listando cambios.", {})

    def crear_cambio(self, data: dict):
        try:
            if not data.get('descripcion') or not data.get('fecha_cambio'):
                return self.tools.output(400, "fecha_cambio y descripcion son requeridos.", {})
            cambio = self.querys.cctv_crear_cambio(data)
            return self.tools.output(201, "Cambio registrado.", cambio)
        except Exception as e:
            print(f"[Cctv] crear_cambio: {e}")
            return self.tools.output(500, "Error registrando cambio.", {})

    # ── Revisiones ─────────────────────────────────────────────────────────────

    def listar_revisiones(self):
        try:
            data = self.querys.cctv_listar_revisiones()
            return self.tools.output(200, "Revisiones obtenidas.", data)
        except Exception as e:
            print(f"[Cctv] listar_revisiones: {e}")
            return self.tools.output(500, "Error listando revisiones.", {})

    def crear_revision(self, data: dict):
        try:
            if not data.get('fecha_revision'):
                return self.tools.output(400, "fecha_revision es requerido.", {})
            revision = self.querys.cctv_crear_revision(data)
            return self.tools.output(201, "Revisión registrada.", revision)
        except Exception as e:
            print(f"[Cctv] crear_revision: {e}")
            return self.tools.output(500, "Error registrando revisión.", {})

    # ── Incidentes ─────────────────────────────────────────────────────────────

    def listar_incidentes(self):
        try:
            data = self.querys.cctv_listar_incidentes()
            return self.tools.output(200, "Incidentes obtenidos.", data)
        except Exception as e:
            print(f"[Cctv] listar_incidentes: {e}")
            return self.tools.output(500, "Error listando incidentes.", {})

    def crear_incidente(self, data: dict):
        try:
            if not data.get('titulo') or not data.get('descripcion') or not data.get('fecha_incidente'):
                return self.tools.output(400, "fecha_incidente, titulo y descripcion son requeridos.", {})
            incidente = self.querys.cctv_crear_incidente(data)
            return self.tools.output(201, "Incidente registrado.", incidente)
        except Exception as e:
            print(f"[Cctv] crear_incidente: {e}")
            return self.tools.output(500, "Error registrando incidente.", {})

    def actualizar_incidente(self, id_incidente: int, data: dict):
        try:
            incidente = self.querys.cctv_actualizar_incidente(id_incidente, data)
            if not incidente:
                return self.tools.output(404, "Incidente no encontrado.", {})
            return self.tools.output(200, "Incidente actualizado.", incidente)
        except Exception as e:
            print(f"[Cctv] actualizar_incidente: {e}")
            return self.tools.output(500, "Error actualizando incidente.", {})
