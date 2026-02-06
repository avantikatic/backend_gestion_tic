from Utils.tools import Tools
from Utils.querys import Querys

class GestionContinuidad:

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.querys = Querys(self.db)

    def obtener_estados_gsc(self):
        """
        Obtiene todos los estados disponibles para el módulo de Gestión de Seguridad y Continuidad
        """
        try:
            estados = self.querys.obtener_estados_gsc()
            
            return self.tools.output(200, "Estados obtenidos exitosamente.", estados)
                
        except Exception as e:
            print(f"Error obteniendo estados GSC: {e}")
            return self.tools.output(500, "Error obteniendo estados.", {})

    def obtener_sistemas_afectados_gsc(self):
        """
        Obtiene todos los sistemas afectados disponibles para el módulo de Gestión de Seguridad y Continuidad
        """
        try:
            sistemas = self.querys.obtener_sistemas_afectados_gsc()
            
            return self.tools.output(200, "Sistemas afectados obtenidos exitosamente.", sistemas)
                
        except Exception as e:
            print(f"Error obteniendo sistemas afectados GSC: {e}")
            return self.tools.output(500, "Error obteniendo sistemas afectados.", {})

    def obtener_modulos_gsc(self):
        """
        Obtiene todos los módulos disponibles para el módulo de Gestión de Seguridad y Continuidad
        """
        try:
            modulos = self.querys.obtener_modulos_gsc()
            
            return self.tools.output(200, "Módulos obtenidos exitosamente.", modulos)
                
        except Exception as e:
            print(f"Error obteniendo módulos GSC: {e}")
            return self.tools.output(500, "Error obteniendo módulos.", {})

    def obtener_tipos_evidencia_gsc(self):
        """
        Obtiene todos los tipos de evidencia disponibles para el módulo GSC
        """
        try:
            tipos = self.querys.obtener_tipos_evidencia_gsc()
            
            return self.tools.output(200, "Tipos de evidencia obtenidos exitosamente.", tipos)
                
        except Exception as e:
            print(f"Error obteniendo tipos de evidencia GSC: {e}")
            return self.tools.output(500, "Error obteniendo tipos de evidencia.", {})

    def obtener_origenes_plataforma_gsc(self):
        """
        Obtiene todos los orígenes de plataforma disponibles para alertas en el módulo GSC
        """
        try:
            origenes = self.querys.obtener_origenes_plataforma_gsc()
            
            return self.tools.output(200, "Orígenes de plataforma obtenidos exitosamente.", origenes)
                
        except Exception as e:
            print(f"Error obteniendo orígenes de plataforma GSC: {e}")
            return self.tools.output(500, "Error obteniendo orígenes de plataforma.", {})

    def obtener_fuentes_seguridad_gsc(self):
        """
        Obtiene todas las fuentes de seguridad disponibles para el módulo SEG
        """
        try:
            fuentes = self.querys.obtener_fuentes_seguridad_gsc()
            
            return self.tools.output(200, "Fuentes de seguridad obtenidas exitosamente.", fuentes)
                
        except Exception as e:
            print(f"Error obteniendo fuentes de seguridad GSC: {e}")
            return self.tools.output(500, "Error obteniendo fuentes de seguridad.", {})

    def obtener_impactos_gsc(self):
        """
        Obtiene todos los niveles de impacto disponibles para el módulo SEG
        """
        try:
            impactos = self.querys.obtener_impactos_gsc()
            
            return self.tools.output(200, "Niveles de impacto obtenidos exitosamente.", impactos)
                
        except Exception as e:
            print(f"Error obteniendo impactos GSC: {e}")
            return self.tools.output(500, "Error obteniendo niveles de impacto.", {})

    def obtener_riesgos_gsc(self):
        """
        Obtiene todos los niveles de riesgo disponibles para el módulo MNT
        """
        try:
            riesgos = self.querys.obtener_riesgos_gsc()
            
            return self.tools.output(200, "Niveles de riesgo obtenidos exitosamente.", riesgos)
                
        except Exception as e:
            print(f"Error obteniendo riesgos GSC: {e}")
            return self.tools.output(500, "Error obteniendo niveles de riesgo.", {})

    def crear_registro_gsc(self, data: dict):
        """
        Crea un registro GSC completo con todas sus secciones
        """
        try:
            resultado = self.querys.crear_registro_gsc_completo(data)
            
            if resultado['success']:
                return self.tools.output(201, resultado['message'], {'id_registro': resultado['id_registro']})
            else:
                return self.tools.output(400, resultado['message'], {})
                
        except Exception as e:
            print(f"Error creando registro GSC: {e}")
            import traceback
            traceback.print_exc()
            return self.tools.output(500, "Error creando registro GSC.", {})

    def obtener_registro_gsc(self, id_registro: int):
        """
        Obtiene un registro GSC completo por su ID
        """
        try:
            registro = self.querys.obtener_registro_gsc_completo(id_registro)
            
            if registro:
                return self.tools.output(200, "Registro obtenido exitosamente.", registro)
            else:
                return self.tools.output(404, "Registro no encontrado.", {})
                
        except Exception as e:
            print(f"Error obteniendo registro GSC: {e}")
            return self.tools.output(500, "Error obteniendo registro.", {})

    def listar_registros_gsc(self, filtros: dict = None):
        """
        Lista registros GSC con filtros opcionales
        """
        try:
            registros = self.querys.listar_registros_gsc(filtros)
            
            return self.tools.output(200, "Registros obtenidos exitosamente.", registros)
                
        except Exception as e:
            print(f"Error listando registros GSC: {e}")
            return self.tools.output(500, "Error listando registros.", {})

    def obtener_contadores_gsc(self, filtros: dict = None):
        """
        Obtiene contadores de registros por estado para un módulo específico
        """
        try:
            contadores = self.querys.obtener_contadores_gsc(filtros)
            
            return self.tools.output(200, "Contadores obtenidos exitosamente.", contadores)
                
        except Exception as e:
            print(f"Error obteniendo contadores GSC: {e}")
            return self.tools.output(500, "Error obteniendo contadores.", {})

    def actualizar_registro_gsc(self, id_registro: int, data: dict):
        """
        Actualiza un registro GSC existente
        """
        try:
            resultado = self.querys.actualizar_registro_gsc_completo(id_registro, data)
            
            if resultado['success']:
                return self.tools.output(200, resultado['message'], {})
            else:
                return self.tools.output(400, resultado['message'], {})
                
        except Exception as e:
            print(f"Error actualizando registro GSC: {e}")
            return self.tools.output(500, "Error actualizando registro.", {})

    def eliminar_registro_gsc(self, id_registro: int):
        """
        Elimina (desactiva) un registro GSC
        """
        try:
            resultado = self.querys.eliminar_registro_gsc(id_registro)
            
            if resultado['success']:
                return self.tools.output(200, resultado['message'], {})
            else:
                return self.tools.output(404, resultado['message'], {})
                
        except Exception as e:
            print(f"Error eliminando registro GSC: {e}")
            return self.tools.output(500, "Error eliminando registro.", {})
