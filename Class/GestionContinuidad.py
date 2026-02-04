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