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
