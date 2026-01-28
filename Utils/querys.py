from Utils.tools import Tools, CustomException
from sqlalchemy import text, func, case, extract, and_, or_, Date, cast
from datetime import datetime, date
import json
import traceback
from Models.IntranetGraphTokenModel import IntranetGraphTokenModel as TokenModel
from Models.IntranetCorreosMicrosoftModel import IntranetCorreosMicrosoftModel as CorreosMicrosoftModel
from Models.IntranetSyncLogModel import IntranetSyncLogModel as SyncLogModel
from Models.IntranetEstadosTickets import IntranetEstadosTickets
from Models.IntranetUsuariosGestionTicModel import IntranetUsuariosGestionTicModel
from Models.IntranetTipoPrioridadModel import IntranetTipoPrioridadModel
from Models.IntranetTipoSoporteModel import IntranetTipoSoporteModel
from Models.IntranetTipoTicketModel import IntranetTipoTicketModel
from Models.IntranetPerfilesMacroprocesoModel import IntranetPerfilesMacroprocesoModel
from Models.IntranetTipoNivelModel import IntranetTipoNivelModel
from Models.IntranetObservacionesInformeGestionModel import IntranetObservacionesInformeGestionModel
from Models.IntranetCausasInformeGestionModel import IntranetCausasInformeGestion
from Models.IntranetAniosInformeGestionModel import IntranetAniosInformeGestion
from Models.IntranetOrigenEstrategicoModel import IntranetOrigenEstrategicoModel
from Models.IntranetTiposServicioModel import IntranetTiposServicioModel
from Models.IntranetProveedoresModel import IntranetProveedoresModel
from Models.IntranetProductosServiciosModel import IntranetProductosServiciosModel
from Models.IntranetMetodosPagoModel import IntranetMetodosPagoModel
from Models.IntranetLicenciasModel import IntranetLicenciasModel
from Models.IntranetLicenciasHistorialModel import IntranetLicenciasHistorialModel
from Models.IntranetTipoRevisionModel import IntranetTipoRevisionModel
from Models.IntranetRevisionesModel import IntranetRevisionesModel
from Models.IntranetVersionesLicenciasModel import IntranetVersionesLicenciasModel

import hashlib

class Querys:

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.query_params = dict()

    # Query para obtener la información del activo por código
    def get_token(self):
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                response = dict()
                
                sql = self.db.query(
                    TokenModel
                ).filter(
                    TokenModel.estado == 1
                ).order_by(
                    TokenModel.id.desc()
                ).first()

                if sql:
                    response = sql.to_dict()

                return response

            except Exception as e:
                retry_count += 1
                print(f"Error en conexión a BD (intento {retry_count}/{max_retries}): {e}")
                
                if retry_count < max_retries:
                    # Cerrar conexión actual e intentar reconectar
                    try:
                        self.db.close()
                    except:
                        pass
                    
                    # Esperar un poco antes del siguiente intento
                    import time
                    time.sleep(1)
                    
                    # Reinicializar la conexión
                    from Config.db import get_database
                    self.db = next(get_database())
                else:
                    # Si se agotaron los reintentos, lanzar excepción
                    raise CustomException(f"Error de conexión a BD después de {max_retries} intentos: {e}")
        
        return dict()

    # Query para desactivar token expirado
    def desactivar_token(self, token_id: int):
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                token_record = self.db.query(
                    TokenModel).filter(TokenModel.id == token_id).first()
                if token_record:
                    token_record.estado = 0
                    self.db.commit()
                    return True
                return False
                
            except Exception as e:
                retry_count += 1
                print(f"Error desactivando token (intento {retry_count}/{max_retries}): {e}")
                
                try:
                    self.db.rollback()
                except:
                    pass
                
                if retry_count < max_retries:
                    # Cerrar conexión actual e intentar reconectar
                    try:
                        self.db.close()
                    except:
                        pass
                    
                    # Esperar un poco antes del siguiente intento
                    import time
                    time.sleep(1)
                    
                    # Reinicializar la conexión
                    from Config.db import get_database
                    self.db = next(get_database())
                else:
                    raise CustomException(f"Error desactivando token después de {max_retries} intentos: {e}")
        
        return False

    # Query para insertar datos en cualquier tabla
    def insertar_datos(self, model: any, data: dict):
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                new_record = model(data)
                self.db.add(new_record)
                self.db.commit()
                self.db.refresh(new_record)
                return new_record
                
            except Exception as e:
                retry_count += 1
                print(f"Error insertando datos (intento {retry_count}/{max_retries}): {e}")
                
                try:
                    self.db.rollback()
                except:
                    pass
                
                if retry_count < max_retries:
                    # Cerrar conexión actual e intentar reconectar
                    try:
                        self.db.close()
                    except:
                        pass
                    
                    # Esperar un poco antes del siguiente intento
                    import time
                    time.sleep(1)
                    
                    # Reinicializar la conexión
                    from Config.db import get_database
                    self.db = next(get_database())
                else:
                    raise CustomException(f"Error insertando datos después de {max_retries} intentos: {e}")
        
        return None

    # ============= MÉTODOS PARA CORREOS MICROSOFT =============

    # Query para generar hash del contenido del correo
    def generar_hash_contenido(self, subject, body_preview, from_email):
        """Genera un hash del contenido del correo para detectar cambios"""
        contenido = f"{subject}{body_preview}{from_email}"
        return hashlib.sha256(contenido.encode()).hexdigest()
    
    # Query para obtener un correo por su message_id de Microsoft
    def obtener_correo_por_message_id(self, message_id):
        """Obtiene un correo por su message_id de Microsoft"""
        try:
            correo = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.message_id == message_id
            ).first()
            
            return correo.to_dict() if correo else None
            
        except Exception as e:
            print(f"Error obteniendo correo por message_id: {e}")
            return None

    # Query para obtener correos desde la base de datos con filtros y paginación
    def obtener_correos_bd(self, limite=100, offset=0, estado=None):
        """Obtiene correos desde la base de datos con filtros y paginación"""
        try:
            # Filtrar correos activos y no descartados (estado != 0)
            query = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.ticket == 0,
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.estado != 0  # Excluir correos descartados
            )
            
            # Filtro por estado específico si se especifica
            if estado:
                query = query.filter(CorreosMicrosoftModel.estado == estado)
            
            # Ordenar por fecha recibida (más recientes primero)
            query = query.order_by(CorreosMicrosoftModel.received_date.desc())
            
            # Paginación
            correos = query.offset(offset).limit(limite).all()
            
            # Convertir a formato frontend
            return [correo.to_frontend_format() for correo in correos]
            
        except Exception as e:
            print(f"Error obteniendo correos de BD: {e}")
            return []

    # Query para insertar un nuevo correo
    def insertar_correo(self, correo_data):
        """Inserta un nuevo correo en la base de datos"""
        try:
            # Generar hash del contenido
            hash_contenido = self.generar_hash_contenido(
                correo_data.get('subject', ''),
                correo_data.get('body_preview', ''),
                correo_data.get('from_email', '')
            )
            correo_data['hash_contenido'] = hash_contenido
            
            nuevo_correo = CorreosMicrosoftModel(correo_data)
            self.db.add(nuevo_correo)
            self.db.commit()
            self.db.refresh(nuevo_correo)
            
            return nuevo_correo.to_dict()
            
        except Exception as e:
            self.db.rollback()
            print(f"Error insertando correo: {e}")
            return None
    
    # Query para actualizar un correo existente
    def actualizar_correo(self, message_id, datos_actualizacion):
        """Actualiza un correo existente"""
        try:
            correo = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.message_id == message_id
            ).first()
            
            if correo:
                # Actualizar campos
                for campo, valor in datos_actualizacion.items():
                    if hasattr(correo, campo):
                        setattr(correo, campo, valor)
                
                correo.updated_at = datetime.now()
                self.db.commit()
                return correo.to_dict()
            
            return None
            
        except Exception as e:
            self.db.rollback()
            print(f"Error actualizando correo: {e}")
            return None
    
    # Query para obtener todos los message_ids existentes en BD
    def obtener_message_ids_existentes(self):
        """Obtiene todos los message_ids existentes en BD"""
        try:
            result = self.db.query(CorreosMicrosoftModel.message_id).all()
            return {row[0] for row in result}  # Set para búsqueda rápida
            
        except Exception as e:
            print(f"Error obteniendo message_ids existentes: {e}")
            return set()
    
    # Query para marcar un correo como procesado o cambiar su estado
    def marcar_correo_procesado(self, message_id, nuevo_estado='procesado'):
        """Marca un correo como procesado o cambia su estado"""
        return self.actualizar_correo(message_id, {'estado': nuevo_estado})
    
    # Query para obtener un ticket por su ticket_id (id de la tabla)
    def obtener_ticket_por_id(self, ticket_id):
        """
        Obtiene un ticket por su ticket_id (id de la tabla)
        """
        try:
            ticket = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.id == ticket_id
            ).first()
            
            if ticket:
                return ticket.to_dict()
            
            return None
            
        except Exception as e:
            print(f"Error obteniendo ticket por ID {ticket_id}: {e}")
            return None

    # Query para obtener un correo por su message_id
    def obtener_correo_por_message_id(self, message_id):
        """
        Obtiene un correo por su message_id
        """
        try:
            correo = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.message_id == message_id
            ).first()
            
            if correo:
                return correo.to_dict()
            
            return None
            
        except Exception as e:
            print(f"Error obteniendo correo por message_id {message_id}: {e}")
            return None

    # Query para registrar una respuesta enviada a un correo
    def registrar_respuesta_correo(self, message_id, respuesta, ticket_id=None):
        """
        Registra una respuesta enviada a un correo en la base de datos
        """
        try:
            # Por ahora solo actualizar el timestamp para indicar que se respondió
            datos_actualizacion = {
                'updated_at': datetime.now()
            }
            
            return self.actualizar_correo(message_id, datos_actualizacion)
            
        except Exception as e:
            print(f"Error registrando respuesta de correo: {e}")
            return None
    
    # Query para descartar un correo (marcarlo como inactivo)
    def descartar_correo(self, message_id):
        """Marca un correo como descartado (estado = 0) para que no aparezca en la bandeja"""
        try:
            resultado = self.actualizar_correo(message_id, {
                'activo': 0,
                'fecha_actualizacion': datetime.now()
            })
            
            if resultado:
                print(f"Correo {message_id} marcado como descartado")
                return resultado
            else:
                print(f"No se encontró el correo {message_id} para descartar")
                return None
                
        except Exception as e:
            print(f"Error descartando correo {message_id}: {e}")
            return None
    
    # Query para convertir un correo en ticket
    def convertir_correo_ticket(self, message_id):
        """Marca un correo como convertido a ticket (ticket = 1) y genera ticket_id"""
        try:
            resultado = self.actualizar_correo(message_id, {
                'ticket': 1,
                'fecha_actualizacion': datetime.now()
            })
            
            if resultado:
                # Generar ticket_id en formato TCK-XXXX basado en el ID del correo
                ticket_id_display = f"TCK-{resultado['id']:04d}"
                ticket_id_numero = str(resultado['id'])  # Solo el número puro
                
                print(f"Correo {message_id} convertido a ticket {ticket_id_display}")
                
                # Agregar ambos formatos al resultado
                resultado['ticket_id'] = ticket_id_numero  # Para uso interno
                resultado['ticket_id_display'] = ticket_id_display  # Para mostrar al usuario
                
                return resultado
            else:
                print(f"No se encontró el correo {message_id} para convertir")
                return None
                
        except Exception as e:
            print(f"Error convirtiendo correo {message_id} a ticket: {e}")
            return None
    
    # Query para obtener correos convertidos en tickets con filtrado optimizado por vista
    def obtener_tickets_correos(self, vista=None, limite=100, offset=0, tecnico_id=None):
        """
        Obtiene correos convertidos en tickets desde la base de datos
        Filtrado optimizado por vista para máximo rendimiento
        Incluye JOIN con IntranetEstadosTickets para obtener el nombre del estado
        """
        try:
            # Query base con JOINs: correos activos convertidos a tickets + información completa
            query = self.db.query(
                CorreosMicrosoftModel,
                IntranetEstadosTickets.nombre.label('estado_nombre'),
                IntranetUsuariosGestionTicModel.nombre.label('tecnico_nombre'),
                IntranetTipoPrioridadModel.nombre.label('prioridad_nombre'),
                IntranetTipoSoporteModel.nombre.label('tipo_soporte_nombre'),
                IntranetTipoTicketModel.nombre.label('tipo_ticket_nombre'),
                IntranetPerfilesMacroprocesoModel.nombre.label('macroproceso_nombre')
            ).outerjoin(
                IntranetEstadosTickets, 
                CorreosMicrosoftModel.estado == IntranetEstadosTickets.id
            ).outerjoin(
                IntranetUsuariosGestionTicModel,
                CorreosMicrosoftModel.asignado == IntranetUsuariosGestionTicModel.id
            ).outerjoin(
                IntranetTipoPrioridadModel,
                CorreosMicrosoftModel.prioridad == IntranetTipoPrioridadModel.id
            ).outerjoin(
                IntranetTipoSoporteModel,
                CorreosMicrosoftModel.tipo_soporte == IntranetTipoSoporteModel.id
            ).outerjoin(
                IntranetTipoTicketModel,
                CorreosMicrosoftModel.tipo_ticket == IntranetTipoTicketModel.id
            ).outerjoin(
                IntranetPerfilesMacroprocesoModel,
                CorreosMicrosoftModel.macroproceso == IntranetPerfilesMacroprocesoModel.id
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1
            )
            
            # Aplicar filtros específicos por vista
            if vista == 'todos':
                # Ya tenemos el filtro base
                pass
            elif vista == 'sin':
                # Sin asignar: donde asignado es NULL o vacío
                query = query.filter(
                    CorreosMicrosoftModel.asignado.is_(None)
                )
            elif vista == 'abiertos':
                # Estado = 1 (Abierto)
                query = query.filter(CorreosMicrosoftModel.estado == 1)
            elif vista == 'proceso':
                # Estado = 2 (En Proceso)
                query = query.filter(CorreosMicrosoftModel.estado == 2)
            elif vista == 'comp':
                # Estado = 4 (Completado)
                query = query.filter(CorreosMicrosoftModel.estado == 3)
            elif vista and vista.startswith('tecnico_'):
                # Filtro por técnico específico: tecnico_1, tecnico_2, etc.
                tecnico_id_from_vista = vista.replace('tecnico_', '')
                try:
                    tecnico_id_int = int(tecnico_id_from_vista)
                    query = query.filter(CorreosMicrosoftModel.asignado == tecnico_id_int)
                except ValueError:
                    # Si no es un número válido, no aplicar filtro
                    pass
            
            # Filtro adicional por tecnico_id específico (parámetro directo)
            if tecnico_id:
                query = query.filter(CorreosMicrosoftModel.asignado == tecnico_id)
            
            # Ordenar por fecha recibida (más recientes primero)
            query = query.order_by(CorreosMicrosoftModel.received_date.desc())
            
            # Obtener total para paginación (sin JOIN para mejor performance en count)
            count_query = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1
            )
            
            # Aplicar los mismos filtros para el conteo
            if vista == 'sin':
                count_query = count_query.filter(CorreosMicrosoftModel.asignado.is_(None))
            elif vista == 'abiertos':
                count_query = count_query.filter(CorreosMicrosoftModel.estado == 1)
            elif vista == 'proceso':
                count_query = count_query.filter(CorreosMicrosoftModel.estado == 2)
            elif vista == 'comp':
                count_query = count_query.filter(CorreosMicrosoftModel.estado == 3)
            elif vista and vista.startswith('tecnico_'):
                # Aplicar el mismo filtro de técnico para el conteo
                tecnico_id_from_vista = vista.replace('tecnico_', '')
                try:
                    tecnico_id_int = int(tecnico_id_from_vista)
                    count_query = count_query.filter(CorreosMicrosoftModel.asignado == tecnico_id_int)
                except ValueError:
                    pass
                    
            # Filtro adicional por tecnico_id para el conteo
            if tecnico_id:
                count_query = count_query.filter(CorreosMicrosoftModel.asignado == tecnico_id)
            
            total = count_query.count()
            
            # Aplicar paginación y obtener resultados
            resultados = query.offset(offset).limit(limite).all()
            
            # Convertir a formato frontend con información adicional de todos los JOINs
            tickets = []
            for correo, estado_nombre, tecnico_nombre, prioridad_nombre, tipo_soporte_nombre, tipo_ticket_nombre, macroproceso_nombre in resultados:
                ticket_data = correo.to_frontend_format()
                # Agregar información del estado
                ticket_data['estado_nombre'] = estado_nombre or '-'
                ticket_data['estadoTicket'] = estado_nombre or '-'  # Para compatibilidad
                # Agregar información del técnico asignado
                ticket_data['tecnico_nombre'] = tecnico_nombre or '-'
                ticket_data['asignadoNombre'] = tecnico_nombre or '-'  # Para compatibilidad
                # Agregar información de prioridad
                ticket_data['prioridad_nombre'] = prioridad_nombre or '-'
                # Agregar información de tipo de soporte
                ticket_data['tipo_soporte_nombre'] = tipo_soporte_nombre or '-'
                # Agregar información de tipo de ticket
                ticket_data['tipo_ticket_nombre'] = tipo_ticket_nombre or '-'
                # Agregar información de macroproceso
                ticket_data['macroproceso_nombre'] = macroproceso_nombre or '-'
                tickets.append(ticket_data)
            
            return {
                'tickets': tickets,
                'total': total,
                'limite': limite,
                'offset': offset,
                'vista': vista
            }
            
        except Exception as e:
            print(f"Error obteniendo tickets de correos: {e}")
            return {
                'tickets': [],
                'total': 0,
                'limite': limite,
                'offset': offset,
                'vista': vista
            }
    
    # Query para obtener todos los estados de tickets
    def obtener_estados_tickets(self):
        """
        Obtiene todos los estados de tickets disponibles desde IntranetEstadosTickets
        """
        try:
            estados = self.db.query(IntranetEstadosTickets).filter(
                IntranetEstadosTickets.estado == 1
            ).all()
            
            return [{'id': estado.id, 'nombre': estado.nombre} for estado in estados]
            
        except Exception as e:
            print(f"Error obteniendo estados de tickets: {e}")
            return []
    
    # Query para obtener todos los técnicos de gestión TIC
    def obtener_tecnicos_gestion_tic(self):
        """
        Obtiene todos los técnicos disponibles desde IntranetUsuariosGestionTicModel
        """
        try:
            tecnicos = self.db.query(IntranetUsuariosGestionTicModel).filter(
                IntranetUsuariosGestionTicModel.estado == 1
            ).all()
            
            return [{'id': tecnico.id, 'nombre': tecnico.nombre} for tecnico in tecnicos]
            
        except Exception as e:
            print(f"Error obteniendo técnicos de gestión TIC: {e}")
            return []
    
    # Querys para logs de sincronización
    def obtener_ultimo_sync_exitoso(self):
        """Obtiene información del último sync exitoso"""
        try:
            ultimo_sync = self.db.query(SyncLogModel).filter(
                SyncLogModel.estado == 'exitoso'
            ).order_by(SyncLogModel.fecha_fin.desc()).first()
            
            return ultimo_sync.to_dict() if ultimo_sync else None
            
        except Exception as e:
            print(f"Error obteniendo último sync: {e}")
            return None
    
    # Query para crear un nuevo log de sincronización
    def crear_log_sync(self, tipo_sync='incremental'):
        """Crea un nuevo registro de sincronización"""
        try:
            log_data = {
                'tipo_sync': tipo_sync,
                'fecha_inicio': datetime.now(),
                'estado': 1
            }
            
            nuevo_log = SyncLogModel(log_data)
            self.db.add(nuevo_log)
            self.db.commit()
            self.db.refresh(nuevo_log)
            
            return nuevo_log.id
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando log de sync: {e}")
            return None
    
    # Query para finalizar un log de sincronización
    def finalizar_log_sync(self, log_id, correos_nuevos=0, correos_actualizados=0, 
                          correos_eliminados=0, estado=1, mensaje_error=None):
        """Finaliza un log de sincronización"""
        try:
            log_sync = self.db.query(SyncLogModel).filter(
                SyncLogModel.id == log_id
            ).first()
            
            if log_sync:
                log_sync.fecha_fin = datetime.now()
                log_sync.correos_nuevos = correos_nuevos
                log_sync.correos_actualizados = correos_actualizados
                log_sync.correos_eliminados = correos_eliminados
                log_sync.estado = estado
                log_sync.mensaje_error = mensaje_error
                
                self.db.commit()
                return log_sync.to_dict()
            
            return None
            
        except Exception as e:
            self.db.rollback()
            print(f"Error finalizando log de sync: {e}")
            return None

    # Querys para obtener listas de prioridades, tipos de soporte, tipos de ticket y macroprocesos
    def obtener_prioridades(self):
        """
        Obtiene todas las prioridades disponibles desde IntranetPrioridades
        """
        try:
            prioridades = self.db.query(IntranetTipoPrioridadModel).filter(
                IntranetTipoPrioridadModel.estado == 1
            ).all()
            
            return [{'id': prioridad.id, 'nombre': prioridad.nombre} for prioridad in prioridades]
            
        except Exception as e:
            print(f"Error obteniendo prioridades: {e}")
            return []

    # Query para obtener tipos de soporte
    def obtener_tipo_soporte(self):
        """
        Obtiene todos los tipos de soporte disponibles desde IntranetTipoSoporte
        """
        try:
            tipos_soporte = self.db.query(IntranetTipoSoporteModel).filter(
                IntranetTipoSoporteModel.estado == 1
            ).all()
            
            return [{'id': tipo.id, 'nombre': tipo.nombre} for tipo in tipos_soporte]
            
        except Exception as e:
            print(f"Error obteniendo tipos de soporte: {e}")
            return []

    # Query para obtener tipos de ticket
    def obtener_tipo_ticket(self):
        """
        Obtiene todos los tipos de ticket disponibles desde IntranetTipoTicket
        """
        try:
            tipos_ticket = self.db.query(IntranetTipoTicketModel).filter(
                IntranetTipoTicketModel.estado == 1
            ).all()
            
            return [{'id': tipo.id, 'nombre': tipo.nombre} for tipo in tipos_ticket]
        except Exception as e:
            print(f"Error obteniendo tipos de ticket: {e}")
            return []

    # Query para obtener macroprocesos
    def obtener_macroprocesos(self):
        """
        Obtiene todos los macroprocesos disponibles (valores estáticos por ahora)
        """
        try:
            # Valores estáticos por ahora
            macroprocesos = self.db.query(IntranetPerfilesMacroprocesoModel).filter(
                IntranetPerfilesMacroprocesoModel.estado == 1
            ).all()
            return [{'id': macro.id, 'nombre': macro.nombre} for macro in macroprocesos]

        except Exception as e:
            print(f"Error obteniendo macroprocesos: {e}")
            return self.tools.output(500, "Error obteniendo macroprocesos.", {})

    # Query para obtener tipos de nivel
    def obtener_tipo_nivel(self):
        """
        Obtiene todos los tipos de nivel disponibles desde IntranetTipoNivel
        """
        try:
            tipos_nivel = self.db.query(IntranetTipoNivelModel).filter(
                IntranetTipoNivelModel.estado == 1
            ).all()
            
            return [{'id': tipo.id, 'nombre': tipo.nombre} for tipo in tipos_nivel]
            
        except Exception as e:
            print(f"Error obteniendo tipos de nivel: {e}")
            return []

    # Query para obtener orígenes estratégicos
    def obtener_origen_estrategico(self):
        """
        Obtiene todos los orígenes estratégicos disponibles
        """
        try:
            origenes = self.db.query(IntranetOrigenEstrategicoModel).filter(
                IntranetOrigenEstrategicoModel.estado == 1
            ).all()
            
            return [{'id': origen.id, 'nombre': origen.nombre} for origen in origenes]
            
        except Exception as e:
            print(f"Error obteniendo orígenes estratégicos: {e}")
            return []

    # Query para filtrar tickets con optimización usando IDs exactos
    def filtrar_tickets_optimizado(self, filtros: dict):
        """
        Filtra tickets usando los campos reales de la tabla intranet_correos_microsoft
        Optimizado para usar los IDs exactos que envía el frontend
        
        Campos de la tabla:
        - ticket (0/1) - Solo tickets convertidos
        - asignado (ID del técnico)
        - prioridad (ID de prioridad)
        - tipo_soporte (ID de tipo soporte)
        - tipo_ticket (ID de tipo ticket) 
        - macroproceso (ID de macroproceso)
        - estado (ID numérico del estado)
        """
        try:
            # Query base con JOINs para obtener nombres
            base_query = """
            SELECT DISTINCT
                icm.id,
                icm.message_id,
                icm.subject,
                icm.from_name,
                icm.from_email,
                icm.body_content,
                icm.received_date,
                icm.created_at,
                icm.updated_at,
                icm.ticket,
                icm.estado,
                icm.asignado,
                icm.prioridad,
                icm.tipo_soporte,
                icm.tipo_ticket,
                icm.origen_estrategico,
                icm.macroproceso,
                icm.fecha_vencimiento,
                icm.sla,
                icm.nivel_id,
                icm.fecha_cierre,
                
                -- JOINs para obtener nombres
                itp.nombre as prioridad_nombre,
                its.nombre as tipo_soporte_nombre,
                itt.nombre as tipo_ticket_nombre,
                ipm.nombre as macroproceso_nombre,
                iugt.nombre as asignado_nombre,
                ioe.nombre as origen_estrategico_nombre,
                -- Mapeo de estados
                CASE 
                    WHEN icm.estado = 1 THEN 'Abierto'
                    WHEN icm.estado = 2 THEN 'En Proceso' 
                    WHEN icm.estado = 3 THEN 'Completado'
                    WHEN icm.estado = 4 THEN 'Cerrado'
                    ELSE 'Abierto'
                END as estado_nombre

            FROM intranet_correos_microsoft icm
            
            -- LEFT JOINs para obtener nombres
            LEFT JOIN intranet_tipo_prioridad itp ON icm.prioridad = itp.id AND itp.estado = 1
            LEFT JOIN intranet_tipo_soporte its ON icm.tipo_soporte = its.id AND its.estado = 1  
            LEFT JOIN intranet_tipo_ticket itt ON icm.tipo_ticket = itt.id AND itt.estado = 1
            LEFT JOIN intranet_perfiles_macroproceso ipm ON icm.macroproceso = ipm.id AND ipm.estado = 1
            LEFT JOIN intranet_usuarios_gestion_tic iugt ON icm.asignado = iugt.id AND iugt.estado = 1
            LEFT JOIN intranet_origen_estrategico ioe ON icm.origen_estrategico = ioe.id AND ioe.estado = 1             
            WHERE icm.activo = 1 
            AND icm.ticket = 1
            """
            
            params = {}
            
            # 1. Filtro de vista base
            vista = filtros.get('vista', 'todos')
            if vista == 'sin':
                base_query += " AND icm.asignado IS NULL"
            elif vista == 'abiertos':
                base_query += " AND icm.estado = 1"
            elif vista == 'proceso':
                base_query += " AND icm.estado = 2"
            elif vista == 'comp':
                base_query += " AND icm.estado = 3"
            elif vista.startswith('tecnico_'):
                tecnico_id = int(vista.replace('tecnico_', ''))
                base_query += " AND icm.asignado = :tecnico_id"
                params['tecnico_id'] = tecnico_id
            
            # 2. Filtros específicos usando campos reales
            if filtros.get('q'):
                search_term = f"%{filtros['q']}%"
                base_query += """ AND (
                    CAST(icm.id AS NVARCHAR) LIKE :search_term OR
                    icm.subject LIKE :search_term OR  
                    icm.from_name LIKE :search_term OR
                    icm.from_email LIKE :search_term
                )"""
                params['search_term'] = search_term
                
            if filtros.get('estado'):
                base_query += " AND icm.estado = :estado_filtro"
                params['estado_filtro'] = filtros['estado']
                
            if filtros.get('asignado'):
                base_query += " AND icm.asignado = :asignado_filtro"
                params['asignado_filtro'] = filtros['asignado']
                
            if filtros.get('tipo_soporte'):
                base_query += " AND icm.tipo_soporte = :tipo_soporte_filtro"
                params['tipo_soporte_filtro'] = filtros['tipo_soporte']
                
            if filtros.get('macroproceso'):
                base_query += " AND icm.macroproceso = :macroproceso_filtro"
                params['macroproceso_filtro'] = filtros['macroproceso']
                
            if filtros.get('tipo_ticket'):
                base_query += " AND icm.tipo_ticket = :tipo_ticket_filtro"
                params['tipo_ticket_filtro'] = filtros['tipo_ticket']
            
            # 3. Contar total (antes de paginación)
            # Extraer solo la parte WHERE y filtros de la query base
            where_clause = ""
            if 'WHERE' in base_query:
                where_clause = "WHERE" + base_query.split('WHERE', 1)[1]
            
            count_query = f"""
            SELECT COUNT(DISTINCT icm.id) as total
            FROM intranet_correos_microsoft icm
            LEFT JOIN intranet_usuarios_gestion_tic iugt ON icm.asignado = iugt.id AND iugt.estado = 1
            {where_clause}
            """
            
            total_result = self.db.execute(text(count_query), params).fetchone()
            total = total_result[0] if total_result else 0
            
            # 4. Agregar ordenación y paginación (SQL Server syntax)
            limite = filtros.get('limite', 100)
            offset = filtros.get('offset', 0)
            
            base_query += f"""
            ORDER BY icm.created_at DESC
            OFFSET {offset} ROWS
            FETCH NEXT {limite} ROWS ONLY
            """
            
            # 5. Ejecutar query principal
            result = self.db.execute(text(base_query), params)
            tickets = []
            
            for row in result:
                # Convertir row a diccionario usando nombres de columnas
                row_dict = dict(row._mapping)
                
                ticket_dict = {
                    'id': row_dict.get('id'),
                    'message_id': row_dict.get('message_id'), 
                    'subject': row_dict.get('subject'),
                    'from_name': row_dict.get('from_name'),
                    'from_email': row_dict.get('from_email'),
                    'body': row_dict.get('body_content'),
                    'received_at': row_dict.get('received_date').strftime('%Y-%m-%d %H:%M:%S') if row_dict.get('received_date') else None,
                    'created_at': row_dict.get('created_at').strftime('%Y-%m-%d') if row_dict.get('created_at') else None,
                    'updated_at': row_dict.get('updated_at').strftime('%Y-%m-%d %H:%M:%S') if row_dict.get('updated_at') else None,
                    'ticket_id': row_dict.get('id'),
                    'ticket_id_display': f"TCK-{row_dict.get('id'):04d}",
                    'ticket': row_dict.get('ticket'),
                    'estado': row_dict.get('estado'),
                    'asignado': row_dict.get('asignado'),
                    'prioridad': row_dict.get('prioridad'),
                    'tipo_soporte': row_dict.get('tipo_soporte'),
                    'tipo_ticket': row_dict.get('tipo_ticket'),
                    'origen_estrategico': row_dict.get('origen_estrategico'),
                    'macroproceso': row_dict.get('macroproceso'),
                    'fecha_vencimiento': row_dict.get('fecha_vencimiento').strftime('%Y-%m-%d') if row_dict.get('fecha_vencimiento') else None,
                    'sla': row_dict.get('sla'),
                    'nivel_id': row_dict.get('nivel_id'),
                    'fecha_cierre': row_dict.get('fecha_cierre').strftime('%Y-%m-%d %H:%M:%S') if row_dict.get('fecha_cierre') else None,
                    'prioridad_nombre': row_dict.get('prioridad_nombre'),
                    'tipo_soporte_nombre': row_dict.get('tipo_soporte_nombre'),
                    'tipo_ticket_nombre': row_dict.get('tipo_ticket_nombre'),
                    'macroproceso_nombre': row_dict.get('macroproceso_nombre'),
                    'asignadoNombre': row_dict.get('asignado_nombre'),
                    'origen_estrategico_nombre': row_dict.get('origen_estrategico_nombre'),
                    'estadoTicket': row_dict.get('estado_nombre')
                }
                tickets.append(ticket_dict)
            
            # 6. Preparar respuesta
            return {
                'tickets': tickets,
                'total': total,
                'limite': filtros.get('limite', 100),
                'offset': filtros.get('offset', 0),
                'filtros_aplicados': {k: v for k, v in filtros.items() 
                                   if k not in ['limite', 'offset'] and v is not None}
            }
            
        except Exception as e:
            print(f"Error en filtrar_tickets_optimizado: {e}")
            raise e

    # ===== FUNCIONES PARA MANEJO DE RESPUESTAS EN HILOS =====
    
    # Query para obtener un ticket por su conversation_id
    def obtener_ticket_por_conversation_id(self, conversation_id):
        """
        Busca un ticket existente basado en el conversation_id
        Returns: dict con datos del ticket o None si no existe
        """
        try:
            if not conversation_id:
                return None
                
            # Por ahora buscaremos por similitud de asunto, ya que no tenemos conversation_id en BD
            # En el futuro se puede agregar la columna conversation_id
            print(f"🔍 Buscando ticket por conversation_id: {conversation_id}")
            
            # Buscaremos cualquier ticket existente (temporalmente deshabilitado)
            # TODO: Agregar columna conversation_id a la tabla
            return None
            
            result = self.db.execute(sql, {"conversation_id": conversation_id}).fetchone()
            
            if result:
                return {
                    'id': result[0],
                    'message_id': result[1],
                    'subject': result[2],
                    'from_email': result[3],
                    'from_name': result[4],
                    'conversation_id': result[5],
                    'created_at': result[6]
                }
                
            return None
            
        except Exception as e:
            print(f"Error en obtener_ticket_por_conversation_id: {e}")
            return None

    # Query para registrar una respuesta entrante en el historial del ticket
    def registrar_respuesta_entrante_ticket(self, respuesta_data):
        """
        Registra una respuesta entrante en el historial del ticket
        """
        try:
            # Por ahora, vamos a insertarlo como un correo normal pero marcado como respuesta
            # En el futuro se puede crear una tabla específica para respuestas
            
            # Crear entrada usando el constructor correcto del modelo
            correo_data = {
                'message_id': respuesta_data.get('message_id'),
                'subject': f"[RESPUESTA] {respuesta_data.get('subject', '')}",
                'from_email': respuesta_data.get('from_email'),
                'from_name': respuesta_data.get('from_name'),
                'received_date': respuesta_data.get('received_date'),
                'body_preview': respuesta_data.get('subject', '')[:100],
                'body_content': respuesta_data.get('body_content'),
                'estado': 2  # Estado 2 = Respuesta procesada (no aparece en buzón)
            }
            
            correo_respuesta = CorreosMicrosoftModel(correo_data)
            
            self.db.add(correo_respuesta)
            self.db.commit()
            
            print(f"✅ Respuesta registrada para ticket {respuesta_data.get('ticket_id')}")
            return True
            
        except Exception as e:
            print(f"Error registrando respuesta entrante: {e}")
            self.db.rollback()
            return False

    # Query para actualizar la fecha de última actividad de un ticket
    def actualizar_ultima_actividad_ticket(self, ticket_id):
        """
        Actualiza la fecha de última actividad de un ticket
        """
        try:
            sql = text("""
                UPDATE intranet_correos_microsoft 
                SET updated_at = NOW()
                WHERE id = :ticket_id
            """)
            
            self.db.execute(sql, {"ticket_id": ticket_id})
            self.db.commit()
            
            return True
            
        except Exception as e:
            print(f"Error actualizando última actividad: {e}")
            self.db.rollback()
            return False

    # Query para buscar tickets con subject similar
    def buscar_ticket_por_subject_similar(self, subject_limpio, from_email):
        """
        Busca tickets con subject similar al proporcionado
        Útil para detectar hilos cuando conversation_id no coincide
        """
        try:
            sql = text("""
                SELECT TOP 1 id, subject, from_email, conversation_id, created_at
                FROM intranet_correos_microsoft 
                WHERE subject LIKE :subject_pattern
                AND from_email = :from_email
                AND estado = 1
                AND created_at >= DATEADD(day, -7, GETDATE())
                ORDER BY created_at DESC
            """)
            
            # Buscar con patrón LIKE para subjects similares
            subject_pattern = f"%{subject_limpio}%"
            
            result = self.db.execute(sql, {
                "subject_pattern": subject_pattern,
                "from_email": from_email
            }).fetchone()
            
            if result:
                return {
                    'id': result[0],
                    'subject': result[1],
                    'from_email': result[2],
                    'conversation_id': result[3],
                    'created_at': result[4]
                }
                
            return None
            
        except Exception as e:
            print(f"Error buscando ticket por subject similar: {e}")
            return None

    # Query para buscar el ticket más reciente de un remitente específico
    def buscar_ticket_reciente_por_email(self, from_email, days=7):
        """
        Busca el ticket más reciente de un remitente específico
        """
        try:
            sql = text("""
                SELECT TOP 1 id, subject, from_email, conversation_id, created_at
                FROM intranet_correos_microsoft 
                WHERE from_email = :from_email
                AND estado = 1
                AND created_at >= DATEADD(day, :days_back, GETDATE())
                ORDER BY created_at DESC
            """)
            
            result = self.db.execute(sql, {
                "from_email": from_email,
                "days_back": -days
            }).fetchone()
            
            if result:
                return {
                    'id': result[0],
                    'subject': result[1],
                    'from_email': result[2],
                    'conversation_id': result[3],
                    'created_at': result[4]
                }
                
            return None
            
        except Exception as e:
            print(f"Error buscando ticket reciente por email: {e}")
            return None

    # Query para obtener métricas del dashboard
    def obtener_metricas_dashboard(self, fecha_inicio=None, fecha_fin=None):
        """
        Obtiene métricas del dashboard para tickets con las condiciones especificadas:
        - Total de tickets: activo=1 AND ticket=1
        - Gestión: tipo_ticket=1
        - Estratégicos: tipo_ticket=2
        - Prioridad alta: prioridad=3
        - Estados: Abiertos (estado=1), En proceso (estado=2), Completados (estado=3)
        - Top 3 tipos de soporte más frecuentes
        - Top 3 macroprocesos más frecuentes
        - Top 3 prioridades más frecuentes
        - Top 3 asignados más frecuentes
        - Distribución por tipo de ticket (Gestión y Estratégico)
        """
        try:
            # Opción 1: SQLAlchemy ORM (más legible, orientado a objetos)
            
            # Query base con filtros
            query = self.db.query(
                func.count().label('total_tickets'),
                func.sum(case((CorreosMicrosoftModel.tipo_ticket == 1, 1), else_=0)).label('gestion'),
                func.sum(case((CorreosMicrosoftModel.tipo_ticket == 2, 1), else_=0)).label('estrategicos'),
                func.sum(case((CorreosMicrosoftModel.prioridad == 3, 1), else_=0)).label('prioridad_alta'),
                func.sum(case((CorreosMicrosoftModel.estado == 1, 1), else_=0)).label('abiertos'),
                func.sum(case((CorreosMicrosoftModel.estado == 2, 1), else_=0)).label('en_proceso'),
                func.sum(case((CorreosMicrosoftModel.estado == 3, 1), else_=0)).label('completados')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1
            )
            
            # Agregar filtros de fecha si se proporcionan
            if fecha_inicio:
                query = query.filter(func.date(CorreosMicrosoftModel.received_date_time) >= fecha_inicio)
            
            if fecha_fin:
                query = query.filter(func.date(CorreosMicrosoftModel.received_date_time) <= fecha_fin)
            
            # Ejecutar query
            result = query.one()
            
            # Query para obtener los 3 tipos de soporte más frecuentes
            query_tipos_soporte = self.db.query(
                CorreosMicrosoftModel.tipo_soporte,
                IntranetTipoSoporteModel.nombre.label('nombre_tipo_soporte'),
                func.count(CorreosMicrosoftModel.tipo_soporte).label('cantidad')
            ).join(
                IntranetTipoSoporteModel,
                CorreosMicrosoftModel.tipo_soporte == IntranetTipoSoporteModel.id
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.tipo_soporte.isnot(None),
                CorreosMicrosoftModel.tipo_soporte != 0
            )
            
            # Agregar filtros de fecha para tipos de soporte
            if fecha_inicio:
                query_tipos_soporte = query_tipos_soporte.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) >= fecha_inicio
                )
            
            if fecha_fin:
                query_tipos_soporte = query_tipos_soporte.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) <= fecha_fin
                )
            
            # Agrupar, ordenar y limitar a los 3 primeros
            tipos_soporte_result = query_tipos_soporte.group_by(
                CorreosMicrosoftModel.tipo_soporte,
                IntranetTipoSoporteModel.nombre
            ).order_by(
                func.count(CorreosMicrosoftModel.tipo_soporte).desc()
            ).limit(3).all()
            
            # Formatear tipos de soporte
            tipos_soporte = []
            for tipo in tipos_soporte_result:
                tipos_soporte.append({
                    'id': tipo[0],
                    'nombre': tipo[1],
                    'cantidad': int(tipo[2])
                })
            
            # Query para obtener los 3 macroprocesos más frecuentes
            query_macroprocesos = self.db.query(
                CorreosMicrosoftModel.macroproceso,
                IntranetPerfilesMacroprocesoModel.nombre.label('nombre_macroproceso'),
                func.count(CorreosMicrosoftModel.macroproceso).label('cantidad')
            ).join(
                IntranetPerfilesMacroprocesoModel,
                CorreosMicrosoftModel.macroproceso == IntranetPerfilesMacroprocesoModel.id
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.macroproceso.isnot(None),
                CorreosMicrosoftModel.macroproceso != 0
            )
            
            # Agregar filtros de fecha para macroprocesos
            if fecha_inicio:
                query_macroprocesos = query_macroprocesos.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) >= fecha_inicio
                )
            
            if fecha_fin:
                query_macroprocesos = query_macroprocesos.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) <= fecha_fin
                )
            
            # Agrupar, ordenar y limitar a los 3 primeros
            macroprocesos_result = query_macroprocesos.group_by(
                CorreosMicrosoftModel.macroproceso,
                IntranetPerfilesMacroprocesoModel.nombre
            ).order_by(
                func.count(CorreosMicrosoftModel.macroproceso).desc()
            ).limit(3).all()
            
            # Formatear macroprocesos
            macroprocesos = []
            for macro in macroprocesos_result:
                macroprocesos.append({
                    'id': macro[0],
                    'nombre': macro[1],
                    'cantidad': int(macro[2])
                })
            
            # Query para obtener las 3 prioridades más frecuentes
            query_prioridades = self.db.query(
                CorreosMicrosoftModel.prioridad,
                IntranetTipoPrioridadModel.nombre.label('nombre_prioridad'),
                func.count(CorreosMicrosoftModel.prioridad).label('cantidad')
            ).join(
                IntranetTipoPrioridadModel,
                CorreosMicrosoftModel.prioridad == IntranetTipoPrioridadModel.id
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.prioridad.isnot(None),
                CorreosMicrosoftModel.prioridad != 0
            )
            
            # Agregar filtros de fecha para prioridades
            if fecha_inicio:
                query_prioridades = query_prioridades.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) >= fecha_inicio
                )
            
            if fecha_fin:
                query_prioridades = query_prioridades.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) <= fecha_fin
                )
            
            # Agrupar, ordenar y limitar a los 3 primeros
            prioridades_result = query_prioridades.group_by(
                CorreosMicrosoftModel.prioridad,
                IntranetTipoPrioridadModel.nombre
            ).order_by(
                func.count(CorreosMicrosoftModel.prioridad).desc()
            ).limit(3).all()
            
            # Formatear prioridades
            prioridades = []
            for prio in prioridades_result:
                prioridades.append({
                    'id': prio[0],
                    'nombre': prio[1],
                    'cantidad': int(prio[2])
                })
            
            # Query para obtener los 3 asignados más frecuentes
            query_asignados = self.db.query(
                CorreosMicrosoftModel.asignado,
                IntranetUsuariosGestionTicModel.nombre.label('nombre_asignado'),
                func.count(CorreosMicrosoftModel.asignado).label('cantidad')
            ).join(
                IntranetUsuariosGestionTicModel,
                CorreosMicrosoftModel.asignado == IntranetUsuariosGestionTicModel.id
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.asignado.isnot(None),
                CorreosMicrosoftModel.asignado != 0
            )
            
            # Agregar filtros de fecha para asignados
            if fecha_inicio:
                query_asignados = query_asignados.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) >= fecha_inicio
                )
            
            if fecha_fin:
                query_asignados = query_asignados.filter(
                    func.date(CorreosMicrosoftModel.received_date_time) <= fecha_fin
                )
            
            # Agrupar, ordenar y limitar a los 3 primeros
            asignados_result = query_asignados.group_by(
                CorreosMicrosoftModel.asignado,
                IntranetUsuariosGestionTicModel.nombre
            ).order_by(
                func.count(CorreosMicrosoftModel.asignado).desc()
            ).limit(3).all()
            
            # Formatear asignados
            asignados = []
            for asig in asignados_result:
                asignados.append({
                    'id': asig[0],
                    'nombre': asig[1],
                    'cantidad': int(asig[2])
                })
            
            if result:
                metricas = {
                    'totals': {
                        'total': int(result[0] or 0),
                        'gestion': int(result[1] or 0),
                        'estrategicos': int(result[2] or 0),
                        'prioridad_alta': int(result[3] or 0)
                    },
                    'estados': {
                        'abiertos': int(result[4] or 0),
                        'en_proceso': int(result[5] or 0),
                        'completados': int(result[6] or 0)
                    },
                    'tipos_soporte': tipos_soporte,
                    'macroprocesos': macroprocesos,
                    'prioridades': prioridades,
                    'asignados': asignados
                }
                
                return metricas
            
            # Si no hay resultados, devolver métricas vacías
            return {
                'totals': {
                    'total': 0,
                    'gestion': 0,
                    'estrategicos': 0,
                    'prioridad_alta': 0
                },
                'estados': {
                    'abiertos': 0,
                    'en_proceso': 0,
                    'completados': 0
                },
                'tipos_soporte': [],
                'macroprocesos': [],
                'prioridades': [],
                'asignados': []
            }
            
        except Exception as e:
            print(f"Error obteniendo métricas del dashboard: {e}")
            raise CustomException(f"Error obteniendo métricas: {str(e)}")

    # query para obtener indicadores de gestión mensual
    def obtener_indicadores_gestion(self, anio):
        """
        Obtiene indicadores de gestión mensual optimizado con SQLAlchemy ORM.
        - Total de tickets completados por mes
        - Tickets cerrados oportunamente (fecha_cierre <= fecha_vencimiento)
        - Tickets cerrados no oportunamente (fecha_cierre > fecha_vencimiento)
        - Tickets pendientes (abiertos/en proceso)
        - Tickets ingresados
        - Tickets abiertos/en proceso al final del mes
        - Porcentaje de cumplimiento y acumulado
        Args:
            anio: Año para el cual obtener los indicadores
        """
        try:
            # CorreosMicrosoftModel ya está importado a nivel de módulo
            meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

            # 1. Tickets completados por mes
            completados_q = self.db.query(
                extract('month', CorreosMicrosoftModel.fecha_cierre).label('mes'),
                func.count().label('total_completados'),
                func.sum(case(
                    (and_(CorreosMicrosoftModel.fecha_vencimiento != None, cast(CorreosMicrosoftModel.fecha_cierre, Date) <= cast(CorreosMicrosoftModel.fecha_vencimiento, Date)), 1),
                    else_=0)).label('oportunos'),
                func.sum(case(
                    (and_(CorreosMicrosoftModel.fecha_vencimiento != None, cast(CorreosMicrosoftModel.fecha_cierre, Date) > cast(CorreosMicrosoftModel.fecha_vencimiento, Date)), 1),
                    else_=0)).label('no_oportunos'),
                func.sum(case(
                    (CorreosMicrosoftModel.fecha_vencimiento == None, 1),
                    else_=0)).label('sin_fecha_vencimiento')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.estado == 3,
                CorreosMicrosoftModel.tipo_ticket == 1,
                CorreosMicrosoftModel.fecha_cierre != None,
                func.extract('year', CorreosMicrosoftModel.fecha_cierre) == anio
            ).group_by(extract('month', CorreosMicrosoftModel.fecha_cierre))

            completados = {int(row.mes): {
                'total_completados': row.total_completados,
                'oportunos': row.oportunos,
                'no_oportunos': row.no_oportunos,
                'sin_fecha_vencimiento': row.sin_fecha_vencimiento
            } for row in completados_q.all()}

            # 2. Tickets pendientes (abiertos/en proceso) por mes (fecha_vencimiento)
            pendientes_q = self.db.query(
                extract('month', CorreosMicrosoftModel.fecha_vencimiento).label('mes'),
                func.count().label('total_pendientes')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.estado.in_([1, 2]),
                CorreosMicrosoftModel.fecha_vencimiento != None,
                func.extract('year', CorreosMicrosoftModel.fecha_vencimiento) == anio
            ).group_by(extract('month', CorreosMicrosoftModel.fecha_vencimiento))

            pendientes = {int(row.mes): row.total_pendientes for row in pendientes_q.all()}

            # 3. Tickets ingresados por mes (received_date)
            ingresados_q = self.db.query(
                extract('month', CorreosMicrosoftModel.received_date).label('mes'),
                func.count().label('total_ingresados')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.received_date != None,
                func.extract('year', CorreosMicrosoftModel.received_date) == anio
            ).group_by(extract('month', CorreosMicrosoftModel.received_date))

            ingresados = {int(row.mes): row.total_ingresados for row in ingresados_q.all()}

            # 4. Tickets abiertos/en proceso al final del mes (fecha_vencimiento)
            abiertos_q = self.db.query(
                extract('month', CorreosMicrosoftModel.fecha_vencimiento).label('mes'),
                func.count().label('total_abiertos')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.estado.in_([1, 2]),
                CorreosMicrosoftModel.fecha_vencimiento != None,
                func.extract('year', CorreosMicrosoftModel.fecha_vencimiento) == anio
            ).group_by(extract('month', CorreosMicrosoftModel.fecha_vencimiento))

            abiertos = {int(row.mes): row.total_abiertos for row in abiertos_q.all()}

            # 5. Obtener porcentaje_meta desde dbo.consecutivos (tipo='META')
            porcentaje_meta = 0
            try:
                sql_meta = "SELECT siguiente FROM dbo.consecutivos WHERE tipo = 'META';"
                result_meta = self.db.execute(text(sql_meta)).fetchone()
                if result_meta and result_meta[0] is not None:
                    porcentaje_meta = float(result_meta[0])
            except Exception as e:
                print(f"Error obteniendo porcentaje_meta: {e}")
                porcentaje_meta = None

            # 6. Procesamiento y armado de indicadores
            indicadores = []
            total_oportunos_acumulado = 0
            total_completados_acumulado = 0
            total_no_oportunos_acumulado = 0
            total_sin_respuesta_acumulado = 0
            total_pendientes_acumulado = 0
            total_a_vencer_acumulado = 0
            total_ingresados_acumulado = 0

            for i in range(1, 13):
                datos = completados.get(i)
                pendientes_mes = pendientes.get(i, 0)
                ingresados_mes = ingresados.get(i, 0)
                abiertos_mes = abiertos.get(i, 0)

                if datos:
                    total_completados_acumulado += datos['total_completados']
                    total_oportunos_acumulado += datos['oportunos']
                    total_no_oportunos_acumulado += datos['no_oportunos']
                    total_pendientes_acumulado += pendientes_mes
                    total_ingresados_acumulado += ingresados_mes
                    total_a_vencer_mes = datos['total_completados'] + pendientes_mes
                    total_a_vencer_acumulado += total_a_vencer_mes
                    sin_respuesta = datos['total_completados'] - datos['oportunos'] - datos['no_oportunos'] + pendientes_mes
                    total_sin_respuesta_acumulado += sin_respuesta
                    porcentaje = round((datos['oportunos'] / datos['total_completados'] * 100), 2) if datos['total_completados'] > 0 else 0
                    porcentaje_acumulado = round((total_oportunos_acumulado / total_a_vencer_acumulado * 100), 2) if total_a_vencer_acumulado > 0 else 0
                    indicadores.append({
                        'mes': meses[i-1],
                        'mes_numero': i,
                        'total_completados': total_a_vencer_mes,
                        'oportunos': datos['oportunos'],
                        'no_oportunos': datos['no_oportunos'],
                        'sin_respuesta': sin_respuesta,
                        'total_ingresados': ingresados_mes,
                        'tickets_abiertos': abiertos_mes,
                        'porcentaje': porcentaje,
                        'porcentaje_acumulado': porcentaje_acumulado,
                        'porcentaje_meta': porcentaje_meta
                    })
                else:
                    total_pendientes_acumulado += pendientes_mes
                    total_ingresados_acumulado += ingresados_mes
                    total_a_vencer_acumulado += pendientes_mes
                    if pendientes_mes > 0:
                        total_sin_respuesta_acumulado += pendientes_mes
                    porcentaje_acumulado = round((total_oportunos_acumulado / total_a_vencer_acumulado * 100), 2) if total_a_vencer_acumulado > 0 else 0
                    indicadores.append({
                        'mes': meses[i-1],
                        'mes_numero': i,
                        'total_completados': pendientes_mes,
                        'oportunos': 0,
                        'no_oportunos': 0,
                        'sin_respuesta': pendientes_mes,
                        'total_ingresados': ingresados_mes,
                        'tickets_abiertos': abiertos_mes,
                        'porcentaje': 0,
                        'porcentaje_acumulado': porcentaje_acumulado,
                        'porcentaje_meta': porcentaje_meta
                    })

            return {
                'anio': anio,
                'indicadores': indicadores,
                'totales': {
                    'total_completados': total_a_vencer_acumulado,
                    'oportunos': total_oportunos_acumulado,
                    'no_oportunos': total_no_oportunos_acumulado,
                    'sin_respuesta': total_sin_respuesta_acumulado,
                    'total_ingresados': total_ingresados_acumulado,
                    'porcentaje_global': round((total_oportunos_acumulado / total_a_vencer_acumulado * 100), 2) if total_a_vencer_acumulado > 0 else 0
                }
            }
        except Exception as e:
            print(f"Error obteniendo indicadores de gestión: {e}")
            raise CustomException(f"Error obteniendo indicadores de gestión: {str(e)}")

    # Query para obtener indicadores estratégicos (mismo formato que gestión)
    def obtener_indicadores_estrategicos(self, anio):
        """
        Obtiene indicadores de tickets estratégicos (tipo_ticket = 2)
        Retorna la misma estructura que obtener_indicadores_gestion
        Cuenta tickets por origen_estrategico: 1=Proyectos, 2=ACPM, 3=Actividades informe gestión
        """
        try:
            meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

            # 1. Tickets completados por mes y origen (fecha_cierre)
            completados_q = self.db.query(
                extract('month', CorreosMicrosoftModel.fecha_cierre).label('mes'),
                CorreosMicrosoftModel.origen_estrategico,
                func.count().label('total_completados'),
                func.sum(case(
                    (and_(CorreosMicrosoftModel.fecha_vencimiento != None, cast(CorreosMicrosoftModel.fecha_cierre, Date) <= cast(CorreosMicrosoftModel.fecha_vencimiento, Date)), 1),
                    else_=0)).label('oportunos'),
                func.sum(case(
                    (and_(CorreosMicrosoftModel.fecha_vencimiento != None, cast(CorreosMicrosoftModel.fecha_cierre, Date) > cast(CorreosMicrosoftModel.fecha_vencimiento, Date)), 1),
                    else_=0)).label('no_oportunos')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.estado == 3,
                CorreosMicrosoftModel.tipo_ticket == 2,
                CorreosMicrosoftModel.fecha_cierre != None,
                func.extract('year', CorreosMicrosoftModel.fecha_cierre) == anio
            ).group_by(
                extract('month', CorreosMicrosoftModel.fecha_cierre),
                CorreosMicrosoftModel.origen_estrategico
            )

            # Organizar completados por mes
            completados_data = {}
            for row in completados_q.all():
                mes = int(row.mes)
                if mes not in completados_data:
                    completados_data[mes] = []
                completados_data[mes].append({
                    'origen': row.origen_estrategico,
                    'total': row.total_completados,
                    'oportunos': row.oportunos,
                    'no_oportunos': row.no_oportunos
                })

            # 2. Tickets pendientes por mes y origen (fecha_vencimiento)
            pendientes_q = self.db.query(
                extract('month', CorreosMicrosoftModel.fecha_vencimiento).label('mes'),
                CorreosMicrosoftModel.origen_estrategico,
                func.count().label('total_pendientes')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.estado.in_([1, 2]),
                CorreosMicrosoftModel.tipo_ticket == 2,
                CorreosMicrosoftModel.fecha_vencimiento != None,
                func.extract('year', CorreosMicrosoftModel.fecha_vencimiento) == anio
            ).group_by(
                extract('month', CorreosMicrosoftModel.fecha_vencimiento),
                CorreosMicrosoftModel.origen_estrategico
            )

            # Organizar pendientes por mes
            pendientes_data = {}
            for row in pendientes_q.all():
                mes = int(row.mes)
                if mes not in pendientes_data:
                    pendientes_data[mes] = []
                pendientes_data[mes].append({
                    'origen': row.origen_estrategico,
                    'total': row.total_pendientes
                })

            # 3. Tickets ingresados por mes
            ingresados_q = self.db.query(
                extract('month', CorreosMicrosoftModel.received_date).label('mes'),
                func.count().label('total_ingresados')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.tipo_ticket == 2,
                CorreosMicrosoftModel.received_date != None,
                func.extract('year', CorreosMicrosoftModel.received_date) == anio
            ).group_by(extract('month', CorreosMicrosoftModel.received_date))

            ingresados = {int(row.mes): row.total_ingresados for row in ingresados_q.all()}

            # 4. Tickets abiertos/en proceso al final del mes
            abiertos_q = self.db.query(
                extract('month', CorreosMicrosoftModel.fecha_vencimiento).label('mes'),
                func.count().label('total_abiertos')
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.estado.in_([1, 2]),
                CorreosMicrosoftModel.tipo_ticket == 2,
                CorreosMicrosoftModel.fecha_vencimiento != None,
                func.extract('year', CorreosMicrosoftModel.fecha_vencimiento) == anio
            ).group_by(extract('month', CorreosMicrosoftModel.fecha_vencimiento))

            abiertos = {int(row.mes): row.total_abiertos for row in abiertos_q.all()}

            # 5. Obtener porcentaje_meta
            porcentaje_meta = 0
            try:
                sql_meta = "SELECT siguiente FROM dbo.consecutivos WHERE tipo = 'META';"
                result_meta = self.db.execute(text(sql_meta)).fetchone()
                if result_meta and result_meta[0] is not None:
                    porcentaje_meta = float(result_meta[0])
            except Exception as e:
                print(f"Error obteniendo porcentaje_meta: {e}")
                porcentaje_meta = None

            # 6. Procesamiento y armado de indicadores
            indicadores = []
            total_oportunos_acumulado = 0
            total_completados_acumulado = 0
            total_no_oportunos_acumulado = 0
            total_sin_respuesta_acumulado = 0
            total_pendientes_acumulado = 0
            total_a_vencer_acumulado = 0
            total_ingresados_acumulado = 0

            for i in range(1, 13):
                # Datos del mes
                lista_completados = completados_data.get(i, [])
                lista_pendientes = pendientes_data.get(i, [])
                
                ingresados_mes = ingresados.get(i, 0)
                abiertos_mes = abiertos.get(i, 0)

                # Inicializar contadores del mes
                mes_proyectos = 0
                mes_acpm = 0
                mes_actividades = 0
                
                mes_total_completados = 0
                mes_oportunos = 0
                mes_no_oportunos = 0
                
                mes_pendientes = 0

                # Procesar completados
                for item in lista_completados:
                    origen = item['origen']
                    total = item['total']
                    
                    mes_total_completados += total
                    mes_oportunos += item['oportunos']
                    mes_no_oportunos += item['no_oportunos']
                    
                    if origen == 1:
                        mes_proyectos += total
                    elif origen == 2:
                        mes_acpm += total
                    elif origen == 3:
                        mes_actividades += total

                # Procesar pendientes
                for item in lista_pendientes:
                    origen = item['origen']
                    total = item['total']
                    
                    mes_pendientes += total
                    
                    if origen == 1:
                        mes_proyectos += total
                    elif origen == 2:
                        mes_acpm += total
                    elif origen == 3:
                        mes_actividades += total

                # Cálculos totales del mes
                total_a_vencer_mes = mes_total_completados + mes_pendientes
                sin_respuesta_mes = mes_total_completados - mes_oportunos - mes_no_oportunos + mes_pendientes
                
                # Acumulados
                total_completados_acumulado += mes_total_completados
                total_oportunos_acumulado += mes_oportunos
                total_no_oportunos_acumulado += mes_no_oportunos
                total_pendientes_acumulado += mes_pendientes
                total_ingresados_acumulado += ingresados_mes
                total_a_vencer_acumulado += total_a_vencer_mes
                total_sin_respuesta_acumulado += sin_respuesta_mes

                # Porcentajes
                porcentaje = round((mes_oportunos / total_a_vencer_mes * 100), 2) if total_a_vencer_mes > 0 else 0
                porcentaje_acumulado = round((total_oportunos_acumulado / total_a_vencer_acumulado * 100), 2) if total_a_vencer_acumulado > 0 else 0

                indicadores.append({
                    'mes': meses[i-1],
                    'mes_numero': i,
                    'total_completados': total_a_vencer_mes,
                    'oportunos': mes_oportunos,
                    'no_oportunos': mes_no_oportunos,
                    'sin_respuesta': sin_respuesta_mes,
                    'total_ingresados': ingresados_mes,
                    'tickets_abiertos': abiertos_mes,
                    'porcentaje': porcentaje,
                    'porcentaje_acumulado': porcentaje_acumulado,
                    'porcentaje_meta': porcentaje_meta,
                    # Desglose por origen (Closed + Pending)
                    'proyectos': mes_proyectos,
                    'acpm': mes_acpm,
                    'actividades_informe': mes_actividades
                })

            return {
                'anio': anio,
                'indicadores': indicadores,
                'totales': {
                    'total_completados': total_a_vencer_acumulado,
                    'oportunos': total_oportunos_acumulado,
                    'no_oportunos': total_no_oportunos_acumulado,
                    'sin_respuesta': total_sin_respuesta_acumulado,
                    'total_ingresados': total_ingresados_acumulado,
                    'porcentaje_global': round((total_oportunos_acumulado / total_a_vencer_acumulado * 100), 2) if total_a_vencer_acumulado > 0 else 0,
                    # Totales por origen
                    'proyectos': sum(i['proyectos'] for i in indicadores),
                    'acpm': sum(i['acpm'] for i in indicadores),
                    'actividades_informe': sum(i['actividades_informe'] for i in indicadores)
                }
            }
        except Exception as e:
            print(f"Error obteniendo indicadores estratégicos: {e}")
            raise CustomException(f"Error obteniendo indicadores estratégicos: {str(e)}")

    # Query para obtener tickets del periodo (filtrado por mes y tipo gestión)
    def obtener_tickets_periodo(self, anio, mes, tipo_ticket=1, page=1, limit=5):
        """
        Obtiene los tickets del periodo especificado (año y mes)
        Filtros:
        - Activo = 1
        - Ticket = 1
        - Tipo Ticket = tipo_ticket (1=Gestión, 2=Estratégico)
        - Mes/Año de received_date
        """
        try:
            offset = (page - 1) * limit

            # Base query
            base_query = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.tipo_ticket == tipo_ticket,
                func.extract('year', CorreosMicrosoftModel.received_date) == anio,
                func.extract('month', CorreosMicrosoftModel.received_date) == mes
            )

            # Total count for pagination
            total_records = base_query.count()

            # Fetch paginated tickets
            tickets_query = base_query.order_by(
                CorreosMicrosoftModel.received_date.desc()
            ).offset(offset).limit(limit).all()

            tickets = []
            if tickets_query:
                # Obtener IDs para consultas masivas
                prioridad_ids = {t.prioridad for t in tickets_query if t.prioridad}
                estado_ids = {t.estado for t in tickets_query if t.estado}
                asignado_ids = {t.asignado for t in tickets_query if t.asignado}
                tipo_soporte_ids = {t.tipo_soporte for t in tickets_query if t.tipo_soporte}
                macroproceso_ids = {t.macroproceso for t in tickets_query if t.macroproceso}

                # Consultas auxiliares
                prioridades = {p.id: p.nombre for p in self.db.query(IntranetTipoPrioridadModel).filter(IntranetTipoPrioridadModel.id.in_(prioridad_ids)).all()}
                estados = {e.id: e.nombre for e in self.db.query(IntranetEstadosTickets).filter(IntranetEstadosTickets.id.in_(estado_ids)).all()}
                usuarios = {u.id: u.nombre for u in self.db.query(IntranetUsuariosGestionTicModel).filter(IntranetUsuariosGestionTicModel.id.in_(asignado_ids)).all()}
                tipos_soporte = {ts.id: ts.nombre for ts in self.db.query(IntranetTipoSoporteModel).filter(IntranetTipoSoporteModel.id.in_(tipo_soporte_ids)).all()}
                macroprocesos = {m.id: m.nombre for m in self.db.query(IntranetPerfilesMacroprocesoModel).filter(IntranetPerfilesMacroprocesoModel.id.in_(macroproceso_ids)).all()}

                for t in tickets_query:
                    ticket_dict = t.to_frontend_format()
                    ticket_dict['prioridad_nombre'] = prioridades.get(t.prioridad, '')
                    ticket_dict['estado_nombre'] = estados.get(t.estado, '')
                    ticket_dict['responsable_nombre'] = usuarios.get(t.asignado, 'Sin asignar')
                    ticket_dict['tipo_soporte_nombre'] = tipos_soporte.get(t.tipo_soporte, '')
                    ticket_dict['macroproceso_nombre'] = macroprocesos.get(t.macroproceso, '')
                    # Formatear fecha corta dd/mm
                    if t.received_date:
                        ticket_dict['fecha_corta'] = t.received_date.strftime('%d/%m')
                    tickets.append(ticket_dict)

            # Resumen de estados (sobre el total del mes, no solo la página)
            resumen_query = self.db.query(
                CorreosMicrosoftModel.estado,
                func.count(CorreosMicrosoftModel.id)
            ).filter(
                CorreosMicrosoftModel.activo == 1,
                CorreosMicrosoftModel.ticket == 1,
                CorreosMicrosoftModel.tipo_ticket == tipo_ticket,
                func.extract('year', CorreosMicrosoftModel.received_date) == anio,
                func.extract('month', CorreosMicrosoftModel.received_date) == mes
            ).group_by(CorreosMicrosoftModel.estado).all()

            resumen_dict = {estado: count for estado, count in resumen_query}
            
            # Mapeo de estados (ajustar IDs según tu DB)
            # 1: Abierto, 2: En Proceso, 3: Resuelto, 4: Cerrado
            total = total_records
            cerrados = resumen_dict.get(3, 0) + resumen_dict.get(4, 0)
            en_progreso = resumen_dict.get(2, 0)
            abiertos = resumen_dict.get(1, 0)

            resumen = {
                'total': total,
                'cerrados': cerrados,
                'en_progreso': en_progreso,
                'abiertos': abiertos
            }

            return {
                'tickets': tickets, 
                'resumen': resumen,
                'pagination': {
                    'page': page,
                    'limit': limit,
                    'total_records': total_records,
                    'total_pages': (total_records + limit - 1) // limit
                }
            }

        except Exception as e:
            print(f"Error en obtener_tickets_periodo: {str(e)}")
            return {'tickets': [], 'resumen': {}, 'pagination': {}}

    # Query para obtener observación de un mes
    def obtener_observacion_mes(self, anio, mes):
        """
        Obtiene la observación de un mes específico
        """
        try:
            observacion = self.db.query(IntranetObservacionesInformeGestionModel).filter(
                IntranetObservacionesInformeGestionModel.anio == anio,
                IntranetObservacionesInformeGestionModel.mes == mes,
                IntranetObservacionesInformeGestionModel.estado == 1
            ).first()
            
            return observacion.to_dict() if observacion else None
            
        except Exception as e:
            print(f"Error obteniendo observación del mes: {e}")
            return None
    
    # Query para guardar o actualizar observación de un mes
    def guardar_observacion_mes(self, anio, mes, observaciones):
        """
        Guarda o actualiza la observación de un mes específico
        """
        try:
            # Buscar si ya existe una observación para este año y mes
            observacion_existente = self.db.query(IntranetObservacionesInformeGestionModel).filter(
                IntranetObservacionesInformeGestionModel.anio == anio,
                IntranetObservacionesInformeGestionModel.mes == mes,
                IntranetObservacionesInformeGestionModel.estado == 1
            ).first()
            
            if observacion_existente:
                # Actualizar
                observacion_existente.observaciones = observaciones
                self.db.commit()
                return observacion_existente.to_dict()
            else:
                # Crear nueva
                nueva_observacion = IntranetObservacionesInformeGestionModel({
                    'anio': anio,
                    'mes': mes,
                    'observaciones': observaciones,
                    'estado': 1
                })
                self.db.add(nueva_observacion)
                self.db.commit()
                self.db.refresh(nueva_observacion)
                return nueva_observacion.to_dict()
                
        except Exception as e:
            self.db.rollback()
            print(f"Error guardando observación del mes: {e}")
            raise CustomException(f"Error guardando observación: {str(e)}")

    # Query para obtener análisis de causas de un año
    def obtener_analisis_causas(self, anio, tipo_ticket=1):
        """
        Obtiene todos los análisis de causas y acciones de un año específico y tipo de ticket
        """
        try:
            analisis = self.db.query(IntranetCausasInformeGestion).filter(
                IntranetCausasInformeGestion.anio == anio,
                IntranetCausasInformeGestion.tipo_ticket == tipo_ticket,
                IntranetCausasInformeGestion.estado == 1
            ).order_by(
                IntranetCausasInformeGestion.mes.asc()
            ).all()
            
            return [a.to_dict() for a in analisis] if analisis else []
            
        except Exception as e:
            print(f"Error obteniendo análisis de causas: {e}")
            return []
    
    # Query para verificar si existe un análisis para año+mes
    def verificar_analisis_existe(self, anio, mes, tipo_ticket=1):
        """
        Verifica si ya existe un análisis para el año, mes y tipo de ticket especificados
        """
        try:
            existe = self.db.query(IntranetCausasInformeGestion).filter(
                IntranetCausasInformeGestion.anio == anio,
                IntranetCausasInformeGestion.mes == mes,
                IntranetCausasInformeGestion.tipo_ticket == tipo_ticket,
                IntranetCausasInformeGestion.estado == 1
            ).first()
            
            return existe is not None
            
        except Exception as e:
            print(f"Error verificando existencia de análisis: {e}")
            return False
    
    # Query para guardar o actualizar análisis de causas
    def guardar_analisis_causas(self, id_analisis, anio, mes, analisis, acciones, responsable, fecha_compromiso, seguimiento, tipo_ticket=1):
        """
        Guarda o actualiza un análisis de causas y acciones
        """
        try:
            if id_analisis:
                # Actualizar existente
                analisis_existente = self.db.query(IntranetCausasInformeGestion).filter(
                    IntranetCausasInformeGestion.id == id_analisis,
                    IntranetCausasInformeGestion.estado == 1
                ).first()
                
                if not analisis_existente:
                    raise CustomException("Análisis no encontrado")
                
                analisis_existente.analisis = analisis
                analisis_existente.acciones = acciones
                analisis_existente.responsable = responsable
                analisis_existente.fecha_compromiso = fecha_compromiso
                analisis_existente.seguimiento = seguimiento
                # No actualizamos tipo_ticket, se asume que es el mismo
                
                self.db.commit()
                return analisis_existente.to_dict()
            else:
                # Verificar si ya existe uno para este mes/año/tipo
                if self.verificar_analisis_existe(anio, mes, tipo_ticket):
                     raise CustomException("Ya existe un análisis para este periodo y tipo de indicador")

                # Crear nuevo
                nuevo_analisis = IntranetCausasInformeGestion(
                    anio=anio,
                    mes=mes,
                    analisis=analisis,
                    acciones=acciones,
                    responsable=responsable,
                    fecha_compromiso=fecha_compromiso,
                    seguimiento=seguimiento,
                    tipo_ticket=tipo_ticket,
                    estado=1
                )
                self.db.add(nuevo_analisis)
                self.db.commit()
                self.db.refresh(nuevo_analisis)
                return nuevo_analisis.to_dict()
                
        except Exception as e:
            self.db.rollback()
            print(f"Error guardando análisis de causas: {e}")
            raise CustomException(f"Error guardando análisis de causas: {str(e)}")

    # Query para obtener todos los años disponibles (activos) ordenados descendentemente
    def obtener_anios_disponibles(self):
        """Obtiene todos los años disponibles (activos) ordenados descendentemente"""
        try:
            anios = self.db.query(IntranetAniosInformeGestion).filter(
                IntranetAniosInformeGestion.estado == 1
            ).order_by(
                IntranetAniosInformeGestion.anio.desc()
            ).all()
            
            return [anio.to_dict() for anio in anios] if anios else []
            
        except Exception as e:
            print(f"Error obteniendo años disponibles: {e}")
            return []

    # Query para crear un nuevo año en la base de datos
    def crear_anio(self, anio, descripcion=None):
        """Crea un nuevo año en la base de datos"""
        try:
            # Validar que el año no exista ya
            anio_existente = self.db.query(IntranetAniosInformeGestion).filter(
                IntranetAniosInformeGestion.anio == anio
            ).first()
            
            if anio_existente:
                raise CustomException(f"El año {anio} ya existe en el sistema")
            
            # Crear nuevo año
            nuevo_anio = IntranetAniosInformeGestion(
                anio=anio,
                descripcion=descripcion,
                estado=1
            )
            
            self.db.add(nuevo_anio)
            self.db.commit()
            self.db.refresh(nuevo_anio)
            
            return nuevo_anio.to_dict()
            
        except CustomException as ce:
            self.db.rollback()
            raise ce
        except Exception as e:
            self.db.rollback()
            print(f"Error creando año: {e}")
            raise CustomException(f"Error creando año: {str(e)}")

    # ===================================================
    # QUERIES PARA MÓDULO DE LICENCIAS
    # ===================================================

    def crear_licencia(self, data):
        """Crea una nueva licencia en la base de datos"""
        from Models.IntranetLicenciasHistorialModel import IntranetLicenciasHistorialModel
        import json
        
        try:
            # Crear la licencia (el ID se genera automáticamente)
            nueva_licencia = IntranetLicenciasModel(data)
            self.db.add(nueva_licencia)
            self.db.commit()
            self.db.refresh(nueva_licencia)
            
            # Registrar en el historial con nombres legibles
            usuario = data.get('usuario', 'Sistema')
            responsable = data.get('responsable', {})
            responsable_str = f"{responsable.get('nombre', '')} - {responsable.get('cargo', '')}" if responsable else ""
            
            cambios_iniciales = {
                'Tipo de Servicio': {'anterior': None, 'nuevo': data.get('tipoServicio', '')},
                'Proveedor': {'anterior': None, 'nuevo': data.get('proveedor', '')},
                'Producto': {'anterior': None, 'nuevo': data.get('producto', '')},
                'Cantidad': {'anterior': None, 'nuevo': str(data.get('cantidad', 0))},
                'Frecuencia': {'anterior': None, 'nuevo': data.get('frecuencia', '')},
                'Fecha de Compra': {'anterior': None, 'nuevo': data.get('fechaCompra', '')},
                'Fecha de Vencimiento': {'anterior': None, 'nuevo': data.get('fechaVencimiento', '')},
                'Valor': {'anterior': None, 'nuevo': str(data.get('valor', 0))},
                'Método de Pago': {'anterior': None, 'nuevo': data.get('metodoPago', '')},
                'Responsable': {'anterior': None, 'nuevo': responsable_str}
            }
            
            historial = IntranetLicenciasHistorialModel(
                licencia_id=nueva_licencia.id,
                usuario=usuario,
                accion='Creación',
                cambios=json.dumps(cambios_iniciales, ensure_ascii=False)
            )
            self.db.add(historial)
            self.db.commit()
            
            return nueva_licencia.to_dict()
            
        except CustomException as ce:
            self.db.rollback()
            raise ce
        except Exception as e:
            self.db.rollback()
            print(f"Error creando licencia: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error creando licencia: {str(e)}")

    def obtener_licencias(self, filtros=None, page=1, per_page=5):
        """Obtiene todas las licencias con filtros opcionales, paginación y KPIs"""
        
        try:
            # Query con JOIN a tabla terceros para obtener nombre del proveedor
            sql_base = """
                SELECT 
                    il.*,
                    its.nombre as tipo_servicio_nombre,
                    t.nombres as proveedor_nombre,
                    ips.nombre as producto_nombre,
                    imp.nombre as metodo_pago_nombre
                FROM intranet_licencias il
                LEFT JOIN intranet_tipos_servicio its ON il.tipo_servicio_id = its.id
                LEFT JOIN terceros t ON il.proveedor_id = t.nit
                LEFT JOIN intranet_productos_servicios ips ON il.producto_id = ips.id
                LEFT JOIN intranet_metodos_pago imp ON il.metodo_pago_id = imp.id
                WHERE il.estado = 1
            """
            
            # Aplicar filtros si se proporcionan
            condiciones = []
            params = {}
            
            if filtros:
                if 'incluirBajas' in filtros and not filtros['incluirBajas']:
                    condiciones.append("il.baja = 0")
                
                if filtros.get('proveedorId'):
                    condiciones.append("il.proveedor_id = :proveedorId")
                    params['proveedorId'] = filtros['proveedorId']
                
                if filtros.get('tipoServicioId'):
                    condiciones.append("il.tipo_servicio_id = :tipoServicioId")
                    params['tipoServicioId'] = filtros['tipoServicioId']
            
            if condiciones:
                sql_base += " AND " + " AND ".join(condiciones)
            
            sql_base += " ORDER BY il.id DESC"
            
            # Ejecutar consulta
            result = self.db.execute(text(sql_base), params).fetchall()
            total = len(result)
            
            # Calcular KPIs sobre todas las licencias
            from datetime import datetime, timedelta, date
            kpis = self._calcular_kpis_from_raw(result)
            
            # Aplicar paginación
            offset = (page - 1) * per_page
            resultados_paginados = result[offset:offset + per_page]
            
            # Construir respuesta
            licencias = []
            for row in resultados_paginados:
                lic_dict = {
                    'id': row.id,
                    'tipoServicioId': row.tipo_servicio_id,
                    'tipoServicio': row.tipo_servicio_nombre,
                    'proveedorId': row.proveedor_id,
                    'proveedor': row.proveedor_nombre,
                    'productoId': row.producto_id,
                    'producto': row.producto_nombre,
                    'cantidad': row.cantidad,
                    'frecuencia': row.frecuencia,
                    'fechaCompra': row.fecha_compra.isoformat() if row.fecha_compra else None,
                    'fechaVencimiento': row.fecha_vencimiento.isoformat() if row.fecha_vencimiento else None,
                    'valor': float(row.valor) if row.valor else 0,
                    'metodoPagoId': row.metodo_pago_id,
                    'metodoPago': row.metodo_pago_nombre,
                    'responsable': {'nombre': row.responsable_nombre, 'cargo': row.responsable_cargo} if row.responsable_nombre else None,
                    'observaciones': row.observaciones,
                    'baja': bool(row.baja),
                    'fechaBaja': row.fecha_baja.isoformat() if row.fecha_baja else None,
                    'motivoBaja': row.motivo_baja
                }
                
                # Agregar historial de cada licencia
                historial = self.db.query(IntranetLicenciasHistorialModel).filter(
                    IntranetLicenciasHistorialModel.licencia_id == row.id
                ).order_by(IntranetLicenciasHistorialModel.fecha.desc()).all()
                lic_dict['historial'] = [h.to_dict() for h in historial]
                
                licencias.append(lic_dict)
            
            return {
                'licencias': licencias,
                'total': total,
                'page': page,
                'per_page': per_page,
                'total_pages': (total + per_page - 1) // per_page,
                'kpis': kpis
            }
            

            
        except Exception as e:
            print(f"Error obteniendo licencias: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo licencias: {str(e)}")
    
    def _calcular_kpis(self, licencias):
        """Calcula KPIs sobre todas las licencias"""
        from datetime import datetime, date
        
        hoy = date.today()
        total = len(licencias)
        criticas = 0
        proximas = 0
        vigentes = 0
        costo_anual_total = 0
        
        for lic in licencias:
            # Calcular estado según fecha de vencimiento (usar snake_case del modelo)
            if lic.fecha_vencimiento:
                dias_restantes = (lic.fecha_vencimiento - hoy).days
                
                if dias_restantes < 0 or dias_restantes <= 8:
                    criticas += 1
                elif dias_restantes <= 30:
                    proximas += 1
                else:
                    vigentes += 1
            
            # Calcular costo anual (solo no dadas de baja)
            if not lic.baja and lic.valor:
                valor = float(lic.valor)
                if lic.frecuencia and lic.frecuencia.lower() == 'mensual':
                    costo_anual_total += valor * 12
                else:
                    costo_anual_total += valor
        
        return {
            'total': total,
            'criticas': criticas,
            'proximas': proximas,
            'vigentes': vigentes,
            'costo_anual_total': round(costo_anual_total, 2)
        }
    
    def _calcular_kpis_from_raw(self, rows):
        """Calcula KPIs desde resultados raw de consulta SQL"""
        from datetime import date
        
        hoy = date.today()
        total = len(rows)
        criticas = 0
        proximas = 0
        vigentes = 0
        costo_anual_total = 0
        
        for row in rows:
            # Calcular estado según fecha de vencimiento
            if row.fecha_vencimiento:
                dias_restantes = (row.fecha_vencimiento - hoy).days
                
                if dias_restantes < 0 or dias_restantes <= 8:
                    criticas += 1
                elif dias_restantes <= 30:
                    proximas += 1
                else:
                    vigentes += 1
            
            # Calcular costo anual (solo no dadas de baja)
            if not row.baja and row.valor:
                valor = float(row.valor)
                if row.frecuencia and row.frecuencia.lower() == 'mensual':
                    costo_anual_total += valor * 12
                else:
                    costo_anual_total += valor
        
        return {
            'total': total,
            'criticas': criticas,
            'proximas': proximas,
            'vigentes': vigentes,
            'costo_anual_total': round(costo_anual_total, 2)
        }

    def obtener_todas_licencias_excel(self, filtros=None):
        """Obtiene todas las licencias sin paginación para exportar a Excel"""
        try:
            # Query con JOIN a tabla terceros
            sql_base = """
                SELECT 
                    il.*,
                    its.nombre as tipo_servicio_nombre,
                    t.nombres as proveedor_nombre,
                    ips.nombre as producto_nombre,
                    imp.nombre as metodo_pago_nombre
                FROM intranet_licencias il
                LEFT JOIN intranet_tipos_servicio its ON il.tipo_servicio_id = its.id
                LEFT JOIN terceros t ON il.proveedor_id = t.nit
                LEFT JOIN intranet_productos_servicios ips ON il.producto_id = ips.id
                LEFT JOIN intranet_metodos_pago imp ON il.metodo_pago_id = imp.id
                WHERE il.estado = 1
            """
            
            # Aplicar filtros
            condiciones = []
            params = {}
            
            if filtros:
                if 'incluirBajas' in filtros and not filtros['incluirBajas']:
                    condiciones.append("il.baja = 0")
                
                if filtros.get('proveedorId'):
                    condiciones.append("il.proveedor_id = :proveedorId")
                    params['proveedorId'] = filtros['proveedorId']
                
                if filtros.get('tipoServicioId'):
                    condiciones.append("il.tipo_servicio_id = :tipoServicioId")
                    params['tipoServicioId'] = filtros['tipoServicioId']
            
            if condiciones:
                sql_base += " AND " + " AND ".join(condiciones)
            
            sql_base += " ORDER BY il.id DESC"
            
            # Ejecutar consulta
            result = self.db.execute(text(sql_base), params).fetchall()
            
            # Construir respuesta
            licencias = []
            for row in result:
                lic_dict = {
                    'id': row.id,
                    'tipoServicioId': row.tipo_servicio_id,
                    'tipoServicio': row.tipo_servicio_nombre,
                    'proveedorId': row.proveedor_id,
                    'proveedor': row.proveedor_nombre,
                    'productoId': row.producto_id,
                    'producto': row.producto_nombre,
                    'cantidad': row.cantidad,
                    'frecuencia': row.frecuencia,
                    'fechaCompra': row.fecha_compra.isoformat() if row.fecha_compra else None,
                    'fechaVencimiento': row.fecha_vencimiento.isoformat() if row.fecha_vencimiento else None,
                    'valor': float(row.valor) if row.valor else 0,
                    'metodoPagoId': row.metodo_pago_id,
                    'metodoPago': row.metodo_pago_nombre,
                    'responsable': {'nombre': row.responsable_nombre, 'cargo': row.responsable_cargo} if row.responsable_nombre else None,
                    'observaciones': row.observaciones,
                    'baja': bool(row.baja),
                    'fechaBaja': row.fecha_baja.isoformat() if row.fecha_baja else None,
                    'motivoBaja': row.motivo_baja
                }
                licencias.append(lic_dict)
            
            return licencias
            
        except Exception as e:
            print(f"Error obteniendo licencias para Excel: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo licencias: {str(e)}")

    def obtener_licencia_por_id(self, licencia_id):
        """Obtiene una licencia específica por su ID"""

        try:
            sql = """
                SELECT 
                    il.*,
                    its.nombre as tipo_servicio_nombre,
                    t.nombres as proveedor_nombre,
                    ips.nombre as producto_nombre,
                    imp.nombre as metodo_pago_nombre
                FROM intranet_licencias il
                LEFT JOIN intranet_tipos_servicio its ON il.tipo_servicio_id = its.id
                LEFT JOIN terceros t ON il.proveedor_id = t.nit
                LEFT JOIN intranet_productos_servicios ips ON il.producto_id = ips.id
                LEFT JOIN intranet_metodos_pago imp ON il.metodo_pago_id = imp.id
                WHERE il.id = :licencia_id AND il.estado = 1
            """
            
            result = self.db.execute(text(sql), {'licencia_id': licencia_id}).fetchone()
            
            if not result:
                raise CustomException("Licencia no encontrada.")
            
            lic_dict = {
                'id': result.id,
                'tipoServicioId': result.tipo_servicio_id,
                'tipoServicio': result.tipo_servicio_nombre,
                'proveedorId': result.proveedor_id,
                'proveedor': result.proveedor_nombre,
                'productoId': result.producto_id,
                'producto': result.producto_nombre,
                'cantidad': result.cantidad,
                'frecuencia': result.frecuencia,
                'fechaCompra': result.fecha_compra.isoformat() if result.fecha_compra else None,
                'fechaVencimiento': result.fecha_vencimiento.isoformat() if result.fecha_vencimiento else None,
                'valor': float(result.valor) if result.valor else 0,
                'metodoPagoId': result.metodo_pago_id,
                'metodoPago': result.metodo_pago_nombre,
                'responsable': {'nombre': result.responsable_nombre, 'cargo': result.responsable_cargo} if result.responsable_nombre else None,
                'observaciones': result.observaciones,
                'baja': bool(result.baja),
                'fechaBaja': result.fecha_baja.isoformat() if result.fecha_baja else None,
                'motivoBaja': result.motivo_baja
            }
            
            # Agregar historial a la respuesta
            historial = self.db.query(IntranetLicenciasHistorialModel).filter(
                IntranetLicenciasHistorialModel.licencia_id == licencia_id
            ).order_by(IntranetLicenciasHistorialModel.fecha.desc()).all()
            lic_dict['historial'] = [h.to_dict() for h in historial]
            
            return lic_dict
            
        except CustomException as ce:
            raise ce
        except Exception as e:
            print(f"Error obteniendo licencia por ID: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo licencia: {str(e)}")

    def actualizar_licencia(self, licencia_id, data):
        """Actualiza una licencia existente"""

        try:
            licencia = self.db.query(IntranetLicenciasModel).filter(
                IntranetLicenciasModel.id == licencia_id,
                IntranetLicenciasModel.estado == 1
            ).first()
            
            if not licencia:
                raise CustomException("Licencia no encontrada.")
            
            # Función auxiliar para normalizar fechas para comparación
            def normalizar_fecha(fecha):
                if fecha is None:
                    return None
                if isinstance(fecha, str):
                    return fecha
                return str(fecha)
            
            cambios_detectados = {}
            
            # Actualizar campos con IDs y detectar cambios (guardando nombres legibles)
            if 'tipoServicioId' in data:
                if licencia.tipo_servicio_id != data['tipoServicioId']:
                    # Obtener nombres para el historial
                    tipo_anterior = self.db.query(IntranetTiposServicioModel).filter(
                        IntranetTiposServicioModel.id == licencia.tipo_servicio_id
                    ).first()
                    tipo_nuevo = self.db.query(IntranetTiposServicioModel).filter(
                        IntranetTiposServicioModel.id == data['tipoServicioId']
                    ).first()
                    cambios_detectados['Tipo de Servicio'] = {
                        'anterior': tipo_anterior.nombre if tipo_anterior else str(licencia.tipo_servicio_id),
                        'nuevo': tipo_nuevo.nombre if tipo_nuevo else str(data['tipoServicioId'])
                    }
                licencia.tipo_servicio_id = data['tipoServicioId']
                
            if 'proveedorId' in data:
                if licencia.proveedor_id != data['proveedorId']:
                    prov_anterior = self.db.query(IntranetProveedoresModel).filter(
                        IntranetProveedoresModel.id == licencia.proveedor_id
                    ).first()
                    prov_nuevo = self.db.query(IntranetProveedoresModel).filter(
                        IntranetProveedoresModel.id == data['proveedorId']
                    ).first()
                    cambios_detectados['Proveedor'] = {
                        'anterior': prov_anterior.nombre if prov_anterior else str(licencia.proveedor_id),
                        'nuevo': prov_nuevo.nombre if prov_nuevo else str(data['proveedorId'])
                    }
                licencia.proveedor_id = data['proveedorId']
                
            if 'productoId' in data:
                if licencia.producto_id != data['productoId']:
                    prod_anterior = self.db.query(IntranetProductosServiciosModel).filter(
                        IntranetProductosServiciosModel.id == licencia.producto_id
                    ).first()
                    prod_nuevo = self.db.query(IntranetProductosServiciosModel).filter(
                        IntranetProductosServiciosModel.id == data['productoId']
                    ).first()
                    cambios_detectados['Producto'] = {
                        'anterior': prod_anterior.nombre if prod_anterior else str(licencia.producto_id),
                        'nuevo': prod_nuevo.nombre if prod_nuevo else str(data['productoId'])
                    }
                licencia.producto_id = data['productoId']
                
            if 'cantidad' in data:
                if licencia.cantidad != data['cantidad']:
                    cambios_detectados['Cantidad'] = {
                        'anterior': licencia.cantidad,
                        'nuevo': data['cantidad']
                    }
                licencia.cantidad = data['cantidad']
                
            if 'frecuencia' in data:
                if licencia.frecuencia != data['frecuencia']:
                    cambios_detectados['Frecuencia'] = {
                        'anterior': licencia.frecuencia,
                        'nuevo': data['frecuencia']
                    }
                licencia.frecuencia = data['frecuencia']
                
            if 'fechaCompra' in data:
                fecha_actual = normalizar_fecha(licencia.fecha_compra)
                fecha_nueva = normalizar_fecha(data['fechaCompra'])
                if fecha_actual != fecha_nueva:
                    cambios_detectados['Fecha de Compra'] = {
                        'anterior': fecha_actual,
                        'nuevo': fecha_nueva
                    }
                licencia.fecha_compra = data['fechaCompra']
                
            if 'fechaVencimiento' in data:
                fecha_actual = normalizar_fecha(licencia.fecha_vencimiento)
                fecha_nueva = normalizar_fecha(data['fechaVencimiento'])
                if fecha_actual != fecha_nueva:
                    cambios_detectados['Fecha de Vencimiento'] = {
                        'anterior': fecha_actual,
                        'nuevo': fecha_nueva
                    }
                licencia.fecha_vencimiento = data['fechaVencimiento']
                
            if 'valor' in data:
                if licencia.valor != data['valor']:
                    cambios_detectados['Valor'] = {
                        'anterior': licencia.valor,
                        'nuevo': data['valor']
                    }
                licencia.valor = data['valor']
                
            if 'metodoPagoId' in data:
                if licencia.metodo_pago_id != data['metodoPagoId']:
                    metodo_anterior = self.db.query(IntranetMetodosPagoModel).filter(
                        IntranetMetodosPagoModel.id == licencia.metodo_pago_id
                    ).first()
                    metodo_nuevo = self.db.query(IntranetMetodosPagoModel).filter(
                        IntranetMetodosPagoModel.id == data['metodoPagoId']
                    ).first()
                    cambios_detectados['Método de Pago'] = {
                        'anterior': metodo_anterior.nombre if metodo_anterior else str(licencia.metodo_pago_id),
                        'nuevo': metodo_nuevo.nombre if metodo_nuevo else str(data['metodoPagoId'])
                    }
                licencia.metodo_pago_id = data['metodoPagoId']
                
            if 'responsable' in data and isinstance(data['responsable'], dict):
                nuevo_nombre = data['responsable'].get('nombre')
                nuevo_cargo = data['responsable'].get('cargo')
                if licencia.responsable_nombre != nuevo_nombre or licencia.responsable_cargo != nuevo_cargo:
                    cambios_detectados['Responsable'] = {
                        'anterior': f"{licencia.responsable_nombre or ''} - {licencia.responsable_cargo or ''}".strip(' -'),
                        'nuevo': f"{nuevo_nombre or ''} - {nuevo_cargo or ''}".strip(' -')
                    }
                licencia.responsable_nombre = nuevo_nombre
                licencia.responsable_cargo = nuevo_cargo
                
            if 'observaciones' in data:
                obs_actual = licencia.observaciones or ""
                obs_nueva = data['observaciones'] or ""
                if obs_actual.strip() != obs_nueva.strip():
                    cambios_detectados['Observaciones'] = {
                        'anterior': obs_actual,
                        'nuevo': obs_nueva
                    }
                licencia.observaciones = data['observaciones']
                
            if 'baja' in data:
                baja_original = licencia.baja  # Guardar valor original antes de actualizar
                if licencia.baja != data['baja']:
                    cambios_detectados['Estado Baja'] = {
                        'anterior': 'Sí' if licencia.baja else 'No',
                        'nuevo': 'Sí' if data['baja'] else 'No'
                    }
                licencia.baja = data['baja']
                
            if 'fechaBaja' in data:
                fecha_actual = normalizar_fecha(licencia.fecha_baja)
                fecha_nueva = normalizar_fecha(data['fechaBaja'])
                if fecha_actual != fecha_nueva:
                    cambios_detectados['Fecha Baja'] = {
                        'anterior': fecha_actual,
                        'nuevo': fecha_nueva
                    }
                licencia.fecha_baja = data['fechaBaja']
                
            if 'motivoBaja' in data:
                motivo_actual = licencia.motivo_baja or ""
                motivo_nuevo = data['motivoBaja'] or ""
                if motivo_actual.strip() != motivo_nuevo.strip():
                    cambios_detectados['Motivo Baja'] = {
                        'anterior': motivo_actual,
                        'nuevo': motivo_nuevo
                    }
                licencia.motivo_baja = data['motivoBaja']
            
            licencia.updated_at = datetime.now()
            
            # Determinar el tipo de acción
            accion = 'Edición'
            if 'baja' in data:
                baja_original = locals().get('baja_original', licencia.baja)
                if data['baja'] == 1:
                    accion = 'Baja'
                elif data['baja'] == 0 and baja_original == 1:
                    accion = 'Reactivación'
            
            # Crear registro en el historial solo si hubo cambios
            if cambios_detectados:
                # Convertir valores no serializables (fechas, etc.) a strings
                cambios_serializables = {}
                for campo, valores in cambios_detectados.items():
                    cambios_serializables[campo] = {
                        'anterior': str(valores['anterior']) if valores['anterior'] is not None else None,
                        'nuevo': str(valores['nuevo']) if valores['nuevo'] is not None else None
                    }
                
                historial = IntranetLicenciasHistorialModel(
                    licencia_id=licencia_id,
                    usuario="Jeyson Martinez",  # TODO: Obtener del usuario autenticado
                    accion=accion,
                    cambios=json.dumps(cambios_serializables, ensure_ascii=False)
                )
                self.db.add(historial)
            
            self.db.commit()
            self.db.refresh(licencia)
            
            return licencia.to_dict()
            
        except CustomException as ce:
            self.db.rollback()
            raise ce
        except Exception as e:
            self.db.rollback()
            print(f"Error actualizando licencia: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error actualizando licencia: {str(e)}")

    def obtener_historial_licencia(self, licencia_id):
        """Obtiene el historial de cambios de una licencia"""
        
        try:
            historial = self.db.query(IntranetLicenciasHistorialModel).filter(
                IntranetLicenciasHistorialModel.licencia_id == licencia_id
            ).order_by(IntranetLicenciasHistorialModel.fecha.desc()).all()
            
            return [h.to_dict() for h in historial]
            
        except Exception as e:
            print(f"Error obteniendo historial de licencia: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo historial de licencia: {str(e)}")

    def eliminar_licencia(self, licencia_id):
        """Elimina (marca como inactiva) una licencia"""

        try:
            licencia = self.db.query(IntranetLicenciasModel).filter(
                IntranetLicenciasModel.id == licencia_id,
                IntranetLicenciasModel.estado == 1
            ).first()
            
            if not licencia:
                raise CustomException("Licencia no encontrada.")
            
            licencia.estado = 0
            licencia.updated_at = datetime.now()
            
            self.db.commit()
            
            return True
            
        except CustomException as ce:
            self.db.rollback()
            raise ce
        except Exception as e:
            self.db.rollback()
            print(f"Error eliminando licencia: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error eliminando licencia: {str(e)}")

    # ===================================================
    # QUERIES PARA CATÁLOGOS DE LICENCIAS
    # ===================================================

    def obtener_tipos_servicio(self):
        """Obtiene todos los tipos de servicio activos"""
        
        try:
            tipos = self.db.query(IntranetTiposServicioModel).filter(
                IntranetTiposServicioModel.estado == 1
            ).order_by(IntranetTiposServicioModel.nombre.asc()).all()
            
            return [t.to_dict() for t in tipos]
            
        except Exception as e:
            print(f"Error obteniendo tipos de servicio: {e}")
            raise CustomException(f"Error obteniendo tipos de servicio: {str(e)}")

    def obtener_proveedores(self):
        """Obtiene todos los proveedores activos"""

        try:
            sql = """
                SELECT nit AS id, nombres AS nombre 
                FROM terceros WHERE concepto_1 in (1, 3) ORDER BY nombre;"""

            result = self.db.execute(text(sql)).fetchall()
            return [dict(row._mapping) for row in result] if result else []
        except CustomException as e:
            traceback.print_exc()
            print(f"Error al obtener proveedores: {e}")
            raise CustomException(f"{e}")
        finally:
            self.db.close()

    def obtener_productos_servicios(self):
        """Obtiene todos los productos/servicios activos"""

        try:
            productos = self.db.query(IntranetProductosServiciosModel).filter(
                IntranetProductosServiciosModel.estado == 1
            ).order_by(IntranetProductosServiciosModel.nombre.asc()).all()
            
            return [p.to_dict() for p in productos]
            
        except Exception as e:
            print(f"Error obteniendo productos/servicios: {e}")
            raise CustomException(f"Error obteniendo productos/servicios: {str(e)}")

    def obtener_metodos_pago(self):
        """Obtiene todos los métodos de pago activos"""
        
        try:
            metodos = self.db.query(IntranetMetodosPagoModel).filter(
                IntranetMetodosPagoModel.estado == 1
            ).order_by(IntranetMetodosPagoModel.nombre.asc()).all()
            
            return [m.to_dict() for m in metodos]
            
        except Exception as e:
            print(f"Error obteniendo métodos de pago: {e}")
            raise CustomException(f"Error obteniendo métodos de pago: {str(e)}")

    def crear_proveedor(self, nombre):
        """Crea un nuevo proveedor"""
        
        try:
            # Verificar que no exista
            existe = self.db.query(IntranetProveedoresModel).filter(
                IntranetProveedoresModel.nombre == nombre
            ).first()
            
            if existe:
                return existe.to_dict()  # Devolver el existente si ya existe
            
            nuevo_proveedor = IntranetProveedoresModel()
            nuevo_proveedor.nombre = nombre
            
            self.db.add(nuevo_proveedor)
            self.db.commit()
            self.db.refresh(nuevo_proveedor)
            
            return nuevo_proveedor.to_dict()
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando proveedor: {e}")
            raise CustomException(f"Error creando proveedor: {str(e)}")

    def crear_producto_servicio(self, nombre):
        """Crea un nuevo producto/servicio"""
        
        try:
            # Verificar que no exista
            existe = self.db.query(IntranetProductosServiciosModel).filter(
                IntranetProductosServiciosModel.nombre == nombre
            ).first()
            
            if existe:
                return existe.to_dict()  # Devolver el existente si ya existe
            
            nuevo_producto = IntranetProductosServiciosModel()
            nuevo_producto.nombre = nombre
            
            self.db.add(nuevo_producto)
            self.db.commit()
            self.db.refresh(nuevo_producto)
            
            return nuevo_producto.to_dict()
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando producto/servicio: {e}")
            raise CustomException(f"Error creando producto/servicio: {str(e)}")

    def crear_tipo_servicio(self, nombre):
        """Crea un nuevo tipo de servicio"""
        
        try:
            # Verificar que no exista
            existe = self.db.query(IntranetTiposServicioModel).filter(
                IntranetTiposServicioModel.nombre == nombre
            ).first()
            
            if existe:
                return existe.to_dict()
            
            nuevo_tipo = IntranetTiposServicioModel()
            nuevo_tipo.nombre = nombre
            
            self.db.add(nuevo_tipo)
            self.db.commit()
            self.db.refresh(nuevo_tipo)
            
            return nuevo_tipo.to_dict()
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando tipo de servicio: {e}")
            raise CustomException(f"Error creando tipo de servicio: {str(e)}")

    def crear_metodo_pago(self, nombre):
        """Crea un nuevo método de pago"""
        
        try:
            # Verificar que no exista
            existe = self.db.query(IntranetMetodosPagoModel).filter(
                IntranetMetodosPagoModel.nombre == nombre
            ).first()
            
            if existe:
                return existe.to_dict()
            
            nuevo_metodo = IntranetMetodosPagoModel()
            nuevo_metodo.nombre = nombre
            
            self.db.add(nuevo_metodo)
            self.db.commit()
            self.db.refresh(nuevo_metodo)
            
            return nuevo_metodo.to_dict()
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando método de pago: {e}")
            raise CustomException(f"Error creando método de pago: {str(e)}")

    # ===============================================
    # ENDPOINTS PARA REVISIONES GENERALES
    # ===============================================

    def obtener_tipos_revision(self):
        """Obtiene todos los tipos de revisión activos"""
        try:
            tipos = self.db.query(IntranetTipoRevisionModel).filter(
                IntranetTipoRevisionModel.estado == 1
            ).all()
            
            return [tipo.to_dict() for tipo in tipos]
            
        except Exception as e:
            print(f"Error obteniendo tipos de revisión: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo tipos de revisión: {str(e)}")

    def crear_revision(self, data):
        """Crea una nueva revisión general"""
        try:
            nueva_revision = IntranetRevisionesModel()
            nueva_revision.fecha = datetime.strptime(data['fecha'], '%Y-%m-%d').date()
            nueva_revision.tipo_revision_id = data['tipo_revision_id']
            nueva_revision.observaciones = data.get('observaciones', '')
            nueva_revision.usuario = data['usuario']
            
            self.db.add(nueva_revision)
            self.db.commit()
            self.db.refresh(nueva_revision)
            
            # Obtener el nombre del tipo de revisión para la respuesta
            tipo = self.db.query(IntranetTipoRevisionModel).filter(
                IntranetTipoRevisionModel.id == nueva_revision.tipo_revision_id
            ).first()
            
            resultado = nueva_revision.to_dict()
            resultado['tipo'] = tipo.nombre if tipo else None
            
            return resultado
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando revisión: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error creando revisión: {str(e)}")

    def obtener_revisiones(self, page=1, per_page=5):
        """Obtiene revisiones con paginación"""
        try:
            # Calcular offset
            offset = (page - 1) * per_page
            
            # Consulta con JOIN para obtener el nombre del tipo
            query = self.db.query(
                IntranetRevisionesModel,
                IntranetTipoRevisionModel.nombre.label('tipo_nombre')
            ).join(
                IntranetTipoRevisionModel,
                IntranetRevisionesModel.tipo_revision_id == IntranetTipoRevisionModel.id
            ).filter(
                IntranetRevisionesModel.estado == 1
            ).order_by(
                IntranetRevisionesModel.fecha.desc(),
                IntranetRevisionesModel.created_at.desc()
            )
            
            # Total de registros
            total = query.count()
            
            # Aplicar paginación
            revisiones_paginadas = query.limit(per_page).offset(offset).all()
            
            # Construir respuesta
            revisiones = []
            for revision, tipo_nombre in revisiones_paginadas:
                rev_dict = revision.to_dict()
                rev_dict['tipo'] = tipo_nombre
                revisiones.append(rev_dict)
            
            return {
                'revisiones': revisiones,
                'total': total,
                'page': page,
                'per_page': per_page,
                'total_pages': (total + per_page - 1) // per_page  # Redondeo hacia arriba
            }
            
        except Exception as e:
            print(f"Error obteniendo revisiones: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo revisiones: {str(e)}")

    def eliminar_revision(self, revision_id):
        """Elimina (marca como inactiva) una revisión"""
        try:
            revision = self.db.query(IntranetRevisionesModel).filter(
                IntranetRevisionesModel.id == revision_id
            ).first()
            
            if not revision:
                raise CustomException("Revisión no encontrada")
            
            revision.estado = 0
            self.db.commit()
            
            return {"message": "Revisión eliminada correctamente"}
            
        except Exception as e:
            self.db.rollback()
            print(f"Error eliminando revisión: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error eliminando revisión: {str(e)}")

    # ==================== VERSIONES DE LICENCIAS ====================
    
    def crear_version(self, data):
        """Crea una nueva versión del control de licencias"""
        try:
            nueva_version = IntranetVersionesLicenciasModel()
            nueva_version.fecha = datetime.strptime(data['fecha'], '%Y-%m-%d').date()
            nueva_version.version = data['version']
            nueva_version.descripcion = data.get('descripcion', '')
            
            self.db.add(nueva_version)
            self.db.commit()
            self.db.refresh(nueva_version)
            
            return {
                'id': nueva_version.id,
                'fecha': nueva_version.fecha.strftime('%Y-%m-%d'),
                'version': nueva_version.version,
                'descripcion': nueva_version.descripcion
            }
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando versión: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error creando versión: {str(e)}")

    def obtener_versiones(self, page=1, per_page=5):
        """Obtiene versiones con paginación"""
        try:
            # Calcular offset
            offset = (page - 1) * per_page
            
            query = self.db.query(
                IntranetVersionesLicenciasModel
            ).filter(
                IntranetVersionesLicenciasModel.estado == 1
            ).order_by(
                IntranetVersionesLicenciasModel.fecha.desc()
            )
            
            # Total de registros
            total = query.count()
            
            # Aplicar paginación
            versiones_paginadas = query.limit(per_page).offset(offset).all()
            
            # Construir respuesta
            versiones = []
            for version in versiones_paginadas:
                versiones.append({
                    'id': version.id,
                    'fecha': version.fecha.strftime('%Y-%m-%d'),
                    'version': version.version,
                    'descripcion': version.descripcion
                })
            
            return {
                'versiones': versiones,
                'total': total,
                'page': page,
                'per_page': per_page,
                'total_pages': (total + per_page - 1) // per_page
            }
            
        except Exception as e:
            print(f"Error obteniendo versiones: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error obteniendo versiones: {str(e)}")

    def eliminar_version(self, version_id):
        """Marca una versión como inactiva (estado = 0)"""
        try:
            version = self.db.query(IntranetVersionesLicenciasModel).filter(
                IntranetVersionesLicenciasModel.id == version_id
            ).first()
            
            if not version:
                raise CustomException("Versión no encontrada")
            
            version.estado = 0
            self.db.commit()
            
            return {"message": "Versión eliminada correctamente"}
            
        except Exception as e:
            self.db.rollback()
            print(f"Error eliminando versión: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise CustomException(f"Error eliminando versión: {str(e)}")
