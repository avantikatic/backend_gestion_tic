from Utils.tools import Tools, CustomException
from Utils.file_handler import FileHandler
from sqlalchemy import text, func, case, extract, and_, or_, Date, cast, collate
from sqlalchemy.exc import IntegrityError
from datetime import datetime, date
import json
import traceback
import os
from pathlib import Path
import pytz
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
from Models.IntranetObservacionesMantenimientoModel import IntranetObservacionesMantenimientoModel
from Models.IntranetCausasMantenimientoModel import IntranetCausasMantenimientoModel
from Models.IntranetActivos import IntranetActivos
from Models.IntranetCausasInformeGestionModel import IntranetCausasInformeGestion
from Models.IntranetAniosInformeGestionModel import IntranetAniosInformeGestion
from Models.IntranetOrigenEstrategicoModel import IntranetOrigenEstrategicoModel
from Models.IntranetOrdenesTrabajo import IntranetOrdenesTrabajo
from Models.IntranetActividadesOrdenesTrabajo import IntranetActividadesOrdenesTrabajo
from Models.IntranetTiposServicioModel import IntranetTiposServicioModel
from Models.IntranetProveedoresModel import IntranetProveedoresModel
from Models.IntranetProductosServiciosModel import IntranetProductosServiciosModel
from Models.IntranetMetodosPagoModel import IntranetMetodosPagoModel
from Models.IntranetLicenciasModel import IntranetLicenciasModel
from Models.IntranetLicenciasHistorialModel import IntranetLicenciasHistorialModel
from Models.IntranetTipoRevisionModel import IntranetTipoRevisionModel
from Models.IntranetRevisionesModel import IntranetRevisionesModel
from Models.IntranetVersionesLicenciasModel import IntranetVersionesLicenciasModel
from Models.IntranetGscEstadosModel import IntranetGscEstados
from Models.IntranetGscSistemasAfectadosModel import IntranetGscSistemasAfectados
from Models.IntranetGscModulosModel import IntranetGscModulos
from Models.IntranetGscTiposEvidenciaModel import IntranetGscTiposEvidencia
from Models.IntranetGscOrigenesPlataformaModel import IntranetGscOrigenesPlataforma
from Models.IntranetGscFuentesSeguridadModel import IntranetGscFuentesSeguridad
from Models.IntranetGscImpactosModel import IntranetGscImpactos
from Models.IntranetGscRiesgosModel import IntranetGscRiesgos
from Models.IntranetTipoMonedaModel import IntranetTipoMoneda
from Models.IntranetGscRegistrosModel import IntranetGscRegistros
from Models.IntranetGscRegistrosSistemasModel import IntranetGscRegistrosSistemas
from Models.IntranetGscEvidenciasModel import IntranetGscEvidencias
from Models.IntranetGscEvidenciasTicketModel import IntranetGscEvidenciasTicket
from Models.IntranetGscEvidenciasCorreoModel import IntranetGscEvidenciasCorreo
from Models.IntranetGscEvidenciasAlertaModel import IntranetGscEvidenciasAlerta
from Models.IntranetGscEvidenciasCapturaModel import IntranetGscEvidenciasCaptura
from Models.IntranetGscEvidenciasOtroModel import IntranetGscEvidenciasOtro
from Models.IntranetGscRegistrosSeguridadModel import IntranetGscRegistrosSeguridad
from Models.IntranetGscRegistrosDisponibilidadModel import IntranetGscRegistrosDisponibilidad
from Models.IntranetGscRegistrosMantenimientoModel import IntranetGscRegistrosMantenimiento
from Models.IntranetGscRegistrosDisasterRecoveryModel import IntranetGscRegistrosDisasterRecovery
from Models.IntranetGscResultadosModel import IntranetGscResultados

import hashlib

class Querys:

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.query_params = dict()
        self.colombia_tz = pytz.timezone('America/Bogota')
    
    def _get_fecha_colombia(self):
        """
        Obtiene la fecha y hora actual en zona horaria de Colombia.
        Retorna datetime naive (sin timezone info) para compatibilidad con BD.
        """
        return datetime.now(self.colombia_tz).replace(tzinfo=None)
    
    def _convertir_fecha_a_colombia(self, fecha_str):
        """
        Convierte una fecha string del frontend a zona horaria de Colombia.
        Si la fecha viene en UTC o sin timezone, la ajusta a Colombia.
        Si ya es datetime, lo convierte a Colombia.
        
        Args:
            fecha_str: String de fecha en formato ISO o datetime object
        
        Returns:
            datetime naive en zona horaria Colombia o None si es inválido
        """
        if not fecha_str:
            return None
        
        try:
            # Si ya es un datetime object
            if isinstance(fecha_str, datetime):
                fecha = fecha_str
            else:
                # Parsear string ISO (puede tener 'Z' o offset)
                fecha_str = str(fecha_str).replace('Z', '+00:00')
                fecha = datetime.fromisoformat(fecha_str)
            
            # Si tiene timezone info, convertir a Colombia
            if fecha.tzinfo is not None:
                fecha_colombia = fecha.astimezone(self.colombia_tz)
                return fecha_colombia.replace(tzinfo=None)
            else:
                # Si no tiene timezone, asumir que ya es hora de Colombia
                return fecha
                
        except Exception as e:
            print(f"⚠️ Error convirtiendo fecha {fecha_str}: {e}")
            return None

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
    def generar_hash_contenido(self, subject, body_preview, from_email, received_date=None):
        """Genera un hash del contenido del correo para detectar cambios"""
        fecha_str = received_date.isoformat() if hasattr(received_date, 'isoformat') else str(received_date or '')
        contenido = f"{subject}{body_preview}{from_email}{fecha_str}"
        return hashlib.sha256(contenido.encode()).hexdigest()
    
    # Query para obtener un correo por su message_id de Microsoft
    def obtener_correo_por_message_id(self, message_id):
        """Obtiene un correo por su message_id de Microsoft (case-sensitive)"""
        try:
            correo = self.db.query(CorreosMicrosoftModel).filter(
                collate(CorreosMicrosoftModel.message_id, 'Latin1_General_CS_AS') == message_id
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
                correo_data.get('from_email', ''),
                correo_data.get('received_date')
            )
            correo_data['hash_contenido'] = hash_contenido
            
            nuevo_correo = CorreosMicrosoftModel(correo_data)
            self.db.add(nuevo_correo)
            self.db.commit()
            self.db.refresh(nuevo_correo)
            
            return nuevo_correo.to_dict()

        except IntegrityError:
            # El message_id ya existe (duplicado por sync concurrente o respuesta ya registrada)
            self.db.rollback()
            message_id = correo_data.get('message_id')
            print(f"⚠️ Correo con message_id '{message_id}' ya existe en BD, omitiendo inserción.")
            correo_existente = self.db.query(CorreosMicrosoftModel).filter(
                CorreosMicrosoftModel.message_id == message_id
            ).first()
            return correo_existente.to_dict() if correo_existente else None

        except Exception as e:
            self.db.rollback()
            print(f"Error insertando correo: {e}")
            return None
    
    # Query para actualizar un correo existente
    def actualizar_correo(self, message_id, datos_actualizacion):
        """Actualiza un correo existente (case-sensitive)"""
        try:
            correo = self.db.query(CorreosMicrosoftModel).filter(
                collate(CorreosMicrosoftModel.message_id, 'Latin1_General_CS_AS') == message_id
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
        Obtiene un correo por su message_id (case-sensitive)
        """
        try:
            correo = self.db.query(CorreosMicrosoftModel).filter(
                collate(CorreosMicrosoftModel.message_id, 'Latin1_General_CS_AS') == message_id
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
        Registra una respuesta entrante en el historial del ticket.
        NOTA: Ya no inserta en intranet_correos_microsoft para evitar duplicados de message_id.
        El correo aparecerá en bandeja normalmente con estado=1 (insertado por sincronizar_correos_inteligente).
        """
        # Solo loggear — la inserción real la hace insertar_correo en el flujo principal
        print(f"📥 Respuesta entrante registrada para ticket {respuesta_data.get('ticket_id')} | subject: {respuesta_data.get('subject', '')}")
        return True

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

    # Query para obtener observación de mantenimiento de un mes
    def obtener_observacion_mes_mantenimiento(self, anio, mes):
        try:
            observacion = self.db.query(IntranetObservacionesMantenimientoModel).filter(
                IntranetObservacionesMantenimientoModel.anio == anio,
                IntranetObservacionesMantenimientoModel.mes == mes,
                IntranetObservacionesMantenimientoModel.estado == 1
            ).first()
            return observacion.to_dict() if observacion else None
        except Exception as e:
            print(f"Error obteniendo observación de mantenimiento: {e}")
            return None

    # Query para guardar o actualizar observación de mantenimiento de un mes
    def guardar_observacion_mes_mantenimiento(self, anio, mes, observaciones):
        try:
            observacion_existente = self.db.query(IntranetObservacionesMantenimientoModel).filter(
                IntranetObservacionesMantenimientoModel.anio == anio,
                IntranetObservacionesMantenimientoModel.mes == mes,
                IntranetObservacionesMantenimientoModel.estado == 1
            ).first()

            if observacion_existente:
                observacion_existente.observaciones = observaciones
                self.db.commit()
                return observacion_existente.to_dict()
            else:
                nueva = IntranetObservacionesMantenimientoModel({
                    'anio': anio,
                    'mes': mes,
                    'observaciones': observaciones,
                    'estado': 1
                })
                self.db.add(nueva)
                self.db.commit()
                self.db.refresh(nueva)
                return nueva.to_dict()

        except Exception as e:
            self.db.rollback()
            print(f"Error guardando observación de mantenimiento: {e}")
            raise CustomException(f"Error guardando observación de mantenimiento: {str(e)}")

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

    # Query para obtener análisis de causas de mantenimiento de un año
    def obtener_analisis_causas_mantenimiento(self, anio):
        try:
            registros = self.db.query(IntranetCausasMantenimientoModel).filter(
                IntranetCausasMantenimientoModel.anio == anio,
                IntranetCausasMantenimientoModel.estado == 1
            ).order_by(
                IntranetCausasMantenimientoModel.mes.asc()
            ).all()
            return [r.to_dict() for r in registros]
        except Exception as e:
            print(f"Error obteniendo análisis de causas de mantenimiento: {e}")
            return []

    # Query para verificar si ya existe un análisis de mantenimiento para año+mes
    def verificar_analisis_mantenimiento_existe(self, anio, mes):
        try:
            existe = self.db.query(IntranetCausasMantenimientoModel).filter(
                IntranetCausasMantenimientoModel.anio == anio,
                IntranetCausasMantenimientoModel.mes == mes,
                IntranetCausasMantenimientoModel.estado == 1
            ).first()
            return existe is not None
        except Exception as e:
            print(f"Error verificando existencia de análisis de mantenimiento: {e}")
            return False

    # Query para guardar o actualizar análisis de causas de mantenimiento
    def guardar_analisis_causas_mantenimiento(self, id_analisis, anio, mes, analisis, acciones, responsable, fecha_compromiso, seguimiento):
        try:
            if id_analisis:
                registro = self.db.query(IntranetCausasMantenimientoModel).filter(
                    IntranetCausasMantenimientoModel.id == id_analisis,
                    IntranetCausasMantenimientoModel.estado == 1
                ).first()

                if not registro:
                    raise CustomException("Análisis de mantenimiento no encontrado")

                registro.analisis         = analisis
                registro.acciones         = acciones
                registro.responsable      = responsable
                registro.fecha_compromiso = fecha_compromiso
                registro.seguimiento      = seguimiento
                self.db.commit()
                return registro.to_dict()
            else:
                if self.verificar_analisis_mantenimiento_existe(anio, mes):
                    raise CustomException("Ya existe un análisis para este mes y año de mantenimiento")

                nuevo = IntranetCausasMantenimientoModel()
                nuevo.anio             = anio
                nuevo.mes              = mes
                nuevo.analisis         = analisis
                nuevo.acciones         = acciones
                nuevo.responsable      = responsable
                nuevo.fecha_compromiso = fecha_compromiso
                nuevo.seguimiento      = seguimiento
                nuevo.estado           = 1
                self.db.add(nuevo)
                self.db.commit()
                self.db.refresh(nuevo)
                return nuevo.to_dict()

        except Exception as e:
            self.db.rollback()
            print(f"Error guardando análisis de causas de mantenimiento: {e}")
            raise CustomException(f"Error guardando análisis de mantenimiento: {str(e)}")

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
                    'tipoMonedaId': row.tipo_moneda_id,
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
                    'tipoMonedaId': row.tipo_moneda_id,
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
                'tipoMonedaId': result.tipo_moneda_id,
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
                
            if 'tipoMonedaId' in data:
                if licencia.tipo_moneda_id != data['tipoMonedaId']:
                    moneda_anterior = self.db.query(IntranetTipoMoneda).filter(
                        IntranetTipoMoneda.id == licencia.tipo_moneda_id
                    ).first()
                    moneda_nueva = self.db.query(IntranetTipoMoneda).filter(
                        IntranetTipoMoneda.id == data['tipoMonedaId']
                    ).first()
                    cambios_detectados['Tipo de Moneda'] = {
                        'anterior': moneda_anterior.nombre if moneda_anterior else str(licencia.tipo_moneda_id),
                        'nuevo': moneda_nueva.nombre if moneda_nueva else str(data['tipoMonedaId'])
                    }
                licencia.tipo_moneda_id = data['tipoMonedaId']
                
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

    def obtener_tipos_moneda(self):
        """Obtiene todos los tipos de moneda activos"""
        
        try:
            tipos = self.db.query(IntranetTipoMoneda).filter(
                IntranetTipoMoneda.activo == True
            ).order_by(IntranetTipoMoneda.codigo.asc()).all()
            
            return [t.to_dict() for t in tipos]
            
        except Exception as e:
            print(f"Error obteniendo tipos de moneda: {e}")
            raise CustomException(f"Error obteniendo tipos de moneda: {str(e)}")

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

    # Querys para Gestión de Seguridad y Continuidad (GSC)
    def obtener_estados_gsc(self):
        """
        Obtiene todos los estados disponibles para el módulo GSC
        """
        try:
            estados = self.db.query(IntranetGscEstados).filter(
                IntranetGscEstados.activo == True
            ).all()
            
            return [{'id': estado.id, 'nombre': estado.nombre} for estado in estados]
            
        except Exception as e:
            print(f"Error obteniendo estados GSC: {e}")
            return []

    def obtener_sistemas_afectados_gsc(self):
        """
        Obtiene todos los sistemas afectados disponibles para el módulo GSC
        """
        try:
            sistemas = self.db.query(IntranetGscSistemasAfectados).filter(
                IntranetGscSistemasAfectados.activo == True
            ).all()
            
            return [{'id': sistema.id, 'nombre': sistema.nombre, 'descripcion': sistema.descripcion} for sistema in sistemas]
            
        except Exception as e:
            print(f"Error obteniendo sistemas afectados GSC: {e}")
            return []

    def obtener_modulos_gsc(self):
        """
        Obtiene todos los módulos disponibles para el módulo GSC ordenados por campo 'orden'
        """
        try:
            modulos = self.db.query(IntranetGscModulos).filter(
                IntranetGscModulos.activo == True
            ).order_by(IntranetGscModulos.orden).all()
            
            return [{'id': modulo.id, 'codigo': modulo.codigo, 'nombre': modulo.nombre, 'descripcion': modulo.descripcion, 'color_clase': modulo.color_clase, 'orden': modulo.orden} for modulo in modulos]
            
        except Exception as e:
            print(f"Error obteniendo módulos GSC: {e}")
            return []
    def obtener_tipos_evidencia_gsc(self):
        """
        Obtiene todos los tipos de evidencia disponibles para el módulo GSC
        """
        try:
            tipos = self.db.query(IntranetGscTiposEvidencia).filter(
                IntranetGscTiposEvidencia.activo == True
            ).all()
            
            return [tipo.to_dict() for tipo in tipos]
            
        except Exception as e:
            print(f"Error obteniendo tipos de evidencia GSC: {e}")
            return []

    def obtener_origenes_plataforma_gsc(self):
        """
        Obtiene todos los orígenes de plataforma disponibles para alertas en el módulo GSC
        """
        try:
            origenes = self.db.query(IntranetGscOrigenesPlataforma).filter(
                IntranetGscOrigenesPlataforma.activo == True
            ).all()
            
            return [origen.to_dict() for origen in origenes]
            
        except Exception as e:
            print(f"Error obteniendo orígenes de plataforma GSC: {e}")
            return []

    def obtener_fuentes_seguridad_gsc(self):
        """
        Obtiene todas las fuentes de seguridad disponibles para el módulo SEG
        """
        try:
            fuentes = self.db.query(IntranetGscFuentesSeguridad).filter(
                IntranetGscFuentesSeguridad.activo == True
            ).all()
            
            return [fuente.to_dict() for fuente in fuentes]
            
        except Exception as e:
            print(f"Error obteniendo fuentes de seguridad GSC: {e}")
            return []

    def obtener_impactos_gsc(self):
        """
        Obtiene todos los niveles de impacto disponibles para el módulo SEG ordenados por campo 'orden'
        """
        try:
            impactos = self.db.query(IntranetGscImpactos).filter(
                IntranetGscImpactos.activo == True
            ).order_by(IntranetGscImpactos.orden).all()
            
            return [impacto.to_dict() for impacto in impactos]
            
        except Exception as e:
            print(f"Error obteniendo impactos GSC: {e}")
            return []

    def obtener_riesgos_gsc(self):
        """
        Obtiene todos los niveles de riesgo disponibles para el módulo MNT ordenados por campo 'orden'
        """
        try:
            riesgos = self.db.query(IntranetGscRiesgos).filter(
                IntranetGscRiesgos.activo == True
            ).order_by(IntranetGscRiesgos.orden).all()
            
            return [riesgo.to_dict() for riesgo in riesgos]
            
        except Exception as e:
            print(f"Error obteniendo riesgos GSC: {e}")
            return []

    # ========================================
    # MÉTODOS CRUD PARA REGISTROS GSC
    # ========================================

    def crear_registro_gsc_completo(self, data: dict):
        """
        Crea un registro GSC completo con todas sus secciones:
        1. Información general (registro principal)
        2. Sistemas afectados (relación muchos a muchos)
        3. Evidencias (con sus datos específicos según tipo)
        4. Datos específicos del módulo (SEG/DISP/MNT/DR)
        5. Resultados iniciales (bitácora - opcional)
        
        Parámetros:
            data: Diccionario con la estructura completa del registro
                - resultados_iniciales: Lista de strings con textos de resultados (opcional)
        
        Retorna:
            dict: {'success': bool, 'id_registro': int, 'message': str}
        """
        try:
            # Convertir array de correos CC a JSON string si es necesario
            correos_cc_json = None
            if data.get('correos_cc'):
            
                if isinstance(data['correos_cc'], list):
                    correos_cc_json = json.dumps(data['correos_cc'])
                else:
                    correos_cc_json = data['correos_cc']  # Ya es string
            
            # PASO 1: Crear registro principal
            fecha_actual = self._get_fecha_colombia()
            
            registro_data = {
                'id_modulo': data.get('id_modulo'),
                'resumen': data.get('resumen'),
                'descripcion': data.get('descripcion'),
                'id_estado': data.get('id_estado'),
                'notificar_gerencia': data.get('notificar_gerencia', False),
                'enviar_contactos_empresa': data.get('enviar_contactos_empresa', False),
                'correos_cc': correos_cc_json,  # JSON string con array de correos
                'usuario_creacion': data.get('usuario_creacion'),
                'fecha_creacion': fecha_actual,
                'fecha_actualizacion': fecha_actual
            }
            
            # Asignar fecha según el estado seleccionado (zona horaria Colombia)
            estado_id = data.get('id_estado')
            
            if estado_id == 1:  # Abierto
                registro_data['fecha_abierto'] = fecha_actual
            elif estado_id == 2:  # En análisis
                registro_data['fecha_en_analisis'] = fecha_actual
            elif estado_id == 3:  # Mitigado
                registro_data['fecha_mitigado'] = fecha_actual
            elif estado_id == 4:  # Cerrado
                registro_data['fecha_cerrado'] = fecha_actual
            
            nuevo_registro = IntranetGscRegistros(registro_data)
            self.db.add(nuevo_registro)
            self.db.flush()  # Para obtener el ID sin hacer commit
            
            id_registro = nuevo_registro.id
            
            # PASO 2: Asociar sistemas afectados
            sistemas_afectados = data.get('sistemas_afectados', [])
            
            for id_sistema in sistemas_afectados:
                sistema_rel = IntranetGscRegistrosSistemas({
                    'id_registro': id_registro,
                    'id_sistema': id_sistema
                })
                self.db.add(sistema_rel)
                self.db.flush()

            # PASO 3: Crear evidencias con sus datos específicos
            evidencias = data.get('evidencias', [])
            for evidencia_data in evidencias:
                # Crear evidencia base
                evidencia = IntranetGscEvidencias({
                    'id_registro': id_registro,
                    'id_tipo_evidencia': evidencia_data.get('id_tipo_evidencia'),
                    'observacion': evidencia_data.get('observacion'),
                    'fecha_evidencia': self._convertir_fecha_a_colombia(evidencia_data.get('fecha_evidencia'))
                })
                self.db.add(evidencia)
                self.db.flush()
                
                id_evidencia = evidencia.id
                tipo_evidencia = evidencia_data.get('id_tipo_evidencia')
                datos_especificos = evidencia_data.get('datos_especificos', {})
                
                # Crear datos específicos según el tipo de evidencia
                if tipo_evidencia == 1:  # Ticket
                    ticket = IntranetGscEvidenciasTicket({
                        'id_evidencia': id_evidencia,
                        'numero_ticket': datos_especificos.get('numero_ticket'),
                        'plataforma': datos_especificos.get('plataforma'),
                        'url_ticket': datos_especificos.get('url_ticket')
                    })
                    self.db.add(ticket)
                    
                elif tipo_evidencia == 2:  # Correo
                    correo = IntranetGscEvidenciasCorreo({
                        'id_evidencia': id_evidencia,
                        'asunto': datos_especificos.get('asunto'),
                        'remitente': datos_especificos.get('remitente'),
                        'destinatarios': datos_especificos.get('destinatarios'),
                        'fecha_envio': self._convertir_fecha_a_colombia(datos_especificos.get('fecha_envio'))
                    })
                    self.db.add(correo)
                    
                elif tipo_evidencia == 3:  # Alerta
                    alerta = IntranetGscEvidenciasAlerta({
                        'id_evidencia': id_evidencia,
                        'id_origen_plataforma': datos_especificos.get('id_origen_plataforma'),
                        'nombre_alerta': datos_especificos.get('nombre_alerta'),
                        'severidad': datos_especificos.get('severidad'),
                        'fecha_alerta': self._convertir_fecha_a_colombia(datos_especificos.get('fecha_alerta')),
                        'codigo_alerta': datos_especificos.get('codigo_alerta')
                    })
                    self.db.add(alerta)
                    
                elif tipo_evidencia == 4:  # Captura
                    # Guardar archivo físico
                    file_handler = FileHandler()
                    base64_data = datos_especificos.get('archivo_base64')
                    nombre_original = datos_especificos.get('nombre_archivo', 'captura.png')
                    
                    if base64_data:
                        archivo_info = file_handler.guardar_imagen_base64(base64_data, nombre_original)
                        captura = IntranetGscEvidenciasCaptura({
                            'id_evidencia': id_evidencia,
                            'nombre_archivo': archivo_info['nombre_archivo'],
                            'ruta_archivo': archivo_info['ruta_relativa'],
                            'archivo_base64': None,  # No guardar base64 en BD
                            'tipo_mime': archivo_info['tipo_mime'],
                            'tamano_bytes': archivo_info['tamano_bytes']
                        })
                        self.db.add(captura)
                    
                elif tipo_evidencia == 5:  # Otro
                    otro = IntranetGscEvidenciasOtro({
                        'id_evidencia': id_evidencia,
                        'descripcion_tipo': datos_especificos.get('descripcion_tipo'),
                        'detalles': datos_especificos.get('detalles'),
                        'referencia': datos_especificos.get('referencia')
                    })
                    self.db.add(otro)
            
            # PASO 4: Crear datos específicos del módulo
            datos_modulo = data.get('datos_modulo', {})
            codigo_modulo = self._obtener_codigo_modulo(data.get('id_modulo'))
            
            if codigo_modulo == 'SEG':  # Seguridad
                seguridad = IntranetGscRegistrosSeguridad({
                    'id_registro': id_registro,
                    'fecha_hora_incidente': self._convertir_fecha_a_colombia(datos_modulo.get('fecha_hora_incidente')),
                    'id_fuente_seguridad': datos_modulo.get('id_fuente_seguridad'),
                    'tipo_amenaza': datos_modulo.get('tipo_amenaza'),
                    'id_impacto': datos_modulo.get('id_impacto'),
                    'responsable_tic': datos_modulo.get('responsable_tic'),
                    'acciones_tomadas': datos_modulo.get('acciones_tomadas')
                })
                self.db.add(seguridad)
                
            elif codigo_modulo == 'DISP':  # Disponibilidad
                disponibilidad = IntranetGscRegistrosDisponibilidad({
                    'id_registro': id_registro,
                    'servicio_afectado': datos_modulo.get('servicio_afectado'),
                    'tipo_evento': datos_modulo.get('tipo_evento'),
                    'tiempo_indisponible_min': datos_modulo.get('tiempo_indisponible_min', 0),
                    'sla_afectado': datos_modulo.get('sla_afectado', False),
                    'acciones': datos_modulo.get('acciones'),
                    'causa_raiz': datos_modulo.get('causa_raiz')
                })
                self.db.add(disponibilidad)
                
            elif codigo_modulo == 'MNT':  # Mantenimiento
                mantenimiento = IntranetGscRegistrosMantenimiento({
                    'id_registro': id_registro,
                    'area': datos_modulo.get('area'),
                    'tipo_mantenimiento': datos_modulo.get('tipo_mantenimiento'),
                    'descripcion': datos_modulo.get('descripcion'),
                    'fecha_inicio': self._convertir_fecha_a_colombia(datos_modulo.get('fecha_inicio')),
                    'fecha_fin': self._convertir_fecha_a_colombia(datos_modulo.get('fecha_fin')),
                    'requiere_parada': datos_modulo.get('requiere_parada', False),
                    'id_riesgo': datos_modulo.get('id_riesgo'),
                    'sistemas_componentes': datos_modulo.get('sistemas_componentes'),
                    'responsable_ejecucion': datos_modulo.get('responsable_ejecucion')
                })
                self.db.add(mantenimiento)
                
            elif codigo_modulo == 'DR':  # Disaster Recovery
                dr = IntranetGscRegistrosDisasterRecovery({
                    'id_registro': id_registro,
                    'escenario': datos_modulo.get('escenario'),
                    'fecha_inicio': self._convertir_fecha_a_colombia(datos_modulo.get('fecha_inicio')),
                    'fecha_fin': self._convertir_fecha_a_colombia(datos_modulo.get('fecha_fin')),
                    'objetivo': datos_modulo.get('objetivo'),
                    'resultado': datos_modulo.get('resultado'),
                    'hallazgos': datos_modulo.get('hallazgos'),
                    'lecciones_aprendidas': datos_modulo.get('lecciones_aprendidas'),
                    'rto_objetivo': datos_modulo.get('rto_objetivo'),
                    'rto_real': datos_modulo.get('rto_real'),
                    'rpo_objetivo': datos_modulo.get('rpo_objetivo'),
                    'rpo_real': datos_modulo.get('rpo_real')
                })
                self.db.add(dr)
            
            # PASO 5: Crear resultados iniciales si se proporcionan
            resultados_iniciales = data.get('resultados_iniciales', [])
            if resultados_iniciales and len(resultados_iniciales) > 0:
                for texto_resultado in resultados_iniciales:
                    if texto_resultado and texto_resultado.strip():
                        resultado = IntranetGscResultados({
                            'id_registro': id_registro,
                            'texto': texto_resultado.strip(),
                            'created_at': self._get_fecha_colombia(),
                            'activo': True
                        })
                        self.db.add(resultado)
            
            # PASO 6: Commit de toda la transacción
            self.db.commit()
            
            # PASO 7: Enviar notificación si al menos uno de los checkbox está activo
            if data.get('notificar_gerencia', False) or data.get('enviar_contactos_empresa', False):
                self.enviar_notificacion_gerencia_gsc(id_registro)
            
            return {
                'success': True,
                'id_registro': id_registro,
                'message': 'Registro GSC creado exitosamente'
            }
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando registro GSC completo: {e}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'id_registro': None,
                'message': f'Error creando registro: {str(e)}'
            }

    def _obtener_codigo_modulo(self, id_modulo: int):
        """
        Método auxiliar para obtener el código del módulo por su ID
        """
        try:
            modulo = self.db.query(IntranetGscModulos).filter(
                IntranetGscModulos.id == id_modulo
            ).first()
            
            return modulo.codigo if modulo else None
            
        except Exception as e:
            print(f"Error obteniendo código de módulo: {e}")
            return None

    def obtener_registro_gsc_completo(self, id_registro: int):
        """
        Obtiene un registro GSC completo con todas sus relaciones
        """
        try:
            # Obtener registro principal
            registro = self.db.query(IntranetGscRegistros).filter(
                IntranetGscRegistros.id == id_registro,
                IntranetGscRegistros.activo == True
            ).first()
            
            if not registro:
                return None
            
            resultado = registro.to_dict()
            
            # Obtener sistemas afectados con JOIN
            sistemas = self.db.query(
                IntranetGscSistemasAfectados
            ).join(
                IntranetGscRegistrosSistemas,
                IntranetGscRegistrosSistemas.id_sistema == IntranetGscSistemasAfectados.id
            ).filter(
                IntranetGscRegistrosSistemas.id_registro == id_registro
            ).all()
            
            resultado['sistemas_afectados'] = [{'id': s.id, 'nombre': s.nombre} for s in sistemas]
            
            # Obtener evidencias
            evidencias = self.db.query(IntranetGscEvidencias).filter(
                IntranetGscEvidencias.id_registro == id_registro,
                IntranetGscEvidencias.activo == True
            ).all()
            
            evidencias_list = []
            for evidencia in evidencias:
                ev_dict = evidencia.to_dict()
                
                # Obtener datos específicos según el tipo
                tipo = evidencia.id_tipo_evidencia
                
                if tipo == 1:  # Ticket
                    ticket = self.db.query(IntranetGscEvidenciasTicket).filter(
                        IntranetGscEvidenciasTicket.id_evidencia == evidencia.id
                    ).first()
                    if ticket:
                        ev_dict['datos_especificos'] = ticket.to_dict()
                        
                elif tipo == 2:  # Correo
                    correo = self.db.query(IntranetGscEvidenciasCorreo).filter(
                        IntranetGscEvidenciasCorreo.id_evidencia == evidencia.id
                    ).first()
                    if correo:
                        ev_dict['datos_especificos'] = correo.to_dict()
                        
                elif tipo == 3:  # Alerta
                    alerta = self.db.query(IntranetGscEvidenciasAlerta).filter(
                        IntranetGscEvidenciasAlerta.id_evidencia == evidencia.id
                    ).first()
                    if alerta:
                        ev_dict['datos_especificos'] = alerta.to_dict()
                        
                elif tipo == 4:  # Captura
                    captura = self.db.query(IntranetGscEvidenciasCaptura).filter(
                        IntranetGscEvidenciasCaptura.id_evidencia == evidencia.id
                    ).first()
                    if captura:
                        captura_dict = captura.to_dict()
                        # En lugar de enviar base64, enviar URL para obtener la imagen
                        captura_dict['imagen_url'] = f"/gestion-continuidad/obtener_imagen_evidencia/{captura.nombre_archivo}"
                        captura_dict['archivo_base64'] = None  # No enviar base64
                        ev_dict['datos_especificos'] = captura_dict
                        
                elif tipo == 5:  # Otro
                    otro = self.db.query(IntranetGscEvidenciasOtro).filter(
                        IntranetGscEvidenciasOtro.id_evidencia == evidencia.id
                    ).first()
                    if otro:
                        ev_dict['datos_especificos'] = otro.to_dict()
                
                evidencias_list.append(ev_dict)
            
            resultado['evidencias'] = evidencias_list
            
            # Obtener datos específicos del módulo
            codigo_modulo = self._obtener_codigo_modulo(registro.id_modulo)
            
            if codigo_modulo == 'SEG':
                seguridad = self.db.query(IntranetGscRegistrosSeguridad).filter(
                    IntranetGscRegistrosSeguridad.id_registro == id_registro
                ).first()
                if seguridad:
                    resultado['datos_modulo'] = seguridad.to_dict()
                    
            elif codigo_modulo == 'DISP':
                disponibilidad = self.db.query(IntranetGscRegistrosDisponibilidad).filter(
                    IntranetGscRegistrosDisponibilidad.id_registro == id_registro
                ).first()
                if disponibilidad:
                    resultado['datos_modulo'] = disponibilidad.to_dict()
                    
            elif codigo_modulo == 'MNT':
                mantenimiento = self.db.query(IntranetGscRegistrosMantenimiento).filter(
                    IntranetGscRegistrosMantenimiento.id_registro == id_registro
                ).first()
                if mantenimiento:
                    resultado['datos_modulo'] = mantenimiento.to_dict()
                    
            elif codigo_modulo == 'DR':
                dr = self.db.query(IntranetGscRegistrosDisasterRecovery).filter(
                    IntranetGscRegistrosDisasterRecovery.id_registro == id_registro
                ).first()
                if dr:
                    resultado['datos_modulo'] = dr.to_dict()
            
            # Parsear correos_cc de JSON a array
            if resultado.get('correos_cc'):
                try:
                
                    resultado['correos_cc'] = json.loads(resultado['correos_cc'])
                except:
                    resultado['correos_cc'] = []
            else:
                resultado['correos_cc'] = []
            
            return resultado
            
        except Exception as e:
            print(f"Error obteniendo registro GSC completo: {e}")
            import traceback
            traceback.print_exc()
            return None

    def listar_registros_gsc(self, filtros: dict = None):
        """
        Lista registros GSC con filtros opcionales
        
        Parámetros:
            filtros: {
                'id_modulo': int,
                'id_estado': int,
                'q': str (búsqueda),
                'fecha_desde': datetime,
                'fecha_hasta': datetime,
                'limite': int,
                'offset': int
            }
        
        Retorna:
            {
                'registros': [...],
                'total': int,
                'pagina': int,
                'total_paginas': int
            }
        """
        try:
            query = self.db.query(IntranetGscRegistros).filter(
                IntranetGscRegistros.activo == True
            )
            
            if filtros:
                if filtros.get('id_modulo'):
                    query = query.filter(IntranetGscRegistros.id_modulo == filtros['id_modulo'])
                
                if filtros.get('id_estado'):
                    query = query.filter(IntranetGscRegistros.id_estado == filtros['id_estado'])
                
                if filtros.get('fecha_desde'):
                    query = query.filter(IntranetGscRegistros.fecha_creacion >= filtros['fecha_desde'])
                
                if filtros.get('fecha_hasta'):
                    query = query.filter(IntranetGscRegistros.fecha_creacion <= filtros['fecha_hasta'])
                
                # Búsqueda por texto en resumen y descripción
                if filtros.get('q'):
                    busqueda = f"%{filtros['q']}%"
                    query = query.filter(
                        or_(
                            IntranetGscRegistros.resumen.like(busqueda),
                            IntranetGscRegistros.descripcion.like(busqueda)
                        )
                    )
            
            # Contar total antes de paginar
            total = query.count()
            
            # Ordenar por id descendente
            query = query.order_by(IntranetGscRegistros.id.desc())
            
            # Paginación
            limite = filtros.get('limite', 10) if filtros else 10
            offset = filtros.get('offset', 0) if filtros else 0
            
            query = query.limit(limite).offset(offset)
            
            registros = query.all()
            
            # Enriquecer cada registro con información adicional
            resultado = []
            for registro in registros:
                reg_dict = registro.to_dict()
                
                # Contar evidencias activas
                cantidad_evidencias = self.db.query(IntranetGscEvidencias).filter(
                    IntranetGscEvidencias.id_registro == registro.id,
                    IntranetGscEvidencias.activo == True
                ).count()
                
                reg_dict['tiene_evidencias'] = cantidad_evidencias > 0
                reg_dict['cantidad_evidencias'] = cantidad_evidencias
                
                # Contar sistemas afectados
                cantidad_sistemas = self.db.query(IntranetGscRegistrosSistemas).filter(
                    IntranetGscRegistrosSistemas.id_registro == registro.id
                ).count()
                
                reg_dict['cantidad_sistemas_afectados'] = cantidad_sistemas
                
                resultado.append(reg_dict)
            
            # Calcular metadatos de paginación
            pagina_actual = (offset // limite) + 1 if limite > 0 else 1
            total_paginas = (total + limite - 1) // limite if limite > 0 else 1
            
            return {
                'registros': resultado,
                'total': total,
                'pagina': pagina_actual,
                'total_paginas': total_paginas,
                'por_pagina': limite
            }
            
        except Exception as e:
            print(f"Error listando registros GSC: {e}")
            import traceback
            traceback.print_exc()
            return {
                'registros': [],
                'total': 0,
                'pagina': 1,
                'total_paginas': 1,
                'por_pagina': 10
            }

    def obtener_contadores_gsc(self, filtros: dict):
        """
        Obtiene contadores de registros por estado
        Si se especifica id_modulo, filtra por ese módulo
        Si no, retorna totales globales de todos los módulos
        Usado para KPIs en el dashboard
        """
        try:
            id_modulo = filtros.get('id_modulo') if filtros else None

            # Obtener IDs de estados
            estados = self.db.query(IntranetGscEstados).filter(
                IntranetGscEstados.activo == True).all()
            estado_map = {e.nombre: e.id for e in estados}

            # Query base para contar
            query_base = self.db.query(
                func.count(IntranetGscRegistros.id)).filter(
                    IntranetGscRegistros.activo == True)
            
            # Si hay filtro de módulo, aplicarlo
            if id_modulo:
                # Total de registros del módulo
                total = query_base.filter(
                    IntranetGscRegistros.id_modulo == id_modulo
                ).scalar() or 0

                # Contar por cada estado
                abiertos = query_base.filter(
                    IntranetGscRegistros.id_modulo == id_modulo,
                    IntranetGscRegistros.id_estado == estado_map.get('Abierto')
                ).scalar() or 0

                en_analisis = query_base.filter(
                    IntranetGscRegistros.id_modulo == id_modulo,
                    IntranetGscRegistros.id_estado == estado_map.get('En análisis')
                ).scalar() or 0

                mitigados = query_base.filter(
                    IntranetGscRegistros.id_modulo == id_modulo,
                    IntranetGscRegistros.id_estado == estado_map.get('Mitigado')
                ).scalar() or 0

                cerrados = query_base.filter(
                    IntranetGscRegistros.id_modulo == id_modulo,
                    IntranetGscRegistros.id_estado == estado_map.get('Cerrado')
                ).scalar() or 0
            else:
                # Totales globales de todos los módulos
                total = query_base.scalar() or 0

                abiertos = query_base.filter(
                    IntranetGscRegistros.id_estado == estado_map.get('Abierto')
                ).scalar() or 0

                en_analisis = query_base.filter(
                    IntranetGscRegistros.id_estado == estado_map.get('En análisis')
                ).scalar() or 0

                mitigados = query_base.filter(
                    IntranetGscRegistros.id_estado == estado_map.get('Mitigado')
                ).scalar() or 0

                cerrados = query_base.filter(
                    IntranetGscRegistros.id_estado == estado_map.get('Cerrado')
                ).scalar() or 0

            return {
                'total': total,
                'abiertos': abiertos,
                'en_analisis': en_analisis,
                'mitigados': mitigados,
                'cerrados': cerrados
            }

        except Exception as e:
            print(f"Error obteniendo contadores GSC: {e}")
            import traceback
            traceback.print_exc()
            return {
                'total': 0,
                'abiertos': 0,
                'en_analisis': 0,
                'mitigados': 0,
                'cerrados': 0
            }

    def actualizar_registro_gsc_completo(self, id_registro: int, data: dict):
        """
        Actualiza un registro GSC completo
        """
        try:
            # Obtener registro existente
            registro = self.db.query(IntranetGscRegistros).filter(
                IntranetGscRegistros.id == id_registro,
                IntranetGscRegistros.activo == True
            ).first()
            
            if not registro:
                return {'success': False, 'message': 'Registro no encontrado'}
            
            # Actualizar campos del registro principal
            if 'resumen' in data:
                registro.resumen = data['resumen']
            if 'descripcion' in data:
                registro.descripcion = data['descripcion']
            if 'id_estado' in data:
                registro.id_estado = data['id_estado']
                # Actualizar fecha según nuevo estado (zona horaria Colombia)
                estado_id = data['id_estado']
                fecha_actual = self._get_fecha_colombia()
                if estado_id == 1:
                    registro.fecha_abierto = fecha_actual
                elif estado_id == 2:
                    registro.fecha_en_analisis = fecha_actual
                elif estado_id == 3:
                    registro.fecha_mitigado = fecha_actual
                elif estado_id == 4:
                    registro.fecha_cerrado = fecha_actual
            if 'notificar_gerencia' in data:
                registro.notificar_gerencia = data['notificar_gerencia']
            if 'enviar_contactos_empresa' in data:
                registro.enviar_contactos_empresa = data['enviar_contactos_empresa']
            if 'correos_cc' in data:
                # Convertir array a JSON string si es necesario
            
                if isinstance(data['correos_cc'], list):
                    registro.correos_cc = json.dumps(data['correos_cc'])
                else:
                    registro.correos_cc = data['correos_cc']  # Ya es string
            
            registro.fecha_actualizacion = self._get_fecha_colombia()
            registro.usuario_actualizacion = data.get('usuario_actualizacion')
            
            # Actualizar sistemas afectados si se proporcionan
            if 'sistemas_afectados' in data:
                # Eliminar relaciones actuales
                self.db.query(IntranetGscRegistrosSistemas).filter(
                    IntranetGscRegistrosSistemas.id_registro == id_registro
                ).delete()
                
                # Crear nuevas relaciones
                for id_sistema in data['sistemas_afectados']:
                    sistema_rel = IntranetGscRegistrosSistemas({
                        'id_registro': id_registro,
                        'id_sistema': id_sistema
                    })
                    self.db.add(sistema_rel)
            
            # Actualizar evidencias si se proporcionan
            if 'evidencias' in data:
                # Desactivar todas las evidencias actuales (soft delete)
                evidencias_actuales = self.db.query(IntranetGscEvidencias).filter(
                    IntranetGscEvidencias.id_registro == id_registro,
                    IntranetGscEvidencias.activo == True
                ).all()
                
                for ev in evidencias_actuales:
                    ev.activo = False
                
                # Crear las nuevas evidencias
                evidencias = data.get('evidencias', [])
                for evidencia_data in evidencias:
                    
                    # Crear evidencia base
                    evidencia = IntranetGscEvidencias({
                        'id_registro': id_registro,
                        'id_tipo_evidencia': evidencia_data.get('id_tipo_evidencia'),
                        'observacion': evidencia_data.get('observacion'),
                        'fecha_evidencia': self._convertir_fecha_a_colombia(evidencia_data.get('fecha_evidencia'))
                    })
                    self.db.add(evidencia)
                    self.db.flush()
                    
                    id_evidencia = evidencia.id
                    tipo_evidencia = evidencia_data.get('id_tipo_evidencia')
                    datos_especificos = evidencia_data.get('datos_especificos', {})
                    
                    # Crear datos específicos según el tipo de evidencia
                    if tipo_evidencia == 1:  # Ticket
                        ticket = IntranetGscEvidenciasTicket({
                            'id_evidencia': id_evidencia,
                            'numero_ticket': datos_especificos.get('numero_ticket'),
                            'plataforma': datos_especificos.get('plataforma'),
                            'url_ticket': datos_especificos.get('url_ticket')
                        })
                        self.db.add(ticket)
                        
                    elif tipo_evidencia == 2:  # Correo
                        correo = IntranetGscEvidenciasCorreo({
                            'id_evidencia': id_evidencia,
                            'asunto': datos_especificos.get('asunto'),
                            'remitente': datos_especificos.get('remitente'),
                            'destinatarios': datos_especificos.get('destinatarios'),
                            'fecha_envio': self._convertir_fecha_a_colombia(datos_especificos.get('fecha_envio'))
                        })
                        self.db.add(correo)
                        
                    elif tipo_evidencia == 3:  # Alerta
                        alerta = IntranetGscEvidenciasAlerta({
                            'id_evidencia': id_evidencia,
                            'id_origen_plataforma': datos_especificos.get('id_origen_plataforma'),
                            'nombre_alerta': datos_especificos.get('nombre_alerta'),
                            'severidad': datos_especificos.get('severidad'),
                            'fecha_alerta': self._convertir_fecha_a_colombia(datos_especificos.get('fecha_alerta')),
                            'codigo_alerta': datos_especificos.get('codigo_alerta')
                        })
                        self.db.add(alerta)
                        
                    elif tipo_evidencia == 4:  # Captura
                        # Guardar archivo físico solo si es base64 nuevo
                        file_handler = FileHandler()
                        base64_data = datos_especificos.get('archivo_base64')
                        nombre_original = datos_especificos.get('nombre_archivo', 'captura.png')
                        
                        if base64_data:
                            # Detectar si es una URL (imagen existente) o base64 nuevo
                            if base64_data.startswith('http://') or base64_data.startswith('https://') or base64_data.startswith('/gestion-continuidad/'):
                                # Es una URL, extraer el nombre del archivo de la URL
                                # Formato: .../obtener_imagen_evidencia/captura_20260204_103045_123456.png
                                nombre_archivo = base64_data.split('/')[-1]
                                ruta_relativa = f'/gestion-continuidad/obtener_imagen_evidencia/{nombre_archivo}'
                                
                                # Obtener info del archivo existente
                                ruta_completa = file_handler.obtener_ruta_completa(nombre_archivo)
                                if ruta_completa.exists():
                                    tamano_bytes = ruta_completa.stat().st_size
                                else:
                                    tamano_bytes = 0
                                
                                captura = IntranetGscEvidenciasCaptura({
                                    'id_evidencia': id_evidencia,
                                    'nombre_archivo': nombre_archivo,
                                    'ruta_archivo': ruta_relativa,
                                    'archivo_base64': None,
                                    'tipo_mime': 'image/png',
                                    'tamano_bytes': tamano_bytes
                                })
                            else:
                                # Es base64 nuevo, guardarlo
                                archivo_info = file_handler.guardar_imagen_base64(base64_data, nombre_original)
                                captura = IntranetGscEvidenciasCaptura({
                                    'id_evidencia': id_evidencia,
                                    'nombre_archivo': archivo_info['nombre_archivo'],
                                    'ruta_archivo': archivo_info['ruta_relativa'],
                                    'archivo_base64': None,  # No guardar base64 en BD
                                    'tipo_mime': archivo_info['tipo_mime'],
                                    'tamano_bytes': archivo_info['tamano_bytes']
                                })
                            
                            self.db.add(captura)
                        
                    elif tipo_evidencia == 5:  # Otro
                        otro = IntranetGscEvidenciasOtro({
                            'id_evidencia': id_evidencia,
                            'descripcion_tipo': datos_especificos.get('descripcion_tipo'),
                            'detalles': datos_especificos.get('detalles'),
                            'referencia': datos_especificos.get('referencia')
                        })
                        self.db.add(otro)
            
            # Actualizar datos del módulo si se proporcionan
            if 'datos_modulo' in data:
                codigo_modulo = self._obtener_codigo_modulo(registro.id_modulo)
                datos_modulo = data['datos_modulo']
                
                if codigo_modulo == 'SEG':
                    seguridad = self.db.query(IntranetGscRegistrosSeguridad).filter(
                        IntranetGscRegistrosSeguridad.id_registro == id_registro
                    ).first()
                    if seguridad:
                        for key, value in datos_modulo.items():
                            # Convertir fechas a zona horaria Colombia
                            if key == 'fecha_hora_incidente' and value:
                                value = self._convertir_fecha_a_colombia(value)
                            setattr(seguridad, key, value)
                            
                elif codigo_modulo == 'DISP':
                    disponibilidad = self.db.query(IntranetGscRegistrosDisponibilidad).filter(
                        IntranetGscRegistrosDisponibilidad.id_registro == id_registro
                    ).first()
                    if disponibilidad:
                        for key, value in datos_modulo.items():
                            setattr(disponibilidad, key, value)
                            
                elif codigo_modulo == 'MNT':
                    mantenimiento = self.db.query(IntranetGscRegistrosMantenimiento).filter(
                        IntranetGscRegistrosMantenimiento.id_registro == id_registro
                    ).first()
                    if mantenimiento:
                        for key, value in datos_modulo.items():
                            # Convertir fechas a zona horaria Colombia
                            if key in ['fecha_inicio', 'fecha_fin'] and value:
                                value = self._convertir_fecha_a_colombia(value)
                            setattr(mantenimiento, key, value)
                            
                elif codigo_modulo == 'DR':
                    dr = self.db.query(IntranetGscRegistrosDisasterRecovery).filter(
                        IntranetGscRegistrosDisasterRecovery.id_registro == id_registro
                    ).first()
                    if dr:
                        for key, value in datos_modulo.items():
                            # Convertir fechas a zona horaria Colombia
                            if key in ['fecha_inicio', 'fecha_fin'] and value:
                                value = self._convertir_fecha_a_colombia(value)
                            setattr(dr, key, value)
            
            self.db.commit()
            
            # Enviar notificación si al menos uno de los checkbox está activo
            if registro.notificar_gerencia or registro.enviar_contactos_empresa:
                self.enviar_notificacion_gerencia_gsc(id_registro)
            
            return {
                'success': True,
                'message': 'Registro actualizado exitosamente'
            }
            
        except Exception as e:
            self.db.rollback()
            print(f"Error actualizando registro GSC: {e}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'message': f'Error actualizando registro: {str(e)}'
            }

    def eliminar_registro_gsc(self, id_registro: int):
        """
        Elimina (desactiva) un registro GSC
        """
        try:
            registro = self.db.query(IntranetGscRegistros).filter(
                IntranetGscRegistros.id == id_registro
            ).first()
            
            if not registro:
                return {'success': False, 'message': 'Registro no encontrado'}
            
            registro.activo = False
            registro.fecha_actualizacion = self._get_fecha_colombia()
            
            self.db.commit()
            
            return {
                'success': True,
                'message': 'Registro eliminado exitosamente'
            }
            
        except Exception as e:
            self.db.rollback()
            print(f"Error eliminando registro GSC: {e}")
            return {
                'success': False,
                'message': f'Error eliminando registro: {str(e)}'
            }

    # ========================================
    # MÉTODOS CRUD PARA RESULTADOS GSC
    # ========================================

    def crear_resultado_gsc(self, data: dict):
        """
        Crea un nuevo resultado (entrada de bitácora) para un registro GSC
        """
        try:
            # Validar que el registro existe y está activo
            registro = self.db.query(IntranetGscRegistros).filter(
                IntranetGscRegistros.id == data.get('id_registro'),
                IntranetGscRegistros.activo == True
            ).first()
            
            if not registro:
                return {'success': False, 'message': 'Registro no encontrado'}
            
            # Crear el resultado con fecha en zona horaria de Colombia
            resultado = IntranetGscResultados({
                'id_registro': data.get('id_registro'),
                'texto': data.get('texto'),
                'created_at': self._get_fecha_colombia(),
                'activo': True
            })
            
            self.db.add(resultado)
            self.db.commit()
            self.db.refresh(resultado)
            
            return {
                'success': True,
                'message': 'Resultado creado exitosamente',
                'id_resultado': resultado.id,
                'data': resultado.to_dict()
            }
            
        except Exception as e:
            self.db.rollback()
            print(f"Error creando resultado GSC: {e}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'message': f'Error creando resultado: {str(e)}'
            }

    def listar_resultados_gsc(self, id_registro: int):
        """
        Lista todos los resultados (entradas de bitácora) de un registro GSC
        Ordenados por fecha de creación descendente (más reciente primero)
        """
        try:
            resultados = self.db.query(IntranetGscResultados).filter(
                IntranetGscResultados.id_registro == id_registro,
                IntranetGscResultados.activo == True
            ).order_by(IntranetGscResultados.created_at.desc()).all()
            
            return [resultado.to_dict() for resultado in resultados]
            
        except Exception as e:
            print(f"Error listando resultados GSC: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _generar_html_notificacion_gsc(self, registro_data: dict, modulo_data: dict, codigo_modulo: str, id_registro: int) -> str:
        """
        Genera el HTML para el correo de notificación a gerencia
        """
        # Obtener nombres desde IDs
        modulo_nombre = ""
        estado_nombre = ""
        
        if 'id_modulo' in registro_data:
            modulo = self.db.query(IntranetGscModulos).filter(
                IntranetGscModulos.id == registro_data['id_modulo']
            ).first()
            if modulo:
                modulo_nombre = modulo.nombre
        
        if 'id_estado' in registro_data:
            estado = self.db.query(IntranetGscEstados).filter(
                IntranetGscEstados.id == registro_data['id_estado']
            ).first()
            if estado:
                estado_nombre = estado.nombre
        
        # Obtener sistemas afectados
        sistemas_query = self.db.query(
            IntranetGscSistemasAfectados
        ).join(
            IntranetGscRegistrosSistemas,
            IntranetGscRegistrosSistemas.id_sistema == IntranetGscSistemasAfectados.id
        ).filter(
            IntranetGscRegistrosSistemas.id_registro == id_registro
        ).all()
        
        sistemas_html = ""
        if sistemas_query:
            sistemas_lista = ", ".join([s.nombre for s in sistemas_query])
            sistemas_html = f"""
            <tr>
                <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold; width: 200px; vertical-align: top;">Sistemas Afectados:</td>
                <td style="padding: 12px; border-bottom: 1px solid #dee2e6;">
                    <div style="background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; border-radius: 4px;">
                        {sistemas_lista}
                    </div>
                </td>
            </tr>
            """
        
        # Verde pastel oscuro: #5a9e7a
        color_principal = "#5a9e7a"
        color_borde = "#4a8a6a"
        
        # Obtener resultados (bitácora) del registro
        resultados = self.listar_resultados_gsc(id_registro)
        resultados_html = ""
        
        if resultados and len(resultados) > 0:
            # Generar filas de resultados
            filas_resultados = ""
            for idx, resultado in enumerate(resultados):
                bg_color = "#ffffff" if idx % 2 == 0 else "#f8f9fa"
                fecha_formateada = resultado['created_at'][:16].replace('T', ' ')  # Formato: "2026-02-09 10:30"
                
                filas_resultados += f"""
                <tr>
                    <td style="padding: 12px; background-color: {bg_color}; font-weight: bold; width: 200px; vertical-align: top; border-bottom: 1px solid #e0e0e0;">{fecha_formateada}</td>
                    <td style="padding: 12px; background-color: {bg_color}; border-bottom: 1px solid #e0e0e0; white-space: pre-wrap;">{resultado['texto']}</td>
                </tr>
                """
            
            resultados_html = f"""
            <div style="background-color: #e8f5e9; padding: 20px; border-radius: 8px; border: 2px solid #81c784; margin-top: 20px;">
                <h3 style="color: #2e7d32; border-bottom: 3px solid #388e3c; padding-bottom: 10px; margin-top: 0;">📝 Resultados (Bitácora)</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    {filas_resultados}
                </table>
            </div>
            """
        
        # Generar sección específica del módulo
        seccion_modulo = ""
        
        if codigo_modulo == 'SEG':
            fuente_nombre = ""
            impacto_nombre = ""
            
            if modulo_data.get('id_fuente_seguridad'):
                fuente = self.db.query(IntranetGscFuentesSeguridad).filter(
                    IntranetGscFuentesSeguridad.id == modulo_data['id_fuente_seguridad']
                ).first()
                if fuente:
                    fuente_nombre = fuente.nombre
            
            if modulo_data.get('id_impacto'):
                impacto = self.db.query(IntranetGscImpactos).filter(
                    IntranetGscImpactos.id == modulo_data['id_impacto']
                ).first()
                if impacto:
                    impacto_nombre = impacto.nombre
            
            fecha_incidente = modulo_data.get('fecha_hora_incidente', '')
            if isinstance(fecha_incidente, datetime):
                fecha_incidente = fecha_incidente.strftime('%Y-%m-%d %H:%M:%S')
            
            seccion_modulo = f"""
            <div style="background-color: #e8f4f8; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 2px solid {color_borde};">
                <h3 style="color: {color_principal}; border-bottom: 3px solid {color_borde}; padding-bottom: 10px; margin-top: 0;">📊 Información de Seguridad</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; width: 200px; border-bottom: 1px solid #e0e0e0;">Fecha/Hora Incidente:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{fecha_incidente}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Fuente:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{fuente_nombre}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold;">Tipo Amenaza:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('tipo_amenaza', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Impacto:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">
                            <span style="background-color: #ffeaa7; padding: 4px 12px; border-radius: 12px; font-weight: bold;">{impacto_nombre}</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold;">Responsable TIC:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('responsable_tic', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold; vertical-align: top;">Acciones Tomadas:</td>
                        <td style="padding: 12px; background-color: #f8f9fa;">{modulo_data.get('acciones_tomadas', '')}</td>
                    </tr>
                </table>
            </div>
            """
        
        elif codigo_modulo == 'DISP':
            seccion_modulo = f"""
            <div style="background-color: #e8f4f8; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 2px solid {color_borde};">
                <h3 style="color: {color_principal}; border-bottom: 3px solid {color_borde}; padding-bottom: 10px; margin-top: 0;">📈 Información de Disponibilidad</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; width: 200px; border-bottom: 1px solid #e0e0e0;">Servicio Afectado:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('servicio_afectado', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Tipo Evento:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('tipo_evento', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold;">Tiempo Indisponible:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">
                            <span style="background-color: #ff7675; color: white; padding: 4px 12px; border-radius: 12px; font-weight: bold;">{modulo_data.get('tiempo_indisponible_min', 0)} minutos</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">SLA Afectado:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">
                            <span style="background-color: {'#ff7675' if modulo_data.get('sla_afectado') else '#55efc4'}; color: white; padding: 4px 12px; border-radius: 12px; font-weight: bold;">{'Sí' if modulo_data.get('sla_afectado') else 'No'}</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; vertical-align: top;">Acciones:</td>
                        <td style="padding: 12px; background-color: #ffffff;">{modulo_data.get('acciones', '')}</td>
                    </tr>
                </table>
            </div>
            """
        
        elif codigo_modulo == 'MNT':
            riesgo_nombre = ""
            if modulo_data.get('id_riesgo'):
                riesgo = self.db.query(IntranetGscRiesgos).filter(
                    IntranetGscRiesgos.id == modulo_data['id_riesgo']
                ).first()
                if riesgo:
                    riesgo_nombre = riesgo.nombre
            
            fecha_inicio = modulo_data.get('fecha_inicio', '')
            fecha_fin = modulo_data.get('fecha_fin', '')
            if isinstance(fecha_inicio, datetime):
                fecha_inicio = fecha_inicio.strftime('%Y-%m-%d %H:%M:%S')
            if isinstance(fecha_fin, datetime):
                fecha_fin = fecha_fin.strftime('%Y-%m-%d %H:%M:%S')
            
            seccion_modulo = f"""
            <div style="background-color: #e8f4f8; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 2px solid {color_borde};">
                <h3 style="color: {color_principal}; border-bottom: 3px solid {color_borde}; padding-bottom: 10px; margin-top: 0;">🔧 Información de Mantenimiento</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; width: 200px; border-bottom: 1px solid #e0e0e0;">Área:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('area', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Tipo Mantenimiento:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('tipo_mantenimiento', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; vertical-align: top;">Descripción:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('descripcion', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Fecha Inicio:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{fecha_inicio}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold;">Fecha Fin:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{fecha_fin}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Interrupción del Servicio:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">
                            <span style="background-color: {'#ff7675' if modulo_data.get('requiere_parada') else '#55efc4'}; color: white; padding: 4px 12px; border-radius: 12px; font-weight: bold;">{'Sí' if modulo_data.get('requiere_parada') else 'No'}</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold;">Riesgo:</td>
                        <td style="padding: 12px; background-color: #ffffff;">
                            <span style="background-color: #ffeaa7; padding: 4px 12px; border-radius: 12px; font-weight: bold;">{riesgo_nombre}</span>
                        </td>
                    </tr>
                </table>
            </div>
            """
        
        elif codigo_modulo == 'DR':
            fecha_inicio = modulo_data.get('fecha_inicio', '')
            fecha_fin = modulo_data.get('fecha_fin', '')
            if isinstance(fecha_inicio, datetime):
                fecha_inicio = fecha_inicio.strftime('%Y-%m-%d %H:%M:%S')
            if isinstance(fecha_fin, datetime):
                fecha_fin = fecha_fin.strftime('%Y-%m-%d %H:%M:%S')
            
            seccion_modulo = f"""
            <div style="background-color: #e8f4f8; padding: 20px; border-radius: 8px; margin-bottom: 25px; border: 2px solid {color_borde};">
                <h3 style="color: {color_principal}; border-bottom: 3px solid {color_borde}; padding-bottom: 10px; margin-top: 0;">🔄 Información de Disaster Recovery</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; width: 200px; border-bottom: 1px solid #e0e0e0;">Escenario:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('escenario', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Fecha Inicio:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{fecha_inicio}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold;">Fecha Fin:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{fecha_fin}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold; vertical-align: top;">Objetivo:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('objetivo', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; vertical-align: top;">Resultado:</td>
                        <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('resultado', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold; vertical-align: top;">Hallazgos:</td>
                        <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">{modulo_data.get('hallazgos', '')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; background-color: #ffffff; font-weight: bold; vertical-align: top;">Lecciones Aprendidas:</td>
                        <td style="padding: 12px; background-color: #ffffff;">{modulo_data.get('lecciones_aprendidas', '')}</td>
                    </tr>
                </table>
            </div>
            """
        
        # HTML completo del correo
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{ 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                    line-height: 1.6; 
                    color: #333; 
                    background-color: #f4f4f4;
                    margin: 0;
                    padding: 0;
                }}
                .container {{ 
                    max-width: 800px; 
                    margin: 20px auto; 
                    background-color: #ffffff;
                    border: 1px solid #dddddd;
                    border-radius: 8px;
                    overflow: hidden;
                }}
                .header {{ 
                    background-color: {color_principal};
                    color: white; 
                    padding: 20px 20px; 
                    text-align: left;
                }}
                .header img {{
                    max-width: 150px;
                    background-color: white;
                    padding: 8px;
                    border-radius: 8px;
                    margin-bottom: 10px;
                }}
                .header h1 {{
                    margin: 8px 0;
                    font-size: 22px;
                    font-weight: 600;
                    color: white;
                }}
                .header p {{
                    margin: 5px 0 0 0;
                    font-size: 15px;
                    color: white;
                }}
                .content {{ 
                    background-color: #ffffff; 
                    padding: 35px;
                }}
                .info-general {{
                    background-color: #fff0f0;
                    padding: 20px;
                    border-radius: 8px;
                    border: 2px solid #ffcccb;
                }}
                .info-general h3 {{
                    color: #e74c3c;
                    border-bottom: 3px solid #c0392b;
                    padding-bottom: 10px;
                    margin-top: 0;
                }}
                .footer {{ 
                    background-color: #2c3e50; 
                    color: #ecf0f1;
                    padding: 20px; 
                    text-align: center; 
                    font-size: 13px;
                }}
                .footer p {{
                    margin: 5px 0;
                }}
                .footer strong {{
                    color: {color_principal};
                    font-size: 16px;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>Notificación de Gestión de Continuidad</h1>
                    <p>Módulo: <strong>{modulo_nombre}</strong></p>
                </div>
                
                <div class="content">
                    {seccion_modulo}
                    
                    <div class="info-general">
                        <h3>📋 Información General</h3>
                        <table style="width: 100%; border-collapse: collapse;">
                            <tr>
                                <td style="padding: 12px; background-color: #ffffff; font-weight: bold; width: 200px; border-bottom: 1px solid #e0e0e0;">Resumen:</td>
                                <td style="padding: 12px; background-color: #ffffff; border-bottom: 1px solid #e0e0e0;">{registro_data.get('resumen', '')}</td>
                            </tr>
                            <tr>
                                <td style="padding: 12px; background-color: #f8f9fa; font-weight: bold;">Estado:</td>
                                <td style="padding: 12px; background-color: #f8f9fa; border-bottom: 1px solid #e0e0e0;">
                                    <span style="background-color: {color_principal}; color: white; padding: 4px 12px; border-radius: 12px; font-weight: bold;">{estado_nombre}</span>
                                </td>
                            </tr>
                            {sistemas_html}
                        </table>
                    </div>
                    
                    {resultados_html}
                </div>
                
                <div class="footer">
                    <p>⚠️ Este mensaje fue generado automáticamente. Por favor, no responda a este correo.</p>
                    <p><strong>AVANTIKA</strong> | Gestión TIC | Macroproceso de Tecnología de la Información y Comunicaciones.</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html

    def enviar_notificacion_gerencia_gsc(self, id_registro: int):
        """
        Envía notificación por correo según los checkbox activos:
        - Ambos activos: envía 1 correo a gerencia + contactos empresa
        - Solo notificar_gerencia: envía solo a gerencia
        - Solo enviar_contactos_empresa: envía solo a contactos empresa
        - Ninguno activo: no envía nada
        """
        try:
            # Obtener el registro completo
            registro = self.db.query(IntranetGscRegistros).filter(
                IntranetGscRegistros.id == id_registro
            ).first()
            
            if not registro:
                print(f"Registro {id_registro} no encontrado")
                return False
            
            # Verificar si debe notificar
            if not registro.notificar_gerencia and not registro.enviar_contactos_empresa:
                return True  # No es error, simplemente no debe notificar
            
            # Obtener código del módulo
            codigo_modulo = self._obtener_codigo_modulo(registro.id_modulo)
            
            # Preparar datos del registro
            registro_data = {
                'id_modulo': registro.id_modulo,
                'resumen': registro.resumen,
                'descripcion': registro.descripcion,
                'id_estado': registro.id_estado
            }
            
            # Obtener datos específicos del módulo
            modulo_data = {}
            
            if codigo_modulo == 'SEG':
                datos = self.db.query(IntranetGscRegistrosSeguridad).filter(
                    IntranetGscRegistrosSeguridad.id_registro == id_registro
                ).first()
                if datos:
                    modulo_data = {
                        'fecha_hora_incidente': datos.fecha_hora_incidente,
                        'id_fuente_seguridad': datos.id_fuente_seguridad,
                        'tipo_amenaza': datos.tipo_amenaza,
                        'id_impacto': datos.id_impacto,
                        'responsable_tic': datos.responsable_tic,
                        'acciones_tomadas': datos.acciones_tomadas
                    }
            
            elif codigo_modulo == 'DISP':
                datos = self.db.query(IntranetGscRegistrosDisponibilidad).filter(
                    IntranetGscRegistrosDisponibilidad.id_registro == id_registro
                ).first()
                if datos:
                    modulo_data = {
                        'servicio_afectado': datos.servicio_afectado,
                        'tipo_evento': datos.tipo_evento,
                        'tiempo_indisponible_min': datos.tiempo_indisponible_min,
                        'sla_afectado': datos.sla_afectado,
                        'acciones': datos.acciones,
                        'causa_raiz': datos.causa_raiz
                    }
            
            elif codigo_modulo == 'MNT':
                datos = self.db.query(IntranetGscRegistrosMantenimiento).filter(
                    IntranetGscRegistrosMantenimiento.id_registro == id_registro
                ).first()
                if datos:
                    modulo_data = {
                        'area': datos.area,
                        'tipo_mantenimiento': datos.tipo_mantenimiento,
                        'descripcion': datos.descripcion,
                        'fecha_inicio': datos.fecha_inicio,
                        'fecha_fin': datos.fecha_fin,
                        'requiere_parada': datos.requiere_parada,
                        'id_riesgo': datos.id_riesgo
                    }
            
            elif codigo_modulo == 'DR':
                datos = self.db.query(IntranetGscRegistrosDisasterRecovery).filter(
                    IntranetGscRegistrosDisasterRecovery.id_registro == id_registro
                ).first()
                if datos:
                    modulo_data = {
                        'escenario': datos.escenario,
                        'fecha_inicio': datos.fecha_inicio,
                        'fecha_fin': datos.fecha_fin,
                        'objetivo': datos.objetivo,
                        'resultado': datos.resultado,
                        'hallazgos': datos.hallazgos,
                        'lecciones_aprendidas': datos.lecciones_aprendidas
                    }
            
            # Generar HTML del correo (pasar id_registro para obtener sistemas afectados)
            html_body = self._generar_html_notificacion_gsc(registro_data, modulo_data, codigo_modulo, id_registro)
            
            # Configurar destinatarios según los checkbox activos
            subject = f"Notificación GSC - {codigo_modulo} - Registro #{id_registro}"
            mail_sender = "tic@avantika.com.co"
            
            # Determinar destinatarios según checkbox activos
            to_email = None
            cc_emails = []
            
            # CASO 1: Ambos checkbox activos - Enviar 1 correo con todos los destinatarios
            if registro.notificar_gerencia and registro.enviar_contactos_empresa:
                to_email = "gerencia@avantika.com.co"
                cc_emails = ["auxiliartic@avantika.com.co", "tic@avantika.com.co"]
                
                # Agregar correos de contactos empresa
                if registro.correos_cc:
                    try:
                        correos_adicionales = json.loads(registro.correos_cc)
                        if isinstance(correos_adicionales, list):
                            cc_emails.extend(correos_adicionales)
                    except Exception as e:
                        print(f"Error parseando correos CC: {e}")
 
            # CASO 2: Solo notificar a gerencia
            elif registro.notificar_gerencia:
                to_email = "gerencia@avantika.com.co"
                cc_emails = ["auxiliartic@avantika.com.co", "tic@avantika.com.co"]
  
            # CASO 3: Solo enviar a contactos empresa
            elif registro.enviar_contactos_empresa:
                if registro.correos_cc:
                    try:
                        correos_adicionales = json.loads(registro.correos_cc)
                        if isinstance(correos_adicionales, list) and len(correos_adicionales) > 0:
                            to_email = correos_adicionales[0]  # Primer correo como destinatario principal
                            cc_emails = correos_adicionales[1:] if len(correos_adicionales) > 1 else []
                            cc_emails.extend(["auxiliartic@avantika.com.co", "tic@avantika.com.co", "sistemas@avantika.com.co"])  # Agregar TIC en CC
                        else:
                            print(f"No hay correos de contactos empresa para enviar")
                            return True
                    except Exception as e:
                        print(f"Error parseando correos CC: {e}")
                        return False
                else:
                    print(f"No hay correos de contactos empresa para enviar")
                    return True
  
            # Si no hay destinatario, no enviar
            if not to_email:
                print(f"No hay destinatarios configurados")
                return True
            
            # Ruta del logo - Usar logo.png en la raíz del proyecto
            logo_path = Path(__file__).parent.parent / "logo.png"

            self.tools.send_email_individual(
                to_email=to_email,
                cc_emails=cc_emails,
                subject=subject,
                body=html_body,
                logo_path=str(logo_path) if logo_path.exists() else None,
                mail_sender=mail_sender,
                db=self.db  # Pasar sesión de BD para usar Graph API
            )
            
            print(f"Notificación enviada exitosamente para registro {id_registro}")
            return True
            
        except Exception as e:
            print(f"Error enviando notificación: {e}")
            traceback.print_exc()
            return False

    # ── Indicador Mantenimiento Preventivo TIC ───────────────────────────────
    def obtener_indicadores_mantenimiento(self, anio):
        """
        Indicador: Cumplimiento del programa de mantenimiento preventivo TIC.
        Solo se evalúan los meses de Marzo (3) y Septiembre (9).

        Para cada mes calcula:
          - total_actividades   : actividades ligadas a OTs preventivas (tipo_mantenimiento=1)
                                  cuya fecha_programacion_desde cae en ese mes/año.
          - actividades_oportunas: subset de las anteriores cuyo created_at (fecha de
                                  ejecución) cae DENTRO del rango [fecha_programacion_desde,
                                  fecha_programacion_hasta] de la OT.
          - porcentaje           : oportunas / total * 100
          - porcentaje_acumulado : acumulado año sobre los dos meses evaluados
        """
        try:
            MESES = [
                {'numero': 3, 'nombre': 'Marzo'},
                {'numero': 9, 'nombre': 'Septiembre'},
            ]

            porcentaje_meta = 90.0

            OT  = IntranetOrdenesTrabajo
            ACT = IntranetActividadesOrdenesTrabajo
            A   = IntranetActivos

            GRUPOS_ACTIVOS = ('13', '16')

            indicadores = []
            total_actividades_acum = 0
            total_oportunas_acum   = 0

            for mes_info in MESES:
                mes_num = mes_info['numero']

                # Filtro compartido: OTs preventivas del mes/año en grupos TIC
                filtro_base = [
                    OT.tipo_mantenimiento == 1,
                    func.extract('year',  OT.fecha_programacion_desde) == anio,
                    func.extract('month', OT.fecha_programacion_desde) == mes_num,
                    A.grupo.in_(GRUPOS_ACTIVOS),
                    A.retirado == 0,
                ]

                # ── Total de actividades a realizar ──────────────────────────
                # Contamos OTs (no actividades): cada OT = 1 actividad de
                # mantenimiento. Incluye las pendientes (estado_ot=1) y las
                # completadas (estado_ot=3), ya que todas "deben realizarse".
                total_q = (
                    self.db.query(func.count(OT.id))
                    .join(A, A.id == OT.activo_id)
                    .filter(*filtro_base)
                    .scalar()
                ) or 0

                # ── Actividades ejecutadas oportunamente ─────────────────────
                # Contamos registros en actividades cuyo created_at cae dentro
                # del rango programado de la OT (ejecutadas a tiempo).
                oportunas_q = (
                    self.db.query(func.count(ACT.id))
                    .join(OT, ACT.orden_trabajo_id == OT.id)
                    .join(A,  A.id == OT.activo_id)
                    .filter(
                        *filtro_base,
                        cast(ACT.created_at, Date) >= OT.fecha_programacion_desde,
                        cast(ACT.created_at, Date) <= OT.fecha_programacion_hasta,
                    )
                    .scalar()
                ) or 0

                total_actividades_acum += total_q
                total_oportunas_acum   += oportunas_q

                porcentaje = (
                    round(oportunas_q / total_q * 100, 2) if total_q > 0 else 0
                )
                porcentaje_acumulado = (
                    round(total_oportunas_acum / total_actividades_acum * 100, 2)
                    if total_actividades_acum > 0 else 0
                )

                indicadores.append({
                    'mes':                  mes_info['nombre'],
                    'mes_numero':           mes_num,
                    'total_actividades':    total_q,
                    'actividades_oportunas': oportunas_q,
                    'porcentaje':           f"{porcentaje}%",
                    'porcentaje_acumulado': f"{porcentaje_acumulado}%",
                    'porcentaje_meta':      porcentaje_meta,
                })

            porcentaje_global = (
                round(total_oportunas_acum / total_actividades_acum * 100, 2)
                if total_actividades_acum > 0 else 0
            )

            return {
                'anio': anio,
                'indicadores': indicadores,
                'totales': {
                    'total_actividades':    total_actividades_acum,
                    'actividades_oportunas': total_oportunas_acum,
                    'porcentaje_global':    f"{porcentaje_global}%",
                    'porcentaje_meta':      porcentaje_meta,
                },
            }

        except Exception as e:
            print(f"Error obteniendo indicadores de mantenimiento: {e}")
            raise CustomException(f"Error obteniendo indicadores de mantenimiento: {str(e)}")

    # ══════════════════════════════════════════════════════════════════════════
    # CONTINGENCIA
    # ══════════════════════════════════════════════════════════════════════════

    # Importaciones diferidas para evitar circularidad al inicio del módulo
    def _cont_models(self):
        from Models.IntranetContingenciaTiposEventoModel          import IntranetContingenciaTiposEvento
        from Models.IntranetContingenciaPrioridadesModel          import IntranetContingenciaPrioridades
        from Models.IntranetContingenciaEstadosEventoModel        import IntranetContingenciaEstadosEvento
        from Models.IntranetContingenciaEstadosAccionModel        import IntranetContingenciaEstadosAccion
        from Models.IntranetContingenciaEstadosDocumentoModel     import IntranetContingenciaEstadosDocumento
        from Models.IntranetContingenciaTiposBitacoraModel        import IntranetContingenciaTiposBitacora
        from Models.IntranetContingenciaResultadosRecuperacionModel import IntranetContingenciaResultadosRecuperacion
        from Models.IntranetContingenciaEventosModel              import IntranetContingenciaEventos
        from Models.IntranetContingenciaAccionesModel             import IntranetContingenciaAcciones
        from Models.IntranetContingenciaDocumentosModel           import IntranetContingenciaDocumentos
        from Models.IntranetContingenciaBitacorasModel            import IntranetContingenciaBitacoras
        from Models.IntranetContingenciaRecuperacionModel         import IntranetContingenciaRecuperacion
        return {
            'TiposEvento':            IntranetContingenciaTiposEvento,
            'Prioridades':            IntranetContingenciaPrioridades,
            'EstadosEvento':          IntranetContingenciaEstadosEvento,
            'EstadosAccion':          IntranetContingenciaEstadosAccion,
            'EstadosDocumento':       IntranetContingenciaEstadosDocumento,
            'TiposBitacora':          IntranetContingenciaTiposBitacora,
            'ResultadosRecuperacion': IntranetContingenciaResultadosRecuperacion,
            'Eventos':                IntranetContingenciaEventos,
            'Acciones':               IntranetContingenciaAcciones,
            'Documentos':             IntranetContingenciaDocumentos,
            'Bitacoras':              IntranetContingenciaBitacoras,
            'Recuperacion':           IntranetContingenciaRecuperacion,
        }

    # ── Catálogos ──────────────────────────────────────────────────────────────

    def cont_obtener_catalogos(self):
        m = self._cont_models()
        def _cat(model):
            return [r.to_dict() for r in
                    self.db.query(model).filter(model.activo == True).order_by(model.orden).all()]
        return {
            'tipos_evento':             _cat(m['TiposEvento']),
            'prioridades':              _cat(m['Prioridades']),
            'estados_evento':           _cat(m['EstadosEvento']),
            'estados_accion':           _cat(m['EstadosAccion']),
            'estados_documento':        _cat(m['EstadosDocumento']),
            'tipos_bitacora':           _cat(m['TiposBitacora']),
            'resultados_recuperacion':  _cat(m['ResultadosRecuperacion']),
        }

    def cont_id_tipo_bitacora(self, nombre: str) -> int:
        """Devuelve el id del tipo de bitácora por nombre (para uso interno)."""
        m = self._cont_models()
        row = self.db.query(m['TiposBitacora']).filter(
            m['TiposBitacora'].nombre == nombre,
            m['TiposBitacora'].activo == True,
        ).first()
        return row.id if row else 1

    # ── Contadores dashboard ───────────────────────────────────────────────────

    def cont_obtener_contadores(self):
        m = self._cont_models()
        E = m['Eventos']
        A = m['Acciones']
        D = m['Documentos']
        B = m['Bitacoras']

        # Buscamos el id del estado "Cerrado" dinámicamente
        cerrado = self.db.query(m['EstadosEvento']).filter(
            m['EstadosEvento'].nombre == 'Cerrado',
            m['EstadosEvento'].activo == True,
        ).first()
        completada = self.db.query(m['EstadosAccion']).filter(
            m['EstadosAccion'].nombre == 'Completada',
            m['EstadosAccion'].activo == True,
        ).first()

        total_eventos     = self.db.query(E).filter(E.activo == True).count()
        eventos_abiertos  = self.db.query(E).filter(
            E.activo == True,
            E.id_estado_evento != cerrado.id if cerrado else True,
        ).count()
        acciones_total     = self.db.query(A).filter(A.activo == True).count()
        acciones_completadas = self.db.query(A).filter(
            A.activo == True,
            A.id_estado_accion == completada.id if completada else -1,
        ).count()
        total_documentos  = self.db.query(D).filter(D.activo == True).count()
        total_bitacoras   = self.db.query(B).filter(B.activo == True).count()

        return {
            'total_eventos':          total_eventos,
            'eventos_abiertos':       eventos_abiertos,
            'acciones_total':         acciones_total,
            'acciones_completadas':   acciones_completadas,
            'total_documentos':       total_documentos,
            'total_bitacoras':        total_bitacoras,
        }

    # ── Eventos ────────────────────────────────────────────────────────────────

    def cont_contar_eventos_anio(self) -> int:
        m = self._cont_models()
        anio = self._get_fecha_colombia().year
        return self.db.query(m['Eventos']).filter(
            func.year(m['Eventos'].fecha_creacion) == anio
        ).count()

    def cont_listar_eventos(self, filtros: dict):
        m = self._cont_models()
        E  = m['Eventos']
        TE = m['TiposEvento']
        PR = m['Prioridades']
        EE = m['EstadosEvento']

        q = (self.db.query(E, TE.nombre, PR.nombre, EE.nombre)
             .join(TE, E.id_tipo_evento   == TE.id)
             .join(PR, E.id_prioridad     == PR.id)
             .join(EE, E.id_estado_evento == EE.id)
             .filter(E.activo == True))

        if filtros.get('id_tipo_evento'):
            q = q.filter(E.id_tipo_evento == filtros['id_tipo_evento'])
        if filtros.get('id_estado_evento'):
            q = q.filter(E.id_estado_evento == filtros['id_estado_evento'])
        if filtros.get('query'):
            term = f"%{filtros['query'].strip().lower()}%"
            q = q.filter(or_(
                E.codigo.ilike(term),
                E.titulo.ilike(term),
                E.area.ilike(term),
                E.responsable.ilike(term),
            ))

        q = q.order_by(E.fecha_creacion.desc())
        rows = q.all()

        resultado = []
        for evento, tipo_nombre, prioridad_nombre, estado_nombre in rows:
            d = evento.to_dict()
            d['tipo_evento']  = tipo_nombre
            d['prioridad']    = prioridad_nombre
            d['estado_evento'] = estado_nombre
            resultado.append(d)
        return resultado

    def cont_crear_evento(self, data: dict) -> dict:
        m = self._cont_models()
        evento = m['Eventos'](data)
        evento.fecha_creacion      = self._get_fecha_colombia()
        evento.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(evento)
        self.db.commit()
        self.db.refresh(evento)
        return evento.to_dict()

    def cont_obtener_evento(self, id_evento: int):
        m  = self._cont_models()
        E  = m['Eventos']
        TE = m['TiposEvento']
        PR = m['Prioridades']
        EE = m['EstadosEvento']

        row = (self.db.query(E, TE.nombre, PR.nombre, EE.nombre)
               .join(TE, E.id_tipo_evento   == TE.id)
               .join(PR, E.id_prioridad     == PR.id)
               .join(EE, E.id_estado_evento == EE.id)
               .filter(E.id == id_evento, E.activo == True)
               .first())

        if not row:
            return None
        evento, tipo_nombre, prioridad_nombre, estado_nombre = row
        d = evento.to_dict()
        d['tipo_evento']   = tipo_nombre
        d['prioridad']     = prioridad_nombre
        d['estado_evento'] = estado_nombre
        return d

    def cont_actualizar_evento(self, id_evento: int, data: dict):
        m = self._cont_models()
        evento = self.db.query(m['Eventos']).filter(
            m['Eventos'].id == id_evento, m['Eventos'].activo == True
        ).first()
        if not evento:
            return None

        campos = ['id_tipo_evento', 'id_prioridad', 'id_estado_evento', 'titulo',
                  'area', 'responsable', 'fecha_inicio', 'rto_objetivo', 'rpo_objetivo',
                  'impacto', 'causa', 'usuario_actualizacion']
        for campo in campos:
            if campo in data:
                setattr(evento, campo, data[campo])
        evento.fecha_actualizacion = self._get_fecha_colombia()
        self.db.commit()
        self.db.refresh(evento)
        return evento.to_dict()

    def cont_eliminar_evento(self, id_evento: int):
        m = self._cont_models()
        now = self._get_fecha_colombia()

        for model in [m['Acciones'], m['Documentos'], m['Bitacoras']]:
            (self.db.query(model)
             .filter(model.id_evento == id_evento)
             .update({'activo': False}, synchronize_session=False))

        rec = self.db.query(m['Recuperacion']).filter(
            m['Recuperacion'].id_evento == id_evento
        ).first()
        if rec:
            rec.activo = False

        evento = self.db.query(m['Eventos']).filter(m['Eventos'].id == id_evento).first()
        if evento:
            evento.activo             = False
            evento.fecha_actualizacion = now
        self.db.commit()

    # ── Acciones ───────────────────────────────────────────────────────────────

    def cont_contar_acciones_anio(self) -> int:
        m = self._cont_models()
        anio = self._get_fecha_colombia().year
        return self.db.query(m['Acciones']).filter(
            func.year(m['Acciones'].fecha_creacion) == anio
        ).count()

    def cont_crear_accion(self, data: dict) -> dict:
        m = self._cont_models()
        accion = m['Acciones'](data)
        accion.fecha_creacion      = self._get_fecha_colombia()
        accion.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(accion)
        self.db.commit()
        self.db.refresh(accion)
        return accion.to_dict()

    def cont_listar_acciones(self, id_evento: int):
        m  = self._cont_models()
        A  = m['Acciones']
        EA = m['EstadosAccion']

        rows = (self.db.query(A, EA.nombre)
                .join(EA, A.id_estado_accion == EA.id)
                .filter(A.id_evento == id_evento, A.activo == True)
                .order_by(A.fecha_creacion.desc())
                .all())

        resultado = []
        for accion, estado_nombre in rows:
            d = accion.to_dict()
            d['estado_accion'] = estado_nombre
            resultado.append(d)
        return resultado

    def cont_actualizar_accion(self, id_accion: int, data: dict):
        m = self._cont_models()
        accion = self.db.query(m['Acciones']).filter(
            m['Acciones'].id == id_accion, m['Acciones'].activo == True
        ).first()
        if not accion:
            return None

        campos = ['id_estado_accion', 'titulo', 'responsable', 'fecha_compromiso',
                  'evidencia', 'control_asociado', 'usuario_actualizacion']
        for campo in campos:
            if campo in data:
                setattr(accion, campo, data[campo])
        accion.fecha_actualizacion = self._get_fecha_colombia()
        self.db.commit()
        self.db.refresh(accion)
        return accion.to_dict()

    def cont_eliminar_accion(self, id_accion: int):
        m = self._cont_models()
        accion = self.db.query(m['Acciones']).filter(m['Acciones'].id == id_accion).first()
        if accion:
            accion.activo             = False
            accion.fecha_actualizacion = self._get_fecha_colombia()
            self.db.commit()

    # ── Recuperación ───────────────────────────────────────────────────────────

    def cont_upsert_recuperacion(self, data: dict) -> dict:
        m = self._cont_models()
        rec = self.db.query(m['Recuperacion']).filter(
            m['Recuperacion'].id_evento == data['id_evento']
        ).first()

        campos = ['id_resultado_recuperacion', 'tiempo_real', 'datos_recuperados',
                  'observaciones', 'servicio_alterno', 'integridad_documental',
                  'informe_final', 'lecciones_aprendidas']

        if rec:
            for campo in campos:
                if campo in data:
                    setattr(rec, campo, data[campo])
            if 'usuario_actualizacion' in data:
                rec.usuario_actualizacion = data['usuario_actualizacion']
            rec.fecha_actualizacion = self._get_fecha_colombia()
            rec.activo = True
        else:
            rec = m['Recuperacion'](data)
            rec.fecha_creacion      = self._get_fecha_colombia()
            rec.fecha_actualizacion = self._get_fecha_colombia()
            self.db.add(rec)

        self.db.commit()
        self.db.refresh(rec)
        return rec.to_dict()

    def cont_obtener_recuperacion(self, id_evento: int):
        m = self._cont_models()
        rec = self.db.query(m['Recuperacion']).filter(
            m['Recuperacion'].id_evento == id_evento,
            m['Recuperacion'].activo == True,
        ).first()
        return rec.to_dict() if rec else None

    # ── Documentos ─────────────────────────────────────────────────────────────

    def cont_contar_documentos_anio(self) -> int:
        m = self._cont_models()
        anio = self._get_fecha_colombia().year
        return self.db.query(m['Documentos']).filter(
            func.year(m['Documentos'].fecha_creacion) == anio
        ).count()

    def cont_crear_documento(self, data: dict) -> dict:
        m = self._cont_models()
        doc = m['Documentos'](data)
        doc.fecha_creacion      = self._get_fecha_colombia()
        doc.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(doc)
        self.db.commit()
        self.db.refresh(doc)
        return doc.to_dict()

    def cont_listar_documentos(self, id_evento: int):
        m  = self._cont_models()
        D  = m['Documentos']
        ED = m['EstadosDocumento']

        rows = (self.db.query(D, ED.nombre)
                .join(ED, D.id_estado_documento == ED.id)
                .filter(D.id_evento == id_evento, D.activo == True)
                .order_by(D.fecha_creacion.desc())
                .all())

        resultado = []
        for doc, estado_nombre in rows:
            d = doc.to_dict()
            d['estado_documento'] = estado_nombre
            resultado.append(d)
        return resultado

    def cont_eliminar_documento(self, id_documento: int):
        m = self._cont_models()
        doc = self.db.query(m['Documentos']).filter(m['Documentos'].id == id_documento).first()
        if doc:
            doc.activo             = False
            doc.fecha_actualizacion = self._get_fecha_colombia()
            self.db.commit()

    # ── Bitácoras ──────────────────────────────────────────────────────────────

    def cont_contar_bitacoras_evento(self, id_evento: int) -> int:
        m = self._cont_models()
        anio = self._get_fecha_colombia().year
        return self.db.query(m['Bitacoras']).filter(
            m['Bitacoras'].id_evento == id_evento,
            func.year(m['Bitacoras'].fecha_creacion) == anio,
        ).count()

    def cont_crear_bitacora(self, data: dict) -> dict:
        m = self._cont_models()
        if not data.get('hora_registro'):
            data['hora_registro'] = self._get_fecha_colombia().strftime('%d/%m/%Y %H:%M')
        bitacora = m['Bitacoras'](data)
        bitacora.fecha_creacion = self._get_fecha_colombia()
        self.db.add(bitacora)
        self.db.commit()
        self.db.refresh(bitacora)
        return bitacora.to_dict()

    def cont_listar_bitacoras(self, id_evento: int):
        m  = self._cont_models()
        B  = m['Bitacoras']
        TB = m['TiposBitacora']

        rows = (self.db.query(B, TB.nombre)
                .join(TB, B.id_tipo_bitacora == TB.id)
                .filter(B.id_evento == id_evento, B.activo == True)
                .order_by(B.fecha_creacion.desc())
                .all())

        resultado = []
        for bitacora, tipo_nombre in rows:
            d = bitacora.to_dict()
            d['tipo_bitacora'] = tipo_nombre
            resultado.append(d)
        return resultado

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # CCTV
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _cctv_models(self):
        from Models.IntranetCctvSedesModel              import IntranetCctvSedes
        from Models.IntranetCctvCamarasModel            import IntranetCctvCamaras
        from Models.IntranetCctvCargosModel             import IntranetCctvCargos
        from Models.IntranetCctvRegistrosCambiosModel   import IntranetCctvRegistrosCambios
        from Models.IntranetCctvRevisionesModel         import IntranetCctvRevisiones
        from Models.IntranetCctvIncidentesModel         import IntranetCctvIncidentes
        from Models.IntranetCctvEstadosCamaraModel      import IntranetCctvEstadosCamara
        from Models.IntranetCctvMetodosBackupModel      import IntranetCctvMetodosBackup
        from Models.IntranetCctvNivelesAccesoModel      import IntranetCctvNivelesAcceso
        from Models.IntranetCctvSeveridadesModel        import IntranetCctvSeveridades
        from Models.IntranetCctvEstadosIncidenteModel   import IntranetCctvEstadosIncidente
        return {
            'Sedes':             IntranetCctvSedes,
            'Camaras':           IntranetCctvCamaras,
            'Cargos':            IntranetCctvCargos,
            'RegistrosCambios':  IntranetCctvRegistrosCambios,
            'Revisiones':        IntranetCctvRevisiones,
            'Incidentes':        IntranetCctvIncidentes,
            'EstadosCamara':     IntranetCctvEstadosCamara,
            'MetodosBackup':     IntranetCctvMetodosBackup,
            'NivelesAcceso':     IntranetCctvNivelesAcceso,
            'Severidades':       IntranetCctvSeveridades,
            'EstadosIncidente':  IntranetCctvEstadosIncidente,
        }

    # ── Catálogos ──────────────────────────────────────────────────────────────

    def cctv_obtener_catalogos(self):
        m = self._cctv_models()
        def _cat(model):
            return [r.to_dict() for r in
                    self.db.query(model).filter(model.activo == True).order_by(model.orden).all()]
        return {
            'estados_camara':    _cat(m['EstadosCamara']),
            'metodos_backup':    _cat(m['MetodosBackup']),
            'niveles_acceso':    _cat(m['NivelesAcceso']),
            'severidades':       _cat(m['Severidades']),
            'estados_incidente': _cat(m['EstadosIncidente']),
        }

    # ── Dashboard ──────────────────────────────────────────────────────────────

    def cctv_dashboard(self):
        m = self._cctv_models()
        S = m['Sedes']
        C = m['Camaras']

        total_sedes   = self.db.query(S).filter(S.activo == True).count()
        total_camaras = self.db.query(C).filter(C.activo == True).count()
        dias_total    = self.db.query(func.sum(C.dias_almacenamiento)).filter(C.activo == True).scalar() or 0
        total_cargos  = self.db.query(m['Cargos']).filter(m['Cargos'].activo == True).count()

        sedes   = self.db.query(S).filter(S.activo == True).order_by(S.nombre).all()
        por_sede = []
        for sede in sedes:
            camaras = self.db.query(C).filter(C.id_sede == sede.id, C.activo == True).all()
            por_sede.append({
                'id_sede':             sede.id,
                'nombre':              sede.nombre,
                'ubicacion_general':   sede.ubicacion_general,
                'camaras':             len(camaras),
                'dias_almacenamiento': sede.dias_almacenamiento_estimado,
                'dias_backup':         sum(c.dias_retencion_backup or 0 for c in camaras),
            })

        return {
            'totales': {
                'sedes':               total_sedes,
                'camaras':             total_camaras,
                'dias_almacenamiento': int(dias_total),
                'cargos_autorizados':  total_cargos,
            },
            'por_sede': por_sede,
        }

    # ── Sedes ──────────────────────────────────────────────────────────────────

    def cctv_listar_sedes(self):
        m = self._cctv_models()
        return [r.to_dict() for r in
                self.db.query(m['Sedes']).filter(m['Sedes'].activo == True)
                .order_by(m['Sedes'].nombre).all()]

    def cctv_crear_sede(self, data: dict) -> dict:
        m = self._cctv_models()
        sede = m['Sedes'](data)
        sede.fecha_creacion      = self._get_fecha_colombia()
        sede.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(sede)
        self.db.commit()
        self.db.refresh(sede)
        return sede.to_dict()

    def cctv_actualizar_sede(self, id_sede: int, data: dict):
        m    = self._cctv_models()
        sede = self.db.query(m['Sedes']).filter(m['Sedes'].id == id_sede, m['Sedes'].activo == True).first()
        if not sede:
            return None
        campos = ['nombre', 'ubicacion_general', 'responsable_operativo',
                  'sistema_grabacion', 'dias_almacenamiento_estimado', 'observaciones',
                  'usuario_actualizacion']
        for campo in campos:
            if campo in data:
                setattr(sede, campo, data[campo])
        sede.fecha_actualizacion = self._get_fecha_colombia()
        self.db.commit()
        self.db.refresh(sede)
        return sede.to_dict()

    def cctv_eliminar_sede(self, id_sede: int):
        m    = self._cctv_models()
        sede = self.db.query(m['Sedes']).filter(m['Sedes'].id == id_sede).first()
        if sede:
            sede.activo              = False
            sede.fecha_actualizacion = self._get_fecha_colombia()
            camaras = self.db.query(m['Camaras']).filter(m['Camaras'].id_sede == id_sede).all()
            for c in camaras:
                c.activo             = False
                c.fecha_actualizacion = self._get_fecha_colombia()
            self.db.commit()

    # ── Cámaras ────────────────────────────────────────────────────────────────

    def _enrich_camara(self, camara, m):
        d   = camara.to_dict()
        s   = self.db.query(m['Sedes']).filter(m['Sedes'].id == camara.id_sede).first()
        ec  = self.db.query(m['EstadosCamara']).filter(m['EstadosCamara'].id == camara.id_estado_camara).first()
        mb  = self.db.query(m['MetodosBackup']).filter(m['MetodosBackup'].id == camara.id_metodo_backup).first()
        d['nombre_sede']         = s.nombre   if s  else ''
        d['estado']              = ec.nombre  if ec else ''
        d['estado_valor']        = ec.valor   if ec else ''
        d['metodo_backup']       = mb.nombre  if mb else ''
        d['metodo_backup_valor'] = mb.valor   if mb else ''
        return d

    def cctv_listar_camaras(self, filtros: dict = {}):
        m = self._cctv_models()
        C = m['Camaras']
        S = m['Sedes']

        q = (self.db.query(C)
             .join(S, C.id_sede == S.id)
             .filter(C.activo == True))

        if filtros.get('id_sede'):
            q = q.filter(C.id_sede == filtros['id_sede'])
        if filtros.get('id_estado_camara'):
            q = q.filter(C.id_estado_camara == filtros['id_estado_camara'])
        if filtros.get('query'):
            term = f"%{filtros['query'].strip().lower()}%"
            q = q.filter(or_(
                C.codigo_equipo_grabacion.ilike(term),
                C.ubicacion_fisica.ilike(term),
                S.nombre.ilike(term),
            ))

        q = q.order_by(S.nombre, C.codigo_equipo_grabacion)
        return [self._enrich_camara(c, m) for c in q.all()]

    def cctv_crear_camara(self, data: dict) -> dict:
        m      = self._cctv_models()
        camara = m['Camaras'](data)
        camara.fecha_creacion      = self._get_fecha_colombia()
        camara.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(camara)
        self.db.commit()
        self.db.refresh(camara)
        return self._enrich_camara(camara, m)

    def cctv_actualizar_camara(self, id_camara: int, data: dict):
        m      = self._cctv_models()
        camara = self.db.query(m['Camaras']).filter(m['Camaras'].id == id_camara, m['Camaras'].activo == True).first()
        if not camara:
            return None
        campos = ['id_sede', 'codigo_equipo_grabacion', 'ubicacion_fisica',
                  'id_estado_camara', 'dias_almacenamiento', 'id_metodo_backup',
                  'dias_retencion_backup', 'fecha_instalacion_actualizacion',
                  'observaciones', 'usuario_actualizacion']
        for campo in campos:
            if campo in data:
                setattr(camara, campo, data[campo])
        camara.fecha_actualizacion = self._get_fecha_colombia()
        self.db.commit()
        self.db.refresh(camara)
        return self._enrich_camara(camara, m)

    def cctv_eliminar_camara(self, id_camara: int):
        m      = self._cctv_models()
        camara = self.db.query(m['Camaras']).filter(m['Camaras'].id == id_camara).first()
        if camara:
            camara.activo             = False
            camara.fecha_actualizacion = self._get_fecha_colombia()
            self.db.commit()

    # ── Cargos ─────────────────────────────────────────────────────────────────

    def _enrich_cargo(self, cargo, m):
        d  = cargo.to_dict()
        na = self.db.query(m['NivelesAcceso']).filter(m['NivelesAcceso'].id == cargo.id_nivel_acceso).first()
        d['nivel_acceso']       = na.nombre if na else ''
        d['nivel_acceso_valor'] = na.valor  if na else ''
        return d

    def cctv_listar_cargos(self):
        m  = self._cctv_models()
        CR = m['Cargos']
        NA = m['NivelesAcceso']

        rows = (self.db.query(CR, NA.nombre, NA.valor)
                .join(NA, CR.id_nivel_acceso == NA.id)
                .filter(CR.activo == True)
                .order_by(CR.nombre).all())

        resultado = []
        for cargo, na_nombre, na_valor in rows:
            d = cargo.to_dict()
            d['nivel_acceso']       = na_nombre
            d['nivel_acceso_valor'] = na_valor
            resultado.append(d)
        return resultado

    def cctv_crear_cargo(self, data: dict) -> dict:
        m     = self._cctv_models()
        cargo = m['Cargos'](data)
        cargo.fecha_creacion      = self._get_fecha_colombia()
        cargo.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(cargo)
        self.db.commit()
        self.db.refresh(cargo)
        return self._enrich_cargo(cargo, m)

    def cctv_actualizar_cargo(self, id_cargo: int, data: dict):
        import json
        m     = self._cctv_models()
        cargo = self.db.query(m['Cargos']).filter(m['Cargos'].id == id_cargo, m['Cargos'].activo == True).first()
        if not cargo:
            return None
        campos_simples = ['nombre', 'id_nivel_acceso', 'justificacion_acceso',
                          'puede_ver_camaras', 'puede_editar_inventario',
                          'puede_administrar_configuracion', 'usuario_actualizacion']
        for campo in campos_simples:
            if campo in data:
                setattr(cargo, campo, data[campo])
        if 'ids_sedes' in data:
            ids = data['ids_sedes'] or []
            cargo.ids_sedes = json.dumps([int(i) for i in ids])
        cargo.fecha_actualizacion = self._get_fecha_colombia()
        self.db.commit()
        self.db.refresh(cargo)
        return self._enrich_cargo(cargo, m)

    def cctv_eliminar_cargo(self, id_cargo: int):
        m     = self._cctv_models()
        cargo = self.db.query(m['Cargos']).filter(m['Cargos'].id == id_cargo).first()
        if cargo:
            cargo.activo             = False
            cargo.fecha_actualizacion = self._get_fecha_colombia()
            self.db.commit()

    # ── Registros de cambios ───────────────────────────────────────────────────

    def cctv_listar_cambios(self):
        m  = self._cctv_models()
        RC = m['RegistrosCambios']
        S  = m['Sedes']
        C  = m['Camaras']

        rows = (self.db.query(RC, S.nombre, C.codigo_equipo_grabacion)
                .outerjoin(S, RC.id_sede   == S.id)
                .outerjoin(C, RC.id_camara == C.id)
                .filter(RC.activo == True)
                .order_by(RC.fecha_creacion.desc())
                .limit(100).all())

        resultado = []
        for cambio, nombre_sede, codigo_camara in rows:
            d = cambio.to_dict()
            d['nombre_sede']   = nombre_sede
            d['codigo_camara'] = codigo_camara
            resultado.append(d)
        return resultado

    def cctv_crear_cambio(self, data: dict) -> dict:
        m      = self._cctv_models()
        cambio = m['RegistrosCambios'](data)
        cambio.fecha_creacion      = self._get_fecha_colombia()
        cambio.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(cambio)
        self.db.commit()
        self.db.refresh(cambio)
        d = cambio.to_dict()
        if cambio.id_sede:
            sede = self.db.query(m['Sedes']).filter(m['Sedes'].id == cambio.id_sede).first()
            d['nombre_sede'] = sede.nombre if sede else None
        return d

    # ── Revisiones ─────────────────────────────────────────────────────────────

    def cctv_listar_revisiones(self):
        m = self._cctv_models()
        return [r.to_dict() for r in
                self.db.query(m['Revisiones']).filter(m['Revisiones'].activo == True)
                .order_by(m['Revisiones'].fecha_creacion.desc()).limit(100).all()]

    def cctv_crear_revision(self, data: dict) -> dict:
        m        = self._cctv_models()
        revision = m['Revisiones'](data)
        revision.fecha_creacion      = self._get_fecha_colombia()
        revision.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(revision)
        self.db.commit()
        self.db.refresh(revision)
        return revision.to_dict()

    # ── Incidentes ─────────────────────────────────────────────────────────────

    def _enrich_incidente(self, inc, m):
        d  = inc.to_dict()
        sv = self.db.query(m['Severidades']).filter(m['Severidades'].id == inc.id_severidad).first()
        ei = self.db.query(m['EstadosIncidente']).filter(m['EstadosIncidente'].id == inc.id_estado_incidente).first()
        d['severidad']       = sv.nombre if sv else ''
        d['severidad_valor'] = sv.valor  if sv else ''
        d['estado']          = ei.nombre if ei else ''
        d['estado_valor']    = ei.valor  if ei else ''
        return d

    def cctv_listar_incidentes(self):
        m  = self._cctv_models()
        I  = m['Incidentes']
        S  = m['Sedes']
        C  = m['Camaras']

        rows = (self.db.query(I, S.nombre, C.codigo_equipo_grabacion)
                .outerjoin(S, I.id_sede   == S.id)
                .outerjoin(C, I.id_camara == C.id)
                .filter(I.activo == True)
                .order_by(I.fecha_creacion.desc())
                .limit(100).all())

        resultado = []
        for inc, nombre_sede, codigo_camara in rows:
            d = self._enrich_incidente(inc, m)
            d['nombre_sede']   = nombre_sede
            d['codigo_camara'] = codigo_camara
            resultado.append(d)
        return resultado

    def cctv_crear_incidente(self, data: dict) -> dict:
        m         = self._cctv_models()
        incidente = m['Incidentes'](data)
        incidente.fecha_creacion      = self._get_fecha_colombia()
        incidente.fecha_actualizacion = self._get_fecha_colombia()
        self.db.add(incidente)
        self.db.commit()
        self.db.refresh(incidente)
        return self._enrich_incidente(incidente, m)

    def cctv_actualizar_incidente(self, id_incidente: int, data: dict):
        m         = self._cctv_models()
        incidente = self.db.query(m['Incidentes']).filter(
            m['Incidentes'].id == id_incidente, m['Incidentes'].activo == True).first()
        if not incidente:
            return None
        campos = ['id_estado_incidente', 'id_severidad', 'accion_correctiva', 'titulo',
                  'descripcion', 'usuario_actualizacion']
        for campo in campos:
            if campo in data:
                setattr(incidente, campo, data[campo])
        incidente.fecha_actualizacion = self._get_fecha_colombia()
        self.db.commit()
        self.db.refresh(incidente)
        return self._enrich_incidente(incidente, m)
