"""
Cliente de API para Human.co
Maneja todas las comunicaciones con la API externa
"""

import requests
import time
import json
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
from config.default_config import DEFAULT_CONFIG, get_api_headers, API_ENDPOINTS


class HumanApiClient:
    """Cliente para interactuar con la API de Human.co"""
    
    def __init__(self, api_key: str = None, base_url: str = None):
        self.api_key = api_key or DEFAULT_CONFIG['api_key']
        self.base_url = base_url or DEFAULT_CONFIG['base_url']
        self.session = requests.Session()
        self.session.headers.update(get_api_headers(self.api_key))
        
        # Configuración de timeouts y reintentos
        self.max_retries = DEFAULT_CONFIG['max_retries']
        self.retry_delay = DEFAULT_CONFIG['retry_delay'] / 1000  # Convertir a segundos
        self.timeout = DEFAULT_CONFIG['request_timeout'] / 1000  # Convertir a segundos
    
    def test_connection(self) -> Tuple[bool, str]:
        """
        Prueba la conexión con la API
        Returns: (success: bool, message: str)
        """
        try:
            response = self._make_request('GET', API_ENDPOINTS['users'], params={'page': 1, 'limit': 1})
            if response:
                return True, "Conexión exitosa con la API"
            else:
                return False, "Error al conectar con la API"
        except Exception as e:
            return False, f"Error de conexión: {str(e)}"
    
    def get_users(self, filters: Dict = None) -> List[Dict]:
        """
        Obtiene la lista de usuarios desde la API usando paginación
        Args:
            filters: Filtros opcionales para usuarios
        Returns:
            Lista de usuarios
        """
        try:
            params = {
                'page': 1,
                'limit': 50,
                'search': 'usernocturno'
            }
            if filters:
                params.update(filters)
            
            all_users = []
            page = 1
            has_more_pages = True
            
            while has_more_pages:
                params['page'] = page
                response = self._make_request('GET', API_ENDPOINTS['users'], params=params)
                
                # La API de usuarios devuelve {count: X, users: [...]} según el proyecto anterior
                if response and 'users' in response and len(response['users']) > 0:
                    #print('response users',response['users'])
                    all_users.extend(response['users'])
                    
                    # Calcular si hay más páginas basándose en count total
                    total_users = response.get('count', 0)
                    users_obtained = page * params['limit']
                    has_more_pages = users_obtained < total_users
                    page += 1
                else:
                    has_more_pages = False
            
            print(f"✅ Obtenidos {len(all_users)} usuarios de la API")
            # for user in all_users:
            #     for seg in user.get("segmentations", []):
            #         if seg.get("group") == "CONVENIO":
            #             print(f"Usuario {user['employeeInternalId']} -> {seg['item']}")

            return all_users
                
        except Exception as e:
            print(f"❌ Error obteniendo usuarios: {str(e)}")
            return []
    
    def get_time_tracking_entries(self, start_date: str, end_date: str, 
                                user_ids: List[str] = None) -> List[Dict]:
        """
        Obtiene las entradas de seguimiento de tiempo
        Args:
            start_date: Fecha de inicio (YYYY-MM-DD)
            end_date: Fecha de fin (YYYY-MM-DD)
            user_ids: Lista opcional de IDs de usuarios
        Returns:
            Lista de entradas de tiempo
        """

        print("ejecutandose get_time_tracking")
        try:
            params = {
                'startDate': start_date,
                'endDate': end_date,
                'limit': 100
                
            }
            
            if user_ids:
                params['userIds'] = ','.join(user_ids)
            
            print("params", params)
            response = self._make_request('GET', API_ENDPOINTS['time_tracking_entries'], params=params)
            if response and 'data' in response:
                entries = response['data']
                print(f"✅ Obtenidas {len(entries)} entradas de tiempo")
                return entries
            else:
                print("❌ No se pudieron obtener entradas de tiempo")
                return []
                
        except Exception as e:
            print(f"❌ Error obteniendo entradas de tiempo: {str(e)}")
            return []
    
    def get_day_summaries(self, start_date: str, end_date: str, 
                         user_ids: List[str] = None) -> List[Dict]:
        """
        Obtiene los resúmenes diarios usando lotes optimizados y procesamiento paralelo
        Args:
            start_date: Fecha de inicio (YYYY-MM-DD)
            end_date: Fecha de fin (YYYY-MM-DD)
            user_ids: Lista opcional de IDs de usuarios
        Returns:
            Lista de resúmenes diarios
        """
        try:
            # Lotes más grandes para mejor rendimiento
            BATCH_SIZE = 15
            all_items = []
            
            if not user_ids:
                # Si no hay user_ids específicos, obtener todos los usuarios
                users = self.get_users()
                user_ids = [u.get('employeeInternalId') for u in users if u.get('employeeInternalId')]
            
            print(f"📋 Procesando {len(user_ids)} empleados en lotes de {BATCH_SIZE}...")
            
            # Crear lotes
            batches = []
            for i in range(0, len(user_ids), BATCH_SIZE):
                batch = user_ids[i:i + BATCH_SIZE]
                batches.append({
                    'batch_number': (i // BATCH_SIZE) + 1,
                    'user_ids': batch,
                    'start_date': start_date,
                    'end_date': end_date
                })
            
            # Procesar lotes en paralelo
            with ThreadPoolExecutor(max_workers=3) as executor:
                future_to_batch = {
                    executor.submit(self._process_batch_summaries, batch): batch 
                    for batch in batches
                }
                
                for future in as_completed(future_to_batch):
                    batch = future_to_batch[future]
                    try:
                        batch_items = future.result()
                        all_items.extend(batch_items)
                        print(f"✅ Lote {batch['batch_number']}: {len(batch_items)} day summaries")
                    except Exception as e:
                        print(f"❌ Error en lote {batch['batch_number']}: {str(e)}")
            
            print(f"✅ Obtenidos {len(all_items)} resúmenes diarios")
            return all_items
                
        except Exception as e:
            print(f"❌ Error obteniendo resúmenes diarios: {str(e)}")
            return []
    
    def _process_batch_summaries(self, batch: Dict) -> List[Dict]:
        """
        Procesa un lote de usuarios para obtener day summaries
        """
        batch_items = []
        
        params = {
            'employeeIds': ','.join(batch['user_ids']),
            'startDate': batch['start_date'],
            'endDate': batch['end_date'],
            'limit': 500,
            'page': 1
        }
        
        # Obtener todas las páginas para este lote
        page = 1
        has_more_pages = True
        
        while has_more_pages:
            params['page'] = page
            
            response = self._make_request('GET', API_ENDPOINTS['day_summaries'], params=params)
            
            if response and 'items' in response and len(response['items']) > 0:
                batch_items.extend(response['items'])
                
                # Verificar si hay más páginas
                has_more_pages = response.get('totalPages', 0) > page
                page += 1
            else:
                has_more_pages = False
            
            # Pausa mínima entre páginas
            if has_more_pages:
                time.sleep(0.1)
        
        return batch_items
    
    def get_time_tracking_parallel_with_users(self, start_date: str, end_date: str, 
                                             users: List[Dict], 
                                             progress_callback=None) -> Dict:
        """
        Obtiene datos de seguimiento de tiempo usando usuarios del cache (optimizado)
        Args:
            start_date: Fecha de inicio
            end_date: Fecha de fin  
            users: Lista de usuarios ya obtenidos del cache
            progress_callback: Callback de progreso
        """
        try:
            print(f"🚀 Iniciando procesamiento paralelo: {start_date} a {end_date}")
            print(f"✅ Usando {len(users)} usuarios desde cache (sin re-descarga)")
            
            # 1. Usuarios ya proporcionados desde cache
            if progress_callback:
                progress_callback(10, "✅ Usuarios obtenidos desde cache...")
            
            if not users:
                return {'success': False, 'error': 'No hay usuarios disponibles'}
            
            print(f"👥 Procesando {len(users)} usuarios")
            
            # 2. Dividir rango de fechas en chunks
            if progress_callback:
                progress_callback(20, "📅 Dividiendo rango de fechas...")
            
            date_chunks = self._split_date_range(start_date, end_date, 30)  # Chunks de 30 días
            print(f"📅 Creados {len(date_chunks)} chunks de fechas")
            
            # 3. Procesar chunks en paralelo
            all_entries = []
            total_chunks = len(date_chunks)
            
            for i, chunk in enumerate(date_chunks):
                if progress_callback:
                    progress = 20 + (60 * (i + 1) / total_chunks)
                    progress_callback(int(progress), f"📊 Procesando chunk {i+1}/{total_chunks}...")
                
                chunk_entries = self.get_day_summaries(
                    chunk['start_date'], 
                    chunk['end_date'], 
                    [u.get('employeeInternalId') for u in users]
                )
                all_entries.extend(chunk_entries)
                
                # Pausa entre chunks para no sobrecargar la API
                if i < total_chunks - 1:
                    time.sleep(DEFAULT_CONFIG['delay_between_batches'] / 1000)
            
            if progress_callback:
                progress_callback(90, "🔧 Consolidando resultados...")
            
            # 4. Consolidar resultados
            result = {
                'success': True,
                'users': {u.get('employeeInternalId'): u for u in users},
                'entries': all_entries,
                'total_users': len(users),
                'total_entries': len(all_entries),
                'date_range': {
                    'start_date': start_date,
                    'end_date': end_date
                }
            }
            
            if progress_callback:
                progress_callback(100, "✅ Procesamiento completado!")
            
            print(f"✅ Procesamiento paralelo completado: {len(all_entries)} entradas")
            return result
            
        except Exception as e:
            error_msg = f"Error en procesamiento paralelo: {str(e)}"
            print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
    
    def _make_request(self, method: str, endpoint: str, params: Dict = None, 
                     data: Dict = None) -> Optional[Dict]:
        """
        Realiza una petición HTTP con reintentos automáticos
        """
        url = f"{self.base_url}{endpoint}"
        print(url)
        
        for attempt in range(self.max_retries):
            try:
                if method.upper() == 'GET':
                    response = self.session.get(url, params=params, timeout=self.timeout)
                elif method.upper() == 'POST':
                    response = self.session.post(url, params=params, json=data, timeout=self.timeout)
                else:
                    raise ValueError(f"Método HTTP no soportado: {method}")
                
                response.raise_for_status()
                return response.json()
                
            except requests.exceptions.RequestException as e:
                print(f"⚠️ Intento {attempt + 1}/{self.max_retries} falló: {str(e)}")
                
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)
                else:
                    print(f"❌ Todos los intentos fallaron para {endpoint}")
                    raise e
        
        return None
    
    def _split_date_range(self, start_date: str, end_date: str, max_days: int = 30) -> List[Dict]:
        """
        Divide un rango de fechas en chunks más pequeños
        """
        chunks = []
        current_start = datetime.strptime(start_date, '%Y-%m-%d')
        end_dt = datetime.strptime(end_date, '%Y-%m-%d')
        
        chunk_number = 1
        while current_start <= end_dt:
            current_end = min(current_start + timedelta(days=max_days - 1), end_dt)
            
            chunks.append({
                'chunk_number': chunk_number,
                'start_date': current_start.strftime('%Y-%m-%d'),
                'end_date': current_end.strftime('%Y-%m-%d'),
                'days': (current_end - current_start).days + 1
            })
            
            current_start = current_end + timedelta(days=1)
            chunk_number += 1
        
        return chunks
