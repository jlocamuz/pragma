"""
Procesador de datos principal
Coordina la obtención de datos de la API y el procesamiento de horas
"""

from typing import Dict, List, Optional, Callable
from datetime import datetime
from core.api_client import HumanApiClient
from core.hours_calculator import ArgentineHoursCalculator
from core.excel_generator import ExcelReportGenerator


class DataProcessor:
    """Procesador principal de datos de asistencia"""
    
    def __init__(self, api_key: str = None, base_url: str = None):
        self.api_client = HumanApiClient(api_key, base_url)
        self.hours_calculator = ArgentineHoursCalculator()
        self.excel_generator = ExcelReportGenerator()
        
        # Cache para optimizar rendimiento
        self._users_cache = None
        self._departments_cache = None
        self._cache_timestamp = None
        self._cache_duration = 300  # 5 minutos
    
    def test_connection(self) -> tuple[bool, str]:
        """Prueba la conexión con la API"""
        return self.api_client.test_connection()
    
    def get_users_list(self, filters: Dict = None, use_cache: bool = True) -> List[Dict]:
        """
        Obtiene la lista de usuarios disponibles con cache
        Args:
            filters: Filtros opcionales para usuarios
            use_cache: Si usar cache o forzar recarga
        Returns:
            Lista de usuarios
        """
        # Verificar cache
        if use_cache and self._is_cache_valid():
            print("✅ Usando usuarios desde cache")
            users = self._users_cache
        else:
            print("🔄 Cargando usuarios desde API...")
            users = self.api_client.get_users(filters)
            self._update_cache(users)
        
        # Aplicar filtros si se especificaron
        if filters:
            return self._apply_user_filters(users, filters)
        
        return users
    
    def _is_cache_valid(self) -> bool:
        """Verifica si el cache es válido"""
        if not self._users_cache or not self._cache_timestamp:
            return False
        
        now = datetime.now()
        cache_age = (now - self._cache_timestamp).total_seconds()
        return cache_age < self._cache_duration
    
    def _update_cache(self, users: List[Dict]):
        """Actualiza el cache de usuarios"""
        self._users_cache = users
        self._cache_timestamp = datetime.now()
        
        # Actualizar cache de departamentos
        departments = set()
        for user in users:
            if user.get('department'):
                departments.add(user['department'])
        self._departments_cache = sorted(list(departments))
    
    def _apply_user_filters(self, users: List[Dict], filters: Dict) -> List[Dict]:
        """Aplica filtros a la lista de usuarios"""
        filtered_users = []
        
        for user in users:
            matches = True
            
            # Filtrar por departamento
            if filters.get('department') and user.get('department') != filters['department']:
                matches = False
            
            # Filtrar por estado activo
            if filters.get('active_only', True) and not user.get('isActive', True):
                matches = False
            
            if matches:
                filtered_users.append(user)
        
        return filtered_users
    
    def process_attendance_report(self, start_date: str, end_date: str, 
                                user_ids: List[str] = None,
                                progress_callback: Callable = None) -> Dict:
        """
        Procesa un reporte completo de asistencia
        Args:
            start_date: Fecha de inicio (YYYY-MM-DD)
            end_date: Fecha de fin (YYYY-MM-DD)
            user_ids: Lista opcional de IDs de usuarios
            progress_callback: Función de callback para progreso
        Returns:
            Diccionario con el resultado del procesamiento
        """
        try:
            if progress_callback:
                progress_callback(0, "Iniciando procesamiento...")
            
            # 1. Obtener datos de la API usando procesamiento paralelo
            if progress_callback:
                progress_callback(5, "Conectando con la API...")
            
            # Usar usuarios del cache en lugar de re-descargar
            if user_ids:
                # Si hay user_ids específicos, usar solo esos
                cached_users = self.get_users_list()
                filtered_users = [u for u in cached_users if u.get('employeeInternalId') in user_ids]
            else:
                # Usar todos los usuarios del cache
                filtered_users = self.get_users_list()
            
            api_result = self.api_client.get_time_tracking_parallel_with_users(
                start_date, end_date, filtered_users,
                lambda p, m: progress_callback(5 + int(p * 0.6), m) if progress_callback else None
            )
            
            if not api_result['success']:
                return {
                    'success': False,
                    'error': api_result.get('error', 'Error desconocido en la API'),
                    'stage': 'api_fetch'
                }
            
            if progress_callback:
                progress_callback(70, "Procesando datos de empleados...")
            
            # 2. Procesar datos de cada empleado
            processed_employees = {}
            users_data = api_result['users']
            entries_data = api_result['entries']
            
            print(f"📊 Usuarios obtenidos: {len(users_data)}")
            print(f"📊 Entradas obtenidas: {len(entries_data)}")
            
            # Agrupar entradas por empleado usando el campo correcto
            entries_by_employee = {}
            for entry in entries_data:
                employee_id = entry.get('employeeId')  # Campo correcto según debugging
                if employee_id:
                    if employee_id not in entries_by_employee:
                        entries_by_employee[employee_id] = []
                    entries_by_employee[employee_id].append(entry)
            
            print(f"📊 Empleados con entradas: {len(entries_by_employee)}")
            
            total_employees = len(users_data)
            processed_count = 0
            
            for employee_id, employee_info in users_data.items():
                if progress_callback:
                    progress = 70 + int((processed_count / total_employees) * 20)
                    employee_name = f"{employee_info.get('firstName', '')} {employee_info.get('lastName', '')}"
                    progress_callback(progress, f"Procesando {employee_name}...")
                
                # Obtener entradas del empleado
                employee_entries = entries_by_employee.get(employee_id, [])
                
                # Procesar datos del empleado
                employee_data = self.hours_calculator.process_employee_data(
                    employee_entries, employee_info, 0, None
                )
                
                processed_employees[employee_id] = employee_data
                processed_count += 1
            
            if progress_callback:
                progress_callback(90, "Generando reporte Excel...")
            
            # 3. Generar reporte Excel
            excel_path = self.excel_generator.generate_report(
                processed_employees, start_date, end_date
            )
            
            if progress_callback:
                progress_callback(100, "¡Reporte completado!")
            
            # 4. Calcular estadísticas finales
            stats = self._calculate_final_stats(processed_employees)
            
            return {
                'success': True,
                'excel_path': excel_path,
                'stats': stats,
                'processed_employees': len(processed_employees),
                'date_range': {
                    'start_date': start_date,
                    'end_date': end_date
                },
                'api_stats': {
                    'total_users': api_result['total_users'],
                    'total_entries': api_result['total_entries']
                }
            }
            
        except Exception as e:
            error_msg = f"Error en procesamiento: {str(e)}"
            print(f"❌ {error_msg}")
            return {
                'success': False,
                'error': error_msg,
                'stage': 'processing'
            }
    
    def _calculate_final_stats(self, processed_employees: Dict) -> Dict:
        """Calcula estadísticas finales del reporte"""
        if not processed_employees:
            return {
                'total_employees': 0,
                'total_hours_worked': 0,
                'total_regular_hours': 0,
                'total_extra_hours_50': 0,
                'total_extra_hours_100': 0,
                'total_night_hours': 0,
                'total_pending_hours': 0,
                'avg_hours_per_employee': 0
            }
        
        total_employees = len(processed_employees)
        total_hours_worked = sum(emp['totals']['total_hours_worked'] for emp in processed_employees.values())
        total_regular_hours = sum(emp['totals']['total_regular_hours'] for emp in processed_employees.values())
        total_extra_hours_50 = sum(emp['totals']['total_extra_hours_50'] for emp in processed_employees.values())
        total_extra_hours_100 = sum(emp['totals']['total_extra_hours_100'] for emp in processed_employees.values())
        total_night_hours = sum(emp['totals']['total_night_hours'] for emp in processed_employees.values())
        total_pending_hours = sum(emp['compensations']['remaining_pending_hours'] for emp in processed_employees.values())
        
        return {
            'total_employees': total_employees,
            'total_hours_worked': round(total_hours_worked, 2),
            'total_regular_hours': round(total_regular_hours, 2),
            'total_extra_hours_50': round(total_extra_hours_50, 2),
            'total_extra_hours_100': round(total_extra_hours_100, 2),
            'total_night_hours': round(total_night_hours, 2),
            'total_pending_hours': round(total_pending_hours, 2),
            'avg_hours_per_employee': round(total_hours_worked / total_employees, 2) if total_employees > 0 else 0
        }
    
    def get_available_filters(self, progress_callback: Callable = None) -> Dict:
        """
        Obtiene los filtros disponibles basados en los usuarios
        Args:
            progress_callback: Función de callback para progreso
        Returns:
            Diccionario con opciones de filtrado
        """
        try:
            if progress_callback:
                progress_callback(20, "📋 Obteniendo usuarios...")
            
            users = self.get_users_list()
            
            if progress_callback:
                progress_callback(60, f"📊 Procesando {len(users)} usuarios...")
            
            # Extraer departamentos únicos
            departments = set()
            locations = set()
            job_titles = set()
            
            for user in users:
                if user.get('department'):
                    departments.add(user['department'])
                if user.get('location'):
                    locations.add(user['location'])
                if user.get('jobTitle'):
                    job_titles.add(user['jobTitle'])
            
            if progress_callback:
                progress_callback(80, "🔧 Configurando filtros...")
            
            return {
                'departments': sorted(list(departments)),
                'locations': sorted(list(locations)),
                'job_titles': sorted(list(job_titles)),
                'total_users': len(users)
            }
            
        except Exception as e:
            print(f"❌ Error obteniendo filtros: {str(e)}")
            return {
                'departments': [],
                'locations': [],
                'job_titles': [],
                'total_users': 0
            }
    
    def filter_users_by_criteria(self, criteria: Dict) -> List[str]:
        """
        Filtra usuarios según criterios específicos
        Args:
            criteria: Diccionario con criterios de filtrado
        Returns:
            Lista de IDs de usuarios que cumplen los criterios
        """
        try:
            users = self.get_users_list()
            filtered_user_ids = []
            
            for user in users:
                matches = True
                
                # Filtrar por departamento
                if criteria.get('department') and user.get('department') != criteria['department']:
                    matches = False
                
                # Filtrar por ubicación
                if criteria.get('location') and user.get('location') != criteria['location']:
                    matches = False
                
                # Filtrar por título de trabajo
                if criteria.get('job_title') and user.get('jobTitle') != criteria['job_title']:
                    matches = False
                
                # Filtrar por estado activo
                if criteria.get('active_only', True) and not user.get('isActive', True):
                    matches = False
                
                if matches:
                    filtered_user_ids.append(user.get('employeeInternalId'))
            
            return filtered_user_ids
            
        except Exception as e:
            print(f"❌ Error filtrando usuarios: {str(e)}")
            return []
    
    def validate_date_range(self, start_date: str, end_date: str) -> Dict:
        """
        Valida un rango de fechas
        Args:
            start_date: Fecha de inicio (YYYY-MM-DD)
            end_date: Fecha de fin (YYYY-MM-DD)
        Returns:
            Diccionario con resultado de validación
        """
        try:
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
            
            errors = []
            warnings = []
            
            # Validaciones básicas
            if start_dt > end_dt:
                errors.append("La fecha de inicio debe ser anterior a la fecha de fin")
            
            # Validar que no sea muy futuro
            now = datetime.now()
            if end_dt > now:
                warnings.append("La fecha de fin es futura, algunos datos pueden no estar disponibles")
            
            # Validar que no sea muy antiguo
            min_date = datetime(2024, 1, 1)
            if start_dt < min_date:
                warnings.append("Fechas muy antiguas pueden tener datos limitados")
            
            # Validar rango no muy amplio
            diff_days = (end_dt - start_dt).days + 1
            if diff_days > 365:
                warnings.append("Rangos muy amplios pueden afectar el rendimiento")
            
            return {
                'is_valid': len(errors) == 0,
                'errors': errors,
                'warnings': warnings,
                'day_count': diff_days
            }
            
        except ValueError as e:
            return {
                'is_valid': False,
                'errors': [f"Formato de fecha inválido: {str(e)}"],
                'warnings': [],
                'day_count': 0
            }
    
    def get_user_count(self, department: str = None) -> int:
        """
        Obtiene el conteo de usuarios total o por departamento
        Args:
            department: Departamento específico (opcional)
        Returns:
            Número de usuarios
        """
        try:
            if department:
                # Filtrar por departamento
                filters = {'department': department}
                users = self.get_users_list(filters)
                return len(users)
            else:
                # Todos los usuarios
                users = self.get_users_list()
                return len(users)
                
        except Exception as e:
            print(f"❌ Error obteniendo conteo de usuarios: {str(e)}")
            return 0
    
    def get_departments_with_counts(self) -> Dict[str, int]:
        """
        Obtiene departamentos con sus respectivos conteos de usuarios
        Returns:
            Diccionario {departamento: cantidad_usuarios}
        """
        try:
            users = self.get_users_list()
            department_counts = {}
            
            for user in users:
                dept = user.get('department', 'Sin Departamento')
                if dept in department_counts:
                    department_counts[dept] += 1
                else:
                    department_counts[dept] = 1
            
            return department_counts
            
        except Exception as e:
            print(f"❌ Error obteniendo departamentos con conteos: {str(e)}")
            return {}
    
    def refresh_cache(self):
        """Fuerza la actualización del cache de usuarios"""
        print("🔄 Refrescando cache de usuarios...")
        self.get_users_list(use_cache=False)
