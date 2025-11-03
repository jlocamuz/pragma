DEFAULT_CONFIG = {
    # API furlong
    'api_key': 'Njc0Mjg4OTpfMmZweXdOMEhOaDU0YzRRRW8xNXZhNjFadTNwOHhyZg==',

    'base_url': 'https://api-prod.humand.co/public/api/v1',

    "local_timezone": "America/Argentina/Buenos_Aires",



    "hora_viatico1_comida1": 6, 
    # Jornada
    'jornada_completa_horas': 8,
    'tolerancia_minutos': 20,
    'fragmento_minutos': 30,

    # Horarios especiales
    'hora_nocturna_inicio': 22,  # 21:00
    'hora_nocturna_fin': 6,      # 06:00
    'sabado_limite_hora': 13,

    # Zona horaria
    'timezone': 'America/Argentina/Buenos_Aires',

    # Red/Requests
    'max_retries': 3,
    'retry_delay': 1000,
    'request_timeout': 30000,

    # Paralelismo
    'max_workers': 6,
    'batch_size_users': 10,
    'batch_size_dates': 7,
    'delay_between_retries': 1000,
    'delay_between_batches': 500,

    # Archivos
    'output_directory': '~/Downloads',
    'filename_format': 'reporte_{start_date}_{end_date}.xlsx',

    # UI
    'window_width': 800,
    'window_height': 600,
    'theme': 'default',


}

# Headers API
def get_api_headers(api_key=None):
    if api_key is None:
        api_key = DEFAULT_CONFIG['api_key']
    return {
        'Authorization': f'Basic {api_key}',
        'Content-Type': 'application/json'
    }

# Endpoints
API_ENDPOINTS = {
    'users': '/users',
    'time_tracking_entries': '/time-tracking/entries',
    'day_summaries': '/time-tracking/day-summaries'
}


