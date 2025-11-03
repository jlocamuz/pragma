# 📊 Generador de Reportes de Horas Extras - Argentina

Sistema automatizado para generar reportes de asistencia según normativa laboral argentina.

## 🚀 Instalación

### Requisitos Previos
- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### Pasos de Instalación

1. **Extraer el archivo ZIP** en la ubicación deseada

2. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Ejecutar la aplicación**:
   ```bash
   python src/main.py
   ```

## 🔧 Configuración

### API Key
La API key está preconfigurada en `src/config/default_config.py`. Si necesitas cambiarla:

```python
'api_key': 'TU_NUEVA_API_KEY_AQUI'
```

### Configuración de Salida
Los reportes se guardan por defecto en `~/Downloads`. Para cambiar:

```python
'output_directory': '/ruta/personalizada'
```

## 📋 Uso

1. **Iniciar la aplicación** - Se abrirá la interfaz gráfica
2. **Seleccionar fechas** - Usar los campos o presets disponibles
3. **Elegir filtros** - Todos los usuarios o por departamento
4. **Generar reporte** - Presionar el botón "GENERAR REPORTE"
5. **Abrir archivo** - El sistema preguntará si deseas abrir el Excel generado

## 📊 Características

### Normativa Argentina Aplicada
- **Jornada completa**: 8 horas diarias
- **Horas extras 50%**: Primeras 2 horas extras en días laborables
- **Horas extras 100%**: Horas adicionales, sábados tarde, domingos y feriados
- **Horas nocturnas**: 21:00 a 06:00
- **Compensaciones**: Sistema automático de compensación de horas

### Hojas del Reporte Excel
1. **Resumen Consolidado**: Totales por empleado
2. **Detalle Diario**: Registro día a día de cada empleado
3. **Estadísticas**: Gráficos y análisis general
4. **Configuración**: Parámetros utilizados y normativa aplicada

## 🛠️ Solución de Problemas

### Error de Conexión
- Verificar conexión a internet
- Comprobar que la API key sea válida

### Error de Dependencias
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### Error de Permisos
- En macOS/Linux: `chmod +x src/main.py`
- Ejecutar como administrador si es necesario

## 📞 Soporte

Para reportar problemas o solicitar mejoras, contactar al desarrollador.

---
**Versión**: 2.0  
**Última actualización**: Agosto 2025
